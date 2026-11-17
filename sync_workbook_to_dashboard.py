import argparse
import getpass
import http.client
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from workbook_sync import MAX_WORKBOOK_BYTES, WorkbookSyncError, parse_workbook


DEFAULT_ENDPOINT = "https://offload-dashboard.onrender.com/api/upload"
TASK_NAME = "OffloadDashboardWorkbookSync"
CONFIG_DIRECTORY = Path(
    os.environ.get("LOCALAPPDATA")
    or os.environ.get("APPDATA")
    or Path.home()
) / "OffloadDashboard"
CONFIG_PATH = CONFIG_DIRECTORY / "sync-config.json"


def _read_workbook(path):
    workbook_path = Path(path).expanduser()
    if not workbook_path.is_file():
        raise WorkbookSyncError(f"Workbook file does not exist: {workbook_path}")
    size = workbook_path.stat().st_size
    if size <= 0 or size > MAX_WORKBOOK_BYTES:
        raise WorkbookSyncError("The workbook must be between 1 byte and 30 MB.")
    content = workbook_path.read_bytes()
    parse_workbook(content)
    return content


def configure(workbook_path, endpoint):
    workbook_path = Path(workbook_path).expanduser().resolve()
    _read_workbook(workbook_path)
    token = getpass.getpass("Paste the Render DASHBOARD_UPLOAD_TOKEN (input is hidden): ").strip()
    confirmation = getpass.getpass("Paste it again to confirm: ").strip()
    if len(token) < 32 or token != confirmation:
        raise WorkbookSyncError("The upload token must match both entries and be at least 32 characters.")
    if not endpoint.startswith("https://"):
        raise WorkbookSyncError("The dashboard upload endpoint must use HTTPS.")

    CONFIG_DIRECTORY.mkdir(parents=True, exist_ok=True)
    content = json.dumps({
        "workbookPath": str(workbook_path),
        "endpoint": endpoint,
        "token": token,
    }, indent=2)
    file_descriptor, temporary_path = tempfile.mkstemp(
        prefix=".sync-config-", suffix=".tmp", dir=CONFIG_DIRECTORY
    )
    try:
        with os.fdopen(file_descriptor, "w", encoding="utf-8", newline="\n") as output:
            output.write(content)
        os.replace(temporary_path, CONFIG_PATH)
    except Exception:
        try:
            os.unlink(temporary_path)
        except OSError:
            pass
        raise
    print(f"Sync configuration saved to {CONFIG_PATH}")


def sync_once():
    if not CONFIG_PATH.is_file():
        raise WorkbookSyncError(
            f"Sync is not configured. Run this script with --configure first."
        )
    try:
        config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        workbook_path = config["workbookPath"]
        endpoint = config["endpoint"]
        token = config["token"]
    except (OSError, KeyError, json.JSONDecodeError) as exc:
        raise WorkbookSyncError(f"Could not read sync configuration at {CONFIG_PATH}.") from exc
    if not endpoint.startswith("https://") or not token:
        raise WorkbookSyncError("The saved sync configuration is incomplete.")

    content = _read_workbook(workbook_path)
    request = Request(
        endpoint,
        data=content,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Length": str(len(content)),
            "User-Agent": "OffloadDashboardLocalSync/1.0",
        },
        method="POST",
    )
    try:
        with urlopen(request, timeout=120) as response:
            result = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        try:
            result = json.loads(exc.read().decode("utf-8"))
            message = result.get("error", f"Dashboard upload failed with HTTP {exc.code}.")
        except (UnicodeDecodeError, json.JSONDecodeError):
            message = f"Dashboard upload failed with HTTP {exc.code}."
        raise WorkbookSyncError(message) from exc
    except (URLError, TimeoutError, OSError, http.client.HTTPException) as exc:
        raise WorkbookSyncError(
            "Could not reach the dashboard upload API. Check the internet connection and dashboard URL."
        ) from exc
    if not result.get("ok"):
        raise WorkbookSyncError(result.get("error", "The dashboard rejected the workbook upload."))
    print(
        f"Dashboard updated: {result['rows']:,} records "
        f"(version {result['version'][:12]})."
    )


def allow_task_on_battery(task_name):
    """schtasks cannot set power options; new tasks otherwise skip runs on battery."""
    script = (
        f"$t = Get-ScheduledTask -TaskName '{task_name}'; "
        "$t.Settings.DisallowStartIfOnBatteries = $false; "
        "$t.Settings.StopIfGoingOnBatteries = $false; "
        "$t.Settings.StartWhenAvailable = $true; "
        "$t.Settings.ExecutionTimeLimit = 'PT30M'; "
        "Set-ScheduledTask -InputObject $t | Out-Null"
    )
    subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
        check=True, capture_output=True, text=True,
    )


def install_task():
    if not CONFIG_PATH.is_file():
        raise WorkbookSyncError("Configure the local sync before installing its scheduled task.")
    script_path = Path(__file__).resolve()
    interpreter = Path(sys.executable)
    if interpreter.with_name("pythonw.exe").is_file():
        interpreter = interpreter.with_name("pythonw.exe")
    action = f'"{interpreter}" "{script_path}" --sync'
    subprocess.run(
        [
            "schtasks.exe", "/Create", "/SC", "MINUTE", "/MO", "5",
            "/TN", TASK_NAME, "/TR", action, "/F", "/IT", "/RL", "LIMITED",
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    allow_task_on_battery(TASK_NAME)
    print("Scheduled workbook sync every 5 minutes while you are signed in.")


def remove_task():
    subprocess.run(
        ["schtasks.exe", "/Delete", "/TN", TASK_NAME, "/F"],
        check=True,
        capture_output=True,
        text=True,
    )
    print("Scheduled workbook sync removed.")


def main():
    parser = argparse.ArgumentParser(
        description="Sync a locally available Excel workbook to the offload dashboard."
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--configure", action="store_true", help="Save the workbook path and upload token.")
    mode.add_argument("--sync", action="store_true", help="Upload the configured workbook now.")
    mode.add_argument("--install-task", action="store_true", help="Schedule uploads every five minutes.")
    mode.add_argument("--remove-task", action="store_true", help="Remove the scheduled uploads.")
    parser.add_argument("--workbook", help="Local path to the synced or downloaded Excel workbook.")
    parser.add_argument("--endpoint", default=DEFAULT_ENDPOINT, help=argparse.SUPPRESS)
    args = parser.parse_args()

    try:
        if args.configure:
            if not args.workbook:
                parser.error("--workbook is required with --configure")
            configure(args.workbook, args.endpoint)
            sync_once()
        elif args.install_task:
            install_task()
        elif args.remove_task:
            remove_task()
        else:
            sync_once()
    except (OSError, subprocess.CalledProcessError, WorkbookSyncError) as exc:
        print(f"Workbook sync failed: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
