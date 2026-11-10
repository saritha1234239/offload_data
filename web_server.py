import json
import os
import socket
import subprocess
import threading
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

from workbook_sync import WorkbookAutoSync, WorkbookSyncError, fetch_graph_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_WORKBOOK_PATH = "https://criticalogindia-my.sharepoint.com/personal/saptarshi_maji_criticalog_com/Documents/Desktop/Midmile%20Management/LH%20performance%20tracker.xlsx?web=1"
DATA_PATH = os.path.join(BASE_DIR, "workbook-data.js")
WORKBOOK_REFRESH_SECONDS = int(os.environ.get("WORKBOOK_REFRESH_SECONDS", "300"))


def create_workbook_sync():
    graph_env = {
        "tenant ID": os.environ.get("MICROSOFT_TENANT_ID", "").strip(),
        "client ID": os.environ.get("MICROSOFT_CLIENT_ID", "").strip(),
        "client secret": os.environ.get("MICROSOFT_CLIENT_SECRET", ""),
        "drive ID": os.environ.get("MICROSOFT_GRAPH_DRIVE_ID", "").strip(),
        "file item ID": os.environ.get("MICROSOFT_GRAPH_ITEM_ID", "").strip(),
    }
    configured_graph_values = {
        label: value for label, value in graph_env.items() if value
    }
    if configured_graph_values:
        missing = [label for label, value in graph_env.items() if not value]
        if missing:
            missing_settings = ", ".join(
                {
                    "tenant ID": "MICROSOFT_TENANT_ID",
                    "client ID": "MICROSOFT_CLIENT_ID",
                    "client secret": "MICROSOFT_CLIENT_SECRET",
                    "drive ID": "MICROSOFT_GRAPH_DRIVE_ID",
                    "file item ID": "MICROSOFT_GRAPH_ITEM_ID",
                }[label]
                for label in missing
            )
            return WorkbookAutoSync(
                "",
                DATA_PATH,
                WORKBOOK_REFRESH_SECONDS,
                source="microsoft_graph",
                configuration_error=f"Missing required Render settings: {missing_settings}.",
            )

        return WorkbookAutoSync(
            "",
            DATA_PATH,
            WORKBOOK_REFRESH_SECONDS,
            fetcher=lambda: fetch_graph_workbook(
                tenant_id=graph_env["tenant ID"],
                client_id=graph_env["client ID"],
                client_secret=graph_env["client secret"],
                drive_id=graph_env["drive ID"],
                item_id=graph_env["file item ID"],
            ),
            source="microsoft_graph",
        )

    return WorkbookAutoSync(
        os.environ.get("WORKBOOK_DOWNLOAD_URL", ""),
        DATA_PATH,
        WORKBOOK_REFRESH_SECONDS,
    )


WORKBOOK_AUTO_SYNC = create_workbook_sync()


def resolve_workbook_path():
    configured = os.environ.get("WORKBOOK_PATH")
    if configured:
        return configured
    if DEFAULT_WORKBOOK_PATH.startswith(("http://", "https://")):
        return DEFAULT_WORKBOOK_PATH

    candidates = [
        r"C:\Users\MSaritha\OneDrive - CriticaLog India Pvt Ltd\Downloads\LH performance tracker (2).xlsx",
        r"C:\Users\MSaritha\OneDrive - CriticaLog India Pvt Ltd\Downloads\LH performance tracker.xlsx",
        os.path.join(BASE_DIR, "LH performance tracker (2).xlsx"),
        os.path.join(BASE_DIR, "LH performance tracker.xlsx"),
        os.path.join(BASE_DIR, "data", "LH performance tracker (2).xlsx"),
        os.path.join(BASE_DIR, "data", "LH performance tracker.xlsx"),
        DEFAULT_WORKBOOK_PATH,
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return DEFAULT_WORKBOOK_PATH


WORKBOOK_PATH = resolve_workbook_path()
PORT = int(os.environ.get("PORT", "8000"))
HOST = os.environ.get("HOST", "0.0.0.0")


def run_workbook_refresh(workbook_path=None):
    if WORKBOOK_AUTO_SYNC.configuration_error:
        return {
            "ok": False,
            "returnCode": 503,
            "output": WORKBOOK_AUTO_SYNC.configuration_error,
        }
    if WORKBOOK_AUTO_SYNC.configured:
        try:
            result = WORKBOOK_AUTO_SYNC.refresh_once()
            return {
                "ok": True,
                "returnCode": 0,
                "rows": result["rows"],
                "version": result["version"],
            }
        except WorkbookSyncError as exc:
            return {"ok": False, "returnCode": 502, "output": str(exc)}

    target_path = workbook_path or WORKBOOK_PATH
    if os.name != "nt":
        return {
            "ok": False,
            "workbookPath": None,
            "returnCode": 501,
            "output": "Workbook refresh requires Windows and Excel. This host serves the bundled workbook-data.js snapshot.",
        }

    script_path = os.path.join(BASE_DIR, "refresh-dashboard-data.ps1")
    command = [
        "powershell.exe",
        "-ExecutionPolicy",
        "Bypass",
        "-NonInteractive",
        "-File",
        script_path,
        "-Once",
        "-WorkbookPath",
        target_path,
    ]
    result = subprocess.run(command, capture_output=True, text=True)
    output = (result.stdout or "") + (result.stderr or "")
    return {
        "ok": result.returncode == 0,
        "workbookPath": target_path,
        "returnCode": result.returncode,
        "output": output.strip(),
    }


class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_GET(self):
        if self.path == "/" or self.path == "":
            self.path = "/dashboard.html"

        if self.path.startswith("/api/health"):
            payload = {
                "status": "ok",
                "port": PORT,
                "workbookAutoSync": WORKBOOK_AUTO_SYNC.status(),
            }
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(json.dumps(payload).encode("utf-8"))
            return

        if self.path.startswith("/api/refresh"):
            if WORKBOOK_AUTO_SYNC.configured and os.name != "nt":
                result = {
                    "ok": False,
                    "returnCode": 405,
                    "output": "Hosted workbook refresh runs on the configured schedule.",
                }
            else:
                result = run_workbook_refresh()
            status = 200 if result["ok"] else result["returnCode"] if result["returnCode"] in (405, 501, 502) else 500
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(json.dumps(result).encode("utf-8"))
            return

        if self.path.endswith((".html", ".js", ".css")):
            cache_bust = f"?cacheBust={int(time.time())}"
            if "cacheBust=" not in self.path:
                self.path = f"{self.path}{cache_bust}"

        return super().do_GET()

    def log_message(self, format, *args):
        print(f"[web] {self.address_string()} {format % args}")


def get_local_ip():
    host_name = socket.gethostname()
    try:
        return socket.gethostbyname(host_name)
    except OSError:
        return "127.0.0.1"


if __name__ == "__main__":
    print(f"Starting dashboard web service on http://{HOST}:{PORT}")
    print(f"Serving folder: {BASE_DIR}")
    if os.name == "nt":
        print(f"Workbook: {WORKBOOK_PATH}")
        try:
            run_workbook_refresh()
            print("Initial workbook refresh completed.")
        except Exception as exc:
            print(f"Initial refresh warning: {exc}")
    else:
        print("Serving the bundled workbook-data.js snapshot until workbook auto-sync succeeds.")

    httpd = ThreadingHTTPServer((HOST, PORT), DashboardHandler)
    WORKBOOK_AUTO_SYNC.start()
    print(f"Local network access: http://{get_local_ip()}:{PORT}")
    try:
        httpd.serve_forever()
    finally:
        WORKBOOK_AUTO_SYNC.stop()
