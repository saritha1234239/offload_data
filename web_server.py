import json
import os
import socket
import subprocess
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_WORKBOOK_PATH = "https://criticalogindia-my.sharepoint.com/personal/saptarshi_maji_criticalog_com/Documents/Desktop/Midmile%20Management/LH%20performance%20tracker.xlsx?web=1"


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
            payload = {"status": "ok", "port": PORT}
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(json.dumps(payload).encode("utf-8"))
            return

        if self.path.startswith("/api/refresh"):
            result = run_workbook_refresh()
            status = 200 if result["ok"] else 501 if result["returnCode"] == 501 else 500
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
        print("Serving the bundled workbook-data.js snapshot; workbook refresh is unavailable on this host.")

    httpd = ThreadingHTTPServer((HOST, PORT), DashboardHandler)
    print(f"Local network access: http://{get_local_ip()}:{PORT}")
    httpd.serve_forever()
