import json
import os
import sys
import tempfile
import threading
import types
import unittest
from datetime import datetime
from io import BytesIO
from pathlib import Path
from http.server import ThreadingHTTPServer
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from unittest.mock import patch

from openpyxl import Workbook

import web_server
import sync_workbook_to_dashboard as local_sync
from workbook_sync import (
    WorkbookAutoSync,
    WorkbookSyncError,
    _GraphRedirectHandler,
    _date_variants,
    fetch_graph_workbook,
    fetch_workbook,
    parse_workbook,
    refresh_from_url,
    workbook_data_script,
)


HEADERS = (
    "Date", "Org", "Dest", "Mawb", "Vendor_name", "Flight_Number", "ETD", "ETA",
    "Bags", "CD weight", "Total Cons", "Offloaded bags", "Impacted cons",
    "Next_flight", "ETD", "ETA", "NDC_Remarks", "Remarks",
)


def workbook_fixture():
    workbook = Workbook()
    workbook.remove(workbook.active)
    for region in ("North", "South", "West", "East"):
        sheet = workbook.create_sheet(f"Daily - {region} - New")
        sheet.append(HEADERS)
        sheet.append(HEADERS)
        sheet.append((
            "5/11/2026", "BOM", "DEL", "ABC123", "AIR INDIA", "AI 101",
            "08:00", "09:30", 10, 2.5, 8, 3, 2, "AI 201", "10:00",
            "11:30", "NDC note", "Remark",
        ))
        sheet.append((
            datetime(2026, 11, 6), "BLR", "DEL", "DEF456", "AKASA", "QP 10",
            None, None, 4, None, None, 1, None, None, None, None, None, None,
        ))
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


class WorkbookSyncTests(unittest.TestCase):
    def _start_upload_server(self):
        server = ThreadingHTTPServer(("127.0.0.1", 0), web_server.DashboardHandler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)
        return f"http://127.0.0.1:{server.server_port}"

    def test_parses_all_regions_skips_duplicate_headers_and_resolves_ambiguous_dates(self):
        data = parse_workbook(workbook_fixture())
        self.assertEqual(len(data["records"]), 8)
        first_records = {record["region"]: record for record in data["records"] if record["sourceRow"] == 3}
        self.assertEqual(first_records["North"]["date"], "2026-05-11")
        self.assertEqual(first_records["South"]["date"], "2026-11-05")
        self.assertEqual(first_records["West"]["date"], "2026-11-05")
        self.assertEqual(first_records["East"]["date"], "2026-11-05")
        self.assertEqual(first_records["South"]["dateVariants"], ["2026-11-05", "2026-05-11"])
        self.assertEqual(first_records["South"]["lane"], "BOM-DEL")
        self.assertEqual(first_records["South"]["etd"], 8 / 24)
        self.assertEqual(first_records["South"]["sourceValues"]["raw_1_Date"], "5/11/2026")
        dated_cell_record = next(
            record for record in data["records"]
            if record["region"] == "South" and record["sourceRow"] == 4
        )
        self.assertEqual(dated_cell_record["sourceValues"]["raw_1_Date"], 46332)
        self.assertEqual(sum(source["shipments"] for source in data["sources"]), 8)

    def test_serializes_the_dashboard_global_data_shape(self):
        data = parse_workbook(workbook_fixture())
        version, script = workbook_data_script(data)
        self.assertEqual(len(version), 64)
        lines = script.splitlines()
        self.assertEqual(json.loads(lines[1].split(" = ", 1)[1].rstrip(";")), data["columns"])
        self.assertEqual(json.loads(lines[2].split(" = ", 1)[1].rstrip(";")), data["records"])
        self.assertIn("window.offloadWorkbookSources", lines[3])

    def test_missing_expected_sheet_rejects_workbook(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Daily - North - New"
        sheet.append(HEADERS)
        stream = BytesIO()
        workbook.save(stream)
        with self.assertRaisesRegex(WorkbookSyncError, "Daily-South-new"):
            parse_workbook(stream.getvalue())

    def test_failed_refresh_keeps_the_last_good_export(self):
        with patch("workbook_sync.fetch_workbook", return_value=workbook_fixture()):
            with tempfile.TemporaryDirectory() as directory:
                destination = Path(directory) / "workbook-data.js"
                destination.write_text("last-good", encoding="utf-8")
                result = refresh_from_url("https://example.test/workbook.xlsx", str(destination))
                self.assertEqual(result["rows"], 8)
                self.assertIn("window.offloadWorkbookRecords", destination.read_text(encoding="utf-8"))

                with patch("workbook_sync.fetch_workbook", return_value=b"not an Excel file"):
                    with self.assertRaises(WorkbookSyncError):
                        refresh_from_url("https://example.test/workbook.xlsx", str(destination))
                self.assertIn("window.offloadWorkbookRecords", destination.read_text(encoding="utf-8"))

    def test_auto_sync_requires_a_configured_public_source(self):
        sync = WorkbookAutoSync("", "unused.js", 300)
        self.assertFalse(sync.configured)
        with self.assertRaisesRegex(WorkbookSyncError, "not configured"):
            sync.refresh_once()

    def test_sign_in_redirect_is_rejected_without_exposing_its_url(self):
        class RedirectResponse(BytesIO):
            headers = {}

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                self.close()

            def geturl(self):
                return "https://login.microsoftonline.com/tenant/signin?token=private"

        with patch(
            "workbook_sync.urlopen",
            side_effect=lambda *_args, **_kwargs: RedirectResponse(b"sign-in page"),
        ):
            with self.assertRaisesRegex(WorkbookSyncError, "redirected the dashboard to Microsoft sign-in") as raised:
                fetch_workbook("https://example.test/shared.xlsx")
        self.assertNotIn("token=private", str(raised.exception))

    def test_public_share_link_uses_direct_download_and_accepts_only_excel_content(self):
        class WorkbookResponse(BytesIO):
            headers = {}

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                self.close()

            def geturl(self):
                return "https://example.test/download/workbook.xlsx"

        requests = []

        def open_request(request, timeout):
            requests.append(request)
            return WorkbookResponse(workbook_fixture())

        with patch("workbook_sync.urlopen", side_effect=open_request):
            content = fetch_workbook("https://example.test/shared.xlsx?e=public")

        self.assertEqual(content, workbook_fixture())
        self.assertEqual(len(requests), 1)
        self.assertIn("e=public", requests[0].full_url)
        self.assertIn("download=1", requests[0].full_url)
        self.assertIn("Mozilla/5.0", requests[0].get_header("User-agent"))

    def test_web_page_response_is_reported_as_not_a_workbook(self):
        class WebPageResponse(BytesIO):
            headers = {}

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                self.close()

            def geturl(self):
                return "https://example.test/shared.xlsx"

        with patch(
            "workbook_sync.urlopen",
            side_effect=lambda *_args, **_kwargs: WebPageResponse(b"<html>Sign in</html>"),
        ):
            with self.assertRaisesRegex(WorkbookSyncError, "web page instead of an Excel file"):
                fetch_workbook("https://example.test/shared.xlsx")

    def test_authenticated_local_upload_refreshes_dashboard_data_and_status(self):
        with tempfile.TemporaryDirectory() as directory:
            data_path = Path(directory) / "workbook-data.js"
            sync = WorkbookAutoSync("", str(data_path), 300, source="local_upload")
            with (
                patch.object(web_server, "UPLOAD_TOKEN", "test-upload-token-value-at-least-32"),
                patch.object(web_server, "UPLOAD_TOKEN_CONFIGURED", True),
                patch.object(web_server, "DATA_PATH", str(data_path)),
                patch.object(web_server, "WORKBOOK_AUTO_SYNC", sync),
            ):
                server_url = self._start_upload_server()
                request = Request(
                    f"{server_url}/api/upload",
                    data=workbook_fixture(),
                    headers={
                        "Authorization": "Bearer test-upload-token-value-at-least-32",
                        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    },
                    method="POST",
                )
                with urlopen(request, timeout=5) as response:
                    result = json.loads(response.read().decode("utf-8"))
                with urlopen(f"{server_url}/api/health", timeout=5) as response:
                    health = json.loads(response.read().decode("utf-8"))

            self.assertTrue(result["ok"])
            self.assertEqual(result["rows"], 8)
            self.assertIn("window.offloadWorkbookRecords", data_path.read_text(encoding="utf-8"))
            self.assertEqual(health["workbookAutoSync"]["source"], "local_upload")
            self.assertEqual(health["workbookAutoSync"]["rows"], 8)
            self.assertIsNotNone(health["workbookAutoSync"]["lastSuccessAt"])
            self.assertTrue(health["workbookUpload"]["enabled"])

    def test_local_upload_rejects_bad_token_and_invalid_workbook_without_replacing_data(self):
        with tempfile.TemporaryDirectory() as directory:
            data_path = Path(directory) / "workbook-data.js"
            data_path.write_text("last-good", encoding="utf-8")
            sync = WorkbookAutoSync("", str(data_path), 300, source="local_upload")
            with (
                patch.object(web_server, "UPLOAD_TOKEN", "test-upload-token-value-at-least-32"),
                patch.object(web_server, "UPLOAD_TOKEN_CONFIGURED", True),
                patch.object(web_server, "DATA_PATH", str(data_path)),
                patch.object(web_server, "WORKBOOK_AUTO_SYNC", sync),
            ):
                server_url = self._start_upload_server()
                bad_token_request = Request(
                    f"{server_url}/api/upload",
                    data=workbook_fixture(),
                    headers={
                        "Authorization": "Bearer wrong-token",
                        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    },
                    method="POST",
                )
                with self.assertRaises(HTTPError) as unauthorized:
                    urlopen(bad_token_request, timeout=5)
                self.assertEqual(unauthorized.exception.code, 401)

                invalid_workbook_request = Request(
                    f"{server_url}/api/upload",
                    data=b"not an Excel workbook",
                    headers={
                        "Authorization": "Bearer test-upload-token-value-at-least-32",
                        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    },
                    method="POST",
                )
                with self.assertRaises(HTTPError) as invalid_workbook:
                    urlopen(invalid_workbook_request, timeout=5)
                self.assertEqual(invalid_workbook.exception.code, 400)

            self.assertEqual(data_path.read_text(encoding="utf-8"), "last-good")

    def test_local_sync_configuration_performs_an_authenticated_upload(self):
        class UploadResponse(BytesIO):
            def __enter__(self):
                return self

            def __exit__(self, *_args):
                self.close()

        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "tracker.xlsx"
            workbook_path.write_bytes(workbook_fixture())
            config_path = Path(directory) / "sync-config.json"
            requests = []

            def open_request(request, timeout):
                requests.append(request)
                return UploadResponse(json.dumps({
                    "ok": True,
                    "rows": 8,
                    "version": "a" * 64,
                }).encode("utf-8"))

            with (
                patch.object(local_sync, "CONFIG_DIRECTORY", Path(directory)),
                patch.object(local_sync, "CONFIG_PATH", config_path),
                patch(
                    "sync_workbook_to_dashboard.getpass.getpass",
                    side_effect=["test-upload-token-value-at-least-32"] * 2,
                ),
                patch("sync_workbook_to_dashboard.urlopen", side_effect=open_request),
            ):
                local_sync.configure(workbook_path, "https://dashboard.example.test/api/upload")
                local_sync.sync_once()

            self.assertEqual(len(requests), 1)
            self.assertEqual(
                requests[0].get_header("Authorization"),
                "Bearer test-upload-token-value-at-least-32",
            )
            self.assertEqual(requests[0].data, workbook_fixture())
            saved_config = json.loads(config_path.read_text(encoding="utf-8"))
            self.assertEqual(saved_config["workbookPath"], str(workbook_path.resolve()))

    def test_graph_download_uses_bearer_token_and_file_scoped_endpoint(self):
            class FakeApplication:
                def __init__(self, client_id, authority, client_credential):
                    self.client_id = client_id
                    self.authority = authority
                    self.client_credential = client_credential

                def acquire_token_for_client(self, scopes):
                    self.scopes = scopes
                    return {"access_token": "test-access-token"}

            fake_msal = types.SimpleNamespace(
                ConfidentialClientApplication=FakeApplication
            )

            class Response(BytesIO):
                headers = {"Content-Length": "10"}

                def __enter__(self):
                    return self

                def __exit__(self, *_args):
                    self.close()

            response = Response(b"workbook!!")
            captured_requests = []

            def open_request(request, timeout):
                captured_requests.append(request)
                return response

            with patch.dict(sys.modules, {"msal": fake_msal}):
                opener = types.SimpleNamespace(open=open_request)
                with patch("workbook_sync.build_opener", return_value=opener):
                    content = fetch_graph_workbook(
                        "tenant-id", "client-id", "secret", "drive/id", "item id"
                    )

            self.assertEqual(content, b"workbook!!")
            request = captured_requests[0]
            self.assertEqual(
                request.full_url,
                "https://graph.microsoft.com/v1.0/drives/drive%2Fid/items/item%20id/content",
            )
            self.assertEqual(request.get_header("Authorization"), "Bearer test-access-token")
            self.assertNotIn("secret", request.headers.values())

    def test_graph_redirect_does_not_forward_bearer_token_to_download_host(self):
            request = Request(
                "https://graph.microsoft.com/v1.0/drives/drive/items/item/content",
                headers={"Authorization": "Bearer test-access-token"},
            )
            redirected = _GraphRedirectHandler().redirect_request(
                request,
                None,
                302,
                "Found",
                {},
                "https://download.sharepoint.com/file",
            )

            self.assertEqual(
                redirected.full_url, "https://download.sharepoint.com/file"
            )
            self.assertIsNone(redirected.get_header("Authorization"))

    def test_graph_authentication_failure_does_not_expose_credentials(self):
            class FakeApplication:
                def __init__(self, *_args, **_kwargs):
                    pass

                def acquire_token_for_client(self, scopes):
                    return {"error_description": "invalid client secret: private-secret"}

            fake_msal = types.SimpleNamespace(
                ConfidentialClientApplication=FakeApplication
            )
            with patch.dict(sys.modules, {"msal": fake_msal}):
                with self.assertRaisesRegex(WorkbookSyncError, "authentication failed") as raised:
                    fetch_graph_workbook("tenant", "client", "private-secret", "drive", "item")
            self.assertNotIn("private-secret", str(raised.exception))

    def test_graph_auto_sync_uses_shared_atomic_export_path(self):
            with tempfile.TemporaryDirectory() as directory:
                destination = Path(directory) / "workbook-data.js"
                sync = WorkbookAutoSync(
                    "",
                    str(destination),
                    300,
                    fetcher=workbook_fixture,
                    source="microsoft_graph",
                )
                result = sync.refresh_once()

                self.assertEqual(result["rows"], 8)
                self.assertEqual(sync.status()["source"], "microsoft_graph")
                self.assertEqual(sync.status()["rows"], 8)
                self.assertIn("window.offloadWorkbookRecords", destination.read_text(encoding="utf-8"))

    def test_partial_graph_configuration_reports_missing_settings(self):
        with patch.dict(os.environ, {"MICROSOFT_CLIENT_ID": "client-id"}, clear=True):
                sync = web_server.create_workbook_sync()

        self.assertFalse(sync.configured)
        self.assertIn("MICROSOFT_TENANT_ID", sync.status()["configurationError"])
        self.assertNotIn("client-id", str(sync.status()))

    def test_complete_graph_configuration_wires_private_fetcher_without_exposing_secret(self):
        settings = {
                "MICROSOFT_TENANT_ID": "tenant-id",
                "MICROSOFT_CLIENT_ID": "client-id",
                "MICROSOFT_CLIENT_SECRET": "private-secret",
                "MICROSOFT_GRAPH_DRIVE_ID": "drive-id",
                "MICROSOFT_GRAPH_ITEM_ID": "item-id",
        }
        with patch.dict(os.environ, settings, clear=True):
                with patch("web_server.fetch_graph_workbook", return_value=b"workbook") as fetch:
                    sync = web_server.create_workbook_sync()
                    self.assertEqual(sync.fetcher(), b"workbook")

        self.assertTrue(sync.configured)
        fetch.assert_called_once_with(
                tenant_id="tenant-id",
                client_id="client-id",
                client_secret="private-secret",
                drive_id="drive-id",
                item_id="item-id",
        )
        self.assertNotIn("private-secret", str(sync.status()))

    def test_dot_separated_ambiguous_dates_keep_both_region_interpretations(self):
        self.assertEqual(
            _date_variants("5.11.2026", "South"),
            ["2026-11-05", "2026-05-11"],
        )


if __name__ == "__main__":
    unittest.main()
