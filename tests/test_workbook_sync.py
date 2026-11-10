import json
import tempfile
import unittest
from datetime import datetime
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from workbook_sync import (
    WorkbookAutoSync,
    WorkbookSyncError,
    _date_variants,
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

        with patch("workbook_sync.urlopen", return_value=RedirectResponse(b"sign-in page")):
            with self.assertRaisesRegex(WorkbookSyncError, "requires sign-in") as raised:
                fetch_workbook("https://example.test/shared.xlsx")
        self.assertNotIn("private", str(raised.exception))

    def test_dot_separated_ambiguous_dates_keep_both_region_interpretations(self):
        self.assertEqual(
            _date_variants("5.11.2026", "South"),
            ["2026-11-05", "2026-05-11"],
        )


if __name__ == "__main__":
    unittest.main()
