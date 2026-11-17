import shutil
import tempfile
import unittest
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

from lccs_tracker import LCCS_TRACKER_QUERY, NDC_REMARK_OPTIONS, TRACKER_COLUMNS
from lccs_tracker_job import TrackerJobError, mysql_query, refresh
from workbook_sync import WorkbookSyncError, _date_variants, parse_workbook


def make_template(path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "LH Performance Tracker"
    sheet["D2"] = "LH Performance Tracker"
    for index, header in enumerate(TRACKER_COLUMNS):
        sheet.cell(4, 4 + index).value = header
    validation = DataValidation(type="list", formula1="=NDC_Remarks_Options")
    validation.add("X5:X5000")
    sheet.add_data_validation(validation)
    workbook.save(path)


def source(shipment_id, mawb, flight="LH 1", day="2026-11-16", etd="08:00:00"):
    return {
        "XAWBID": shipment_id, "XAWSDPRDT": day, "XAWBOOKLOC": "BOM", "XAWPKUPLOC": "DEL",
        "mawb_no": mawb, "VENDORNM": "Vendor", "SECTORMODE": "D", "XAWBMODE": "A",
        "XAWFLTNO": flight, "XAWSDPRTIME": etd, "XAWPACKEGE": 3, "XAWBCDWGT": 10.5, "is_nfo": "",
        "org_region": "WEST",
    }


class TrackerJobTests(unittest.TestCase):
    def setUp(self):
        self.folder = Path(tempfile.mkdtemp())
        self.path = self.folder / "tracker.xlsx"
        make_template(self.path)

    def tearDown(self):
        shutil.rmtree(self.folder)

    def sheet(self):
        return openpyxl.load_workbook(self.path)["LH Performance Tracker"]

    def test_first_refresh_writes_rows_with_hidden_id(self):
        result = refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                         source_rows=[source(2, "M2"), source(1, "M1")])
        self.assertEqual(result, {"added": 2, "updated": 0, "total": 2})
        sheet = self.sheet()
        self.assertEqual(sheet["C4"].value, "LCCS_ID")
        self.assertTrue(sheet.column_dimensions["C"].hidden)
        self.assertEqual([sheet["C5"].value, sheet["C6"].value], ["1", "2"])
        self.assertEqual(sheet["H5"].value, "M1")
        self.assertEqual(sheet["O5"].value, "Prime")

    def test_refresh_keeps_team_entries_and_updates_lccs_columns(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook["LH Performance Tracker"]
        sheet["N5"] = "11:30"
        sheet["S5"] = 2
        sheet["X5"] = NDC_REMARK_OPTIONS[4]
        sheet["Y5"] = "2026-11-16"
        sheet["AA5"] = "Closed"
        workbook.save(self.path)

        result = refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                         source_rows=[source(1, "M1", flight="LH 9"), source(3, "M3")])
        self.assertEqual(result, {"added": 1, "updated": 1, "total": 2})
        sheet = self.sheet()
        self.assertEqual(sheet["L5"].value, "LH 9")
        self.assertEqual(
            [sheet["N5"].value, sheet["S5"].value, sheet["X5"].value, sheet["Y5"].value, sheet["AA5"].value],
            ["11:30", 2, NDC_REMARK_OPTIONS[4], "2026-11-16", "Closed"],
        )

    def test_rows_outside_refreshed_window_are_kept(self):
        refresh(self.path, date(2026, 11, 15), date(2026, 11, 15),
                source_rows=[source(1, "OLD", day="2026-11-15")])
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(2, "NEW")])
        sheet = self.sheet()
        self.assertEqual([sheet["H5"].value, sheet["H6"].value], ["OLD", "NEW"])

    def test_invalid_ndc_typed_in_workbook_stops_refresh_without_overwriting(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        workbook = openpyxl.load_workbook(self.path)
        workbook["LH Performance Tracker"]["X5"] = "typo"
        workbook.save(self.path)
        with self.assertRaisesRegex(TrackerJobError, "LCCS_ID 1"):
            refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        self.assertEqual(self.sheet()["X5"].value, "typo")

    def test_rows_typed_without_id_are_rejected(self):
        workbook = openpyxl.load_workbook(self.path)
        workbook["LH Performance Tracker"]["H5"] = "manual row"
        workbook.save(self.path)
        with self.assertRaisesRegex(TrackerJobError, "no LCCS_ID"):
            refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[])

    def test_region_is_written_to_hidden_column_b(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        sheet = self.sheet()
        self.assertEqual([sheet["B4"].value, sheet["B5"].value], ["Org_Region", "West"])
        self.assertTrue(sheet.column_dimensions["B"].hidden)

    def test_dashboard_parser_reads_refreshed_tracker(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                source_rows=[source(1, "M1"), source(2, "M2", etd="13:30:00")])
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook["LH Performance Tracker"]
        sheet["S5"] = 1
        sheet["T5"] = 2
        sheet["X5"] = NDC_REMARK_OPTIONS[12]
        workbook.save(self.path)

        data = parse_workbook(self.path.read_bytes())
        first, second = data["records"]
        self.assertEqual(data["sources"][0]["name"], "LH Performance Tracker")
        self.assertEqual(
            (first["date"], first["region"], first["lane"], first["mawb"], first["bags"],
             first["offloadedBags"], first["impactedCons"], first["ndcRemarks"], first["sourceRow"]),
            ("2026-11-16", "West", "BOM-DEL", "M1", 3, 1, 2, NDC_REMARK_OPTIONS[12], 5),
        )
        self.assertAlmostEqual(first["etd"], 8 / 24)
        self.assertAlmostEqual(second["etd"], 13.5 / 24)
        self.assertEqual(second["offloadedBags"], 0)

    def test_datetime_departure_dates_are_stored_and_read_as_dates(self):
        from datetime import datetime
        refresh(self.path, date(2026, 10, 20), date(2026, 10, 20),
                source_rows=[source(1, "M1", day=datetime(2026, 10, 20, 0, 0)),
                             source(2, "M2", day="2026-10-21T00:00:00")])
        sheet = self.sheet()
        self.assertEqual([sheet["D5"].value, sheet["D6"].value], ["2026-10-20", "2026-10-21"])
        self.assertEqual(_date_variants("2026-10-20T00:00:00", "North"), ["2026-10-20"])

    def test_dashboard_parser_rejects_tracker_without_rows(self):
        with self.assertRaisesRegex(WorkbookSyncError, "no dashboard rows"):
            parse_workbook(self.path.read_bytes())

    def test_query_placeholders_convert_to_pymysql_style(self):
        converted = mysql_query(LCCS_TRACKER_QUERY)
        self.assertIn("%(from_date)s", converted)
        self.assertIn("%(testing_client_code)s", converted)
        self.assertNotIn(":hub_code", converted)


if __name__ == "__main__":
    unittest.main()
