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
        "XAWFLTNO": flight, "XAWSDPRTIME": etd, "XAWSARVTIME": "10:15:00", "total_cons": 4, "XAWPACKEGE": 3, "XAWBCDWGT": 10.5, "is_nfo": "",
        "org_region": "WEST",
    }


class TrackerJobTests(unittest.TestCase):
    def setUp(self):
        self.folder = Path(tempfile.mkdtemp())
        self.path = self.folder / "tracker.xlsx"
        make_template(self.path)

    def tearDown(self):
        shutil.rmtree(self.folder)

    def sheet(self, name="West"):
        return openpyxl.load_workbook(self.path)[name]

    def test_first_refresh_writes_rows_with_hidden_id(self):
        result = refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                         source_rows=[source(2, "M2"), source(1, "M1")])
        self.assertEqual(result, {"added": 2, "updated": 0, "removed": 0, "total": 2})
        sheet = self.sheet()
        self.assertEqual(sheet["C1"].value, "LCCS_ID")
        self.assertTrue(sheet.column_dimensions["C"].hidden)
        self.assertEqual([sheet["C2"].value, sheet["C3"].value], ["1", "2"])
        self.assertEqual(sheet["H2"].value, "M1")
        self.assertEqual(sheet["O2"].value, "Prime")

    def test_refresh_keeps_team_entries_and_updates_lccs_columns(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook["West"]
        sheet["S2"] = 2
        sheet["X2"] = NDC_REMARK_OPTIONS[4]
        sheet["AA2"] = "Closed"
        workbook.save(self.path)

        result = refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                         source_rows=[source(1, "M1", flight="LH 9"), source(3, "M3")])
        self.assertEqual(result, {"added": 1, "updated": 1, "removed": 0, "total": 2})
        sheet = self.sheet()
        self.assertEqual(sheet["L2"].value, "LH 9")
        self.assertEqual(
            [sheet["N2"].value, sheet["R2"].value, sheet["S2"].value, sheet["X2"].value, sheet["AA2"].value],
            ["10:15:00", 4, 2, NDC_REMARK_OPTIONS[4], "Closed"],
        )

    def test_rows_older_than_the_window_are_removed(self):
        refresh(self.path, date(2026, 11, 14), date(2026, 11, 15),
                source_rows=[source(1, "OLD", day="2026-11-14"), source(3, "KEEP", day="2026-11-15")])
        result = refresh(self.path, date(2026, 11, 15), date(2026, 11, 16), source_rows=[source(2, "NEW")])
        self.assertEqual(result["removed"], 1)
        sheet = self.sheet()
        self.assertEqual([sheet["H2"].value, sheet["H3"].value, sheet["H4"].value], ["KEEP", "NEW", None])

    def test_invalid_ndc_typed_in_workbook_stops_refresh_without_overwriting(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        workbook = openpyxl.load_workbook(self.path)
        workbook["West"]["X2"] = "typo"
        workbook.save(self.path)
        with self.assertRaisesRegex(TrackerJobError, "LCCS_ID 1"):
            refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        self.assertEqual(self.sheet()["X2"].value, "typo")

    def test_rows_typed_without_id_are_rejected(self):
        workbook = openpyxl.load_workbook(self.path)
        workbook["LH Performance Tracker"]["H5"] = "manual row"
        workbook.save(self.path)
        with self.assertRaisesRegex(TrackerJobError, "no LCCS_ID"):
            refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[])

    def test_region_is_written_to_hidden_column_b(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        sheet = self.sheet()
        self.assertEqual([sheet["B1"].value, sheet["B2"].value], ["Org_Region", "West"])
        self.assertTrue(sheet.column_dimensions["B"].hidden)

    def test_dashboard_parser_reads_refreshed_tracker(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                source_rows=[source(1, "M1"), source(2, "M2", etd="13:30:00")])
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook["West"]
        sheet["S2"] = 1
        sheet["T2"] = 2
        sheet["X2"] = NDC_REMARK_OPTIONS[12]
        workbook.save(self.path)

        data = parse_workbook(self.path.read_bytes())
        first, second = data["records"]
        self.assertEqual({s["name"]: s["rows"] for s in data["sources"]}["West"], 2)
        self.assertEqual(
            (first["date"], first["region"], first["lane"], first["mawb"], first["bags"],
             first["offloadedBags"], first["impactedCons"], first["ndcRemarks"], first["sourceRow"]),
            ("2026-11-16", "West", "BOM-DEL", "M1", 3, 1, 2, NDC_REMARK_OPTIONS[12], 2),
        )
        self.assertAlmostEqual(first["etd"], 8 / 24)
        self.assertAlmostEqual(second["etd"], 13.5 / 24)
        self.assertEqual(second["offloadedBags"], 0)

    def test_datetime_departure_dates_are_stored_and_read_as_dates(self):
        from datetime import datetime
        refresh(self.path, date(2026, 10, 20), date(2026, 10, 21),
                source_rows=[source(1, "M1", day=datetime(2026, 10, 20, 0, 0)),
                             source(2, "M2", day="2026-10-21T00:00:00")])
        sheet = self.sheet()
        self.assertEqual([sheet["D2"].value, sheet["D3"].value], ["2026-10-20", "2026-10-21"])
        self.assertEqual(_date_variants("2026-10-20T00:00:00", "North"), ["2026-10-20"])

    def test_dashboard_parser_rejects_tracker_without_rows(self):
        with self.assertRaisesRegex(WorkbookSyncError, "no dashboard rows"):
            parse_workbook(self.path.read_bytes())

    def test_rows_are_split_into_four_region_sheets(self):
        rows = [dict(source(1, "W1"), org_region="WEST"), dict(source(2, "N1"), org_region="NORTH"),
                dict(source(3, "S1"), org_region="SOUTH"), dict(source(4, "S2"), org_region="SOUTH")]
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=rows)
        workbook = openpyxl.load_workbook(self.path)
        self.assertEqual(workbook.sheetnames, ["North", "West", "East", "South"])
        self.assertEqual(workbook["South"]["D1"].value, "Departure_Date")  # title rows removed
        self.assertEqual(workbook["South"].freeze_panes, "D2")
        self.assertEqual([workbook["South"]["H2"].value, workbook["South"]["H3"].value], ["S1", "S2"])
        self.assertEqual(workbook["North"]["H2"].value, "N1")
        self.assertIsNone(workbook["East"]["H2"].value)
        self.assertEqual(workbook["East"]["D1"].value, "Departure_Date")
        self.assertTrue(workbook["East"].column_dimensions["C"].hidden)
        for name in ("North", "West", "East", "South"):
            validations = {v.formula1: str(v.sqref) for v in workbook[name].data_validations.dataValidation}
            self.assertEqual(validations, {"=NDC_Remarks_Options": "X2:X5000", '"Ontime,Delayed"': "AA2:AA5000"}, name)

        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=rows)
        validations = openpyxl.load_workbook(self.path)["South"].data_validations.dataValidation
        self.assertEqual(len(validations), 2)  # not duplicated on later refreshes

        data = parse_workbook(self.path.read_bytes())
        self.assertEqual(sorted(r["region"] for r in data["records"]), ["North", "South", "South", "West"])
        self.assertEqual([s["name"] for s in data["sources"]], ["North", "West", "East", "South"])

    def test_team_entries_survive_split_and_region_change(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        workbook = openpyxl.load_workbook(self.path)
        workbook["West"]["S2"] = 4
        workbook.save(self.path)
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                source_rows=[dict(source(1, "M1"), org_region="EAST")])
        workbook = openpyxl.load_workbook(self.path)
        self.assertIsNone(workbook["West"]["H2"].value)
        self.assertEqual([workbook["East"]["H2"].value, workbook["East"]["S2"].value], ["M1", 4])

    def test_unknown_region_goes_to_unassigned_sheet(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                source_rows=[dict(source(1, "M1"), org_region=None)])
        self.assertEqual(self.sheet("Unassigned")["H2"].value, "M1")
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                source_rows=[dict(source(1, "M1"), org_region="WEST")])
        self.assertNotIn("Unassigned", openpyxl.load_workbook(self.path).sheetnames)

    def test_region_sheets_are_protected_except_team_columns(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        sheet = self.sheet()
        self.assertTrue(sheet.protection.sheet)
        self.assertTrue(sheet.protection.deleteRows)
        self.assertFalse(sheet.protection.autoFilter)
        self.assertTrue(sheet["H2"].protection.locked)  # LCCS column (MAWB)
        self.assertFalse(sheet["S2"].protection.locked)  # Offloaded bags
        self.assertFalse(sheet["AA2"].protection.locked)  # Status

    def test_rows_added_above_the_headers_are_parsed_then_removed(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        workbook = openpyxl.load_workbook(self.path)
        workbook["West"].insert_rows(1, 2)  # an editor added a title above the headers
        workbook["West"]["D1"] = "My title"
        workbook["West"]["S4"] = 2  # team entry on M1, now in row 4
        workbook.save(self.path)
        data = parse_workbook(self.path.read_bytes())
        self.assertEqual([(r["mawb"], r["sourceRow"]) for r in data["records"]], [("M1", 4)])

        result = refresh(self.path, date(2026, 11, 16), date(2026, 11, 16),
                         source_rows=[source(1, "M1"), source(2, "M2")])
        self.assertEqual(result["added"], 1)
        sheet = self.sheet()
        self.assertEqual([sheet["D1"].value, sheet["H2"].value, sheet["S2"].value, sheet["H3"].value],
                         ["Departure_Date", "M1", 2, "M2"])

    def test_duplicate_ids_across_sheets_stop_refresh(self):
        refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])
        workbook = openpyxl.load_workbook(self.path)
        workbook["North"]["C5"] = "1"
        workbook["North"]["H5"] = "M1"
        workbook.save(self.path)
        with self.assertRaisesRegex(TrackerJobError, "more than once"):
            refresh(self.path, date(2026, 11, 16), date(2026, 11, 16), source_rows=[source(1, "M1")])

    def test_query_placeholders_convert_to_pymysql_style(self):
        converted = mysql_query(LCCS_TRACKER_QUERY)
        self.assertIn("%(from_date)s", converted)
        self.assertIn("%(testing_client_code)s", converted)
        self.assertNotIn(":hub_code", converted)


if __name__ == "__main__":
    unittest.main()
