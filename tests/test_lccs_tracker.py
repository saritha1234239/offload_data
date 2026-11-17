import unittest
from datetime import date, time
import csv
from io import StringIO

from lccs_tracker import (
    LCCS_TRACKER_QUERY,
    NDC_REMARK_OPTIONS,
    TRACKER_COLUMNS,
    TrackerDataError,
    map_lccs_row,
    map_lccs_rows,
    tracker_csv,
)


class LccsTrackerTests(unittest.TestCase):
    def test_query_uses_the_php_source_tables_and_join_keys(self):
        self.assertIn("FROM xawbhdr awb", LCCS_TRACKER_QUERY)
        self.assertIn("hub.MHBHUBCD = awb.XAWBOOKLOC", LCCS_TRACKER_QUERY)
        self.assertNotIn("XUHLUSRCD", LCCS_TRACKER_QUERY)
        self.assertIn("awb.XAWBMODE = 'A'", LCCS_TRACKER_QUERY)
        self.assertIn("detail.XAIRWAYBILLNO = awb.XAWBILLNO", LCCS_TRACKER_QUERY)
        self.assertIn("awb.XAWBID AS XAWBID", LCCS_TRACKER_QUERY)

    def test_maps_source_fields_into_exact_report_order_and_preserves_manual_values(self):
        record = map_lccs_row(
            {
                "XAWBID": 42,
                "XAWSDPRDT": date(2026, 11, 16),
                "XAWBOOKLOC": "BOM1",
                "XAWPKUPLOC": "DEL2",
                "mawb_no": "123-45678901",
                "VENDORNM": "Vendor",
                "SECTORMODE": "D",
                "XAWBMODE": "A",
                "XAWFLTNO": "NA",
                "XAWFLTNO2": "LH 761",
                "XAWSDPRTIME": time(9, 30),
                "XAWPACKEGE": 12,
                "XAWBCDWGT": 86.5,
                "is_nfo": "NFO",
            },
            {
                "offloaded_bags": 2,
                "impacted_cons": "3   ",
                "next_flight": "LH 762",
                "revised_etd": "11:00",
                "revised_eta": "14:00",
                "ndc_remarks": NDC_REMARK_OPTIONS[0],
                "status": "Delayed",
            },
        )

        self.assertEqual(len(record["values"]), 24)
        self.assertEqual(record["shipment_id"], "42")
        self.assertEqual(
            record["values"],
            [
                "2026-11-16", "BOM", "DEL", "BOM-DEL", "123-45678901",
                "Vendor", "NFO", "Direct", "LH 761", "09:30:00", "",
                "Lean 1", 12, 86.5, "", 2, "3", "LH 762", "11:00",
                "14:00", NDC_REMARK_OPTIONS[0], "", "", "Delayed",
            ],
        )
        self.assertEqual(len(TRACKER_COLUMNS), 24)
        self.assertEqual(TRACKER_COLUMNS[9], TRACKER_COLUMNS[18])
        self.assertEqual(TRACKER_COLUMNS[10], TRACKER_COLUMNS[19])

    def test_missing_source_columns_remain_blank_instead_of_being_invented(self):
        record = map_lccs_row({
            "XAWBID": "awb-1",
            "XAWBOOKLOC": "BLR",
            "XAWPKUPLOC": "DEL",
            "XAWSDPRTIME": "12:00:00",
            "XAWBMODE": "A",
        })

        self.assertEqual(record["values"][10], "")
        self.assertEqual(record["values"][11], "Lean 1")
        self.assertEqual(record["values"][14], "")
        self.assertEqual(record["values"][21:23], ["", ""])

        incomplete_lane = map_lccs_row({"XAWBID": "awb-2", "XAWBOOKLOC": "BLR"})
        self.assertEqual(incomplete_lane["values"][3], "")

    def test_eta_and_total_cons_come_from_lccs(self):
        values = map_lccs_row({"XAWBID": "1", "XAWSARVTIME": "12:25:00", "total_cons": 7})["values"]
        self.assertEqual((values[10], values[14]), ("12:25:00", 7))
        blank = map_lccs_row({"XAWBID": "2", "XAWSARVTIME": "00:00:00", "total_cons": 0})["values"]
        self.assertEqual((blank[10], blank[14]), ("", ""))
        self.assertIn("awb.XAWSARVTIME AS XAWSARVTIME", LCCS_TRACKER_QUERY)
        self.assertIn("COUNT(DISTINCT cons.XAWDDKTID)", LCCS_TRACKER_QUERY)

    def test_scheduled_field_overrides_survive_source_refresh_mapping(self):
        record = map_lccs_row(
            {
                "XAWBID": "awb-1",
                "XAWSDPRTIME": "08:00:00",
                "XAWFLTNO": "LH 10",
            },
            {
                "flight_number_override": "LH 11",
                "scheduled_etd_override": "09:00",
                "scheduled_eta_override": "10:30",
            },
        )

        self.assertEqual(record["values"][8:11], ["LH 11", "09:00", "10:30"])

    def test_csv_has_exact_24_column_header_order_including_duplicate_etd_eta(self):
        csv_text = tracker_csv([map_lccs_row({"XAWBID": "awb-1"})])
        rows = list(csv.reader(StringIO(csv_text)))

        self.assertEqual(rows[0], list(TRACKER_COLUMNS))
        self.assertEqual(len(rows[0]), 24)
        self.assertEqual(len(rows[1]), 24)

    def test_duplicate_identical_join_rows_are_deduplicated(self):
        source = {
            "XAWBID": "awb-1",
            "XAWBOOKLOC": "BLR",
            "XAWPKUPLOC": "DEL",
            "XAWSDPRTIME": "08:00:00",
        }

        self.assertEqual(len(map_lccs_rows([source, source])), 1)

    def test_conflicting_rows_for_one_identifier_are_rejected(self):
        common = {
            "XAWBID": "awb-1",
            "XAWBOOKLOC": "BLR",
            "XAWPKUPLOC": "DEL",
        }
        with self.assertRaisesRegex(TrackerDataError, "conflicting rows"):
            map_lccs_rows([
                {**common, "XAWFLTNO": "LH 1"},
                {**common, "XAWFLTNO": "LH 2"},
            ])

    def test_missing_stable_identifier_is_rejected(self):
        with self.assertRaisesRegex(TrackerDataError, "missing XAWBID"):
            map_lccs_row({"mawb_no": "123-456"})

    def test_ndc_dropdown_has_all_supplied_options_and_rejects_unknown_values(self):
        self.assertEqual(len(NDC_REMARK_OPTIONS), 37)
        self.assertEqual(
            NDC_REMARK_OPTIONS[0],
            "PUD Vehicle / GSA arrived in the Hub / Location after Airport Cut Off",
        )
        self.assertEqual(NDC_REMARK_OPTIONS[-1], "Vehicle caught by sales tax team")

        with self.assertRaisesRegex(TrackerDataError, "supplied dropdown options"):
            map_lccs_row(
                {"XAWBID": "awb-1"},
                {"ndc_remarks": "Not on the supplied list"},
            )

        selected = map_lccs_row(
            {"XAWBID": "awb-1"},
            {"ndc_remarks": NDC_REMARK_OPTIONS[3]},
        )
        self.assertEqual(selected["values"][20], NDC_REMARK_OPTIONS[3])

    def test_surface_mawbs_get_no_dispatch_slot_like_php_report(self):
        for mode in ("S", "", None):
            with self.subTest(mode=mode):
                record = map_lccs_row({"XAWBID": "s1", "XAWSDPRTIME": "08:00:00", "XAWBMODE": mode})
                self.assertEqual(record["values"][11], "")

    def test_query_excludes_testing_client_and_derives_nfo_from_dockets(self):
        self.assertIn("test_docket.XCLNTCD = :testing_client_code", LCCS_TRACKER_QUERY)
        self.assertIn("NOT EXISTS", LCCS_TRACKER_QUERY)
        self.assertIn("UPPER(nfo_docket.XDKTPROMODE) = 'NFO'", LCCS_TRACKER_QUERY)
        self.assertIn("nfo_docket.XLASTNEGATIVECD != 'AA'", LCCS_TRACKER_QUERY)
        self.assertIn("AS is_nfo", LCCS_TRACKER_QUERY)

    def test_query_selects_air_surface_mode(self):
        self.assertIn("awb.XAWBMODE AS XAWBMODE", LCCS_TRACKER_QUERY)

    def test_dispatch_time_boundaries_match_php_report(self):
        expected = {
            "00:00:00": "Prime",
            "09:00:00": "Prime",
            "09:00:01": "Lean 1",
            "12:00:00": "Lean 1",
            "12:00:01": "Lean 2",
            "invalid": "",
        }
        for dispatch_time, slot in expected.items():
            with self.subTest(dispatch_time=dispatch_time):
                record = map_lccs_row({
                    "XAWBID": dispatch_time,
                    "XAWSDPRTIME": dispatch_time,
                    "XAWBMODE": "A",
                })
                self.assertEqual(record["values"][11], slot)


if __name__ == "__main__":
    unittest.main()
