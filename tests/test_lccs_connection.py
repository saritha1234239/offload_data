import json
import tempfile
import unittest
from pathlib import Path

from lccs_connection import LccsConfigError, load_config, save_config


VALID = {
    "db_type": "mysql",
    "host": "lccs-db.internal",
    "port": "3306",
    "database": "lccs",
    "username": "reader",
    "password": "secret",
}


class LccsConnectionTests(unittest.TestCase):
    def test_round_trip_saves_port_as_number(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sub" / "lccs-config.json"
            save_config(VALID, path)
            self.assertEqual(load_config(path), {**VALID, "port": 3306})
            self.assertEqual(json.loads(path.read_text())["host"], "lccs-db.internal")

    def test_missing_fields_and_bad_port_are_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "c.json"
            with self.assertRaisesRegex(LccsConfigError, "host"):
                save_config({**VALID, "host": " "}, path)
            with self.assertRaisesRegex(LccsConfigError, "password"):
                save_config({**VALID, "password": ""}, path)
            with self.assertRaisesRegex(LccsConfigError, "Port"):
                save_config({**VALID, "port": "abc"}, path)
            self.assertFalse(path.exists())

    def test_unconfigured_load_explains_next_step(self):
        with self.assertRaisesRegex(LccsConfigError, "--configure"):
            load_config(Path(tempfile.gettempdir()) / "does-not-exist-lccs.json")


if __name__ == "__main__":
    unittest.main()
