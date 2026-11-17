import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from lccs_connection import ENV_KEYS, LccsConfigError, load_config, save_config


VALID = {
    "db_type": "mysql",
    "host": "lccs-db.internal",
    "port": "3306",
    "database": "lccs",
    "username": "reader",
    "password": "secret",
}
NO_LCCS_ENV = {name: value for name, value in os.environ.items() if name not in ENV_KEYS.values()}


class LccsConnectionTests(unittest.TestCase):
    def setUp(self):
        patcher = patch.dict(os.environ, NO_LCCS_ENV, clear=True)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_round_trip_saves_to_env_file_and_keeps_other_lines(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / ".env"
            path.write_text("SMTP_HOST=smtp.example.com\n", encoding="utf-8")
            save_config(VALID, path)
            self.assertEqual(load_config(path), {**VALID, "port": 3306})
            text = path.read_text(encoding="utf-8")
            self.assertIn("SMTP_HOST=smtp.example.com", text)
            self.assertIn("LCCS_DB_HOST='lccs-db.internal'", text)

    def test_process_environment_overrides_env_file(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / ".env"
            save_config(VALID, path)
            with patch.dict(os.environ, {"LCCS_DB_HOST": "other-host"}):
                self.assertEqual(load_config(path)["host"], "other-host")

    def test_missing_fields_and_bad_port_are_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / ".env"
            with self.assertRaisesRegex(LccsConfigError, "host"):
                save_config({**VALID, "host": " "}, path)
            with self.assertRaisesRegex(LccsConfigError, "password"):
                save_config({**VALID, "password": ""}, path)
            with self.assertRaisesRegex(LccsConfigError, "Port"):
                save_config({**VALID, "port": "abc"}, path)
            self.assertFalse(path.exists())

    def test_unconfigured_load_explains_next_step(self):
        with self.assertRaisesRegex(LccsConfigError, "--configure"):
            load_config(Path(tempfile.gettempdir()) / "does-not-exist.env")


if __name__ == "__main__":
    unittest.main()
