import argparse
import getpass
import json
import os
import sys
import tempfile
from pathlib import Path


CONFIG_DIRECTORY = Path(
    os.environ.get("LOCALAPPDATA")
    or os.environ.get("APPDATA")
    or Path.home()
) / "OffloadDashboard"
CONFIG_PATH = CONFIG_DIRECTORY / "lccs-config.json"

FIELDS = (
    ("db_type", "Database type (e.g. mysql)", "mysql"),
    ("host", "Database host (name or IP)", None),
    ("port", "Database port", "3306"),
    ("database", "Database (schema) name", None),
    ("username", "Read-only username", None),
)


class LccsConfigError(ValueError):
    pass


def _ask(label, default):
    suffix = f" [{default}]" if default else ""
    value = input(f"{label}{suffix}: ").strip() or (default or "")
    if not value:
        raise LccsConfigError(f"{label} is required.")
    return value


def save_config(settings, path=CONFIG_PATH):
    missing = [key for key, _, _ in FIELDS if not str(settings.get(key, "")).strip()]
    if missing or not settings.get("password"):
        raise LccsConfigError(f"Missing LCCS settings: {', '.join(missing + (['password'] if not settings.get('password') else []))}")
    if not str(settings["port"]).isdigit():
        raise LccsConfigError("Port must be a number.")
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(prefix=".lccs-config-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as output:
            json.dump({**settings, "port": int(settings["port"])}, output, indent=2)
        os.replace(temporary, path)
    except Exception:
        try:
            os.unlink(temporary)
        except OSError:
            pass
        raise


def load_config(path=CONFIG_PATH):
    path = Path(path)
    if not path.is_file():
        raise LccsConfigError("LCCS connection is not configured. Run: python lccs_connection.py --configure")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise LccsConfigError(f"Could not read {path}.") from exc


def configure():
    print("Enter the LCCS read-only database details. Nothing is sent anywhere.")
    settings = {key: _ask(label, default) for key, label, default in FIELDS}
    password = getpass.getpass("Password (input is hidden): ")
    if password != getpass.getpass("Password again: "):
        raise LccsConfigError("Passwords did not match.")
    settings["password"] = password
    save_config(settings)
    print(f"Saved to {CONFIG_PATH} (this Windows user only; not in the repository).")


def show():
    settings = load_config()
    for key, label, _ in FIELDS:
        print(f"{label}: {settings.get(key, '')}")
    print("Password: " + ("(saved)" if settings.get("password") else "(missing)"))


def connect(settings=None):
    settings = settings or load_config()
    if str(settings.get("db_type", "")).lower() not in ("mysql", "mariadb"):
        raise LccsConfigError("Only MySQL/MariaDB LCCS databases are supported.")
    import pymysql

    return pymysql.connect(
        host=settings["host"],
        port=int(settings["port"]),
        user=settings["username"],
        password=settings["password"],
        database=settings["database"],
        connect_timeout=15,
        read_timeout=300,
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor,
    )


REQUIRED_TABLES = ("xawbhdr", "xawbdtl", "xdockethdr", "mhubmst", "xusrhublinkmst")


def test_connection():
    import pymysql

    try:
        connection = connect()
    except pymysql.MySQLError as exc:
        raise LccsConfigError(f"Could not connect to LCCS: {exc.args[-1] if exc.args else exc}") from exc
    try:
        with connection.cursor() as cursor:
            missing = []
            for table in REQUIRED_TABLES:
                try:
                    cursor.execute(f"SELECT 1 FROM `{table}` LIMIT 1")
                    cursor.fetchall()
                except pymysql.MySQLError:
                    missing.append(table)
    finally:
        connection.close()
    if missing:
        raise LccsConfigError("Connected, but cannot read: " + ", ".join(missing))
    print("Connected. All required LCCS tables are readable.")


def main():
    parser = argparse.ArgumentParser(description="Store the LCCS database connection settings on this computer.")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--configure", action="store_true", help="Enter or replace the connection details.")
    mode.add_argument("--show", action="store_true", help="Show saved details (password hidden).")
    mode.add_argument("--test", action="store_true", help="Connect and check the required tables are readable.")
    args = parser.parse_args()
    try:
        if args.configure:
            configure()
        elif args.test:
            test_connection()
        else:
            show()
    except (OSError, LccsConfigError) as exc:
        print(f"LCCS configuration failed: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
