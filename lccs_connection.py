import argparse
import getpass
import os
import sys
from pathlib import Path

from env_file import ENV_FILE, env_value, save_env


# Logs and non-secret job settings; credentials live only in the project .env file.
CONFIG_DIRECTORY = Path(
    os.environ.get("LOCALAPPDATA")
    or os.environ.get("APPDATA")
    or Path.home()
) / "OffloadDashboard"

FIELDS = (
    ("db_type", "Database type (e.g. mysql)", "mysql"),
    ("host", "Database host (name or IP)", None),
    ("port", "Database port", "3306"),
    ("database", "Database (schema) name", None),
    ("username", "Read-only username", None),
)
ENV_KEYS = {
    "db_type": "LCCS_DB_TYPE",
    "host": "LCCS_DB_HOST",
    "port": "LCCS_DB_PORT",
    "database": "LCCS_DB_NAME",
    "username": "LCCS_DB_USER",
    "password": "LCCS_DB_PASSWORD",
}


class LccsConfigError(ValueError):
    pass


def _ask(label, default):
    suffix = f" [{default}]" if default else ""
    value = input(f"{label}{suffix}: ").strip() or (default or "")
    if not value:
        raise LccsConfigError(f"{label} is required.")
    return value


def _validate(settings):
    missing = [key for key in ENV_KEYS if not str(settings.get(key, "")).strip()]
    if missing:
        raise LccsConfigError(f"Missing LCCS settings: {', '.join(missing)}")
    if not str(settings["port"]).isdigit():
        raise LccsConfigError("Port must be a number.")


def save_config(settings, path=ENV_FILE):
    _validate(settings)
    save_env({ENV_KEYS[key]: settings[key] for key in ENV_KEYS}, path)


def load_config(path=ENV_FILE):
    settings = {key: env_value(name, path) for key, name in ENV_KEYS.items()}
    if not settings["host"] or not settings["password"]:
        raise LccsConfigError(
            f"LCCS connection is not configured in {Path(path).name}. "
            "Run: python lccs_connection.py --configure"
        )
    _validate(settings)
    return {**settings, "port": int(settings["port"])}


def configure():
    print("Enter the LCCS read-only database details. Nothing is sent anywhere.")
    settings = {key: _ask(label, default) for key, label, default in FIELDS}
    password = getpass.getpass("Password (input is hidden): ")
    if password != getpass.getpass("Password again: "):
        raise LccsConfigError("Passwords did not match.")
    settings["password"] = password
    save_config(settings)
    print(f"Saved to {ENV_FILE} (ignored by Git).")


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
