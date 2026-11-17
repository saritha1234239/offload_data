"""The project's .env file: the single place this computer keeps credentials.

Values already set in the process environment (for example GitHub Actions
secrets or Render settings) take precedence over the file.
"""
import os
from pathlib import Path

from dotenv import dotenv_values, set_key


ENV_FILE = Path(__file__).with_name(".env")


def env_value(name, path=ENV_FILE):
    value = os.environ.get(name)
    if value is None or not value.strip():
        value = dotenv_values(path).get(name) if Path(path).is_file() else None
    return (value or "").strip()


def save_env(values, path=ENV_FILE):
    """Add or replace keys in the .env file, keeping every other line."""
    path = Path(path)
    path.touch(exist_ok=True)
    for name, value in values.items():
        set_key(str(path), name, str(value), quote_mode="always")
