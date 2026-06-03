"""
============================================================
Test SQL Connection - Daily E-Sales POS Automation
============================================================

Purpose:
    Test if Python can read .env and connect to POS SQL Server.

This script does NOT generate Excel.
This script only checks:
    1. If .env exists
    2. If required settings are readable
    3. If SQL Server connection works
    4. If SELECT 1 works

Run:
    py test_connection.py

or:
    TEST_CONNECTION.bat

============================================================
"""

import os
from pathlib import Path

import pyodbc
from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"


def load_settings() -> dict:
    """
    Load .env settings.
    """
    load_dotenv(ENV_PATH)

    return {
        "SQL_SERVER": os.getenv("SQL_SERVER", "").strip(),
        "SQL_DATABASE": os.getenv("SQL_DATABASE", "").strip(),
        "SQL_USERNAME": os.getenv("SQL_USERNAME", "").strip(),
        "SQL_PASSWORD": os.getenv("SQL_PASSWORD", ""),
        "SQL_DRIVER": os.getenv("SQL_DRIVER", "ODBC Driver 17 for SQL Server").strip(),
        "SQL_ENCRYPT": os.getenv("SQL_ENCRYPT", "no").strip(),
        "SQL_TRUST_SERVER_CERTIFICATE": os.getenv("SQL_TRUST_SERVER_CERTIFICATE", "yes").strip(),

        "POS_BRANCH": os.getenv("POS_BRANCH", "").strip(),
        "POS_COMPANY_NAME": os.getenv("POS_COMPANY_NAME", "").strip(),
        "POS_OPERATOR_NAME": os.getenv("POS_OPERATOR_NAME", "").strip(),

        "DATE_MODE": os.getenv("DATE_MODE", "PREVIOUS_MONTH").strip(),
        "DATE_FROM": os.getenv("DATE_FROM", "").strip(),
        "DATE_TO": os.getenv("DATE_TO", "").strip(),

        "VAT_ROUNDING_MODE": os.getenv("VAT_ROUNDING_MODE", "PER_TRANSACTION").strip(),

        "OUTPUT_DIR": os.getenv("OUTPUT_DIR", "output").strip(),
        "LOG_DIR": os.getenv("LOG_DIR", "logs").strip(),
    }


def print_env_debug(settings: dict) -> None:
    """
    Print important .env values for checking.

    Important:
        Do NOT print actual SQL_PASSWORD.
        Password length lang ang pinapakita for privacy.
    """
    print("========== ENV DEBUG ==========")
    print("ENV_PATH:", ENV_PATH)
    print("ENV_EXISTS:", ENV_PATH.exists())

    print("\n--- SQL SERVER CONNECTION ---")
    print("SQL_SERVER:", settings["SQL_SERVER"])
    print("SQL_DATABASE:", settings["SQL_DATABASE"])
    print("SQL_USERNAME:", settings["SQL_USERNAME"])
    print("SQL_PASSWORD length:", len(settings["SQL_PASSWORD"]) if settings["SQL_PASSWORD"] else 0)
    print("SQL_DRIVER:", settings["SQL_DRIVER"])
    print("SQL_ENCRYPT:", settings["SQL_ENCRYPT"])
    print("SQL_TRUST_SERVER_CERTIFICATE:", settings["SQL_TRUST_SERVER_CERTIFICATE"])

    print("\n--- POS / REPORT HEADER INFO ---")
    print("POS_BRANCH:", settings["POS_BRANCH"])
    print("POS_COMPANY_NAME:", settings["POS_COMPANY_NAME"])
    print("POS_OPERATOR_NAME:", settings["POS_OPERATOR_NAME"])

    print("\n--- DATE SETTINGS ---")
    print("DATE_MODE:", settings["DATE_MODE"])
    print("DATE_FROM:", settings["DATE_FROM"])
    print("DATE_TO:", settings["DATE_TO"])

    print("\n--- VAT / OUTPUT SETTINGS ---")
    print("VAT_ROUNDING_MODE:", settings["VAT_ROUNDING_MODE"])
    print("OUTPUT_DIR:", settings["OUTPUT_DIR"])
    print("LOG_DIR:", settings["LOG_DIR"])
    print("================================")


def build_connection_string(settings: dict) -> str:
    """
    Build SQL Server connection string.
    """
    return (
        f"DRIVER={{{settings['SQL_DRIVER']}}};"
        f"SERVER={settings['SQL_SERVER']};"
        f"DATABASE={settings['SQL_DATABASE']};"
        f"UID={settings['SQL_USERNAME']};"
        f"PWD={settings['SQL_PASSWORD']};"
        f"Encrypt={settings['SQL_ENCRYPT']};"
        f"TrustServerCertificate={settings['SQL_TRUST_SERVER_CERTIFICATE']};"
    )


def validate_required_settings(settings: dict) -> None:
    """
    Check required .env values before connecting.
    """
    required = [
        "SQL_SERVER",
        "SQL_DATABASE",
        "SQL_USERNAME",
        "SQL_PASSWORD",
        "SQL_DRIVER",
    ]

    missing = [key for key in required if not settings[key]]

    if missing:
        raise ValueError("Missing required .env value(s): " + ", ".join(missing))


def main():
    print("Starting SQL connection test...")

    settings = load_settings()
    print_env_debug(settings)
    validate_required_settings(settings)

    conn_str = build_connection_string(settings)

    try:
        print("\nConnecting to SQL Server...")
        conn = pyodbc.connect(conn_str, timeout=15)

        print("Connected successfully.")

        cursor = conn.cursor()
        cursor.execute("SELECT 1 AS TestValue")
        row = cursor.fetchone()

        print("SELECT 1 result:", row.TestValue)

        conn.close()
        print("\nSUCCESS: SQL connection test passed.")

    except Exception as exc:
        print("\nERROR: SQL connection test failed.")
        print(exc)
        raise


if __name__ == "__main__":
    main()