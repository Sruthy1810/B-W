"""Project config. Keeps the original get_input_file / get_output_file API.

Supports either:
  * Excel (.xlsx): a sheet named FILES (or the first sheet) with
        column A = key, column B = value
        e.g.  input_folder | C:/.../input
              Output_file  | C:/.../output.xlsx
  * INI text file (.ini/.cfg/.txt) with a [FILES] section.
"""
from configparser import ConfigParser
from pathlib import Path

CONFIG_FILE = Path(r"C:/Users/Admin/OneDrive - StayAhead/Documents/Projects/B&W/config/config.xlsx")

BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "logs" / "run.log"
ENABLED_PROCESSES = ["payment"]          # later: ["payment", "cost", "temp1"]


def _load_files_section() -> dict:
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(f"Config file not found: {CONFIG_FILE}")

    if CONFIG_FILE.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(CONFIG_FILE, data_only=True, read_only=True)
        ws = wb["FILES"] if "FILES" in wb.sheetnames else wb.active
        values = {}
        for row in ws.iter_rows(values_only=True):
            if row and len(row) > 1 and row[0] is not None and row[1] is not None:
                values[str(row[0]).strip().lower()] = str(row[1]).strip()
        return values

    cp = ConfigParser()
    cp.read(CONFIG_FILE, encoding="utf-8")
    if "FILES" not in cp:
        raise KeyError(f"[FILES] section missing in {CONFIG_FILE}")
    return {k.lower(): v.strip() for k, v in cp["FILES"].items()}


_FILES = _load_files_section()


def _get(key: str) -> str:
    try:
        return _FILES[key.lower()]
    except KeyError:
        raise KeyError(f"'{key}' not found in config ({CONFIG_FILE}). Found: {list(_FILES)}")


def get_input_file():
    return _get("input_folder")


def get_output_file():
    return _get("output_file")