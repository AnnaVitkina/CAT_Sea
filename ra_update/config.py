import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Colab Google Drive CAT_Sea (only used when that path exists — not on Windows)
_COLAB_CAT_SEA = (
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_CAT/CAT_Sea"
)


def _resolve_io_dirs() -> tuple[Path, Path, Path | None]:
    """
    Prefer Colab Drive folders when available; otherwise local repo input/output.
    Never touch the Drive path on Windows (colon in share name → WinError 123).
    """
    if os.name != "nt":
        colab_root = Path(_COLAB_CAT_SEA)
        if colab_root.is_dir():
            return colab_root / "input", colab_root / "output", colab_root
    return ROOT / "input", ROOT / "output", None


INPUT_DIR, OUTPUT_DIR, CAT_SEA_ROOT = _resolve_io_dirs()

# Separate input folders under input/
RA_INPUT_DIR = INPUT_DIR / "rate agreement"
UPDATE_INPUT_DIR = INPUT_DIR / "update"

# Preferred defaults when present (non-interactive / Enter in menu)
RA_FILENAME = "Advanced Export - RA20260513093 v.13 - Sea.xlsx"
UPDATE_FILENAME = "OC_RATES_PIVOT_OB for perkins.xlsx"

RA_SHEET = "Rate card"
GENERAL_INFO_SHEET = "General info"
# BASE first: lane cuts/duplicates from BASE must exist before ETSBAF runs
UPDATE_SHEETS = ("BASE", "ETSBAF")

# Rate Card layout
HEADER_COST_ROW = 3  # cost names (e.g. BAF Fee (22G0 15.09.2026-14.10.2026))
HEADER_FIELD_ROW = 7  # Lane #, KEY, Carrier, p/unit, ...
DATA_START_ROW = 8

# Classic Sea layout defaults (Sea FR omits postal cols — resolve via lane_columns)
COL_LANE = 1
COL_TRANSPOREON = 2
COL_KEY = 3
COL_CARRIER = 4
COL_SERVICE = 5
COL_SERVICE_C = 6
COL_SERVICE_SHORT = 7
COL_VALID_FROM = 8
COL_VALID_TO = 9
COL_ORIGIN_PORT = 10
COL_ORIGIN_POSTAL = 11
COL_ORIGIN_COUNTRY = 12
COL_DEST_PORT = 13
COL_DEST_POSTAL = 14
COL_DEST_COUNTRY = 15

# Countries that trigger auto-add of missing BASE lanes on Sea RAs
SEA_LANE_COUNTRIES = frozenset({"GB", "ZA"})
SEA_FR_DEST_COUNTRIES = frozenset({"FR"})

# Sentinel mins in update file are ignored
IGNORE_MIN_ABS_THRESHOLD = 1_000_000

# Green highlight for changed p/unit cells
CHANGED_FILL_HEX = "92D050"
# Red highlight when multiple container codes share one cost cell
DUAL_CONTAINER_FILL_HEX = "FF0000"
# Orange: identity field mismatch — RA value kept and update appended with /
MISMATCH_FILL_HEX = "F4B183"
