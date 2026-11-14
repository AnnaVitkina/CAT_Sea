"""Defaults for the logical RA compare pipeline (separate from ra_update)."""

from pathlib import Path

from ra_update.config import OUTPUT_DIR, RA_FILENAME, RA_INPUT_DIR

ROOT = Path(__file__).resolve().parent.parent

# Same folders as ra_update.config (Colab Drive if present, else local input/output)
# Baseline: input/rate agreement/
# Candidate: output/ (*_updated_*.xlsx)
BASELINE_DIR = RA_INPUT_DIR
CANDIDATE_DIR = OUTPUT_DIR
DEFAULT_BASELINE_NAME = RA_FILENAME

# Cap large one-side cost lists in the raw JSON
COST_PRESENCE_CAP = 500

# Output names are built as '{RA short}_ {update stem}_logical_diff_*.json/.xlsx'
# e.g. 'RA20260513093 v.15_ Additional lanes UK_logical_diff_refined.json'
