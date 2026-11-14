"""Entry point: apply BASE + ETSBAF updates onto the Rate Agreement Rate Card."""

from __future__ import annotations

import argparse
from pathlib import Path

from ra_update.config import (
    INPUT_DIR,
    OUTPUT_DIR,
    RA_FILENAME,
    RA_INPUT_DIR,
    UPDATE_FILENAME,
    UPDATE_INPUT_DIR,
)
from ra_update.file_select import ensure_input_folders, resolve_xlsx
from ra_update.pipeline import run_pipeline


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Update Rate Agreement Rate Card from OC_RATES_PIVOT BASE/ETSBAF sheets. "
            f"RA files: {RA_INPUT_DIR} · updates: {UPDATE_INPUT_DIR} · "
            f"output: {OUTPUT_DIR}"
        )
    )
    parser.add_argument(
        "--ra",
        type=Path,
        default=None,
        help=f"Path to RA Advanced Export xlsx (default: choose from {RA_INPUT_DIR})",
    )
    parser.add_argument(
        "--update",
        type=Path,
        default=None,
        help=f"Path to update pivot xlsx (default: choose from {UPDATE_INPUT_DIR})",
    )
    parser.add_argument("--output-dir", type=Path, default=None, help="Output directory")
    parser.add_argument(
        "--equipment-decisions",
        type=Path,
        default=None,
        help=(
            "JSON file with create/alias decisions for missing equipment "
            f"(default: {INPUT_DIR / 'equipment_decisions.json'})"
        ),
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help=(
            "Do not prompt for files, missing equipment, or BAF create-until; "
            "uses preferred filenames or the sole file in each folder"
        ),
    )
    args = parser.parse_args()

    ensure_input_folders(RA_INPUT_DIR, UPDATE_INPUT_DIR, OUTPUT_DIR)
    interactive = not args.non_interactive

    ra_path = resolve_xlsx(
        "rate agreement (Advanced Export)",
        RA_INPUT_DIR,
        args.ra,
        interactive=interactive,
        preferred_name=RA_FILENAME,
    )
    update_path = resolve_xlsx(
        "update pivot (OC_RATES_PIVOT)",
        UPDATE_INPUT_DIR,
        args.update,
        interactive=interactive,
        preferred_name=UPDATE_FILENAME,
    )

    print(f"Rate agreement: {ra_path}")
    print(f"Update file   : {update_path}")

    result = run_pipeline(
        ra_path=ra_path,
        update_path=update_path,
        output_dir=args.output_dir,
        interactive_equipment=interactive,
        equipment_decisions_path=args.equipment_decisions,
    )
    print()
    print("=== SUMMARY ===")
    print(f"Updated RA : {result.updated_ra_path}")
    print(
        f"Issues     : {result.issues_path} "
        f"({result.issue_count} issues, {result.note_count} notes)"
    )
    print(f"Changes    : {result.changes_path}")
    print(f"Update rows: {result.processed_update_rows}")
    print(f"Cells set  : {result.changed_cells} changed, {result.already_equal} already equal")


if __name__ == "__main__":
    main()
