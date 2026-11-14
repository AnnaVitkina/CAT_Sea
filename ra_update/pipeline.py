from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from .apply_updates import IssueRow, UpdateApplier
from .baf_create import BafCreateStore, resolve_baf_create_cutoffs
from .config import (
    INPUT_DIR,
    OUTPUT_DIR,
    RA_FILENAME,
    RA_INPUT_DIR,
    UPDATE_FILENAME,
    UPDATE_INPUT_DIR,
    UPDATE_SHEETS,
)
from .equipment_decisions import (
    EquipmentDecisionStore,
    resolve_missing_equipment_interactive,
)
from .ra_model import RateCardWorkbook
from .output_names import pair_stem
from .reporting import count_real_issues_from_stats, write_reports


@dataclass
class PipelineResult:
    updated_ra_path: Path
    issues_path: Path
    changes_path: Path
    processed_update_rows: int
    changed_cells: int
    already_equal: int
    issue_count: int
    note_count: int = 0


def run_pipeline(
    ra_path: Path | None = None,
    update_path: Path | None = None,
    output_dir: Path | None = None,
    *,
    interactive_equipment: bool = True,
    equipment_decisions_path: Path | None = None,
) -> PipelineResult:
    # Prefer new folders; fall back to legacy flat input/ paths if needed
    if ra_path is None:
        preferred_ra = RA_INPUT_DIR / RA_FILENAME
        legacy_ra = INPUT_DIR / RA_FILENAME
        ra_path = preferred_ra if preferred_ra.exists() else legacy_ra
    if update_path is None:
        preferred_up = UPDATE_INPUT_DIR / UPDATE_FILENAME
        legacy_up = INPUT_DIR / UPDATE_FILENAME
        update_path = preferred_up if preferred_up.exists() else legacy_up
    output_dir = output_dir or OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)
    RA_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    UPDATE_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    decisions_path = equipment_decisions_path or (INPUT_DIR / "equipment_decisions.json")

    if not ra_path.exists():
        raise FileNotFoundError(ra_path)
    if not update_path.exists():
        raise FileNotFoundError(update_path)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    print(f"Loading RA: {ra_path.name}")
    ra = RateCardWorkbook(ra_path)
    print(
        f"  title={ra.agreement_title!r} lanes={len(ra.lanes)} "
        f"cost_columns={len(ra.cost_columns)} "
        f"agreement_validity={ra.agreement_valid_from}..{ra.agreement_valid_to}"
    )

    print("Resolving missing equipment cost columns...")
    equipment_store = EquipmentDecisionStore()
    resolve_missing_equipment_interactive(
        ra,
        update_path,
        equipment_store,
        decisions_path=decisions_path,
        interactive=interactive_equipment,
    )

    print("Resolving BAF Fee monthly create cutoffs...")
    baf_create_store = BafCreateStore()
    resolve_baf_create_cutoffs(
        ra,
        update_path,
        baf_create_store,
        decisions_path=decisions_path,
        interactive=interactive_equipment,
        aliases=equipment_store.aliases or None,
    )

    applier = UpdateApplier(
        ra,
        equipment_store=equipment_store,
        baf_create_store=baf_create_store,
    )

    # Record BAF create-until decisions in issues/notes
    for eq, note in sorted(baf_create_store.decision_notes.items()):
        applier.stats.issues.append(
            IssueRow(
                sheet="ETSBAF",
                update_row_index=-1,
                key="",
                rate_col=None,
                equipment=eq,
                charge_type="BAF",
                new_value=None,
                lane_no=None,
                ra_row=None,
                reason=note,
            )
        )

    # Intentionally BASE then ETSBAF so BASE lane splits are visible to ETSBAF.
    for sheet in UPDATE_SHEETS:
        print(f"Processing update sheet: {sheet}")
        df = pd.read_excel(update_path, sheet_name=sheet)
        print(f"  rows={len(df)}")
        applier.process_sheet(sheet, df)

    stats = applier.stats
    changed = sum(1 for c in stats.changes if c.status == "changed")
    equal = sum(1 for c in stats.changes if c.status == "already_equal")

    prefix = pair_stem(ra_path, update_path)
    out_ra = output_dir / f"{prefix}_updated_{stamp}.xlsx"
    print(f"Saving updated RA -> {out_ra.name}")
    ra.save(out_ra)

    report_paths = write_reports(stats, output_dir, stamp, name_prefix=prefix)
    real_issues, notes = count_real_issues_from_stats(stats)
    print(
        f"Done. changed={changed} already_equal={equal} "
        f"issues={real_issues} notes={notes} "
        f"earlier_validity={len(stats.earlier_validity)}"
    )
    print(f"Issues report: {report_paths['issues'].name}")
    print(f"Changes report: {report_paths['changes'].name}")

    return PipelineResult(
        updated_ra_path=out_ra,
        issues_path=report_paths["issues"],
        changes_path=report_paths["changes"],
        processed_update_rows=stats.processed_update_rows,
        changed_cells=changed,
        already_equal=equal,
        issue_count=real_issues,
        note_count=notes,
    )
