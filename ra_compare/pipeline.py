"""Logical RA compare pipeline — separate from ra_update.run_pipeline."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from ra_update.output_names import stem_from_baseline_and_candidate

from .compare import compare_workbooks
from .config import OUTPUT_DIR
from .refine import refine_diff
from .report import print_summary, write_json, write_xlsx


@dataclass
class ComparePipelineResult:
    baseline_path: Path
    candidate_path: Path
    raw_json_path: Path
    refined_json_path: Path
    xlsx_path: Path
    paired_lanes: int
    matched_identical: int
    real_rate_mismatches: int
    currency_only_cost_diffs: int
    identity_logical: int
    lanes_only_baseline: int
    lanes_only_candidate: int
    name_prefix: str


def run_compare_pipeline(
    baseline_path: Path | None = None,
    candidate_path: Path | None = None,
    output_dir: Path | None = None,
    *,
    baseline_label: str = "baseline",
    candidate_label: str = "candidate",
    update_path: Path | None = None,
    name_prefix: str | None = None,
) -> ComparePipelineResult:
    """
    Compare two Rate Card Advanced Export workbooks logically.

    Writes (prefix e.g. 'RA20260513093 v.15_ Additional lanes UK'):
      - {prefix}_logical_diff_raw.json
      - {prefix}_logical_diff_refined.json
      - {prefix}_logical_diff.xlsx
    """
    if baseline_path is None or candidate_path is None:
        raise ValueError(
            "baseline_path and candidate_path are required "
            "(choose via run_compare_pipeline.py interactive menu)"
        )
    baseline_path = Path(baseline_path)
    candidate_path = Path(candidate_path)
    output_dir = Path(output_dir or OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not baseline_path.exists():
        raise FileNotFoundError(f"Baseline RA not found: {baseline_path}")
    if not candidate_path.exists():
        raise FileNotFoundError(f"Candidate RA not found: {candidate_path}")

    prefix = name_prefix or stem_from_baseline_and_candidate(
        baseline_path, candidate_path, update_path=update_path
    )

    raw = compare_workbooks(
        baseline_path,
        candidate_path,
        baseline_label=baseline_label,
        candidate_label=candidate_label,
    )
    refined = refine_diff(raw)

    raw_path = write_json(output_dir / f"{prefix}_logical_diff_raw.json", raw)
    refined_path = write_json(
        output_dir / f"{prefix}_logical_diff_refined.json", refined
    )
    xlsx_path = write_xlsx(
        output_dir / f"{prefix}_logical_diff.xlsx", refined, raw=raw
    )
    print_summary(raw, refined)
    print(f"\nWrote {raw_path}")
    print(f"Wrote {refined_path}")
    print(f"Wrote {xlsx_path}")

    s = refined["summary"]
    return ComparePipelineResult(
        baseline_path=baseline_path,
        candidate_path=candidate_path,
        raw_json_path=raw_path,
        refined_json_path=refined_path,
        xlsx_path=xlsx_path,
        paired_lanes=s["paired_lanes"],
        matched_identical=s["matched_identical"],
        real_rate_mismatches=s["real_rate_mismatches"],
        currency_only_cost_diffs=s["currency_only_cost_diffs"],
        identity_logical=s["identity_logical"],
        lanes_only_baseline=s["lanes_only_baseline"],
        lanes_only_candidate=s["lanes_only_candidate"],
        name_prefix=prefix,
    )
