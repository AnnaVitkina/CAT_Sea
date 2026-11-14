"""Entry point: logical compare of two Rate Card Advanced Export workbooks."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    _REPO_ROOT = Path(__file__).resolve().parent
except NameError:
    _cand = Path("/content/CAT_Sea")
    _REPO_ROOT = _cand if _cand.is_dir() else Path.cwd()
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from ra_compare.config import (
    BASELINE_DIR,
    CANDIDATE_DIR,
    DEFAULT_BASELINE_NAME,
    OUTPUT_DIR,
)
from ra_compare.pipeline import run_compare_pipeline
from ra_update.file_select import ensure_input_folders, list_xlsx, resolve_xlsx
from ra_update.output_names import stem_from_baseline_and_candidate


def _default_interactive() -> bool:
    if Path("/content").exists():
        return False
    return sys.stdin.isatty()


def _newest_updated_ra(folder: Path) -> str | None:
    """Prefer '{pair}_updated_*.xlsx', fall back to legacy RA_UPDATED_*."""
    files = list_xlsx(folder)

    def _is_updated(p: Path) -> bool:
        n = p.name
        return "_updated_" in n.lower() or n.startswith("RA_UPDATED_")

    updated = sorted(
        (p for p in files if _is_updated(p)),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return updated[0].name if updated else None


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Logically compare two Rate Card workbooks "
            "(lanes by KEY/validity/OD; costs by fee/equipment/period — not row/column #). "
            f"Baseline: {BASELINE_DIR} · candidate/output: {CANDIDATE_DIR}"
        )
    )
    parser.add_argument(
        "--baseline",
        type=Path,
        default=None,
        help=f"Reference RA xlsx (default: choose from {BASELINE_DIR})",
    )
    parser.add_argument(
        "--candidate",
        type=Path,
        default=None,
        help=f"Candidate / updated RA xlsx (default: choose from {CANDIDATE_DIR})",
    )
    parser.add_argument(
        "--update",
        type=Path,
        default=None,
        help="Optional update pivot used for output naming (else derived from candidate name)",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Directory for logical_diff JSON + XLSX",
    )
    parser.add_argument(
        "--baseline-label",
        default="baseline",
        help="Label used in notes/meta (default: baseline)",
    )
    parser.add_argument(
        "--candidate-label",
        default="candidate",
        help="Label used in notes/meta (default: candidate)",
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="Use preferred baseline name and newest *_updated_* without prompting (default on Colab)",
    )
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="Force interactive file prompts even on Colab",
    )
    args = parser.parse_args()

    out_dir = args.output_dir or OUTPUT_DIR
    ensure_input_folders(BASELINE_DIR, CANDIDATE_DIR, out_dir)
    if args.interactive:
        interactive = True
    elif args.non_interactive:
        interactive = False
    else:
        interactive = _default_interactive()
    print(f"Mode: {'interactive' if interactive else 'non-interactive'}")
    print(f"Baseline dir: {BASELINE_DIR}")
    print(f"Output dir  : {out_dir}")

    baseline = resolve_xlsx(
        "baseline rate agreement",
        BASELINE_DIR,
        args.baseline,
        interactive=interactive,
        preferred_name=DEFAULT_BASELINE_NAME,
    )

    preferred_candidate = _newest_updated_ra(CANDIDATE_DIR)
    candidate = resolve_xlsx(
        "candidate (updated RA)",
        CANDIDATE_DIR,
        args.candidate,
        interactive=interactive,
        preferred_name=preferred_candidate,
    )

    update_path = Path(args.update) if args.update is not None else None
    prefix = stem_from_baseline_and_candidate(
        baseline, candidate, update_path=update_path
    )
    print(f"Baseline : {baseline}")
    print(f"Candidate: {candidate}")
    if update_path:
        print(f"Update   : {update_path}")
    print(f"Prefix   : {prefix}")

    result = run_compare_pipeline(
        baseline_path=baseline,
        candidate_path=candidate,
        output_dir=out_dir,
        baseline_label=args.baseline_label,
        candidate_label=args.candidate_label,
        update_path=update_path,
        name_prefix=prefix,
    )
    print()
    print("=== COMPARE SUMMARY ===")
    print(f"Baseline   : {result.baseline_path.name}")
    print(f"Candidate  : {result.candidate_path.name}")
    print(f"Raw JSON   : {result.raw_json_path}")
    print(f"Refined    : {result.refined_json_path}")
    print(f"XLSX       : {result.xlsx_path}")
    print(
        f"Paired {result.paired_lanes} · identical {result.matched_identical} · "
        f"rate diffs {result.real_rate_mismatches} · "
        f"currency-only {result.currency_only_cost_diffs} · "
        f"identity logical {result.identity_logical}"
    )


if __name__ == "__main__":
    main()
