from __future__ import annotations

from dataclasses import asdict
from pathlib import Path

import pandas as pd

from .apply_updates import UpdateStats
from .config import UPDATE_SHEETS

ISSUE_COLUMNS = [
    "sheet",
    "update_row_index",
    "key",
    "rate_col",
    "equipment",
    "charge_type",
    "new_value",
    "lane_no",
    "ra_row",
    "reason",
    "reason_category",
]

CHANGE_COLUMNS = [
    "sheet",
    "update_row_index",
    "key",
    "lane_no",
    "ra_row",
    "rate_col",
    "cost_header",
    "punit_col",
    "old_value",
    "new_value",
    "status",
]

EARLIER_VALIDITY_COLUMNS = [
    "sheet",
    "update_row_index",
    "key",
    "lane_no",
    "ra_row",
    "original_valid_from",
    "original_valid_to",
    "effective_valid_from",
    "effective_valid_to",
    "agreement_valid_from",
    "agreement_valid_to",
    "note",
]


# Informational tabs / decision notes — kept in the workbook but NOT counted as issues.
NOTE_CATEGORIES = frozenset(
    {
        "Out of scope",
        "BAF period note",
        "BAF create cutoff",
        "ETS wider than lane",
        "KEY slash/BU resolve",
        "KEY not found",
        "Identity mismatch",
        "SPOT dual validity",
    }
)


def categorize_reason(reason: object) -> str:
    """Map full reason text to a short, stable tab name."""
    text = "" if reason is None or (isinstance(reason, float) and pd.isna(reason)) else str(reason)
    lower = text.lower()

    if "resolved via slash-combined" in lower or "resolved via non-bu sibling" in lower:
        return "KEY slash/BU resolve"
    if lower.strip() == "key not found" or lower.startswith("key not found"):
        return "KEY not found"
    if "identity mismatch" in lower:
        return "Identity mismatch"
    if "no such lane" in lower:
        return "No such lane"
    if "details/validity" in lower or "validity must be identical" in lower:
        return "Details validity"
    if "ambiguous" in lower:
        return "Ambiguous lane"
    if "out of scope" in lower or "not in scope" in lower:
        return "Out of scope"
    if "lane split failed" in lower:
        return "Lane split failed"
    if "lane add failed" in lower or "not addable" in lower:
        return "Lane add"
    if "not covered" in lower or "charge column not covered" in lower or "charge type" in lower:
        return "Not covered"
    if "not applicable on sheet" in lower:
        return "Wrong sheet"
    if "without gaps" in lower or "cost column not found" in lower or "no cost column" in lower:
        return "Cost not found"
    if "dual container" in lower:
        return "Dual container"
    if "baf create until" in lower or "baf create skipped" in lower:
        return "BAF create cutoff"
    if "15th" in lower or "expanded to" in lower or "not aligned to" in lower:
        return "BAF period note"
    if "etsbaf ok" in lower or "wider than lane validity" in lower:
        return "ETS wider than lane"
    if "spot price preferred" in lower:
        return "SPOT dual validity"
    if "potential cost column" in lower and "not creating new columns" in lower:
        return "DFT validity"
    if "dft " in lower and "validity does not match" in lower:
        return "DFT validity"
    if "already written" in lower or "conflicting" in lower:
        return "Write conflict"
    if "empty key" in lower:
        return "Empty KEY"
    if "no non-empty" in lower:
        return "No rates"
    if text.strip():
        return "Other"
    return "Unknown"


def is_note_category(category: object) -> bool:
    return str(category) in NOTE_CATEGORIES


def count_real_issues(issues_df: pd.DataFrame) -> tuple[int, int]:
    """Return (real_issue_count, note_count)."""
    if issues_df is None or issues_df.empty:
        return 0, 0
    if "reason_category" not in issues_df.columns:
        cats = issues_df["reason"].map(categorize_reason) if "reason" in issues_df.columns else []
        series = pd.Series(list(cats))
    else:
        series = issues_df["reason_category"]
    note_mask = series.map(is_note_category)
    return int((~note_mask).sum()), int(note_mask.sum())


def count_real_issues_from_stats(stats: UpdateStats) -> tuple[int, int]:
    """Return (real_issue_count, note_count) from in-memory issue rows."""
    if not stats.issues:
        return 0, 0
    note_n = sum(1 for i in stats.issues if is_note_category(categorize_reason(i.reason)))
    return len(stats.issues) - note_n, note_n


def _safe_sheet_name(name: str, used: set[str]) -> str:
    """Excel sheet names: max 31 chars, unique, no \\ / * ? : [ ]."""
    cleaned = "".join("_" if ch in r"\/*?:[]" else ch for ch in str(name)).strip() or "Sheet"
    cleaned = cleaned[:31]
    base = cleaned
    i = 2
    while cleaned in used:
        suffix = f"_{i}"
        cleaned = f"{base[: 31 - len(suffix)]}{suffix}"
        i += 1
    used.add(cleaned)
    return cleaned


def _empty_issues() -> pd.DataFrame:
    return pd.DataFrame(columns=ISSUE_COLUMNS)


def _write_issues_workbook(
    issues_df: pd.DataFrame,
    earlier_df: pd.DataFrame,
    path: Path,
) -> None:
    used_names: set[str] = set()
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # Overview — separate real issues from informational notes
        if issues_df.empty:
            overview = pd.DataFrame(
                columns=["kind", "sheet", "reason_category", "count"]
            )
        else:
            overview = (
                issues_df.groupby(["sheet", "reason_category"], dropna=False)
                .size()
                .reset_index(name="count")
            )
            overview["kind"] = overview["reason_category"].map(
                lambda c: "note" if is_note_category(c) else "issue"
            )
            overview = overview[["kind", "sheet", "reason_category", "count"]].sort_values(
                ["kind", "sheet", "count"], ascending=[True, True, False]
            )
        overview.to_excel(
            writer, sheet_name=_safe_sheet_name("Overview", used_names), index=False
        )

        # Notes-only and Issues-only summary tabs
        if not issues_df.empty:
            notes_df = issues_df[issues_df["reason_category"].map(is_note_category)]
            real_df = issues_df[~issues_df["reason_category"].map(is_note_category)]
        else:
            notes_df = _empty_issues()
            real_df = _empty_issues()
        real_df.to_excel(
            writer, sheet_name=_safe_sheet_name("Issues only", used_names), index=False
        )
        notes_df.to_excel(
            writer, sheet_name=_safe_sheet_name("Notes only", used_names), index=False
        )

        # Updates whose validity was cut to the RA agreement window (start and/or end)
        earlier_df.to_excel(
            writer,
            sheet_name=_safe_sheet_name("Agreement validity", used_names),
            index=False,
        )

        # One tab per update source sheet (BASE / ETSBAF)
        for sheet in UPDATE_SHEETS:
            part = issues_df[issues_df["sheet"] == sheet] if not issues_df.empty else _empty_issues()
            if part.empty and not issues_df.empty:
                part = _empty_issues()
            part.to_excel(
                writer, sheet_name=_safe_sheet_name(str(sheet), used_names), index=False
            )

        # One tab per reason category
        categories = (
            sorted(issues_df["reason_category"].dropna().unique().tolist())
            if not issues_df.empty
            else []
        )
        if not categories:
            _empty_issues().to_excel(
                writer, sheet_name=_safe_sheet_name("No issues", used_names), index=False
            )
        else:
            for cat in categories:
                part = issues_df[issues_df["reason_category"] == cat]
                tab = _safe_sheet_name(str(cat), used_names)
                part.to_excel(writer, sheet_name=tab, index=False)

        # Combined sheet+reason tabs for finer split
        if not issues_df.empty:
            pairs = (
                issues_df.groupby(["sheet", "reason_category"], dropna=False)
                .size()
                .reset_index(name="count")
            )
            for _, row in pairs.iterrows():
                sheet = str(row["sheet"])
                cat = str(row["reason_category"])
                part = issues_df[
                    (issues_df["sheet"] == sheet) & (issues_df["reason_category"] == cat)
                ]
                tab = _safe_sheet_name(f"{sheet} - {cat}", used_names)
                part.to_excel(writer, sheet_name=tab, index=False)


def write_reports(
    stats: UpdateStats,
    output_dir: Path,
    stamp: str,
    *,
    name_prefix: str | None = None,
) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths: dict[str, Path] = {}

    if name_prefix:
        issues_path = output_dir / f"{name_prefix}_issues_{stamp}.xlsx"
        changes_path = output_dir / f"{name_prefix}_changes_{stamp}.xlsx"
    else:
        issues_path = output_dir / f"update_issues_{stamp}.xlsx"
        changes_path = output_dir / f"update_changes_{stamp}.xlsx"

    issues_df = pd.DataFrame([asdict(i) for i in stats.issues])
    changes_df = pd.DataFrame([asdict(c) for c in stats.changes])
    earlier_df = pd.DataFrame([asdict(e) for e in stats.earlier_validity])

    if issues_df.empty:
        issues_df = _empty_issues()
    else:
        issues_df = issues_df.copy()
        issues_df["reason_category"] = issues_df["reason"].map(categorize_reason)
        for col in ISSUE_COLUMNS:
            if col not in issues_df.columns:
                issues_df[col] = None
        issues_df = issues_df[ISSUE_COLUMNS]

    if changes_df.empty:
        changes_df = pd.DataFrame(columns=CHANGE_COLUMNS)
    else:
        for col in CHANGE_COLUMNS:
            if col not in changes_df.columns:
                changes_df[col] = None
        changes_df = changes_df[CHANGE_COLUMNS]

    if earlier_df.empty:
        earlier_df = pd.DataFrame(columns=EARLIER_VALIDITY_COLUMNS)
    else:
        for col in EARLIER_VALIDITY_COLUMNS:
            if col not in earlier_df.columns:
                earlier_df[col] = None
        earlier_df = earlier_df[EARLIER_VALIDITY_COLUMNS]

    _write_issues_workbook(issues_df, earlier_df, issues_path)

    # Changes: split by update sheet tabs as well
    used_names: set[str] = set()
    with pd.ExcelWriter(changes_path, engine="openpyxl") as writer:
        changes_df.to_excel(
            writer, sheet_name=_safe_sheet_name("All", used_names), index=False
        )
        earlier_df.to_excel(
            writer, sheet_name=_safe_sheet_name("Earlier validity", used_names), index=False
        )
        dual_df = (
            changes_df[changes_df["status"].isin({"dual_container", "dual_validity"})]
            if not changes_df.empty
            else pd.DataFrame(columns=CHANGE_COLUMNS)
        )
        dual_df.to_excel(
            writer, sheet_name=_safe_sheet_name("Dual rates", used_names), index=False
        )
        for sheet in UPDATE_SHEETS:
            part = (
                changes_df[changes_df["sheet"] == sheet]
                if not changes_df.empty
                else pd.DataFrame(columns=CHANGE_COLUMNS)
            )
            part.to_excel(
                writer, sheet_name=_safe_sheet_name(str(sheet), used_names), index=False
            )

    paths["issues"] = issues_path
    paths["changes"] = changes_path
    return paths
