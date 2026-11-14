"""Write compare JSON / XLSX outputs and print a console summary."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd


def write_json(path: Path, payload: dict) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return path


def _safe_sheet(name: str, used: set[str]) -> str:
    base = re.sub(r"[:\\/?*\[\]]", " ", name).strip() or "Sheet"
    base = base[:31]
    out = base
    n = 2
    while out in used:
        suffix = f"_{n}"
        out = (base[: 31 - len(suffix)] + suffix)[:31]
        n += 1
    used.add(out)
    return out


def _records_to_df(rows: list | dict | None) -> pd.DataFrame:
    if rows is None:
        return pd.DataFrame()
    if isinstance(rows, dict):
        return pd.DataFrame([{"key": k, "value": v} for k, v in rows.items()])
    if not rows:
        return pd.DataFrame()
    flat: list[dict] = []
    for item in rows:
        if not isinstance(item, dict):
            flat.append({"value": item})
            continue
        fields = item.get("fields")
        if isinstance(fields, list) and fields:
            base = {k: v for k, v in item.items() if k != "fields"}
            for f in fields:
                row = dict(base)
                if isinstance(f, dict):
                    row["field"] = f.get("field")
                    row["field_baseline"] = f.get("baseline")
                    row["field_candidate"] = f.get("candidate")
                flat.append(row)
        else:
            flat.append(dict(item))
    return pd.DataFrame(flat)


def write_xlsx(path: Path, refined: dict, *, raw: dict | None = None) -> Path:
    """Write refined (and key raw presence lists) compare result to xlsx."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    used: set[str] = set()

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        meta = refined.get("meta") or {}
        summary = refined.get("summary") or {}
        overview = pd.DataFrame(
            [{"section": "meta", "key": k, "value": v} for k, v in meta.items()]
            + [
                {"section": "summary", "key": k, "value": v}
                for k, v in summary.items()
            ]
        )
        overview.to_excel(
            writer, sheet_name=_safe_sheet("Summary", used), index=False
        )

        sheets: list[tuple[str, object]] = [
            ("Rate mismatches", refined.get("rate_mismatches")),
            ("Currency only", refined.get("currency_only_sample")),
            ("Identity logical", refined.get("identity_logical")),
            ("Identity cosmetic", refined.get("identity_cosmetic_sample")),
            ("Lanes only baseline", refined.get("lanes_only_baseline")),
            ("Lanes only candidate", refined.get("lanes_only_candidate")),
            ("Headers only baseline", refined.get("cost_headers_only_baseline")),
            ("Headers only candidate", refined.get("cost_headers_only_candidate")),
            (
                "Costs only baseline",
                refined.get("cost_only_on_lane_baseline_sample"),
            ),
            (
                "Costs only candidate",
                refined.get("cost_only_on_lane_candidate_sample"),
            ),
            ("Ambiguous", refined.get("ambiguous")),
        ]
        if raw:
            sheets.extend(
                [
                    (
                        "Raw costs only baseline",
                        raw.get("lane_costs_only_baseline"),
                    ),
                    (
                        "Raw costs only candidate",
                        raw.get("lane_costs_only_candidate"),
                    ),
                ]
            )

        for title, rows in sheets:
            df = _records_to_df(rows if isinstance(rows, list) else None)
            if df.empty:
                df = pd.DataFrame({"info": ["(none)"]})
            df.to_excel(writer, sheet_name=_safe_sheet(title, used), index=False)

        for title, data in (
            ("Currency only by fee", refined.get("currency_only_by_fee")),
            (
                "Presence baseline by fee",
                refined.get("cost_only_on_lane_baseline_by_fee"),
            ),
            (
                "Presence candidate by fee",
                refined.get("cost_only_on_lane_candidate_by_fee"),
            ),
            (
                "Rate mismatch by fee",
                (refined.get("summary") or {}).get("rate_mismatches_by_fee")
                or (raw or {}).get("summary", {}).get("rate_mismatches_by_fee"),
            ),
        ):
            df = _records_to_df(data if isinstance(data, dict) else None)
            if df.empty:
                continue
            df.to_excel(writer, sheet_name=_safe_sheet(title, used), index=False)

    return path


def print_summary(raw: dict, refined: dict | None = None) -> None:
    m = raw["meta"]
    s = raw["summary"]
    rs = (refined or {}).get("summary", {})

    print("\n=== LOGICAL DIFF SUMMARY ===")
    print(
        f"Baseline ({m.get('baseline_label', 'baseline')}): {m['baseline_lanes']} lanes · "
        f"{m['baseline']}"
    )
    print(
        f"Candidate ({m.get('candidate_label', 'candidate')}): {m['candidate_lanes']} lanes · "
        f"{m['candidate']}"
    )
    print(f"Paired: {s['paired_lanes']} | identical: {s['matched_identical']}")
    print(
        f"Only baseline: {s['lanes_only_baseline']} | "
        f"Only candidate: {s['lanes_only_candidate']}"
    )
    print(
        f"Identity diffs: {s['identity_diff_lanes']} | "
        f"Cost value diffs: {s['cost_value_diff_lanes']}"
    )
    if rs:
        print(
            f"  identity logical: {rs.get('identity_logical', '?')} · "
            f"cosmetic: {rs.get('identity_cosmetic_only', '?')}"
        )
        print(
            f"  real rate mismatches: {rs.get('real_rate_mismatches', '?')} · "
            f"currency-only: {rs.get('currency_only_cost_diffs', '?')}"
        )
    print(
        f"Costs only on baseline lane: {s['lane_costs_only_baseline']} | "
        f"only on candidate lane: {s['lane_costs_only_candidate']}"
    )
    print(
        f"Cost headers only baseline: {s['cost_headers_only_baseline']} | "
        f"only candidate: {s['cost_headers_only_candidate']}"
    )
    print(f"Ambiguous: {s['ambiguous_matches']}")
    print(f"Rate mismatches by fee: {s['rate_mismatches_by_fee']}")
