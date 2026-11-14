"""Classify raw logical diffs into cosmetic vs logical / rate vs currency-only."""

from __future__ import annotations

import re
from collections import Counter

from .keys import token_set


def refine_diff(raw: dict) -> dict:
    """
    Build a refined view on top of compare_workbooks() output.

    - Identity: cosmetic (same tokens / spacing) vs logical
    - Costs: real rate mismatches vs currency-only vs min-only
    """
    cosmetic_id: list[dict] = []
    logical_id: list[dict] = []
    for item in raw.get("identity_diffs", []):
        cos: list[dict] = []
        log: list[dict] = []
        for f in item.get("fields", []):
            a = f.get("baseline") or ""
            b = f.get("candidate") or ""
            if token_set(a) == token_set(b) and token_set(a):
                cos.append(f)
            elif re.sub(r"\s+", "", a) == re.sub(r"\s+", "", b):
                cos.append(f)
            else:
                log.append(f)
        if cos and not log:
            cosmetic_id.append({**item, "fields": cos})
        elif log:
            logical_id.append({**item, "fields": log, "cosmetic_also": cos})

    rate_real: list[dict] = []
    currency_only: list[dict] = []
    min_only: list[dict] = []
    for item in raw.get("cost_diffs", []):
        for diff in item.get("diffs", []):
            entry = {
                "baseline_row": item["baseline_row"],
                "candidate_row": item["candidate_row"],
                "key": item["key"],
                "valid": item["valid"],
                **diff,
            }
            if diff.get("diff_rate"):
                rate_real.append(entry)
            elif diff.get("diff_min"):
                min_only.append(entry)
            elif diff.get("diff_currency"):
                currency_only.append(entry)

    base_presence = raw.get("cost_only_on_lane_baseline", [])
    cand_presence = raw.get("cost_only_on_lane_candidate", [])

    return {
        "meta": raw["meta"],
        "summary": {
            **raw["summary"],
            "identity_cosmetic_only": len(cosmetic_id),
            "identity_logical": len(logical_id),
            "real_rate_mismatches": len(rate_real),
            "currency_only_cost_diffs": len(currency_only),
            "min_only_cost_diffs": len(min_only),
        },
        "lanes_only_baseline": raw.get("lanes_only_baseline", []),
        "lanes_only_candidate": raw.get("lanes_only_candidate", []),
        "ambiguous": raw.get("ambiguous", []),
        "cost_headers_only_baseline": raw.get("cost_headers_only_baseline", []),
        "cost_headers_only_candidate": raw.get("cost_headers_only_candidate", []),
        "identity_logical": logical_id,
        "identity_cosmetic_sample": cosmetic_id[:20],
        "rate_mismatches": rate_real,
        "currency_only_sample": currency_only[:40],
        "currency_only_by_fee": dict(Counter(x["fee"] for x in currency_only)),
        "cost_only_on_lane_baseline_by_fee": dict(
            Counter(x["fee"] for x in base_presence)
        ),
        "cost_only_on_lane_candidate_by_fee": dict(
            Counter(x["fee"] for x in cand_presence)
        ),
        "cost_only_on_lane_baseline_sample": base_presence[:25],
        "cost_only_on_lane_candidate_sample": cand_presence[:25],
    }
