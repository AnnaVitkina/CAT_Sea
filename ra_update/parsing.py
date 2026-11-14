from __future__ import annotations

import pandas as pd

from .charge_rules import ChargeRule, resolve_rate_column
from .config import IGNORE_MIN_ABS_THRESHOLD


def split_multi(value: object) -> list[str]:
    """Split slash-separated multi-values used in KEY / postal fields."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    text = str(value).strip()
    if not text:
        return []
    parts = [p.strip() for p in text.split("/")]
    return [p for p in parts if p]


def value_in_multi(needle: object, haystack: object) -> bool:
    needles = split_multi(needle)
    options = split_multi(haystack)
    if not needles:
        return True  # nothing to validate
    if not options:
        return False
    return any(n in options for n in needles)


def is_ignored_min(value: object) -> bool:
    try:
        return abs(float(value)) >= IGNORE_MIN_ABS_THRESHOLD
    except (TypeError, ValueError):
        return False


def to_number(value: object) -> float | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        value = text
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def numbers_equal(a: object, b: object, tol: float = 1e-9) -> bool:
    na, nb = to_number(a), to_number(b)
    if na is None or nb is None:
        return str(a).strip() == str(b).strip()
    return abs(na - nb) <= tol


def _companion_min_col(rate_col: str) -> str:
    if rate_col.upper().endswith("_RATE"):
        return rate_col[: -len("_RATE")] + "_MIN"
    return rate_col + "_MIN"


def iter_update_rates(row: pd.Series, sheet: str) -> list[dict]:
    """
    Extract actionable rate updates from one update row.

    Each item:
      equipment, charge_type, rate_col, new_value, min_value, min_col,
      rule, unsupported, reason
    """
    results: list[dict] = []
    row_cols = {str(c) for c in row.index}

    for col, raw in row.items():
        col_s = str(col)
        if not col_s.endswith("_RATE"):
            continue
        num = to_number(raw)
        if num is None:
            continue

        min_col = _companion_min_col(col_s)
        min_raw = row.get(min_col) if min_col in row_cols else None
        min_num = to_number(min_raw)
        if min_num is not None and is_ignored_min(min_num):
            min_num = None

        rule, equipment, skip_reason = resolve_rate_column(col_s, sheet)
        if skip_reason:
            results.append(
                {
                    "equipment": equipment,
                    "charge_type": rule.charge_key if rule else None,
                    "rate_col": col_s,
                    "new_value": num,
                    "min_value": min_num,
                    "min_col": min_col if min_num is not None else None,
                    "rule": rule,
                    "unsupported": True,
                    "reason": skip_reason,
                }
            )
            continue

        assert rule is not None
        results.append(
            {
                "equipment": equipment,
                "charge_type": rule.charge_key,
                "rate_col": col_s,
                "new_value": num,
                "min_value": min_num,
                "min_col": min_col if min_num is not None else None,
                "rule": rule,
                "unsupported": False,
                "reason": None,
            }
        )

    return results
