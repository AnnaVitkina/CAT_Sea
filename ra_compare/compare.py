"""
Logical RA comparison: match lanes by identity (not row #) and costs by
fee/equipment/period (not column #).
"""

from __future__ import annotations

import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from ra_update.config import (
    COL_SERVICE,
    COL_SERVICE_SHORT,
    COL_TRANSPOREON,
)
from ra_update.parsing import to_number
from ra_update.ra_model import RaLane, RateCardWorkbook

from .config import COST_PRESENCE_CAP
from .keys import (
    cost_logical_key,
    date_str,
    lane_fingerprint,
    lane_soft_key,
    norm_tokens,
    token_set,
)


def _identity_equivalent(a: object, b: object) -> bool:
    """
    True when identity strings match logically.
    'A / B' vs 'A/B' (spaces around /) is not a real difference.
    """
    sa = "" if a is None else str(a).strip()
    sb = "" if b is None else str(b).strip()
    if sa == sb:
        return True
    ta, tb = token_set(sa), token_set(sb)
    if ta and ta == tb:
        return True
    return re.sub(r"\s+", "", sa) == re.sub(r"\s+", "", sb)


def _cell_num(ra: RateCardWorkbook, row: int, col: int | None):
    if col is None:
        return None
    return to_number(ra.get_value(row, col))


def _cell_rate(ra: RateCardWorkbook, row: int, col: int | None):
    """
    Numeric rate when possible; otherwise keep non-empty text
    (e.g. red dual-validity '520 (…) / 548 (…)').
    """
    if col is None:
        return None
    v = ra.get_value(row, col)
    num = to_number(v)
    if num is not None:
        return num
    if v is None:
        return None
    text = str(v).strip()
    return text or None


def _cell_str(ra: RateCardWorkbook, row: int, col: int | None) -> str:
    if col is None:
        return ""
    v = ra.get_value(row, col)
    return "" if v is None else str(v).strip()


def _lane_costs(ra: RateCardWorkbook, lane: RaLane) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for cc in ra.cost_columns:
        rate = _cell_rate(ra, lane.row, cc.rate_col)
        mn = _cell_num(ra, lane.row, cc.min_col)
        cur = _cell_str(ra, lane.row, cc.currency_col)
        if rate is None and mn is None and not cur:
            continue
        lk = cost_logical_key(cc)
        if lk in out and out[lk].get("rate") is not None:
            continue
        out[lk] = {
            "header": cc.header,
            "fee": cc.fee_name,
            "equipment": cc.equipment,
            "valid_from": date_str(cc.valid_from),
            "valid_to": date_str(cc.valid_to),
            "rate": rate,
            "min": mn,
            "currency": cur,
            "col": cc.start_col,
        }
    return out


def _lane_brief(lane: RaLane, label: str) -> dict:
    return {
        "side": label,
        "row": lane.row,
        "lane_no": lane.lane_no,
        "key": lane.key_raw,
        "carrier": lane.carrier,
        "service_c": lane.service_c,
        "valid_from": date_str(lane.valid_from),
        "valid_to": date_str(lane.valid_to),
        "origin": f"{lane.origin_port}|{lane.origin_postal}|{lane.origin_country}",
        "dest": f"{lane.dest_port}|{lane.dest_postal}|{lane.dest_country}",
    }


def _identity_fields(ra: RateCardWorkbook, lane: RaLane) -> dict[str, str]:
    return {
        "KEY": lane.key_raw,
        "Carrier": lane.carrier,
        "SERVICE": _cell_str(ra, lane.row, COL_SERVICE),
        "SERVICE__C": lane.service_c,
        "Service": _cell_str(ra, lane.row, COL_SERVICE_SHORT),
        "ValidFrom": date_str(lane.valid_from),
        "ValidTo": date_str(lane.valid_to),
        "OriginPort": lane.origin_port,
        "OriginPostal": lane.origin_postal,
        "OriginCountry": lane.origin_country,
        "DestPort": lane.dest_port,
        "DestPostal": lane.dest_postal,
        "DestCountry": lane.dest_country,
        "Transporeon": _cell_str(ra, lane.row, COL_TRANSPOREON),
    }


def _nums_equal(a, b) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    na, nb = to_number(a), to_number(b)
    if na is not None and nb is not None:
        return abs(float(na) - float(nb)) < 1e-9
    # Dual-validity / other non-numeric rate text
    return str(a).strip() == str(b).strip()


def _cost_effectively_blank(cost: dict) -> bool:
    """
    True when a one-sided cost cell is empty/zero (no meaningful rate or min).

    Empty baseline vs candidate rate 0 (+ optional currency) is not a real
    presence difference — treat like both blank.
    """
    if cost.get("min") is not None:
        return False
    rate = cost.get("rate")
    if rate is None:
        return True
    num = to_number(rate)
    if num is not None:
        return abs(float(num)) < 1e-9
    return False


def match_lanes(
    a_lanes: list[RaLane],
    b_lanes: list[RaLane],
    *,
    baseline_label: str = "baseline",
    candidate_label: str = "candidate",
) -> tuple[list[tuple[RaLane, RaLane]], list[RaLane], list[RaLane], list[dict]]:
    """Exact fingerprint match first, then soft KEY-overlap match for leftovers."""
    a_by_fp: dict[str, list[RaLane]] = defaultdict(list)
    b_by_fp: dict[str, list[RaLane]] = defaultdict(list)
    for L in a_lanes:
        a_by_fp[lane_fingerprint(L)].append(L)
    for L in b_lanes:
        b_by_fp[lane_fingerprint(L)].append(L)

    paired: list[tuple[RaLane, RaLane]] = []
    used_a: set[int] = set()
    used_b: set[int] = set()
    ambiguous: list[dict] = []

    for fp in sorted(set(a_by_fp) & set(b_by_fp)):
        aa = a_by_fp[fp]
        bb = b_by_fp[fp]
        n = min(len(aa), len(bb))
        for i in range(n):
            paired.append((aa[i], bb[i]))
            used_a.add(aa[i].row)
            used_b.add(bb[i].row)
        if len(aa) != len(bb):
            ambiguous.append(
                {
                    "type": "fingerprint_count_mismatch",
                    "fingerprint": fp[:200],
                    f"{baseline_label}_count": len(aa),
                    f"{candidate_label}_count": len(bb),
                    f"{baseline_label}_rows": [x.row for x in aa],
                    f"{candidate_label}_rows": [x.row for x in bb],
                }
            )

    rem_a = [L for L in a_lanes if L.row not in used_a]
    rem_b = [L for L in b_lanes if L.row not in used_b]

    b_index: dict[str, list[RaLane]] = defaultdict(list)
    for L in rem_b:
        sk = lane_soft_key(L)
        for tok in norm_tokens(L.key_raw):
            b_index[f"{sk}|{tok}"].append(L)

    for La in rem_a:
        sk = lane_soft_key(La)
        candidates: list[RaLane] = []
        seen: set[int] = set()
        for tok in norm_tokens(La.key_raw):
            for Lb in b_index.get(f"{sk}|{tok}", []):
                if Lb.row in used_b or Lb.row in seen:
                    continue
                seen.add(Lb.row)
                candidates.append(Lb)
        if len(candidates) == 1:
            Lb = candidates[0]
            paired.append((La, Lb))
            used_a.add(La.row)
            used_b.add(Lb.row)
        elif len(candidates) > 1:
            ambiguous.append(
                {
                    "type": "soft_ambiguous",
                    f"{baseline_label}_row": La.row,
                    f"{baseline_label}_key": La.key_raw,
                    f"{baseline_label}_valid": (
                        f"{date_str(La.valid_from)}..{date_str(La.valid_to)}"
                    ),
                    "candidates": [
                        {"row": c.row, "key": c.key_raw, "service": c.service_c}
                        for c in candidates[:8]
                    ],
                }
            )

    only_a = [L for L in a_lanes if L.row not in used_a]
    only_b = [L for L in b_lanes if L.row not in used_b]
    return paired, only_a, only_b, ambiguous


def compare_workbooks(
    baseline_path: Path,
    candidate_path: Path,
    *,
    baseline_label: str = "baseline",
    candidate_label: str = "candidate",
) -> dict:
    """
    Compare two Rate Card workbooks logically.

    JSON uses baseline_label / candidate_label in field names where helpful,
    plus stable keys: lanes_only_baseline, lanes_only_candidate, etc.
    """
    print(f"Loading baseline ({baseline_label}): {baseline_path.name}")
    a = RateCardWorkbook(baseline_path)
    print(f"Loading candidate ({candidate_label}): {candidate_path.name}")
    b = RateCardWorkbook(candidate_path)

    a_cost_keys = {cost_logical_key(c): c.header for c in a.cost_columns}
    b_cost_keys = {cost_logical_key(c): c.header for c in b.cost_columns}
    costs_only_base = sorted(set(a_cost_keys) - set(b_cost_keys))
    costs_only_cand = sorted(set(b_cost_keys) - set(a_cost_keys))

    paired, only_base, only_cand, ambiguous = match_lanes(
        a.lanes,
        b.lanes,
        baseline_label=baseline_label,
        candidate_label=candidate_label,
    )

    identity_diffs: list[dict] = []
    cost_diffs: list[dict] = []
    cost_only_base: list[dict] = []
    cost_only_cand: list[dict] = []
    matched_identical = 0

    for La, Lb in paired:
        id_a = _identity_fields(a, La)
        id_b = _identity_fields(b, Lb)
        field_diffs = []
        for f in id_a:
            if _identity_equivalent(id_a[f], id_b[f]):
                continue
            field_diffs.append(
                {
                    "field": f,
                    "baseline": id_a[f],
                    "candidate": id_b[f],
                }
            )

        ca = _lane_costs(a, La)
        cb = _lane_costs(b, Lb)
        shared = set(ca) & set(cb)
        only_a = set(ca) - set(cb)
        only_b = set(cb) - set(ca)

        lane_cost_diffs = []
        for k in sorted(shared):
            va, vb = ca[k], cb[k]
            rate_diff = not _nums_equal(va["rate"], vb["rate"])
            min_diff = not _nums_equal(va["min"], vb["min"])
            cur_diff = (va["currency"] or "") != (vb["currency"] or "")
            if rate_diff or min_diff or cur_diff:
                lane_cost_diffs.append(
                    {
                        "fee": va["fee"],
                        "equipment": va["equipment"],
                        "period": f"{va['valid_from']}..{va['valid_to']}",
                        "header_baseline": va["header"],
                        "header_candidate": vb["header"],
                        "rate_baseline": va["rate"],
                        "rate_candidate": vb["rate"],
                        "min_baseline": va["min"],
                        "min_candidate": vb["min"],
                        "currency_baseline": va["currency"],
                        "currency_candidate": vb["currency"],
                        "diff_rate": rate_diff,
                        "diff_min": min_diff,
                        "diff_currency": cur_diff,
                    }
                )

        meaningful_only_a = False
        for k in sorted(only_a):
            if _cost_effectively_blank(ca[k]):
                continue
            meaningful_only_a = True
            cost_only_base.append(
                {
                    **_lane_brief(La, "baseline"),
                    "candidate_row": Lb.row,
                    **{
                        kk: ca[k][kk]
                        for kk in ("fee", "equipment", "header", "rate", "min", "currency")
                    },
                    "period": f"{ca[k]['valid_from']}..{ca[k]['valid_to']}",
                }
            )
        meaningful_only_b = False
        for k in sorted(only_b):
            if _cost_effectively_blank(cb[k]):
                continue
            meaningful_only_b = True
            cost_only_cand.append(
                {
                    **_lane_brief(Lb, "candidate"),
                    "baseline_row": La.row,
                    **{
                        kk: cb[k][kk]
                        for kk in ("fee", "equipment", "header", "rate", "min", "currency")
                    },
                    "period": f"{cb[k]['valid_from']}..{cb[k]['valid_to']}",
                }
            )

        if field_diffs:
            identity_diffs.append(
                {
                    "baseline_row": La.row,
                    "candidate_row": Lb.row,
                    "lane_no_baseline": La.lane_no,
                    "lane_no_candidate": Lb.lane_no,
                    "key_baseline": La.key_raw[:120],
                    "key_candidate": Lb.key_raw[:120],
                    "valid": f"{date_str(La.valid_from)}..{date_str(La.valid_to)}",
                    "fields": field_diffs,
                }
            )
        if lane_cost_diffs:
            cost_diffs.append(
                {
                    "baseline_row": La.row,
                    "candidate_row": Lb.row,
                    "lane_no_baseline": La.lane_no,
                    "lane_no_candidate": Lb.lane_no,
                    "key": La.key_raw[:120],
                    "valid": f"{date_str(La.valid_from)}..{date_str(La.valid_to)}",
                    "diffs": lane_cost_diffs,
                }
            )
        if (
            not field_diffs
            and not lane_cost_diffs
            and not meaningful_only_a
            and not meaningful_only_b
        ):
            matched_identical += 1

    fee_rate_mismatch: dict[str, int] = defaultdict(int)
    for item in cost_diffs:
        for d in item["diffs"]:
            if d["diff_rate"]:
                fee_rate_mismatch[d["fee"]] += 1

    return {
        "meta": {
            "baseline": baseline_path.name,
            "candidate": candidate_path.name,
            "baseline_label": baseline_label,
            "candidate_label": candidate_label,
            "baseline_path": str(baseline_path),
            "candidate_path": str(candidate_path),
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "baseline_lanes": len(a.lanes),
            "candidate_lanes": len(b.lanes),
            "baseline_cost_columns": len(a.cost_columns),
            "candidate_cost_columns": len(b.cost_columns),
            "agreement_baseline": a.agreement_title,
            "agreement_candidate": b.agreement_title,
            "agreement_valid_baseline": (
                f"{date_str(a.agreement_valid_from)}..{date_str(a.agreement_valid_to)}"
            ),
            "agreement_valid_candidate": (
                f"{date_str(b.agreement_valid_from)}..{date_str(b.agreement_valid_to)}"
            ),
        },
        "summary": {
            "paired_lanes": len(paired),
            "matched_identical": matched_identical,
            "lanes_only_baseline": len(only_base),
            "lanes_only_candidate": len(only_cand),
            "identity_diff_lanes": len(identity_diffs),
            "cost_value_diff_lanes": len(cost_diffs),
            "lane_costs_only_baseline": len(cost_only_base),
            "lane_costs_only_candidate": len(cost_only_cand),
            "cost_headers_only_baseline": len(costs_only_base),
            "cost_headers_only_candidate": len(costs_only_cand),
            "ambiguous_matches": len(ambiguous),
            "rate_mismatches_by_fee": dict(
                sorted(fee_rate_mismatch.items(), key=lambda x: -x[1])
            ),
        },
        "cost_headers_only_baseline": [
            {"key": k, "header": a_cost_keys[k]} for k in costs_only_base
        ],
        "cost_headers_only_candidate": [
            {"key": k, "header": b_cost_keys[k]} for k in costs_only_cand
        ],
        "lanes_only_baseline": [_lane_brief(L, "baseline") for L in only_base],
        "lanes_only_candidate": [_lane_brief(L, "candidate") for L in only_cand],
        "ambiguous": ambiguous,
        "identity_diffs": identity_diffs,
        "cost_diffs": cost_diffs,
        "cost_only_on_lane_baseline": cost_only_base[:COST_PRESENCE_CAP],
        "cost_only_on_lane_candidate": cost_only_cand[:COST_PRESENCE_CAP],
        "cost_only_on_lane_baseline_truncated": max(
            0, len(cost_only_base) - COST_PRESENCE_CAP
        ),
        "cost_only_on_lane_candidate_truncated": max(
            0, len(cost_only_cand) - COST_PRESENCE_CAP
        ),
    }
