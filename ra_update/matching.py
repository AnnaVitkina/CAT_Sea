from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, timedelta

from .dates import format_ra_date, period_inside, periods_equal
from .lane_identity import locations_match
from .parsing import split_multi, value_in_multi
from .ra_model import RaLane, RateCardWorkbook

# Insert _BU after the OC_* service token: ...-OC_CNTR_CY-CY-GB... → ...-OC_CNTR_CY-CY_BU-GB...
_BU_INSERT_RE = re.compile(r"(OC_[A-Z0-9]+_[A-Z0-9]+-[A-Z0-9]+)-", re.IGNORECASE)


def non_bu_key(key: str) -> str:
    """ONEY-...-CY-CY_BU-GBSOU-... → ONEY-...-CY-CY-GBSOU-..."""
    return (key or "").replace("_BU-", "-", 1)


def bu_key(key: str) -> str | None:
    """ONEY-...-CY-CY-GBSOU-... → ONEY-...-CY-CY_BU-GBSOU-... (None if already BU)."""
    k = (key or "").strip()
    if not k or "_BU-" in k:
        return None
    m = _BU_INSERT_RE.search(k)
    if not m:
        return None
    return k[: m.end(1)] + "_BU" + k[m.end(1) :]


def service_without_bu(service_c: str) -> str:
    """OC_CNTR_CY-CY_BU → OC_CNTR_CY-CY"""
    s = (service_c or "").strip()
    if s.upper().endswith("_BU"):
        return s[: -len("_BU")]
    return s.replace("_BU", "")


def find_lanes_for_key(ra: RateCardWorkbook, key: str) -> tuple[list[RaLane], str | None]:
    """
    Resolve RA lanes for an update KEY.

    1) Exact KEY (including a part of slash-combined KEY cells)
    2) If missing: look for the key / BU↔non-BU sibling in slash-combined KEY rows
    3) If still missing:
         - BU update → standalone non-BU sibling (update that lane; slash-append KEY/SERVICE)
         - non-BU update → standalone BU sibling (same; do not create a second lane)

    When resolved via slash / sibling, caller should apply BU preference write rules
    (prefer non-BU KEY value when validities match).
    """
    key = (key or "").strip()
    if not key:
        return [], None

    direct = ra.find_by_key(key)
    if direct:
        return list(direct), None

    sibling_non_bu = non_bu_key(key) if "_BU-" in key else None
    sibling_bu = bu_key(key) if "_BU-" not in key else None

    def _dedupe(lanes: list[RaLane]) -> list[RaLane]:
        seen: set[int] = set()
        out: list[RaLane] = []
        for lane in lanes:
            if lane.row in seen:
                continue
            seen.add(lane.row)
            out.append(lane)
        return out

    # Slash-combined KEY rows containing this key or its BU↔non-BU sibling
    slash_hits: list[RaLane] = []
    for lane in ra.lanes:
        if "/" not in (lane.key_raw or ""):
            continue
        parts = set(lane.keys)
        if key in parts:
            slash_hits.append(lane)
            continue
        if sibling_non_bu and sibling_non_bu in parts:
            slash_hits.append(lane)
            continue
        if sibling_bu and sibling_bu in parts:
            slash_hits.append(lane)
            continue
        # non-BU update: slash row has BU part that strips to this key
        if "_BU-" not in key:
            for part in parts:
                if "_BU-" in part and non_bu_key(part) == key:
                    slash_hits.append(lane)
                    break

    if slash_hits:
        via = (
            sibling_non_bu
            if sibling_non_bu and all(key not in l.keys for l in slash_hits)
            else key
        )
        return _dedupe(slash_hits), (
            f"KEY resolved via slash-combined KEY row (matched {via}; "
            f"apply BU preference vs non-BU keys)"
        )

    # BU update with no slash pair: use standalone non-BU sibling lanes
    if sibling_non_bu:
        sib_lanes = ra.find_by_key(sibling_non_bu)
        if sib_lanes:
            return list(sib_lanes), (
                f"KEY resolved via non-BU sibling {sibling_non_bu} "
                f"(apply BU preference vs non-BU keys)"
            )

    # non-BU update with no slash pair: use standalone BU sibling (do not add a 2nd lane)
    if sibling_bu:
        sib_lanes = ra.find_by_key(sibling_bu)
        if sib_lanes:
            return list(sib_lanes), (
                f"KEY resolved via BU sibling {sibling_bu} "
                f"(apply BU preference vs non-BU keys)"
            )

    return [], None


@dataclass
class MatchResult:
    lanes: list[RaLane]
    ok: bool
    reason: str | None
    # When True: ETS may be written even though update/cost validity is wider
    # than the matched lane(s). BAF must not use this path.
    ets_wider_than_lane: bool = False


def lanes_overlapping_update(
    update_from: date,
    update_to: date,
    candidates: list[RaLane],
) -> list[RaLane]:
    """Lanes whose validity overlaps the update period (any shared days)."""
    out: list[RaLane] = []
    for lane in candidates:
        if lane.valid_from is None or lane.valid_to is None:
            continue
        if lane.valid_from <= update_to and lane.valid_to >= update_from:
            out.append(lane)
    out.sort(key=lambda l: (l.valid_from, l.valid_to, l.row))
    return out


def _country_ok(update_val: str, ra_val: str) -> bool:
    if not update_val:
        return True
    return update_val.strip().upper() == (ra_val or "").strip().upper()


def _origin_ok(update_origin: str, lane: RaLane) -> bool:
    """Origin from update may match Origin Port or Origin Postal (incl. aliases)."""
    return locations_match(update_origin, lane.origin_port, lane.origin_postal)


def _dest_ok(update_dest: str, lane: RaLane) -> bool:
    """Dest from update may match Destination Port or Postal (incl. aliases)."""
    return locations_match(update_dest, lane.dest_port, lane.dest_postal)


def validate_lane_details_excluding_validity(
    lane: RaLane,
    *,
    carrier: str,
    service_c: str,
    origin_country: str,
    dest_country: str,
    origin_code: str,
    dest_code: str,
) -> tuple[bool, str | None]:
    if carrier and lane.carrier and carrier.strip().upper() != lane.carrier.strip().upper():
        return False, f"Carrier mismatch: update={carrier!r} ra={lane.carrier!r}"

    if service_c and lane.service_c:
        if not value_in_multi(service_c, lane.service_c) and service_c.strip() != lane.service_c.strip():
            # BU ↔ non-BU are the same lane family (slash-append KEY/SERVICE later)
            u = service_without_bu(service_c).strip().upper()
            ra_norm = {service_without_bu(p).strip().upper() for p in split_multi(lane.service_c)}
            if not u or u not in ra_norm:
                return False, f"SERVICE__C mismatch: update={service_c!r} ra={lane.service_c!r}"

    if not _country_ok(origin_country, lane.origin_country):
        return False, (
            f"ORIGIN_COUNTRY mismatch: update={origin_country!r} ra={lane.origin_country!r}"
        )

    if not _country_ok(dest_country, lane.dest_country):
        return False, (
            f"DESTINATION_COUNTRY mismatch: update={dest_country!r} ra={lane.dest_country!r}"
        )

    if not _origin_ok(origin_code, lane):
        return False, (
            f"Origin mismatch: update={origin_code!r} "
            f"ra_port={lane.origin_port!r} ra_postal={lane.origin_postal!r}"
        )

    if not _dest_ok(dest_code, lane):
        return False, (
            f"Destination mismatch: update={dest_code!r} "
            f"ra_port={lane.dest_port!r} ra_postal={lane.dest_postal!r}"
        )

    if lane.valid_from is None or lane.valid_to is None:
        return False, "RA lane missing Valid from/Valid to"

    return True, None


def lanes_cover_without_gaps(
    update_from: date,
    update_to: date,
    candidates: list[RaLane],
) -> list[RaLane] | None:
    """
    Select lanes fully inside the update period whose validities cover it
    without gaps (including lanes created by prior splits/cuts).
    """
    inside = [
        lane
        for lane in candidates
        if lane.valid_from is not None
        and lane.valid_to is not None
        and period_inside(lane.valid_from, lane.valid_to, update_from, update_to)
    ]
    if not inside:
        return None

    inside.sort(key=lambda l: (l.valid_from, l.valid_to, l.row))

    selected: list[RaLane] = []
    covered_until: date | None = None

    for lane in inside:
        assert lane.valid_from is not None and lane.valid_to is not None
        if covered_until is None:
            if lane.valid_from > update_from:
                return None
            selected.append(lane)
            covered_until = lane.valid_to
        else:
            if lane.valid_from > covered_until + timedelta(days=1):
                return None
            if lane.valid_to <= covered_until:
                continue
            selected.append(lane)
            covered_until = lane.valid_to

        if covered_until >= update_to:
            return selected

    return None


def select_lane(
    candidates: list[RaLane],
    *,
    carrier: str,
    service_c: str,
    origin_country: str,
    dest_country: str,
    origin_code: str,
    dest_code: str,
    update_from,
    update_to,
    mode: str,
) -> MatchResult:
    """
    Match one or more RA lanes for an update.

    Order:
      1) exact validity match
      2) contiguous set of lanes inside the update that cover it without gaps
      3) single wider lane that contains the update (caller will cut/duplicate)
    """
    if not candidates:
        return MatchResult([], False, "No such lane (KEY not found in Rate Card)")

    if update_from is None or update_to is None:
        return MatchResult([], False, "Update row missing RATE_EFFECTIVE/EXPIRATION dates")

    detail_ok: list[RaLane] = []
    detail_failures: list[str] = []
    for lane in candidates:
        ok, reason = validate_lane_details_excluding_validity(
            lane,
            carrier=carrier,
            service_c=service_c,
            origin_country=origin_country,
            dest_country=dest_country,
            origin_code=origin_code,
            dest_code=dest_code,
        )
        if ok:
            detail_ok.append(lane)
        else:
            detail_failures.append(f"lane#{lane.lane_no}: {reason}")

    if not detail_ok:
        joined = " | ".join(detail_failures[:5])
        extra = "" if len(detail_failures) <= 5 else f" (+{len(detail_failures) - 5} more)"
        return MatchResult(
            [],
            False,
            f"KEY found but details did not match: {joined}{extra}",
        )

    # 1) Exact validity match
    exact = [
        lane
        for lane in detail_ok
        if periods_equal(update_from, update_to, lane.valid_from, lane.valid_to)
    ]
    if len(exact) == 1:
        return MatchResult(exact, True, None)
    if len(exact) > 1:
        nos = ", ".join(m.lane_no for m in exact)
        return MatchResult(
            [],
            False,
            f"Ambiguous: multiple RA lanes with identical validity ({nos})",
        )

    mode_u = mode.upper()

    # 2) Multi-lane cover without gaps (lanes fully inside the update)
    covered = lanes_cover_without_gaps(update_from, update_to, detail_ok)
    if covered:
        return MatchResult(covered, True, None)

    # 3) Update lies inside a single wider lane -> caller cuts/duplicates
    inside_one = [
        lane
        for lane in detail_ok
        if period_inside(update_from, update_to, lane.valid_from, lane.valid_to)
        and not periods_equal(update_from, update_to, lane.valid_from, lane.valid_to)
    ]
    if len(inside_one) == 1:
        return MatchResult(inside_one, True, None)
    if len(inside_one) > 1:
        nos = ", ".join(m.lane_no for m in inside_one)
        return MatchResult(
            [],
            False,
            f"Ambiguous: multiple RA lanes contain update period ({nos})",
        )

    # Explain failure
    validity_bits = []
    for lane in detail_ok:
        validity_bits.append(
            f"lane#{lane.lane_no}: {format_ra_date(lane.valid_from)}..{format_ra_date(lane.valid_to)}"
        )
    joined = " | ".join(validity_bits[:5])
    extra = "" if len(validity_bits) <= 5 else f" (+{len(validity_bits) - 5} more)"
    if mode_u == "BASE":
        return MatchResult(
            [],
            False,
            (
                f"BASE validity not identical, no contiguous lane cover, and update not "
                f"inside a single wider lane; update={update_from}..{update_to}: {joined}{extra}"
            ),
        )
    return MatchResult(
        [],
        False,
        (
            f"ETSBAF update not inside a single lane and no contiguous lane set covers "
            f"update={update_from}..{update_to}: {joined}{extra}"
        ),
    )
