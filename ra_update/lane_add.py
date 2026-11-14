"""Add missing BASE lanes for applicable Sea / Sea FR RAs."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta

import pandas as pd
from openpyxl.styles import PatternFill

from .config import (
    CHANGED_FILL_HEX,
    COL_CARRIER,
    COL_DEST_COUNTRY,
    COL_DEST_PORT,
    COL_DEST_POSTAL,
    COL_KEY,
    COL_LANE,
    COL_ORIGIN_COUNTRY,
    COL_ORIGIN_PORT,
    COL_ORIGIN_POSTAL,
    COL_SERVICE,
    COL_SERVICE_C,
    COL_SERVICE_SHORT,
    COL_TRANSPOREON,
    COL_VALID_FROM,
    COL_VALID_TO,
    SEA_FR_DEST_COUNTRIES,
    SEA_LANE_COUNTRIES,
)
from .dates import format_ra_date, period_inside, periods_equal
from .lane_identity import (
    DEFAULT_SERVICE_TEXT,
    build_transporeon_id,
    derive_service_short,
    location_port_postal,
    normalize_slash_value,
)
from .lane_split import _copy_row, _next_lane_number, _set_lane_no, _set_validity
from .ra_model import RaLane, RateCardWorkbook

GREEN = PatternFill(start_color=CHANGED_FILL_HEX, end_color=CHANGED_FILL_HEX, fill_type="solid")


@dataclass
class AddAction:
    action: str  # create_lane | cut_added_lane | duplicate_after_add
    key: str
    lane_no: str
    ra_row: int
    valid_from: date
    valid_to: date
    detail: str


def ra_kind(agreement_title: str | None) -> str:
    """Classify RA for missing-lane auto-add rules."""
    title = (agreement_title or "").strip().lower()
    if not title:
        return "unknown"
    if title == "sea fr" or title.startswith("sea fr"):
        return "sea_fr"
    if title == "sea" or title.startswith("sea"):
        return "sea"
    return "other"


def is_missing_lane_addable(
    agreement_title: str | None,
    origin_country: str,
    dest_country: str,
) -> tuple[bool, str | None]:
    """
    Whether a KEY that is absent from the Rate Card may be auto-added.

    Sea: lanes from/to GB or ZA (incl. GB-ZA / ZA-GB).
    Sea FR: lanes to destination country FR.
    """
    kind = ra_kind(agreement_title)
    o = (origin_country or "").strip().upper()
    d = (dest_country or "").strip().upper()

    if kind == "sea":
        if o in SEA_LANE_COUNTRIES or d in SEA_LANE_COUNTRIES:
            return True, None
        return (
            False,
            (
                f"No such lane and not addable to Sea RA "
                f"(need from/to GB or ZA; got {o or '?'}-{d or '?'})"
            ),
        )

    if kind == "sea_fr":
        if d in SEA_FR_DEST_COUNTRIES:
            return True, None
        return (
            False,
            (
                f"No such lane and not addable to Sea FR RA "
                f"(need destination FR; got {o or '?'}-{d or '?'})"
            ),
        )

    title = (agreement_title or "").strip() or "?"
    return (
        False,
        f"No such lane; auto-add not configured for RA {title!r}",
    )


def _set_green(ws, row: int, col: int, value) -> None:
    cell = ws.cell(row, col)
    cell.value = value
    cell.fill = GREEN


def _clear_cost_cells(ra: RateCardWorkbook, row: int) -> None:
    for cost in ra.cost_columns:
        for col in (cost.currency_col, cost.flat_col, cost.punit_col):
            if col is not None:
                ra.ws.cell(row, col).value = None


def create_lane_from_update(
    ra: RateCardWorkbook,
    *,
    key: str,
    carrier: str,
    service_c: str,
    origin_country: str,
    dest_country: str,
    origin_code: str,
    dest_code: str,
    valid_from: date,
    valid_to: date,
) -> tuple[RaLane, AddAction]:
    """Append a new Rate Card lane row populated from update identity fields."""
    ws = ra.ws
    donor = ra.lanes[-1] if ra.lanes else None
    new_row = ws.max_row + 1

    if donor is not None:
        _copy_row(ws, donor.row, new_row)
        _clear_cost_cells(ra, new_row)
    else:
        for col in range(1, ws.max_column + 1):
            ws.cell(new_row, col).value = None

    key_n = normalize_slash_value(key)
    service_c_n = normalize_slash_value(service_c)
    service_short = derive_service_short(service_c_n)
    origin_port, origin_postal = location_port_postal(origin_code)
    dest_port, dest_postal = location_port_postal(dest_code)
    transporeon = build_transporeon_id(
        carrier=carrier,
        service_short=service_short,
        origin_port=origin_port,
        origin_postal=origin_postal,
        dest_port=dest_port,
        dest_postal=dest_postal,
    )
    lane_no = str(_next_lane_number(ra))

    _set_green(ws, new_row, COL_LANE, lane_no)
    _set_green(ws, new_row, COL_TRANSPOREON, transporeon)
    _set_green(ws, new_row, COL_KEY, key_n)
    _set_green(ws, new_row, COL_CARRIER, carrier)
    _set_green(ws, new_row, COL_SERVICE, DEFAULT_SERVICE_TEXT)
    _set_green(ws, new_row, COL_SERVICE_C, service_c_n)
    _set_green(ws, new_row, COL_SERVICE_SHORT, service_short)
    _set_validity(ra, new_row, valid_from, valid_to)
    _set_green(ws, new_row, COL_ORIGIN_PORT, origin_port)
    _set_green(ws, new_row, COL_ORIGIN_POSTAL, origin_postal)
    _set_green(ws, new_row, COL_ORIGIN_COUNTRY, origin_country)
    _set_green(ws, new_row, COL_DEST_PORT, dest_port)
    _set_green(ws, new_row, COL_DEST_POSTAL, dest_postal)
    _set_green(ws, new_row, COL_DEST_COUNTRY, dest_country)

    lane = ra.register_lane_from_row(new_row)
    action = AddAction(
        action="create_lane",
        key=key,
        lane_no=lane.lane_no,
        ra_row=new_row,
        valid_from=valid_from,
        valid_to=valid_to,
        detail=(
            f"Added missing lane {lane.lane_no} "
            f"{format_ra_date(valid_from)}-{format_ra_date(valid_to)}"
        ),
    )
    return lane, action


def choose_widest_period(
    periods: list[tuple[date, date]],
) -> tuple[date, date]:
    """Prefer longest duration; ties → earlier start, later end."""
    return max(
        periods,
        key=lambda p: ((p[1] - p[0]).days, -p[0].toordinal(), p[1].toordinal()),
    )


def resolve_nested_validity_cuts(
    ra: RateCardWorkbook,
    *,
    key: str,
    lanes_by_period: dict[tuple[date, date], RaLane],
) -> list[AddAction]:
    """
    When one added period is strictly inside another, cut the wider lane so
    validity segments do not overlap (earlier/older date cutting).

    Examples:
      wide  01.06.2026-31.05.2027, narrow 15.09.2026-31.05.2027
        -> wide becomes 01.06.2026-14.09.2026
           (and an after-segment if wide extended past narrow end)

      wide  01.06.2026-31.05.2027, narrow 01.06.2026-31.10.2026  (same start)
        -> wide shrinks to 01.11.2026-31.05.2027
           narrow keeps 01.06.2026-31.10.2026
           (exactly 2 lanes; no leftover full-wide row)
    """
    actions: list[AddAction] = []
    # Only original update periods participate as cut sources (not after-segments we add)
    periods = list(lanes_by_period.keys())
    # Process wider outers first so nested chains cut cleanly
    periods_by_width = sorted(
        periods,
        key=lambda p: ((p[1] - p[0]).days, -p[0].toordinal()),
        reverse=True,
    )
    cut_done: set[tuple[tuple[date, date], tuple[date, date]]] = set()

    for outer in periods_by_width:
        for inner in periods:
            pair = (outer, inner)
            if pair in cut_done:
                continue
            if periods_equal(outer[0], outer[1], inner[0], inner[1]):
                continue
            if not period_inside(inner[0], inner[1], outer[0], outer[1]):
                continue

            lane = lanes_by_period.get(outer)
            if lane is None or lane.valid_from is None or lane.valid_to is None:
                continue

            o_from, o_to = outer
            i_from, i_to = inner
            has_before = o_from < i_from
            has_after = o_to > i_to

            if not has_before and not has_after:
                continue

            # Lane may already have been shortened by a previous inner cut
            if lane.valid_to is not None and has_before and lane.valid_to < i_from:
                cut_done.add(pair)
                continue

            cut_done.add(pair)

            # Same start, outer longer (e.g. 01.06-31.05.2027 vs 01.06-31.10.2026):
            # inner lane already covers Jun–Oct; shrink outer to the after segment only
            # → expected 2 lanes: 01.06-31.10 and 01.11-31.05 (no leftover full-wide lane).
            if has_after and not has_before:
                after_from = i_to + timedelta(days=1)
                after_to = o_to
                _set_validity(ra, lane.row, after_from, after_to)
                lane.valid_from = after_from
                lane.valid_to = after_to
                actions.append(
                    AddAction(
                        action="cut_added_lane",
                        key=key,
                        lane_no=lane.lane_no,
                        ra_row=lane.row,
                        valid_from=after_from,
                        valid_to=after_to,
                        detail=(
                            f"Cut added lane {lane.lane_no} to "
                            f"{format_ra_date(after_from)}-{format_ra_date(after_to)} "
                            f"(after nested same-start "
                            f"{format_ra_date(i_from)}-{format_ra_date(i_to)}; "
                            f"was {format_ra_date(o_from)}-{format_ra_date(o_to)})"
                        ),
                    )
                )
                continue

            if has_after:
                after_from = i_to + timedelta(days=1)
                after_to = o_to
                # Duplicate after while costs still match the pre-cut outer lane
                new_row = ra.ws.max_row + 1
                _copy_row(ra.ws, lane.row, new_row)
                new_lane_no = str(_next_lane_number(ra))
                _set_lane_no(ra, new_row, new_lane_no)
                _set_validity(ra, new_row, after_from, after_to)
                after_lane = ra.register_lane_from_row(new_row)
                lanes_by_period[(after_from, after_to)] = after_lane
                actions.append(
                    AddAction(
                        action="duplicate_after_add",
                        key=key,
                        lane_no=after_lane.lane_no,
                        ra_row=after_lane.row,
                        valid_from=after_from,
                        valid_to=after_to,
                        detail=(
                            f"After-segment lane {after_lane.lane_no} "
                            f"{format_ra_date(after_from)}-{format_ra_date(after_to)} "
                            f"from wider {format_ra_date(o_from)}-{format_ra_date(o_to)}"
                        ),
                    )
                )

            if has_before:
                new_to = i_from - timedelta(days=1)
                _set_validity(ra, lane.row, o_from, new_to)
                lane.valid_from = o_from
                lane.valid_to = new_to
                actions.append(
                    AddAction(
                        action="cut_added_lane",
                        key=key,
                        lane_no=lane.lane_no,
                        ra_row=lane.row,
                        valid_from=o_from,
                        valid_to=new_to,
                        detail=(
                            f"Cut added lane {lane.lane_no} to "
                            f"{format_ra_date(o_from)}-{format_ra_date(new_to)} "
                            f"(day before nested "
                            f"{format_ra_date(i_from)}-{format_ra_date(i_to)})"
                        ),
                    )
                )

    return actions


def periods_partially_overlap_without_nesting(
    periods: list[tuple[date, date]],
) -> tuple[date, date] | None:
    """Return a conflicting pair description when overlap is not containment."""
    for i, a in enumerate(periods):
        for b in periods[i + 1 :]:
            if periods_equal(a[0], a[1], b[0], b[1]):
                continue
            # overlap?
            if a[0] <= b[1] and b[0] <= a[1]:
                a_in_b = period_inside(a[0], a[1], b[0], b[1])
                b_in_a = period_inside(b[0], b[1], a[0], a[1])
                if not a_in_b and not b_in_a:
                    return a
    return None


def currency_for_rate(row: pd.Series, rate_col: str) -> str | None:
    """Companion *_CURRENCY for a *_RATE column, when present."""
    if rate_col.upper().endswith("_RATE"):
        cur_col = rate_col[: -len("_RATE")] + "_CURRENCY"
    else:
        cur_col = rate_col + "_CURRENCY"
    # Exact match first, then case-insensitive fallback
    if cur_col in row.index:
        val = row.get(cur_col)
    else:
        upper_map = {str(c).upper(): c for c in row.index}
        hit = upper_map.get(cur_col.upper())
        if hit is None:
            return None
        val = row.get(hit)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    text = str(val).strip()
    return text or None
