"""Split / duplicate RA lanes so an update period has an exact-validity lane."""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from datetime import date, timedelta

from openpyxl.styles import PatternFill

from .config import (
    CHANGED_FILL_HEX,
    COL_LANE,
    COL_VALID_FROM,
    COL_VALID_TO,
)
from .cost_map import CostColumn
from .dates import format_ra_date, periods_equal
from .ra_model import RaLane, RateCardWorkbook

GREEN = PatternFill(start_color=CHANGED_FILL_HEX, end_color=CHANGED_FILL_HEX, fill_type="solid")


def cost_update_leaks_outside_update(
    lane: RaLane,
    costs: list[CostColumn],
    update_from: date,
    update_to: date,
) -> bool:
    """
    True when writing these cost values on `lane` would affect days that are
    outside the update period.

    A cost applies on the intersection of (cost validity, lane validity).
    If that intersection sticks out beyond [update_from, update_to], updating
    the cost on this lane would incorrectly change non-update days -> split.
    """
    if lane.valid_from is None or lane.valid_to is None:
        return False

    for cost in costs:
        if cost.valid_from is None or cost.valid_to is None:
            # Period-less cost applies across the whole lane validity
            applied_from, applied_to = lane.valid_from, lane.valid_to
        else:
            applied_from = max(cost.valid_from, lane.valid_from)
            applied_to = min(cost.valid_to, lane.valid_to)
            if applied_from > applied_to:
                continue  # cost does not apply on this lane at all

        if applied_from < update_from or applied_to > update_to:
            return True
    return False


@dataclass
class SplitAction:
    action: str  # cut_original | duplicate_update | duplicate_after
    source_lane_no: str
    source_row: int
    result_lane_no: str
    result_row: int
    valid_from: date
    valid_to: date
    detail: str


def _next_lane_number(ra: RateCardWorkbook) -> int:
    nums: list[int] = []
    for lane in ra.lanes:
        try:
            nums.append(int(str(lane.lane_no).strip()))
        except (TypeError, ValueError):
            continue
    return (max(nums) + 1) if nums else 1


def _copy_row(ws, source_row: int, target_row: int) -> None:
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        dst.value = src.value
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)
            dst.alignment = copy(src.alignment)


def _set_validity(ra: RateCardWorkbook, row: int, valid_from: date, valid_to: date) -> None:
    ws = ra.ws
    from_cell = ws.cell(row, COL_VALID_FROM)
    to_cell = ws.cell(row, COL_VALID_TO)
    from_cell.value = format_ra_date(valid_from)
    to_cell.value = format_ra_date(valid_to)
    from_cell.fill = GREEN
    to_cell.fill = GREEN


def _set_lane_no(ra: RateCardWorkbook, row: int, lane_no: str) -> None:
    cell = ra.ws.cell(row, COL_LANE)
    cell.value = lane_no
    cell.fill = GREEN


def _refresh_lane_object(ra: RateCardWorkbook, lane: RaLane, valid_from: date, valid_to: date) -> None:
    lane.valid_from = valid_from
    lane.valid_to = valid_to


def find_exact_validity_lane(
    ra: RateCardWorkbook,
    *,
    key: str,
    template: RaLane,
    update_from: date,
    update_to: date,
) -> RaLane | None:
    """Find an existing lane for key with exact update validity and same core details."""
    for lane in ra.find_by_key(key):
        if lane.valid_from != update_from or lane.valid_to != update_to:
            continue
        if lane.carrier != template.carrier:
            continue
        if lane.service_c != template.service_c:
            continue
        if lane.origin_country != template.origin_country:
            continue
        if lane.dest_country != template.dest_country:
            continue
        return lane
    return None


def ensure_exact_validity_lane(
    ra: RateCardWorkbook,
    *,
    key: str,
    lane: RaLane,
    update_from: date,
    update_to: date,
) -> tuple[RaLane, list[SplitAction]]:
    """
    Ensure there is a lane with Valid from/to exactly equal to the update period.

    If the matched lane is wider:
      1) cut original to the pre-update segment (when it exists)
      2) duplicate a lane with the update validity
      3) duplicate a post-update remainder lane when needed (no gaps)
    """
    actions: list[SplitAction] = []

    if lane.valid_from is None or lane.valid_to is None:
        raise ValueError(f"Lane {lane.lane_no} missing validity")

    if periods_equal(lane.valid_from, lane.valid_to, update_from, update_to):
        return lane, actions

    existing = find_exact_validity_lane(
        ra, key=key, template=lane, update_from=update_from, update_to=update_to
    )
    if existing is not None:
        return existing, actions

    l_from, l_to = lane.valid_from, lane.valid_to
    if update_from < l_from or update_to > l_to:
        raise ValueError(
            f"Update {update_from}..{update_to} is not inside lane "
            f"{lane.lane_no} validity {l_from}..{l_to}"
        )

    has_before = update_from > l_from
    has_after = update_to < l_to

    # Case: update starts at lane start → keep original as update period, append after if needed
    if not has_before:
        if has_after:
            # Duplicate the after-segment first while the full original values are intact
            after_from = update_to + timedelta(days=1)
            _duplicate_lane(
                ra,
                source=lane,
                valid_from=after_from,
                valid_to=l_to,
                action_label="duplicate_after",
                actions=actions,
            )
        _set_validity(ra, lane.row, update_from, update_to)
        _refresh_lane_object(ra, lane, update_from, update_to)
        actions.append(
            SplitAction(
                action="cut_original",
                source_lane_no=lane.lane_no,
                source_row=lane.row,
                result_lane_no=lane.lane_no,
                result_row=lane.row,
                valid_from=update_from,
                valid_to=update_to,
                detail="Original lane validity set to update period",
            )
        )
        return lane, actions

    # Case: has before → cut original to before, duplicate update (+ after)
    before_to = update_from - timedelta(days=1)
    original_lane_no = lane.lane_no
    original_row = lane.row

    _set_validity(ra, lane.row, l_from, before_to)
    _refresh_lane_object(ra, lane, l_from, before_to)
    actions.append(
        SplitAction(
            action="cut_original",
            source_lane_no=original_lane_no,
            source_row=original_row,
            result_lane_no=lane.lane_no,
            result_row=lane.row,
            valid_from=l_from,
            valid_to=before_to,
            detail=(
                f"Cut lane {original_lane_no} validity to "
                f"{format_ra_date(l_from)}-{format_ra_date(before_to)}"
            ),
        )
    )

    # Duplicate from the (still fully valued) cut row — costs/other fields unchanged
    update_lane = _duplicate_lane(
        ra,
        source=lane,
        valid_from=update_from,
        valid_to=update_to,
        action_label="duplicate_update",
        actions=actions,
    )

    if has_after:
        after_from = update_to + timedelta(days=1)
        _duplicate_lane(
            ra,
            source=lane,
            valid_from=after_from,
            valid_to=l_to,
            action_label="duplicate_after",
            actions=actions,
        )

    return update_lane, actions


def _duplicate_lane(
    ra: RateCardWorkbook,
    *,
    source: RaLane,
    valid_from: date,
    valid_to: date,
    action_label: str,
    actions: list[SplitAction],
) -> RaLane:
    new_row = ra.ws.max_row + 1
    _copy_row(ra.ws, source.row, new_row)
    new_lane_no = str(_next_lane_number(ra))
    _set_lane_no(ra, new_row, new_lane_no)
    _set_validity(ra, new_row, valid_from, valid_to)

    new_lane = ra.register_lane_from_row(new_row)
    actions.append(
        SplitAction(
            action=action_label,
            source_lane_no=source.lane_no,
            source_row=source.row,
            result_lane_no=new_lane.lane_no,
            result_row=new_row,
            valid_from=valid_from,
            valid_to=valid_to,
            detail=(
                f"Duplicated lane {source.lane_no} -> {new_lane.lane_no} with validity "
                f"{format_ra_date(valid_from)}-{format_ra_date(valid_to)}"
            ),
        )
    )
    return new_lane
