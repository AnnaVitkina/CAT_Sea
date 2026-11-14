"""Select cost columns whose validity periods cover an update period without gaps."""

from __future__ import annotations

from datetime import date, timedelta

from .cost_map import CostColumn
from .dates import format_ra_date


def periods_cover_without_gaps(
    update_from: date,
    update_to: date,
    candidates: list[CostColumn],
) -> list[CostColumn] | None:
    """
    Return a minimal-enough ordered list of cost columns that together cover
    [update_from, update_to] with no gaps (adjacent periods may touch by 1 day).

    Returns None if coverage is impossible.
    """
    overlapping = [
        c
        for c in candidates
        if c.valid_from is not None
        and c.valid_to is not None
        and c.valid_from <= update_to
        and c.valid_to >= update_from
    ]
    if not overlapping:
        return None

    # Prefer longer columns first so a short slice (15.07-30.09) does not
    # win over an existing wider cover (15.07-14.10) with the same start.
    overlapping.sort(
        key=lambda c: (
            c.valid_from,
            -((c.valid_to - c.valid_from).days if c.valid_to and c.valid_from else 0),
            c.punit_col,
        )
    )

    selected: list[CostColumn] = []
    covered_until: date | None = None

    for c in overlapping:
        assert c.valid_from is not None and c.valid_to is not None
        if covered_until is None:
            if c.valid_from > update_from:
                return None  # gap at the start
            selected.append(c)
            covered_until = c.valid_to
        else:
            # next segment must start on/before the day after current coverage
            if c.valid_from > covered_until + timedelta(days=1):
                return None  # gap
            if c.valid_to <= covered_until:
                continue  # fully inside already covered range
            selected.append(c)
            covered_until = c.valid_to

        if covered_until >= update_to:
            return selected

    return None


def describe_cover_failure(
    update_from: date,
    update_to: date,
    fee: str,
    equipment: str,
    candidates: list[CostColumn],
) -> str:
    headers = [c.header for c in candidates]
    return (
        f"No cost column set covers validity "
        f"{format_ra_date(update_from)}-{format_ra_date(update_to)} "
        f"for {fee} ({equipment}) without gaps; candidates={headers}"
    )
