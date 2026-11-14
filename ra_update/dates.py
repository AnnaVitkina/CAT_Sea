from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any

import pandas as pd


def parse_date(value: Any) -> date | None:
    """Parse RA (DD.MM.YYYY) or update (datetime / ISO) dates."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, pd.Timestamp):
        return value.date()

    text = str(value).strip()
    if not text:
        return None

    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    # Excel serial (rare if already parsed)
    try:
        return pd.to_datetime(text, dayfirst=True).date()
    except Exception:
        return None


def format_ra_date(d: date) -> str:
    return d.strftime("%d.%m.%Y")


def period_inside(inner_from: date, inner_to: date, outer_from: date, outer_to: date) -> bool:
    return inner_from >= outer_from and inner_to <= outer_to


def periods_equal(a_from: date, a_to: date, b_from: date, b_to: date) -> bool:
    return a_from == b_from and a_to == b_to


def uncovered_period_segments(
    span_from: date,
    span_to: date,
    covered: list[tuple[date, date]],
) -> list[tuple[date, date]]:
    """
    Sub-periods of [span_from, span_to] not covered by any interval in covered.

    Used when BASE force-creates a continuation lane: existing RA lane
    01.06.2026-31.10.2026 + update 01.06.2026-31.05.2027 → remainder
    01.11.2026-31.05.2027.
    """
    if span_from > span_to:
        return []

    clips: list[tuple[date, date]] = []
    for a, b in covered:
        s = max(a, span_from)
        e = min(b, span_to)
        if s <= e:
            clips.append((s, e))
    if not clips:
        return [(span_from, span_to)]

    clips.sort(key=lambda p: (p[0], p[1]))
    merged: list[tuple[date, date]] = [clips[0]]
    for s, e in clips[1:]:
        prev_s, prev_e = merged[-1]
        if s <= prev_e + timedelta(days=1):
            merged[-1] = (prev_s, max(prev_e, e))
        else:
            merged.append((s, e))

    gaps: list[tuple[date, date]] = []
    cursor = span_from
    for s, e in merged:
        if cursor < s:
            gaps.append((cursor, s - timedelta(days=1)))
        cursor = max(cursor, e + timedelta(days=1))
    if cursor <= span_to:
        gaps.append((cursor, span_to))
    return gaps


def clamp_period_to_agreement(
    start: date,
    end: date,
    agreement_from: date | None,
    agreement_to: date | None,
) -> tuple[date, date] | None:
    """
    Cut [start, end] to the RA agreement window.
    Returns None if the period lies entirely outside the agreement.
    """
    s, e = start, end
    if agreement_from is not None and s < agreement_from:
        s = agreement_from
    if agreement_to is not None and e > agreement_to:
        e = agreement_to
    if s > e:
        return None
    return s, e


def agreement_clamp_note(*, clamped_start: bool, clamped_end: bool) -> str | None:
    """Human-readable note when update validity was cut to the RA agreement window."""
    parts: list[str] = []
    if clamped_start:
        parts.append(
            "Update started before RA agreement validity; "
            "treated as starting from agreement Valid from"
        )
    if clamped_end:
        parts.append(
            "Update ended after RA agreement validity; "
            "treated as ending at agreement Valid to"
        )
    return " | ".join(parts) if parts else None


def clamp_update_to_agreement(
    update_from: date | None,
    update_to: date | None,
    agreement_from: date | None,
    agreement_to: date | None,
) -> tuple[date | None, date | None, bool, bool, str | None]:
    """
    Cut update validity to the RA agreement window (both ends).

      - starts before agreement → start = agreement_from
      - ends after agreement → end = agreement_to

    Returns (effective_from, effective_to, clamped_start, clamped_end, error).
    """
    if update_from is None or update_to is None:
        return update_from, update_to, False, False, None

    effective_from = update_from
    effective_to = update_to
    clamped_start = False
    clamped_end = False

    if agreement_from is not None and effective_from < agreement_from:
        effective_from = agreement_from
        clamped_start = True

    if agreement_to is not None and effective_to > agreement_to:
        effective_to = agreement_to
        clamped_end = True

    if effective_from > effective_to:
        return (
            effective_from,
            effective_to,
            clamped_start,
            clamped_end,
            (
                f"Update period empty after applying RA agreement validity "
                f"{format_ra_date(agreement_from) if agreement_from else '?'} - "
                f"{format_ra_date(agreement_to) if agreement_to else '?'}: "
                f"original={format_ra_date(update_from)}-{format_ra_date(update_to)}"
            ),
        )

    return effective_from, effective_to, clamped_start, clamped_end, None


def parse_validity_period_text(text: object) -> tuple[date | None, date | None]:
    """Parse '01.06.2026 - 31.08.2027' from General info."""
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return None, None
    raw = str(text).strip()
    if not raw:
        return None, None
    for sep in (" - ", " – ", " — ", "-", "–", "—"):
        if sep in raw:
            left, right = raw.split(sep, 1)
            return parse_date(left.strip()), parse_date(right.strip())
    return parse_date(raw), None
