"""BAF monthly (15th→14th) and EU ETS period cost helpers + column creation."""

from __future__ import annotations

from calendar import monthrange
from copy import copy
from datetime import date, timedelta

import pandas as pd
from openpyxl.styles import PatternFill

from .charge_rules import DFT_WEIGHT_RATE_BY, resolve_rate_column
from .config import CHANGED_FILL_HEX, DATA_START_ROW, HEADER_COST_ROW, HEADER_FIELD_ROW
from .cost_map import (
    CostColumn,
    match_period_fee_columns,
    parse_cost_header,
    period_fee_header,
)
from .dates import (
    clamp_period_to_agreement,
    format_ra_date,
    parse_date,
    period_inside,
    uncovered_period_segments,
)
from .parsing import to_number

GREEN = PatternFill(start_color=CHANGED_FILL_HEX, end_color=CHANGED_FILL_HEX, fill_type="solid")


def build_contiguous_periods_from_boundaries(
    effectives: set[date] | list[date],
    expiries: set[date] | list[date],
) -> list[tuple[date, date]]:
    """
    Build non-overlapping, gap-free validity periods for DFT BAF / DFT ETS.

    Starts = unique RATE_EFFECTIVE dates (sorted).
    For each start S_i: end = day before S_{i+1}, else max(RATE_EXPIRATION).

    Example BAF (effectives 15.07 + 15.10, expiries 14.10 + 31.10):
      15.07.2026-14.10.2026, 15.10.2026-31.10.2026
    Example ETS (effectives 15.07 + 01.10 + 15.10, expiries 14.10 + 31.10):
      15.07.2026-30.09.2026, 01.10.2026-14.10.2026, 15.10.2026-31.10.2026
    """
    starts = sorted({d for d in effectives if d is not None})
    ends = {d for d in expiries if d is not None}
    if not starts or not ends:
        return []
    last_end = max(ends)
    out: list[tuple[date, date]] = []
    for i, start in enumerate(starts):
        if i + 1 < len(starts):
            end = starts[i + 1] - timedelta(days=1)
        else:
            end = last_end
        if end < start:
            continue
        out.append((start, end))
    return out


def collect_dft_boundary_dates(
    df: pd.DataFrame,
    *,
    charge_keys: set[str],
    ra=None,
    sheet: str = "ETSBAF",
) -> tuple[set[date], set[date]]:
    """
    Collect RATE_EFFECTIVE / RATE_EXPIRATION dates from update rows that have
    a non-empty rate for the given DFT charge keys (e.g. DFTBAF, DFTEU_ETS).

    When `ra` is provided, only rows that remain after out-of-scope filtering
    and that target an RA lane (existing KEY / sibling) contribute boundaries.
    Otherwise out-of-RA keys inject spurious effectives (e.g. 28.07, 07.08)
    and slice DFT catalog periods that no in-scope lane needs.
    """
    from .equipment_decisions import update_row_targets_ra

    effectives: set[date] = set()
    expiries: set[date] = set()
    keys_u = {k.upper() for k in charge_keys}
    for _, row in df.iterrows():
        if ra is not None and not update_row_targets_ra(ra, row, sheet):
            continue
        vf = parse_date(row.get("RATE_EFFECTIVE_DATE__C"))
        vt = parse_date(row.get("RATE_EXPIRATION_DATE__C"))
        if vf is None or vt is None:
            continue
        hit = False
        for col, raw in row.items():
            col_s = str(col)
            if not col_s.endswith("_RATE"):
                continue
            if to_number(raw) is None:
                continue
            rule, _eq, skip = resolve_rate_column(col_s, sheet)
            if skip or rule is None:
                continue
            if (rule.charge_key or "").upper() in keys_u:
                hit = True
                break
        if hit:
            effectives.add(vf)
            expiries.add(vt)
    return effectives, expiries


def dft_periods_from_update_df(
    df: pd.DataFrame,
    *,
    charge_keys: set[str],
    agreement_from: date | None = None,
    agreement_to: date | None = None,
    ra=None,
    sheet: str = "ETSBAF",
) -> list[tuple[date, date]]:
    """
    Derive DFT cost validity periods from in-scope update rows that target
    the RA, then cut each period to the RA agreement window (do not create
    headers past agreement Valid to / before Valid from).
    """
    effectives, expiries = collect_dft_boundary_dates(
        df, charge_keys=charge_keys, ra=ra, sheet=sheet
    )
    periods = build_contiguous_periods_from_boundaries(effectives, expiries)
    if agreement_from is None and agreement_to is None:
        return periods
    out: list[tuple[date, date]] = []
    for a, b in periods:
        clamped = clamp_period_to_agreement(a, b, agreement_from, agreement_to)
        if clamped is not None:
            out.append(clamped)
    return out


def _add_months(d: date, months: int) -> date:
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    day = min(d.day, monthrange(y, m)[1])
    return date(y, m, day)


def baf_month_end(start_15: date) -> date:
    """14th of the month after a 15th start."""
    nxt = _add_months(date(start_15.year, start_15.month, 1), 1)
    return date(nxt.year, nxt.month, 14)


def baf_month_containing(d: date) -> tuple[date, date]:
    """Standard BAF month that contains day d: 15th → 14th next month."""
    if d.day >= 15:
        start = date(d.year, d.month, 15)
    else:
        prev = _add_months(date(d.year, d.month, 1), -1)
        start = date(prev.year, prev.month, 15)
    return start, baf_month_end(start)


def is_standard_baf_month(valid_from: date, valid_to: date) -> bool:
    if valid_from.day != 15:
        return False
    return valid_to == baf_month_end(valid_from)


def iter_baf_months_covering(update_from: date, update_to: date) -> list[tuple[date, date]]:
    """All standard BAF months that overlap [update_from, update_to]."""
    start, end = baf_month_containing(update_from)
    out: list[tuple[date, date]] = []
    while start <= update_to:
        out.append((start, end))
        start = end + timedelta(days=1)
        end = baf_month_end(start)
    return out


def baf_period_validation_note(update_from: date, update_to: date) -> str | None:
    """
    None if update is exactly one standard month or an exact multi-month
    span from a 15th through a 14th on month boundaries.
    """
    months = iter_baf_months_covering(update_from, update_to)
    if not months:
        return f"BAF period empty for {format_ra_date(update_from)}-{format_ra_date(update_to)}"
    if len(months) == 1 and months[0] == (update_from, update_to):
        return None
    if (
        update_from == months[0][0]
        and update_to == months[-1][1]
        and all(is_standard_baf_month(a, b) for a, b in months)
    ):
        return None
    expanded = ", ".join(f"{format_ra_date(a)}-{format_ra_date(b)}" for a, b in months)
    return (
        f"BAF period {format_ra_date(update_from)}-{format_ra_date(update_to)} "
        f"is not aligned to 15th→14th monthly validity; expanded to: {expanded}"
    )


def _block_width(ra, cost: CostColumn) -> int:
    starts = sorted(c.start_col for c in ra.cost_columns)
    try:
        i = starts.index(cost.start_col)
    except ValueError:
        return 2
    if i + 1 < len(starts):
        return max(1, starts[i + 1] - cost.start_col)
    return 2


def _rate_by_from_template(ra, template: CostColumn) -> str:
    raw = ra.ws.cell(5, template.start_col).value
    if not raw:
        return "Container"
    for line in str(raw).replace("\r", "").split("\n"):
        if line.strip().lower().startswith("rate by:"):
            return line.split(":", 1)[1].strip()
    return "Container"


def find_template_period_cost(
    ra,
    *,
    fee_name: str,
    equipment: str,
) -> CostColumn | None:
    """Prefer same equipment period column; else a compatible fee template."""
    eq_u = (equipment or "").upper()
    if eq_u == "DFT" or not eq_u:
        from .cost_map import weight_period_candidates

        same = weight_period_candidates(ra.cost_columns, fee_name=fee_name)
    else:
        same = [
            c
            for c in ra.cost_columns
            if c.fee_name == fee_name
            and c.equipment == equipment
            and c.valid_from is not None
        ]
    if same:
        # Prefer a standard BAF month template when available
        for c in same:
            if c.valid_from and c.valid_to and is_standard_baf_month(c.valid_from, c.valid_to):
                return c
        return same[0]

    # Fallback: keep DFT/weight templates away from container equipment
    if eq_u == "DFT" or not eq_u:
        any_fee = [
            c
            for c in ra.cost_columns
            if c.fee_name == fee_name and c.valid_from is not None
        ]
    else:
        any_fee = [
            c
            for c in ra.cost_columns
            if c.fee_name == fee_name
            and c.valid_from is not None
            and c.equipment
            and str(c.equipment).upper() != "DFT"
        ]
        if not any_fee:
            any_fee = [
                c
                for c in ra.cost_columns
                if c.fee_name == fee_name and c.valid_from is not None
            ]
    return any_fee[0] if any_fee else None


def _is_weight_rate_by(text: str | None) -> bool:
    return bool(text) and "weight" in str(text).lower()


def _container_rate_by(equipment: str) -> str:
    eq = (equipment or "").strip()
    return f"Container/{eq}" if eq else "Container"


def _rate_by_for_period_cost(
    ra,
    *,
    equipment: str,
    template: CostColumn,
    rate_by: str | None,
) -> str:
    """
    Rate by:
      - DFT / no-equipment → Weight/chargeable kg
      - container equipment (42G0, 42SOC, …) → Container/… never Weight
    """
    eq_u = (equipment or "").upper()
    weight_fee = eq_u == "DFT" or not eq_u

    if weight_fee:
        if rate_by:
            return rate_by
        return DFT_WEIGHT_RATE_BY

    # Container equipment: ignore Weight templates / overrides
    if rate_by and not _is_weight_rate_by(rate_by):
        return rate_by
    tmpl = _rate_by_from_template(ra, template)
    if _is_weight_rate_by(tmpl):
        return _container_rate_by(equipment)
    return tmpl


def create_period_cost_column(
    ra,
    *,
    fee_name: str,
    equipment: str,
    valid_from: date,
    valid_to: date,
    template: CostColumn,
    prolong_from: CostColumn | None = None,
    rate_by: str | None = None,
) -> CostColumn:
    """
    Append a period cost block, e.g.
      BAF Fee (22G0 15.09.2026-14.10.2026)
      BAF Fee (DFT 15.07.2026-14.08.2026)  — weight/*DFTBAF
    with Validity period / Cost to prolong / Rate by metadata.

    Validity is always cut to the RA agreement window before the header is built.
    """
    clamped = clamp_period_to_agreement(
        valid_from,
        valid_to,
        getattr(ra, "agreement_valid_from", None),
        getattr(ra, "agreement_valid_to", None),
    )
    if clamped is None:
        raise ValueError(
            f"Cannot create {fee_name} ({equipment}): period "
            f"{format_ra_date(valid_from)}-{format_ra_date(valid_to)} "
            f"is outside RA agreement validity"
        )
    valid_from, valid_to = clamped

    ws = ra.ws
    width = _block_width(ra, template)
    start = ws.max_column + 1
    header = period_fee_header(fee_name, equipment, valid_from, valid_to)

    # Avoid duplicate create (header text may differ on spaces around date dash)
    existing = match_period_fee_columns(
        ra.cost_columns,
        fee_name=fee_name,
        equipment=equipment,
        valid_from=valid_from,
        valid_to=valid_to,
    )
    if existing:
        return existing[0]

    resolved_rate_by = _rate_by_for_period_cost(
        ra, equipment=equipment, template=template, rate_by=rate_by
    )
    applies_lines = [
        f"Validity period: from {format_ra_date(valid_from)} to {format_ra_date(valid_to)}",
        "",
    ]
    if prolong_from is not None:
        applies_lines.append(f"Cost to prolong: {prolong_from.header}")
        applies_lines.append("")
        applies_lines.append("")
    applies_lines.append("Applies if invoiced by Carrier")
    applies = "\n".join(applies_lines)
    rate_by_text = f"Rate by: {resolved_rate_by}\nRegular rule"

    for offset in range(width):
        src_col = template.start_col + offset
        dst_col = start + offset
        for row in range(1, 8):
            src = ws.cell(row, src_col)
            dst = ws.cell(row, dst_col)
            if src.has_style:
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.fill = copy(src.fill)
                dst.number_format = src.number_format
                dst.protection = copy(src.protection)
                dst.alignment = copy(src.alignment)
        for row in range(DATA_START_ROW, ws.max_row + 1):
            ws.cell(row, dst_col).value = None

    ws.cell(HEADER_COST_ROW, start).value = header
    ws.cell(HEADER_COST_ROW, start).fill = GREEN
    ws.cell(4, start).value = applies
    ws.cell(4, start).fill = GREEN
    ws.cell(5, start).value = rate_by_text
    ws.cell(5, start).fill = GREEN

    for offset in range(width):
        src_col = template.start_col + offset
        dst_col = start + offset
        ws.cell(HEADER_FIELD_ROW, dst_col).value = ws.cell(HEADER_FIELD_ROW, src_col).value
        ws.cell(6, dst_col).value = ws.cell(6, src_col).value

    fee, eq, v_from, v_to = parse_cost_header(header)
    currency_col = flat_col = punit_col = None
    flat_is_min = False
    for offset in range(width):
        c = start + offset
        lab = ws.cell(HEADER_FIELD_ROW, c).value
        if lab is None:
            continue
        low = str(lab).strip().lower()
        if low == "currency" and currency_col is None:
            currency_col = c
        elif low == "flat" and flat_col is None:
            flat_col = c
            marker = ws.cell(6, c).value
            flat_is_min = marker is not None and str(marker).strip().upper() == "MIN"
        elif low == "p/unit" and punit_col is None:
            punit_col = c

    new_cost = CostColumn(
        start_col=start,
        header=header,
        fee_name=fee,
        equipment=eq,
        valid_from=v_from,
        valid_to=v_to,
        currency_col=currency_col,
        flat_col=flat_col,
        flat_is_min=flat_is_min,
        punit_col=punit_col,
    )
    ra.cost_columns.append(new_cost)
    return new_cost


def _dft_fee_columns(ra, fee_name: str) -> list[CostColumn]:
    """
    DFT weight costs: BAF/ETS Fee (DFT …) plus period-only headers some Sea FR
    cards use (no equipment token). Same set as find_cost_columns for DFT.
    """
    from .cost_map import weight_period_candidates

    return weight_period_candidates(ra.cost_columns, fee_name=fee_name)


def _periods_overlap(a0: date, a1: date, b0: date, b1: date) -> bool:
    return a0 <= b1 and a1 >= b0


def resolve_dft_period_columns(
    ra,
    *,
    fee_name: str,
    update_from: date,
    update_to: date,
    dft_periods: list[tuple[date, date]],
    template_fee_fallback: str | None = None,
) -> tuple[list[CostColumn], str | None]:
    """
    DFT BAF / EU ETS cost columns:

      1) Prefer an existing DFT column that fully contains the update.
      2) Else use any existing DFT columns that together cover the update
         without gaps (their own validities — not only pivot-exact headers).
      3) Else use existing exact catalog periods matching derived dft_periods.
      4) If DFT columns overlap the update but still cannot cover without gaps
         → issue; do NOT create sliced columns.
      5) When no DFT column overlaps this update, create only uncovered
         remainders (catalog ∩ update minus existing DFT validity) — never
         recreate earlier catalog slices that overlap other existing columns.
    """
    from .period_cover import periods_cover_without_gaps

    same = _dft_fee_columns(ra, fee_name)

    containers = [
        c
        for c in same
        if c.valid_from is not None
        and c.valid_to is not None
        and period_inside(update_from, update_to, c.valid_from, c.valid_to)
    ]
    if containers:
        best = max(
            containers,
            key=lambda c: (c.valid_to - c.valid_from).days
            if c.valid_from and c.valid_to
            else 0,
        )
        return [best], None

    # Existing RA columns may already cover (e.g. 15.07-14.10 + 15.10-31.10)
    existing_cover = periods_cover_without_gaps(update_from, update_to, same)
    if existing_cover:
        return existing_cover, None

    exact: list[CostColumn] = []
    for m_from, m_to in dft_periods:
        existing = None
        for want_eq in ("DFT", None):
            hits = match_period_fee_columns(
                ra.cost_columns,
                fee_name=fee_name,
                equipment=want_eq,
                valid_from=m_from,
                valid_to=m_to,
            )
            if hits:
                existing = hits[0]
                break
        if existing is not None:
            exact.append(existing)

    covering = periods_cover_without_gaps(update_from, update_to, exact)
    if covering:
        return covering, None

    overlapping = [
        c
        for c in same
        if c.valid_from is not None
        and c.valid_to is not None
        and _periods_overlap(update_from, update_to, c.valid_from, c.valid_to)
    ]
    if overlapping:
        cand_txt = ", ".join(
            f"{c.header} validity="
            f"{format_ra_date(c.valid_from)}-{format_ra_date(c.valid_to)}"
            for c in overlapping
        )
        return [], (
            f"DFT {fee_name}: potential cost column(s) present but validity does not "
            f"match update {format_ra_date(update_from)}-{format_ra_date(update_to)}; "
            f"not creating new columns. candidates=[{cand_txt}]"
        )

    # No overlap with this update → create only uncovered catalog remainders
    template = find_template_period_cost(ra, fee_name=fee_name, equipment="DFT")
    if template is None and template_fee_fallback:
        template = find_template_period_cost(
            ra, fee_name=template_fee_fallback, equipment="DFT"
        )
    if template is None:
        return [], f"No {fee_name} template column available to clone for DFT"

    existing_periods = [
        (c.valid_from, c.valid_to)
        for c in same
        if c.valid_from is not None and c.valid_to is not None
    ]
    # Also treat already-matched exact headers as covered
    for c in exact:
        if c.valid_from is not None and c.valid_to is not None:
            existing_periods.append((c.valid_from, c.valid_to))

    to_create: list[tuple[date, date]] = []
    seen: set[tuple[date, date]] = set()
    for m_from, m_to in dft_periods:
        s = max(m_from, update_from)
        e = min(m_to, update_to)
        if s > e:
            continue
        for rf, rt in uncovered_period_segments(s, e, existing_periods):
            if (rf, rt) in seen:
                continue
            seen.add((rf, rt))
            to_create.append((rf, rt))

    results: list[CostColumn] = list(exact)
    for m_from, m_to in to_create:
        existing = None
        for want_eq in ("DFT", None):
            hits = match_period_fee_columns(
                ra.cost_columns,
                fee_name=fee_name,
                equipment=want_eq,
                valid_from=m_from,
                valid_to=m_to,
            )
            if hits:
                existing = hits[0]
                break
        if existing is not None:
            if existing not in results:
                results.append(existing)
            continue
        if any(
            c.valid_from == m_from and c.valid_to == m_to for c in results
        ):
            continue

        prev = None
        for a, b in reversed(dft_periods):
            if b < m_from:
                prev = (a, b)
                break
        prolong = None
        if prev is not None:
            for want_eq in ("DFT", None):
                hits = match_period_fee_columns(
                    ra.cost_columns,
                    fee_name=fee_name,
                    equipment=want_eq,
                    valid_from=prev[0],
                    valid_to=prev[1],
                )
                if hits:
                    prolong = hits[0]
                    break
        created = create_period_cost_column(
            ra,
            fee_name=fee_name,
            equipment="DFT",
            valid_from=m_from,
            valid_to=m_to,
            template=template,
            prolong_from=prolong,
            rate_by=DFT_WEIGHT_RATE_BY,
        )
        print(f"  Created {created.header}")
        results.append(created)
        existing_periods.append((m_from, m_to))

    covering = periods_cover_without_gaps(update_from, update_to, results + same)
    if covering:
        return covering, None
    return [], (
        f"DFT {fee_name} periods could not cover "
        f"{format_ra_date(update_from)}-{format_ra_date(update_to)}"
    )


def ensure_baf_cost_columns(
    ra,
    *,
    equipment: str,
    update_from: date,
    update_to: date,
    baf_create_store=None,
    dft_periods: list[tuple[date, date]] | None = None,
) -> tuple[list[CostColumn], str | None]:
    """
    Ensure BAF Fee columns exist for periods covering the update.

    Container equipment: standard monthly 15→14 months.
    DFT: prefer existing covering / exact periods; if a potential DFT column
    overlaps but validity does not align → issue (do not create a sliced column).
    Create only truly missing non-overlapping DFT periods.
    """
    eq_u = equipment.upper()
    if eq_u == "DFT" and dft_periods is not None:
        return resolve_dft_period_columns(
            ra,
            fee_name="BAF Fee",
            update_from=update_from,
            update_to=update_to,
            dft_periods=dft_periods,
        )

    note = baf_period_validation_note(update_from, update_to)
    months = iter_baf_months_covering(update_from, update_to)
    # Cut monthly BAF periods to RA agreement validity
    agr_f = getattr(ra, "agreement_valid_from", None)
    agr_t = getattr(ra, "agreement_valid_to", None)
    clipped_months: list[tuple[date, date]] = []
    for m_from, m_to in months:
        clipped = clamp_period_to_agreement(m_from, m_to, agr_f, agr_t)
        if clipped is not None:
            clipped_months.append(clipped)
    months = clipped_months

    template = find_template_period_cost(ra, fee_name="BAF Fee", equipment=equipment)
    if template is None:
        return [], f"No BAF Fee template column available to clone for {equipment}"

    results: list[CostColumn] = []
    skipped: list[str] = []
    for m_from, m_to in months:
        existing_cols = match_period_fee_columns(
            ra.cost_columns,
            fee_name="BAF Fee",
            equipment=eq_u,
            valid_from=m_from,
            valid_to=m_to,
        )
        if existing_cols:
            results.append(existing_cols[0])
            continue

        allow = True
        if baf_create_store is not None:
            allow = baf_create_store.allows_create(equipment, m_to)
        if not allow:
            skipped.append(f"{format_ra_date(m_from)}-{format_ra_date(m_to)}")
            continue

        prev_start = _add_months(date(m_from.year, m_from.month, 1), -1)
        prev_from = date(prev_start.year, prev_start.month, 15)
        prev_to = baf_month_end(prev_from)
        prolong_cols = match_period_fee_columns(
            ra.cost_columns,
            fee_name="BAF Fee",
            equipment=eq_u,
            valid_from=prev_from,
            valid_to=prev_to,
        )
        prolong = prolong_cols[0] if prolong_cols else None
        existing = create_period_cost_column(
            ra,
            fee_name="BAF Fee",
            equipment=equipment,
            valid_from=m_from,
            valid_to=m_to,
            template=template,
            prolong_from=prolong,
            rate_by=None,
        )
        print(f"  Created {existing.header}")
        results.append(existing)

    note_parts = [p for p in (note,) if p]
    if skipped:
        note_parts.append(
            f"BAF create skipped for {equipment} (past cutoff): {', '.join(skipped)}"
        )
    combined = " | ".join(note_parts) if note_parts else None
    return results, combined


def ensure_ets_cost_columns(
    ra,
    *,
    equipment: str,
    update_from: date,
    update_to: date,
    agreement_from: date | None,
    agreement_to: date | None,
    dft_periods: list[tuple[date, date]] | None = None,
) -> tuple[list[CostColumn], str | None]:
    """
    EU ETS: prefer a column covering the update (typically full rate-card /
    agreement-style validity). If none, create for the exact update days.

    DFT: prefer existing covering / exact periods; overlapping-but-misaligned
    validity → issue (do not create sliced columns like 15.07-30.09 when
    15.07-14.10 already exists).
    """
    eq_u = equipment.upper()

    if eq_u == "DFT" and dft_periods is not None:
        return resolve_dft_period_columns(
            ra,
            fee_name="EU ETS Fee",
            update_from=update_from,
            update_to=update_to,
            dft_periods=dft_periods,
            template_fee_fallback="BAF Fee",
        )

    same = [
        c
        for c in ra.cost_columns
        if c.fee_name == "EU ETS Fee" and c.equipment == equipment
    ]

    # 1) Exact period (ignore spaces around date dash in header text)
    exact = match_period_fee_columns(
        ra.cost_columns,
        fee_name="EU ETS Fee",
        equipment=eq_u,
        valid_from=update_from,
        valid_to=update_to,
    )
    if exact:
        return [exact[0]], None

    # 2) Single column that fully contains the update
    containers = [
        c
        for c in same
        if c.valid_from is not None
        and c.valid_to is not None
        and period_inside(update_from, update_to, c.valid_from, c.valid_to)
    ]
    if len(containers) == 1:
        return [containers[0]], None
    if len(containers) > 1:
        # Prefer agreement-wide / longest
        best = max(
            containers,
            key=lambda c: (c.valid_to - c.valid_from).days if c.valid_from and c.valid_to else 0,
        )
        return [best], None

    # 3) Agreement-validity column if present and covers update
    if agreement_from and agreement_to:
        agr = match_period_fee_columns(
            ra.cost_columns,
            fee_name="EU ETS Fee",
            equipment=eq_u,
            valid_from=agreement_from,
            valid_to=agreement_to,
        )
        for c in agr:
            if period_inside(update_from, update_to, c.valid_from, c.valid_to):
                return [c], None

    # 4) Create for exact stated days
    template = find_template_period_cost(ra, fee_name="EU ETS Fee", equipment=equipment)
    if template is None:
        # fall back to any BAF template for layout
        template = find_template_period_cost(ra, fee_name="BAF Fee", equipment=equipment)
    if template is None:
        return [], f"No EU ETS Fee template column available to clone for {equipment}"

    created = create_period_cost_column(
        ra,
        fee_name="EU ETS Fee",
        equipment=equipment,
        valid_from=update_from,
        valid_to=update_to,
        template=template,
        prolong_from=None,
        rate_by=DFT_WEIGHT_RATE_BY if equipment.upper() == "DFT" else None,
    )
    print(f"  Created {created.header}")
    return [created], None
