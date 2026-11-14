from __future__ import annotations

import re
from dataclasses import dataclass

from .charge_rules import resolve_ra_equipment
from .dates import format_ra_date, parse_date


COST_PERIOD_RE = re.compile(
    r"^(?P<name>.+?)\s*\((?P<body>[^)]*)\)\s*$"
)
# Sea FR exports often put spaces around the date dash:
#   BAF Fee (22G0 15.07.2026 - 14.08.2026)
BODY_WITH_PERIOD_RE = re.compile(
    r"^(?P<eq>\S+)\s+(?P<from>\d{2}\.\d{2}\.\d{4})\s*-\s*(?P<to>\d{2}\.\d{2}\.\d{4})$"
)
# Weight / RoRo style: "BAF Fee (15.07.2026-14.08.2026)" — no equipment token
BODY_PERIOD_ONLY_RE = re.compile(
    r"^(?P<from>\d{2}\.\d{2}\.\d{4})\s*-\s*(?P<to>\d{2}\.\d{2}\.\d{4})$"
)
BODY_EQ_ONLY_RE = re.compile(r"^(?P<eq>\S+)$")


@dataclass(frozen=True)
class CostColumn:
    start_col: int  # cost name lives here (row 3)
    header: str
    fee_name: str  # e.g. BAF Fee
    equipment: str | None
    valid_from: object | None
    valid_to: object | None
    currency_col: int | None = None
    flat_col: int | None = None
    # True when row-6 above Flat is labeled MIN (Flat stores minimum)
    flat_is_min: bool = False
    punit_col: int | None = None

    @property
    def rate_col(self) -> int | None:
        """Column that stores the main rate (p/unit, else Flat when Flat is not MIN)."""
        if self.punit_col is not None:
            return self.punit_col
        if self.flat_col is not None and not self.flat_is_min:
            return self.flat_col
        return None

    @property
    def min_col(self) -> int | None:
        """Column that stores MIN (Flat labeled as MIN)."""
        if self.flat_col is not None and self.flat_is_min:
            return self.flat_col
        return None


def parse_cost_header(header: str) -> tuple[str, str | None, object | None, object | None]:
    """
    Parse 'BAF Fee (22G0 15.09.2026-14.10.2026)', 'BAF Fee (15.07.2026-14.08.2026)',
    'Base Rate (42G0)', or 'Biofuel Fee'.
    Returns (fee_name, equipment, valid_from, valid_to).
    """
    text = str(header).strip()
    m = COST_PERIOD_RE.match(text)
    if not m:
        return text, None, None, None

    fee_name = m.group("name").strip()
    body = m.group("body").strip()

    mp = BODY_WITH_PERIOD_RE.match(body)
    if mp:
        return (
            fee_name,
            mp.group("eq").upper(),
            parse_date(mp.group("from")),
            parse_date(mp.group("to")),
        )

    # Period only (no DFT / equipment remark): BAF Fee (15.07.2026-14.08.2026)
    mpo = BODY_PERIOD_ONLY_RE.match(body)
    if mpo:
        return fee_name, None, parse_date(mpo.group("from")), parse_date(mpo.group("to"))

    me = BODY_EQ_ONLY_RE.match(body)
    if me:
        return fee_name, me.group("eq").upper(), None, None

    # Unusual bodies (e.g. BASE_STAT_FRK) — keep as equipment token
    return fee_name, body.upper(), None, None


def period_fee_header(fee_name: str, equipment: str | None, valid_from, valid_to) -> str:
    """
    Build RA cost header including equipment when present:
      BAF Fee (DFT 15.07.2026-14.08.2026)
      BAF Fee (42G0 15.07.2026-14.08.2026)
    Period-only (no equipment) is used only as a match fallback for older cards.
    """
    span = f"{format_ra_date(valid_from)}-{format_ra_date(valid_to)}"
    eq = (equipment or "").strip().upper()
    if not eq:
        return f"{fee_name} ({span})"
    return f"{fee_name} ({eq} {span})"


def match_period_fee_columns(
    cost_columns: list[CostColumn],
    *,
    fee_name: str,
    equipment: str | None,
    valid_from,
    valid_to,
) -> list[CostColumn]:
    """
    Match period fee columns by fee + equipment + dates (not exact header text).

    Sea FR headers often use spaces around the date dash
    (`15.07.2026 - 14.08.2026`); generated headers do not.
    """
    eq = (equipment or "").strip().upper() or None
    out: list[CostColumn] = []
    for c in cost_columns:
        if c.fee_name != fee_name:
            continue
        c_eq = (c.equipment or "").strip().upper() or None
        if c_eq != eq:
            continue
        if c.valid_from == valid_from and c.valid_to == valid_to:
            out.append(c)
    return out


def legacy_dft_period_header(fee_name: str, valid_from, valid_to) -> str:
    """Alias for DFT-tagged header: BAF Fee (DFT 15.07.2026-14.08.2026)."""
    return period_fee_header(fee_name, "DFT", valid_from, valid_to)


def weight_period_candidates(
    cost_columns: list[CostColumn],
    *,
    fee_name: str,
) -> list[CostColumn]:
    """
    Columns for *DFTBAF / *DFTEU_ETS: headers with DFT equipment,
    plus period-only headers that some older cards still use.
    """
    out: list[CostColumn] = []
    for c in cost_columns:
        if c.fee_name != fee_name:
            continue
        if c.valid_from is None or c.valid_to is None:
            continue
        if c.equipment is None or c.equipment == "DFT":
            out.append(c)
    return out


def find_cost_columns(
    cost_columns: list[CostColumn],
    *,
    ra_fee_name: str,
    equipment: str | None,
    valid_from,
    valid_to,
    require_period: bool,
    runtime_aliases: dict[str, str] | None = None,
) -> tuple[list[CostColumn], str | None]:
    """
    Find one or more cost columns for the update.

    For period-based fees (BAF/ETS/DFT):
      1) exact period header match
      2) else a set of candidate periods that cover the update without gaps
      3) else a unique period-less column for that fee+equipment

    `runtime_aliases` maps update equipment -> RA equipment after user confirmation
    (e.g. {"22SOC": "22G0"}). Static alias groups are NOT applied automatically.
    """
    from .period_cover import describe_cover_failure, periods_cover_without_gaps

    fee = ra_fee_name
    eq = equipment.upper() if equipment else None

    # Fee-only costs (Biofuel Fee, Documentation Fee, USTR301 Fee, ...)
    if eq is None:
        matches = [
            c
            for c in cost_columns
            if c.fee_name == fee and c.equipment is None
        ]
        if not matches:
            matches = [c for c in cost_columns if c.header == fee]
        if len(matches) == 1:
            return [matches[0]], None
        if len(matches) > 1:
            return [], f"Ambiguous cost columns for {fee}: {[c.header for c in matches]}"
        # Fee-only only (e.g. weight War Risk Fee) — do not fall back to
        # container-scoped War Risk Fee (22G0)/(42G0)/...
        return [], f"Cost column not found: {fee}"

    update_eq = eq
    # *DFTBAF / *DFTEU_ETS: prefer BAF/ETS Fee (DFT dates); also accept period-only
    if eq == "DFT" and require_period:
        same_fee_eq = weight_period_candidates(cost_columns, fee_name=fee)
        if valid_from is not None and valid_to is not None:
            for want_eq in ("DFT", None):
                by_exact = match_period_fee_columns(
                    cost_columns,
                    fee_name=fee,
                    equipment=want_eq,
                    valid_from=valid_from,
                    valid_to=valid_to,
                )
                if len(by_exact) == 1:
                    return [by_exact[0]], None
                if len(by_exact) > 1:
                    label = period_fee_header(fee, want_eq, valid_from, valid_to)
                    return [], f"Multiple RA cost columns named {label!r}"
        if valid_from is not None and valid_to is not None:
            covered = periods_cover_without_gaps(valid_from, valid_to, same_fee_eq)
            if covered:
                return covered, None
            if same_fee_eq:
                return [], describe_cover_failure(
                    valid_from, valid_to, fee, "DFT", same_fee_eq
                )
        return [], f"Cost column not found: {fee} (DFT <period>)"

    present_for_fee = {
        c.equipment for c in cost_columns if c.fee_name == fee and c.equipment
    }
    resolved = resolve_ra_equipment(
        eq, present_for_fee, runtime_aliases=runtime_aliases
    )
    if resolved is None:
        wanted = f"{fee} ({update_eq})"
        return [], f"Cost column not found: {wanted}"
    eq = resolved

    exact_name = None
    if valid_from is not None and valid_to is not None:
        exact_name = period_fee_header(fee, eq, valid_from, valid_to)
        by_exact = match_period_fee_columns(
            cost_columns,
            fee_name=fee,
            equipment=eq,
            valid_from=valid_from,
            valid_to=valid_to,
        )
        if by_exact:
            if len(by_exact) > 1:
                return [], f"Multiple RA cost columns named {exact_name!r}"
            return [by_exact[0]], None

    same_fee_eq = [
        c
        for c in cost_columns
        if c.fee_name == fee and c.equipment == eq
    ]

    # Non-period header: "Base Rate (BASE_STAT_FRK)", "Suez Canal Charge (SCANAL)"
    no_period_name = f"{fee} ({eq})"
    by_name = [c for c in cost_columns if c.header == no_period_name]
    if by_name and not require_period:
        if len(by_name) == 1:
            return [by_name[0]], None
        return [], f"Ambiguous cost columns named {no_period_name!r}"

    if require_period:
        if valid_from is not None and valid_to is not None:
            covered = periods_cover_without_gaps(valid_from, valid_to, same_fee_eq)
            if covered:
                return covered, None

        period_less = [c for c in same_fee_eq if c.valid_from is None and c.valid_to is None]
        if len(period_less) == 1:
            return [period_less[0]], None
        if not same_fee_eq:
            wanted = exact_name or f"{fee} ({eq} <period>)"
            return [], f"Cost column not found: {wanted}"
        if valid_from is not None and valid_to is not None:
            return [], describe_cover_failure(valid_from, valid_to, fee, eq, same_fee_eq)
        return [], f"Cost column not found: {fee} ({eq})"

    if len(same_fee_eq) == 1:
        return [same_fee_eq[0]], None
    if not same_fee_eq:
        return [], f"Cost column not found: {fee} ({eq})"
        # # Alias-aware not-found message (re-enable with alias resolution above):
        # aliases = "/".join(alias_group) if alias_group else update_eq
        # return [], f"Cost column not found: {fee} ({update_eq}) (aliases tried: {aliases})"
    return [], f"Ambiguous cost columns for {fee} ({eq}): {[c.header for c in same_fee_eq]}"


def find_cost_column(
    cost_columns: list[CostColumn],
    *,
    ra_fee_name: str,
    equipment: str | None,
    valid_from,
    valid_to,
    require_period: bool,
    runtime_aliases: dict[str, str] | None = None,
) -> tuple[CostColumn | None, str | None]:
    """Backward-compatible single-column finder."""
    cols, err = find_cost_columns(
        cost_columns,
        ra_fee_name=ra_fee_name,
        equipment=equipment,
        valid_from=valid_from,
        valid_to=valid_to,
        require_period=require_period,
        runtime_aliases=runtime_aliases,
    )
    if err:
        return None, err
    if not cols:
        return None, "Cost column not found"
    return cols[0], None
