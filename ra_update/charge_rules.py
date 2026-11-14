"""Mapping rules: update Rate Card column names -> RA Rate Card cost headers."""

from __future__ import annotations

import re
from dataclasses import dataclass

# Services explicitly out of scope for this update process
OUT_OF_SCOPE_SERVICES = {
    "OC_CNTR_CFS-CFS",
    "OC_CNTR_D-CY",
}

OUT_OF_SCOPE_SERVICE_REASON = {
    "OC_CNTR_CFS-CFS": "SERVICE__C OC_CNTR_CFS-CFS is not in scope",
    "OC_CNTR_D-CY": "SERVICE__C OC_CNTR_D-CY is out of scope",
}


# Rate-by text for RoRo/DFT weight-based BAF / EU ETS (matches RA "Rate by:" cell).
DFT_WEIGHT_RATE_BY = "Weight/chargeable kg"


@dataclass(frozen=True)
class ChargeRule:
    """How one update *_RATE column maps onto an RA cost header."""

    charge_key: str
    ra_fee_name: str
    # Fixed equipment/token in RA header body, e.g. DFT, BASE_STAT_FRK, PCANAL.
    # For CNTR_* rules this is None and equipment comes from the column name.
    ra_equipment: str | None
    require_period: bool
    sheets: frozenset[str]
    out_of_scope: bool = False
    out_of_scope_reason: str | None = None
    # If True, equipment is parsed from [EQ]CNTR_... column name
    from_equipment_col: bool = False
    # Optional RA "Rate by:" value when creating / verifying the cost column
    rate_by: str | None = None


# DFT / special columns (prefix of column without _RATE/_CURRENCY/_MIN)
DFT_COLUMN_RULES: dict[str, ChargeRule] = {
    "*DFTBASE_SLF_PRP": ChargeRule(
        charge_key="DFTBASE_SLF_PRP",
        ra_fee_name="",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        out_of_scope=True,
        out_of_scope_reason="PRP is out of scope (confirmed in IR2025101300002 (UAT))",
    ),
    "*DFTBASE_STAT_FRK": ChargeRule(
        charge_key="DFTBASE_STAT_FRK",
        ra_fee_name="Base Rate",
        ra_equipment="BASE_STAT_FRK",
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
    "*DFTBIOFUEL": ChargeRule(
        charge_key="DFTBIOFUEL",
        ra_fee_name="Biofuel Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
    "*DFTDOC_FEE": ChargeRule(
        charge_key="DFTDOC_FEE",
        ra_fee_name="Documentation Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
    "*DFTDOC": ChargeRule(
        charge_key="DFTDOC",
        ra_fee_name="Documentation Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
    "*DFTROLL_TR_NONFRK": ChargeRule(
        charge_key="DFTROLL_TR_NONFRK",
        ra_fee_name="RoRo Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
    "*DFTWM_COGH": ChargeRule(
        charge_key="DFTWM_COGH",
        ra_fee_name="Handled Goods Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
    # *DFTWM_SCANAL → fee-only Suez Canal Charge (no equipment), Weight/chargeable kg
    "*DFTWM_SCANAL_O": ChargeRule(
        charge_key="DFTWM_SCANAL",
        ra_fee_name="Suez Canal Charge",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        rate_by=DFT_WEIGHT_RATE_BY,
    ),
    "*DFTWM_SCANAL": ChargeRule(
        charge_key="DFTWM_SCANAL",
        ra_fee_name="Suez Canal Charge",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        rate_by=DFT_WEIGHT_RATE_BY,
    ),
    "*DFTWM_PCANAL_O": ChargeRule(
        charge_key="DFTWM_PCANAL",
        ra_fee_name="Panama Canal Charge",
        ra_equipment="PCANAL",
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
    "*DFTWM_PCANAL": ChargeRule(
        charge_key="DFTWM_PCANAL",
        ra_fee_name="Panama Canal Charge",
        ra_equipment="PCANAL",
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
    # *DFTWM_WAR_RISK → War Risk Fee (no equipment), Rate by: Weight/chargeable kg
    "*DFTWM_WAR_RISK": ChargeRule(
        charge_key="DFTWM_WAR_RISK",
        ra_fee_name="War Risk Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        rate_by=DFT_WEIGHT_RATE_BY,
    ),
    # *DFTBAF → BAF Fee (DFT <period>), Rate by: Weight/chargeable kg
    "*DFTBAF": ChargeRule(
        charge_key="DFTBAF",
        ra_fee_name="BAF Fee",
        ra_equipment="DFT",
        require_period=True,
        sheets=frozenset({"ETSBAF"}),
        rate_by=DFT_WEIGHT_RATE_BY,
    ),
    # *DFTEU_ETS → EU ETS Fee (DFT <period>), Rate by: Weight/chargeable kg
    "*DFTEU_ETS": ChargeRule(
        charge_key="DFTEU_ETS",
        ra_fee_name="EU ETS Fee",
        ra_equipment="DFT",
        require_period=True,
        sheets=frozenset({"ETSBAF"}),
        rate_by=DFT_WEIGHT_RATE_BY,
    ),
    # USTR301 variants if present in pivot
    "*DFTUSTR301": ChargeRule(
        charge_key="USTR301",
        ra_fee_name="USTR301 Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
    "USTR301": ChargeRule(
        charge_key="USTR301",
        ra_fee_name="USTR301 Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
    ),
}

# Equipment CNTR charge suffix -> rule template (equipment filled from column)
CNTR_CHARGE_RULES: dict[str, ChargeRule] = {
    "BASE": ChargeRule(
        charge_key="BASE",
        ra_fee_name="Base Rate",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        from_equipment_col=True,
    ),
    "DTHC": ChargeRule(
        charge_key="DTHC",
        ra_fee_name="Destination Terminal Handling Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        from_equipment_col=True,
    ),
    "OTHC": ChargeRule(
        charge_key="OTHC",
        ra_fee_name="Origin Terminal Handling Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        from_equipment_col=True,
    ),
    "WAR_RISK": ChargeRule(
        charge_key="WAR_RISK",
        ra_fee_name="War Risk Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        from_equipment_col=True,
    ),
    "WAR_RISK_C": ChargeRule(
        charge_key="WAR_RISK",
        ra_fee_name="War Risk Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        from_equipment_col=True,
    ),
    "DCFS": ChargeRule(
        charge_key="DCFS",
        ra_fee_name="Destination CFS Fee",
        ra_equipment=None,
        require_period=False,
        sheets=frozenset({"BASE"}),
        from_equipment_col=True,
    ),
    "BAF": ChargeRule(
        charge_key="BAF",
        ra_fee_name="BAF Fee",
        ra_equipment=None,
        require_period=True,
        sheets=frozenset({"ETSBAF"}),
        from_equipment_col=True,
    ),
    "ETS": ChargeRule(
        charge_key="ETS",
        ra_fee_name="EU ETS Fee",
        ra_equipment=None,
        require_period=True,
        sheets=frozenset({"ETSBAF"}),
        from_equipment_col=True,
    ),
}

CNTR_RATE_RE = re.compile(
    r"^(.+)CNTR_(BASE|DTHC|OTHC|DCFS|WAR_RISK_C|WAR_RISK|BAF|ETS)_RATE$",
    re.IGNORECASE,
)

# Update/RA equipment codes that share the same Rate-by container type.
# Alias only within a group (e.g. 20G0/22G0/22SOC = Container/20FT).
# Solo groups (e.g. 42P1 = Container/40FR) have no alias — must be created.
# Order within each group is preference when several codes exist on the Rate Card.
EQUIPMENT_ALIAS_GROUPS: tuple[tuple[str, ...], ...] = (
    ("20G0", "22G0", "22SOC"),  # Container/20FT
    ("22P1",),  # Container/20FR
    ("22P1B",),  # Container/20FROHOW
    ("22P1H",),  # Container/20FROH
    ("22P1W",),  # Container/20FROW
    ("22RT",),  # Container/20RE
    ("22UT", "22U0"),  # Container/20OT
    ("22UTH",),  # Container/20OTOH
    ("25G0",),  # Container/20HC
    ("42G0", "42P1S", "42SOC"),  # Container/40FT
    ("42P1",),  # Container/40FR
    ("42P1B",),  # Container/40FROHOW
    ("42P1H",),  # Container/40FROH
    ("42P1W",),  # Container/40FTOW
    ("42RT",),  # Container/40RE
    ("42UT", "42U0"),  # Container/40OT
    ("42UTH",),  # Container/40OTOH
    ("45G0", "L2G0"),  # Container/40HC
    ("45RT",),  # Container/40HCRE
    ("45UT", "45U0"),  # Container/40HCOT
    ("45UTH",),  # Container/40HCOTOH
    ("45SOC",),  # Container/40HCSO
)

# Canonical Rate-by label when creating a new equipment cost column
EQUIPMENT_RATE_BY: dict[str, str] = {
    "20G0": "Container/20FT",
    "22G0": "Container/20FT",
    "22SOC": "Container/20FT",
    "22P1": "Container/20FR",
    "22P1B": "Container/20FROHOW",
    "22P1H": "Container/20FROH",
    "22P1W": "Container/20FROW",
    "22RT": "Container/20RE",
    "22UT": "Container/20OT",
    "22U0": "Container/20OT",
    "22UTH": "Container/20OTOH",
    "25G0": "Container/20HC",
    "42G0": "Container/40FT",
    "42P1S": "Container/40FT",
    "42SOC": "Container/40FT",
    "42P1": "Container/40FR",
    "42P1B": "Container/40FROHOW",
    "42P1H": "Container/40FROH",
    "42P1W": "Container/40FTOW",
    "42RT": "Container/40RE",
    "42UT": "Container/40OT",
    "42U0": "Container/40OT",
    "42UTH": "Container/40OTOH",
    "45G0": "Container/40HC",
    "L2G0": "Container/40HC",
    "45RT": "Container/40HCRE",
    "45UT": "Container/40HCOT",
    "45U0": "Container/40HCOT",
    "45UTH": "Container/40HCOTOH",
    "45SOC": "Container/40HCSO",
}

_EQUIPMENT_TO_GROUP: dict[str, tuple[str, ...]] = {
    code: group for group in EQUIPMENT_ALIAS_GROUPS for code in group
}


def equipment_alias_group(equipment: str | None) -> tuple[str, ...]:
    """Return preference-ordered alias group for an equipment code (or itself)."""
    if not equipment:
        return tuple()
    eq = str(equipment).strip().upper()
    return _EQUIPMENT_TO_GROUP.get(eq, (eq,))


def preferred_alias_template(
    equipment: str | None,
    present_equipments: set[str],
) -> str | None:
    """
    Suggest an existing RA equipment in the *same Rate-by alias group*.
    Never cross groups (42P1 must not suggest 22G0).
    """
    if not equipment:
        return None
    eq = str(equipment).strip().upper()
    present = {str(p).strip().upper() for p in present_equipments if p}
    group = equipment_alias_group(eq)
    # Solo group → no alias peer
    if len(group) <= 1:
        return None
    for code in group:
        if code != eq and code in present:
            return code
    return None


def equipment_rate_by(equipment: str | None) -> str | None:
    """Canonical Rate-by for an equipment code, if known."""
    if not equipment:
        return None
    return EQUIPMENT_RATE_BY.get(str(equipment).strip().upper())


def same_equipment_alias_group(a: str | None, b: str | None) -> bool:
    """True when both codes belong to the same static alias group (suggestion table)."""
    if not a or not b:
        return False
    ga = equipment_alias_group(a)
    gb = equipment_alias_group(b)
    if not ga or not gb:
        return False
    return ga == gb


def resolve_ra_equipment(
    update_equipment: str | None,
    present_equipments: set[str],
    *,
    runtime_aliases: dict[str, str] | None = None,
) -> str | None:
    """
    Pick which RA equipment token to use for an update equipment code.

    1) Exact match on RA
    2) Confirmed runtime alias (from user decision)
    3) Otherwise None (do not auto-apply static groups)
    """
    if not update_equipment:
        return None
    eq = str(update_equipment).strip().upper()
    present = {str(p).strip().upper() for p in present_equipments if p}
    if eq in present:
        return eq
    if runtime_aliases:
        mapped = runtime_aliases.get(eq)
        if mapped and mapped in present:
            return mapped
        if mapped:
            return mapped  # allow even if presence check races after create
    return None


def normalize_rate_col_base(col_name: str) -> str:
    """Strip trailing _RATE and normalize."""
    name = str(col_name).strip()
    if name.upper().endswith("_RATE"):
        name = name[: -len("_RATE")]
    return name


def resolve_rate_column(col_name: str, sheet: str) -> tuple[ChargeRule | None, str | None, str | None]:
    """
    Resolve update column to a ChargeRule.

    Returns (rule, equipment_from_col, skip_reason).
    skip_reason set when column is known out-of-scope or wrong sheet / unknown.
    """
    sheet_u = sheet.upper()
    base = normalize_rate_col_base(col_name)
    base_u = base.upper()

    # Exact DFT / special match (case-insensitive on keys)
    for key, rule in DFT_COLUMN_RULES.items():
        if base_u == key.upper():
            if rule.out_of_scope:
                return rule, None, rule.out_of_scope_reason
            if sheet_u not in rule.sheets:
                return (
                    rule,
                    rule.ra_equipment,
                    f"Charge {rule.charge_key} not applicable on sheet {sheet_u}",
                )
            return rule, rule.ra_equipment, None

    # Prefix match for DFT columns with unexpected suffixes already stripped
    # (longest key first)
    dft_keys = sorted(DFT_COLUMN_RULES.keys(), key=len, reverse=True)
    for key in dft_keys:
        if base_u.startswith(key.upper()):
            rule = DFT_COLUMN_RULES[key]
            if rule.out_of_scope:
                return rule, None, rule.out_of_scope_reason
            if sheet_u not in rule.sheets:
                return (
                    rule,
                    rule.ra_equipment,
                    f"Charge {rule.charge_key} not applicable on sheet {sheet_u}",
                )
            return rule, rule.ra_equipment, None

    m = CNTR_RATE_RE.match(str(col_name).strip())
    if m:
        equipment = m.group(1).upper()
        charge_suffix = m.group(2).upper()
        rule = CNTR_CHARGE_RULES.get(charge_suffix)
        if rule is None:
            return None, equipment, f"Charge type {charge_suffix} not covered by current rules"
        if sheet_u not in rule.sheets:
            return (
                rule,
                equipment,
                f"Charge type {rule.charge_key} not applicable on sheet {sheet_u}",
            )
        return rule, equipment, None

    return None, None, "Charge column not covered by current rules"


def service_out_of_scope_reason(service_c: str) -> str | None:
    if not service_c:
        return None
    # Exact service token match (also if multi-value, any out-of-scope token)
    parts = [p.strip() for p in str(service_c).split("/") if p.strip()]
    for part in parts:
        if part in OUT_OF_SCOPE_SERVICES:
            return OUT_OF_SCOPE_SERVICE_REASON[part]
    return None
