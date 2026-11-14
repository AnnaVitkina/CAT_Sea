"""
Lane identity fields: SERVICE short map, port/postal aliases, Transporeon ID,
and sync of update values onto Rate Card rows (merge mismatches with / + orange).
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.styles import PatternFill

from .config import (
    CHANGED_FILL_HEX,
    COL_CARRIER,
    COL_DEST_COUNTRY,
    COL_DEST_PORT,
    COL_DEST_POSTAL,
    COL_KEY,
    COL_ORIGIN_COUNTRY,
    COL_ORIGIN_PORT,
    COL_ORIGIN_POSTAL,
    COL_SERVICE,
    COL_SERVICE_C,
    COL_SERVICE_SHORT,
    COL_TRANSPOREON,
    MISMATCH_FILL_HEX,
)
from .parsing import split_multi

GREEN = PatternFill(start_color=CHANGED_FILL_HEX, end_color=CHANGED_FILL_HEX, fill_type="solid")
ORANGE = PatternFill(start_color=MISMATCH_FILL_HEX, end_color=MISMATCH_FILL_HEX, fill_type="solid")

DEFAULT_SERVICE_TEXT = "not PRECARRIAGE/ONCARRIAGE"

# SERVICE__C (any of) → Service short on Rate Card
SERVICE_SHORT_MAP: dict[str, str] = {
    "OC_CNTR_BC-CY": "BC-CY",
    "OC_CNTR_CFS-CY": "CFS-CY",
    "OC_CNTR_CY-CY": "CY-CY",
    "OC_CNTR_CY-CY_BU": "CY-CY",
    "OC_CNTR_CY-RR": "CY-RR",
    "OC_CNTR_CY-RR_BU": "CY-RR",
    "OC_CNTR_RR-CY": "RR-CY",
    "OC_CNTR_RR-CY_BU": "RR-CY",
    "OC_CNTR_SPOT": "CY-CY",
    "SPOT": "CY-CY",
    "OC_LOLO": "LOLO",
    "OC_LOLO_SPOT": "LOLO",
    "OC_RRBB": "RRBB",
    "OC_RRBB_SPOT": "RRBB",
}

# Port / postal alias groups. First code is the preferred lead when expanding.
# Prefer the full 3-code form when that is the defined group (e.g. CNTXG/CNTSN/CNXGA).
PORT_ALIAS_GROUPS: tuple[tuple[str, ...], ...] = (
    ("CNSHA", "CNSHG", "CNSGH"),
    ("CNTXG", "CNTSN", "CNXGA"),
    ("USORF", "USNFK"),
    ("USSSI", "USBCK"),
)

POSTAL_ALIAS_GROUPS: tuple[tuple[str, ...], ...] = (
    ("USCHI_USORF", "USCHI_USNFK"),
)

_PORT_TO_GROUP: dict[str, tuple[str, ...]] = {
    code.upper(): group for group in PORT_ALIAS_GROUPS for code in group
}
_POSTAL_TO_GROUP: dict[str, tuple[str, ...]] = {
    code.upper(): group for group in POSTAL_ALIAS_GROUPS for code in group
}


def derive_service_short(service_c: str) -> str:
    """Map SERVICE__C to Rate Card Service short (CY-CY, RRBB, ...)."""
    primary = (service_c or "").split("/")[0].strip()
    if not primary:
        return ""
    hit = SERVICE_SHORT_MAP.get(primary)
    if hit:
        return hit
    hit = SERVICE_SHORT_MAP.get(primary.upper())
    if hit:
        return hit
    # Fallback for unlisted codes
    if primary.startswith("OC_CNTR_"):
        rest = primary[len("OC_CNTR_") :]
        if rest.endswith("_BU"):
            rest = rest[: -len("_BU")]
        if rest.upper() == "SPOT":
            return "CY-CY"
        return rest
    if primary.upper() == "SPOT":
        return "CY-CY"
    if primary.startswith("OC_"):
        return primary[3:]
    return primary


def _all_tokens(value: object) -> list[str]:
    """Split on both '//' chunks and '/' alternatives; drop empties."""
    text = "" if value is None else str(value).strip()
    if not text:
        return []
    out: list[str] = []
    for chunk in text.split("//"):
        for part in split_multi(chunk):
            if part:
                out.append(part)
    return out


def join_slash(tokens: list[str]) -> str:
    """Join multi-values with '/' and no spaces: A/B not A / B."""
    return "/".join(t.strip() for t in tokens if t and str(t).strip())


def normalize_slash_value(value: object) -> str:
    """
    Canonical multi-value form: tokens joined by '/' with no surrounding spaces.
    Also collapses '//' and ' / ' from older RA / merge styles.
    """
    return join_slash(_all_tokens(value))


def expand_location_aliases(code: str, *, postal: bool) -> str:
    """
    Expand a location code to its full alias group joined by '/'.
    Example: CNSHA → CNSHA/CNSHG/CNSGH
             CNTXG → CNTXG/CNTSN/CNXGA
    """
    text = (code or "").strip()
    if not text:
        return ""
    group_map = _POSTAL_TO_GROUP if postal else _PORT_TO_GROUP
    # Expand each token (from / or //), then unique-preserve group order
    seen: set[str] = set()
    out: list[str] = []
    for part in _all_tokens(text):
        group = group_map.get(part.upper())
        codes = group if group else (part,)
        for c in codes:
            key = c.upper()
            if key not in seen:
                seen.add(key)
                out.append(c)
    return "/".join(out)


def location_port_postal(code: str) -> tuple[str, str]:
    """
    ORIGIN/DESTINATION_LOCATION_NAME__C → (port, postal).
    Underscore codes go to postal; plain codes to port.
    Both sides get alias expansion.
    """
    text = (code or "").strip()
    if not text:
        return "", ""
    if "_" in text.split("/")[0]:
        return "", expand_location_aliases(text, postal=True)
    return expand_location_aliases(text, postal=False), ""


def _primary_location_code(value: object) -> str:
    """First location token only — Transporeon IDs must not embed alias slash lists."""
    toks = _all_tokens(value)
    return toks[0] if toks else ""


def build_transporeon_id(
    *,
    carrier: str,
    service_short: str,
    origin_port: str,
    origin_postal: str,
    dest_port: str,
    dest_postal: str,
) -> str:
    """
    Carrier + Service + Origin Port + Origin Postal + Dest Port + Dest Postal.

    Uses a single primary code per location field (not CNSHA/CNSHG/CNSGH).
    Alias expansions belong in Origin/Destination Port and Postal columns only.
    """
    return (
        f"{carrier or ''}"
        f"{service_short or ''}"
        f"{_primary_location_code(origin_port)}"
        f"{_primary_location_code(origin_postal)}"
        f"{_primary_location_code(dest_port)}"
        f"{_primary_location_code(dest_postal)}"
    )


def _norm_parts(value: object) -> set[str]:
    return {p.upper() for p in _all_tokens(value)}


def _alias_family(token: str) -> frozenset[str]:
    t = (token or "").strip().upper()
    if not t:
        return frozenset()
    group = _PORT_TO_GROUP.get(t) or _POSTAL_TO_GROUP.get(t)
    if group:
        return frozenset(x.upper() for x in group)
    return frozenset({t})


def _wanted_already_present(existing: object, wanted: str) -> bool:
    """True if every update token is already in the RA cell (split on / and //)."""
    e_norm = _norm_parts(existing)
    w_tokens = _all_tokens(wanted)
    if not w_tokens:
        return True
    return all(t.upper() in e_norm for t in w_tokens)


def _values_equivalent(existing: object, wanted: str) -> bool:
    """
    True when strings match, wanted already present, or RA tokens are the same
    alias family as wanted (e.g. RA=CNSGH or CNTXG/CNTSN/CNXGA//CNTXG/CNTSN,
    wanted=CNSHA/CNSHG/CNSGH or CNTXG/CNTSN/CNXGA → equivalent → expand/clean).
    """
    if wanted is None:
        wanted = ""
    exp = "" if existing is None else str(existing).strip()
    w = str(wanted).strip()
    if exp == w:
        return True
    if not exp and not w:
        return True
    if _wanted_already_present(exp, w) and _wanted_already_present(w, exp):
        return True
    e_parts = _norm_parts(exp)
    w_parts = _norm_parts(w)
    if e_parts == w_parts and bool(w_parts):
        return True
    if not e_parts or not w_parts:
        return False
    # Port/postal: RA tokens ⊆ wanted alias family (allows cleanup to full 3-code form)
    wanted_family: set[str] = set()
    for p in w_parts:
        wanted_family |= _alias_family(p)
    if e_parts <= wanted_family and w_parts <= wanted_family:
        return True
    # Wanted already fully covered by existing tokens (KEY / SERVICE__C / postal)
    if _wanted_already_present(exp, w):
        return True
    return False


def _merge_mismatch(existing: object, wanted: str) -> str:
    """
    Append update value with / only when it adds something new (no spaces).
    Never re-append KEY / SERVICE / postal / port tokens already present.
    """
    exp = "" if existing is None else str(existing).strip()
    w = (wanted or "").strip()
    if not exp:
        return normalize_slash_value(w)
    if not w:
        return normalize_slash_value(exp)
    if _wanted_already_present(exp, w):
        return normalize_slash_value(exp)
    # Append only tokens that are not already in the cell; rebuild without spaces
    e_tokens = _all_tokens(exp)
    e_norm = {t.upper() for t in e_tokens}
    new_tokens = [t for t in _all_tokens(w) if t.upper() not in e_norm]
    if not new_tokens:
        return normalize_slash_value(exp)
    return join_slash(e_tokens + new_tokens)


@dataclass
class IdentitySyncResult:
    field: str
    col: int
    old_value: object
    new_value: str
    status: str  # unchanged | set | expanded | mismatch_merged


def sync_lane_identity_from_update(
    ra,
    lane,
    *,
    key: str,
    carrier: str,
    service_c: str,
    origin_country: str,
    dest_country: str,
    origin_code: str,
    dest_code: str,
) -> list[IdentitySyncResult]:
    """
    Align Rate Card identity columns with update-derived values.

    - Empty RA cell → set update value (green)
    - Equivalent (incl. port alias sets) → normalize to expanded form if needed (green)
    - Different → keep RA value and append update with / (orange)
    """
    ws = ra.ws
    row = lane.row
    service_short = derive_service_short(service_c)
    origin_port, origin_postal = location_port_postal(origin_code)
    dest_port, dest_postal = location_port_postal(dest_code)
    # Transporeon: keep RA value when present. Alias lists go in port/postal cols only.
    # Build a primary-code ID only to fill blank Transporeon cells.
    transporeon = build_transporeon_id(
        carrier=carrier,
        service_short=service_short,
        origin_port=origin_port,
        origin_postal=origin_postal,
        dest_port=dest_port,
        dest_postal=dest_postal,
    )

    # Always store multi-values as A/B (no spaces around /)
    key_n = normalize_slash_value(key or "")
    service_c_n = normalize_slash_value(service_c or "")

    wanted: list[tuple[str, int, str]] = [
        ("KEY", COL_KEY, key_n),
        ("Carrier", COL_CARRIER, carrier or ""),
        ("SERVICE", COL_SERVICE, DEFAULT_SERVICE_TEXT),
        ("SERVICE__C", COL_SERVICE_C, service_c_n),
        ("Service", COL_SERVICE_SHORT, service_short),
        ("Origin Port", COL_ORIGIN_PORT, origin_port),
        ("Origin Postal Code", COL_ORIGIN_POSTAL, origin_postal),
        ("ORIGIN_COUNTRY__C", COL_ORIGIN_COUNTRY, origin_country or ""),
        ("Destination Port", COL_DEST_PORT, dest_port),
        ("Destination Postal Code", COL_DEST_POSTAL, dest_postal),
        ("DESTINATION_COUNTRY__C", COL_DEST_COUNTRY, dest_country or ""),
        ("Transporeon ID", COL_TRANSPOREON, transporeon),
    ]

    # Port fields: clean/expand to full alias group (e.g. CNTXG/CNTSN/CNXGA).
    # KEY / SERVICE__C / postals: never replace a richer RA value; only skip re-add.
    # Transporeon ID: never expand/merge — keep initially present value.
    port_fields = {"Origin Port", "Destination Port"}
    slash_fields = {
        "KEY",
        "SERVICE__C",
        "Origin Port",
        "Origin Postal Code",
        "Destination Port",
        "Destination Postal Code",
    }
    preserve_fields = {"Transporeon ID"}

    results: list[IdentitySyncResult] = []
    for field, col, want in wanted:
        cell = ws.cell(row, col)
        old = cell.value
        old_s = "" if old is None else str(old).strip()
        want = normalize_slash_value(want) if field in slash_fields else (want or "")

        if not want and not old_s:
            results.append(
                IdentitySyncResult(field, col, old, old_s, "unchanged")
            )
            continue

        # Keep original Transporeon ID; only fill when the cell is empty
        if field in preserve_fields:
            if old_s:
                results.append(
                    IdentitySyncResult(field, col, old, old_s, "unchanged")
                )
            elif want:
                cell.value = want
                cell.fill = GREEN
                results.append(IdentitySyncResult(field, col, old, want, "set"))
                _refresh_lane_attr(lane, col, want)
            else:
                results.append(
                    IdentitySyncResult(field, col, old, old_s, "unchanged")
                )
            continue

        # Update token(s) already in the cell → do not append again
        if want and _wanted_already_present(old, want):
            if field in slash_fields:
                compact = normalize_slash_value(old)
                # Prefer full port alias form when equivalent; else just de-space
                target = want if (
                    field in port_fields and _values_equivalent(old, want) and want
                ) else compact
                if target and target != old_s:
                    cell.value = target
                    cell.fill = GREEN
                    results.append(
                        IdentitySyncResult(field, col, old, target, "expanded")
                    )
                    _refresh_lane_attr(lane, col, target)
                    continue
            results.append(
                IdentitySyncResult(field, col, old, old_s or want, "unchanged")
            )
            continue

        if _values_equivalent(old, want):
            # Port alias family → normalize to full N-code form
            if field in port_fields and want and old_s != want:
                cell.value = want
                cell.fill = GREEN
                results.append(
                    IdentitySyncResult(field, col, old, want, "expanded")
                )
                _refresh_lane_attr(lane, col, want)
            elif field in slash_fields:
                compact = normalize_slash_value(old)
                if compact and compact != old_s:
                    cell.value = compact
                    cell.fill = GREEN
                    results.append(
                        IdentitySyncResult(field, col, old, compact, "expanded")
                    )
                    _refresh_lane_attr(lane, col, compact)
                else:
                    results.append(
                        IdentitySyncResult(field, col, old, old_s or want, "unchanged")
                    )
            else:
                results.append(
                    IdentitySyncResult(field, col, old, old_s or want, "unchanged")
                )
            continue

        if not old_s and want:
            cell.value = want
            cell.fill = GREEN
            results.append(IdentitySyncResult(field, col, old, want, "set"))
            _refresh_lane_attr(lane, col, want)
            continue

        merged = _merge_mismatch(old, want)
        if merged == old_s:
            results.append(
                IdentitySyncResult(field, col, old, old_s, "unchanged")
            )
            continue
        cell.value = merged
        cell.fill = ORANGE
        results.append(
            IdentitySyncResult(field, col, old, merged, "mismatch_merged")
        )
        _refresh_lane_attr(lane, col, merged)

    return results


def _refresh_lane_attr(lane, col: int, value: str) -> None:
    """Keep in-memory RaLane fields in sync after identity edits."""
    mapping = {
        COL_KEY: ("key_raw", True),
        COL_CARRIER: ("carrier", False),
        COL_SERVICE_C: ("service_c", False),
        COL_ORIGIN_PORT: ("origin_port", False),
        COL_ORIGIN_POSTAL: ("origin_postal", False),
        COL_ORIGIN_COUNTRY: ("origin_country", False),
        COL_DEST_PORT: ("dest_port", False),
        COL_DEST_POSTAL: ("dest_postal", False),
        COL_DEST_COUNTRY: ("dest_country", False),
    }
    spec = mapping.get(col)
    if not spec:
        return
    attr, is_key = spec
    setattr(lane, attr, value)
    if is_key:
        lane.keys = split_multi(value)


def locations_match(update_code: str, ra_port: str, ra_postal: str) -> bool:
    """True if update location matches RA port or postal (with aliases)."""
    if not (update_code or "").strip():
        return True
    port, postal = location_port_postal(update_code)
    update_parts = _norm_parts(port) | _norm_parts(postal)
    if not update_parts:
        update_parts = _norm_parts(update_code)
    ra_parts = _norm_parts(ra_port) | _norm_parts(ra_postal)
    if not ra_parts:
        return False
    # Match if any update token shares an alias group with any RA token,
    # or exact token overlap.
    if update_parts & ra_parts:
        return True
    for u in update_parts:
        ug = _PORT_TO_GROUP.get(u) or _POSTAL_TO_GROUP.get(u) or (u,)
        for r in ra_parts:
            rg = _PORT_TO_GROUP.get(r) or _POSTAL_TO_GROUP.get(r) or (r,)
            if set(x.upper() for x in ug) & set(x.upper() for x in rg):
                return True
    return False
