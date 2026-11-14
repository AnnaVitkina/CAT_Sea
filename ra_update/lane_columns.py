"""
Resolve Rate Card identity column indices from the field header row.

Sea exports include Origin/Destination Postal Code; Sea FR omits them, so
fixed column numbers shift country into the port slot. Always map by header name.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.worksheet.worksheet import Worksheet

from .config import (
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
    HEADER_FIELD_ROW,
)


def _norm_header(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().upper().replace(" ", "_")


# Canonical attr -> accepted header labels (normalized)
_HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "lane": ("LANE_#", "LANE"),
    "transporeon": ("TRANSPOREON_ID", "TRANSPOREON"),
    "key": ("KEY",),
    "carrier": ("CARRIER",),
    "service": ("SERVICE",),
    "service_c": ("SERVICE__C", "SERVICE_C"),
    "service_short": ("SERVICE",),  # short name also labeled "Service"; disambiguate below
    "valid_from": ("VALID_FROM",),
    "valid_to": ("VALID_TO",),
    "origin_port": ("ORIGIN_PORT",),
    "origin_postal": ("ORIGIN_POSTAL_CODE", "ORIGIN_POSTAL"),
    "origin_country": ("ORIGIN_COUNTRY__C", "ORIGIN_COUNTRY"),
    "dest_port": ("DESTINATION_PORT", "DEST_PORT"),
    "dest_postal": ("DESTINATION_POSTAL_CODE", "DESTINATION_POSTAL", "DEST_POSTAL"),
    "dest_country": ("DESTINATION_COUNTRY__C", "DESTINATION_COUNTRY", "DEST_COUNTRY"),
}


@dataclass(frozen=True)
class LaneColumns:
    """1-based Rate Card columns for lane identity fields. Postal may be None."""

    lane: int
    transporeon: int
    key: int
    carrier: int
    service: int
    service_c: int
    service_short: int
    valid_from: int
    valid_to: int
    origin_port: int
    origin_postal: int | None
    origin_country: int
    dest_port: int
    dest_postal: int | None
    dest_country: int


def _default_columns() -> LaneColumns:
    """Classic Sea layout (with postal columns)."""
    return LaneColumns(
        lane=COL_LANE,
        transporeon=COL_TRANSPOREON,
        key=COL_KEY,
        carrier=COL_CARRIER,
        service=COL_SERVICE,
        service_c=COL_SERVICE_C,
        service_short=COL_SERVICE_SHORT,
        valid_from=COL_VALID_FROM,
        valid_to=COL_VALID_TO,
        origin_port=COL_ORIGIN_PORT,
        origin_postal=COL_ORIGIN_POSTAL,
        origin_country=COL_ORIGIN_COUNTRY,
        dest_port=COL_DEST_PORT,
        dest_postal=COL_DEST_POSTAL,
        dest_country=COL_DEST_COUNTRY,
    )


def resolve_lane_columns(
    ws: Worksheet,
    header_row: int = HEADER_FIELD_ROW,
) -> LaneColumns:
    """
    Map identity fields from header labels. Falls back to classic Sea indices
    when a required header is missing.
    """
    by_norm: dict[str, list[int]] = {}
    for col in range(1, min(ws.max_column, 40) + 1):
        raw = ws.cell(header_row, col).value
        if raw is None or str(raw).strip() == "":
            continue
        # Cost block labels (Currency/Flat/p/unit) — stop once we hit them
        lab = str(raw).strip().lower()
        if lab in {"currency", "flat", "p/unit"}:
            break
        by_norm.setdefault(_norm_header(raw), []).append(col)

    defaults = _default_columns()

    def pick(attr: str, default: int | None, *, prefer_last: bool = False) -> int | None:
        aliases = _HEADER_ALIASES[attr]
        hits: list[int] = []
        for alias in aliases:
            hits.extend(by_norm.get(alias, []))
        if not hits:
            return default
        return hits[-1] if prefer_last else hits[0]

    # "Service" appears twice in spirit: SERVICE (long) vs Service (short).
    # Headers are distinct after normalize only if casing/spacing differs —
    # Sea uses SERVICE then SERVICE__C then Service → SERVICE, SERVICE__C, SERVICE.
    # Prefer column 5 for long SERVICE and column 7 for short when both map to SERVICE.
    service_hits = by_norm.get("SERVICE", [])
    if len(service_hits) >= 2:
        service_col = service_hits[0]
        service_short_col = service_hits[1]
    elif len(service_hits) == 1:
        # Only one "Service" label — keep defaults for the other
        service_col = service_hits[0]
        service_short_col = defaults.service_short
        if service_col == defaults.service_short:
            service_col = defaults.service
    else:
        service_col = defaults.service
        service_short_col = defaults.service_short

    origin_country = pick("origin_country", defaults.origin_country)
    dest_country = pick("dest_country", defaults.dest_country)
    origin_port = pick("origin_port", defaults.origin_port)
    dest_port = pick("dest_port", defaults.dest_port)
    if origin_country is None or dest_country is None or origin_port is None or dest_port is None:
        return defaults

    return LaneColumns(
        lane=pick("lane", defaults.lane) or defaults.lane,
        transporeon=pick("transporeon", defaults.transporeon) or defaults.transporeon,
        key=pick("key", defaults.key) or defaults.key,
        carrier=pick("carrier", defaults.carrier) or defaults.carrier,
        service=service_col,
        service_c=pick("service_c", defaults.service_c) or defaults.service_c,
        service_short=service_short_col,
        valid_from=pick("valid_from", defaults.valid_from) or defaults.valid_from,
        valid_to=pick("valid_to", defaults.valid_to) or defaults.valid_to,
        origin_port=origin_port,
        origin_postal=pick("origin_postal", None),
        origin_country=origin_country,
        dest_port=dest_port,
        dest_postal=pick("dest_postal", None),
        dest_country=dest_country,
    )
