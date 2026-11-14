"""Logical lane / cost identity keys (row and column order independent)."""

from __future__ import annotations

import re
from datetime import date

from ra_update.lane_identity import _all_tokens
from ra_update.ra_model import RaLane


def date_str(v: date | None) -> str:
    return v.isoformat() if v else ""


def norm_tokens(text: str) -> frozenset[str]:
    return frozenset(t.upper() for t in _all_tokens(text) if t)


def primary_service(service_c: str) -> str:
    toks = _all_tokens(service_c)
    return (toks[0] if toks else "").upper()


def lane_fingerprint(lane: RaLane) -> str:
    """Stable logical identity for a lane (order-independent KEY tokens)."""
    keys = sorted(norm_tokens(lane.key_raw))
    return "|".join(
        [
            ",".join(keys),
            (lane.carrier or "").upper(),
            primary_service(lane.service_c),
            date_str(lane.valid_from),
            date_str(lane.valid_to),
            ",".join(sorted(norm_tokens(lane.origin_port))),
            ",".join(sorted(norm_tokens(lane.dest_port))),
            (lane.origin_country or "").upper(),
            (lane.dest_country or "").upper(),
        ]
    )


def lane_soft_key(lane: RaLane) -> str:
    """Broader key for fuzzy pairing leftovers after exact fingerprint match."""
    return "|".join(
        [
            (lane.carrier or "").upper(),
            date_str(lane.valid_from),
            date_str(lane.valid_to),
            (lane.origin_country or "").upper(),
            (lane.dest_country or "").upper(),
        ]
    )


def cost_logical_key(cc) -> str:
    eq = (cc.equipment or "").upper()
    return "|".join(
        [
            re.sub(r"\s+", " ", (cc.fee_name or "").strip().upper()),
            eq,
            date_str(cc.valid_from),
            date_str(cc.valid_to),
        ]
    )


def token_set(text: str) -> frozenset[str]:
    return frozenset(t.upper() for t in _all_tokens(text or "") if t)
