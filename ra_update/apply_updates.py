from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any

import pandas as pd
from openpyxl.styles import PatternFill

from .charge_rules import service_out_of_scope_reason
from .config import CHANGED_FILL_HEX, DUAL_CONTAINER_FILL_HEX
from .cost_map import find_cost_columns
from .dates import (
    agreement_clamp_note,
    clamp_update_to_agreement,
    format_ra_date,
    parse_date,
    period_inside,
    periods_equal,
    uncovered_period_segments,
)
from .baf_create import BafCreateStore
from .equipment_decisions import (
    EquipmentDecisionStore,
    ensure_weight_suez_canal_charge,
    ensure_weight_war_risk_fee,
)
from .period_fees import (
    dft_periods_from_update_df,
    ensure_baf_cost_columns,
    ensure_ets_cost_columns,
)
from .lane_add import (
    choose_widest_period,
    create_lane_from_update,
    currency_for_rate,
    is_missing_lane_addable,
    periods_partially_overlap_without_nesting,
    resolve_nested_validity_cuts,
)
from .lane_identity import sync_lane_identity_from_update
from .lane_split import cost_update_leaks_outside_update, ensure_exact_validity_lane
from .matching import (
    find_lanes_for_key,
    lanes_overlapping_update,
    select_lane,
    service_without_bu,
    validate_lane_details_excluding_validity,
)
from .parsing import iter_update_rates, numbers_equal, to_number
from .ra_model import RateCardWorkbook


GREEN = PatternFill(start_color=CHANGED_FILL_HEX, end_color=CHANGED_FILL_HEX, fill_type="solid")
RED = PatternFill(
    start_color=DUAL_CONTAINER_FILL_HEX,
    end_color=DUAL_CONTAINER_FILL_HEX,
    fill_type="solid",
)


def _format_rate_display(value: object) -> str:
    num = to_number(value)
    if num is None:
        # Strip trailing validity bracket from prior dual text if present
        text = str(value).strip()
        return text
    if float(num).is_integer():
        return str(int(num))
    return str(num)


def _format_validity_span(valid_from: date | None, valid_to: date | None) -> str:
    return f"{format_ra_date(valid_from)}-{format_ra_date(valid_to)}"


def _is_zero_clear(value: object) -> bool:
    """Update rate/min 0 means clear the RA cell (not write numeric 0)."""
    num = to_number(value)
    return num is not None and abs(float(num)) < 1e-12


def _cell_is_empty(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def _format_rate_with_validity(
    value: object,
    valid_from: date | None,
    valid_to: date | None,
) -> str:
    """e.g. 548 (15.09.2026-14.10.2026)"""
    rate = _format_rate_display(value)
    if valid_from is None or valid_to is None:
        return rate
    return f"{rate} ({_format_validity_span(valid_from, valid_to)})"


def _is_dft_equipment(equipment: str | None) -> bool:
    return (equipment or "").strip().upper() == "DFT"


def _key_has_bu(key: str) -> bool:
    """True when KEY marks a BU service (e.g. OC_CNTR_CY-CY_BU-...)."""
    k = (key or "").upper()
    return "_BU-" in k or k.endswith("_BU")


def _key_has_spot(key: str) -> bool:
    """True when KEY marks a SPOT service (e.g. OC_RRBB_SPOT-... / OC_CNTR_SPOT-...)."""
    k = (key or "").upper()
    return (
        "_SPOT-" in k
        or "_SPOT/" in k
        or k.endswith("_SPOT")
        or "-SPOT-" in k
        or "/SPOT-" in k
    )


def _join_keys(keys: list[str]) -> str:
    seen: set[str] = set()
    out: list[str] = []
    for k in keys:
        if not k or k in seen:
            continue
        seen.add(k)
        out.append(k)
    return " | ".join(out)


def _periods_intersect(
    a_from: date | None,
    a_to: date | None,
    b_from: date | None,
    b_to: date | None,
) -> bool:
    if a_from is None or a_to is None or b_from is None or b_to is None:
        return True
    return a_from <= b_to and b_from <= a_to


@dataclass
class WrittenCellMeta:
    """Provenance for a cell written earlier in this run."""

    value: Any
    keys: list[str] = field(default_factory=list)
    valid_from: date | None = None
    valid_to: date | None = None
    equipment: str | None = None
    change_indexes: list[int] = field(default_factory=list)


@dataclass
class PendingLaneAdd:
    idx: Any
    row: pd.Series
    meta: dict
    rates: list[dict]
    earlier_clamped: bool  # True if start and/or end was cut to agreement
    original_from: date | None
    original_to: date | None
    clamp_note: str | None = None
    # BASE only: KEY exists but validity does not match → create a new lane period
    force_create: bool = False


def _cost_applies_on_lane(cost, lane) -> bool:
    """Whether a cost header period overlaps the lane validity (period-less always applies)."""
    if cost.valid_from is None or cost.valid_to is None:
        return True
    if lane.valid_from is None or lane.valid_to is None:
        return True
    return cost.valid_from <= lane.valid_to and cost.valid_to >= lane.valid_from


@dataclass
class IssueRow:
    sheet: str
    update_row_index: int
    key: str
    rate_col: str | None
    equipment: str | None
    charge_type: str | None
    new_value: Any
    lane_no: str | None
    ra_row: int | None
    reason: str


@dataclass
class ChangeRow:
    sheet: str
    update_row_index: int
    key: str
    lane_no: str
    ra_row: int
    rate_col: str
    cost_header: str
    punit_col: int
    old_value: Any
    new_value: Any
    status: str  # changed | already_equal


@dataclass
class EarlierValidityRow:
    sheet: str
    update_row_index: int
    key: str
    lane_no: str
    ra_row: int
    original_valid_from: str
    original_valid_to: str
    effective_valid_from: str
    effective_valid_to: str
    agreement_valid_from: str
    agreement_valid_to: str
    note: str


@dataclass
class UpdateStats:
    changes: list[ChangeRow] = field(default_factory=list)
    issues: list[IssueRow] = field(default_factory=list)
    earlier_validity: list[EarlierValidityRow] = field(default_factory=list)
    processed_update_rows: int = 0


class UpdateApplier:
    def __init__(
        self,
        ra: RateCardWorkbook,
        equipment_store: EquipmentDecisionStore | None = None,
        baf_create_store: BafCreateStore | None = None,
    ):
        self.ra = ra
        self.equipment_store = equipment_store or EquipmentDecisionStore()
        self.baf_create_store = baf_create_store or BafCreateStore()
        # cells already written in this run: (row, col) -> value
        self.written: dict[tuple[int, int], Any] = {}
        # equipment code that produced each written rate cell (for alias dual-values)
        self.written_equipment: dict[tuple[int, int], str] = {}
        # provenance for conflict rules (same value / BU prefer / intersecting validity)
        self.written_meta: dict[tuple[int, int], WrittenCellMeta] = {}
        # cells seeded from a wider-period template during lane add (may be overridden
        # by the period-specific update row on that same lane)
        self._seeded_cells: set[tuple[int, int]] = set()
        self.stats = UpdateStats()
        # DFT BAF/ETS validity catalogs (from all update effectives/expiries)
        self._dft_baf_periods: list[tuple[date, date]] | None = None
        self._dft_ets_periods: list[tuple[date, date]] | None = None

    def _find_costs(
        self,
        *,
        ra_fee_name,
        equipment,
        valid_from,
        valid_to,
        require_period,
        sheet_name: str | None = None,
        idx=None,
        key: str | None = None,
        rate_col: str | None = None,
    ):
        aliases = self.equipment_store.aliases or None
        # Resolve runtime alias for period ensure/create
        eq = equipment
        if eq and aliases and eq.upper() in aliases:
            eq = aliases[eq.upper()]
        eq_u = str(eq).upper() if eq else ""

        # BAF: container = monthly 15→14; DFT = contiguous periods from all *DFTBAF rows
        if ra_fee_name == "BAF Fee" and require_period and valid_from and valid_to and eq:
            cols, note = ensure_baf_cost_columns(
                self.ra,
                equipment=eq_u,
                update_from=valid_from,
                update_to=valid_to,
                baf_create_store=self.baf_create_store,
                dft_periods=self._dft_baf_periods if eq_u == "DFT" else None,
            )
            if note and sheet_name is not None:
                self._issue(
                    sheet=sheet_name,
                    update_row_index=idx,
                    key=key or "",
                    rate_col=rate_col,
                    equipment=equipment,
                    charge_type="BAF",
                    new_value=None,
                    lane_no=None,
                    ra_row=None,
                    reason=note,
                )
            if cols:
                return cols, None
            return [], note or "BAF Fee cost columns could not be ensured"

        # EU ETS: prefer covering column; else create.
        # DFT uses contiguous periods from all *DFTEU_ETS rows (not exact-row create).
        if ra_fee_name == "EU ETS Fee" and require_period and valid_from and valid_to and eq:
            if eq_u == "DFT" and self._dft_ets_periods is not None:
                created, cerr = ensure_ets_cost_columns(
                    self.ra,
                    equipment=eq_u,
                    update_from=valid_from,
                    update_to=valid_to,
                    agreement_from=self.ra.agreement_valid_from,
                    agreement_to=self.ra.agreement_valid_to,
                    dft_periods=self._dft_ets_periods,
                )
                if created:
                    return created, None
                return [], cerr or "DFT EU ETS Fee cost columns could not be ensured"

            cols, err = find_cost_columns(
                self.ra.cost_columns,
                ra_fee_name=ra_fee_name,
                equipment=equipment,
                valid_from=valid_from,
                valid_to=valid_to,
                require_period=require_period,
                runtime_aliases=aliases,
            )
            if cols:
                return cols, None
            created, cerr = ensure_ets_cost_columns(
                self.ra,
                equipment=eq_u,
                update_from=valid_from,
                update_to=valid_to,
                agreement_from=self.ra.agreement_valid_from,
                agreement_to=self.ra.agreement_valid_to,
            )
            if created:
                return created, None
            return [], cerr or err or "EU ETS Fee cost column not found"

        # *DFTWM_WAR_RISK → fee-only War Risk Fee (Weight/chargeable kg); create if missing
        if ra_fee_name == "War Risk Fee" and not eq_u and not require_period:
            cols, err = find_cost_columns(
                self.ra.cost_columns,
                ra_fee_name=ra_fee_name,
                equipment=None,
                valid_from=valid_from,
                valid_to=valid_to,
                require_period=False,
                runtime_aliases=aliases,
            )
            if cols:
                return cols, None
            created, cerr = ensure_weight_war_risk_fee(self.ra)
            if created:
                return [created], None
            return [], cerr or err or "War Risk Fee (weight) cost column not found"

        # *DFTWM_SCANAL → fee-only Suez Canal Charge (Weight/chargeable kg); create if missing
        if ra_fee_name == "Suez Canal Charge" and not eq_u and not require_period:
            cols, err = find_cost_columns(
                self.ra.cost_columns,
                ra_fee_name=ra_fee_name,
                equipment=None,
                valid_from=valid_from,
                valid_to=valid_to,
                require_period=False,
                runtime_aliases=aliases,
            )
            if cols:
                return cols, None
            # Legacy RA header "Suez Canal Charge (SCANAL)"
            legacy = [
                c
                for c in self.ra.cost_columns
                if c.fee_name == "Suez Canal Charge"
                and (c.equipment or "").upper() == "SCANAL"
            ]
            if legacy:
                return legacy, None
            created, cerr = ensure_weight_suez_canal_charge(self.ra)
            if created:
                return [created], None
            return [], cerr or err or "Suez Canal Charge (weight) cost column not found"

        return find_cost_columns(
            self.ra.cost_columns,
            ra_fee_name=ra_fee_name,
            equipment=equipment,
            valid_from=valid_from,
            valid_to=valid_to,
            require_period=require_period,
            runtime_aliases=aliases,
        )

    def process_sheet(self, sheet_name: str, df: pd.DataFrame) -> None:
        mode = sheet_name.upper()
        if mode not in {"BASE", "ETSBAF"}:
            raise ValueError(f"Unsupported update sheet {sheet_name}")

        pending_adds: dict[str, list[PendingLaneAdd]] | None = (
            {} if mode == "BASE" else None
        )

        # DFT BAF/ETS: derive cost validities from all update effectives/expiries
        # and create catalog columns up front (do not wait for a successful lane match).
        if mode == "ETSBAF":
            self._dft_baf_periods = dft_periods_from_update_df(
                df,
                charge_keys={"DFTBAF"},
                agreement_from=self.ra.agreement_valid_from,
                agreement_to=self.ra.agreement_valid_to,
            )
            self._dft_ets_periods = dft_periods_from_update_df(
                df,
                charge_keys={"DFTEU_ETS"},
                agreement_from=self.ra.agreement_valid_from,
                agreement_to=self.ra.agreement_valid_to,
            )
            # Prefetch: reuse covering DFT columns; never invent sliced columns
            # when a potential overlapping column already exists (that is an issue).
            if self._dft_baf_periods:
                txt = ", ".join(
                    f"{format_ra_date(a)}-{format_ra_date(b)}"
                    for a, b in self._dft_baf_periods
                )
                print(f"  DFT BAF periods: {txt}")
                u_from, u_to = self._dft_baf_periods[0][0], self._dft_baf_periods[-1][1]
                cols, note = ensure_baf_cost_columns(
                    self.ra,
                    equipment="DFT",
                    update_from=u_from,
                    update_to=u_to,
                    baf_create_store=None,
                    dft_periods=self._dft_baf_periods,
                )
                print(f"  DFT BAF columns ready: {len(cols)}")
                if note:
                    print(f"  DFT BAF: {note}")
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=-1,
                        key="",
                        rate_col=None,
                        equipment="DFT",
                        charge_type="DFTBAF",
                        new_value=None,
                        lane_no=None,
                        ra_row=None,
                        reason=note,
                    )
            if self._dft_ets_periods:
                txt = ", ".join(
                    f"{format_ra_date(a)}-{format_ra_date(b)}"
                    for a, b in self._dft_ets_periods
                )
                print(f"  DFT EU ETS periods: {txt}")
                u_from, u_to = self._dft_ets_periods[0][0], self._dft_ets_periods[-1][1]
                cols, note = ensure_ets_cost_columns(
                    self.ra,
                    equipment="DFT",
                    update_from=u_from,
                    update_to=u_to,
                    agreement_from=self.ra.agreement_valid_from,
                    agreement_to=self.ra.agreement_valid_to,
                    dft_periods=self._dft_ets_periods,
                )
                print(f"  DFT EU ETS columns ready: {len(cols)}")
                if note:
                    print(f"  DFT EU ETS: {note}")
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=-1,
                        key="",
                        rate_col=None,
                        equipment="DFT",
                        charge_type="DFTEU_ETS",
                        new_value=None,
                        lane_no=None,
                        ra_row=None,
                        reason=note,
                    )
        else:
            self._dft_baf_periods = None
            self._dft_ets_periods = None

        for idx, row in df.iterrows():
            self.stats.processed_update_rows += 1
            self._process_update_row(sheet_name, mode, idx, row, pending_adds)

        if pending_adds:
            self._flush_pending_lane_adds(sheet_name, pending_adds)

    def _meta(self, row: pd.Series) -> dict:
        def s(col: str) -> str:
            v = row.get(col)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return ""
            return str(v).strip()

        # Port/postal source: LOCATION first (zone only as fallback)
        dest = s("DESTINATION_LOCATION_NAME__C") or s("DESTINATION_ZONE__C")
        origin = s("ORIGIN_LOCATION_NAME__C") or s("ORIGIN_ZONE__C")
        return {
            "key": s("KEY"),
            "carrier": s("CARRIER"),
            "service_c": s("SERVICE__C"),
            "origin_country": s("ORIGIN_COUNTRY__C"),
            "dest_country": s("DESTINATION_COUNTRY__C"),
            "origin_code": origin,
            "dest_code": dest,
            "valid_from": parse_date(row.get("RATE_EFFECTIVE_DATE__C")),
            "valid_to": parse_date(row.get("RATE_EXPIRATION_DATE__C")),
        }

    def _issue(self, **kwargs) -> None:
        self.stats.issues.append(IssueRow(**kwargs))

    def _process_update_row(
        self,
        sheet_name: str,
        mode: str,
        idx,
        row: pd.Series,
        pending_adds: dict[str, list[PendingLaneAdd]] | None = None,
    ) -> None:
        meta = self._meta(row)
        key = meta["key"]
        original_from = meta["valid_from"]
        original_to = meta["valid_to"]

        (
            effective_from,
            effective_to,
            clamped_start,
            clamped_end,
            clamp_err,
        ) = clamp_update_to_agreement(
            original_from,
            original_to,
            self.ra.agreement_valid_from,
            self.ra.agreement_valid_to,
        )
        meta["valid_from"] = effective_from
        meta["valid_to"] = effective_to
        earlier_clamped = clamped_start or clamped_end
        clamp_note = agreement_clamp_note(
            clamped_start=clamped_start, clamped_end=clamped_end
        )

        if clamp_err:
            self._issue(
                sheet=sheet_name,
                update_row_index=idx,
                key=key,
                rate_col=None,
                equipment=None,
                charge_type=None,
                new_value=None,
                lane_no=None,
                ra_row=None,
                reason=clamp_err,
            )
            return

        scope_reason = service_out_of_scope_reason(meta["service_c"])
        if scope_reason:
            self._issue(
                sheet=sheet_name,
                update_row_index=idx,
                key=key,
                rate_col=None,
                equipment=None,
                charge_type=None,
                new_value=None,
                lane_no=None,
                ra_row=None,
                reason=scope_reason,
            )
            return

        rates = iter_update_rates(row, sheet_name)

        if not key:
            self._issue(
                sheet=sheet_name,
                update_row_index=int(idx) if str(idx).isdigit() else idx,
                key="",
                rate_col=None,
                equipment=None,
                charge_type=None,
                new_value=None,
                lane_no=None,
                ra_row=None,
                reason="Update row has empty KEY",
            )
            return

        if not rates:
            self._issue(
                sheet=sheet_name,
                update_row_index=idx,
                key=key,
                rate_col=None,
                equipment=None,
                charge_type=None,
                new_value=None,
                lane_no=None,
                ra_row=None,
                reason="No non-empty *_RATE values on update row",
            )
            return

        # Out-of-scope / unsupported rates do not require a lane match by themselves,
        # but if every rate is unsupported we still record them without lane lookup.
        actionable = [r for r in rates if not r["unsupported"]]
        if not actionable:
            for r in rates:
                self._issue(
                    sheet=sheet_name,
                    update_row_index=idx,
                    key=key,
                    rate_col=r["rate_col"],
                    equipment=r["equipment"],
                    charge_type=r["charge_type"],
                    new_value=r["new_value"],
                    lane_no=None,
                    ra_row=None,
                    reason=r["reason"] or "Not covered by current rules",
                )
            return

        candidates, key_resolve_note = find_lanes_for_key(self.ra, key)
        service_for_match = meta["service_c"]
        # BU KEY resolved onto non-BU / slash-paired lanes: match service without _BU
        if key_resolve_note and _key_has_bu(key):
            service_for_match = service_without_bu(meta["service_c"])
        match = select_lane(
            candidates,
            carrier=meta["carrier"],
            service_c=service_for_match,
            origin_country=meta["origin_country"],
            dest_country=meta["dest_country"],
            origin_code=meta["origin_code"],
            dest_code=meta["dest_code"],
            update_from=meta["valid_from"],
            update_to=meta["valid_to"],
            mode=mode,
        )
        if match.ok and key_resolve_note:
            self._issue(
                sheet=sheet_name,
                update_row_index=idx,
                key=key,
                rate_col=None,
                equipment=None,
                charge_type=None,
                new_value=None,
                lane_no=",".join(l.lane_no for l in match.lanes[:5]) or None,
                ra_row=match.lanes[0].row if match.lanes else None,
                reason=key_resolve_note,
            )

        ets_wider_than_lane = False

        if not match.ok:
            # Always surface explicit out-of-scope charges separately
            remaining = []
            for r in rates:
                reason = (r.get("reason") or "").lower()
                if r["unsupported"] and ("out of scope" in reason or "not in scope" in reason):
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=idx,
                        key=key,
                        rate_col=r["rate_col"],
                        equipment=r["equipment"],
                        charge_type=r["charge_type"],
                        new_value=r["new_value"],
                        lane_no=None,
                        ra_row=None,
                        reason=r["reason"] or "Out of scope",
                    )
                else:
                    remaining.append(r)

            if remaining:
                reason = match.reason or "Lane match failed"
                # BASE: defer missing KEY or unmatched validity so we can auto-add
                # a new lane for the (agreement-clamped) update period.
                base_defer_create = pending_adds is not None and (
                    reason.startswith("No such lane")
                    or reason.startswith("BASE validity not identical")
                )
                if base_defer_create:
                    pending_adds.setdefault(key, []).append(
                        PendingLaneAdd(
                            idx=idx,
                            row=row,
                            meta=dict(meta),
                            rates=list(rates),
                            earlier_clamped=earlier_clamped,
                            original_from=original_from,
                            original_to=original_to,
                            clamp_note=clamp_note,
                            force_create=reason.startswith(
                                "BASE validity not identical"
                            ),
                        )
                    )
                    return

                # ETSBAF: KEY absent from RA → note only (not an issue)
                if mode == "ETSBAF" and "KEY not found" in reason:
                    rate_cols = ", ".join(r["rate_col"] for r in remaining)
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=idx,
                        key=key,
                        rate_col=rate_cols,
                        equipment=None,
                        charge_type=None,
                        new_value=None,
                        lane_no=None,
                        ra_row=None,
                        reason="KEY not found",
                    )
                    return

                # ETSBAF: update may be wider than lane / no contiguous cover.
                # Apply container ETS + DFT BAF/ETS on overlapping lanes; other charges fail.
                overlap_ok_charges = {"ETS", "DFTBAF", "DFTEU_ETS"}
                ets_rates = [
                    r
                    for r in remaining
                    if not r.get("unsupported")
                    and (r.get("charge_type") or "").upper() in overlap_ok_charges
                ]
                other_rates = [r for r in remaining if r not in ets_rates]

                if mode == "ETSBAF" and ets_rates and meta["valid_from"] and meta["valid_to"]:
                    detail_ok = []
                    for lane in candidates:
                        ok, _ = validate_lane_details_excluding_validity(
                            lane,
                            carrier=meta["carrier"],
                            service_c=service_for_match,
                            origin_country=meta["origin_country"],
                            dest_country=meta["dest_country"],
                            origin_code=meta["origin_code"],
                            dest_code=meta["dest_code"],
                        )
                        if ok:
                            detail_ok.append(lane)
                    overlap = lanes_overlapping_update(
                        meta["valid_from"], meta["valid_to"], detail_ok
                    )
                    if overlap:
                        match = type(match)(overlap, True, reason, ets_wider_than_lane=True)
                        ets_wider_than_lane = True
                        rates = ets_rates + [
                            r for r in rates if r.get("unsupported")
                        ]
                        # Non-ETS/non-DFT actionable rates still reported as failures
                        if other_rates:
                            rate_cols = ", ".join(r["rate_col"] for r in other_rates)
                            self._issue(
                                sheet=sheet_name,
                                update_row_index=idx,
                                key=key,
                                rate_col=rate_cols,
                                equipment=None,
                                charge_type=None,
                                new_value=None,
                                lane_no=",".join(l.lane_no for l in overlap[:3]),
                                ra_row=overlap[0].row,
                                reason=reason,
                            )
                    else:
                        rate_cols = ", ".join(r["rate_col"] for r in remaining)
                        self._issue(
                            sheet=sheet_name,
                            update_row_index=idx,
                            key=key,
                            rate_col=rate_cols,
                            equipment=None,
                            charge_type=None,
                            new_value=None,
                            lane_no=None,
                            ra_row=None,
                            reason=reason,
                        )
                        return
                else:
                    rate_cols = ", ".join(r["rate_col"] for r in remaining)
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=idx,
                        key=key,
                        rate_col=rate_cols,
                        equipment=None,
                        charge_type=None,
                        new_value=None,
                        lane_no=None,
                        ra_row=None,
                        reason=reason,
                    )
                    return
            else:
                return

        matched_lanes = list(match.lanes)
        ets_wider_than_lane = ets_wider_than_lane or bool(
            getattr(match, "ets_wider_than_lane", False)
        )
        if not matched_lanes:
            self._issue(
                sheet=sheet_name,
                update_row_index=idx,
                key=key,
                rate_col=", ".join(r["rate_col"] for r in actionable),
                equipment=None,
                charge_type=None,
                new_value=None,
                lane_no=None,
                ra_row=None,
                reason="Lane match returned no lanes",
            )
            return

        # Broad identity sync: SERVICE short, ports/postals (+ aliases), Transporeon.
        # Mismatches keep RA value and append update with / (orange).
        for lane in matched_lanes:
            syncs = sync_lane_identity_from_update(
                self.ra,
                lane,
                key=key,
                carrier=meta["carrier"],
                service_c=meta["service_c"],
                origin_country=meta["origin_country"],
                dest_country=meta["dest_country"],
                origin_code=meta["origin_code"],
                dest_code=meta["dest_code"],
            )
            for syn in syncs:
                if syn.status == "unchanged":
                    continue
                self.stats.changes.append(
                    ChangeRow(
                        sheet=sheet_name,
                        update_row_index=idx,
                        key=key,
                        lane_no=lane.lane_no,
                        ra_row=lane.row,
                        rate_col="(identity)",
                        cost_header=syn.field,
                        punit_col=syn.col,
                        old_value=syn.old_value,
                        new_value=syn.new_value,
                        status=syn.status,
                    )
                )
                if syn.status == "mismatch_merged":
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=idx,
                        key=key,
                        rate_col=syn.field,
                        equipment=None,
                        charge_type=None,
                        new_value=syn.new_value,
                        lane_no=lane.lane_no,
                        ra_row=lane.row,
                        reason=(
                            f"Identity mismatch on {syn.field}: "
                            f"RA={syn.old_value!r} update appended → {syn.new_value!r}"
                        ),
                    )

        # Track write targets per original matched lane (may become split result)
        write_targets = {id(lane): lane for lane in matched_lanes}
        ref_lane = matched_lanes[0]

        for r in rates:
            if r["unsupported"]:
                self._issue(
                    sheet=sheet_name,
                    update_row_index=idx,
                    key=key,
                    rate_col=r["rate_col"],
                    equipment=r["equipment"],
                    charge_type=r["charge_type"],
                    new_value=r["new_value"],
                    lane_no=ref_lane.lane_no,
                    ra_row=ref_lane.row,
                    reason=r["reason"] or "Not covered by current rules",
                )
                continue

            rule = r["rule"]
            costs, err = self._find_costs(
                ra_fee_name=rule.ra_fee_name,
                equipment=r["equipment"],
                valid_from=meta["valid_from"],
                valid_to=meta["valid_to"],
                require_period=rule.require_period,
                sheet_name=sheet_name,
                idx=idx,
                key=key,
                rate_col=r["rate_col"],
            )
            if not costs:
                self._issue(
                    sheet=sheet_name,
                    update_row_index=idx,
                    key=key,
                    rate_col=r["rate_col"],
                    equipment=r["equipment"],
                    charge_type=r["charge_type"],
                    new_value=r["new_value"],
                    lane_no=ref_lane.lane_no,
                    ra_row=ref_lane.row,
                    reason=err or "Cost column not found",
                )
                continue

            for lane in matched_lanes:
                write_lane = write_targets[id(lane)]
                u_from, u_to = meta["valid_from"], meta["valid_to"]

                # Wider lane contains the update -> cut/duplicate to exact validity
                # before writing (BASE always; ETSBAF when cost update would leak).
                lane_wider = (
                    u_from is not None
                    and u_to is not None
                    and write_lane.valid_from is not None
                    and write_lane.valid_to is not None
                    and period_inside(u_from, u_to, write_lane.valid_from, write_lane.valid_to)
                    and not periods_equal(
                        u_from, u_to, write_lane.valid_from, write_lane.valid_to
                    )
                )
                needs_split = False
                # ETS-wider-than-lane: do not cut/expand the lane to the update period
                if lane_wider and not ets_wider_than_lane:
                    if mode == "BASE":
                        # BASE costs are period-less on the lane -> must isolate update segment
                        needs_split = True
                    elif mode == "ETSBAF" and cost_update_leaks_outside_update(
                        write_lane, costs, u_from, u_to
                    ):
                        needs_split = True

                if needs_split:
                    try:
                        write_lane, split_actions = ensure_exact_validity_lane(
                            self.ra,
                            key=key,
                            lane=write_lane,
                            update_from=u_from,
                            update_to=u_to,
                        )
                    except ValueError as exc:
                        self._issue(
                            sheet=sheet_name,
                            update_row_index=idx,
                            key=key,
                            rate_col=r["rate_col"],
                            equipment=r["equipment"],
                            charge_type=r["charge_type"],
                            new_value=r["new_value"],
                            lane_no=lane.lane_no,
                            ra_row=lane.row,
                            reason=f"Lane split failed: {exc}",
                        )
                        continue

                    write_targets[id(lane)] = write_lane
                    for action in split_actions:
                        self.stats.changes.append(
                            ChangeRow(
                                sheet=sheet_name,
                                update_row_index=idx,
                                key=key,
                                lane_no=action.result_lane_no,
                                ra_row=action.result_row,
                                rate_col="(lane split)",
                                cost_header=action.detail,
                                punit_col=0,
                                old_value=f"{action.source_lane_no}@{action.source_row}",
                                new_value=(
                                    f"{format_ra_date(action.valid_from)}-"
                                    f"{format_ra_date(action.valid_to)}"
                                ),
                                status=action.action,
                            )
                        )

                currency = currency_for_rate(row, r["rate_col"])

                for cost in costs:
                    if not _cost_applies_on_lane(cost, write_lane):
                        continue

                    # Currency from update *_CURRENCY companion (esp. newly created costs)
                    self._write_currency(
                        lane=write_lane,
                        cost=cost,
                        currency=currency,
                    )

                    # Main rate -> p/unit, or Flat when Flat is the rate (e.g. Documentation Fee)
                    rate_target = cost.rate_col
                    if rate_target is not None:
                        self._write_cell(
                            sheet_name=sheet_name,
                            idx=idx,
                            key=key,
                            lane=write_lane,
                            rate_col=r["rate_col"],
                            cost_header=f"{cost.header} [{'p/unit' if cost.punit_col == rate_target else 'Flat'}]",
                            value_col=rate_target,
                            new_value=r["new_value"],
                            equipment=r.get("equipment"),
                            valid_from=meta.get("valid_from"),
                            valid_to=meta.get("valid_to"),
                        )
                    else:
                        self._issue(
                            sheet=sheet_name,
                            update_row_index=idx,
                            key=key,
                            rate_col=r["rate_col"],
                            equipment=r["equipment"],
                            charge_type=r["charge_type"],
                            new_value=r["new_value"],
                            lane_no=write_lane.lane_no,
                            ra_row=write_lane.row,
                            reason=f"No rate column (p/unit or Flat) for {cost.header!r}",
                        )

                    # MIN -> Flat when Flat is labeled MIN (BASE_STAT_FRK, RoRo Fee, ...)
                    min_val = r.get("min_value")
                    min_target = cost.min_col
                    if min_val is not None and min_target is not None:
                        self._write_cell(
                            sheet_name=sheet_name,
                            idx=idx,
                            key=key,
                            lane=write_lane,
                            rate_col=r.get("min_col") or f"{r['rate_col']}[MIN]",
                            cost_header=f"{cost.header} [Flat/MIN]",
                            value_col=min_target,
                            new_value=min_val,
                            equipment=r.get("equipment"),
                            valid_from=meta.get("valid_from"),
                            valid_to=meta.get("valid_to"),
                        )
                    elif min_val is not None and min_target is None:
                        self._issue(
                            sheet=sheet_name,
                            update_row_index=idx,
                            key=key,
                            rate_col=r.get("min_col"),
                            equipment=r["equipment"],
                            charge_type=r["charge_type"],
                            new_value=min_val,
                            lane_no=write_lane.lane_no,
                            ra_row=write_lane.row,
                            reason=(
                                f"Update has MIN={min_val} but RA cost {cost.header!r} "
                                "has no Flat/MIN column"
                            ),
                        )

        if ets_wider_than_lane:
            u_from, u_to = meta["valid_from"], meta["valid_to"]
            seen_rows: set[int] = set()
            for lane in matched_lanes:
                target = write_targets[id(lane)]
                if target.row in seen_rows:
                    continue
                seen_rows.add(target.row)
                self._issue(
                    sheet=sheet_name,
                    update_row_index=idx,
                    key=key,
                    rate_col=None,
                    equipment=None,
                    charge_type="ETS",
                    new_value=None,
                    lane_no=target.lane_no,
                    ra_row=target.row,
                    reason=(
                        "ETSBAF OK: EU ETS cost/update validity is wider than lane validity; "
                        "ETS rate applied on overlapping lane without expanding lane dates. "
                        f"update={format_ra_date(u_from)}-{format_ra_date(u_to)}, "
                        f"lane={format_ra_date(target.valid_from)}-"
                        f"{format_ra_date(target.valid_to)}"
                    ),
                )

        if earlier_clamped and clamp_note:
            agr_from = self.ra.agreement_valid_from
            agr_to = self.ra.agreement_valid_to
            seen_rows: set[int] = set()
            for lane in matched_lanes:
                target = write_targets[id(lane)]
                if target.row in seen_rows:
                    continue
                seen_rows.add(target.row)
                self.stats.earlier_validity.append(
                    EarlierValidityRow(
                        sheet=sheet_name,
                        update_row_index=idx,
                        key=key,
                        lane_no=target.lane_no,
                        ra_row=target.row,
                        original_valid_from=format_ra_date(original_from) if original_from else "",
                        original_valid_to=format_ra_date(original_to) if original_to else "",
                        effective_valid_from=format_ra_date(effective_from) if effective_from else "",
                        effective_valid_to=format_ra_date(effective_to) if effective_to else "",
                        agreement_valid_from=format_ra_date(agr_from) if agr_from else "",
                        agreement_valid_to=format_ra_date(agr_to) if agr_to else "",
                        note=clamp_note,
                    )
                )

    def _flush_pending_lane_adds(
        self,
        sheet_name: str,
        pending_adds: dict[str, list[PendingLaneAdd]],
    ) -> None:
        # Non-BU keys first so a later BU KEY can attach to that lane
        # (slash-append KEY/SERVICE) instead of creating a second lane.
        ordered_keys = sorted(
            pending_adds.keys(),
            key=lambda k: (1 if _key_has_bu(k) else 0, k),
        )
        for key in ordered_keys:
            items = pending_adds.get(key) or []
            if not items:
                continue

            force_create = any(i.force_create for i in items)

            # Before creating a BU lane: if non-BU (or slash) sibling now exists
            # (including just created above), treat BU rows as updates of that lane.
            # Skip when force_create: we need a new validity segment, not a re-match.
            if _key_has_bu(key) and not force_create:
                cands, _note = find_lanes_for_key(self.ra, key)
                if cands:
                    for item in items:
                        self._process_update_row(
                            sheet_name,
                            "BASE",
                            item.idx,
                            item.row,
                            pending_adds=None,
                        )
                    continue

            sample = items[0]
            # Validity-segment create: KEY already on RA → allow new period lane
            if force_create and find_lanes_for_key(self.ra, key)[0]:
                addable, why_not = True, None
            else:
                addable, why_not = is_missing_lane_addable(
                    self.ra.agreement_title,
                    sample.meta.get("origin_country") or "",
                    sample.meta.get("dest_country") or "",
                )
            if not addable:
                for item in items:
                    remaining = [r for r in item.rates if not r["unsupported"]]
                    rate_cols = ", ".join(r["rate_col"] for r in remaining) if remaining else None
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=item.idx,
                        key=key,
                        rate_col=rate_cols,
                        equipment=None,
                        charge_type=None,
                        new_value=None,
                        lane_no=None,
                        ra_row=None,
                        reason=why_not or "No such lane",
                    )
                continue

            try:
                self._add_missing_lanes_for_key(sheet_name, key, items)
            except ValueError as exc:
                for item in items:
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=item.idx,
                        key=key,
                        rate_col=None,
                        equipment=None,
                        charge_type=None,
                        new_value=None,
                        lane_no=None,
                        ra_row=None,
                        reason=f"Lane add failed: {exc}",
                    )

    def _add_missing_lanes_for_key(
        self,
        sheet_name: str,
        key: str,
        items: list[PendingLaneAdd],
    ) -> None:
        """
        Create missing BASE lanes for one KEY.

        Multi-period (nested) example KHNN:
          wide  01.06.2026-31.05.2027  (DTHC/OTHC)
          narrow 15.09.2026-31.05.2027 (BASE)
          1) create both lanes with costs from the wider update row
          2) cut wide validity to day before narrow start
          3) apply each update row's own costs onto its validity lane

        BU KEY: never create a second lane when the non-BU sibling is already
        present — re-apply as an update (slash KEY/SERVICE + cost conflict rules),
        unless force_create (BASE validity segment missing → new lane).
        """
        force_create = any(i.force_create for i in items)
        if _key_has_bu(key) and not force_create:
            cands, _note = find_lanes_for_key(self.ra, key)
            if cands:
                for item in items:
                    self._process_update_row(
                        sheet_name,
                        "BASE",
                        item.idx,
                        item.row,
                        pending_adds=None,
                    )
                return

        for item in items:
            vf, vt = item.meta.get("valid_from"), item.meta.get("valid_to")
            if vf is None or vt is None:
                raise ValueError("Update row missing effective validity dates")

        # BASE validity miss with KEY already on RA: create only the uncovered
        # remainder (e.g. existing 01.06-31.10 + update 01.06-31.05.2027 →
        # new lane 01.11-31.05.2027), then re-apply update on all covering lanes.
        if force_create:
            self._force_create_validity_remainders(sheet_name, key, items)
            return

        period_items: dict[tuple[date, date], list[PendingLaneAdd]] = {}
        for item in items:
            period = (item.meta["valid_from"], item.meta["valid_to"])
            period_items.setdefault(period, []).append(item)

        periods = list(period_items.keys())
        conflict = periods_partially_overlap_without_nesting(periods)
        if conflict is not None:
            raise ValueError(
                f"Update periods overlap without nesting "
                f"({format_ra_date(conflict[0])}-{format_ra_date(conflict[1])}); "
                "cannot auto-add"
            )

        widest_period = choose_widest_period(periods)
        widest_item = period_items[widest_period][0]
        w_meta = widest_item.meta

        lanes_by_period: dict[tuple[date, date], Any] = {}
        seeded_for_key: set[tuple[int, int]] = set()

        # 1) Create one lane per distinct update validity; seed costs from widest row
        for period in sorted(periods, key=lambda p: (p[0], p[1])):
            vf, vt = period
            lane, action = create_lane_from_update(
                self.ra,
                key=key,
                carrier=w_meta.get("carrier") or "",
                service_c=w_meta.get("service_c") or "",
                origin_country=w_meta.get("origin_country") or "",
                dest_country=w_meta.get("dest_country") or "",
                origin_code=w_meta.get("origin_code") or "",
                dest_code=w_meta.get("dest_code") or "",
                valid_from=vf,
                valid_to=vt,
            )
            lanes_by_period[period] = lane
            self.stats.changes.append(
                ChangeRow(
                    sheet=sheet_name,
                    update_row_index=widest_item.idx,
                    key=key,
                    lane_no=action.lane_no,
                    ra_row=action.ra_row,
                    rate_col="(lane add)",
                    cost_header=action.detail,
                    punit_col=0,
                    old_value=None,
                    new_value=f"{format_ra_date(vf)}-{format_ra_date(vt)}",
                    status=action.action,
                )
            )

            before_written = set(self.written.keys())
            self._apply_rates_to_lane(
                sheet_name=sheet_name,
                idx=widest_item.idx,
                key=key,
                row=widest_item.row,
                lane=lane,
                rates=widest_item.rates,
                valid_from=widest_period[0],
                valid_to=widest_period[1],
                allow_seed_overwrite=False,
            )
            newly = set(self.written.keys()) - before_written
            seeded_for_key |= newly
            self._seeded_cells |= newly

        # 2) Cut wider nested lanes (preserve earlier/older validity logic)
        cut_actions = resolve_nested_validity_cuts(
            self.ra, key=key, lanes_by_period=lanes_by_period
        )
        for action in cut_actions:
            self.stats.changes.append(
                ChangeRow(
                    sheet=sheet_name,
                    update_row_index=widest_item.idx,
                    key=key,
                    lane_no=action.lane_no,
                    ra_row=action.ra_row,
                    rate_col="(lane add)",
                    cost_header=action.detail,
                    punit_col=0,
                    old_value=None,
                    new_value=(
                        f"{format_ra_date(action.valid_from)}-"
                        f"{format_ra_date(action.valid_to)}"
                    ),
                    status=action.action,
                )
            )

        # 3) Apply each update row's own costs onto its validity lane
        #    (may override wider-template seed on that lane)
        for item in items:
            period = (item.meta["valid_from"], item.meta["valid_to"])
            lane = lanes_by_period[period]
            self._apply_rates_to_lane(
                sheet_name=sheet_name,
                idx=item.idx,
                key=key,
                row=item.row,
                lane=lane,
                rates=item.rates,
                valid_from=period[0],
                valid_to=period[1],
                allow_seed_overwrite=True,
            )

            if item.earlier_clamped:
                agr_from = self.ra.agreement_valid_from
                agr_to = self.ra.agreement_valid_to
                note = item.clamp_note or agreement_clamp_note(
                    clamped_start=True, clamped_end=False
                )
                if note and "(applied while adding missing lane)" not in note:
                    note = f"{note} (applied while adding missing lane)"
                self.stats.earlier_validity.append(
                    EarlierValidityRow(
                        sheet=sheet_name,
                        update_row_index=item.idx,
                        key=key,
                        lane_no=lane.lane_no,
                        ra_row=lane.row,
                        original_valid_from=(
                            format_ra_date(item.original_from) if item.original_from else ""
                        ),
                        original_valid_to=(
                            format_ra_date(item.original_to) if item.original_to else ""
                        ),
                        effective_valid_from=format_ra_date(period[0]),
                        effective_valid_to=format_ra_date(period[1]),
                        agreement_valid_from=format_ra_date(agr_from) if agr_from else "",
                        agreement_valid_to=format_ra_date(agr_to) if agr_to else "",
                        note=note or "",
                    )
                )

        # Seed markers no longer needed for this key
        self._seeded_cells -= seeded_for_key

    def _force_create_validity_remainders(
        self,
        sheet_name: str,
        key: str,
        items: list[PendingLaneAdd],
    ) -> None:
        """
        KEY exists but BASE validity does not match: create only uncovered
        remainder segments, then re-apply each update across all covering lanes
        (existing get cost updates; new remainder gets costs added).
        """
        sample = items[0]
        meta = sample.meta
        cands, _note = find_lanes_for_key(self.ra, key)
        detail_ok: list[Any] = []
        for lane in cands:
            ok, _reason = validate_lane_details_excluding_validity(
                lane,
                carrier=meta.get("carrier") or "",
                service_c=meta.get("service_c") or "",
                origin_country=meta.get("origin_country") or "",
                dest_country=meta.get("dest_country") or "",
                origin_code=meta.get("origin_code") or "",
                dest_code=meta.get("dest_code") or "",
            )
            if ok:
                detail_ok.append(lane)

        covered: list[tuple[date, date]] = [
            (lane.valid_from, lane.valid_to)
            for lane in detail_ok
            if lane.valid_from is not None and lane.valid_to is not None
        ]

        # Prefer longest update row as cost seed for newly created remainders
        seed_item = max(
            items,
            key=lambda i: (
                (i.meta["valid_to"] - i.meta["valid_from"]).days,
                -i.meta["valid_from"].toordinal(),
            ),
        )
        seeded_for_key: set[tuple[int, int]] = set()
        created_periods: set[tuple[date, date]] = set()

        for item in items:
            u_from = item.meta["valid_from"]
            u_to = item.meta["valid_to"]
            for vf, vt in uncovered_period_segments(u_from, u_to, covered):
                if (vf, vt) in created_periods:
                    continue
                # Exact remainder already present (e.g. prior item created it)
                if any(
                    lane.valid_from == vf and lane.valid_to == vt for lane in detail_ok
                ):
                    created_periods.add((vf, vt))
                    continue

                lane, action = create_lane_from_update(
                    self.ra,
                    key=key,
                    carrier=meta.get("carrier") or "",
                    service_c=meta.get("service_c") or "",
                    origin_country=meta.get("origin_country") or "",
                    dest_country=meta.get("dest_country") or "",
                    origin_code=meta.get("origin_code") or "",
                    dest_code=meta.get("dest_code") or "",
                    valid_from=vf,
                    valid_to=vt,
                )
                created_periods.add((vf, vt))
                covered.append((vf, vt))
                detail_ok.append(lane)
                self.stats.changes.append(
                    ChangeRow(
                        sheet=sheet_name,
                        update_row_index=item.idx,
                        key=key,
                        lane_no=action.lane_no,
                        ra_row=action.ra_row,
                        rate_col="(lane add)",
                        cost_header=(
                            f"{action.detail} (remainder after existing validity)"
                        ),
                        punit_col=0,
                        old_value=None,
                        new_value=f"{format_ra_date(vf)}-{format_ra_date(vt)}",
                        status=action.action,
                    )
                )

                before_written = set(self.written.keys())
                self._apply_rates_to_lane(
                    sheet_name=sheet_name,
                    idx=seed_item.idx,
                    key=key,
                    row=seed_item.row,
                    lane=lane,
                    rates=seed_item.rates,
                    valid_from=vf,
                    valid_to=vt,
                    allow_seed_overwrite=False,
                )
                newly = set(self.written.keys()) - before_written
                seeded_for_key |= newly
                self._seeded_cells |= newly

        # Re-match: existing + remainder should cover the update without gaps
        for item in items:
            self._process_update_row(
                sheet_name,
                "BASE",
                item.idx,
                item.row,
                pending_adds=None,
            )
            if item.earlier_clamped:
                agr_from = self.ra.agreement_valid_from
                agr_to = self.ra.agreement_valid_to
                note = item.clamp_note or agreement_clamp_note(
                    clamped_start=True, clamped_end=False
                )
                if note and "(applied while adding missing lane)" not in note:
                    note = f"{note} (applied while adding missing lane)"
                period = (item.meta["valid_from"], item.meta["valid_to"])
                self.stats.earlier_validity.append(
                    EarlierValidityRow(
                        sheet=sheet_name,
                        update_row_index=item.idx,
                        key=key,
                        lane_no="",
                        ra_row=0,
                        original_valid_from=(
                            format_ra_date(item.original_from) if item.original_from else ""
                        ),
                        original_valid_to=(
                            format_ra_date(item.original_to) if item.original_to else ""
                        ),
                        effective_valid_from=format_ra_date(period[0]),
                        effective_valid_to=format_ra_date(period[1]),
                        agreement_valid_from=format_ra_date(agr_from) if agr_from else "",
                        agreement_valid_to=format_ra_date(agr_to) if agr_to else "",
                        note=note or "",
                    )
                )

        self._seeded_cells -= seeded_for_key

    def _apply_rates_to_lane(
        self,
        *,
        sheet_name: str,
        idx,
        key: str,
        row: pd.Series,
        lane,
        rates: list[dict],
        valid_from: date,
        valid_to: date,
        allow_seed_overwrite: bool,
    ) -> None:
        for r in rates:
            if r["unsupported"]:
                self._issue(
                    sheet=sheet_name,
                    update_row_index=idx,
                    key=key,
                    rate_col=r["rate_col"],
                    equipment=r["equipment"],
                    charge_type=r["charge_type"],
                    new_value=r["new_value"],
                    lane_no=lane.lane_no,
                    ra_row=lane.row,
                    reason=r["reason"] or "Not covered by current rules",
                )
                continue

            rule = r["rule"]
            costs, err = self._find_costs(
                ra_fee_name=rule.ra_fee_name,
                equipment=r["equipment"],
                valid_from=valid_from,
                valid_to=valid_to,
                require_period=rule.require_period,
                sheet_name=sheet_name,
                idx=idx,
                key=key,
                rate_col=r["rate_col"],
            )
            if not costs:
                self._issue(
                    sheet=sheet_name,
                    update_row_index=idx,
                    key=key,
                    rate_col=r["rate_col"],
                    equipment=r["equipment"],
                    charge_type=r["charge_type"],
                    new_value=r["new_value"],
                    lane_no=lane.lane_no,
                    ra_row=lane.row,
                    reason=err or "Cost column not found",
                )
                continue

            currency = currency_for_rate(row, r["rate_col"])

            for cost in costs:
                if not _cost_applies_on_lane(cost, lane):
                    continue

                self._write_currency(
                    lane=lane,
                    cost=cost,
                    currency=currency,
                    allow_seed_overwrite=allow_seed_overwrite,
                    mark_as_seed=not allow_seed_overwrite,
                )

                rate_target = cost.rate_col
                if rate_target is not None:
                    self._write_cell(
                        sheet_name=sheet_name,
                        idx=idx,
                        key=key,
                        lane=lane,
                        rate_col=r["rate_col"],
                        cost_header=(
                            f"{cost.header} "
                            f"[{'p/unit' if cost.punit_col == rate_target else 'Flat'}]"
                        ),
                        value_col=rate_target,
                        new_value=r["new_value"],
                        equipment=r.get("equipment"),
                        allow_seed_overwrite=allow_seed_overwrite,
                        mark_as_seed=not allow_seed_overwrite,
                        valid_from=valid_from,
                        valid_to=valid_to,
                    )
                else:
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=idx,
                        key=key,
                        rate_col=r["rate_col"],
                        equipment=r["equipment"],
                        charge_type=r["charge_type"],
                        new_value=r["new_value"],
                        lane_no=lane.lane_no,
                        ra_row=lane.row,
                        reason=f"No rate column (p/unit or Flat) for {cost.header!r}",
                    )

                min_val = r.get("min_value")
                min_target = cost.min_col
                if min_val is not None and min_target is not None:
                    self._write_cell(
                        sheet_name=sheet_name,
                        idx=idx,
                        key=key,
                        lane=lane,
                        rate_col=r.get("min_col") or f"{r['rate_col']}[MIN]",
                        cost_header=f"{cost.header} [Flat/MIN]",
                        value_col=min_target,
                        new_value=min_val,
                        equipment=r.get("equipment"),
                        allow_seed_overwrite=allow_seed_overwrite,
                        mark_as_seed=not allow_seed_overwrite,
                        valid_from=valid_from,
                        valid_to=valid_to,
                    )
                elif min_val is not None and min_target is None:
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=idx,
                        key=key,
                        rate_col=r.get("min_col"),
                        equipment=r["equipment"],
                        charge_type=r["charge_type"],
                        new_value=min_val,
                        lane_no=lane.lane_no,
                        ra_row=lane.row,
                        reason=(
                            f"Update has MIN={min_val} but RA cost {cost.header!r} "
                            "has no Flat/MIN column"
                        ),
                    )

    def _remember_write(
        self,
        cell_key: tuple[int, int],
        *,
        value: Any,
        key: str,
        valid_from: date | None,
        valid_to: date | None,
        equipment: str | None,
        change_index: int,
    ) -> None:
        self.written[cell_key] = value
        if equipment:
            self.written_equipment[cell_key] = equipment
        meta = self.written_meta.get(cell_key)
        if meta is None:
            self.written_meta[cell_key] = WrittenCellMeta(
                value=value,
                keys=[key] if key else [],
                valid_from=valid_from,
                valid_to=valid_to,
                equipment=equipment,
                change_indexes=[change_index],
            )
        else:
            meta.value = value
            if key and key not in meta.keys:
                meta.keys.append(key)
            meta.valid_from = valid_from
            meta.valid_to = valid_to
            if equipment:
                meta.equipment = equipment
            meta.change_indexes.append(change_index)

    def _sync_change_keys(self, cell_key: tuple[int, int]) -> str:
        meta = self.written_meta.get(cell_key)
        if not meta:
            return ""
        joined = _join_keys(meta.keys)
        for i in meta.change_indexes:
            if 0 <= i < len(self.stats.changes):
                self.stats.changes[i].key = joined
        return joined

    def _append_change(
        self,
        *,
        sheet_name: str,
        idx,
        key: str,
        lane,
        rate_col: str,
        cost_header: str,
        value_col: int,
        old_value: Any,
        new_value: Any,
        status: str,
    ) -> int:
        self.stats.changes.append(
            ChangeRow(
                sheet=sheet_name,
                update_row_index=idx,
                key=key,
                lane_no=lane.lane_no,
                ra_row=lane.row,
                rate_col=rate_col,
                cost_header=cost_header,
                punit_col=value_col,
                old_value=old_value,
                new_value=new_value,
                status=status,
            )
        )
        return len(self.stats.changes) - 1

    def _write_currency(
        self,
        *,
        lane,
        cost,
        currency: str | None,
        allow_seed_overwrite: bool = False,
        mark_as_seed: bool = False,
    ) -> None:
        """
        Fill RA Currency from update *_CURRENCY when present.

        Used for newly created cost columns and empty currency cells.
        Does not overwrite a non-empty currency unless seed overwrite applies.
        """
        if not currency or cost.currency_col is None:
            return
        cur_key = (lane.row, cost.currency_col)
        existing_cur = self.ra.get_value(lane.row, cost.currency_col)
        empty = existing_cur in (None, "")
        seeded = allow_seed_overwrite and cur_key in self._seeded_cells
        if not empty and not seeded:
            return
        if cur_key in self.written and not seeded:
            return
        self.ra.set_value(lane.row, cost.currency_col, currency)
        self.ra.ws.cell(lane.row, cost.currency_col).fill = GREEN
        self.written[cur_key] = currency
        if mark_as_seed:
            self._seeded_cells.add(cur_key)
        elif seeded:
            self._seeded_cells.discard(cur_key)

    def _write_cell(
        self,
        *,
        sheet_name: str,
        idx,
        key: str,
        lane,
        rate_col: str,
        cost_header: str,
        value_col: int,
        new_value: float,
        equipment: str | None = None,
        allow_seed_overwrite: bool = False,
        mark_as_seed: bool = False,
        valid_from: date | None = None,
        valid_to: date | None = None,
    ) -> None:
        cell_key = (lane.row, value_col)
        old_value = self.ra.get_value(lane.row, value_col)
        eq = (equipment or "").strip().upper() or None
        clear = _is_zero_clear(new_value)

        # Strict rule: never rewrite a cell already written in this run,
        # except same-value key merge, non-BU prefer, seed overwrite, or alias dual.
        if cell_key in self.written:
            prev = self.written[cell_key]
            prev_text = str(prev)
            prev_eq = self.written_equipment.get(cell_key)
            prev_meta = self.written_meta.get(cell_key)
            is_dual_text = isinstance(prev, str) and " / " in prev_text and "(" in prev_text

            # Update 0 → clear: empty/prior-clear merges; else overwrite with empty
            if clear:
                prev_cleared = _cell_is_empty(prev) or _is_zero_clear(prev)
                if prev_cleared:
                    if prev_meta is not None and key and key not in prev_meta.keys:
                        prev_meta.keys.append(key)
                    joined = self._sync_change_keys(cell_key) or key
                    ch_i = self._append_change(
                        sheet_name=sheet_name,
                        idx=idx,
                        key=joined,
                        lane=lane,
                        rate_col=rate_col,
                        cost_header=cost_header,
                        value_col=value_col,
                        old_value=old_value,
                        new_value=None,
                        status="already_equal",
                    )
                    if prev_meta is not None:
                        prev_meta.change_indexes.append(ch_i)
                        self._sync_change_keys(cell_key)
                    return
                self.ra.set_value(lane.row, value_col, None)
                self.ra.ws.cell(lane.row, value_col).fill = GREEN
                self._seeded_cells.discard(cell_key)
                joined = _join_keys(
                    (prev_meta.keys if prev_meta else []) + [key]
                )
                ch_i = self._append_change(
                    sheet_name=sheet_name,
                    idx=idx,
                    key=joined or key,
                    lane=lane,
                    rate_col=rate_col,
                    cost_header=cost_header,
                    value_col=value_col,
                    old_value=prev,
                    new_value=None,
                    status="cleared",
                )
                self._remember_write(
                    cell_key,
                    value=None,
                    key=key,
                    valid_from=valid_from,
                    valid_to=valid_to,
                    equipment=eq,
                    change_index=ch_i,
                )
                if prev_meta is not None and key not in prev_meta.keys:
                    prev_meta.keys.append(key)
                self._sync_change_keys(cell_key)
                return

            # 1) Same value → merge keys onto the change row(s)
            if numbers_equal(prev, new_value):
                if prev_meta is not None and key and key not in prev_meta.keys:
                    prev_meta.keys.append(key)
                joined = self._sync_change_keys(cell_key) or key
                ch_i = self._append_change(
                    sheet_name=sheet_name,
                    idx=idx,
                    key=joined,
                    lane=lane,
                    rate_col=rate_col,
                    cost_header=cost_header,
                    value_col=value_col,
                    old_value=old_value,
                    new_value=new_value,
                    status="already_equal",
                )
                if prev_meta is not None:
                    prev_meta.change_indexes.append(ch_i)
                    self._sync_change_keys(cell_key)
                return

            if is_dual_text and eq is not None and (
                f"({eq})" in prev_text or (prev_eq is not None and eq == prev_eq)
            ):
                self._append_change(
                    sheet_name=sheet_name,
                    idx=idx,
                    key=key,
                    lane=lane,
                    rate_col=rate_col,
                    cost_header=cost_header,
                    value_col=value_col,
                    old_value=old_value,
                    new_value=prev,
                    status="already_equal",
                )
                return

            # Confirmed runtime aliases only (not static groups)
            if (
                eq
                and prev_eq
                and eq != prev_eq
                and self.equipment_store.are_aliased(eq, prev_eq)
            ):
                left = prev_text if is_dual_text else _format_rate_display(prev)
                dual_val = f"{left} / {_format_rate_display(new_value)} ({eq})"
                self.ra.set_value(lane.row, value_col, dual_val)
                self.ra.ws.cell(lane.row, value_col).fill = RED
                self._seeded_cells.discard(cell_key)
                ch_i = self._append_change(
                    sheet_name=sheet_name,
                    idx=idx,
                    key=key,
                    lane=lane,
                    rate_col=rate_col,
                    cost_header=cost_header,
                    value_col=value_col,
                    old_value=prev,
                    new_value=dual_val,
                    status="dual_container",
                )
                self._remember_write(
                    cell_key,
                    value=dual_val,
                    key=key,
                    valid_from=valid_from,
                    valid_to=valid_to,
                    equipment=eq,
                    change_index=ch_i,
                )
                self._issue(
                    sheet=sheet_name,
                    update_row_index=idx,
                    key=key,
                    rate_col=rate_col,
                    equipment=eq,
                    charge_type=None,
                    new_value=dual_val,
                    lane_no=lane.lane_no,
                    ra_row=lane.row,
                    reason=(
                        f"Dual container rate on {cost_header!r}: "
                        f"kept {_format_rate_display(prev)} ({prev_eq}) and added "
                        f"{_format_rate_display(new_value)} ({eq})"
                    ),
                )
                return

            if allow_seed_overwrite and cell_key in self._seeded_cells:
                self._seeded_cells.discard(cell_key)
                # fall through and overwrite seed
            else:
                prev_keys = list(prev_meta.keys) if prev_meta else []
                prev_from = prev_meta.valid_from if prev_meta else None
                prev_to = prev_meta.valid_to if prev_meta else None
                same_validity = periods_equal(
                    prev_from, prev_to, valid_from, valid_to
                )
                intersecting = _periods_intersect(
                    prev_from, prev_to, valid_from, valid_to
                )
                if prev_meta and prev_meta.keys:
                    prev_has_non_bu = any(not _key_has_bu(k) for k in prev_meta.keys)
                    prev_has_bu = any(_key_has_bu(k) for k in prev_meta.keys)
                else:
                    prev_has_non_bu = False
                    prev_has_bu = False
                new_is_bu = _key_has_bu(key)
                keys_differ = bool(prev_keys) and key not in prev_keys

                # 2) Different keys, same validity → prefer non-BU value
                if keys_differ and same_validity and (prev_has_bu ^ new_is_bu):
                    # Prefer incoming non-BU over existing BU-only write
                    prefer_new = (not new_is_bu) and prev_has_bu and not prev_has_non_bu
                    # Keep existing non-BU when incoming is BU
                    prefer_keep = prev_has_non_bu and new_is_bu
                    if prefer_new:
                        write_val: Any = (
                            int(new_value)
                            if float(new_value).is_integer()
                            else float(new_value)
                        )
                        self.ra.set_value(lane.row, value_col, write_val)
                        self.ra.ws.cell(lane.row, value_col).fill = GREEN
                        self._seeded_cells.discard(cell_key)
                        if prev_meta is not None and key not in prev_meta.keys:
                            prev_meta.keys.append(key)
                        joined = _join_keys(
                            (prev_meta.keys if prev_meta else prev_keys) + [key]
                        )
                        ch_i = self._append_change(
                            sheet_name=sheet_name,
                            idx=idx,
                            key=joined,
                            lane=lane,
                            rate_col=rate_col,
                            cost_header=cost_header,
                            value_col=value_col,
                            old_value=prev,
                            new_value=write_val,
                            status="changed",
                        )
                        self._remember_write(
                            cell_key,
                            value=write_val,
                            key=key,
                            valid_from=valid_from,
                            valid_to=valid_to,
                            equipment=eq,
                            change_index=ch_i,
                        )
                        self._sync_change_keys(cell_key)
                        return
                    if prefer_keep:
                        if prev_meta is not None and key not in prev_meta.keys:
                            prev_meta.keys.append(key)
                        joined = self._sync_change_keys(cell_key) or _join_keys(
                            prev_keys + [key]
                        )
                        self._append_change(
                            sheet_name=sheet_name,
                            idx=idx,
                            key=joined,
                            lane=lane,
                            rate_col=rate_col,
                            cost_header=cost_header,
                            value_col=value_col,
                            old_value=old_value,
                            new_value=prev,
                            status="already_equal",
                        )
                        return

                # 3) Different validity but intersecting
                #    SPOT involved → write SPOT price only; note dual validity.
                #    Non-DFT: keep both rates with validity brackets, red fill.
                #    DFT: leave as issue for now (no dual write).
                if (
                    prev_from is not None
                    and prev_to is not None
                    and valid_from is not None
                    and valid_to is not None
                    and not same_validity
                    and intersecting
                ):
                    joined = _join_keys(prev_keys + [key])
                    if _is_dft_equipment(eq) or _is_dft_equipment(prev_eq):
                        self._issue(
                            sheet=sheet_name,
                            update_row_index=idx,
                            key=joined,
                            rate_col=rate_col,
                            equipment=eq,
                            charge_type=None,
                            new_value=new_value,
                            lane_no=lane.lane_no,
                            ra_row=lane.row,
                            reason=(
                                f"Conflicting updates on {cost_header!r} "
                                f"(row={lane.row}, col={value_col}): "
                                f"existing={prev!r} from {prev_keys} "
                                f"validity {format_ra_date(prev_from)}-"
                                f"{format_ra_date(prev_to)}; "
                                f"attempted={new_value!r} from {key!r} "
                                f"validity {format_ra_date(valid_from)}-"
                                f"{format_ra_date(valid_to)}. "
                                f"Validities intersect with different periods "
                                f"(DFT: dual write not applied)."
                            ),
                        )
                        return

                    if is_dual_text:
                        left = prev_text
                    else:
                        left = _format_rate_with_validity(prev, prev_from, prev_to)
                    right = _format_rate_with_validity(
                        new_value, valid_from, valid_to
                    )
                    dual_desc = f"{left} / {right}"

                    prev_has_spot = any(_key_has_spot(k) for k in prev_keys)
                    new_is_spot = _key_has_spot(key)
                    if prev_has_spot or new_is_spot:
                        # Prefer SPOT numeric rate; document the dual pair in a note
                        if new_is_spot and not prev_has_spot:
                            write_val: Any = (
                                int(new_value)
                                if float(new_value).is_integer()
                                else float(new_value)
                            )
                            chosen_from, chosen_to = valid_from, valid_to
                            status = "changed"
                        elif prev_has_spot and not new_is_spot:
                            write_val = prev
                            chosen_from, chosen_to = prev_from, prev_to
                            status = "already_equal"
                        else:
                            # Both SPOT → keep latest incoming SPOT rate
                            write_val = (
                                int(new_value)
                                if float(new_value).is_integer()
                                else float(new_value)
                            )
                            chosen_from, chosen_to = valid_from, valid_to
                            status = "changed"

                        if status == "changed":
                            self.ra.set_value(lane.row, value_col, write_val)
                            self.ra.ws.cell(lane.row, value_col).fill = GREEN
                            self._seeded_cells.discard(cell_key)
                        ch_i = self._append_change(
                            sheet_name=sheet_name,
                            idx=idx,
                            key=joined,
                            lane=lane,
                            rate_col=rate_col,
                            cost_header=cost_header,
                            value_col=value_col,
                            old_value=prev,
                            new_value=write_val,
                            status=status,
                        )
                        self._remember_write(
                            cell_key,
                            value=write_val,
                            key=key,
                            valid_from=chosen_from,
                            valid_to=chosen_to,
                            equipment=eq,
                            change_index=ch_i,
                        )
                        if prev_meta is not None and key not in prev_meta.keys:
                            prev_meta.keys.append(key)
                        self._sync_change_keys(cell_key)
                        self._issue(
                            sheet=sheet_name,
                            update_row_index=idx,
                            key=joined,
                            rate_col=rate_col,
                            equipment=eq,
                            charge_type=None,
                            new_value=write_val,
                            lane_no=lane.lane_no,
                            ra_row=lane.row,
                            reason=(
                                f"SPOT price preferred on {cost_header!r} "
                                f"(row={lane.row}, col={value_col}): wrote {write_val!r} "
                                f"(SPOT). Dual validity was {dual_desc}; "
                                f"SPOT rate kept, non-SPOT not written as dual cell."
                            ),
                        )
                        return

                    dual_val = dual_desc
                    self.ra.set_value(lane.row, value_col, dual_val)
                    self.ra.ws.cell(lane.row, value_col).fill = RED
                    self._seeded_cells.discard(cell_key)
                    ch_i = self._append_change(
                        sheet_name=sheet_name,
                        idx=idx,
                        key=joined,
                        lane=lane,
                        rate_col=rate_col,
                        cost_header=cost_header,
                        value_col=value_col,
                        old_value=prev,
                        new_value=dual_val,
                        status="dual_validity",
                    )
                    self._remember_write(
                        cell_key,
                        value=dual_val,
                        key=key,
                        valid_from=valid_from,
                        valid_to=valid_to,
                        equipment=eq,
                        change_index=ch_i,
                    )
                    self._issue(
                        sheet=sheet_name,
                        update_row_index=idx,
                        key=joined,
                        rate_col=rate_col,
                        equipment=eq,
                        charge_type=None,
                        new_value=dual_val,
                        lane_no=lane.lane_no,
                        ra_row=lane.row,
                        reason=(
                            f"Dual validity rates on {cost_header!r} "
                            f"(row={lane.row}, col={value_col}): {dual_val}. "
                            f"Validities intersect with different periods; "
                            f"both values kept."
                        ),
                    )
                    return

                self._issue(
                    sheet=sheet_name,
                    update_row_index=idx,
                    key=_join_keys(prev_keys + [key]) or key,
                    rate_col=rate_col,
                    equipment=eq,
                    charge_type=None,
                    new_value=new_value,
                    lane_no=lane.lane_no,
                    ra_row=lane.row,
                    reason=(
                        f"Cell already written this run (row={lane.row}, col={value_col}, "
                        f"cost={cost_header!r}): existing_write={prev!r}, attempted={new_value!r}. "
                        "Not covered by rules / conflicting updates."
                    ),
                )
                return

        # Update 0 → empty the RA cell (do not write numeric 0)
        if clear:
            already_empty = _cell_is_empty(old_value)
            if not already_empty:
                self.ra.set_value(lane.row, value_col, None)
                self.ra.ws.cell(lane.row, value_col).fill = GREEN
            if mark_as_seed:
                self._seeded_cells.add(cell_key)
            ch_i = self._append_change(
                sheet_name=sheet_name,
                idx=idx,
                key=key,
                lane=lane,
                rate_col=rate_col,
                cost_header=cost_header,
                value_col=value_col,
                old_value=old_value,
                new_value=None,
                status="already_equal" if already_empty else "cleared",
            )
            self._remember_write(
                cell_key,
                value=None,
                key=key,
                valid_from=valid_from,
                valid_to=valid_to,
                equipment=eq,
                change_index=ch_i,
            )
            return

        if numbers_equal(old_value, new_value):
            write_num = to_number(new_value)
            if mark_as_seed:
                self._seeded_cells.add(cell_key)
            ch_i = self._append_change(
                sheet_name=sheet_name,
                idx=idx,
                key=key,
                lane=lane,
                rate_col=rate_col,
                cost_header=cost_header,
                value_col=value_col,
                old_value=old_value,
                new_value=new_value,
                status="already_equal",
            )
            self._remember_write(
                cell_key,
                value=write_num,
                key=key,
                valid_from=valid_from,
                valid_to=valid_to,
                equipment=eq,
                change_index=ch_i,
            )
            return

        write_val = int(new_value) if float(new_value).is_integer() else float(new_value)
        self.ra.set_value(lane.row, value_col, write_val)
        self.ra.ws.cell(lane.row, value_col).fill = GREEN
        if mark_as_seed:
            self._seeded_cells.add(cell_key)
        ch_i = self._append_change(
            sheet_name=sheet_name,
            idx=idx,
            key=key,
            lane=lane,
            rate_col=rate_col,
            cost_header=cost_header,
            value_col=value_col,
            old_value=old_value,
            new_value=write_val,
            status="changed",
        )
        self._remember_write(
            cell_key,
            value=write_val,
            key=key,
            valid_from=valid_from,
            valid_to=valid_to,
            equipment=eq,
            change_index=ch_i,
        )
