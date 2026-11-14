from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .config import (
    DATA_START_ROW,
    GENERAL_INFO_SHEET,
    HEADER_COST_ROW,
    HEADER_FIELD_ROW,
    RA_SHEET,
)
from .cost_map import CostColumn, parse_cost_header
from .dates import parse_date, parse_validity_period_text
from .lane_columns import LaneColumns, resolve_lane_columns
from .parsing import split_multi


@dataclass
class RaLane:
    row: int
    lane_no: str
    key_raw: str
    keys: list[str]
    carrier: str
    service_c: str
    valid_from: date | None
    valid_to: date | None
    origin_port: str
    origin_postal: str
    origin_country: str
    dest_port: str
    dest_postal: str
    dest_country: str


class RateCardWorkbook:
    def __init__(self, path: Path):
        self.path = path
        # keep_vba=False; data_only=False so we can write values back
        self.wb = load_workbook(path)
        if RA_SHEET not in self.wb.sheetnames:
            raise ValueError(f"Sheet {RA_SHEET!r} not found in {path.name}")
        self.ws: Worksheet = self.wb[RA_SHEET]
        self.cols: LaneColumns = resolve_lane_columns(self.ws)
        self.agreement_title, self.agreement_valid_from, self.agreement_valid_to = (
            self._load_agreement_info()
        )
        self.cost_columns = self._build_cost_columns()
        self.lanes = self._build_lanes()
        self._key_index: dict[str, list[RaLane]] = {}
        for lane in self.lanes:
            for k in lane.keys:
                self._key_index.setdefault(k, []).append(lane)

    def _cell_at(self, row: int, col: int | None) -> str:
        if col is None:
            return ""
        return self._as_str(self._cell(row, col))

    def _load_agreement_info(self) -> tuple[str | None, date | None, date | None]:
        if GENERAL_INFO_SHEET not in self.wb.sheetnames:
            return None, None, None
        gi = self.wb[GENERAL_INFO_SHEET]
        title: str | None = None
        valid_from: date | None = None
        valid_to: date | None = None
        for row in gi.iter_rows(min_row=1, max_row=gi.max_row, max_col=2, values_only=True):
            label = str(row[0]).strip().lower() if row[0] is not None else ""
            value = row[1] if len(row) > 1 else None
            if "agreement title" in label and value is not None:
                title = str(value).strip() or None
            if "validity period" in label:
                valid_from, valid_to = parse_validity_period_text(value)
        if title is None:
            # Fallback: sheet cell A2 often holds the RA short name (e.g. Sea)
            a2 = gi.cell(2, 1).value
            if a2 is not None and str(a2).strip():
                title = str(a2).strip()
        return title, valid_from, valid_to

    def _cell(self, row: int, col: int):
        return self.ws.cell(row, col).value

    def _as_str(self, value) -> str:
        if value is None:
            return ""
        return str(value).strip()

    def _build_cost_columns(self) -> list[CostColumn]:
        """
        Build cost column metadata from Rate Card headers.

        Field layout within a cost block (until the next cost header):
          Currency | Flat (optional; may be MIN) | p/unit (optional)
        Row 6 above Flat may say MIN -> Flat stores minimum, p/unit stores rate.
        Documentation Fee has Flat only (rate in Flat, no MIN label).
        """
        ws = self.ws
        field_row = HEADER_FIELD_ROW
        cost_row = HEADER_COST_ROW
        min_marker_row = 6
        columns: list[CostColumn] = []

        header_cols = [
            c
            for c in range(1, ws.max_column + 1)
            if ws.cell(cost_row, c).value not in (None, "")
        ]

        for i, col in enumerate(header_cols):
            header = str(ws.cell(cost_row, col).value).strip()
            fee_name, equipment, v_from, v_to = parse_cost_header(header)
            block_end = header_cols[i + 1] if i + 1 < len(header_cols) else ws.max_column + 1

            currency_col = None
            flat_col = None
            flat_is_min = False
            punit_col = None

            for c in range(col, block_end):
                label = ws.cell(field_row, c).value
                if label is None:
                    continue
                lab = str(label).strip().lower()
                if lab == "currency" and currency_col is None:
                    currency_col = c
                elif lab == "flat" and flat_col is None:
                    flat_col = c
                    marker = ws.cell(min_marker_row, c).value
                    flat_is_min = marker is not None and str(marker).strip().upper() == "MIN"
                elif lab == "p/unit" and punit_col is None:
                    punit_col = c

            # Need at least one writable value column
            if flat_col is None and punit_col is None:
                continue

            columns.append(
                CostColumn(
                    start_col=col,
                    header=header,
                    fee_name=fee_name,
                    equipment=equipment,
                    valid_from=v_from,
                    valid_to=v_to,
                    currency_col=currency_col,
                    flat_col=flat_col,
                    flat_is_min=flat_is_min,
                    punit_col=punit_col,
                )
            )
        return columns

    def _build_lanes(self) -> list[RaLane]:
        lanes: list[RaLane] = []
        for row in range(DATA_START_ROW, self.ws.max_row + 1):
            key_raw = self._cell_at(row, self.cols.key)
            lane_no = self._cell_at(row, self.cols.lane)
            if not key_raw and not lane_no:
                continue
            lanes.append(self._lane_from_row(row))
        return lanes

    def find_by_key(self, key: str) -> list[RaLane]:
        return list(self._key_index.get(str(key).strip(), []))

    def _lane_from_row(self, row: int) -> RaLane:
        c = self.cols
        key_raw = self._cell_at(row, c.key)
        return RaLane(
            row=row,
            lane_no=self._cell_at(row, c.lane),
            key_raw=key_raw,
            keys=split_multi(key_raw),
            carrier=self._cell_at(row, c.carrier),
            service_c=self._cell_at(row, c.service_c),
            valid_from=parse_date(self._cell(row, c.valid_from)),
            valid_to=parse_date(self._cell(row, c.valid_to)),
            origin_port=self._cell_at(row, c.origin_port),
            origin_postal=self._cell_at(row, c.origin_postal),
            origin_country=self._cell_at(row, c.origin_country),
            dest_port=self._cell_at(row, c.dest_port),
            dest_postal=self._cell_at(row, c.dest_postal),
            dest_country=self._cell_at(row, c.dest_country),
        )

    def register_lane_from_row(self, row: int) -> RaLane:
        """Add a newly inserted worksheet row into the in-memory lane index."""
        lane = self._lane_from_row(row)
        self.lanes.append(lane)
        for k in lane.keys:
            self._key_index.setdefault(k, []).append(lane)
        return lane

    def get_value(self, row: int, col: int):
        return self.ws.cell(row, col).value

    def set_value(self, row: int, col: int, value):
        self.ws.cell(row, col).value = value

    def get_punit(self, row: int, punit_col: int):
        return self.get_value(row, punit_col)

    def set_punit(self, row: int, punit_col: int, value):
        self.set_value(row, punit_col, value)

    def save(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(path)
