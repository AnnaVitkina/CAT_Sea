"""
Resolve missing equipment cost columns interactively (or via a decisions file).

Aliases are same Rate-by only (e.g. 20G0/22G0/22SOC = Container/20FT).
Solo types (e.g. 42P1 = Container/40FR alone) have no alias peer and should be created;
a layout donor (e.g. 42G0) is only used to clone column structure, not as an alias.

For each update equipment code that has rates but no matching RA cost header:
  1) Prefer CREATE when no same-type alias exists; clone layout from a size donor
  2) Prefer ALIAS when a same Rate-by peer is already on the RA
  3) Apply that decision for the rest of the run
"""

from __future__ import annotations

import json
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import pandas as pd
from openpyxl.styles import PatternFill

from .charge_rules import (
    EQUIPMENT_RATE_BY,
    equipment_alias_group,
    equipment_rate_by,
    preferred_alias_template,
    resolve_rate_column,
)
from .config import (
    CHANGED_FILL_HEX,
    DATA_START_ROW,
    HEADER_COST_ROW,
    HEADER_FIELD_ROW,
    UPDATE_SHEETS,
)
from .cost_map import CostColumn, find_cost_columns, parse_cost_header
from .dates import format_ra_date
from .parsing import to_number

GREEN = PatternFill(start_color=CHANGED_FILL_HEX, end_color=CHANGED_FILL_HEX, fill_type="solid")

@dataclass
class EquipmentDecision:
    equipment: str
    action: str  # create | alias | skip
    target: str | None = None  # alias target, or create template equipment
    fees_created: list[str] = field(default_factory=list)
    note: str = ""


@dataclass
class EquipmentDecisionStore:
    """Runtime decisions: aliases and created equipment codes."""

    aliases: dict[str, str] = field(default_factory=dict)  # update_eq -> ra_eq
    created: set[str] = field(default_factory=set)
    skipped: set[str] = field(default_factory=set)
    decisions: list[EquipmentDecision] = field(default_factory=list)

    def resolve_equipment(self, equipment: str | None) -> str | None:
        if not equipment:
            return equipment
        eq = str(equipment).strip().upper()
        return self.aliases.get(eq, eq)

    def are_aliased(self, a: str | None, b: str | None) -> bool:
        if not a or not b:
            return False
        a_u, b_u = a.strip().upper(), b.strip().upper()
        if a_u == b_u:
            return False
        return self.resolve_equipment(a_u) == self.resolve_equipment(b_u)


def suggest_alias_equipment(equipment: str, present: set[str]) -> str | None:
    """Same Rate-by group member already on the RA (alias peer), else None."""
    return preferred_alias_template(equipment, present)


def suggest_layout_template(equipment: str, present: set[str]) -> str | None:
    """
    Equipment to clone *column layout* from when creating a new code.
    Prefers same-group alias; else a common size donor (not treated as alias).
    """
    alias = preferred_alias_template(equipment, present)
    if alias:
        return alias
    eq = (equipment or "").strip().upper()
    present_u = {str(p).strip().upper() for p in present if p}
    # Prefer same size family for layout only (40xx → 42G0/45G0, 20/22 → 22G0)
    if eq.startswith("45") or eq.startswith("L2"):
        for cand in ("45G0", "42G0", "22G0"):
            if cand in present_u:
                return cand
    if eq.startswith("42"):
        for cand in ("42G0", "45G0", "22G0"):
            if cand in present_u:
                return cand
    if eq.startswith(("20", "22", "25")):
        for cand in ("22G0", "20G0", "45G0"):
            if cand in present_u:
                return cand
    return sorted(present_u)[0] if present_u else None


def rate_by_for_new_equipment(new_eq: str, template_rate_by: str | None) -> str:
    known = equipment_rate_by(new_eq) or EQUIPMENT_RATE_BY.get(
        str(new_eq).strip().upper()
    )
    if known:
        return known
    if template_rate_by:
        text = template_rate_by.strip()
        if "weight" in text.lower():
            return f"Container/{new_eq}"
        return text
    return f"Container/{new_eq}"


def _prompt(text: str, default: str = "") -> str:
    """
    Ask via input() so Colab / Jupyter cells work (stdin often isatty()=False).
    Falls back to readline on a real TTY if input() is unavailable.
    """
    import sys

    suffix = f" [{default}]" if default else ""
    prompt = f"{text}{suffix}: "
    try:
        raw = input(prompt).strip()
    except EOFError:
        raise SystemExit(
            "Interactive equipment prompts need usable input(). "
            "In Colab use `%run run_pipeline.py` (not !python), "
            "or pass --non-interactive with input/equipment_decisions.json."
        ) from None
    except Exception:
        # Non-notebook fallback
        print(prompt, end="", flush=True)
        try:
            raw = sys.stdin.readline()
            if raw == "":
                raise EOFError("empty stdin")
            raw = raw.strip()
        except EOFError:
            raise SystemExit(
                "Interactive equipment prompts need usable input(). "
                "In Colab use `%run run_pipeline.py` (not !python), "
                "or pass --non-interactive with input/equipment_decisions.json."
            ) from None
    return raw if raw else default


def load_decisions_file(path: Path) -> dict[str, dict]:
    """
    Optional JSON:
      {"22SOC": {"action": "alias", "target": "22G0"},
       "42SOC": {"action": "create", "target": "42G0"}}
    """
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        return {}
    return {
        str(k).upper(): v
        for k, v in data.items()
        if isinstance(v, dict) and not str(k).startswith("_")
    }


def save_decisions_file(path: Path, store: EquipmentDecisionStore) -> None:
    # Preserve sibling keys such as _baf_create_until
    existing: dict = {}
    if path.exists():
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                existing = {
                    k: v
                    for k, v in loaded.items()
                    if str(k).startswith("_")
                }
        except json.JSONDecodeError:
            existing = {}
    payload = dict(existing)
    for d in store.decisions:
        payload[d.equipment] = {
            "action": d.action,
            "target": d.target,
            "fees_created": d.fees_created,
            "note": d.note,
        }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def collect_missing_equipment_fees(
    ra,
    update_path: Path,
) -> dict[str, set[str]]:
    """equipment -> set of ra_fee_name that are missing on the Rate Card."""
    missing: dict[str, set[str]] = {}
    present_eq = {c.equipment for c in ra.cost_columns if c.equipment}

    for sheet in UPDATE_SHEETS:
        df = pd.read_excel(update_path, sheet_name=sheet)
        for _, row in df.iterrows():
            for col, raw in row.items():
                col_s = str(col)
                if not col_s.endswith("_RATE"):
                    continue
                if to_number(raw) is None:
                    continue
                rule, equipment, skip = resolve_rate_column(col_s, sheet)
                if skip or rule is None or not equipment:
                    continue
                # Period fees (BAF / EU ETS) are created on demand with monthly /
                # exact-day logic — not during equipment create/alias prompts.
                if rule.require_period:
                    continue
                eq = equipment.upper()
                # Exact equipment already on RA for this fee?
                cols, err = find_cost_columns(
                    ra.cost_columns,
                    ra_fee_name=rule.ra_fee_name,
                    equipment=eq,
                    valid_from=None,
                    valid_to=None,
                    require_period=False,
                )
                if cols:
                    continue
                # Only treat as missing-equipment when no exact header exists
                if any(
                    c.fee_name == rule.ra_fee_name and c.equipment == eq
                    for c in ra.cost_columns
                ):
                    continue
                # Skip if equipment itself exists for some fee (partial) — still list missing fees
                missing.setdefault(eq, set()).add(rule.ra_fee_name)
    return missing


def _template_cost_for_fee(ra, fee_name: str, template_eq: str) -> CostColumn | None:
    for c in ra.cost_columns:
        if c.fee_name == fee_name and c.equipment == template_eq:
            if c.valid_from is None and c.valid_to is None:
                return c
    for c in ra.cost_columns:
        if c.fee_name == fee_name and c.equipment == template_eq:
            return c
    return None


def _block_width(ra, cost: CostColumn) -> int:
    starts = sorted(c.start_col for c in ra.cost_columns)
    try:
        i = starts.index(cost.start_col)
    except ValueError:
        return 2
    if i + 1 < len(starts):
        return max(1, starts[i + 1] - cost.start_col)
    # last block: count non-empty field labels
    width = 0
    for c in range(cost.start_col, ra.ws.max_column + 2):
        if c > cost.start_col and ra.ws.cell(HEADER_COST_ROW, c).value not in (None, ""):
            break
        lab = ra.ws.cell(HEADER_FIELD_ROW, c).value
        if lab is None and c > cost.start_col:
            # allow one empty then stop if past known value cols
            if cost.punit_col and c > cost.punit_col:
                break
            if cost.flat_col and c > (cost.punit_col or cost.flat_col):
                break
        width += 1
        if width >= 4:
            break
    return max(width, 2)


def create_fee_only_cost_column(
    ra,
    *,
    fee_name: str,
    template: CostColumn,
    rate_by: str,
) -> CostColumn:
    """
    Append a fee-only cost block (no equipment token), e.g.
      War Risk Fee
    with Rate by: Weight/chargeable kg.
    """
    for c in ra.cost_columns:
        if c.header == fee_name or (c.fee_name == fee_name and c.equipment is None):
            return c

    ws = ra.ws
    width = _block_width(ra, template)
    start = ws.max_column + 1
    header = fee_name
    applies = ws.cell(4, template.start_col).value
    if applies is None:
        applies = "Applies if invoiced by Carrier"
    rate_by_text = f"Rate by: {rate_by}\nRegular rule"

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


def ensure_weight_fee_only(
    ra,
    *,
    fee_name: str,
    legacy_equipment: str | None = None,
) -> tuple[CostColumn | None, str | None]:
    """
    Ensure a fee-only cost column exists with Rate by: Weight/chargeable kg
    (header = fee name only, no equipment token).

    Also accepts a legacy header like 'Suez Canal Charge (SCANAL)' if present.
    """
    from .charge_rules import DFT_WEIGHT_RATE_BY

    existing = [
        c
        for c in ra.cost_columns
        if c.fee_name == fee_name and c.equipment is None
    ]
    if not existing:
        existing = [c for c in ra.cost_columns if c.header == fee_name]
    if not existing and legacy_equipment:
        leg = legacy_equipment.upper()
        existing = [
            c
            for c in ra.cost_columns
            if c.fee_name == fee_name and (c.equipment or "").upper() == leg
        ]
    if existing:
        return existing[0], None

    template = next(
        (c for c in ra.cost_columns if c.fee_name == fee_name),
        None,
    )
    if template is None:
        template = next(
            (
                c
                for c in ra.cost_columns
                if c.fee_name in {"War Risk Fee", "Base Rate", "BAF Fee", "EU ETS Fee"}
            ),
            None,
        )
    if template is None:
        return None, f"No template column available to create weight {fee_name}"

    created = create_fee_only_cost_column(
        ra,
        fee_name=fee_name,
        template=template,
        rate_by=DFT_WEIGHT_RATE_BY,
    )
    print(f"  Created {created.header!r} (Rate by: {DFT_WEIGHT_RATE_BY})")
    return created, None


def ensure_weight_war_risk_fee(ra) -> tuple[CostColumn | None, str | None]:
    """*DFTWM_WAR_RISK → fee-only War Risk Fee (Weight/chargeable kg)."""
    return ensure_weight_fee_only(ra, fee_name="War Risk Fee")


def ensure_weight_suez_canal_charge(ra) -> tuple[CostColumn | None, str | None]:
    """*DFTWM_SCANAL → fee-only Suez Canal Charge (Weight/chargeable kg)."""
    return ensure_weight_fee_only(
        ra, fee_name="Suez Canal Charge", legacy_equipment="SCANAL"
    )


def create_cost_column_from_template(
    ra,
    *,
    fee_name: str,
    new_equipment: str,
    template: CostColumn,
    rate_by: str,
) -> CostColumn:
    """
    Append a new cost block cloned from `template`, with header like
    'Base Rate (22SOC)' and Rate by: Container/20SOC.
    """
    ws = ra.ws
    width = _block_width(ra, template)
    start = ws.max_column + 1

    # Build header (period-less BASE-style; period templates keep period text swapped)
    if template.valid_from is not None and template.valid_to is not None:
        header = (
            f"{fee_name} ({new_equipment} "
            f"{format_ra_date(template.valid_from)}-{format_ra_date(template.valid_to)})"
        )
    else:
        header = f"{fee_name} ({new_equipment})"

    applies = ws.cell(4, template.start_col).value
    if applies is None:
        applies = "Applies if invoiced by Carrier"
    rate_by_text = f"Rate by: {rate_by}\nRegular rule"

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

        # Clear lane data in new columns
        for row in range(DATA_START_ROW, ws.max_row + 1):
            ws.cell(row, dst_col).value = None

    ws.cell(HEADER_COST_ROW, start).value = header
    ws.cell(HEADER_COST_ROW, start).fill = GREEN
    ws.cell(4, start).value = applies
    ws.cell(5, start).value = rate_by_text
    ws.cell(5, start).fill = GREEN

    # Field labels from template block
    for offset in range(width):
        src_col = template.start_col + offset
        dst_col = start + offset
        ws.cell(HEADER_FIELD_ROW, dst_col).value = ws.cell(HEADER_FIELD_ROW, src_col).value
        ws.cell(6, dst_col).value = ws.cell(6, src_col).value

    # Register metadata
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


def resolve_missing_equipment_interactive(
    ra,
    update_path: Path,
    store: EquipmentDecisionStore,
    *,
    decisions_path: Path | None = None,
    interactive: bool = True,
    ask_fn: Callable[[str, str], str] | None = None,
) -> EquipmentDecisionStore:
    """
    For each missing equipment, create columns or register an alias.
    `ask_fn(prompt, default) -> answer` overrides stdin (for tests / agent).
    """
    ask = ask_fn or _prompt
    # Decisions file is only auto-applied in --non-interactive mode.
    # Interactive runs always ask in the current terminal.
    preloaded = (
        load_decisions_file(decisions_path)
        if (decisions_path and not interactive)
        else {}
    )

    missing = collect_missing_equipment_fees(ra, update_path)
    if not missing:
        print("  No missing equipment cost columns detected.")
        return store

    present = {c.equipment for c in ra.cost_columns if c.equipment}
    print(f"  Missing equipment cost columns for: {', '.join(sorted(missing))}")

    for eq in sorted(missing):
        fees = sorted(missing[eq])
        group = equipment_alias_group(eq)
        rate_by = equipment_rate_by(eq)
        alias_peer = suggest_alias_equipment(eq, present)
        layout_donor = suggest_layout_template(eq, present)

        preload = preloaded.get(eq)
        if preload and not interactive:
            action = str(preload.get("action", "")).strip().lower()
            target = preload.get("target")
            target = str(target).strip().upper() if target else (
                alias_peer or layout_donor
            )
            print(f"  Using saved decision for {eq}: {action} -> {target}")
        elif interactive:
            print()
            print(f"Equipment {eq} has update rates but no RA cost column(s).")
            print(f"  Missing fees: {', '.join(fees)}")
            if rate_by:
                print(f"  Rate by: {rate_by}")
            print(f"  Alias group (same Rate by): {group}")
            if alias_peer:
                print(
                    f"  Same-type alias already on RA: {alias_peer} "
                    f"(you may alias instead of create)"
                )
                default_ans = "n"
            else:
                print(
                    "  No same-type alias on RA — should create this equipment."
                )
                if layout_donor:
                    print(
                        f"  Layout donor for cloning column structure: {layout_donor} "
                        f"(not an alias)"
                    )
                default_ans = "y"
            ans = ask(
                f"Create new cost columns for {eq}? (y=create / n=alias)",
                default_ans,
            ).lower()
            if ans.startswith("y"):
                action = "create"
                target = layout_donor
            else:
                action = "alias"
                target = ask(
                    f"Alias {eq} onto which existing equipment?",
                    alias_peer or layout_donor or "",
                ).upper()
        else:
            # Non-interactive, no preload: create solo types; alias when peer exists
            if alias_peer:
                action, target = "alias", alias_peer
                print(f"  Auto-alias {eq} -> {alias_peer} (same Rate-by group)")
            elif layout_donor:
                action, target = "create", layout_donor
                print(
                    f"  Auto-create {eq} (no same-type alias; layout from {layout_donor})"
                )
            else:
                print(f"  Skipping {eq} (non-interactive, no layout donor).")
                store.skipped.add(eq)
                store.decisions.append(
                    EquipmentDecision(eq, "skip", note="non-interactive, no donor")
                )
                continue

        if action == "create":
            if not target:
                print(f"  Cannot create {eq}: no layout donor equipment available.")
                store.skipped.add(eq)
                store.decisions.append(
                    EquipmentDecision(eq, "skip", note="no template")
                )
                continue
            created_fees: list[str] = []
            for fee in fees:
                # Already created earlier in this loop?
                exists = [
                    c
                    for c in ra.cost_columns
                    if c.fee_name == fee and c.equipment == eq
                ]
                if exists:
                    created_fees.append(fee)
                    continue
                tmpl = _template_cost_for_fee(ra, fee, target)
                if tmpl is None:
                    print(f"  Skip create {fee} ({eq}): no template {fee} ({target})")
                    continue
                tmpl_rate_by = ra.ws.cell(5, tmpl.start_col).value
                tmpl_rate_by_s = None
                if tmpl_rate_by:
                    # "Rate by: Container/20FT\nRegular rule"
                    for line in str(tmpl_rate_by).replace("\r", "").split("\n"):
                        if line.strip().lower().startswith("rate by:"):
                            tmpl_rate_by_s = line.split(":", 1)[1].strip()
                            break
                rate_by = rate_by_for_new_equipment(eq, tmpl_rate_by_s)
                new_c = create_cost_column_from_template(
                    ra,
                    fee_name=fee,
                    new_equipment=eq,
                    template=tmpl,
                    rate_by=rate_by,
                )
                created_fees.append(fee)
                print(f"  Created cost column {new_c.header!r} (Rate by: {rate_by})")
            store.created.add(eq)
            present.add(eq)
            store.decisions.append(
                EquipmentDecision(
                    eq,
                    "create",
                    target=target,
                    fees_created=created_fees,
                    note=f"cloned from {target}",
                )
            )
        elif action == "alias":
            if not target:
                print(f"  Cannot alias {eq}: empty target.")
                store.skipped.add(eq)
                continue
            if target not in present and not any(
                c.equipment == target for c in ra.cost_columns
            ):
                print(f"  Warning: alias target {target} not found on RA; storing anyway.")
            store.aliases[eq] = target
            print(f"  Aliasing {eq} -> {target} for this run")
            store.decisions.append(
                EquipmentDecision(eq, "alias", target=target, note="confirmed alias")
            )
        else:
            store.skipped.add(eq)
            store.decisions.append(EquipmentDecision(eq, "skip", note=f"action={action}"))

    if decisions_path:
        save_decisions_file(decisions_path, store)
        print(f"  Saved equipment decisions -> {decisions_path.name}")

    return store
