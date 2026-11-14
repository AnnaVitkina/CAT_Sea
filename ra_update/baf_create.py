"""Interactive stopper: choose until which BAF monthly cost column to create."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd

from .charge_rules import resolve_rate_column
from .cost_map import period_fee_header
from .dates import format_ra_date, parse_date
from .equipment_decisions import update_row_targets_ra
from .parsing import to_number
from .period_fees import iter_baf_months_covering


BAF_DECISIONS_KEY = "_baf_create_until"


@dataclass
class BafCreateStore:
    """
    Per-equipment create cutoff for new BAF Fee monthly columns.

    create_until[eq] = date  -> create months with valid_to <= date
    create_until[eq] = None  -> create none (only update existing)
    missing key              -> create all (no stopper applied yet)
    """

    create_until: dict[str, date | None] = field(default_factory=dict)
    proposed: dict[str, list[tuple[date, date]]] = field(default_factory=dict)
    decision_notes: dict[str, str] = field(default_factory=dict)

    def has_decision(self, equipment: str) -> bool:
        return equipment.upper() in self.create_until

    def allows_create(self, equipment: str, month_to: date) -> bool:
        eq = equipment.upper()
        if eq not in self.create_until:
            return True
        until = self.create_until[eq]
        if until is None:
            return False
        return month_to <= until

    def note_for(self, equipment: str) -> str | None:
        return self.decision_notes.get(equipment.upper())


def _tty_prompt(text: str, default: str = "") -> str:
    """Use input() so Colab / Jupyter work (stdin often isatty()=False)."""
    import sys

    suffix = f" [{default}]" if default else ""
    prompt = f"{text}{suffix}: "
    try:
        raw = input(prompt).strip()
    except EOFError:
        raise SystemExit(
            "Interactive BAF create prompts need usable input(). "
            "In Colab use `%run run_pipeline.py` (not !python), "
            "or pass --non-interactive to create all missing BAF months."
        ) from None
    except Exception:
        print(prompt, end="", flush=True)
        try:
            raw = sys.stdin.readline()
            if raw == "":
                raise EOFError("empty stdin")
            raw = raw.strip()
        except EOFError:
            raise SystemExit(
                "Interactive BAF create prompts need usable input(). "
                "In Colab use `%run run_pipeline.py` (not !python), "
                "or pass --non-interactive to create all missing BAF months."
            ) from None
    return raw if raw else default


def collect_missing_baf_months(
    ra,
    update_path: Path,
    *,
    aliases: dict[str, str] | None = None,
) -> dict[str, list[tuple[date, date]]]:
    """
    equipment -> sorted unique standard BAF months that updates need but RA lacks.
    """
    aliases = {k.upper(): v.upper() for k, v in (aliases or {}).items()}
    needed: dict[str, set[tuple[date, date]]] = {}

    df = pd.read_excel(update_path, sheet_name="ETSBAF")
    # DFT BAF catalog is created automatically (not via interactive 15→14 cutoff).
    ignored_eq: set[str] = set()
    for _, row in df.iterrows():
        applies = update_row_targets_ra(ra, row, "ETSBAF")
        vf = parse_date(row.get("RATE_EFFECTIVE_DATE__C"))
        vt = parse_date(row.get("RATE_EXPIRATION_DATE__C"))
        if vf is None or vt is None:
            continue
        for col, raw in row.items():
            col_s = str(col)
            if not col_s.endswith("_RATE"):
                continue
            if to_number(raw) is None:
                continue
            rule, equipment, skip = resolve_rate_column(col_s, "ETSBAF")
            if skip or rule is None or rule.ra_fee_name != "BAF Fee":
                continue
            eq = (rule.ra_equipment or equipment or "").strip().upper()
            if not eq:
                continue
            eq = aliases.get(eq, eq)
            if eq == "DFT":
                continue  # DFT periods created up front in process_sheet
            if not applies:
                ignored_eq.add(eq)
                continue
            for m_from, m_to in iter_baf_months_covering(vf, vt):
                headers = [period_fee_header("BAF Fee", eq, m_from, m_to)]
                if any(c.header in headers for c in ra.cost_columns):
                    continue
                needed.setdefault(eq, set()).add((m_from, m_to))

    unused = sorted(eq for eq in ignored_eq if eq not in needed)
    if unused:
        print(
            "  Skipping BAF create prompt for "
            f"{', '.join(unused)} "
            "(only on out-of-scope or non-matching KEY rows)."
        )
    return {eq: sorted(months) for eq, months in sorted(needed.items())}


def load_baf_create_until(path: Path) -> dict[str, date | None]:
    """
    From equipment_decisions.json:
      "_baf_create_until": {"42P1": "14.02.2027", "42G0": null}
    null / "" / "none" => create none.
    """
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        return {}
    raw = data.get(BAF_DECISIONS_KEY)
    if not isinstance(raw, dict):
        return {}
    out: dict[str, date | None] = {}
    for eq, val in raw.items():
        key = str(eq).strip().upper()
        if val is None or str(val).strip().lower() in {"", "none", "0"}:
            out[key] = None
            continue
        parsed = parse_date(val)
        if parsed is None:
            continue
        out[key] = parsed
    return out


def save_baf_create_until(path: Path, store: BafCreateStore) -> None:
    """Merge _baf_create_until into the decisions JSON (preserve other keys)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    data: dict = {}
    if path.exists():
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                data = loaded
        except json.JSONDecodeError:
            data = {}
    payload = {}
    for eq, until in store.create_until.items():
        payload[eq] = format_ra_date(until) if until is not None else None
    data[BAF_DECISIONS_KEY] = payload
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def _prompt_cutoff(
    equipment: str,
    months: list[tuple[date, date]],
) -> date | None:
    """
    Ask until which proposed month to create.
    Returns month end date inclusive, or None to create none.
    """
    print()
    print("-" * 72)
    print(f"BAF Fee ({equipment}): missing monthly cost columns that would be created:")
    for i, (m_from, m_to) in enumerate(months, start=1):
        print(
            f"  [{i}] BAF Fee ({equipment} "
            f"{format_ra_date(m_from)}-{format_ra_date(m_to)})"
        )
    n = len(months)
    print(
        f"Create through which month? Enter 1-{n} "
        f"(or 0 = create none of these)"
    )
    raw = _tty_prompt("Choice", default=str(n))
    try:
        choice = int(raw)
    except ValueError:
        choice = n
    if choice <= 0:
        return None
    if choice > n:
        choice = n
    return months[choice - 1][1]


def resolve_baf_create_cutoffs(
    ra,
    update_path: Path,
    store: BafCreateStore,
    *,
    decisions_path: Path | None = None,
    interactive: bool = True,
    aliases: dict[str, str] | None = None,
) -> BafCreateStore:
    """
    Pre-scan missing BAF months; prompt create-until per equipment each run.

    Cutoffs are not preloaded or persisted — enter them every interactive run.
    Non-interactive: create all missing months (no decisions-file preload).
    """
    _ = decisions_path  # kept for call-site compatibility; cutoffs never load/save

    missing = collect_missing_baf_months(ra, update_path, aliases=aliases)
    if not missing:
        print("  No missing BAF monthly cost columns to create.")
        return store

    for eq, months in missing.items():
        store.proposed[eq] = list(months)
        proposed_txt = ", ".join(
            f"{format_ra_date(a)}-{format_ra_date(b)}" for a, b in months
        )

        if not interactive:
            store.create_until[eq] = months[-1][1]
            store.decision_notes[eq] = _decision_note(
                eq, months[-1][1], months, months, [], source="non-interactive all"
            )
            print(
                f"  BAF Fee ({eq}): non-interactive — creating all {len(months)} "
                f"missing month(s): {proposed_txt}"
            )
            continue

        until = _prompt_cutoff(eq, months)
        store.create_until[eq] = until
        if until is None:
            creating, skipped = [], months
        else:
            creating = [(a, b) for a, b in months if b <= until]
            skipped = [(a, b) for a, b in months if b > until]
        store.decision_notes[eq] = _decision_note(
            eq, until, months, creating, skipped, source="user choice"
        )
        if until is None:
            print(f"  -> will not create new BAF Fee ({eq}) months")
        else:
            print(
                f"  -> create through {format_ra_date(until)} "
                f"({len(creating)} create, {len(skipped)} skip)"
            )

    return store


def _decision_note(
    equipment: str,
    until: date | None,
    proposed: list[tuple[date, date]],
    creating: list[tuple[date, date]],
    skipped: list[tuple[date, date]],
    *,
    source: str,
) -> str:
    prop = ", ".join(f"{format_ra_date(a)}-{format_ra_date(b)}" for a, b in proposed)
    cre = ", ".join(f"{format_ra_date(a)}-{format_ra_date(b)}" for a, b in creating) or "(none)"
    sk = ", ".join(f"{format_ra_date(a)}-{format_ra_date(b)}" for a, b in skipped) or "(none)"
    until_txt = format_ra_date(until) if until is not None else "none"
    return (
        f"BAF create until {until_txt} for {equipment} ({source}); "
        f"proposed=[{prop}]; creating=[{cre}]; skipped=[{sk}]"
    )
