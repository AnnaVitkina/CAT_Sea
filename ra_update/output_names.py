"""Readable output filenames from RA + update workbook names."""

from __future__ import annotations

import re
from pathlib import Path

# Windows-forbidden filename characters
_BAD = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
_RA_VER = re.compile(r"(RA\d+\s+v\.\d+)", re.IGNORECASE)


def short_ra_label(path: Path | str) -> str:
    """
    'Advanced Export - RA20260513093 v.15 - Sea.xlsx' → 'RA20260513093 v.15'
    """
    stem = Path(path).stem.strip()
    m = _RA_VER.search(stem)
    if m:
        return m.group(1)
    # Fallback: strip common Advanced Export wrappers
    text = re.sub(r"^Advanced Export\s*-\s*", "", stem, flags=re.I)
    text = re.sub(r"\s*-\s*Sea(?:\s*FR)?\s*$", "", text, flags=re.I)
    return text.strip() or stem


def short_update_label(path: Path | str) -> str:
    """'Additional lanes UK.xlsx' → 'Additional lanes UK'"""
    return Path(path).stem.strip()


def pair_label(ra_path: Path | str, update_path: Path | str) -> str:
    """
    'RA20260513093 v.15_ Additional lanes UK'
    (space after underscore, as requested).
    """
    return f"{short_ra_label(ra_path)}_ {short_update_label(update_path)}"


def safe_filename_stem(label: str, *, max_len: int = 120) -> str:
    """Make a label safe as a file stem (keep spaces; drop illegal chars)."""
    text = _BAD.sub(" ", str(label)).strip()
    text = re.sub(r"\s+", " ", text)
    if len(text) > max_len:
        text = text[:max_len].rstrip(" ._")
    return text or "output"


def pair_stem(ra_path: Path | str, update_path: Path | str) -> str:
    return safe_filename_stem(pair_label(ra_path, update_path))


def stem_from_baseline_and_candidate(
    baseline_path: Path | str,
    candidate_path: Path | str,
    *,
    update_path: Path | str | None = None,
) -> str:
    """
    Prefer explicit update_path; else parse '{ra}_ {update}_updated…' from
    candidate name; else baseline RA + candidate stem.
    """
    if update_path is not None:
        return pair_stem(baseline_path, update_path)

    cand = Path(candidate_path)
    ra = short_ra_label(baseline_path)
    # e.g. 'RA20260513093 v.15_ Additional lanes UK_updated.xlsx'
    m = re.match(
        re.escape(ra) + r"_ (.+?)_updated(?:_\d{8}_\d{6})?$",
        cand.stem,
        flags=re.I,
    )
    if m:
        return safe_filename_stem(f"{ra}_ {m.group(1)}")

    # Legacy RA_UPDATED_* → still use RA + candidate stem
    return safe_filename_stem(f"{ra}_ {cand.stem}")
