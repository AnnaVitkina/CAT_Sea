"""Interactive / non-interactive selection of input Excel files."""

from __future__ import annotations

from pathlib import Path


def list_xlsx(folder: Path) -> list[Path]:
    """Sorted *.xlsx in folder, skipping Excel lock files (~$)."""
    if not folder.is_dir():
        return []
    return sorted(
        p
        for p in folder.glob("*.xlsx")
        if p.is_file() and not p.name.startswith("~$")
    )


def ensure_input_folders(*folders: Path) -> None:
    for folder in folders:
        folder.mkdir(parents=True, exist_ok=True)


def choose_xlsx(
    label: str,
    folder: Path,
    *,
    preferred_name: str | None = None,
) -> Path:
    """
    Prompt the user to pick an .xlsx from `folder`.
    If preferred_name exists, it is offered as the default (Enter).
    """
    ensure_input_folders(folder)
    files = list_xlsx(folder)
    if not files:
        raise FileNotFoundError(
            f"No .xlsx files in {folder}. Place the {label} file there and retry."
        )

    default_idx = 0
    if preferred_name:
        for i, p in enumerate(files):
            if p.name == preferred_name:
                default_idx = i
                break

    print()
    print(f"Select {label} from: {folder}")
    for i, p in enumerate(files, start=1):
        mark = " (default)" if i - 1 == default_idx else ""
        print(f"  [{i}] {p.name}{mark}")

    while True:
        raw = input(f"Choice [1-{len(files)}] (Enter={default_idx + 1}): ").strip()
        if not raw:
            return files[default_idx]
        if raw.isdigit():
            n = int(raw)
            if 1 <= n <= len(files):
                return files[n - 1]
        print(f"  Enter a number 1-{len(files)}, or press Enter for default.")


def resolve_xlsx(
    label: str,
    folder: Path,
    explicit: Path | None,
    *,
    interactive: bool,
    preferred_name: str | None = None,
) -> Path:
    """
    Resolve an input workbook path.

    - explicit CLI path wins
    - interactive → numbered menu from folder
    - non-interactive → preferred_name if present, else sole file in folder
    """
    if explicit is not None:
        path = explicit.expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(path)
        return path

    ensure_input_folders(folder)
    files = list_xlsx(folder)

    if interactive:
        return choose_xlsx(label, folder, preferred_name=preferred_name)

    if preferred_name:
        preferred = folder / preferred_name
        if preferred.exists():
            return preferred.resolve()

    if len(files) == 1:
        return files[0].resolve()

    if not files:
        raise FileNotFoundError(
            f"No .xlsx in {folder}. Pass -- path or add a {label} file."
        )
    names = ", ".join(p.name for p in files)
    raise FileNotFoundError(
        f"Multiple {label} files in {folder} ({names}). "
        f"Pass an explicit path or run interactively."
    )
