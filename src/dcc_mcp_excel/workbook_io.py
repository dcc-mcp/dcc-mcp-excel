"""Reusable helpers for writing and reading workbook artifacts.

Shared by the headless compiler and the read-back verifier so both sides
agree on what "the same value" means: text is compared stripped, numbers
with a tolerance (openpyxl round-trips floats through the XML text layer).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from .workbook_ir import WorkbookEnvelope, WorkbookIr

XLSX_SUFFIX = ".xlsx"
# Excel rejects worksheet titles containing : \ / ? * [ ], and caps them at 31
# characters. Compiled names are sanitized so a document id can never produce
# an unopenable workbook.
_INVALID_TITLE_CHARS = re.compile(r"[\\/*?:\[\]]")
MAX_TITLE_LEN = 31
FLOAT_TOLERANCE = 1e-9


def sanitize_sheet_title(name: str, *, taken: set[str] | None = None) -> str:
    """Return an Excel-legal, unique worksheet title."""
    cleaned = _INVALID_TITLE_CHARS.sub("-", name).strip("'") or "Sheet"
    cleaned = cleaned[:MAX_TITLE_LEN]
    used = taken or set()
    if cleaned not in used:
        return cleaned
    prefix = cleaned[: MAX_TITLE_LEN - 4]
    for index in range(2, 1000):
        candidate = f"{prefix}-{index}"
        if candidate not in used:
            return candidate
    raise ValueError(f"could not derive a unique sheet title from {name!r}")


def resolve_output_path(envelope: WorkbookEnvelope, out_dir: str | Path, *, suffix: str = XLSX_SUFFIX) -> Path:
    """Artifact path for an envelope: `<out_dir>/<safe-stem><suffix>`."""
    from .workbook_ir import artifact_stem

    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    return out / f"{artifact_stem(envelope.document_id)}{suffix}"


def table_display_name(spec: Any) -> str:
    """The openpyxl table name an IR table spec is written under.

    `TableSpec.name` is optional, so the compiler and the read-back gate must
    derive the same fallback from the same place: two rules here would make
    the gate report a correctly-written unnamed table as missing.
    """
    name = getattr(spec, "name", None)
    if name:
        return name
    reference = str(getattr(spec, "range", ""))
    _sheet, _sep, address = reference.rpartition("!")
    return "Table" + (address or reference).replace("$", "").replace(":", "_")


def cell_values_match(expected: Any, actual: Any) -> bool:
    """Compare one expected IR value against a value read back from xlsx."""
    if expected is None:
        return actual is None
    if isinstance(expected, bool):
        return actual is expected
    if isinstance(expected, (int, float)):
        if isinstance(actual, bool) or not isinstance(actual, (int, float)):
            return False
        return abs(float(expected) - float(actual)) <= FLOAT_TOLERANCE
    if isinstance(expected, str):
        return isinstance(actual, str) and expected.strip() == actual.strip()
    return expected == actual


def parse_a1(reference: str) -> tuple[str | None, str]:
    """Split an A1 reference into (sheet-qualifier, address)."""
    sheet, _, address = reference.rpartition("!")
    return (sheet or None, address or reference)


def worksheet_row_count(document: WorkbookIr) -> int:
    """Total populated rows across every worksheet (validation/reporting)."""
    return sum(len(ws.rows) for ws in document.worksheets)
