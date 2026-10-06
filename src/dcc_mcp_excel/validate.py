"""Structural validation for workbooks and artifacts (proposal §17/§18.1).

Two layers, kept separate on purpose:

- `validate_envelope` — contract + structural checks on the Workbook IR,
  before any file is produced.
- `validate_artifacts` — produced files exist and are non-empty.

Honest reporting: every check returns pass/fail with a reason, and
host-limited features are reported as declared-but-unverified rather than
checked and passed.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .capabilities import HOST_LIMITED
from .workbook_ir import DOCUMENT_KIND, IR_VERSION, WorkbookEnvelope

MAX_SHEET_NAME_LEN = 31


@dataclass(frozen=True)
class Check:
    name: str
    ok: bool
    detail: str = ""


def validate_envelope(envelope: WorkbookEnvelope) -> dict[str, Any]:
    """Validate a WorkbookEnvelope; returns {ok, checks, warnings}."""
    checks: list[Check] = []
    warnings: list[str] = []

    checks.append(Check("schema_version", envelope.schema_version == IR_VERSION, envelope.schema_version))
    checks.append(Check("document_kind", envelope.kind == DOCUMENT_KIND, envelope.kind))
    sheets = envelope.document.worksheets
    checks.append(Check("has_worksheets", len(sheets) > 0, f"{len(sheets)} worksheets"))
    checks.append(Check("has_title", bool(envelope.metadata.title.strip()), envelope.metadata.title))

    for sheet in sheets:
        label = f"worksheet '{sheet.name}'"
        if len(sheet.name) > MAX_SHEET_NAME_LEN:
            checks.append(
                Check(
                    f"{label}.name_length",
                    False,
                    f"{len(sheet.name)} chars exceeds Excel's {MAX_SHEET_NAME_LEN}-char limit",
                )
            )
        if not sheet.rows:
            warnings.append(f"{label}: no rows; the sheet will be created empty")
        widths = {len(row) for row in sheet.rows}
        if len(widths) > 1:
            warnings.append(f"{label}: ragged rows (widths {sorted(widths)}); short rows leave trailing cells empty")

    for spec in envelope.document.charts:
        warnings.append(
            f"chart on '{spec.worksheet}' ({spec.kind}): declared in the IR and validated, "
            f"but native rendering is {HOST_LIMITED} — open the workbook in Excel to materialize it"
        )
    for spec in envelope.document.pivots:
        warnings.append(
            f"pivot '{spec.name}' on '{spec.worksheet}': declared in the IR and validated, "
            f"but pivot construction is {HOST_LIMITED}"
        )
    if envelope.document.calculation_policy.mode == "manual":
        warnings.append("calculation_policy.mode='manual': formulas stay unevaluated until Excel recalculates")

    ok = all(c.ok for c in checks)
    return {"ok": ok, "checks": [c.__dict__ for c in checks], "warnings": warnings}


def validate_artifacts(paths: list[str | Path]) -> dict[str, Any]:
    """Check produced artifacts exist and are non-empty."""
    results = []
    for raw in paths:
        p = Path(raw)
        results.append({"path": str(p), "ok": p.is_file() and p.stat().st_size > 0})
    return {"ok": all(r["ok"] for r in results), "artifacts": results}
