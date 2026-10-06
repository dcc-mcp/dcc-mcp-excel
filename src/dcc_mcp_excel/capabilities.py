"""Capability grading — the honest boundary of the v0.1.0 surface (PIP-4276).

Two grades, and the difference is evidence, not confidence:

- `verified` — covered by a CI-green headless test on a Linux runner. No
  Office installation, no COM, no Windows-only path.
- `host_limited` — requires the desktop Excel application (COM). The
  artifact is structurally valid without it, but the value only materializes
  when Excel opens it. Never reported as verified.

Anything not listed here is unimplemented, not silently degraded: the
compiler raises on specs it cannot write rather than dropping them.
"""

from __future__ import annotations

from dataclasses import dataclass

VERIFIED = "verified"
HOST_LIMITED = "host_limited"
UNIMPLEMENTED = "unimplemented"


@dataclass(frozen=True)
class Capability:
    name: str
    grade: str
    summary: str
    evidence: str
    requires_office: bool

    @property
    def is_verified(self) -> bool:
        return self.grade == VERIFIED


CAPABILITIES: tuple[Capability, ...] = (
    Capability(
        name="workbook.compile",
        grade=VERIFIED,
        summary="Workbook IR (office-ir/1.0) → XLSX through the headless Open XML backend.",
        evidence="tests/test_compiler.py + tests/test_readback.py, ubuntu-latest, no Office installed.",
        requires_office=False,
    ),
    Capability(
        name="workbook.read_back",
        grade=VERIFIED,
        summary="Reopen a compiled XLSX and compare every cell against the source IR.",
        evidence="tests/test_readback.py, ubuntu-latest; the 1.0 write-then-read-back gate.",
        requires_office=False,
    ),
    Capability(
        name="workbook.rows.write",
        grade=VERIFIED,
        summary="Positional grid writes (text / number / boolean / empty cells).",
        evidence="tests/test_compiler.py::test_compile_writes_rows_and_values.",
        requires_office=False,
    ),
    Capability(
        name="workbook.formulas.write",
        grade=VERIFIED,
        summary="Formula text is written to the sheet; the formula string round-trips.",
        evidence="tests/test_readback.py::test_read_back_detects_formulas_present.",
        requires_office=False,
    ),
    Capability(
        name="workbook.named_ranges",
        grade=VERIFIED,
        summary="Defined names are written and read back.",
        evidence="tests/test_compiler.py::test_compile_writes_named_ranges.",
        requires_office=False,
    ),
    Capability(
        name="workbook.validation",
        grade=VERIFIED,
        summary="Data-validation rules (list/whole/decimal/date/custom) are written to the sheet.",
        evidence="tests/test_compiler.py::test_compile_writes_list_validation.",
        requires_office=False,
    ),
    Capability(
        name="workbook.conditional_format",
        grade=VERIFIED,
        summary="cell_value / color_scale / data_bar / formula rules are written to the sheet.",
        evidence="tests/test_compiler.py::test_compile_writes_conditional_formats.",
        requires_office=False,
    ),
    Capability(
        name="excel.workbook.calculate",
        grade=HOST_LIMITED,
        summary="Native formula recalculation. Cached values are produced by Excel, not by openpyxl.",
        evidence="Not covered by CI. xlwings backend is not wired in v0.1.0.",
        requires_office=True,
    ),
    Capability(
        name="excel.chart.generate",
        grade=HOST_LIMITED,
        summary="Native chart objects and their refresh on open.",
        evidence="Not covered by CI. IR charts are parsed and validated, not rendered in v0.1.0.",
        requires_office=True,
    ),
    Capability(
        name="excel.pivot.refresh",
        grade=HOST_LIMITED,
        summary="Pivot table creation and refresh.",
        evidence="Not covered by CI. IR pivots are parsed and validated, not built in v0.1.0.",
        requires_office=True,
    ),
    Capability(
        name="excel.pdf.export",
        grade=HOST_LIMITED,
        summary="High-fidelity PDF export including pagination.",
        evidence="Not covered by CI; requires the desktop Excel COM backend.",
        requires_office=True,
    ),
    Capability(
        name="excel.graph.session",
        grade=UNIMPLEMENTED,
        summary="Microsoft Graph Workbook sessions (proposal §6.3).",
        evidence="Out of scope for v0.1.0: needs Azure app registration and tenant consent.",
        requires_office=False,
    ),
)

_BY_NAME = {capability.name: capability for capability in CAPABILITIES}


def get(name: str) -> Capability | None:
    return _BY_NAME.get(name)


def by_grade(grade: str) -> tuple[Capability, ...]:
    return tuple(c for c in CAPABILITIES if c.grade == grade)


def report() -> dict[str, object]:
    """Machine-readable grading summary, exposed through the skill and CLI."""
    return {
        "schema": "dcc-mcp-capability-grade/1",
        "verified": [c.name for c in by_grade(VERIFIED)],
        "host_limited": [c.name for c in by_grade(HOST_LIMITED)],
        "unimplemented": [c.name for c in by_grade(UNIMPLEMENTED)],
        "capabilities": [
            {
                "name": c.name,
                "grade": c.grade,
                "summary": c.summary,
                "evidence": c.evidence,
                "requires_office": c.requires_office,
            }
            for c in CAPABILITIES
        ],
    }
