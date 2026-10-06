"""Read-back verification — the "write then read it back" 1.0 gate.

Compiling an XLSX proves a file was written; it does not prove the file
contains what the IR asked for. This module reopens the artifact with
openpyxl and compares it against the envelope cell by cell, then reports
structured features separately because Excel only materializes some of them
on open (see `capabilities`).
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .workbook_io import cell_values_match, sanitize_sheet_title
from .workbook_ir import WorkbookEnvelope

MISMATCH_LIMIT = 20


@dataclass(frozen=True)
class Mismatch:
    sheet: str
    address: str
    expected: Any
    actual: Any


@dataclass(frozen=True)
class ReadbackReport:
    path: str
    sheets: tuple[str, ...]
    checked_cells: int
    mismatches: tuple[Mismatch, ...]
    formulas_present: int
    named_ranges: tuple[str, ...]
    validations: int
    conditional_formats: int

    @property
    def ok(self) -> bool:
        return not self.mismatches

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "path": self.path,
            "sheets": list(self.sheets),
            "checked_cells": self.checked_cells,
            "mismatches": [
                {"sheet": m.sheet, "address": m.address, "expected": m.expected, "actual": m.actual}
                for m in self.mismatches
            ],
            "formulas_present": self.formulas_present,
            "named_ranges": list(self.named_ranges),
            "validations": self.validations,
            "conditional_formats": self.conditional_formats,
        }


def _address(row: int, column: int) -> str:
    return f"{get_column_letter(column)}{row}"


def read_back(envelope: WorkbookEnvelope, xlsx_path: str | Path) -> ReadbackReport:
    """Reopen `xlsx_path` and compare it against `envelope`."""
    path = Path(xlsx_path)
    if not path.is_file():
        raise FileNotFoundError(f"workbook artifact not found: {path}")
    workbook = load_workbook(filename=str(path), data_only=False)
    try:
        worksheet_irs = envelope.document.worksheets
        if len(workbook.sheetnames) != len(worksheet_irs):
            raise ValueError(
                f"sheet count mismatch: artifact has {len(workbook.sheetnames)}, IR has {len(worksheet_irs)}"
            )

        expected_titles = []
        taken: set[str] = set()
        for worksheet_ir in worksheet_irs:
            title = sanitize_sheet_title(worksheet_ir.name, taken=taken)
            taken.add(title)
            expected_titles.append(title)
        # Lengths are known equal from the check above, so a plain zip is
        # exact here. zip(strict=True) would say the same thing but is 3.10+,
        # and this package supports 3.9.
        for index, (expected, actual) in enumerate(zip(expected_titles, workbook.sheetnames), start=1):
            if expected != actual:
                raise ValueError(f"worksheet {index}: expected title '{expected}', artifact has '{actual}'")

        checked = 0
        mismatches: list[Mismatch] = []
        for worksheet_ir, title in zip(worksheet_irs, workbook.sheetnames):
            sheet = workbook[title]
            for row_index, row in enumerate(worksheet_ir.rows, start=1):
                for column_index, expected_value in enumerate(row, start=1):
                    checked += 1
                    actual_value = sheet.cell(row=row_index, column=column_index).value
                    if not cell_values_match(expected_value, actual_value):
                        mismatches.append(
                            Mismatch(
                                sheet=title,
                                address=_address(row_index, column_index),
                                expected=expected_value,
                                actual=actual_value,
                            )
                        )
                        if len(mismatches) >= MISMATCH_LIMIT:
                            break
                if len(mismatches) >= MISMATCH_LIMIT:
                    break

        formulas = 0
        for title in workbook.sheetnames:
            for row in workbook[title].iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formulas += 1

        validations = sum(len(workbook[title].data_validations.dataValidation) for title in workbook.sheetnames)
        conditional_formats = sum(len(workbook[title].conditional_formatting._cf_rules) for title in workbook.sheetnames)

        return ReadbackReport(
            path=str(path),
            sheets=tuple(workbook.sheetnames),
            checked_cells=checked,
            mismatches=tuple(mismatches),
            formulas_present=formulas,
            named_ranges=tuple(workbook.defined_names),
            validations=validations,
            conditional_formats=conditional_formats,
        )
    finally:
        workbook.close()
