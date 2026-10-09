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

from .workbook_io import cell_values_match, sanitize_sheet_title, table_display_name
from .workbook_ir import WorkbookEnvelope

MISMATCH_LIMIT = 20


@dataclass(frozen=True)
class Mismatch:
    sheet: str
    address: str
    expected: Any
    actual: Any


@dataclass(frozen=True)
class FeatureMismatch:
    """A structured feature the IR asked for that the artifact does not carry."""

    feature: str
    sheet: str
    reference: str
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
    feature_mismatches: tuple[FeatureMismatch, ...] = ()

    @property
    def ok(self) -> bool:
        return not self.mismatches and not self.feature_mismatches

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
            "feature_mismatches": [
                {
                    "feature": f.feature,
                    "sheet": f.sheet,
                    "reference": f.reference,
                    "expected": f.expected,
                    "actual": f.actual,
                }
                for f in self.feature_mismatches
            ],
        }


def _address(row: int, column: int) -> str:
    return f"{get_column_letter(column)}{row}"


def _normalize_range(reference: str) -> str:
    """Fold an A1 reference to a comparable form (no `$`, no sheet qualifier)."""
    _sheet, _sep, address = reference.rpartition("!")
    return (address or reference).replace("$", "").upper()


def _cf_ranges(sheet: Any) -> set[str]:
    """Every conditional-formatting range on `sheet`, normalized."""
    return {str(entry.sqref).replace("$", "").upper() for entry in sheet.conditional_formatting._cf_rules}


def _validation_ranges(sheet: Any) -> set[str]:
    """Every data-validation range on `sheet`, normalized."""
    ranges: set[str] = set()
    for validation in sheet.data_validations.dataValidation:
        ranges.update(_normalize_range(str(ref)) for ref in validation.sqref.ranges)
    return ranges


def _check_features(
    envelope: WorkbookEnvelope,
    workbook: Any,
    ir_name_to_title: dict[str, str],
) -> tuple[FeatureMismatch, ...]:
    """Compare the structured features the IR declared against the artifact.

    Counting them is not verification: a feature dropped during compile leaves
    the count lower and the gate still green. Every declared feature must be
    found on its own sheet at its own reference.
    """
    document = envelope.document
    mismatches: list[FeatureMismatch] = []

    for spec in document.formulas:
        sheet = workbook[ir_name_to_title[spec.worksheet]]
        expected = spec.formula
        actual = sheet[_normalize_range(spec.cell)].value
        if actual != expected:
            mismatches.append(
                FeatureMismatch("formula", spec.worksheet, spec.cell, expected, actual)
            )

    for spec in document.validations:
        sheet = workbook[ir_name_to_title[spec.worksheet]]
        if _normalize_range(spec.range) not in _validation_ranges(sheet):
            mismatches.append(
                FeatureMismatch(
                    "validation", spec.worksheet, spec.range, spec.range, sorted(_validation_ranges(sheet))
                )
            )

    for spec in document.conditional_formats:
        sheet = workbook[ir_name_to_title[spec.worksheet]]
        if _normalize_range(spec.range) not in _cf_ranges(sheet):
            mismatches.append(
                FeatureMismatch(
                    "conditional_format",
                    spec.worksheet,
                    spec.range,
                    spec.range,
                    sorted(_cf_ranges(sheet)),
                )
            )

    for spec in document.tables:
        sheet = workbook[ir_name_to_title[spec.worksheet]]
        titles = {str(name).lower() for name in sheet.tables}
        if table_display_name(spec).lower() not in titles:
            mismatches.append(
                FeatureMismatch(
                    "table", spec.worksheet, spec.range, table_display_name(spec), sorted(titles)
                )
            )

    for spec in document.named_ranges:
        if spec.name not in workbook.defined_names:
            mismatches.append(
                FeatureMismatch("named_range", "", spec.refers_to, spec.name, list(workbook.defined_names))
            )

    return tuple(mismatches)


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
        # Feature specs address sheets by their IR name, which may differ from
        # the compiled title once sanitization kicks in.
        ir_name_to_title = {ws.name: title for ws, title in zip(worksheet_irs, expected_titles)}
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
            feature_mismatches=_check_features(envelope, workbook, ir_name_to_title),
        )
    finally:
        workbook.close()
