"""Headless Open XML compiler — Workbook IR → XLSX (ADR 004).

Contract-first: the compiler is an *implementation* of the workbook
contract. It writes structure only — cells, formulas, named ranges, tables,
validation and conditional formatting. Native recomputation, chart refresh
and pivots require the Excel COM host and are reported as `host_limited`
rather than being faked (see `dcc_mcp_excel.capabilities`).

Nothing is silently dropped: a feature spec naming an unknown worksheet
raises `UnknownWorksheetError` instead of being skipped.

The dependency on openpyxl is opt-in exactly like python-pptx in
dcc-mcp-powerpoint: package import never pulls it, only this module does.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
)
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table as OpenpyxlTable
from openpyxl.worksheet.table import TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet as OpenpyxlWorksheet

from .workbook_io import XLSX_SUFFIX, sanitize_sheet_title
from .workbook_ir import (
    WorkbookEnvelope,
    WorkbookIr,
)

DEFAULT_SHEET_COLUMN_WIDTH = 12.0
_CALCULATION_MODES = {"auto", "manual"}


class ConditionalFormatError(ValueError):
    """A conditional-format spec has params the headless writer cannot apply."""


def _split_range(reference: str) -> str:
    """Strip a leading sheet qualifier and `$` anchors from a range."""
    _sheet, _sep, address = reference.rpartition("!")
    return (address or reference).replace("$", "")


def _write_rows(sheet: OpenpyxlWorksheet, rows: tuple[tuple[Any, ...], ...]) -> int:
    for row_index, row in enumerate(rows, start=1):
        for column_index, value in enumerate(row, start=1):
            if value is None:
                continue
            sheet.cell(row=row_index, column=column_index, value=value)
    return len(rows)


def _write_cells(
    document: WorkbookIr,
    ir_name: str,
    sheets_by_ir_name: dict[str, OpenpyxlWorksheet],
) -> None:
    """Write `document.cells` (addressed writes) once, after every sheet exists.

    Unqualified addresses land on the first worksheet; a `Sheet!A1` qualifier
    is resolved against the **IR** worksheet names so a title that had to be
    sanitized is still addressable.
    """
    for cell in document.cells:
        sheet_ref, _sep, address = cell.address.rpartition("!")
        target = sheets_by_ir_name[ir_name]
        if sheet_ref:
            name = sheet_ref.strip("'")
            if name not in sheets_by_ir_name:
                raise KeyError(f"cell '{cell.address}' names unknown worksheet '{sheet_ref}'")
            target = sheets_by_ir_name[name]
        handle = target[address.replace("$", "")]
        if cell.is_formula:
            handle.value = cell.formula
        else:
            handle.value = cell.value


def _write_formulas(sheet: OpenpyxlWorksheet, ir_name: str, document: WorkbookIr) -> None:
    for formula in document.formulas:
        if formula.worksheet != ir_name:
            continue
        sheet[formula.cell.replace("$", "")].value = formula.formula


def _write_validations(sheet: OpenpyxlWorksheet, ir_name: str, document: WorkbookIr) -> None:
    for spec in document.validations:
        if spec.worksheet != ir_name:
            continue
        validation = DataValidation(type=spec.kind, allow_blank=True)
        if spec.kind == "list":
            values = spec.params.get("values") or spec.params.get("formula1")
            if values is None:
                raise KeyError(f"validation on '{sheet.title}' needs params.values for a list rule")
            if isinstance(values, (list, tuple)):
                validation.formula1 = '"' + ",".join(str(v) for v in values) + '"'
            else:
                validation.formula1 = str(values)
        else:
            for key in ("operator", "formula1", "formula2"):
                if key in spec.params:
                    setattr(validation, key, spec.params[key])
        sheet.add_data_validation(validation)
        validation.add(_split_range(spec.range))


def _conditional_rule(spec: Any) -> Any:
    kind = spec.kind
    params = spec.params
    if kind == "cell_value":
        operator = params.get("operator", "equal")
        formula = params.get("formula")
        if formula is None:
            raise ConditionalFormatError("cell_value needs params.formula")
        return CellIsRule(operator=operator, formula=[formula], **_style_kwargs(params))
    if kind == "color_scale":
        colors = params.get("colors") or ["F8696B", "FFEB84", "63BE7B"]
        return ColorScaleRule(
            start_type=params.get("start_type", "min"),
            start_value=params.get("start_value"),
            start_color=colors[0],
            mid_type=params.get("mid_type", "percentile"),
            mid_value=params.get("mid_value", 50),
            mid_color=colors[1] if len(colors) > 2 else colors[-1],
            end_type=params.get("end_type", "max"),
            end_value=params.get("end_value"),
            end_color=colors[-1],
        )
    if kind == "data_bar":
        return DataBarRule(
            start_type=params.get("start_type", "num"),
            start_value=params.get("start_value", 0),
            end_type=params.get("end_type", "num"),
            end_value=params.get("end_value", 100),
            color=params.get("color", "638EC6"),
        )
    if kind == "formula":
        formula = params.get("formula")
        if formula is None:
            raise ConditionalFormatError("formula needs params.formula")
        return FormulaRule(formula=[formula], **_style_kwargs(params))
    raise ConditionalFormatError(f"unsupported conditional format kind '{kind}'")


def _style_kwargs(params: dict[str, Any]) -> dict[str, Any]:
    """Extract the openpyxl differential-style keys an IR rule may carry."""
    allowed = ("font", "border", "fill")
    return {key: params[key] for key in allowed if key in params}


def _write_conditional_formats(sheet: OpenpyxlWorksheet, ir_name: str, document: WorkbookIr) -> None:
    for spec in document.conditional_formats:
        if spec.worksheet != ir_name:
            continue
        rule = _conditional_rule(spec)
        sheet.conditional_formatting.add(_split_range(spec.range), rule)


def _write_tables(sheet: OpenpyxlWorksheet, ir_name: str, document: WorkbookIr) -> None:
    """Materialize the IR tables as real openpyxl tables on `sheet`.

    Tables were parsed and validated but never written, so the read-back gate
    now reports them as missing. An unnamed table gets a name derived from its
    range; openpyxl requires one.
    """
    for spec in document.tables:
        if spec.worksheet != ir_name:
            continue
        name = spec.name or f"Table{_split_range(spec.range).replace(':', '_')}"
        table = OpenpyxlTable(displayName=name, ref=_split_range(spec.range))
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False
        )
        sheet.add_table(table)


def _write_named_ranges(workbook: Workbook, document: WorkbookIr) -> None:
    for spec in document.named_ranges:
        if spec.name in workbook.defined_names:
            raise ValueError(f"duplicate named range '{spec.name}'")
        workbook.defined_names.add(DefinedName(spec.name, attr_text=spec.refers_to))


def _autosize_columns(sheet: OpenpyxlWorksheet, rows: tuple[tuple[Any, ...], ...]) -> None:
    widths: dict[int, int] = {}
    for row in rows:
        for index, value in enumerate(row, start=1):
            widths[index] = max(widths.get(index, 0), len(str(value)) if value is not None else 0)
    for index, width in widths.items():
        sheet.column_dimensions[get_column_letter(index)].width = max(DEFAULT_SHEET_COLUMN_WIDTH, float(min(width + 2, 60)))


class UnknownWorksheetError(ValueError):
    """A feature spec names a worksheet the IR does not define."""


class WorkbookCompiler:
    """Compiles a WorkbookEnvelope into an XLSX file via openpyxl."""

    def __init__(self, envelope: WorkbookEnvelope, workbook: Workbook | None = None) -> None:
        self.envelope = envelope
        if workbook is None:
            workbook = Workbook()
            workbook.remove(workbook.active)
        self.wb = workbook
        self.titles: set[str] = set()
        self.sheets_by_ir_name: dict[str, OpenpyxlWorksheet] = {}
        self.summary: dict[str, Any] = {"sheets": [], "rows": 0}

    def _require_known_worksheets(self) -> None:
        """Reject feature specs naming a worksheet the IR does not define.

        Titles are sanitized on the way into the artifact (`Shots 10/06`
        becomes `Shots 10-06`), so matching a spec against the compiled title
        silently skips it. Validate against the IR names up front instead: an
        unwritable spec is an error, never a dropped feature.
        """
        known = [worksheet.name for worksheet in self.envelope.document.worksheets]
        known_set = set(known)
        for label, specs in (
            ("table", self.envelope.document.tables),
            ("formula", self.envelope.document.formulas),
            ("validation", self.envelope.document.validations),
            ("conditional_format", self.envelope.document.conditional_formats),
            ("chart", self.envelope.document.charts),
            ("pivot", self.envelope.document.pivots),
        ):
            for spec in specs:
                if spec.worksheet not in known_set:
                    raise UnknownWorksheetError(
                        f"{label} names unknown worksheet '{spec.worksheet}'; "
                        f"IR defines: {', '.join(repr(name) for name in known)}"
                    )

    def _apply_metadata(self) -> None:
        metadata = self.envelope.metadata
        self.wb.properties.title = metadata.title
        self.wb.properties.creator = metadata.author or "DCC-MCP"
        self.wb.properties.subject = f"dcc-mcp-excel workbook ({self.envelope.document_id})"
        policy = self.envelope.document.calculation_policy
        if policy.full_calc_on_load:
            self.wb.calculation.fullCalcOnLoad = True

    def _compile_sheet(self, worksheet_ir: Any) -> None:
        title = sanitize_sheet_title(worksheet_ir.name, taken=self.titles)
        self.titles.add(title)
        sheet = self.wb.create_sheet(title=title)
        self.sheets_by_ir_name[worksheet_ir.name] = sheet
        rows_written = _write_rows(sheet, worksheet_ir.rows)
        _autosize_columns(sheet, worksheet_ir.rows)
        document = self.envelope.document
        _write_formulas(sheet, worksheet_ir.name, document)
        _write_validations(sheet, worksheet_ir.name, document)
        _write_conditional_formats(sheet, worksheet_ir.name, document)
        _write_tables(sheet, worksheet_ir.name, document)
        self.summary["sheets"].append({"name": title, "rows": rows_written})
        self.summary["rows"] += rows_written

    def compile(self, out_path: str | Path) -> Path:
        document = self.envelope.document
        self._require_known_worksheets()
        self._apply_metadata()
        for worksheet_ir in document.worksheets:
            self._compile_sheet(worksheet_ir)
        if document.cells:
            # `document.cells` is a document-level list, not a per-sheet one:
            # it is written exactly once, and only after every sheet exists so
            # a qualified address can target any of them.
            _write_cells(document, document.worksheets[0].name, self.sheets_by_ir_name)
        _write_named_ranges(self.wb, document)
        out = Path(out_path)
        if out.suffix.lower() != XLSX_SUFFIX:
            out = out.with_suffix(XLSX_SUFFIX)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(str(out))
        return out


def compile_workbook(envelope: WorkbookEnvelope, out_path: str | Path) -> Path:
    """Compile a WorkbookEnvelope to XLSX via the headless Open XML backend."""
    return WorkbookCompiler(envelope).compile(out_path)
