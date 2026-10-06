"""Workbook IR contract — mirrors the dcc-mcp-office-ir workbook schema 1:1.

Contract-first: this module is the *domain core* of the Excel adapter. The
compiler (headless Open XML implementation) and the COM renderer both
consume this contract, never raw coordinates.

JSON shape (snake_case, office-ir/1.0), see the Rust crate
dcc-mcp-office-ir for the authoritative schema:

    {
      "schema_version": "office-ir/1.0",
      "kind": "workbook",
      "document_id": "draft:shot-list",
      "metadata": {"title": "Shot list", "author": "...", "language": "en"},
      "document": {
        "worksheets": [
          {"name": "Shots", "rows": [["shot", "status"], ["sh010", "ip"]]}
        ],
        "tables": [{"worksheet": "Shots", "range": "A1:B2"}],
        "named_ranges": [],
        "formulas": [],
        "validations": [],
        "conditional_formats": [],
        "charts": [],
        "pivots": [],
        "calculation_policy": {"mode": "auto", "full_calc_on_load": false}
      },
      "outputs": ["xlsx"]
    }
"""

from __future__ import annotations

import json
import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Union

IR_VERSION: str = "office-ir/1.0"
DOCUMENT_KIND: str = "workbook"

# A1-style cell or range reference, e.g. "B7" or "A1:C9". An optional
# sheet-qualified form ("Summary!$A$1") is accepted for named ranges. The row
# group excludes 0 because Excel rows are 1-based: "A0" is not a cell.
_A1_CELL_RE = re.compile(r"^(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)?!?\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$")
_A1_RANGE_RE = re.compile(
    r"^(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)(?::(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*))?$"
)


class IrValidationError(ValueError):
    """A Workbook IR document violated the contract. Carries a json-path hint."""

    def __init__(self, path: str, message: str) -> None:
        super().__init__(f"[{path}] {message}")
        self.path = path
        self.message = message


@dataclass(frozen=True)
class TemplateRef:
    uri: str
    version: str


@dataclass(frozen=True)
class Resource:
    id: str
    uri: str
    mime: str | None = None


@dataclass(frozen=True)
class Metadata:
    title: str
    author: str = ""
    language: str = "en"


@dataclass(frozen=True)
class Cell:
    """One populated cell: A1 address plus a validated literal or formula."""

    address: str
    value: str | float | bool | None = None
    formula: str | None = None

    @property
    def is_formula(self) -> bool:
        return self.formula is not None


@dataclass(frozen=True)
class TableSpec:
    worksheet: str
    range: str
    name: str | None = None


@dataclass(frozen=True)
class NamedRange:
    name: str
    refers_to: str


@dataclass(frozen=True)
class FormulaSpec:
    worksheet: str
    cell: str
    formula: str


@dataclass(frozen=True)
class ValidationSpec:
    worksheet: str
    range: str
    kind: str
    params: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ConditionalFormatSpec:
    worksheet: str
    range: str
    kind: str
    params: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ChartSpec:
    worksheet: str
    kind: str
    data_range: str
    title: str | None = None


@dataclass(frozen=True)
class PivotSpec:
    worksheet: str
    name: str
    source: str


@dataclass(frozen=True)
class CalculationPolicy:
    mode: str = "auto"
    full_calc_on_load: bool = False


@dataclass(frozen=True)
class Worksheet:
    name: str
    rows: tuple[tuple[str | float | bool | None, ...], ...] = ()


@dataclass(frozen=True)
class WorkbookIr:
    """Workbook document payload: grid data plus structured features.

    Cells are parsed first because every row is a positional list; features
    (tables, formulas, ...) address the grid by A1 reference and are kept
    verbatim so a later COM/headless backend can apply them.
    """

    worksheets: tuple[Worksheet, ...]
    tables: tuple[TableSpec, ...] = ()
    named_ranges: tuple[NamedRange, ...] = ()
    formulas: tuple[FormulaSpec, ...] = ()
    validations: tuple[ValidationSpec, ...] = ()
    conditional_formats: tuple[ConditionalFormatSpec, ...] = ()
    charts: tuple[ChartSpec, ...] = ()
    pivots: tuple[PivotSpec, ...] = ()
    calculation_policy: CalculationPolicy = field(default_factory=CalculationPolicy)
    cells: tuple[Cell, ...] = ()


@dataclass(frozen=True)
class WorkbookEnvelope:
    schema_version: str
    kind: str
    document_id: str
    metadata: Metadata
    document: WorkbookIr
    template: TemplateRef | None = None
    resources: tuple[Resource, ...] = ()
    outputs: tuple[str, ...] = ("xlsx",)


_FEATURE_KINDS: dict[str, frozenset[str]] = {
    "validation": frozenset({"list", "whole", "decimal", "date", "custom"}),
    "conditional_format": frozenset({"cell_value", "color_scale", "data_bar", "formula"}),
    "chart": frozenset({"bar", "line", "pie", "scatter", "column"}),
}

CALCULATION_MODES = frozenset({"auto", "manual"})

CELL_LITERALS = (str, int, float, bool)

# A cell holds a literal or nothing. Written as a runtime Optional[Union[...]]
# rather than a PEP 604 union because this is an assignment, not an
# annotation: `from __future__ import annotations` defers annotations only, so
# `str | None` here would raise TypeError on the 3.9 floor.
_CELL_KIND = Optional[Union[str, float, bool]]


def _require(mapping: dict[str, Any], key: str, path: str) -> Any:
    if key not in mapping:
        raise IrValidationError(path, f"missing required key '{key}'")
    return mapping[key]


def _require_str(mapping: dict[str, Any], key: str, path: str) -> str:
    value = _require(mapping, key, path)
    if not isinstance(value, str):
        raise IrValidationError(path, f"'{key}' must be a string, got {type(value).__name__}")
    return value


def _require_a1(value: Any, path: str, key: str) -> str:
    if not isinstance(value, str) or not _A1_RANGE_RE.match(value.strip()):
        raise IrValidationError(path, f"'{key}' must be an A1-style reference, got {value!r}")
    return value.strip()


def _require_a1_cell(value: Any, path: str, key: str) -> str:
    if not isinstance(value, str) or not _A1_CELL_RE.match(value.strip()):
        raise IrValidationError(path, f"'{key}' must be an A1-style cell address, got {value!r}")
    return value.strip()


def parse_cell_value(raw: Any, path: str) -> _CELL_KIND:
    """Normalize one literal cell: string, number, bool or empty."""
    if raw is None:
        return None
    if isinstance(raw, bool):
        return raw
    if isinstance(raw, (int, float)):
        # NaN and infinity have no xlsx representation; JSON can carry both
        # through non-strict parsers, so reject them at the contract edge.
        if not math.isfinite(raw):
            raise IrValidationError(path, f"{raw!r} is not a representable cell value")
        return raw
    if isinstance(raw, str):
        return raw
    raise IrValidationError(path, f"cell value must be string/number/bool/null, got {type(raw).__name__}")


def parse_worksheet(raw: Any, index: int) -> Worksheet:
    path = f"document.worksheets[{index}]"
    if not isinstance(raw, dict):
        raise IrValidationError(path, "worksheet must be an object")
    name = _require_str(raw, "name", path)
    if not name.strip():
        raise IrValidationError(path, "'name' must be a non-empty string")
    rows_raw = raw.get("rows", [])
    if not isinstance(rows_raw, list):
        raise IrValidationError(f"{path}.rows", "'rows' must be a list of lists")
    rows: list[tuple[_CELL_KIND, ...]] = []
    for row_index, row in enumerate(rows_raw):
        if not isinstance(row, list):
            raise IrValidationError(f"{path}.rows[{row_index}]", "row must be a list")
        rows.append(tuple(parse_cell_value(v, f"{path}.rows[{row_index}][{c}]") for c, v in enumerate(row)))
    return Worksheet(name=name, rows=tuple(rows))


def parse_cells(raw: Any) -> tuple[Cell, ...]:
    """Parse `document.cells`: [{address, value|formula}]."""
    if raw is None:
        return ()
    if not isinstance(raw, list):
        raise IrValidationError("document.cells", "'cells' must be a list")
    cells: list[Cell] = []
    for index, entry in enumerate(raw):
        path = f"document.cells[{index}]"
        if not isinstance(entry, dict):
            raise IrValidationError(path, "cell must be an object")
        address = _require_a1_cell(entry.get("address"), path, "address")
        formula = entry.get("formula")
        value = entry.get("value")
        if formula is not None:
            if not isinstance(formula, str) or not formula.startswith("="):
                raise IrValidationError(path, "'formula' must be a string starting with '='")
            cells.append(Cell(address=address, formula=formula))
            continue
        if "value" not in entry:
            raise IrValidationError(path, "cell needs either 'value' or 'formula'")
        cells.append(Cell(address=address, value=parse_cell_value(value, path)))
    return tuple(cells)


def parse_features(raw: dict[str, Any]) -> dict[str, Any]:
    """Parse the structured feature lists that overlay the cell grid."""
    path = "document"

    def entries(key: str) -> list[Any]:
        value = raw.get(key, [])
        if not isinstance(value, list):
            raise IrValidationError(f"{path}.{key}", f"'{key}' must be a list")
        return value

    tables: list[TableSpec] = []
    for i, item in enumerate(entries("tables")):
        ipath = f"{path}.tables[{i}]"
        if not isinstance(item, dict):
            raise IrValidationError(ipath, "table must be an object")
        tables.append(
            TableSpec(
                worksheet=_require_str(item, "worksheet", ipath),
                range=_require_a1(item.get("range"), ipath, "range"),
                name=item.get("name"),
            )
        )

    named_ranges: list[NamedRange] = []
    for i, item in enumerate(entries("named_ranges")):
        ipath = f"{path}.named_ranges[{i}]"
        if not isinstance(item, dict):
            raise IrValidationError(ipath, "named range must be an object")
        named_ranges.append(
            NamedRange(
                name=_require_str(item, "name", ipath),
                refers_to=_require_a1(item.get("refers_to"), ipath, "refers_to"),
            )
        )

    formulas: list[FormulaSpec] = []
    for i, item in enumerate(entries("formulas")):
        ipath = f"{path}.formulas[{i}]"
        if not isinstance(item, dict):
            raise IrValidationError(ipath, "formula must be an object")
        formula = _require_str(item, "formula", ipath)
        if not formula.startswith("="):
            raise IrValidationError(ipath, "'formula' must start with '='")
        formulas.append(
            FormulaSpec(
                worksheet=_require_str(item, "worksheet", ipath),
                cell=_require_a1_cell(item.get("cell"), ipath, "cell"),
                formula=formula,
            )
        )

    def kinded(key: str, feature: str, factory: Any) -> list[Any]:
        result: list[Any] = []
        for i, item in enumerate(entries(key)):
            ipath = f"{path}.{key}[{i}]"
            if not isinstance(item, dict):
                raise IrValidationError(ipath, f"{feature} must be an object")
            kind = _require_str(item, "kind", ipath)
            if kind not in _FEATURE_KINDS[feature]:
                known = ", ".join(sorted(_FEATURE_KINDS[feature]))
                raise IrValidationError(ipath, f"unknown {feature} kind '{kind}'; known: {known}")
            params = item.get("params", {})
            if not isinstance(params, dict):
                raise IrValidationError(ipath, "'params' must be an object")
            result.append(
                factory(
                    worksheet=_require_str(item, "worksheet", ipath),
                    range=_require_a1(item.get("range"), ipath, "range"),
                    kind=kind,
                    params=dict(params),
                )
            )
        return result

    validations = kinded("validations", "validation", ValidationSpec)
    conditional_formats = kinded("conditional_formats", "conditional_format", ConditionalFormatSpec)

    charts: list[ChartSpec] = []
    for i, item in enumerate(entries("charts")):
        ipath = f"{path}.charts[{i}]"
        if not isinstance(item, dict):
            raise IrValidationError(ipath, "chart must be an object")
        kind = _require_str(item, "kind", ipath)
        if kind not in _FEATURE_KINDS["chart"]:
            known = ", ".join(sorted(_FEATURE_KINDS["chart"]))
            raise IrValidationError(ipath, f"unknown chart kind '{kind}'; known: {known}")
        charts.append(
            ChartSpec(
                worksheet=_require_str(item, "worksheet", ipath),
                kind=kind,
                data_range=_require_a1(item.get("data_range"), ipath, "data_range"),
                title=item.get("title"),
            )
        )

    pivots: list[PivotSpec] = []
    for i, item in enumerate(entries("pivots")):
        ipath = f"{path}.pivots[{i}]"
        if not isinstance(item, dict):
            raise IrValidationError(ipath, "pivot must be an object")
        pivots.append(
            PivotSpec(
                worksheet=_require_str(item, "worksheet", ipath),
                name=_require_str(item, "name", ipath),
                source=_require_str(item, "source", ipath),
            )
        )

    policy_raw = raw.get("calculation_policy", {})
    if not isinstance(policy_raw, dict):
        raise IrValidationError(f"{path}.calculation_policy", "must be an object")
    mode = str(policy_raw.get("mode", "auto"))
    if mode not in CALCULATION_MODES:
        raise IrValidationError(f"{path}.calculation_policy.mode", f"expected auto|manual, got '{mode}'")
    policy = CalculationPolicy(mode=mode, full_calc_on_load=bool(policy_raw.get("full_calc_on_load", False)))

    return {
        "tables": tuple(tables),
        "named_ranges": tuple(named_ranges),
        "formulas": tuple(formulas),
        "validations": tuple(validations),
        "conditional_formats": tuple(conditional_formats),
        "charts": tuple(charts),
        "pivots": tuple(pivots),
        "calculation_policy": policy,
    }


def parse_workbook(raw: Any) -> WorkbookIr:
    path = "document"
    if not isinstance(raw, dict):
        raise IrValidationError(path, "must be an object")
    worksheets_raw = raw.get("worksheets")
    if not isinstance(worksheets_raw, list) or not worksheets_raw:
        raise IrValidationError(f"{path}.worksheets", "must be a non-empty list")
    worksheets = tuple(parse_worksheet(ws, i) for i, ws in enumerate(worksheets_raw))
    names = [ws.name for ws in worksheets]
    duplicates = {n for n in names if names.count(n) > 1}
    if duplicates:
        raise IrValidationError(f"{path}.worksheets", f"duplicate worksheet names: {sorted(duplicates)}")
    return WorkbookIr(worksheets=worksheets, cells=parse_cells(raw.get("cells")), **parse_features(raw))


def parse_envelope(raw: Any) -> WorkbookEnvelope:
    if not isinstance(raw, dict):
        raise IrValidationError("$", "envelope must be an object")
    version = _require_str(raw, "schema_version", "$")
    if version != IR_VERSION:
        raise IrValidationError("$.schema_version", f"expected '{IR_VERSION}', got '{version}'")
    kind = _require_str(raw, "kind", "$")
    if kind != DOCUMENT_KIND:
        raise IrValidationError("$.kind", f"expected '{DOCUMENT_KIND}', got '{kind}'")
    metadata_raw = _require(raw, "metadata", "$")
    if not isinstance(metadata_raw, dict):
        raise IrValidationError("$.metadata", "must be an object")
    template_raw = raw.get("template")
    template = None
    if template_raw is not None:
        if not isinstance(template_raw, dict):
            raise IrValidationError("$.template", "must be an object")
        template = TemplateRef(
            uri=_require_str(template_raw, "uri", "$.template"),
            version=str(template_raw.get("version", "0.0.0")),
        )
    return WorkbookEnvelope(
        schema_version=version,
        kind=kind,
        document_id=_require_str(raw, "document_id", "$"),
        metadata=Metadata(
            title=_require_str(metadata_raw, "title", "$.metadata"),
            author=str(metadata_raw.get("author", "")),
            language=str(metadata_raw.get("language", "en")),
        ),
        template=template,
        resources=tuple(
            Resource(id=str(r.get("id", i)), uri=str(r.get("uri", ""))) for i, r in enumerate(raw.get("resources", []))
        ),
        document=parse_workbook(_require(raw, "document", "$")),
        outputs=tuple(str(o) for o in raw.get("outputs", ["xlsx"])),
    )


def artifact_stem(document_id: str) -> str:
    """Safe filesystem stem for a document id (ids may contain ':')."""
    return re.sub(r"[^A-Za-z0-9._-]+", "-", document_id).strip("-") or "workbook"


def load_workbook_ir(source: str | Path | dict[str, Any]) -> WorkbookEnvelope:
    """Load and validate a Workbook IR document from a JSON file or a mapping."""
    if isinstance(source, dict):
        return parse_envelope(source)
    path = Path(source)
    if not path.is_file():
        raise IrValidationError("$", f"input file not found: {path}")
    try:
        with path.open("r", encoding="utf-8") as handle:
            raw = json.load(handle)
    except json.JSONDecodeError as exc:
        raise IrValidationError("$", f"invalid JSON: {exc}") from exc
    return parse_envelope(raw)
