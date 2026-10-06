"""Workbook IR contract tests — pure, no Excel, no filesystem."""

from __future__ import annotations

import pytest

from dcc_mcp_excel.workbook_ir import (
    IR_VERSION,
    IrValidationError,
    WorkbookEnvelope,
    WorkbookIr,
    Worksheet,
    load_workbook_ir,
    parse_envelope,
)


def _envelope(**overrides) -> WorkbookEnvelope:
    document = overrides.pop("document", None) or WorkbookIr(worksheets=(Worksheet(name="Shots", rows=(("shot", "sh010"),)),))
    base = {
        "schema_version": IR_VERSION,
        "kind": "workbook",
        "document_id": "draft:shot-list",
        "metadata": {"title": "Shot list"},
        "document": document,
    }
    base.update(overrides)
    return WorkbookEnvelope(**base)  # type: ignore[arg-type]


def test_minimal_envelope_parses() -> None:
    raw = {
        "schema_version": "office-ir/1.0",
        "kind": "workbook",
        "document_id": "draft:shot-list",
        "metadata": {"title": "Shot list"},
        "document": {"worksheets": [{"name": "Shots", "rows": [["shot", "status"], ["sh010", "ip"]]}]},
    }
    envelope = parse_envelope(raw)
    assert envelope.document.worksheets[0].name == "Shots"
    assert envelope.document.worksheets[0].rows == (("shot", "status"), ("sh010", "ip"))
    assert envelope.outputs == ("xlsx",)


@pytest.mark.parametrize(
    "mutate",
    [
        {"schema_version": "office-ir/2.0"},
        {"kind": "presentation"},
        {"metadata": {"author": "no title"}},
        {"document": {"worksheets": []}},
    ],
)
def test_envelope_rejects_contract_violations(mutate) -> None:
    raw = {
        "schema_version": "office-ir/1.0",
        "kind": "workbook",
        "document_id": "draft:x",
        "metadata": {"title": "x"},
        "document": {"worksheets": [{"name": "S", "rows": []}]},
    }
    raw.update(mutate)
    with pytest.raises(IrValidationError):
        parse_envelope(raw)


def test_missing_key_reports_json_path() -> None:
    with pytest.raises(IrValidationError, match=r"\[\$\] missing required key 'document_id'"):
        parse_envelope(
            {
                "schema_version": "office-ir/1.0",
                "kind": "workbook",
                "metadata": {"title": "x"},
                "document": {"worksheets": [{"name": "S"}]},
            }
        )


def test_duplicate_worksheet_names_rejected() -> None:
    with pytest.raises(IrValidationError, match="duplicate worksheet names"):
        parse_envelope(
            {
                "schema_version": "office-ir/1.0",
                "kind": "workbook",
                "document_id": "draft:x",
                "metadata": {"title": "x"},
                "document": {"worksheets": [{"name": "S", "rows": []}, {"name": "S", "rows": []}]},
            }
        )


def test_worksheet_rows_accept_nulls_and_types() -> None:
    envelope = parse_envelope(
        {
            "schema_version": "office-ir/1.0",
            "kind": "workbook",
            "document_id": "draft:x",
            "metadata": {"title": "x"},
            "document": {"worksheets": [{"name": "S", "rows": [["t", 1, 2.5, True, None]]}]},
        }
    )
    assert envelope.document.worksheets[0].rows[0] == ("t", 1, 2.5, True, None)


def test_worksheet_row_must_be_a_list() -> None:
    with pytest.raises(IrValidationError, match=r"rows\[0\]"):
        parse_envelope(
            {
                "schema_version": "office-ir/1.0",
                "kind": "workbook",
                "document_id": "draft:x",
                "metadata": {"title": "x"},
                "document": {"worksheets": [{"name": "S", "rows": ["not-a-row"]}]},
            }
        )


@pytest.mark.parametrize("bad", [float("nan"), float("inf"), float("-inf")])
def test_non_finite_numbers_rejected(bad) -> None:
    """NaN/infinity have no xlsx representation; reject them at the contract."""
    with pytest.raises(IrValidationError, match="not a representable cell value"):
        parse_envelope(
            {
                "schema_version": "office-ir/1.0",
                "kind": "workbook",
                "document_id": "draft:x",
                "metadata": {"title": "x"},
                "document": {"worksheets": [{"name": "S", "rows": [[bad]]}]},
            }
        )


def test_bad_cell_value_type_rejected() -> None:
    with pytest.raises(IrValidationError, match="cell value must be"):
        parse_envelope(
            {
                "schema_version": "office-ir/1.0",
                "kind": "workbook",
                "document_id": "draft:x",
                "metadata": {"title": "x"},
                "document": {"worksheets": [{"name": "S", "rows": [[{"nested": 1}]]}]},
            }
        )


def test_unknown_feature_kind_rejected() -> None:
    for key, kind in (("validations", "bananas"), ("conditional_formats", "sparkle"), ("charts", "treemap")):
        with pytest.raises(IrValidationError, match=f"unknown .* kind '{kind}'"):
            parse_envelope(
                {
                    "schema_version": "office-ir/1.0",
                    "kind": "workbook",
                    "document_id": "draft:x",
                    "metadata": {"title": "x"},
                    "document": {
                        "worksheets": [{"name": "S", "rows": []}],
                        key: [{"worksheet": "S", "range": "A1:A2", "kind": kind}],
                    },
                }
            )


@pytest.mark.parametrize("bad", ["not-a-ref", "1A", "A0", ""])
def test_malformed_a1_references_rejected(bad) -> None:
    with pytest.raises(IrValidationError, match="A1-style"):
        parse_envelope(
            {
                "schema_version": "office-ir/1.0",
                "kind": "workbook",
                "document_id": "draft:x",
                "metadata": {"title": "x"},
                "document": {
                    "worksheets": [{"name": "S", "rows": []}],
                    "tables": [{"worksheet": "S", "range": bad}],
                },
            }
        )


def test_sheet_qualified_range_accepted() -> None:
    envelope = parse_envelope(
        {
            "schema_version": "office-ir/1.0",
            "kind": "workbook",
            "document_id": "draft:x",
            "metadata": {"title": "x"},
            "document": {
                "worksheets": [{"name": "Summary", "rows": []}],
                "named_ranges": [{"name": "Total", "refers_to": "Summary!$A$1:$B$9"}],
            },
        }
    )
    assert envelope.document.named_ranges[0].refers_to == "Summary!$A$1:$B$9"


def test_formula_must_start_with_equals() -> None:
    with pytest.raises(IrValidationError, match="must start with"):
        parse_envelope(
            {
                "schema_version": "office-ir/1.0",
                "kind": "workbook",
                "document_id": "draft:x",
                "metadata": {"title": "x"},
                "document": {
                    "worksheets": [{"name": "S", "rows": []}],
                    "formulas": [{"worksheet": "S", "cell": "C2", "formula": "SUM(A1)"}],
                },
            }
        )


def test_addressed_cells_parse_value_and_formula() -> None:
    envelope = parse_envelope(
        {
            "schema_version": "office-ir/1.0",
            "kind": "workbook",
            "document_id": "draft:x",
            "metadata": {"title": "x"},
            "document": {
                "worksheets": [{"name": "S", "rows": []}],
                "cells": [{"address": "B2", "value": 42}, {"address": "C2", "formula": "=B2*2"}],
            },
        }
    )
    cells = envelope.document.cells
    assert cells[0].value == 42 and not cells[0].is_formula
    assert cells[1].formula == "=B2*2" and cells[1].is_formula


def test_addressed_cell_needs_value_or_formula() -> None:
    with pytest.raises(IrValidationError, match="needs either 'value' or 'formula'"):
        parse_envelope(
            {
                "schema_version": "office-ir/1.0",
                "kind": "workbook",
                "document_id": "draft:x",
                "metadata": {"title": "x"},
                "document": {"worksheets": [{"name": "S", "rows": []}], "cells": [{"address": "B2"}]},
            }
        )


def test_calculation_policy_mode_is_constrained() -> None:
    with pytest.raises(IrValidationError, match="expected auto|manual"):
        parse_envelope(
            {
                "schema_version": "office-ir/1.0",
                "kind": "workbook",
                "document_id": "draft:x",
                "metadata": {"title": "x"},
                "document": {
                    "worksheets": [{"name": "S", "rows": []}],
                    "calculation_policy": {"mode": "sometimes"},
                },
            }
        )


def test_load_reports_missing_file(tmp_path) -> None:
    with pytest.raises(IrValidationError, match="input file not found"):
        load_workbook_ir(tmp_path / "nope.json")


def test_load_reports_invalid_json(tmp_path) -> None:
    bad = tmp_path / "bad.json"
    bad.write_text("{not json", encoding="utf-8")
    with pytest.raises(IrValidationError, match="invalid JSON"):
        load_workbook_ir(bad)


def test_artifact_stem_sanitizes_document_id() -> None:
    from dcc_mcp_excel.workbook_ir import artifact_stem

    assert artifact_stem("draft:shot-list") == "draft-shot-list"
    assert artifact_stem("///") == "workbook"
