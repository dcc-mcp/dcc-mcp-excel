"""Read-back verification tests — the 1.0 write-then-read-back gate."""

from __future__ import annotations

from pathlib import Path

import pytest

from dcc_mcp_excel.compiler import compile_workbook
from dcc_mcp_excel.readback import read_back
from dcc_mcp_excel.workbook_io import cell_values_match
from dcc_mcp_excel.workbook_ir import (
    Metadata,
    WorkbookEnvelope,
    WorkbookIr,
    Worksheet,
    parse_envelope,
)

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"


def _envelope(document: WorkbookIr, document_id: str = "draft:test") -> WorkbookEnvelope:
    return WorkbookEnvelope(
        schema_version="office-ir/1.0",
        kind="workbook",
        document_id=document_id,
        metadata=Metadata(title="Test"),
        document=document,
    )


def test_read_back_matches_every_cell(tmp_path: Path) -> None:
    envelope = _envelope(
        WorkbookIr(
            worksheets=(
                Worksheet(name="Shots", rows=(("shot", "status", "frames"), ("sh010", "ip", 120))),
                Worksheet(name="Summary", rows=(("total", 1),)),
            )
        )
    )
    out = compile_workbook(envelope, tmp_path / "rb.xlsx")
    report = read_back(envelope, out)
    assert report.ok, report.to_dict()
    # Every positioned cell in the IR is checked: 3x2 on Shots + 1x2 on Summary.
    assert report.checked_cells == 8
    assert report.sheets == ("Shots", "Summary")


def test_read_back_detects_mismatches(tmp_path: Path) -> None:
    """The gate must fail, not warn, when the artifact does not match the IR."""
    envelope = _envelope(WorkbookIr(worksheets=(Worksheet(name="S", rows=(("expected",),)),)))
    out = compile_workbook(envelope, tmp_path / "rb.xlsx")

    # Same artifact, different IR: one cell disagrees.
    altered = _envelope(WorkbookIr(worksheets=(Worksheet(name="S", rows=(("changed",),)),)))
    report = read_back(altered, out)
    assert not report.ok
    assert report.mismatches[0].address == "A1"
    assert report.mismatches[0].expected == "changed"
    assert report.mismatches[0].actual == "expected"


def test_read_back_rejects_sheet_count_mismatch(tmp_path: Path) -> None:
    envelope = _envelope(WorkbookIr(worksheets=(Worksheet(name="A", rows=(("x",),)),)))
    out = compile_workbook(envelope, tmp_path / "rb.xlsx")
    wider = _envelope(
        WorkbookIr(
            worksheets=(Worksheet(name="A", rows=(("x",),)), Worksheet(name="B", rows=(("y",),))),
        )
    )
    with pytest.raises(ValueError, match="sheet count mismatch"):
        read_back(wider, out)


def test_read_back_rejects_title_mismatch(tmp_path: Path) -> None:
    envelope = _envelope(WorkbookIr(worksheets=(Worksheet(name="Alpha", rows=(("x",),)),)))
    out = compile_workbook(envelope, tmp_path / "rb.xlsx")
    renamed = _envelope(WorkbookIr(worksheets=(Worksheet(name="Beta", rows=(("x",),)),)))
    with pytest.raises(ValueError, match="expected title"):
        read_back(renamed, out)


def test_read_back_detects_formulas_present(tmp_path: Path) -> None:
    envelope = parse_envelope(
        {
            "schema_version": "office-ir/1.0",
            "kind": "workbook",
            "document_id": "draft:f",
            "metadata": {"title": "f"},
            "document": {
                "worksheets": [{"name": "S", "rows": [["a", 2]]}],
                "formulas": [{"worksheet": "S", "cell": "C1", "formula": "=B1*2"}],
            },
        }
    )
    out = compile_workbook(envelope, tmp_path / "f.xlsx")
    report = read_back(envelope, out)
    assert report.ok
    # The formula text survives; its value does not (evaluation is host_limited).
    assert report.formulas_present == 1


def test_read_back_reports_features(tmp_path: Path) -> None:
    envelope = parse_envelope(
        {
            "schema_version": "office-ir/1.0",
            "kind": "workbook",
            "document_id": "draft:feat",
            "metadata": {"title": "feat"},
            "document": {
                "worksheets": [{"name": "S", "rows": [["n", 3]]}],
                "named_ranges": [{"name": "Total", "refers_to": "S!$A$1:$B$1"}],
                "validations": [{"worksheet": "S", "range": "A2:A9", "kind": "list", "params": {"values": ["a", "b"]}}],
                "conditional_formats": [
                    {"worksheet": "S", "range": "B2:B9", "kind": "cell_value", "params": {"formula": ["5"]}}
                ],
            },
        }
    )
    out = compile_workbook(envelope, tmp_path / "feat.xlsx")
    report = read_back(envelope, out)
    assert report.ok
    assert report.named_ranges == ("Total",)
    assert report.validations == 1
    assert report.conditional_formats == 1


def test_read_back_missing_file_raises(tmp_path: Path) -> None:
    envelope = _envelope(WorkbookIr(worksheets=(Worksheet(name="S"),)))
    with pytest.raises(FileNotFoundError):
        read_back(envelope, tmp_path / "absent.xlsx")


def test_read_back_example_workbook(tmp_path: Path) -> None:
    from dcc_mcp_excel.workbook_ir import load_workbook_ir

    envelope = load_workbook_ir(EXAMPLES / "shot_list.json")
    out = compile_workbook(envelope, tmp_path / "example.xlsx")
    report = read_back(envelope, out)
    assert report.ok, report.to_dict()


@pytest.mark.parametrize(
    "expected,actual,match",
    [
        (None, None, True),
        ("a", "a", True),
        ("a", " a ", True),
        ("a", "b", False),
        (1, 1, True),
        (1, 1.0000000001, True),
        (1, 2, False),
        (True, True, True),
        (True, 1, False),
        (1, True, False),
        ("1", 1, False),
    ],
)
def test_cell_values_match(expected, actual, match) -> None:
    assert cell_values_match(expected, actual) is match
