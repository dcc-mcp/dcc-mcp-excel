"""Structural validation tests."""

from __future__ import annotations

from pathlib import Path

from dcc_mcp_excel.validate import validate_artifacts, validate_envelope
from dcc_mcp_excel.workbook_ir import (
    Metadata,
    WorkbookEnvelope,
    WorkbookIr,
    Worksheet,
    parse_envelope,
)

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"


def _envelope(document: WorkbookIr, title: str = "Test") -> WorkbookEnvelope:
    return WorkbookEnvelope(
        schema_version="office-ir/1.0",
        kind="workbook",
        document_id="draft:test",
        metadata=Metadata(title=title),
        document=document,
    )


def test_minimal_envelope_passes() -> None:
    report = validate_envelope(_envelope(WorkbookIr(worksheets=(Worksheet(name="S", rows=(("a",),)),))))
    assert report["ok"]


def test_empty_title_fails() -> None:
    report = validate_envelope(_envelope(WorkbookIr(worksheets=(Worksheet(name="S"),)), title="  "))
    assert not report["ok"]
    assert any(c["name"] == "has_title" for c in report["checks"])


def test_long_sheet_name_fails() -> None:
    report = validate_envelope(_envelope(WorkbookIr(worksheets=(Worksheet(name="X" * 40),))))
    assert not report["ok"]
    assert any("name_length" in c["name"] for c in report["checks"])


def test_empty_sheet_warns() -> None:
    report = validate_envelope(_envelope(WorkbookIr(worksheets=(Worksheet(name="S", rows=()),))))
    assert report["ok"]
    assert any("no rows" in w for w in report["warnings"])


def test_ragged_rows_warn() -> None:
    report = validate_envelope(
        _envelope(WorkbookIr(worksheets=(Worksheet(name="S", rows=(("a", "b"), ("c",))),)))
    )
    assert report["ok"]
    assert any("ragged rows" in w for w in report["warnings"])


def test_host_limited_features_warn_instead_of_passing() -> None:
    """Charts and pivots are reported as unverified, never as checked-green."""
    document = parse_envelope(
        {
            "schema_version": "office-ir/1.0",
            "kind": "workbook",
            "document_id": "draft:x",
            "metadata": {"title": "x"},
            "document": {
                "worksheets": [{"name": "S", "rows": [["a", 1]]}],
                "charts": [{"worksheet": "S", "kind": "bar", "data_range": "A1:B1"}],
                "pivots": [{"worksheet": "S", "name": "P", "source": "A1:B1"}],
            },
        }
    ).document
    report = validate_envelope(_envelope(document))
    assert report["ok"]
    joined = " ".join(report["warnings"])
    assert "host_limited" in joined
    assert "chart" in joined and "pivot" in joined
    # No fabricated check claims the chart was verified.
    assert all("chart" not in c["name"] for c in report["checks"])


def test_manual_calculation_warns() -> None:
    document = parse_envelope(
        {
            "schema_version": "office-ir/1.0",
            "kind": "workbook",
            "document_id": "draft:x",
            "metadata": {"title": "x"},
            "document": {
                "worksheets": [{"name": "S", "rows": [["a", 1]]}],
                "calculation_policy": {"mode": "manual"},
            },
        }
    ).document
    report = validate_envelope(_envelope(document))
    assert any("manual" in w for w in report["warnings"])


def test_example_workbook_validates() -> None:
    from dcc_mcp_excel.workbook_ir import load_workbook_ir

    report = validate_envelope(load_workbook_ir(EXAMPLES / "shot_list.json"))
    assert report["ok"], report


def test_validate_artifacts(tmp_path: Path) -> None:
    good = tmp_path / "a.xlsx"
    good.write_bytes(b"PK\x03\x04")
    empty = tmp_path / "b.xlsx"
    empty.write_bytes(b"")
    report = validate_artifacts([str(good), str(empty), str(tmp_path / "missing.xlsx")])
    assert not report["ok"]
    assert [r["ok"] for r in report["artifacts"]] == [True, False, False]
