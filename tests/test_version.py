"""Package-level smoke tests (no Excel installation required)."""

from __future__ import annotations

import pytest

from dcc_mcp_excel import __version__
from dcc_mcp_excel._standalone_entry import main
from dcc_mcp_excel.capabilities import HOST_LIMITED, VERIFIED, by_grade, report


def test_version_is_pep440() -> None:
    parts = __version__.split(".")
    assert len(parts) == 3
    assert all(p.isdigit() for p in parts)


def test_cli_version_matches_package_version(capsys) -> None:
    with pytest.raises(SystemExit) as exit_info:
        main(["dcc-mcp-excel", "--version"])
    assert exit_info.value.code == 0
    assert capsys.readouterr().out.strip() == f"dcc-mcp-excel {__version__}"


def test_capability_report_shape() -> None:
    body = report()
    assert body["schema"] == "dcc-mcp-capability-grade/1"
    assert "workbook.compile" in body["verified"]
    assert "excel.workbook.calculate" in body["host_limited"]
    assert "excel.graph.session" in body["unimplemented"]
    # A capability is never listed under two grades.
    grades = [*body["verified"], *body["host_limited"], *body["unimplemented"]]
    assert len(grades) == len(set(grades))


def test_host_limited_capabilities_are_not_verified() -> None:
    """The PIP-4276 rule: no host_limited capability may claim verification."""
    for capability in by_grade(HOST_LIMITED):
        assert capability.requires_office
        assert capability.grade != VERIFIED


def test_every_capability_carries_evidence() -> None:
    for capability in by_grade(VERIFIED) + by_grade(HOST_LIMITED):
        assert capability.evidence.strip(), f"{capability.name} has no evidence line"
