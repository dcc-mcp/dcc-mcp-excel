"""Host matrix + preflight tests (stdlib-only, platform independent)."""

from __future__ import annotations

from dcc_mcp_excel.host_matrix import (
    HOST_MATRIX,
    find_excel_executable,
    headless_available,
    preflight,
)


def test_matrix_has_a_verified_headless_entry() -> None:
    verified = [entry for entry in HOST_MATRIX if entry.status == "verified"]
    assert verified, "the host matrix must name at least one verified configuration"
    assert all(entry.platform == "any" for entry in verified)


def test_matrix_records_every_office_build() -> None:
    office = [entry for entry in HOST_MATRIX if entry.app == "Excel"]
    assert {entry.version for entry in office} >= {"Microsoft 365 (Office 16)", "2019 / 2021 (Office 16)", "2016 (Office 16)"}
    for entry in office:
        assert entry.status == "host_limited"
        assert entry.notes.strip()


def test_find_excel_executable_returns_path_or_none() -> None:
    found = find_excel_executable()
    assert found is None or found.is_file()


def test_preflight_reports_both_backends() -> None:
    status = preflight()
    assert status["app"] == "excel"
    assert set(status["checks"]) == {"headless_openxml", "desktop_excel"}
    assert isinstance(status["verified"], list) and status["verified"]
    assert isinstance(status["host_limited"], list) and status["host_limited"]


def test_preflight_degrades_without_office() -> None:
    """A machine without Excel lists every host_limited capability as degraded."""
    status = preflight()
    if status["checks"]["desktop_excel"]:
        assert status["degraded_without_office"] == []
    else:
        assert set(status["degraded_without_office"]) == set(status["host_limited"])


def test_preflight_ok_tracks_headless_only() -> None:
    """`ok` reflects the headless backend, so a Linux runner is not a failure."""
    status = preflight()
    assert status["ok"] is headless_available()
