"""office-host client tests — stdlib-only, no host binary required."""

from __future__ import annotations

from pathlib import Path

from dcc_mcp_excel.host_client import (
    HOST_EXE,
    _abs,
    _matching_response,
    compile_workbook,
    find_host_binary,
    handshake,
    rpc,
)


def test_abs_returns_absolute_paths() -> None:
    resolved = _abs("relative/ir.json")
    assert Path(resolved).is_absolute()


def test_matching_response_selects_by_id() -> None:
    stdout = "\n".join(
        [
            json_line("other", {"ok": False}),
            json_line("req", {"ok": True}),
        ]
    )
    assert _matching_response(stdout, "req")["result"] == {"ok": True}


def test_matching_response_rejects_duplicates() -> None:
    stdout = "\n".join([json_line("req", {"a": 1}), json_line("req", {"a": 2})])
    try:
        _matching_response(stdout, "req")
    except ValueError as exc:
        assert "2 responses" in str(exc)
    else:  # pragma: no cover - the call above must raise
        raise AssertionError("duplicate responses must be rejected")


def test_matching_response_rejects_invalid_json() -> None:
    try:
        _matching_response("not json\n", "req")
    except ValueError as exc:
        assert "invalid JSON" in str(exc)
    else:  # pragma: no cover
        raise AssertionError("invalid JSON must be rejected")


def json_line(identifier: str, result: dict) -> str:
    import json

    return json.dumps({"jsonrpc": "2.0", "id": identifier, "result": result})


def test_rpc_reports_missing_host_without_raising(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("dcc_mcp_excel.host_client.find_host_binary", lambda: None)
    outcome = rpc("office.host.ping", {})
    assert outcome["success"] is False
    assert outcome["backend"] is None
    assert HOST_EXE in outcome["reason"]
    assert "OFFICE_HOST_NOT_FOUND" in outcome["reason"]


def test_find_host_binary_prefers_env(monkeypatch, tmp_path: Path) -> None:
    fake = tmp_path / HOST_EXE
    fake.write_bytes(b"")
    monkeypatch.setenv("DCC_OFFICE_HOST", str(fake))
    assert find_host_binary() == str(fake)


def test_client_helpers_report_missing_host(monkeypatch) -> None:
    monkeypatch.setattr("dcc_mcp_excel.host_client.find_host_binary", lambda: None)
    assert handshake()["success"] is False
    assert compile_workbook("ir.json", "out.xlsx")["success"] is False
