"""Client for the Rust `dcc-mcp-office` Excel sidecar (ADR 004).

The sidecar owns the headless xlsx implementation shared with the rest of
the Office family; this adapter is the thin application layer on top. When
the sidecar binary is absent, `rpc()` reports that explicitly instead of
falling back to a second local writer — one implementation, one drift
surface.

Resolution order: `DCC_OFFICE_HOST` → `$ORIGIN/lib/dcc-office-host.exe`
(PyOxidizer layout) → PATH. Mirrors `dcc_mcp_powerpoint.host_client`.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

HOST_EXE = "dcc-office-host.exe"
OFFICE_HOST_ENV = "DCC_OFFICE_HOST"
DEFAULT_APP = "excel"
DEFAULT_TIMEOUT_SEC = 300


def find_host_binary() -> str | None:
    """Locate the office-host executable (see module docstring for the order)."""
    override = os.environ.get(OFFICE_HOST_ENV)
    if override and Path(override).is_file():
        return override
    exe_dir = Path(sys.executable).resolve().parent
    bundled = exe_dir / "lib" / HOST_EXE
    if bundled.is_file():
        return str(bundled)
    return shutil.which(HOST_EXE)


def _rpc_work_dir() -> Path:
    """Per-call work directory for file-based stdio redirection.

    Plain mkdir rather than tempfile.mkdtemp: some confined environments
    (agent sandboxes) deny the mkdtemp/chmod paths while allowing ordinary
    directory creation. The caller removes it best-effort.
    """
    base = Path(os.environ.get("DCC_OFFICE_HOST_TMP") or os.environ.get("TEMP") or os.environ.get("TMP") or ".")
    base = base.resolve()
    for _attempt in range(10):
        candidate = base / f"dcc-office-host-{os.getpid()}-{time.time_ns()}"
        try:
            candidate.mkdir()
            return candidate
        except FileExistsError:
            continue
    raise OSError(f"could not create a host work directory under {base}")


def _abs(path: str | Path) -> str:
    """Absolute path for every path handed to the host.

    The host's COM backends resolve relative paths against the Office
    process working directory (usually System32), so relative input is
    guaranteed to fail — normalize on the client side.
    """
    return str(Path(path).resolve())


def _matching_response(stdout: str, request_id: str) -> dict[str, Any]:
    """Return the JSON-RPC response from the host's NDJSON output.

    Hosts write the command response followed by zero or more event
    notifications. Selecting by request id keeps the one-shot client
    compatible with both a single-response host and that stream.
    """
    responses: list[dict[str, Any]] = []
    for line_number, line in enumerate(stdout.splitlines(), start=1):
        if not line.strip():
            continue
        try:
            payload = json.loads(line)
        except json.JSONDecodeError as exc:
            raise ValueError(f"invalid JSON on host output line {line_number}: {exc}") from exc
        if isinstance(payload, dict) and payload.get("id") == request_id:
            responses.append(payload)
    if not responses:
        raise ValueError(f"host output did not contain response id {request_id!r}")
    if len(responses) != 1:
        raise ValueError(f"host output contained {len(responses)} responses for id {request_id!r}")
    return responses[0]


def rpc(method: str, params: dict[str, Any], *, app: str = DEFAULT_APP) -> dict[str, Any]:
    """One JSON-RPC exchange with the host over stdin/stdout.

    Stdio is redirected through temporary files instead of OS pipes: the
    wire contract is unchanged, while confined environments that block
    anonymous pipes can still talk to the host and large output cannot
    deadlock.
    """
    binary = find_host_binary()
    if binary is None:
        return {
            "success": False,
            "backend": None,
            "reason": f"OFFICE_HOST_NOT_FOUND: {HOST_EXE} not found; set {OFFICE_HOST_ENV} "
            "or install the dcc-mcp-office runtime",
        }
    request = {"jsonrpc": "2.0", "id": "req", "method": method, "params": params}
    work = _rpc_work_dir()
    try:
        stdin_path = work / "request.json"
        stdout_path = work / "response.json"
        stderr_path = work / "stderr.txt"
        stdin_path.write_text(json.dumps(request), encoding="utf-8")
        try:
            with (
                stdin_path.open("r", encoding="utf-8") as stdin_file,
                stdout_path.open("w", encoding="utf-8") as stdout_file,
                stderr_path.open("w", encoding="utf-8") as stderr_file,
            ):
                proc = subprocess.run(
                    [binary, f"--app={app}", "--stdio"],
                    stdin=stdin_file,
                    stdout=stdout_file,
                    stderr=stderr_file,
                    timeout=DEFAULT_TIMEOUT_SEC,
                    check=False,
                )
        except subprocess.TimeoutExpired as exc:
            return {"success": False, "backend": "office_host", "reason": f"host timed out: {exc}"}
        stdout = stdout_path.read_text(encoding="utf-8")
        stderr = stderr_path.read_text(encoding="utf-8")
    finally:
        shutil.rmtree(work, ignore_errors=True)
    if proc.returncode != 0:
        return {"success": False, "backend": "office_host", "reason": stderr.strip() or "host exited non-zero"}
    try:
        payload = _matching_response(stdout, "req")
    except ValueError as exc:
        return {"success": False, "backend": "office_host", "reason": f"invalid host output: {exc}"}
    result = payload.get("result")
    if result is None:
        return {"success": False, "backend": "office_host", "reason": str(payload.get("error", "empty result"))}
    return {"success": True, "backend": "office_host", "result": result}


def ping() -> dict[str, Any]:
    return rpc("office.host.ping", {})


def handshake(app: str = DEFAULT_APP) -> dict[str, Any]:
    """office.host.handshake — protocol + capability manifest."""
    return rpc("office.host.handshake", {"requested_app": app}, app=app)


def compile_workbook(ir_path: str | Path, output_path: str | Path) -> dict[str, Any]:
    """workbook.compile — Workbook IR JSON → XLSX through the shared host."""
    return rpc(
        "office.command.execute",
        {
            "capability": "workbook.compile",
            "input": {"ir": _abs(ir_path), "output": _abs(output_path)},
        },
    )
