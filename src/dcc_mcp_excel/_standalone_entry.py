"""Dual-purpose entry point: CLI commands + skill-script passthrough.

Mirrors `dcc_mcp_powerpoint._standalone_entry` (ADR 003 pattern) so the
PyOxidizer bundle behaves the same way: `dcc-mcp-excel <command>` for the
CLI, and `dcc-mcp-excel <script.py> ...` to run a skill script with the
embedded interpreter (gateway `execute_script` parity).
"""

from __future__ import annotations

import argparse
import json
import os
import runpy
import sys
from collections.abc import Sequence
from pathlib import Path

from . import __version__
from .capabilities import report
from .host_client import compile_workbook
from .host_matrix import preflight
from .server_launcher import ServeConfig, serve
from .validate import validate_artifacts

_PYTHON_SCRIPT_SUFFIXES = frozenset({".py", ".pyw"})


def _is_skill_script_invocation(argv: Sequence[str]) -> bool:
    if len(argv) < 2:
        return False
    script = Path(argv[1])
    return script.suffix.lower() in _PYTHON_SCRIPT_SUFFIXES and script.is_file()


def _run_skill_script(argv: Sequence[str]) -> None:
    script = str(Path(argv[1]).resolve())
    original_argv = sys.argv
    sys.argv = [script, *argv[2:]]
    try:
        runpy.run_path(script, run_name="__main__")
    finally:
        sys.argv = original_argv


def _cmd_compile(args: argparse.Namespace) -> int:
    result = compile_workbook(args.input, args.output)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("success") else 1


def _cmd_capabilities(args: argparse.Namespace) -> int:
    print(json.dumps(report(), ensure_ascii=False))
    return 0


def _cmd_preflight(args: argparse.Namespace) -> int:
    status = preflight()
    print(json.dumps(status, ensure_ascii=False))
    return 0 if status.get("ok") else 1


def _cmd_validate(args: argparse.Namespace) -> int:
    target = Path(args.input)
    paths = sorted(str(p) for p in target.rglob("*") if p.is_file()) if target.is_dir() else [args.input]
    report_body = validate_artifacts(paths)
    print(json.dumps({"success": report_body["ok"], "context": report_body}, ensure_ascii=False))
    return 0 if report_body["ok"] else 1


def _cmd_serve(args: argparse.Namespace) -> int:
    return serve(
        ServeConfig(
            server=args.server,
            mcp_port=args.mcp_port,
            registry_dir=args.registry_dir,
            pid_file=args.pid_file,
        )
    )


def main(argv: Sequence[str] | None = None) -> None:
    resolved_argv = list(sys.argv if argv is None else argv)
    os.environ.setdefault("DCC_MCP_PYTHON_EXECUTABLE", sys.executable)
    if _is_skill_script_invocation(resolved_argv):
        _run_skill_script(resolved_argv)
        return

    parser = argparse.ArgumentParser(
        prog="dcc-mcp-excel",
        description="Standalone Excel adapter (thin layer over the dcc-mcp-office runtime)",
    )
    sub = parser.add_subparsers(dest="command", required=True)
    c = sub.add_parser("compile", help="Workbook IR JSON -> XLSX via the shared office host")
    c.add_argument("--input", required=True)
    c.add_argument("--output", required=True)
    c.set_defaults(func=_cmd_compile)
    p = sub.add_parser("capabilities", help="print the capability grading report")
    p.set_defaults(func=_cmd_capabilities)
    d = sub.add_parser("preflight", help="host matrix + start-up self-check")
    d.set_defaults(func=_cmd_preflight)
    v = sub.add_parser("validate", help="validate artifacts exist and are non-empty")
    v.add_argument("--input", required=True)
    v.set_defaults(func=_cmd_validate)
    s = sub.add_parser("serve", help="start and register the Excel MCP adapter")
    s.add_argument("--server", help="absolute path to dcc-mcp-server")
    s.add_argument("--mcp-port", type=int, default=0, help="MCP port; 0 lets the OS choose")
    s.add_argument("--registry-dir", help="shared DCC-MCP FileRegistry directory")
    s.add_argument("--pid-file", help="write the server PID to this file")
    s.set_defaults(func=_cmd_serve)
    parser.add_argument("--version", action="version", version=f"dcc-mcp-excel {__version__}")
    args = parser.parse_args(resolved_argv[1:])
    raise SystemExit(args.func(args))


if __name__ == "__main__":
    main()
