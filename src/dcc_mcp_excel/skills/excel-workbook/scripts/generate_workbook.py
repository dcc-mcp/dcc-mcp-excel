"""excel-workbook / generate_workbook — Workbook IR → XLSX + read-back.

Parameter resolution order (dcc-mcp-core execute_script convention):
1. stdin JSON: {"input": ..., "output_dir": ..., "verify": ...}
2. CLI flags: --input --out --verify/--no-verify
"""

from __future__ import annotations

import argparse
import json
import sys

from dcc_mcp_excel.compiler import compile_workbook
from dcc_mcp_excel.readback import read_back
from dcc_mcp_excel.validate import validate_artifacts, validate_envelope
from dcc_mcp_excel.workbook_io import resolve_output_path


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8.

    On Windows, a piped subprocess stdout defaults to the ANSI codepage
    (charmap) and fails on CJK text. The gateway reads JSON from stdout, so
    the encoding is part of the script contract.
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def run(params: dict) -> None:
    from dcc_mcp_excel.workbook_ir import load_workbook_ir

    envelope = load_workbook_ir(params["input"])
    report = validate_envelope(envelope)
    if not report["ok"]:
        print(
            json.dumps(
                {"success": False, "message": "workbook IR failed structural validation", "context": report},
                ensure_ascii=False,
            )
        )
        return

    out_dir = params.get("output_dir", "output")
    xlsx = compile_workbook(envelope, resolve_output_path(envelope, out_dir))
    artifacts = [str(xlsx)]

    readback_report = None
    if params.get("verify", True):
        readback_report = read_back(envelope, xlsx)

    artifact_report = validate_artifacts(artifacts)
    success = artifact_report["ok"] and (readback_report is None or readback_report.ok)
    message = f"workbook '{envelope.document_id}' compiled ({len(envelope.document.worksheets)} worksheets"
    message += f", {len(readback_report.mismatches)} cell mismatches)" if readback_report else ")"

    print(
        json.dumps(
            {
                "success": success,
                "message": message,
                "context": {
                    "document_id": envelope.document_id,
                    "artifact": str(xlsx),
                    "artifacts": artifacts,
                    "backend": "openxml",
                    "read_back": readback_report.to_dict() if readback_report else None,
                    "validation": report,
                    "artifacts_ok": artifact_report,
                },
            },
            ensure_ascii=False,
        )
    )


def main() -> None:
    _force_utf8_stdio()
    params: dict = {}
    if not sys.stdin.isatty():
        raw = sys.stdin.read()
        if raw.strip():
            try:
                params = json.loads(raw)
            except json.JSONDecodeError:
                params = {}
    if not params:
        parser = argparse.ArgumentParser(description="Workbook IR → XLSX + read-back verification")
        parser.add_argument("--input", required=True, help="Workbook IR JSON path")
        parser.add_argument("--out", dest="output_dir", default="output", help="output directory")
        parser.add_argument("--verify", dest="verify", action=argparse.BooleanOptionalAction, default=True)
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
