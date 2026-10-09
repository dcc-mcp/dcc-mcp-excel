"""excel-capabilities / preflight — host matrix + start-up self-check."""

from __future__ import annotations

import json
import sys

from dcc_mcp_excel.host_matrix import preflight


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def run(params: dict) -> None:
    del params  # no inputs; the probe reads the local machine
    status = preflight()
    checks = status["checks"]
    message = "headless Open XML backend ready"
    if not checks["desktop_excel"]:
        message += "; desktop Excel absent — host-limited capabilities degrade"
    print(json.dumps({"success": bool(status["ok"]), "message": message, "context": status}, ensure_ascii=False))


def main() -> None:
    _force_utf8_stdio()
    try:
        run({})
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
