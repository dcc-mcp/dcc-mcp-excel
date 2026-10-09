"""excel-capabilities / capabilities — print the graded capability report."""

from __future__ import annotations

import json
import sys

from dcc_mcp_excel.capabilities import HOST_LIMITED, VERIFIED, by_grade, report


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def run(params: dict) -> None:
    del params  # no inputs; the report is derived from the shipped contract
    body = report()
    verified = [c.name for c in by_grade(VERIFIED)]
    host_limited = [c.name for c in by_grade(HOST_LIMITED)]
    print(
        json.dumps(
            {
                "success": True,
                "message": f"{len(verified)} verified, {len(host_limited)} host_limited",
                "context": body,
            },
            ensure_ascii=False,
        )
    )


def main() -> None:
    _force_utf8_stdio()
    try:
        run({})
    except Exception as exc:  # noqa: BLE001 — surface as structured error
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
