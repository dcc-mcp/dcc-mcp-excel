"""Runtime purity guard — importing the package must not pull dev deps.

Dependency policy: openpyxl / pywin32 are opt-in. The headless compile path
is the product surface on Linux CI, but the *package* must stay importable
in a DCC that pins its own openpyxl. The check runs in a clean interpreter
so pytest's own imports cannot mask a violation.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

PROBE = (
    "import sys; "
    "sys.path.insert(0, r'{src}'); "
    "import dcc_mcp_excel; "
    "bad = [m for m in ('openpyxl', 'win32com', 'pythoncom') if m in sys.modules]; "
    "assert not bad, 'opt-in deps imported at package import: ' + repr(bad); "
    "print('pure')"
).format(src=str(ROOT / "src"))


def test_package_import_pulls_no_opt_in_deps() -> None:
    proc = subprocess.run(
        [sys.executable, "-c", PROBE],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=120,
        check=False,
    )
    assert proc.returncode == 0, proc.stderr
    assert "pure" in proc.stdout


def test_preflight_works_without_opt_in_deps() -> None:
    """The stdlib-only preflight path must work in the clean interpreter too."""
    probe = (
        "import sys, json; "
        "sys.path.insert(0, r'{src}'); "
        "from dcc_mcp_excel.host_matrix import preflight; "
        "print(json.loads(json.dumps(preflight()))['app'])"
    ).format(src=str(ROOT / "src"))
    proc = subprocess.run(
        [sys.executable, "-c", probe],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=120,
        check=False,
    )
    assert proc.returncode == 0, proc.stderr
    assert proc.stdout.strip() == "excel"
