"""Shared pytest fixtures — make the src package importable in-tree."""

from __future__ import annotations

import os
import sys
from collections.abc import Iterator
from pathlib import Path

import pytest

SRC = Path(__file__).resolve().parent.parent / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


@pytest.fixture
def skill_env() -> Iterator[dict[str, str]]:
    """Environment for launching a skill script as a subprocess.

    CI installs the package (`pip install -e .[dev]`), so the script resolves
    `dcc_mcp_excel` through site-packages. A bare checkout has no install, so
    PYTHONPATH carries the in-tree `src` — one contract, both situations.
    """
    env = os.environ.copy()
    existing = env.get("PYTHONPATH")
    env["PYTHONPATH"] = f"{SRC}{os.pathsep}{existing}" if existing else str(SRC)
    yield env
