"""Host matrix and preflight self-check (the 1.0 standard gate).

The matrix is data, not code paths: adding an Office build means adding one
entry. `preflight()` is the start-up self-check an adapter runs before it
publishes itself — it never raises on a missing desktop Office, it reports
which capabilities degrade because of it.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .capabilities import HOST_LIMITED, VERIFIED, by_grade

ADAPTER_APP = "excel"
EXCEL_EXE_CANDIDATES = (
    Path("C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"),
    Path("C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXE"),
)


@dataclass(frozen=True)
class HostEntry:
    app: str
    version: str
    platform: str
    status: str
    notes: str


HOST_MATRIX: tuple[HostEntry, ...] = (
    HostEntry(
        app="Excel",
        version="Microsoft 365 (Office 16)",
        platform="windows",
        status="host_limited",
        notes="Desktop COM path: native recalculation, chart/pivot refresh and PDF export. Not exercised by CI.",
    ),
    HostEntry(
        app="Excel",
        version="2019 / 2021 (Office 16)",
        platform="windows",
        status="host_limited",
        notes="Same COM surface as Microsoft 365; expected to work, not CI-verified.",
    ),
    HostEntry(
        app="Excel",
        version="2016 (Office 16)",
        platform="windows",
        status="host_limited",
        notes="Same COM surface; not CI-verified.",
    ),
    HostEntry(
        app="headless",
        version="openpyxl",
        platform="any",
        status="verified",
        notes="Open XML compile + read-back verification. No Office installation required; runs on ubuntu-latest in CI.",
    ),
)


def find_excel_executable() -> Path | None:
    """Absolute path to a desktop Excel installation, or None when absent."""
    for candidate in EXCEL_EXE_CANDIDATES:
        if candidate.is_file():
            return candidate
    return None


def headless_available() -> bool:
    """True when the headless Open XML backend can be imported."""
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        return False
    return True


def preflight() -> dict[str, object]:
    """Start-up self-check: report what this machine can actually do.

    Never raises for a missing desktop Excel. A degraded host is a report,
    not a failure — the headless surface stays usable without it.
    """
    excel = find_excel_executable()
    headless = headless_available()
    checks = {
        "headless_openxml": headless,
        "desktop_excel": excel is not None,
    }
    degraded = [c.name for c in by_grade(HOST_LIMITED)] if excel is None else []
    return {
        "ok": headless,
        "app": ADAPTER_APP,
        "excel_path": str(excel) if excel else None,
        "checks": checks,
        "verified": [c.name for c in by_grade(VERIFIED)],
        "host_limited": [c.name for c in by_grade(HOST_LIMITED)],
        "degraded_without_office": degraded,
        "host_matrix": [
            {
                "app": entry.app,
                "version": entry.version,
                "platform": entry.platform,
                "status": entry.status,
                "notes": entry.notes,
            }
            for entry in HOST_MATRIX
        ],
    }
