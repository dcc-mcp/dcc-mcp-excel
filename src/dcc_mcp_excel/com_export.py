"""Desktop COM export — XLSX → PDF via Excel (host-limited).

Mirrors `dcc_mcp_powerpoint.render`: the COM backend never claims success
silently. When Excel is unavailable it reports an explicit reason
(`OFFICE_APP_NOT_INSTALLED`) and leaves the compiled XLSX in place — a
missing export is a reported degradation, never a fabricated artifact.

This module is opt-in: pywin32 is a dev/test-only dependency and is
imported inside the function body, never at package import time.
"""

from __future__ import annotations

import functools
import logging
from pathlib import Path
from typing import Any

from .host_matrix import find_excel_executable

logger = logging.getLogger(__name__)

OFFICE_UNAVAILABLE = "OFFICE_APP_NOT_INSTALLED"

# Excel file-format constant for PDF (xlTypePDF).
_XL_TYPE_PDF = 0


@functools.lru_cache(maxsize=1)
def office_available() -> bool:
    """True when Excel is installed and COM dispatch works (cached)."""
    if find_excel_executable() is None:
        return False
    try:
        import win32com.client

        app = win32com.client.DispatchEx("Excel.Application")
        try:
            return bool(getattr(app, "Version", ""))
        finally:
            # Dedicated probe instance: quit immediately, never leak a
            # process and never touch a user-started Excel.
            app.Quit()
    except Exception as exc:  # noqa: BLE001 — any COM failure means "unavailable"
        logger.debug("Excel COM probe failed: %s", exc)
        return False


def export_pdf(xlsx_path: str | Path, out_dir: str | Path) -> dict[str, Any]:
    """Export a workbook to PDF with the desktop COM backend.

    Returns a result context: backend, pdf path, office version — or an
    explicit unavailable reason.
    """
    # COM runs inside Excel's process: relative paths resolve against
    # Excel's working directory, not ours — always hand absolute paths to
    # the COM surface.
    xlsx = Path(xlsx_path).resolve()
    if not xlsx.is_file():
        return {"success": False, "backend": None, "reason": f"input not found: {xlsx}"}
    out = Path(out_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    if not office_available():
        return {
            "success": False,
            "backend": None,
            "reason": f"{OFFICE_UNAVAILABLE}: Excel COM is unavailable; XLSX compiled, PDF export skipped",
        }

    import win32com.client

    app = win32com.client.DispatchEx("Excel.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = False
        workbook = app.Workbooks.Open(str(xlsx))
        try:
            pdf_path = out / f"{xlsx.stem}.pdf"
            workbook.ExportAsFixedFormat(_XL_TYPE_PDF, str(pdf_path))
        finally:
            workbook.Close(SaveChanges=False)
        version = str(getattr(app, "Version", ""))
    finally:
        app.Quit()

    if not pdf_path.is_file():
        return {"success": False, "backend": "excel_com", "reason": f"PDF export produced no file: {pdf_path}"}
    return {
        "success": True,
        "backend": "excel_com",
        "pdf": str(pdf_path),
        "office_version": version,
    }
