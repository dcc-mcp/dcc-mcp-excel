"""dcc-mcp-excel — Excel adapter for the DCC-MCP ecosystem.

Thin application layer over `dcc-mcp-office`. This package owns Excel
semantics only: the Workbook IR contract, the headless Open XML compile +
read-back path, capability grading and the office-host launcher. The shared
machinery (protocol, IR envelope, C# COM runtime, Open XML worker) comes
from `dcc-mcp-office`.

Capability surface (v0.1.0):
- workbook_ir: the Workbook IR contract (mirrors dcc-mcp-office-ir)
- compiler: Workbook IR -> XLSX (headless Open XML, openpyxl; opt-in import)
- readback: write-then-read-back verification (opt-in import)
- capabilities: verified / host_limited grading (stdlib-only)
- host_matrix: host matrix + preflight self-check (stdlib-only)
- validate: structural validation reports (stdlib-only)
- host_client: stdlib-only JSON-RPC client for the shared office host
- com_export: desktop COM PDF export (pywin32, dev/test-only, Windows)
- server_launcher: registers the bundled skill packs with dcc-mcp-server

Dependency policy (mirrors dcc-mcp-powerpoint): openpyxl / pywin32 are
opt-in. Importing this package pulls neither.
"""

from __future__ import annotations

from .capabilities import CAPABILITIES, Capability, report
from .host_matrix import HOST_MATRIX, preflight
from .validate import validate_artifacts, validate_envelope
from .workbook_ir import IrValidationError, WorkbookEnvelope, load_workbook_ir

__version__ = "0.1.0"

__all__ = [
    "CAPABILITIES",
    "HOST_MATRIX",
    "Capability",
    "IrValidationError",
    "WorkbookEnvelope",
    "__version__",
    "load_workbook_ir",
    "preflight",
    "report",
    "validate_artifacts",
    "validate_envelope",
]
