# Install the Excel adapter

The adapter is an out-of-process Excel service. It does not install an Office
add-in and it never attaches to a user-owned Excel process. The official
`dcc-mcp-server` owns MCP transport, FileRegistry registration and
heartbeats; this package owns Excel skill packs and workstation generation.

## Requirements

- Python 3.9+ on Windows, macOS or Linux.
- **Headless workbook generation does not require Excel.** Open XML compile,
  read-back verification and structural inspection run anywhere, including on
  a CI runner with no Office installed.
- Desktop Microsoft Excel is required only for the `host_limited` surfaces:
  native formula recalculation, chart and pivot rendering, and PDF export.
- Current `dcc-mcp-cli` and `dcc-mcp-server` from the same DCC-MCP release.

Confirm the runtime before installing the adapter:

```powershell
dcc-mcp-cli doctor
dcc-mcp-server --version
```

## Install from PyPI

```bash
pip install "dcc-mcp-excel[headless]"
```

The base package keeps a stdlib-only runtime so it can be installed inside a
DCC that pins its own openpyxl; the `headless` extra adds openpyxl for
workbook compilation.

Verify what this machine can actually do before an Excel run:

```powershell
dcc-mcp-excel preflight
```

`ok=true` means the headless backend is ready. `desktop_excel=false` is
expected on Linux/macOS and on CI: only the `host_limited` capabilities
degrade, and the report lists exactly which ones.

## Source checkout

For adapter development, use an isolated virtual environment:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[dev,sidecar]"
.\.venv\Scripts\dcc-mcp-excel.exe serve
```

## Verification

Registration is not proven by package installation. Verify the live route in
order, copying the tool slug returned by `search` rather than constructing it:

```powershell
dcc-mcp-cli doctor
dcc-mcp-cli list
dcc-mcp-cli wait-ready --dcc-type excel
dcc-mcp-cli search --dcc-type excel --query "generate workbook"
dcc-mcp-cli describe <slug-from-search>
```

If `list` has no Excel row, inspect the adapter process output and the log
paths reported by `dcc-mcp-cli doctor`. Do not fall back to generic desktop
automation.

## Known limits

- openpyxl writes formulas as text; cached values appear only after Excel
  recalculates the file.
- Charts and pivots declared in the Workbook IR are validated but not
  rendered by the headless writer.
- Pagination, macro execution and VBA are outside the headless surface.
