# AGENTS.md — dcc-mcp-excel

> Progressive disclosure: this file is a **map**, not an encyclopedia.

## 30-Second Summary

`dcc-mcp-excel` is the **thin Excel adapter** over `dcc-mcp-office`. It owns
Excel application semantics only: workbook generation from structured data,
the headless XLSX compile + read-back path, capability grading and the
office-host launcher. Shared machinery (protocol, IR envelope, C# COM
runtime, jobs, security policy) comes from `dcc-mcp-office` + `dcc-mcp-core`.

**Current status:** v0.1.0 — headless path live. Workbook IR → Open XML
compile → read-back verification runs end to end with **no Excel
installation**, and the gate runs on a Linux runner in CI. Native
recalculation, chart/pivot rendering and PDF export are `host_limited`:
declared, graded, and never claimed as verified.

## Repo Map

| Path | What it is |
|---|---|
| `src/dcc_mcp_excel/workbook_ir.py` | Workbook IR contract (mirrors dcc-mcp-office-ir) |
| `src/dcc_mcp_excel/workbook_io.py` | shared write/read helpers (titles, value comparison) |
| `src/dcc_mcp_excel/compiler.py` | headless Open XML compiler: Workbook IR → XLSX (openpyxl) |
| `src/dcc_mcp_excel/readback.py` | write-then-read-back verification (the 1.0 gate) |
| `src/dcc_mcp_excel/capabilities.py` | verified / host_limited / unimplemented grading |
| `src/dcc_mcp_excel/host_matrix.py` | host matrix + preflight self-check |
| `src/dcc_mcp_excel/com_export.py` | desktop COM PDF export (pywin32, Windows only) |
| `src/dcc_mcp_excel/host_client.py` | stdlib-only JSON-RPC client for the shared office host |
| `src/dcc_mcp_excel/server_launcher.py` | registers the bundled skills with `dcc-mcp-server` |
| `src/dcc_mcp_excel/skills/excel-workbook/` | SKILL.md + tools.yaml + scripts (generate/validate/inspect) |
| `src/dcc_mcp_excel/skills/excel-capabilities/` | capability report + preflight scripts |
| `examples/` | shot-list Workbook IR + generated XLSX |
| `tests/` | pytest (headless; COM only via subprocess boundary) |
| `docs/adr/` | adapter-level decisions |

## Upstream Dependencies

- `dcc-mcp-core` (pip) — gateway, skills runtime, sidecar lifecycle.
- `dcc-mcp-office` — Rust crates (`dcc-mcp-office-protocol`,
  `dcc-mcp-office-ir`, `dcc-mcp-office-tools`) + the `office-host` runtime.

## Capabilities owned here

- `workbook.compile` — Workbook IR → XLSX (headless) → read-back verification.
- `workbook.read_back` — reopen the artifact and compare every cell.
- `excel.workbook.calculate` / `excel.chart.generate` / `excel.pivot.refresh`
  — `host_limited` in v0.1.0 (COM backend not wired).
- Graph Workbook sessions (proposal §6.3) — out of scope; needs Azure app
  registration and tenant consent.

## Dependency policy

- **Package import is stdlib-only.** openpyxl and pywin32 are **opt-in**:
  `compiler` / `readback` import openpyxl inside the module, `com_export`
  imports pywin32 inside the function body. A test pins this
  (`tests/test_runtime_purity.py`).
- One xlsx implementation only (ADR 004): no second writer inside the
  adapter, even as a fallback. A missing shared host is reported, not worked
  around.

## Test

```bash
pip install -e ".[dev]"
pytest
ruff check src tests
```

The headless gate must stay green on **Linux** — that is the proof the
adapter does not depend on a desktop Excel installation.
