# dcc-mcp-excel

Excel adapter for the DCC-MCP ecosystem — the **thin application layer** over
[dcc-mcp-office](https://github.com/dcc-mcp/dcc-mcp-office): structured
workbook generation from production data, headless XLSX compilation and
read-back verification.

**Status: v0.1.0 — headless path live.** The compile pipeline runs end to end
without Excel installed: Workbook IR (`office-ir/1.0`) → headless Open XML
compile (openpyxl) → write-then-read-back verification → structural report.
Skill scripts execute through the dcc-mcp gateway.

Two grades of capability, and the difference is evidence:

| Grade | What it means | Examples |
|---|---|---|
| `verified` | CI-green on a Linux runner, no Office installed | workbook compile, read-back, grid/formula/named-range/validation/conditional-format writes |
| `host_limited` | Requires desktop Excel (COM); the artifact is valid but the value only materializes when Excel opens it | native recalculation, chart and pivot rendering, PDF export |

Run `dcc-mcp-excel capabilities` (or the `excel-capabilities` skill) for the
full graded list.

<!-- dcc-mcp-coverage-pointer:start -->
<!-- Generated from dcc-mcp-catalog.yml by scripts/generate_adapter_pointer.py in dcc-mcp/dcc-mcp-core. Do not edit by hand. -->
## Part of the DCC-MCP host matrix

**dcc-mcp-excel** — Excel adapter for DCC-MCP — headless Workbook IR to XLSX compile
with read-back verification over the dcc-mcp-office runtime.

It is one of **47 host adapters** in the DCC-MCP catalog. Every adapter speaks the same
MCP protocol and builds on the same core runtime contract; each one exposes the tools
its own host needs on top of that.

- [All host adapters and install metadata](https://dcc-mcp.github.io/ecosystem)
- [Host matrix on the core README](https://github.com/dcc-mcp/dcc-mcp-core#readme)
- [Showcase](https://dcc-mcp.github.io/showcase)

This block is generated from the catalog entry in
[`dcc-mcp-catalog.yml`](https://github.com/dcc-mcp/dcc-mcp-core/blob/main/dcc-mcp-catalog.yml).
Re-run the generator after changing the catalog.
<!-- dcc-mcp-coverage-pointer:end -->

## Why headless first

The `dcc-mcp-office` shared core already ships a CI-green headless xlsx
writer (`skills/office-generate-production-dashboard`). Building the adapter
on that path means the 1.0 gate — host matrix, preflight self-check,
write-then-read-back — is verifiable in CI on a plain Linux runner instead of
depending on a licensed desktop Office install.

Desktop Excel stays the path for the things only Excel can do: evaluate a
formula, render a chart, refresh a pivot, paginate a PDF. Those are declared
and graded `host_limited`, never claimed as verified.

## Install

```bash
pip install "dcc-mcp-excel"
# headless compilation needs the openpyxl extra:
pip install "dcc-mcp-excel[headless]"
```

See [install.md](./install.md) for the verified live route.

## Skill packs

- `excel-workbook` — compile a Workbook IR into an editable XLSX, then reopen
  it and verify every cell against the source IR; validate an IR; inventory an
  existing workbook.
- `excel-capabilities` — the graded capability report plus the host matrix and
  the start-up preflight self-check.

## Agent usage (via dcc-mcp-cli / gateway)

1. Start the registered adapter (the launcher binds the bundled skill packs
   and script runtime to the official `dcc-mcp-server`):
   `dcc-mcp-excel serve`
2. Run a skill script directly (gateway `execute_script` contract —
   stdin JSON or CLI flags):
   ```bash
   python src/dcc_mcp_excel/skills/excel-workbook/scripts/generate_workbook.py \
     --input examples/shot_list.json --out out
   ```
3. Validate the packs with the official linter:
   ```bash
   dcc-mcp-cli lint src/dcc_mcp_excel/skills --warnings-as-errors --non-interactive
   ```

## From production tracking to a spreadsheet

`dcc-mcp-fpt` reads Flow Production Tracking; this adapter turns what it
returns into something you can send. `tests/test_e2e.py` pins that chain with
a real `find_entities` payload: entities → Workbook IR → XLSX → read-back.

## Development

```bash
pip install -e ".[dev]"
pytest
ruff check src tests
```

## License

MIT
