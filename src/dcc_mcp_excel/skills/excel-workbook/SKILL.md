---
name: excel-workbook
description: >-
  Compile a Workbook IR JSON (office-ir/1.0, kind: workbook) into a native
  XLSX through the headless Open XML backend, then read the artifact back and
  verify every cell against the source IR. Use whenever the agent must
  produce an .xlsx from structured data — production reports, shot lists,
  timesheets, budget tables. No Excel installation required.
license: MIT
allowed-tools: Bash Read
metadata:
  dcc-mcp:
    dcc: excel
    layer: domain
    stage: authoring
    version: 0.1.0
    tags:
      - excel
      - workbook
      - xlsx
      - generate
      - readback
      - openxml
    search-hint: >-
      generate xlsx, make spreadsheet, excel report, shot list, timesheet,
      compile workbook, read back workbook
    tools: tools.yaml
---

# excel-workbook (Authoring stage)

Workbook generation through the designed pipeline:

1. data planner picks worksheets and their grid layout
2. Workbook IR document (contract: dcc-mcp-office-ir workbook schema)
3. headless Open XML compile → XLSX (no Excel installation needed)
4. read-back verification: reopen the artifact and compare every cell
   against the source IR
5. structural validation report

The headless path is `verified`: it runs on a Linux runner in CI with no
Office installed. Native recalculation, chart/pivot refresh and PDF export
are `host_limited` and are reported, never faked — see `excel-capabilities`.

## Related skills

- `excel-capabilities` — the graded capability report for this adapter

## Input contract

- `input` — path to a Workbook IR JSON envelope
  (`schema_version: office-ir/1.0`, `kind: workbook`)
- `output_dir` — artifact directory
- `verify` — run the read-back verification step (default: true)

```json
{
  "schema_version": "office-ir/1.0",
  "kind": "workbook",
  "document_id": "draft:shot-list",
  "metadata": {"title": "Shot list", "author": "pipeline", "language": "en"},
  "document": {
    "worksheets": [
      {"name": "Shots", "rows": [["shot", "status"], ["sh010", "ip"]]}
    ],
    "tables": [{"worksheet": "Shots", "range": "A1:B2"}],
    "formulas": [{"worksheet": "Shots", "cell": "C2", "formula": "=COUNTA(A2:A2)"}]
  },
  "outputs": ["xlsx"]
}
```

## Decision rules

- model the data as a grid first; never hand-place cell coordinates before
  the shape of the table is settled
- one worksheet per logical table; sheet names are sanitized to Excel's
  31-character limit and de-duplicated automatically
- prefer `document.worksheets[].rows` for bulk data and `document.cells`
  or `document.formulas` for addressed, sparse writes
- formulas are written as text: the value only materializes when Excel opens
  the file. Say so in the hand-off rather than implying a computed number.

## Scripts

- `generate_workbook` — IR → XLSX + read-back verification + validation report
- `validate_workbook` — validate a Workbook IR without generating
- `inspect_workbook` — read-only inventory of an existing XLSX

## Validation rules

- envelope contract enforced at load (`workbook_ir`): bad schema version,
  unknown feature kinds, malformed A1 references, duplicate sheet names and
  wrong value types are hard errors carrying a json-path hint
- read-back compares every cell; a mismatch fails the run instead of
  warning
- artifacts must exist and be non-empty

## Known limits

- openpyxl does not evaluate formulas: cached values are absent until Excel
  recalculates (`excel.workbook.calculate` is `host_limited`)
- charts and pivots declared in the IR are parsed and validated, but are not
  rendered by the headless writer
- the compiler raises on specs it cannot write rather than silently dropping
  them
- conditional formatting is written as a rule; Excel renders it on open

## Agent-visible summary

Result context: artifact path, sheets written, rows written, read-back
verdict (cells checked, mismatches), validation checks + warnings, and the
capability grading for anything host-limited.
