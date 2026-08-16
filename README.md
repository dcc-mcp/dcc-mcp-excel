# dcc-mcp-excel

Excel adapter for the DCC-MCP ecosystem — thin application layer over
[dcc-mcp-office](https://github.com/dcc-mcp/dcc-mcp-office).

**Status: planned — not started.** This repository is a placeholder created
as part of the Office Automation Platform repo split (see
`dcc-mcp-office/docs/adr/006-shared-office-core-split.md`). Work starts in
Phase 2, blocked on the `dcc-mcp-office` M1 (COM MVP) release.

## Scope (proposal §11.2)

- - `excel.workbook.calculate` — native formula recalc
- `excel.table.update` / `excel.chart.generate` / `excel.pivot.refresh`
- Graph Workbook sessions for cloud scenarios (proposal §6.3)
- production dashboards via the `office-generate-production-dashboard` skill

## Upstream

- `dcc-mcp-office` — protocol, IR, C# runtime, Open XML worker, security
  policy, generic skills.
- `dcc-mcp-core` — gateway, jobs, artifacts, skills runtime, lifecycle.

## License

MIT
