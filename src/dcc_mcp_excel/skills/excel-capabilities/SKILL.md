---
name: excel-capabilities
description: >-
  Report which Excel capabilities are verified by CI and which require a
  desktop Excel installation (host_limited), plus the host matrix and the
  start-up preflight self-check. Use before promising an Excel capability to
  a user, and when an Excel operation degrades or is unavailable.
license: MIT
allowed-tools: Bash Read
metadata:
  dcc-mcp:
    dcc: excel
    layer: infrastructure
    stage: diagnostics
    version: 0.1.0
    tags:
      - excel
      - capabilities
      - host-matrix
      - preflight
      - diagnostics
    search-hint: >-
      excel capability, is it verified, host limited, do I need Excel,
      preflight, host matrix, why is recalculation missing
    tools: tools.yaml
---

# excel-capabilities (Diagnostics stage)

The grade of a capability is evidence, not confidence:

| Grade | Meaning |
|---|---|
| `verified` | Covered by a CI-green headless test on a Linux runner. No Office installation involved. |
| `host_limited` | Requires desktop Excel (COM). The artifact is structurally valid without it, but the value only materializes when Excel opens the file. |
| `unimplemented` | Not built. Reported instead of silently degrading. |

v0.1.0: headless compile, read-back verification, grid/formula/named-range/
validation/conditional-format writes are `verified`. Native recalculation,
chart and pivot materialization, and PDF export are `host_limited`.
Microsoft Graph Workbook sessions are `unimplemented` by design (proposal
§6.3 needs Azure app registration and tenant consent).

## Do not over-promise

- Do not tell a user a computed number is in the file until Excel has opened
  it. openpyxl writes the formula text; it does not evaluate it.
- Do not describe a chart or pivot as "created" — the IR declares and
  validates them; Excel materializes them.
- A `preflight` failure is a report, not a crash: the headless surface stays
  usable without desktop Excel.

## Scripts

- `capabilities` — print the graded capability report
- `preflight` — host matrix + start-up self-check (what this machine can do)

## After Failure

- `preflight ok=false` means the headless backend itself is missing
  (openpyxl not installed) — install the package with its `dev` extra.
- `desktop_excel=false` is expected on Linux/macOS and on CI: only the
  `host_limited` capabilities degrade.
