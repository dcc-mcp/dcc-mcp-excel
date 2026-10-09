# ADR 004 — 无头 xlsx 只保留一份实现，落在共享核心

- **Status**: Accepted
- **Date**: 2026-10-06
- **Related**: dcc-mcp-office ADR-006, PIP-4281, PIP-4276

## Context

`dcc-mcp-office` v0.2.3 已经有一条 CI 绿的无头 xlsx 路径
（`skills/office-generate-production-dashboard/scripts/generate_dashboard.py`，
Python + openpyxl，在 ubuntu-latest 上跑通）。adapter 侧需要同样的能力。

薄适配层是生态约定：adapter 只带应用语义，重型实现留在共享核心。如果在
adapter 内再起一套 openpyxl writer，短期更快，长期必然语义漂移 —— 两套
实现对同一个 Workbook IR 的边界处理（ragged rows、sheet 名长度、公式
求值语义）迟早分叉，而 ADR 006 已明确「keeping exactly one COM
implementation」。

## Decision

1. **无头 xlsx writer 只有一个**：位于 `dcc-mcp-office`（共享核心或其
   `skills/`）。adapter 通过一个薄的 `host_client` 调它，不在仓内另起一套。
2. **adapter 内不实现 fallback**。共享 host 缺失时 `rpc()` 返回
   `OFFICE_HOST_NOT_FOUND` 并带上修复提示，绝不静默回退到本地 writer。
3. **本次过渡期**：`compiler.py` 直接调 openpyxl，结构与共享核心的
   dashboard 脚本对齐（同样的 `office-ir/1.0` envelope、同样的 A1 寻址）。
   共享核心 `workbook.compile` capability 落地后，`compiler.py` 换成一次
   `office.command.execute` RPC —— 调用面不变，实现位置变。
4. **能力分级写进契约**：无头可验证的标 `verified`，只有 Excel 能做
   （recalc / chart / pivot / PDF）标 `host_limited`，永不冒充 verified。

## Consequences

- 只有一份实现，语义漂移面消失。
- CI 可以在 Linux runner 上验证 1.0 门禁（写后回读），不依赖 Office 授权。
- 共享 host 未就绪时 adapter 的降级是显式的、可诊断的，不是静默的。
- 过渡期 `compiler.py` 与共享脚本存在一段代码相似期；切换点被
  `host_client.compile_workbook()` 签名固定，切换只需替换函数体。
