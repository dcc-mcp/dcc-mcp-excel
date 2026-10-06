# ADR 005 — 能力分级写进代码，不是写进文档

- **Status**: Accepted
- **Date**: 2026-10-06
- **Related**: PIP-4276, PIP-4281

## Context

Excel 有一批能力只有装了 Excel 才成立：原生公式重算、图表渲染、数据透视
刷新、分页 PDF 导出。无头 Open XML 路径写得出结构，但写不出这些值。

如果在文档里「说明」而代码里不区分，会出现两类失败：

1. agent 读 README 后向用户承诺「表里已经有算好的数」，实际是公式文本；
2. 后续维护者把 `host_limited` 能力当成已验证能力继续往上搭。

`dcc-mcp-powerpoint` 的 `render_deck` 已有先例：Office 不可用时返回显式
reason，绝不产出假产物。

## Decision

1. **`capabilities.py` 是唯一分级来源**：每条能力带
   `grade` / `summary` / `evidence` / `requires_office`，`verified` 与
   `host_limited` 两集合不重叠（有测试钉住）。
2. **分级可被机器读取**：`dcc-mcp-excel capabilities` 与
   `excel-capabilities` skill 输出同一份 `dcc-mcp-capability-grade/1` JSON。
3. **`host_limited` 能力在校验里报 warning，不报 passing check**。
   `validate_envelope` 遇到 IR 里的 chart / pivot 只声明「已声明、未验证」，
   不生成一条名为 chart 的绿色检查 —— 测试钉住这一点。
4. **缺失的 desktop Excel 是报告不是失败**：`preflight()` 的 `ok` 只反映
   无头后端，Linux runner 上不算挂。
5. **`unimplemented` 单独一档**，用于 Graph Workbook 会话这类「评估过、
   明确不做」的能力，避免与「还没做」混淆。

## Consequences

- agent 在承诺 Excel 能力前可以先查分级，而不是猜。
- 把 `host_limited` 冒充 `verified` 会直接测试失败。
- 新增能力必须带 evidence 字段（有测试钉住），分级无法退化成装饰。
