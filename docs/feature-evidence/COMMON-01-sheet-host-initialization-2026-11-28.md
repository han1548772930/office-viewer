# COMMON-01 Sheets 宿主初始化顺序

日期：2026-11-28。状态：已修复代码，未验收。

用户报告 `Cannot access 'painterRequest' before initialization`。实际调用链为 [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 调用 [use-sheet-table-workbench.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-table-workbench.ts)，后者注册 immediate / sync watcher 时同步调用完整 `state()`，读取 `tableHostBlocked`，间接访问当时尚未初始化的 `painterRequest`。

本批把完整 `useSheetFormatPainter` 初始化前移到 `tableHostBlocked` 和 `useSheetTableWorkbench` 之前，按真实依赖顺序创建状态；不延迟同步 watcher，也不以空值或延迟定时器规避权限状态。相邻条件格式、验证和排序宿主继续在其依赖之后初始化。格式刷中的表格状态引用只用于后续焦点回调，初始同步 watcher 不读取它。

[sheet-table-selector.browser.test.mjs](../../packages/office-studio/tests/sheet-table-selector.browser.test.mjs) 新增初次挂载、独立工作台挂载、接入 Studio 后创建表格和取消的草稿，检查未抛异常、入口状态、内容和历史不变；夹具增加受控 Studio 开关。只读 SFC/模板/TypeScript 语法、mjs 语法和限定 diff 空白检查已执行，未运行测试、类型检查、构建、dev、浏览器或其他验收。新增入口数量为 0，公开模型、命令和 Worker/WASM 协议保持。
