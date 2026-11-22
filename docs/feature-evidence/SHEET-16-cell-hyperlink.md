# SHEET-16 单元格超链接

## 固定版依据

固定来源为 `E:\rust_test\univer`，版本 `v1.6.0`，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

- `packages/sheets-hyper-link-ui/src/menu/schema.ts` 将工具栏入口注册在 `RibbonInsertGroup.MEDIA`，布局为 `row: 1, column: 2, rowSpan: 2`；同时注册主区域右键菜单入口。
- `packages/sheets-hyper-link-ui/src/menu/menu.ts` 注册 `InsertLinkShortcut`，快捷键为 Ctrl/Cmd+K，并用 `getLinkDisable$` 判断当前单元格权限与状态。
- `packages/sheets-hyper-link-ui/src/utils/index.ts:getShouldDisableCellLink` 禁止公式、数组公式、含内容控件及 checkbox/list/listMultiple 验证单元格；其他验证单元格会经过写入验证。
- `packages/sheets-hyper-link-ui/src/views/CellLinkEdit.tsx:CellLinkEdit` 展示 display、tooltip、类型与 payload，并支持 URL、单元格范围、工作表、定义名称等目标及富文本选区。
- `packages/sheets-hyper-link/src/commands/commands/add-hyper-link.command.ts:AddHyperLinkCommand` 将链接写为 `CustomRangeType.HYPERLINK` 富文本并组织单元格与链接模型 mutation 的撤销重做。
- `packages/sheets-hyper-link/src/types/interfaces/i-hyper-link.ts` 定义链接 payload、display、tooltip、位置和 id。
- `packages/engine-render/src/components/docs/layout/style/custom-range.ts:getCustomRangeStyle` 默认使用 `blue.600` 并绘制下划线；`textStyleMode`/`textColorMode` 会保留文本自身样式。本项目当前仅复现默认链接外观，颜色暂使用 `#0563c1`，尚未实现这些富文本样式开关及画布主题映射。

## 当前实现

Ribbon 插入入口位于 MEDIA 分组，支持 Ctrl/Cmd+K、外部 HTTP(S)/mailto 与工作簿内工作表单元格范围目标。单元格链接以 `#gid=<sheetId>&range=<A1-range>` 保存，可在编辑弹窗中更新/移除，通过右键菜单、Enter 或 Ctrl/Cmd+点击导航。内部导航校验目标工作表可见性与查看权限，切换工作表、选择目标范围并滚动至可视区域。工作簿模型和 Excel 视图适配保留该 payload，XLSX writer 将有效目标映射为标准 worksheet `location`。

新增与移除经 `OfficeStudio` 工作簿命令进入内容历史，并执行工作表/范围保护和数据验证写入校验。链接文字显示为蓝色并带下划线。`FWorkbook.navigateToSheetHyperlink` 提供内部工作表导航 API，对应上游 `packages/sheets-hyper-link-ui/src/facade/f-workbook.ts`。上游桌面悬停预览与权限检查见 `packages/sheets-hyper-link-ui/src/controllers/popup.controller.ts`。

文件入口为 `packages/office-studio/src/features/sheet-workbench/sheet-hyperlink.ts`、`packages/office-studio/src/api/office-studio.ts`、`packages/office-studio/src/services/excelViewAdapter.ts`、`packages/office-studio/src/rendering/spreadsheet/scene.ts`、`packages/office-studio/src/components/ExcelWorkbench.vue` 与 `packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts`。

## 差异与验收

本地模型仍是单元格级链接字段，不具备 Univer `CustomRangeType.HYPERLINK` 的局部富文本链接、tooltip、多条链接/custom range 或悬停预览编辑弹层。内部目标目前支持工作表 ID + A1 单元格/矩形范围，不支持定义名称或上游 parser 的其它内部目标形式。公式禁用和 checkbox/list/listMultiple 验证规则禁用已对齐上游判断。shared/array formula 与 cell custom blocks 在本地模型中没有等价字段，尚未核实对应文件导入后的限制行为。单元格链接编辑 UI 使用 antdv-next Modal，悬停 popup 仍待实现。

契约覆盖外部 URL 安全规范化、公式和受限验证单元格拒绝、单元格模型视图往返、命令撤销/重做/重开、内部 payload 解析及导航选区。已执行 `sheet-hyperlink.contract.test.mjs`，5 项通过。Rust writer 的 `hyperlink` 过滤用例和 `writes_univer_internal_link_payload_as_an_xlsx_location` 精确用例分别通过，验证 Univer internal payload 可写成 XLSX worksheet location。命令/API 契约、Canvas Ctrl/Cmd+点击、浏览器焦点/滚动、实际 XLSX 样本往返和不同权限主体仍待集中验收。此功能保持未验收，不代表超链接全量对齐。
