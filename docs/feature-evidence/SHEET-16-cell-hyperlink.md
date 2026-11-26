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

<a id="media-hyperlink-type-review-2026-11-27"></a>

## 2026-11-27 插入媒体组复核

本批修正链接文字的存储类型。原实现把数值或布尔值改成显示文字后仍复制原 `dataType`，现统一写入 `dataType: 'text'`，并按本项目非公式值写入契约清除 `cachedDataType`。数字、布尔及以 `=` 开头的链接标签都保持文字；移除链接保留文字和类型。

固定版入口与内容命令分别计数，不能把这些数量当作功能完成证明：

| 类别 | 数量与入口 | 来源 |
| --- | --- | --- |
| 常驻 Ribbon | 1；MEDIA 第 1 行、第 2 列，跨 2 行，order 1；`sheet.operation.insert-hyper-link-toolbar` | [schema.ts](../../../univer/packages/sheets-hyper-link-ui/src/menu/schema.ts)、[popup.operations.ts](../../../univer/packages/sheets-hyper-link-ui/src/commands/operations/popup.operations.ts) |
| 上下文 | 1；MAIN_AREA / OTHERS，order 0，共用上述 operation | 同一 schema |
| 快捷键 | 1；Ctrl/Cmd+K | [menu.ts](../../../univer/packages/sheets-hyper-link-ui/src/menu/menu.ts) |
| 内容命令 | 添加、更新、取消 3 个；富文本编辑另有 3 个对应命令 | 下表 |

| 固定版命令 ID | 参数 | 来源 |
| --- | --- | --- |
| `sheets.command.add-hyper-link` | `{ unitId, subUnitId, link: { id, row, column, payload, display?, tooltip? } }` | [add-hyper-link.command.ts](../../../univer/packages/sheets-hyper-link/src/commands/commands/add-hyper-link.command.ts) |
| `sheets.command.update-hyper-link` | `{ unitId, subUnitId, id, row, column, payload: { payload, display?, tooltip? } }` | [update-hyper-link.command.ts](../../../univer/packages/sheets-hyper-link/src/commands/commands/update-hyper-link.command.ts) |
| `sheets.command.cancel-hyper-link` | `{ unitId, subUnitId, id, row, column }` | [remove-hyper-link.command.ts](../../../univer/packages/sheets-hyper-link/src/commands/commands/remove-hyper-link.command.ts) |

上述 3 个命令写入单元格时均指定 `CellValueType.STRING`，本批据此修正本地字符串类型；本地缓存清除规则来自 [model-utils.ts](../../packages/office-studio/src/api/model-utils.ts) 的 `applyApiCellTypes`，没有增加从 feature 到 API 的倒置依赖。本地新增/更新继续共用 [setSheetCellHyperlink](../../packages/office-studio/src/features/sheet-workbench/sheet-hyperlink.ts)，移除仍使用已有内容命令及历史。

内部目标仍有差异：[固定 parser](../../../univer/packages/sheets-hyper-link/src/services/parser.service.ts) 支持 `gid`、`range`、`rangeid`、`unitid`，本地只支持当前工作簿的工作表和 A1 范围，不具备定义名称或跨工作簿目标。此入口针对单元格，不能据此声称绘图对象超链接已实现。局部富文本、多链接、tooltip 和悬停编辑仍未验收。

新增 [模型/视图合同草稿](../../packages/office-studio/tests/sheet-hyperlink.contract.test.mjs) 与 [命令合同草稿](../../packages/office-studio/tests/sheet-hyperlink-command.contract.test.mjs)，覆盖数值/布尔旧类型、公式形状的显示文字、缓存类型清除、移除、撤销重做和快照重开。本批没有运行测试、类型检查、Rust/WASM、浏览器或 XLSX 样本；上文历史通过记录不代表新增用例已经通过。UI 未改动，主题、焦点与输入法仍留待统一浏览器验收。
