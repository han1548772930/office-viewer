# SHEET-16 单元格超链接（外部链接切片）

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

插入 Ribbon MEDIA 分组及网格位置、Ctrl/Cmd+K、外部 HTTP(S)/mailto 整单元格链接编辑、已有链接打开/移除、Canvas 蓝色下划线渲染均有对应实现。新增与移除经 `OfficeStudio` 工作簿命令进入同一内容历史，执行工作表与范围保护校验，写入/撤销/重做继续经已有模型适配和文件写出链路。

文件入口为 `packages/office-studio/src/features/sheet-workbench/sheet-hyperlink.ts`、`packages/office-studio/src/api/office-studio.ts`、`packages/office-studio/src/services/excelViewAdapter.ts`、`packages/office-studio/src/rendering/spreadsheet/scene.ts`、`packages/office-studio/src/components/ExcelWorkbench.vue` 与 `packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts`。

## 差异与验收

当前切片只支持外部 URL 且作用于整单元格纯文本。它没有 Univer 的富文本局部链接/custom range、tooltip、多条链接模型、编辑弹层预览、单元格范围/工作表/定义名称内部导航、链接上下文菜单或可访问演示中的权限细节。公式与 checkbox/list/listMultiple 单元格明确禁用，避免覆盖公式并遵循上游判断。图像入口仍未实现。

新增契约测试覆盖外部 URL 规范化、安全协议限制、上游禁用状态、模型视图往返，以及命令新增/移除、撤销重做和模型重开。按主任务约定本次未运行测试或构建，运行行为、浏览器导航及 Office 文件样本保存/重开均未验收。不得将此切片记作超链接全量对齐。
