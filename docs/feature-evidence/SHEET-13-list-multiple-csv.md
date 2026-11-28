# SHEET-13 多选列表 CSV 校验

状态：2026-11-27 修正多选校验的 CSV 解析错接；2026-11-28 已写入专用列表内容命令、受控多选弹层、搜索/顺序/清除/重开态、Canvas 标签和规则项/颜色编辑，见 [完整控件记录](SHEET-13-validation-cell-input-2026-11-28.md)。新草稿均未运行，SHEET-13 未验收。固定来源为 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

## 固定版行为

[ListMultipleValidator](../../../univer/packages/sheets-data-validation/src/validators/list-multiple-validator.ts) 继承 ListValidator，并关闭普通字体渲染。[ListValidator](../../../univer/packages/sheets-data-validation/src/validators/list-validator.ts) 的 `parseCellValue` 使用 `deserializeListOptions`；单选比较完整标签，多选逐项校验，内联和引用候选均忽略大小写匹配。

[serializeListOptions / deserializeListOptions](../../../univer/packages/sheets/src/basics/utils.ts) 使用 JSON 字符串数组保存选项，读取时先解析 JSON，旧值才回退逗号拆分。固定版内联规则和多选单元格都使用这套 JSON 表示，不能把本地 CSV 声称为同一保存协议。

[dropdown-manager.service.ts](../../../univer/packages/sheets-data-validation-ui/src/services/dropdown-manager.service.ts) 的列表分支把多选值序列化后写入命中的单格，同时清除富文本和公式；多选保存返回 false，使弹层保持打开。重开时按忽略大小写匹配，把存值映射回候选的原标签。[ListDropDown.tsx](../../../univer/packages/sheets-ui/src/views/dropdown/ListDropDown.tsx) 点击候选会切换所选项，再按完整候选列表的顺序输出，去重并移除已不在候选中的项。多选默认值读取见 [list-dropdown/utils.ts](../../../univer/packages/sheets-ui/src/views/dropdown/list-dropdown/utils.ts)，标签绘制见 [dropdown-multiple-widget.ts](../../../univer/packages/sheets-data-validation-ui/src/views/widgets/dropdown-multiple-widget.ts)。

## 本地协议与小修

[validation-engine.ts](../../packages/office-studio/src/features/sheet-validation/validation-engine.ts) 已有的 parser / serializer 服务于本地公式列表，协议如下。

| 数据位置 | 表示与调用 |
| --- | --- |
| 内联 `formula1` | `serializeValidationListItems(labels)`：逐项 CSV 转义后再包一对外层引号 |
| 多选单元格值 | 原始 CSV；序列化调用为 `serializeValidationListItems(labels).slice(1, -1)` |
| 解析多选单元格 | 先给原始 CSV 包一对外层引号，再传给 `parseValidationListItems`；parser 只移除这层公式包装，保留字段引号供 CSV 解析 |
| 引用候选 | `formula1` 使用 `=$A$1:$A$4` 等范围引用；由现有候选查询读取单元格值 |

例如单个 `Green, dark` 标签的单元格 CSV 为 `"Green, dark"`，公式列表为 `""Green, dark""`。直接把单元格 CSV 传给现有 parser 会误把字段引号当成公式包装移除，标签再次被拆为两项。

2026-11-27 把 `listMultiple` 分支的 `String(actual).split(',')` 替换为现有 parser，并在调用处说明包装层。逗号、转义引号和标签自身的首尾引号由同一 CSV 解析处理；单选仍比较整个标签，大小写归一化和逐项候选匹配沿用现有逻辑。这项历史校验小修没有新增 parser 或 JSON 兼容层；2026-11-28 的专用 planner/owner/UI 同样复用此 CSV 协议。

真实内容入口为 [sheet-cell-input-commands.ts](../../packages/office-studio/src/api/sheet-cell-input-commands.ts)，调用 [sheet-cell-input-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-input-command.ts) 的输入解析、权限和数据验证，再由统一提交写入工作簿历史。候选规则更新使用既有 Facade 数据验证命令。

## 草稿与静态检查

[sheet-quick-validation.contract.test.mjs](../../packages/office-studio/tests/sheet-quick-validation.contract.test.mjs) 新增两组真实命令草稿，分别从内联和区域引用候选开始。候选为 `Red`、`Green, dark`、`It"s blue`、`"Quoted"`，覆盖单个含逗号标签、单个含引号标签、标签自身带首尾引号、全部四项和大小写混合项。输入使用原始 CSV，直接调用 `sheet.command.commit-cell-input` 并检查实际存值与验证结果。

草稿还将候选更新为新内联规则，移除 `Green, dark` 并加入 `Approved`，检查现有内容变为非法、拒绝被移除标签且快照不变、接受新多项、规则和内容各一条历史、两次 undo / redo，以及原生 WorkbookModel 保存重开后保留新值和候选。原生重开不代表 XLSX 往返验收。

本轮仅静态复读生产差异、真实命令草稿、固定来源、文档链接和限定 `git diff --check`。测试、类型、构建、Rust/WASM、dev、浏览器与性能检查均后置，没有本轮运行通过记录。

## 未完成与差异

[SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) 已接专用 Ant 列表覆盖层，多选直接调用 typed 选择命令，按完整候选顺序序列化原始 CSV、去重、保持弹层并重建当前会话；单选完整标签按文本保存，重开恢复候选匹配态。存在此控件时旧 AutoComplete 候选为空，不再整段普通文字替换。实现及新契约/浏览器草稿见完整控件记录，尚未运行验收。

彩色标签绘制、搜索、清除、列表项颜色/顺序编辑与共享公式引用的动态/命名候选已写入；2026-11-27 续接见 [下拉操作与动态源记录](SHEET-13-checkbox-feedback-dropdown-2026-11-28.md)。标签自动尺寸、旋转/缩进、纯 INDEX authoring、完整数组/spill、候选数字格式和大源 Worker 仍需推进。本地 CSV 与固定版 JSON 数组有明确表示差异，未添加跨协议互操作。第三方 XLSX 样本、权限与协作、主题切换、焦点恢复、Windows IME 和大范围性能均未验收；SHEET-13 保持部分实现。
