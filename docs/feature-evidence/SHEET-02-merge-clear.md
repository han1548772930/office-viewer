# SHEET-02 / SHEET-04：合并、取消合并与分类清除

记录日期：2026-11-22（Asia/Shanghai）。固定参考 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地只读路径 `E:\rust_test\univer`。开发进度：待统一验证；验收状态：**未验收**。按用户安排，本批仅实现和编写用例，没有运行测试、类型检查、构建、Rust/WASM、浏览器或性能验证。

## 1. 入口与固定源码

| 来源 | 已追踪契约与入口位置 |
| --- | --- |
| [Sheets schema](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/schema.ts) | START / LAYOUT 的合并父菜单 order 9，grid r2c3；子项顺序为全部、垂直、水平、取消合并。MAIN_AREA、ROW_HEADER、COL_HEADER 的分类清除子项为内容、格式、全部。父菜单与子项分开计数，不新增 Ribbon 根声明。 |
| [共享 UI schema](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/schema.ts) | HISTORY 的清除格式 order 3、grid r2c2，与格式刷上下排列。本地开始 / 工具中的额外“清除内容”是项目扩展，不计为上游常驻入口。 |
| [merge.menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/merge.menu.ts) | 父菜单激活态读末选区 primary 的合并状态；工作簿、工作表、值/样式及范围编辑权限和交叉选区影响禁用，绘图焦点影响显隐。 |
| [clear.menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/clear.menu.ts) | 内容、格式、全部分开判断值/样式权限，不只依赖控件是否禁用。 |
| [add-worksheet-merge.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/add-worksheet-merge.command.ts) | `AddWorksheetMergeCommand`、All/Vertical/Horizontal、`getClearContentMutationParamForRange` 和 `facadeMerge`：方向、首项内容、清非主格、重叠选项、确认和历史。`Dimension.COLUMNS = 0` 对应垂直，`ROWS = 1` 对应水平。 |
| [remove-worksheet-merge.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/remove-worksheet-merge.command.ts) | 所有与请求选区相交的合并均移除；主格样式恢复到原区域，内容不回填；模型和选择变化组织撤销事务。 |
| [clear-selection-content.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/clear-selection-content.command.ts)、[format](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/clear-selection-format.command.ts)、[all](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/clear-selection-all.command.ts) | 内容清除保留样式；格式和全部清除移除相交合并；筛选隐藏行从内容作用范围中排除，手动隐藏另行处理。 |
| [超链接拦截器](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-hyper-link/src/controllers/set-range.controller.ts)、[验证拦截器](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation/src/controllers/dv.controller.ts) | 三类清除都移除普通链接；clear-all 裁减验证范围。本次尚未确认对应条件格式 clear interceptor，不能将本地条件格式裁减写成已核实上游行为。 |
| [FRange](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/facade/f-range.ts) | `clear()` 缺省清全部；`contentsOnly` / `formatOnly` 只有恰好一个为真时单独清除。`merge/mergeAcross/mergeVertically` 缺省 defaultMerge=true、isForceMerge=false。 |

## 2. 本地模型、命令与 UI

- [sheet-range-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-command.ts) 校验显式目标及合法坐标，提供矩形裁减、筛选行排除、权限与不相交矩形的稀疏单元格划分。
- [sheet-merge-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-merge-command.ts) 提供计划、预览及 model/view 两种适配。`sheet.command.merge`、`merge-all`、`merge-horizontal`、`merge-vertical`、`unmerge` 与固定版 add/remove-worksheet-merge 别名共用行为。
- [sheet-clear-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clear-command.ts) 统一 `clear-content`、`clear-format`、`clear-all`、`clear-range` 及三种 clear-selection 别名。内容清除删除值、公式、缓存类型/显示、富文本、链接和迷你图；保留样式和未知 cell metadata，失效 formulaError。
- [sheet-range-editor-commands.ts](../../packages/office-studio/src/api/sheet-range-editor-commands.ts) 注册事务，显式 unitId/sheetId/ranges 优先，防止非活动表的 API 误改当前表；候选模型由 [sheet-range-editor-validation.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-editor-validation.ts) 验证后提交。
- [use-sheet-range-editor.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-range-editor.ts) 只管理 UI 请求及适配。Ribbon、Delete/Backspace、右键和 Facade 进入相同注册命令；独立工作台使用相同 feature 算法后交给宿主保存。
- [SheetMergeConfirmationDialog.vue](../../packages/office-studio/src/components/spreadsheet/SheetMergeConfirmationDialog.vue) 使用受控 `antdv-next` Modal / Alert，继承 ConfigProvider。取消不写历史；动画关闭后归焦。目标工作簿/工作表/运行实例切换或停用清理请求；确认保留原命令、目标、范围与 baseRevision，版本或权限变化拒绝提交并显示具体错误。

合并保留按行列顺序找到的首项非空完整单元格，包含 0、false、公式和富文本。搬移后的主格更新身份与坐标；非主格保留自身样式。独立 view 搬移使用显式 null/空边框/完整默认样式，避免 [excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) 从旧主格继承链接、metadata 或样式。取消合并恢复主格样式，保留实际边框边与显式 null，不恢复被丢弃内容；撤销可恢复合并前全部数据。

权限覆盖请求范围和全部受影响的旧合并，即使请求只含其中一格。clear-all 裁减验证及条件格式后保留范围外资源，重定位相对公式；验证按序列化的首范围，条件格式按 Canvas 的行列最小范围定位，色阶/数据条/图标阈值公式亦重定位，绝对引用和列表/复选来源不移动。

`OfficeStudio.commitWorkbook` 对未显式要求重新计算且候选未变化的事务提前返回 false，避免无操作清除重算无关公式和新增历史。显式 recalculate 仍执行；该变化影响全部工作簿事务，集中验证须覆盖各 feature 的无操作和重算场景。

## 3. 差异与未核实项

| 项目 | 本地选择与未核实边界 |
| --- | --- |
| 首项内容 | 固定 helper 实际保留首项非空内容，Facade 注释写“左上格”。本地跟随 helper；不能仅按注释断言上游实际行为。 |
| 确认触发 | 上游检查非主格有内容，本地仅真正丢失内容数量大于 0 时确认。空主格只有一个 donor 时直接搬移并合并，不提示丢失。 |
| 重叠 | Facade 维持严格默认重叠拒绝；UI 可强制替换相交旧合并。重复同形 UI 合并是无操作。 |
| 资源 | 普通链接三种清除均移除；clear-all 对条件格式裁减是本地产品选择，尚无上游 interceptor 证据。线程评论、Note、图表等其他资源删除语义仍须逐项核实，不据本切片宣称已覆盖。 |
| 取消合并边框 | 本地保留非主格已有实际边和显式 null，并将主格外围边分配到原区域外围；不同于机械复制完整主格 border。Office 独立文件效果未验证。 |
| 选择 | 合并保留实际主格并从真实模型推导合并形状；取消合并通过统一内容历史将活动范围 primary 移到起点，undo / redo 还原各自完整选择。显式 API 目标不与该表当前选区相交时保留其选择；本地非末活动范围及确认期间选择切换的差异见下文。 |

## 4. 用例与未运行的验收

[sheet-range-editor.contract.test.mjs](../../packages/office-studio/tests/sheet-range-editor.contract.test.mjs) 已编写选项/别名、输入、矩形裁减、稀疏划分、方向、多选交叉、0/false/公式/富文本、确认原子性、扩展权限、边框恢复、旧模型适配、全表稀疏清除、筛选、资源/相对公式、Facade、版本冲突、验证拒绝、历史隔离和无操作重算用例。

[sheet-range-editor.browser.test.mjs](../../packages/office-studio/tests/sheet-range-editor.browser.test.mjs) 当前有 13 个静态调用声明（一个外层 for 注册 classic/grid 两项，按源码数组预计 14 个用例），包含真实 OfficeStudio / ExcelWorkbench / Canvas2D 的确认取消/接受与归焦、编辑历史和内部重开、Delete/右键/API、筛选与手动隐藏、版本/权限变化、多选取消合并、工作簿切换、暗色 classic/grid、独立 view 保真，以及清除格式和合并的输入会话。**全部尚未运行**，循环样本和声明数量不代表通过或功能完成。

本切片新增的可替换控件均为 `antdv-next`，未引入原生表单控件。主题、焦点、系统 IME、生产 WebGPU、实际 XLSX 导出及第三方 Office 重开均未验证；功能清单保持未验收。

## 5. 性能与继续工作

矩形裁减、移除集合和稀疏划分不按所选面积逐格清除，但尚未做同数据集前后基准，不能声称优化或性能达标。非默认样式的取消合并仍可能展开整块区域生成样式空格；巨量水平/垂直合并仍生成大量矩形；导入的重叠合并在稀疏划分阶段仍拒绝，完整诊断/恢复需继续补齐。紧凑范围样式、Worker/Rust/WASM 路径、取消与内存预算及真实大样本均未验收。

<a id="sheets-clear-formatting-dispatch-review"></a>

## 2026-11-27 START/HISTORY 清除格式续核

固定 `ui.command.clear-formatting` 经 [共享格式刷 controller](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/controllers/format-painter.controller.ts) 调用 `FormatPainterSessionService.clearFormatting`，再由 Sheets host 派发 `sheet.command.clear-selection-format`。命令生成 `s: null`；[cell-style.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/basics/cell-style.ts) 的 `handleStyle` 继续清目标 richText run 样式。

合并删除须读完整 interceptor 链路：[merge-cell.controller.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/controllers/merge-cell.controller.ts) 的 `_initCommandInterceptor` 同时处理 ClearSelectionFormat / All，移除相交 merge，并登记逆向 AddWorksheetMergeMutation。只看 clear 命令主体会漏掉该行为；本地格式清除移除相交合并已有固定源码依据，不登记为差异。

本地 [sheet-range-editor-commands.ts](../../packages/office-studio/src/api/sheet-range-editor-commands.ts) 已有真实 sheet clear ID、模型事务、过滤行与权限判断、内容历史及内部保存链路；[sheet-clear-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clear-command.ts) 清作者样式、引用和已建模 run 字体，让目标恢复继承。Docs 在 Ribbon 适配中另映射到 `doc.command.clear-formatting`，OfficeStudio 尚未注册共享 `ui.command.clear-formatting`。该公开跨产品派发缺口由主线登记，本批保留有效 sheet / doc 路径，不增加仅指向 Sheets 的共享别名。

本节只补来源与差异，没有修改清除生产逻辑或增加验收结果。测试、浏览器、文件样本、真实输入法、主题 / 焦点和性能均未运行，状态仍为 **未验收**。

<a id="sheets-clear-formatting-input-session-review"></a>

## 2026-11-27 START/HISTORY 清除格式输入会话续修

本批续核 START/HISTORY 第 4 根：固定 [ui/menu/schema.ts:35](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/schema.ts#L35) 的 clear-formatting 为 classic order 3、grid r2c2；本地 `sheet.command.clear-format` 保持相同位置与单按钮。常驻、插件扩展、上下文、溢出根和子项增量均为 0。

固定 [ClearFormattingMenuItemFactory:42](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/format-painter.menu.ts#L42) 的 disabled$ 读 session.canClear；[session:81–88](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/services/format-painter/format-painter-session.service.ts#L81) 的能力与执行都依 active adapter.canStart。[Sheets host:76–83](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/format-painter/format-painter.controller.ts#L76) 复用 ClearSelectionFormatMenuItemFactory 的 disabled$；[menu-util:163/191/225](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/menu-util.ts#L163) 默认不支持单元格编辑，检查可见编辑器、FX 焦点、绘图和权限。该真实链路给出输入禁用依据，不只根据按钮名称推断。

原本地清除格式按钮只依 hasSheet 禁用，`executeRibbon` 会先 finishEdit，因此可先提交输入草稿再清样式。[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 现在复用已有 `canQuickFormatSelection`，包括活动工作台、可编辑的完整选区、单元格/公式输入、组字与绘图状态；将能力传入 [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 新增的可选内部 `clearFormatEditable`。清除格式按钮依该值禁用，`clearModeForCommand` 识别到 format 的 UI 命令在 finishEdit 前依同一能力返回，拦截本地 clear-format 和固定 clear-selection-format 两个 ID，不改变普通字体草稿能力。省略新增可选状态的原 Ribbon 调用方保留既有行为。

通过后的清除仍由 `use-sheet-range-editor` → 原注册 sheet clear 命令或 standalone view 计划完成，沿原过滤行、合并、权限、run 字体、内容历史和内部保存链路；公开 API、领域模型、命令 ID、事件字段、Office 和 Worker/WASM 协议没有新增。共享 `ui.command.clear-formatting` 尚未注册且跨产品派发缺口仍在，不能把这次 Sheets UI 接线登记为共享实现。`clear-range` 的可选 formatOnly 不属于本批两个入口的输入禁用修复。

本批只改上述两份生产文件、[原 Ribbon 合同草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)、[原范围编辑浏览器草稿](../../packages/office-studio/tests/sheet-range-editor.browser.test.mjs)及本页，生产增量为 Excel `+2/-0`、Ribbon `+2/-1`，合计 `+4/-1`。合同新增 1 个声明，直接检查真实 schema 的独立能力、缺表禁用、原位置与实际 invocation；浏览器新增 1 个声明，含 production/standalone × cell/formula × classic/grid 八个输入样本，通过真实 Input/TextArea、原生禁用点击与真实 WorkbenchRibbon 的 `execute` 事件验证草稿/模型/版本/历史不变、合成 composition、明转暗、输入结束仍禁用、Esc 取消后清除、生产 undo/redo/内部 JSON 重开和 standalone 一次 view 发布/JSON 读回。既有权限变更用例补清除格式禁用和 execute 拒绝。事件直接经真实 Vue emitter 到 Excel，未给产品新增测试接口，也不将合成事件视为 Windows IME 验收。

当前静态声明为 Ribbon 合同 25、浏览器 12（含一个为 classic/grid 注册两个用例的外层 for；按源码数组预计 13 个浏览器用例），八个输入循环样本不另计独立声明；全部运行 0。可替换控件仍为原 `antdv-next` Button/Input/TextArea，复用统一 ConfigProvider/主题；[docs/llm.md](../llm.md) 的 Button Disabled 指引已查阅。只完成源码/diff 阅读与声明计数，静态语法/引用/文档收尾由主线统一；未运行测试、类型检查、构建、dev、浏览器、Rust/WASM、Office 或性能。原生 IME、实际主题/焦点、两个渲染后端、Office 独立重开与性能仍未验收，SHEET-02/04 状态不变。

## 2026-11-27 布局组合并续核

本次重新按固定 [merge.menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/merge.menu.ts)、[merge.menu.spec.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/__tests__/merge.menu.spec.ts)、add / remove command 和 [handle-merge-operation.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/utils/handle-merge-operation.ts) 追踪布局组的合并根与四个子项。grid r2c3、classic order 9，子项仍按全部、垂直、水平、取消排列；与另外五个根的分类计数见 [布局组续核](SHEET-04-text-rotation-shrink-to-fit.md#2026-11-27-布局组续核)。

固定父菜单激活态只读末选区 `primary.actualRow / actualColumn` 在当前 worksheet 合并模型中的状态，不读可能过期的 selection merge 标志。上游测试明确覆盖：选区虽相交合并但主格普通时不激活、多选末段改变主格、合并主格及非主格、取消合并后的撤销重做、mutation、切表及无活动工作簿。本地原先用 `canUnmerge`（任一目标相交且可编辑）兼作 activated，本次主线复用 `use-sheet-range-editor.ts` 的 `merges` 与 `sheetRangeContainsCell` 增加 `primary()` / `mergeActive`，从 ExcelWorkbench 的 selectedKey 传入实际主格；`sheet-ribbon.ts` 仅将 activated 接到 mergeActive，保留 canMerge / canUnmerge 和真实命令协议。

执行核对继续保留既有实现与第 3 节的显式差异，没有改合并领域算法。全部 / 水平 / 垂直按模式生成矩形，重叠选项、首项内容及清除非主格进入同一个 model / view 计划；取消命令移除全部相交合并并恢复样式，真实事务提供撤销、重做与内部保存重开。权限覆盖目标及受影响旧合并，不用 activated 代替执行权限。固定父菜单在绘图焦点时隐藏，本地 `sheet-ribbon.ts` 已有 `hidden: state.drawingFocused`，原合同也覆盖五个布局根的绘图隐藏；先前“没有同等 hidden 接线”的记录有误，现已纠正。本地子项按可合并 / 可取消状态单独禁用，固定子项没有相同逐项 disabled 声明，该逐项能力判断仍为保留差异。

选择契约已读到具体来源：固定 `AddMergeRedoSelectionsOperationFactory` 保留原选区与 actual 主格并重建末选区的合并 primary，undo 恢复此前 selections；固定取消命令把末选区的 redo primary 改为该选区左上单格，undo 恢复原 selections。本地现已补取消合并的选择编排，UI 继续从真实模型重算 mergeActive；实现与保留差异见下一节，运行验收仍后置。

主线在 [sheet-ui-ribbon.contract.test.mjs](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs) 及 [sheet-cell-layout.browser.test.mjs](../../packages/office-studio/tests/sheet-cell-layout.browser.test.mjs) 补 activated / 真实工作台 primary 草稿。按既定安排未运行测试、类型、构建、Rust/WASM、浏览器或 Office 样本；主题、焦点、IME、完整资源联动与性能仍未验收。本轮包含工作台状态及取消合并选择历史接线，不能将其作为 SHEET-02 / SHEET-04 整体验收。

### 合并选择与统一内容历史

[office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 的 `commitWorkbook` 已持有完整 `selections.before / after`，统一历史执行时同步恢复模型与各自选择，并在内容事务结束后发布选择事件。本次复用该机制，没有增加选择历史、修改 selection store 或扩展公共输入。合并 Add 路径继续使用默认选择快照保留所有范围、activeRangeIndex 与 actual primary；[facade.ts](../../packages/office-studio/src/api/facade.ts) 的 `FWorksheet.getPrimaryCell` 从当时的 mergedRanges 推导合并形状及主格标志，撤销重做重新读取实际模型。

[sheet-range-editor-commands.ts](../../packages/office-studio/src/api/sheet-range-editor-commands.ts) 仅在取消合并注册中增加选择编排：当前选区与实际移除的合并相交时，将活动范围起点作为 after.primary，before 保留原完整多选和 actual primary。成功取消与内容变更进入同一历史项；undo 恢复原 primary / 范围，redo 恢复起点 / 范围。确认待定、取消、版本失败、权限失败和没有实际移除的无操作都不发布该 after 选择。

保留差异：固定版只在末选区持有 primary，本地允许非末 activeRangeIndex，取消后使用该活动范围起点并保留其他选区。固定取消命令即使显式目标不相交当前选择，也会重建当前末选区的 primary；本地仅编排与实际移除合并相交的选择，保留无关活动表及目标表的已保存选择。固定 Add helper 直接持有 selection 中的合并旗标，本地由模型推导真实形状，不保留过期旗标。确认请求仍捕获原目标范围与版本；接受时若仅选择发生变化，本地保存接受时的实际选择，不用捕获范围覆盖后来选择。隐藏 / 筛选轴继续遵循已有可见主格规则，未据此宣称所有上游选择细节或第三方 Office 行为都已验收。

[sheet-range-editor.contract.test.mjs](../../packages/office-studio/tests/sheet-range-editor.contract.test.mjs) 新增 3 个真实 Studio / Facade 合同草稿，分别包含三种合并模式的实际主格与形状、末段 / 非末活动范围的取消及完整 undo / redo、无相交显式 API 与非活动表选择隔离；扩展既有确认用例覆盖待定 / 取消 / 版本失败的选择原子性，以及选择切换后接受捕获范围的历史。主线已在独立 view 的 `context.applyView(candidate)` 与 `context.changed(...)` 之间，仅成功取消时调用 `resetPrimary()`；ExcelWorkbench 以实际 activeSelectionIndex 设置 selectedKey / selectionAnchor 到起点，归入同一次宿主历史。真实独立工作台的成功 / Add / 失败浏览器草稿由主线补充，运行验证继续后置，本节为实现与静态源码证据。

<a id="sheets-merge-input-session-review"></a>

## 2026-11-27 布局组合并输入会话续修

本批继续合并父菜单与全部 / 垂直 / 水平 / 取消四个子项，保持 grid r2c3、classic order 9 和既有顺序。常驻、插件扩展、上下文、溢出根和子项计数增量均为 0。绘图焦点隐藏此前已接线，本批只纠正上节过期证据，没有重复添加显隐实现。

固定 [CellMergeMenuItemFactory:47–55](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/merge.menu.ts#L47) 的 editDisabled 复用 `getCurrentRangeDisable$`；[menu-util:163/191–192](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/menu-util.ts#L163) 缺省 supportCellEdit=false，检查可见单元格编辑器和 FX 焦点，[225–226](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/menu-util.ts#L225) 继续检查绘图。父菜单还合并交叉选区与排他范围的禁用；本地沿已有 canMerge / canUnmerge 和权限判断，不宣称已完全复刻这些上游状态。

原本地父项和四个子项只依 fontEditable / canMerge / canUnmerge，普通单元格或公式输入时仍可用，`executeRibbon` 会先 finishEdit 提交草稿再打开合并确认。[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 现复用 `canQuickFormatSelection`，包括活动工作台、完整选区权限、单元格/公式输入、组字与绘图状态；传入 [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 的可选内部 `mergeEditable`，使父项与四子项共同禁用。UI `executeRibbon` 在 finishEdit 前使用同一能力拦截 [SHEET_MERGE_COMMAND_IDS](../../packages/office-studio/src/features/sheet-workbench/sheet-merge-command.ts) 的 5 个本地命令和 6 个别名，迟到 execute 不提交草稿、不建确认请求、不改模型或历史。省略可选状态的原 Ribbon 调用方保留既有行为。

输入结束后的合并仍由 `use-sheet-range-editor` → 原 model / view 计划 → 注册命令和统一内容历史执行，确认取消/接受、首项 donor、真实 primary 的 activated、取消合并选择编排、权限与保存路径保持原契约。普通字体仍允许已有草稿格式行为；程序化 OfficeStudio / Facade 合并没有新增 UI 输入限制。公开 API、命令 ID、模型、事件字段、内容历史、Office 和 Worker/WASM 协议均无新增，只有可选内部 UI 能力。固定子项没有同样的逐项 disabled，本地保留 canMerge / canUnmerge 子项限制，未以此声称全量一致。

本批共上述两份生产文件、[原 Ribbon 合同草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)、[原范围编辑浏览器草稿](../../packages/office-studio/tests/sheet-range-editor.browser.test.mjs) 和本页五文件；生产增量 Excel `+2/-0`、Ribbon `+7/-5`，合计 `+9/-5`。合同新增 1 个声明，检查五项禁用、可选状态、省略状态行为、缺表/字体权限、原位置、activated 和 canMerge / canUnmerge 的不同组合。浏览器新增 1 个声明，包含 production/standalone × cell/formula × classic/grid 八个循环样本，使用原 Input/TextArea、原生 disabled 点击和真实 WorkbenchRibbon 的 Vue execute emitter 覆盖全部 11 个 ID；断言普通输入/合成 composition/明转暗时草稿、模型、版本、历史、确认请求和 valueChanges 不变，composition 结束仍编辑时继续禁用，Esc 后真实确认一次，生产 undo/redo/内部 JSON 重开和 standalone 一次 view 发布/JSON 读回。未给产品新增测试接口，standalone 不以 OfficeStudio reopen 冒充宿主重开，也不把合成 composition 当作 Windows IME 验收。

当前静态声明为 Ribbon 合同 26、浏览器 13（一个外层 for 注册两个变体，按源码数组预计 14 个浏览器用例），全部运行 0。可替换控件仍为既有 `antdv-next` Button/Input/TextArea/Modal，复用统一 ConfigProvider 与主题，没有新表单或输入桥。仅做源码/diff 阅读与 rg 计数，静态语法/引用/文档收尾由主线统一；未运行测试、类型检查、构建、dev、浏览器、Rust/WASM、Office 或基准。真实系统 IME、明暗主题/焦点运行、WebGPU/Canvas2D、Office 独立文件重开和性能仍 **未验收**，SHEET-02/04 整体状态不变。

<a id="sheets-clear-content-input-session-review"></a>

## 2026-11-27 START/工具 清除内容输入会话续修

重新读取固定 [clear.menu.ts:45](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/clear.menu.ts#L45) 的内容清除 factory：值编辑权限与排他范围影响禁用；其 [menu-util.ts:163/191–192](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/menu-util.ts#L163) 缺省 supportCellEdit=false，单元格编辑器可见或公式编辑焦点时禁用。[内容清除命令](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/clear-selection-content.command.ts#L36) 仍按可见范围、mutation 与 undo/redo 工作。本地开始 / 工具中的“清除内容”沿用此功能语义，属于既有项目扩展根，不声称固定版有同名 Ribbon 常驻根。本批新增根 0；当前工具组五根按数组顺序为保护范围、快捷键面板、搜索功能、清除内容、范围主题；清除内容保持第 4 项、grid r1c3 跨两行，classic 顺序不变。

原扩展根仅由 hasSheet 控制，普通单元格或公式输入时的 UI execute 会先 finishEdit 提交草稿，再清空选区。[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 现向 [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 传可选内部 clearContentEditable=canQuickFormatSelection，根项结合缺表禁用；原调用方省略此状态时保留既有行为。提交前守卫由格式两项扩到 [SHEET_CLEAR_COMMAND_IDS](../../packages/office-studio/src/features/sheet-workbench/sheet-clear-command.ts) 的全部 7 个 ID，包含内容、格式、全部、clear-range 与 3 个别名。迟到 UI execute 不提交普通/组字草稿、不清值或格式，不改变模型、revision、历史或 standalone view。绘图焦点与权限沿既有 canQuickFormatSelection 判断，不增加上游没有的绘图隐藏。

输入退出后的内容清除仍经 [use-sheet-range-editor.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-range-editor.ts) → model/view 算法 → [注册事务](../../packages/office-studio/src/api/sheet-range-editor-commands.ts)；保留值/公式/缓存清理、样式及未知 metadata、筛选行排除、资源裁减、权限、内容历史和内部保存。程序化 OfficeStudio / Facade、Delete/Backspace 和右键的既有入口没有新增本 UI 会话限制；值/样式细分权限与排他范围的完整差异仍沿前文保留。仅修改同步 UI 能力和分发顺序，没有更换模型权威来源、公开命令/API/事件、历史格式或 Worker/WASM 协议。

本批五文件：上述两份生产文件、[原 Ribbon 合同草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)、[原范围编辑浏览器草稿](../../packages/office-studio/tests/sheet-range-editor.browser.test.mjs) 和本页。生产增量 Excel +2/-1、Ribbon +2/-1，合计 +4/-2。合同新增 1 声明，检查可选状态、独立字体能力、缺表、五根顺序、位置及真实 invocation；浏览器新增 1 声明，在 production/standalone × cell/formula × classic/grid 八个循环样本中覆盖全部 7 个 ID 与 clear-range 的两种显式选项，断言普通输入、合成 composition、明转暗及组字结束仍编辑时不写内容/历史，Esc 后真实清值、样式与邻格保留、一次生产历史/standalone 发布、生产 undo/redo/内部 JSON 重开和 standalone JSON 读回。没有新增产品测试接口。当前 rg 计数为 Ribbon 29 声明/29 计划、范围编辑浏览器 14 声明/15 计划（其中 1 个外层 for 注册 classic/grid 两例）；八个输入循环样本不另计声明，全部运行 0。

本批可替换 UI 仍由既有 antdv-next Button/Input/TextArea/Modal 与统一 ConfigProvider 承载，没有新增表单或原生输入桥。只做源码/diff 阅读与 rg 计数；静态语法/引用收尾由主线统一，未运行测试、类型检查、构建、dev、浏览器、Rust/WASM、Office 或基准。合成组字不代表原生 Windows IME；内部 JSON 不代表 Office 文件往返。真实明暗主题/焦点、IME、WebGPU/Canvas2D、Office 重开及性能均保留 **未验收**，SHEET-02/04 总状态不变。
