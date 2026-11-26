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

[sheet-range-editor.browser.test.mjs](../../packages/office-studio/tests/sheet-range-editor.browser.test.mjs) 编写 10 组真实 OfficeStudio / ExcelWorkbench / Canvas2D 用例：确认取消/接受与归焦、编辑历史和内部重开、Delete/右键/API、筛选与手动隐藏、版本/权限变化、多选取消合并、工作簿切换、暗色 classic/grid、独立 view 保真。**全部尚未运行**，用例数量不代表通过或功能完成。

本切片新增的可替换控件均为 `antdv-next`，未引入原生表单控件。主题、焦点、系统 IME、生产 WebGPU、实际 XLSX 导出及第三方 Office 重开均未验证；功能清单保持未验收。

## 5. 性能与继续工作

矩形裁减、移除集合和稀疏划分不按所选面积逐格清除，但尚未做同数据集前后基准，不能声称优化或性能达标。非默认样式的取消合并仍可能展开整块区域生成样式空格；巨量水平/垂直合并仍生成大量矩形；导入的重叠合并在稀疏划分阶段仍拒绝，完整诊断/恢复需继续补齐。紧凑范围样式、Worker/Rust/WASM 路径、取消与内存预算及真实大样本均未验收。

## 2026-11-27 布局组合并续核

本次重新按固定 [merge.menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/merge.menu.ts)、[merge.menu.spec.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/__tests__/merge.menu.spec.ts)、add / remove command 和 [handle-merge-operation.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/utils/handle-merge-operation.ts) 追踪布局组的合并根与四个子项。grid r2c3、classic order 9，子项仍按全部、垂直、水平、取消排列；与另外五个根的分类计数见 [布局组续核](SHEET-04-text-rotation-shrink-to-fit.md#2026-11-27-布局组续核)。

固定父菜单激活态只读末选区 `primary.actualRow / actualColumn` 在当前 worksheet 合并模型中的状态，不读可能过期的 selection merge 标志。上游测试明确覆盖：选区虽相交合并但主格普通时不激活、多选末段改变主格、合并主格及非主格、取消合并后的撤销重做、mutation、切表及无活动工作簿。本地原先用 `canUnmerge`（任一目标相交且可编辑）兼作 activated，本次主线复用 `use-sheet-range-editor.ts` 的 `merges` 与 `sheetRangeContainsCell` 增加 `primary()` / `mergeActive`，从 ExcelWorkbench 的 selectedKey 传入实际主格；`sheet-ribbon.ts` 仅将 activated 接到 mergeActive，保留 canMerge / canUnmerge 和真实命令协议。

执行核对继续保留既有实现与第 3 节的显式差异，没有改合并领域算法。全部 / 水平 / 垂直按模式生成矩形，重叠选项、首项内容及清除非主格进入同一个 model / view 计划；取消命令移除全部相交合并并恢复样式，真实事务提供撤销、重做与内部保存重开。权限覆盖目标及受影响旧合并，不用 activated 代替执行权限。固定父菜单在绘图焦点时隐藏，本地没有同等 hidden 接线；本地子项按可合并 / 可取消状态单独禁用，固定子项没有相同逐项 disabled 声明。这些 UI 差异仍未对齐。

选择契约已读到具体来源：固定 `AddMergeRedoSelectionsOperationFactory` 保留原选区与 actual 主格并重建末选区的合并 primary，undo 恢复此前 selections；固定取消命令把末选区的 redo primary 改为该选区左上单格，undo 恢复原 selections。本地现已补取消合并的选择编排，UI 继续从真实模型重算 mergeActive；实现与保留差异见下一节，运行验收仍后置。

主线在 [sheet-ui-ribbon.contract.test.mjs](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs) 及 [sheet-cell-layout.browser.test.mjs](../../packages/office-studio/tests/sheet-cell-layout.browser.test.mjs) 补 activated / 真实工作台 primary 草稿。按既定安排未运行测试、类型、构建、Rust/WASM、浏览器或 Office 样本；主题、焦点、IME、完整资源联动与性能仍未验收。本轮包含工作台状态及取消合并选择历史接线，不能将其作为 SHEET-02 / SHEET-04 整体验收。

### 合并选择与统一内容历史

[office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 的 `commitWorkbook` 已持有完整 `selections.before / after`，统一历史执行时同步恢复模型与各自选择，并在内容事务结束后发布选择事件。本次复用该机制，没有增加选择历史、修改 selection store 或扩展公共输入。合并 Add 路径继续使用默认选择快照保留所有范围、activeRangeIndex 与 actual primary；[facade.ts](../../packages/office-studio/src/api/facade.ts) 的 `FWorksheet.getPrimaryCell` 从当时的 mergedRanges 推导合并形状及主格标志，撤销重做重新读取实际模型。

[sheet-range-editor-commands.ts](../../packages/office-studio/src/api/sheet-range-editor-commands.ts) 仅在取消合并注册中增加选择编排：当前选区与实际移除的合并相交时，将活动范围起点作为 after.primary，before 保留原完整多选和 actual primary。成功取消与内容变更进入同一历史项；undo 恢复原 primary / 范围，redo 恢复起点 / 范围。确认待定、取消、版本失败、权限失败和没有实际移除的无操作都不发布该 after 选择。

保留差异：固定版只在末选区持有 primary，本地允许非末 activeRangeIndex，取消后使用该活动范围起点并保留其他选区。固定取消命令即使显式目标不相交当前选择，也会重建当前末选区的 primary；本地仅编排与实际移除合并相交的选择，保留无关活动表及目标表的已保存选择。固定 Add helper 直接持有 selection 中的合并旗标，本地由模型推导真实形状，不保留过期旗标。确认请求仍捕获原目标范围与版本；接受时若仅选择发生变化，本地保存接受时的实际选择，不用捕获范围覆盖后来选择。隐藏 / 筛选轴继续遵循已有可见主格规则，未据此宣称所有上游选择细节或第三方 Office 行为都已验收。

[sheet-range-editor.contract.test.mjs](../../packages/office-studio/tests/sheet-range-editor.contract.test.mjs) 新增 3 个真实 Studio / Facade 合同草稿，分别包含三种合并模式的实际主格与形状、末段 / 非末活动范围的取消及完整 undo / redo、无相交显式 API 与非活动表选择隔离；扩展既有确认用例覆盖待定 / 取消 / 版本失败的选择原子性，以及选择切换后接受捕获范围的历史。主线已在独立 view 的 `context.applyView(candidate)` 与 `context.changed(...)` 之间，仅成功取消时调用 `resetPrimary()`；ExcelWorkbench 以实际 activeSelectionIndex 设置 selectedKey / selectionAnchor 到起点，归入同一次宿主历史。真实独立工作台的成功 / Add / 失败浏览器草稿由主线补充，运行验证继续后置，本节为实现与静态源码证据。
