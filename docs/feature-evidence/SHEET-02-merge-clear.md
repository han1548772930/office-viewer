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
| 选择 | 本地确认复用捕获的原范围，普通命令不复制上游 selection interceptor 的内部实现；API 不暗中改当前选择。全部上游 redo selection 细节仍待核实。 |

## 4. 用例与未运行的验收

[sheet-range-editor.contract.test.mjs](../../packages/office-studio/tests/sheet-range-editor.contract.test.mjs) 已编写选项/别名、输入、矩形裁减、稀疏划分、方向、多选交叉、0/false/公式/富文本、确认原子性、扩展权限、边框恢复、旧模型适配、全表稀疏清除、筛选、资源/相对公式、Facade、版本冲突、验证拒绝、历史隔离和无操作重算用例。

[sheet-range-editor.browser.test.mjs](../../packages/office-studio/tests/sheet-range-editor.browser.test.mjs) 编写 10 组真实 OfficeStudio / ExcelWorkbench / Canvas2D 用例：确认取消/接受与归焦、编辑历史和内部重开、Delete/右键/API、筛选与手动隐藏、版本/权限变化、多选取消合并、工作簿切换、暗色 classic/grid、独立 view 保真。**全部尚未运行**，用例数量不代表通过或功能完成。

本切片新增的可替换控件均为 `antdv-next`，未引入原生表单控件。主题、焦点、系统 IME、生产 WebGPU、实际 XLSX 导出及第三方 Office 重开均未验证；功能清单保持未验收。

## 5. 性能与继续工作

矩形裁减、移除集合和稀疏划分不按所选面积逐格清除，但尚未做同数据集前后基准，不能声称优化或性能达标。非默认样式的取消合并仍可能展开整块区域生成样式空格；巨量水平/垂直合并仍生成大量矩形；导入的重叠合并在稀疏划分阶段仍拒绝，完整诊断/恢复需继续补齐。紧凑范围样式、Worker/Rust/WASM 路径、取消与内存预算及真实大样本均未验收。
