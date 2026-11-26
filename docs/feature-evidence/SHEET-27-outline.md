# SHEET-27：行列大纲、公开 API 与结构生命周期

记录日期：2026-11-26。状态：本批局部实现已落盘，整体未验收。先静态清点入口与差异，再实现本地 Vue/TypeScript 模型和局部 Ant 控件；不安装、执行或依赖商业 SDK。

本阶段所有测试、类型、构建、浏览器、Rust/WASM、服务与性能验证后置，仅写草稿。入口数和源码阅读不作为完成证明。

## 固定入口清点

固定 OSS：`E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。以下 Pro 发布类型/ES 为项目已有只读研究缓存。

| 分类 | 标签与 ID | 分组、顺序及布局 | 状态、命令与来源 |
| --- | --- | --- | --- |
| Sheets 扩展根 1 | 组合；`sheet.toolbar.dimension-outline.group` | Data / ORGANIZATION，order 3，grid row 2 / column 2；classic 同一 selector | 当前工作簿且单一选区才显示，权限决定禁用；2 子项分别调用 `sheet.command.outline-selected-dimension-shortcut` / `sheet.command.clear-selected-dimension-outlines-shortcut`。来源：`sheets-outline-ui/lib/es/index.js` 的 `vt/yt/Ht` |
| 行头上下文根 1 | `sheet.contextMenu.dimension-outline` | ROW_HEADER / LAYOUT，order 8 | 组合行、取消行分组，2 子项；`outline-selected-rows` / `clear-selected-rows-outlines`；`xt/Ct/Mt/Ht` |
| 列头上下文根 1 | 同一菜单 ID、不同宿主 | COL_HEADER / LAYOUT，order 8 | 组合列、取消列分组，2 子项；`outline-selected-columns` / `clear-selected-columns-outlines`；`St/Ct/Mt/Ht` |
| 快捷键 2 | 组合 / 取消组合 | Windows Alt+Shift+Right / Left；Mac Ctrl/Cmd+Shift+K / J | sheet editor focused；单一整行/整列直接操作，普通单一单元格选区打开轴选择。来源 `On/kn/ft/ht` |
| Canvas 上下文控件 | 层级按钮、组折叠按钮、组线选择 | gutter level 20px / button 14px / line hit 6px；由当前层级/组数生成 | 权限、折叠祖先与范围条件；`set-dimension-outline-collapsed` 与 `set-selections`；`outline-view-model.d.ts`、`cn/yn` |

分类口径：常驻 OSS 根 0、Pro 扩展根 1、上下文根 2；三个 selector 共 6 子项，快捷键 2；未发现本 feature 专属溢出根，通用 Ribbon More 保留同一入口。动态 gutter 不与静态菜单数相加。Docs 没有本 Sheets 插件入口；不代表 Docs 其他大纲能力完成。

固定公共 [FWorksheet 类型](../../.research.local/univer-1.6.0/packages/sheets-outline/lib/types/facade/f-worksheet.d.ts) 共 6 方法：`addRowOutline`、`addColumnOutline`、`removeDimensionOutline`、`setDimensionOutlineCollapsed`、`clearDimensionOutlines`、`getDimensionOutlines`。起点 0-based，添加参数为数量，清除为 inclusive [start,end]，raw snapshot 含 unitId/subUnitId；无效添加忽略，其他编辑复用命令/历史/权限。

核心 [类型/算法](../../.research.local/univer-1.6.0/packages/sheets-outline/lib/types/common/algorithm.d.ts) 与 `lib/es/index.js` 的 `E/De/O/k/A/j/Ie`：最多 8 个嵌套组，允许相等范围作为不同层；相邻组合保留最先相邻组 ID/状态；同范围 collapsed 同步；insert/delete/move 变换，move 分裂拒绝，结构后同范围合并并 OR collapsed。`SHEET_OUTLINE_PLUGIN` 资源按 sheet 保存 outlineList。

## 实现前差异

- 当前仅 `features/sheet-outline/sheet-outline.ts` 及统一 change-outline 注册；有嵌套、基础折叠和独立 derived outlineHiddenRows/Columns。
- 当前 Data 单独分组有 8 根按钮/selector，与固定单 selector 不同。现有更多入口保留实际语义，固定来源 selector/快捷键和可复查状态需补入。
- 相邻合并更换既有 ID；同范围 collapsed 未同步；缺类型化失败、公开 snapshot/resource 适配与六项 API 接线。
- `sheet-axis-structure-resources.ts` 用一般范围拆分并分配新组 ID；需改为 outline 连续变换、预先拒绝分裂，结构去重与复制契约单列。
- 当前汇总位置是本地 Office 增强；固定 outline 类型/资源没有汇总方向字段。保留本项目模型字段并记录此差异。
- 压缩隐藏 spans、Ant 轴选择/汇总设置、owner 取消与焦点恢复、真实 Windows/Mac 快捷键/两后端和 Office 原生资源均尚待实施/验证。

以上为实施前基线，不代表本批修改后的状态。Root 负责共享域、Facade/Studio、Ribbon/工作台、Canvas/scene 和顶层清单接线；本子任务仅修改 outline feature、局部组件、专属草稿与本证据。

## 本批实际实现与边界

本子任务共涉及 **19 个文件**：11 个 feature 文件、5 个契约草稿、1 个浏览器草稿、1 个浏览器配置及本证据。仅 `sheet-outline.ts` 为既有 feature 文件，其余为本批新增。没有修改 Root 所属的共享文件，也没有重新实现上一批切片器查询。

| 文件 | 已落盘行为与接口 |
| --- | --- |
| [sheet-outline.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline.ts) | 保留相邻既有组 ID/折叠状态；同范围折叠同步；行列独立的八层嵌套、显式展开/折叠、按 ID 删除、inclusive clear；两轴汇总设置原子候选与受保护范围检查；压缩隐藏区间与旧 Set 接口并存 |
| [sheet-outline-types.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline-types.ts) | `DimensionOutlineAxis`、七类稳定原因/错误码、scope/extent/tree/resource；快照组深拷贝及复制 ID 校验 |
| [sheet-outline-structure.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline-structure.ts) | 复用共享轴 permutation；插删边界变换、整组移动、预先拒绝分裂；结构后同范围保留首 ID 并 OR collapsed；完整源组复制、新 ID、扩展字段深拷贝；Office 上限拒绝截断 |
| [sheet-outline-resources.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline-resources.ts) | 数量到 inclusive 范围；带当前 unitId/subUnitId 的防御性公开快照；`SHEET_OUTLINE_PLUGIN` 序列化/整份候选导入；空资源清除所有当前 sheet、未知 sheet 拒绝、目标 owner 重绑定、未知组字段保留；本地 summary 扩展 |
| [sheet-outline-view.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline-view.ts) | 扁平 viewport view 含 depth/parentId、summary anchor/boundary 和祖先隐藏；children 保留结构树 metadata；全轴越界 summary 返回 null 但保留 boundary；折叠原因穿过同范围中间层；确定性 level 1…maxDepth+1 |
| [sheet-outline-command.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline-command.ts) | 4 个固定 core、8 个固定 UI、3 个本地命令；同一模型候选计算轴 prompt/capability；单选/多选/整行列语义；owner 字段透传及类型校验 |
| [sheet-outline-menu.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline-menu.ts) | 固定 Ribbon selector、行列头 context 根及两个快捷键 descriptor；单选、权限、交叉分组/clear no-op 决定状态；实际控件交由共享 Ant Ribbon/菜单渲染 |
| [sheet-outline-panel-owner.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline-panel-owner.ts) | owner 捕获、发布前重查权限与候选、真实 change-outline 命令；预览/提交复用同一 groupId；仅自己提交后的精确 revision+1 和候选结果可结束旧交互 |
| [use-sheet-outline-dialog.ts](../../packages/office-studio/src/features/sheet-outline/use-sheet-outline-dialog.ts) | 草稿代次、关闭/重开/换 host/观察者重入、内容/Undo/Redo/历史失效、显式重新读取；连接节点和合法 owner 下 nextTick 恢复焦点；scope dispose 释放订阅 |
| [SheetOutlineActionDialog.vue](../../packages/office-studio/src/features/sheet-outline/SheetOutlineActionDialog.vue) | 真实 group/remove 轴选择及错误/权限/crossing 提示；props `open/studio/workbookId/sheetId/ranges/operation`；`update:open` 与无参 `changed` |
| [SheetOutlineSettingsDialog.vue](../../packages/office-studio/src/features/sheet-outline/SheetOutlineSettingsDialog.vue) | 行、列汇总方向草稿和一次 summary-settings 提交；props `open/studio/workbookId/sheetId`；`update:open` 与无参 `changed` |

共享接线应使用以下接口，不能把本子任务的候选 helpers 当成已经验收的 Studio/Facade/Canvas：

```ts
type ChangeSheetOutlineInput = {
  unitId: string; sheetId: string; action: SheetOutlineAction;
  expectedRevision?: number;
  expectedHistoryGeneration?: number;
  expectedUserId?: string;
  host?: 'api' | 'workbench';
}
resolveSheetOutlineCommand(id, value, context)
getSheetOutlinePromptChoices(context, operation)
createSheetOutlineRibbonItem(context)
createSheetOutlineContextMenuItem(context, axis)
transformSheetOutlineForAxisEdit(state, edit, beforeExtent, afterExtent?, allocateId?)
getSheetOutlineHiddenSpans(state, axis, count)
getSheetOutlineViewModels(state, axis, count, includeHidden?)
getSheetOutlineLevelControls(state, axis)
serializeSheetOutlineResource / deserializeSheetOutlineResource
serializeWorkbookOutlineResource / deserializeWorkbookOutlineResource
```

`OutlineWorksheet`、panel 和 resource scope 的 extent 使用领域 `rowCount/columnCount`，缺省分别为 Office 上限；core 命令、Facade、布局和结构接线需使用一致的真实 extent，不能用 viewport 最小显示行列数改变公开范围。

Root 已协调 Data / ORGANIZATION 中“大纲设置”及完整分组菜单接线，包含 group/remove/collapse/expand/settings；固定来源 selector 仍是两子项，新增公开可见入口要按本地增强单列。Root 计划新增 `outlineHiddenRowSpans/outlineHiddenColumnSpans`（inclusive）并保留旧 arrays 兼容，Canvas、命中、滚动、冻结、打印等消费路径由其接线。此处仅记录责任和协议，运行态、最终布局及接线结果待统一验收。

## 明确差异与未核实项

- 固定发布类型和资源没有 summary 方向。本地继续支持 `summaryAbove/summaryLeft`，资源中采用 `officeStudio:{schemaVersion:1,...}`；不是冒称上游字段。Office 补充规格可按 [Outline.SummaryRow](https://learn.microsoft.com/en-us/office/vba/api/excel.outline.summaryrow)、[Outline.SummaryColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.outline.summarycolumn) 复查；本批未访问这些网页或运行 Office 验证。
- 本地 level 1 全折叠、maxDepth+1 全展开，最多九个按钮；固定 UI 最深按钮使用 toggle。该本地选择已写代码注释，不冒称相同内部交互；[Outline.ShowLevels](https://learn.microsoft.com/en-us/office/vba/api/excel.outline.showlevels) 的实机边界仍未核实。
- 轴复制为本地增强：只复制完全包含在源范围内的组，新的 ID 和嵌套扩展数据独立；部分源组不造出新组。固定本批类型/算法可证实 insert/delete/move，没有据此宣称其复制策略。
- 接受旧资源成员缺省 unitId/subUnitId；提供 owner 时验证其类型，最终公开快照绑定目标 workbook/sheet。固定 SDK 所有导入路径如何处理旧资源未核实。
- 原生 XLSX outlineLevel/hidden/collapsed/summaryBelow/summaryRight 的 parser/writer 映射、复杂 Office 样本和独立文件重开未完成；内部资源/JSON 草稿不等于原生 Office 保真验收。
- 当前嵌套/资源元数据逻辑在 TypeScript，复用单一模型候选。大数量组的 Worker/Rust 热路径、规模阈值和两种渲染后端实测尚缺，不能宣称性能提升、零拷贝或大型文件已达标。

## 验证草稿与实际执行

2026-11-26 初批新增 **68 条测试声明草稿**，由静态文本计数确认，全部未运行；下表保留初批记录，不含 2026-11-27 续批新增的 4 条：

| 草稿 | 条数 | 行为 |
| --- | --- | --- |
| [模型生命周期](../../packages/office-studio/tests/sheet-outline-model-lifecycle.contract.test.mjs) | 14 | 稳定相邻 ID、同范围层、八/九层与独立轴、错误、level、百万行压缩 span、summary/boundary、祖先隐藏、原子权限和 no-op |
| [结构与资源](../../packages/office-studio/tests/sheet-outline-structure-resources.contract.test.mjs) | 17 | 插删/完整移动/分裂拒绝、OR 去重、copy/ID/深拷贝、Office 上限、数量/inclusive、owner/legacy、空插件/错误/未知 sheet/原型式 ID、summary schema |
| [命令与菜单](../../packages/office-studio/tests/sheet-outline-command-menu.contract.test.mjs) | 12 | 4 core/8 UI 语义、prompt/整行列/多选、交叉/保护/clear capability、快捷键路由、owner 字段、原子本地命令和标签 |
| [owner 与弹窗控制器](../../packages/office-studio/tests/sheet-outline-panel-owner.contract.test.mjs) | 12 | revision/history/user/host、准备/提交同 ID、发布后重入、取消、权限复查、重开替换、不自动覆盖陈旧草稿、scope 释放；使用确定性 host fixture |
| [真实 Facade/Studio 集成](../../packages/office-studio/tests/sheet-outline-facade-lifecycle.contract.test.mjs) | 8 | 六 API、ignored invalid add、八层、保护、内容历史、native/public 重开、结构原子拒绝、owner 与 BeforeCommandExecute 重入；待执行时才创建 SSR harness |
| [真实工作台弹窗](../../tests/browser/sheet-outline-dialogs.spec.ts) | 5 | 快捷键 prompt/选轴、取消与焦点、陈旧/reread、两轴 summary 一次提交/Undo/Redo/资源、切 sheet 后取消；[独立配置](../../tests/browser/sheet-outline-dialogs.playwright.config.ts) 未启动 |

已有 11 条 `sheet-outline.contract.test.mjs` 和 1 条 `sheet-outline-workbench.spec.ts` 未在本子任务修改/运行；部分旧断言依赖 arrays 和旧 Ribbon，Root 保留兼容并在接线后复查。声明条数仅说明草稿覆盖设计，不证明功能或 UI 验收。

实际执行仅为静态读取、文本编辑、文件状态和声明数复查。没有执行测试、类型检查、构建、Rust/WASM/Cargo、dev/服务、浏览器、基准或 `git diff --check`。新增可替换控件均为 `antdv-next` Form/FormItem/Select/Button/Alert/Modal，样式引用统一主题 token；没有新增原生表单、文本/IME 桥或替代 Canvas 内核。主题切换、弹层/焦点、Canvas 输入法、两种渲染后端、Mac 快捷键、真实文件往返和性能仍为 **未验收**。最终产品能力由 Root 汇总共享接线和统一验证，不以这份证据标记全部 SHEET-27 完成。

## 2026-11-27 Organization 续核

固定 Pro 缓存 [outline UI ES](../../.research.local/univer-1.6.0/packages/sheets-outline-ui/lib/es/index.js) 的 `On/kn` 要求 Windows Alt+Shift+Right/Left，Mac Ctrl/Cmd+Shift+K/J。[use-sheet-outline-workbench.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-outline-workbench.ts) 原 Mac 分支错误要求不按 Shift，本轮已改为要求 Shift；现有 IME guard 和 Windows 分支保留。[命令菜单契约草稿](../../packages/office-studio/tests/sheet-outline-command-menu.contract.test.mjs) 的既有快捷键用例现直接调用真实 helper，覆盖 Windows 两键、Mac Ctrl/Meta 两种修饰键、Shift 大写 K/J，以及缺 Shift、非 Mac 和 composition 拒绝。用例数量仍为 12，本轮未运行。

同一固定 UI 的 `H/Vt` 和 [outline core ES](../../.research.local/univer-1.6.0/packages/sheets-outline/lib/es/index.js) 的 `sheet.command.clear-dimension-outlines` 均只取消完全包含在选区中的组。部分相交外层组保留是已核实契约；没有可删组时上游返回 false 并反馈错误。本地 [sheet-outline.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline.ts) 的 inclusive containment 与既有 partial-clear no-op 草稿符合该删除边界，本轮未修改。

| 本地分类 | 可复查数量与实际接线 | 当前行为与剩余验收 |
| --- | --- | --- |
| Organization 大纲根 | 固定组合 selector 1、本地增强分组显示 1、本地增强大纲设置 1 | 固定 selector 仍为 order 3 / grid row 2、column 2；新增增强项不能代替缺失的表格入口 |
| 固定 selector 子项 | 组合、取消组合共 2；行列头上下文各 1 根、各 2 子项；快捷键 2 | 按真实候选/保护判断状态，不把入口数量当成功能完成 |
| 本地分组显示子项 | 12 个直接项：10 个动作、2 个汇总位置子菜单；另有 4 个嵌套位置动作 | [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 已复用行列上下文菜单的 4 个分组/取消子项；6 个折叠/展开动作和 4 个汇总位置动作通过 getSheetOutlineUpdate 判断真实候选、no-op 和保护范围。根在缺少上下文或工作表无编辑权限时禁用；运行交互未验收 |
| 本地设置入口 | open-outline-settings 保留 assertCanEditWorksheet；[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的 outlineSettingsAvailable 已同时检查 Studio、workbook ID、大纲上下文和 canEditWorksheetProtection | 独立“大纲设置”根已接入工作表编辑权限；弹层、主题与焦点恢复未验收 |

已读取 [sheet-outline-commands.ts](../../packages/office-studio/src/api/sheet-outline-commands.ts)：真实编辑进入统一 CHANGE_SHEET_OUTLINE_COMMAND → commitWorkbook，具 owner/revision/history/user 与活动宿主检查，并由 getSheetOutlineUpdate 检查动作的保护范围。未改变上述接口、内容历史或资源协议。顶部表格根的缺口、布局冲突和创建方案另记于 [SHEET-18 创建入口续核](SHEET-18-table-selector-entry-review.md)。

菜单状态按实际变更组的完整行/列范围检查保护，不能只检查当前单元格选区。折叠选择包含选区的最小组；可编辑内层组不受外层其他区域的保护阻断，汇总位置变更仍检查该轴全部组。多选禁用依赖单一选区的动作，汇总位置继续按整轴能力判断。

| 续批既有草稿文件 | 新增声明 | 覆盖设计与实际执行 |
| --- | --- | --- |
| [Ribbon 契约](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs) | 2 | 无上下文/无组、部分包含取消、真实行命令参数、折叠/展开 no-op、多选与整轴汇总；完整外层保护范围、另一轴可编辑、可编辑内层组、工作表保护。全部未运行 |
| [工作台浏览器](../../tests/browser/sheet-outline-dialogs.spec.ts) | 2 | 通过原生 WorkbookModel 快照重开真实 fixture；行组受保护时禁用、列组允许折叠，仅开菜单不改内容或历史，列折叠仅增加一次 revision/history；工作表保护禁用独立设置和分组显示根。文件当前共 7 条，全部未运行 |

浏览器草稿改用真实独立根或通用“更多”对话框查找设置入口，并读取 aria-expanded 避免重复打开时关闭弹层；保护 fixture 使用领域类型要求的 onlyMe / othersCanView 及完整必填字段。快捷键契约仍为 12 条，没有新增测试文件。

本轮完成源码、字段、变更范围和声明数静态复核。此前命令菜单草稿的 node --check 语法检查通过，仅为语法检查；本次续批没有运行测试、类型检查、构建、Rust/WASM、dev/服务、浏览器或基准。没有新增 UI 或控件迁移，沿用现有 antdv-next Ribbon 和弹层。快捷键、菜单状态、创建/取消、撤销重做、文件重开、主题、焦点及输入法的运行验收继续后置，整体保持 **未验收**。
