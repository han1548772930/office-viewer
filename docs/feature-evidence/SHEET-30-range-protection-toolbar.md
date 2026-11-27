# SHEET-30：保护范围顶部新建入口

记录日期：2026-11-27（Asia/Shanghai）。固定参考 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；只读源码位于 `E:\rust_test\univer`。本批修正既有顶部入口的新建意图；SHEET-30 整体仍 **未验收**。按当前阶段约定，测试、类型检查、构建、dev、浏览器、Rust/WASM、Office 样本和基准均未运行。

<a id="range-protection-toolbar-create-intent-review"></a>

## 固定入口与调用链

| 来源 | 已读取的契约 |
| --- | --- |
| [sheets-ui/menu/schema.ts:372](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/schema.ts#L372) | START / OTHERS 的保护根 order 0，grid r1c1 跨两行并显示标签，factory 为 sheetPermissionToolbarMenuFactory。 |
| [permission.menu.ts:49](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/permission.menu.ts#L49)、[permission-menu-util.ts:230](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/permission-menu-util.ts#L230) | 根命令为 sheet.command.add-range-protection-from-toolbar。单元格编辑器可见、公式栏焦点、绘图、缺表/选区、WorkbookCreateProtectPermission、工作表保护和保护范围相交共同影响禁用；显隐来自 Sheets 宿主。 |
| [range-protection.command.ts:36](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/range-protection.command.ts#L36)、[sheet-permission-open-panel.operation.ts:31](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/operations/sheet-permission-open-panel.operation.ts#L31) | toolbar 通过 sheet-permission.operation.openPanel 传 showDetail=true，未传旧 rule；弹层使用权限详情视图。查看所有规则命令传 showDetail=false。 |
| [SheetPermissionPanel.tsx:56](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/views/permission/SheetPermissionPanel.tsx#L56)、[PermissionDetailPanel.tsx:47](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/views/permission/panel-detail/PermissionDetailPanel.tsx#L47)、[util.ts:97](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/views/permission/util.ts#L97) | 无旧 rule 时生成当前 selections 的新规则，描述/ID/permissionId 为空，缺省 onlyMe / othersCanView；从 sheet bar 打开时另用整表范围。 |
| [PermissionDetailFooterPart.tsx:201](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/views/permission/panel-detail/PermissionDetailFooterPart.tsx#L201)、[add-range-protection.command.ts:27](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/add-range-protection.command.ts#L27) | 固定详情先由 authzIoService 创建权限对象，再执行 AddRangeProtectionCommand → AddRangeProtectionMutation → RangeProtectionRuleModel。所读新增命令没有直接 pushUndoRedo，完整上游保护新增历史/服务持久化尚未核实，不能假定与本地内容历史一致。 |

本批入口计数：既有顶部根 1，本批新增根 0，顶部子项 0；本地根仍为开始 / 工具第 1 项，classic 顺序和 grid r1c1 跨两行位置不变。工具组另有快捷键面板、搜索功能、清除内容与范围主题四根，属于各自功能。本地右键新增/查看两项沿原入口保留，本批没有增加右键或上下文入口。此计数只说明入口归属。

## 本地小修与模型保存

原 [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 的根只发 sheet.operation.open-range-protection。[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的默认选区分支置 protectionStartNew=false；[SheetRangeProtectionPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetRangeProtectionPanel.vue) 打开时 loadRules 会选第一条旧规则。已有规则属于另一用户时，顶部“保护范围”会进入旧规则只读表单，即使当前选区尚无保护。

根现在传既有 `{ mode: 'add' }` 意图；默认选区分支设置当前范围并尊重该意图，面板按既有 startNew 调用 beginNewRule，显示当前选区的空白新建草稿。原带 sheetId 的右键新增/查看分支保持原行为，无 mode 的默认管理入口仍选择旧规则。本批生产修改仅涉及 ExcelWorkbench.vue 与 sheet-ribbon.ts，增量分别为 +5/-1 和 +1/-1，合计 +6/-2。

新建保存复用既有面板 apply → [FWorksheet.setRangeProtection](../../packages/office-studio/src/api/facade.ts) → [createSheetRangeProtectionRule / setSheetRangeProtectionRule](../../packages/office-studio/src/features/sheet-protection/range-protection.ts) → [sheet.command.set-range-protection](../../packages/office-studio/src/api/office-studio.ts)。命令核对当前用户及旧规则所有者，提交 WorkbookModel.rangeProtectionRules；注册事务经 commitWorkbook → prepareWorkbookChange/publish 进入本地内容历史。面板读取保存后的规则，旧规则和单元格内容保留。真实所有者校验、用户权限、历史和模型算法没有另建实现。

模型与视图的保护规则字段由 [excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) 携带。浏览器草稿保存真实当前模型的内部 JSON，并经原 FOffice.createWorkbook 重开；这只覆盖内部模型读回，不代表 XLSX 保护或跨用户服务持久化。入口属于 Vue 主线程 UI 编排，规则构造和现有事务保持同步；公开 API、模型字段、事件、内容历史与 Worker/WASM 协议没有新增接口，沿用既有 operation 的 mode 字段。

## 草稿与未验收项

[原 Ribbon 合同草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs) 新增 1 声明，检查真实新建 invocation、原位置及缺表行为；保留主线新追加的 ARABIC/ROMAN/BASE/DECIMAL 名称派发。[原范围编辑浏览器草稿](../../packages/office-studio/tests/sheet-range-editor.browser.test.mjs) 新增 1 声明，内部覆盖 classic/grid 两个样本：存在另一用户的旧 A1:C1 规则时，顶部入口直接为 D3 新建；打开不写历史，保存保留旧规则及内容，一次生产历史、undo/redo、内部 JSON 重开及真实右键查看旧规则均有断言。明转暗发生在面板打开期间，但未据此声称主题已验收。

本批共五文件：ExcelWorkbench.vue、sheet-ribbon.ts、上述两份原草稿和本页。rg 当前计数为 Ribbon 30 声明/30 计划；范围编辑浏览器 15 声明/16 计划（其中 1 个外层 for 注册 classic/grid 两例）。本批两个布局样本在同一声明内，不另计独立声明；全部运行 0。只完成源码/diff 阅读和 rg 计数，语法/引用/AST 与全局清单由主线统一收尾。

现有保护表单使用 antdv-next Modal、Form/FormItem、Input、Select 和 Button；本批没有新增控件或原生桥。统一 ConfigProvider 已提供库控件主题，现有面板自定义选中、边框与反馈仍含固定颜色，完整 token 适配及关闭后焦点恢复尚未核实。原生 IME、明暗主题/焦点、Canvas2D/WebGPU、Office 文件、真实服务和性能保持未验收。

SHEET-30 尚有明确缺口：顶部根的专属创建权限与任意交叉保护判定尚未统一采用固定 factory 的条件；本地顶部当前将多选区取一个外包矩形，未保留上游逐块 selections；没有 OfficeStudio 的独立宿主尚未接入此面板；上游详情可编辑范围、服务权限对象和完整保护历史/Office 保真也未闭环。前述批次修正误进入旧规则的问题，后续输入状态见下一节；完整工作簿/工作表/范围保护及 Pro 目标继续保留。

<a id="range-protection-input-session-review"></a>

## 2026-11-27 保护新建输入与现有编辑能力续修

再读固定 permission-menu-util.ts:230–299 与 permission.menu.ts:49–59：保护根在可见单元格编辑器、FX 焦点或绘图焦点时禁用，并另查 WorkbookCreateProtectPermission、工作表保护与范围相交。固定 hidden 仍只依 Sheets 宿主，不增加绘图隐藏。本地新建意图批没有这些输入约束，迟到 execute 会先 finishEdit 提交草稿再打开新建表单。

ExcelWorkbench 现向 sheet-ribbon 传可选内部 rangeProtectionEditable=canQuickFormatSelection，根结合缺表禁用。现有 canQuickFormatSelection 包含活动宿主、有效选区、工作表/范围编辑权限、单元格及公式编辑、组字和绘图选区；省略可选状态的 Ribbon 调用方保留既有行为。在 executeRibbon 入口一次识别 `{ mode: 'add' }` 新建请求，先用同一能力拦截，再将该布尔值交原带 sheetId / 默认选区两分支设置 startNew。普通输入、组字或不可编辑选区的迟到新建不提交内容，不产生保护表单、模型变化或历史。原无 mode 的查看/管理仍可查看另一用户的旧规则，本批没有把查看误改为编辑权限入口。

本批复用现有用户权限与同步规则事务；没有创建 WorkbookCreateProtectPermission 的本地替代物、扩展权限模型、改变 API 或 Worker/WASM 协议。canQuickFormatSelection 的一般编辑权限不等于上游专属创建权限；允许编辑的既有保护范围相交和已保护工作表等状态仍有差异，继续列为缺口。入口根增量 0，位置、标签及新建 mode 均沿上一批。

本批五文件：ExcelWorkbench.vue、sheet-ribbon.ts、两份原 Ribbon / 范围编辑草稿和本页。生产增量分别 +6/-3、+2/-1，合计 +8/-4。合同新增 1 声明，覆盖可选能力、禁用后的真实 invocation、缺表、位置及其他字体能力；浏览器新增 1 声明，包含 cell/formula × classic/grid 四个输入样本及一个真实保护选区样本。原生 disabled 点击及两种新建参数的迟到 execute 在普通/合成组字/明转暗时保持草稿、模型、revision 和历史，Esc 后可正常打开当前 D3 新建；另一用户保护的 B1 禁止新建，原管理入口仍打开 A1:C1 只读表单。上一批真实新建、保存、历史及内部重开草稿保留。当前 Ribbon 31 声明/31 计划、范围编辑浏览器 16 声明/17 计划，全部运行 0。

可替换控件继续为现有 antdv-next Button/Input/TextArea/Modal/Form/Select，没有新增表单或原生输入桥。只完成源码回读和 rg 计数；静态语法/引用收尾由主线统一，未运行测试、类型检查、构建、dev、浏览器、Rust/WASM、Office 或基准。合成 composition 不代表 Windows 原生 IME；完整主题 token、焦点恢复、两渲染后端、Office/服务保真与专属权限仍 **未验收**。
