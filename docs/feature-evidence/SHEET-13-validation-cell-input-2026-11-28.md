# SHEET-13：日期、列表控件与规则面板闭环

日期：2026-11-28。状态：已写入生产实现、工作台接线和草稿，未验收。所有运行验证按当前 AGENTS.md 统一后置。OfficeStudio、根导出及 locale 由主任务协调注册；静态接线不代表运行或文件样本通过。

<a id="validation-cell-input-2026-11-28"></a>

## 固定入口与分类

固定来源为 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`：[schema.ts](../../../univer/packages/sheets-data-validation-ui/src/menu/schema.ts)、[dv.menu.ts](../../../univer/packages/sheets-data-validation-ui/src/menu/dv.menu.ts)、[data-validation.operation.ts](../../../univer/packages/sheets-data-validation-ui/src/commands/operations/data-validation.operation.ts)。

| 分类 | 固定入口/顺序 | 本地落点与范围 |
| --- | --- | --- |
| 常驻 Data/RULES | `sheet.menu.data-validation`，order 0，grid r1c1；管理子项 order 0，添加子项 order 1 | 管理 `sheet.operation.open-data-validation`；添加沿默认 decimal/equal/100 创建再按新规则 ID 打开 |
| 常驻 Insert/CELL | 复选框 order 0、下拉 order 1、日期 order 2；三个节点没有 gridLayout | [快捷入口证据](SHEET-13-quick-dropdown-date.md) 及 [复选框证据](SHEET-13-checkbox-entry-review.md)，本批没有新增常驻根 |
| 选中单元格上下文 | DATE/时间格式和 LIST/LIST_MULTIPLE 由 dropdown manager 与 widget 定位，不属于 Ribbon 常驻根 | 主格上的一个互斥日期/列表触发 Button；可见合并格、冻结 pane、滚动及缩放定位；Alt+Down 共用 owner |
| 弹层内部 | DateDropdown 的 dirty/确认；ListDropdown 的搜索、选中、编辑入口 | 日期按模式显示 1–3 个 Ant 输入并明确确认/取消；列表显示搜索、N 个候选 Checkbox、清除/关闭及有权限时的规则编辑 |
| 规则侧栏 | 固定 [DataValidationDetail.tsx](../../../univer/packages/sheets-data-validation-ui/src/views/components/DataValidationDetail.tsx) 和 [ListFormulaInput.tsx](../../../univer/packages/sheets-data-validation-ui/src/views/components/formula-input/ListFormulaInput.tsx) | 312px 侧栏；规则列表/详情、应用区域、类型、操作数、showTime、列表项/颜色/顺序、错误/输入提示、删除和清除全部 |

可复查计数：数据验证功能共有上述 4 个常驻根，管理根含 2 个普通菜单项；本批新增常驻/插件扩展/溢出根均为 0。日期与列表上下文各有一种互斥触发器，实际只绑定当前可见目标，不把全表单元格数或 N 个候选当作新增根。classic/grid 复用同一入口与命令；窄屏溢出归属沿共享 Ribbon。固定管理命令是 `data-validation.operation.open-validation-panel`，添加是 `data-validation.command.addRuleAndOpen`；固定 show/hide dropdown operations 为 `sheet.operation.show-data-validation-dropdown` / `sheet.operation.hide-data-validation-dropdown`，2026-11-27 续接已注册同名公开 UI operations，详见 [本批操作证据](SHEET-13-checkbox-feedback-dropdown-2026-11-28.md)。

固定菜单仅在 Sheets 显示，其禁用点为 WorkbookEditable、WorksheetSetCellStyle、WorksheetEdit、RangeProtectionEdit。本地另检查工作台、用户、完整 View/Edit 权限及输入/IME/其它弹层状态。列表与日期 Button 出现要求当前规则允许 showDropDown 且 renderMode 非 text；UI 输入全部由 antdv-next 管理。

## 命令、模型和历史

[validation-widget-input.ts](../../packages/office-studio/src/features/sheet-validation/validation-widget-input.ts) 复用既有输入会话 reader、合并主格、样式继承和数据验证。新内容命令见 [sheet-validation-input-commands.ts](../../packages/office-studio/src/api/sheet-validation-input-commands.ts)：

- `sheet.command.commit-date-input` 接收 unchanged/date/time/date-time/serial，直接写 numeric serial；日期采用既有 Gregorian 算术，Dayjs 仅适配控件。
- `sheet.command.commit-validation-list-input` 接收完整标签数组，按完整候选顺序、忽略大小写匹配并去重。标签强制文本写入，`1` 和 `=1+1` 不变成数字或公式。多选即时提交后保留弹层并重建 canonical session；单选成功关闭。

每次内容写入清除被替换的公式、富文本和值缓存，保留不相关 style/metadata，一次 commitWorkbook、一次所属工作簿历史和一次 SheetValueChanged。规划、mutator 与 prepareView 重查 workbook/sheet 身份、revision/history、用户、完整 JSON 基线、规则和目标 View/Edit。API 可编辑非活动工作簿；UI host 必须仍为当前工作台。STOP 或取消不提交；warning/info 以 issue-bearing accepted 结果写入。

日期 serial 0/60 显示真实兼容日期标签、日历留空；近似锚点只用于选择真实日期。时间修改保留伪日期，纯日期修改直接保留原 fractional day；未修改确认保留原精度、数值字符串类型和格式。已修改时同类别格式保留，否则分别为 `yyyy-MM-dd`、`yyyy-MM-dd hh:mm:ss`、`HH:mm:ss` 和 `[h]:mm:ss`。日期详情见 [实施记录](SHEET-13-date-picker-entry-plan.md)。本地独立 time 规则允许有限非负 serial，有 operator 时比较实际 serial；固定版没有独立 time-validator，此扩展的 Office 样本效果未核实。

本地列表表示仍为 CSV：[validation-list-input.ts](../../packages/office-studio/src/features/sheet-validation/validation-list-input.ts) 与 [validation-engine.ts](../../packages/office-studio/src/features/sheet-validation/validation-engine.ts) 共用 parser/serializer。固定版 [utils.ts](../../../univer/packages/sheets/src/basics/utils.ts) 使用 JSON 字符串数组，这是明确协议差异，未加 JSON 互操作层。详细例子见 [CSV 证据](SHEET-13-list-multiple-csv.md)。

[validation-rule-command.ts](../../packages/office-studio/src/features/sheet-validation/validation-rule-command.ts) / [sheet-data-validation-commands.ts](../../packages/office-studio/src/api/sheet-data-validation-commands.ts) 保留 set/replace/remove/clear/clear-quick-data-validation 的现有五个 ID 与 boolean 返回。DTO 拷贝与纯计划独立可测；原规则完整范围及新范围一并检查 View/Edit；局部 clear 仅改指定可访问区域并保留外侧残余、公式锚点和 metadata。无实际修改不创建历史；已有未改动的导入公式可保留，编辑后的无效公式拒绝。

## owner、权限与 UI

[use-sheet-validation-cell.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-validation-cell.ts) 绑定目标主格、规则签名、用户、workbook/history 和输入会话。旧 submit/clear/close/after-close 回调不能作用于新 owner。外部无关内容可推进 model/revision；目标内容、格式、日期系统或规则冲突保留原日期草稿并 disabled。切格/表/本/用户/输入会话、历史清空、同 ID 模型替换或工作台失活结束旧 owner。standalone 只接受一次精确已发布快照引用回传；同字节外部替换仍失效。

[use-sheet-data-validation-workbench.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-data-validation-workbench.ts) 管理整个规则列表会话。UI 只取得完整覆盖范围可 View 的规则；可见但无完整 Edit 权限的规则不能编辑/删除。存在任何隐藏规则或不可编辑规则时禁用“清除全部”，不能只清可见子集后声称清除整个工作表。Canvas 取区仅改变选区，不重新绑定 owner；其它外部内容/规则/历史变更结束旧完整列表，避免以旧列表覆盖新模型。提交和观察者重入后只更新仍属当前 owner 的视图；正常关闭才恢复所属画布焦点。

[SheetDataValidationPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataValidationPanel.vue) 使用 Ant Form/FormItem、Input/Input.TextArea、Select、Checkbox、ColorPicker、Collapse、Button 和 Alert。内联 CSV 与逐项编辑同步；重排时颜色随标签移动；未改动的 formula1 原字节、formula2 空位/尾部位置和其它 metadata 保留。引用选项只读标签并允许颜色编辑，源预览受 View 权限约束。Select/ColorPicker 弹层放在所属 sidebar 中，表单和弹层的 composition 统一阻止保存/关闭；ColorPicker 使用可聚焦的 Ant Button 触发。

[SheetValidationCellControl.vue](../../packages/office-studio/src/components/spreadsheet/SheetValidationCellControl.vue) 使用 Ant Popover、Button、Form、Input、Checkbox、Tag 与 Alert；搜索、上下/Home/End 导航、清除、单/多选和规则编辑均路由 owner。[validation-cell-tags.ts](../../packages/office-studio/src/rendering/spreadsheet/validation-cell-tags.ts) 在 Canvas 绘制独立标签、候选颜色、换行、主题及字体装饰，替换 raw CSV 文本；[scene.ts](../../packages/office-studio/src/rendering/spreadsheet/scene.ts) 保留选区/fill handle，屏蔽不可 View 的文本/标签，并过滤条件格式计算用的不可 View 源格。此掩码只说明 Canvas 及新控件，不能推导全局 Facade、公式栏和统计已完成权限隔离。

[validation-input.ts](../../packages/office-studio/src/i18n/sheets/locales/validation-input.ts) 首批提供 19 locales × 23 个新文案，续接增加 Checkbox/prompt/invalid 文案后为 19 × 37。控件沿统一 ConfigProvider 和 theme token；日期 locale 沿既有 Dayjs/Ant 适配。原 Canvas native input 事件桥没有扩建第二个输入 DOM；替换控件由 Vue/Ant 拥有。

<a id="validation-cell-drafts-2026-11-28"></a>

## 草稿与静态检查

| 草稿 | 可审阅的真实断言 | 状态 |
| --- | --- | --- |
| [sheet-validation-input.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-input.contract.test.mjs) | serial 0/59/60/61、1904、隐藏时间/小数秒、数值纠正/duration、完整标签/CSV、STOP、merge/源权限、reader、一次历史/内部重开及准备阶段撤权 | 未运行 |
| [sheet-validation-cell-host.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-cell-host.contract.test.mjs) | authority、原输入取消/结束、多选重建会话、旧回调、冲突/失活、观察者重入、standalone 精确 echo | 未运行 |
| [spreadsheet-validation-cells.contract.test.mjs](../../packages/office-studio/tests/rendering/spreadsheet-validation-cells.contract.test.mjs) | CSV/颜色/主题/装饰、有界可见命令、合并冻结/缩放/隐藏、内容 View 掩码、选区/fill handle | 未运行 |
| [sheet-data-validation-commands.contract.test.mjs](../../packages/office-studio/tests/sheet-data-validation-commands.contract.test.mjs) | 新规则/替换/移除/局部分类清除、原/新范围 View/Edit、冻结 DTO/getter、陈旧 guard、未改导入公式、取消/撤权、隔离历史/内部重开 | 未运行 |
| [sheet-data-validation-host.contract.test.mjs](../../packages/office-studio/tests/sheet-data-validation-host.contract.test.mjs) | 可见列表、完整范围权限、清除全部条件、同 owner 连续编辑、取消、Canvas 取区、失活/外部替换/观察者重入及引用预览权限 | 未运行 |
| [sheet-validation-panel-state.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-panel-state.contract.test.mjs)、[sheet-validation-panel-ui.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-panel-ui.contract.test.mjs) | 区域/类型规则、CSV item 身份与颜色重排、保留原字节/空位、showTime/锚点和拒绝空列表；旧源字符串 UI 断言改为规则行为 | 未运行 |
| [sheet-validation-cell.browser.test.mjs](../../packages/office-studio/tests/sheet-validation-cell.browser.test.mjs) / [fixture](../../packages/office-studio/tests/fixtures/sheet-validation-cell.html) | classic/grid 的真实 Studio/Canvas/Ant 日期、时间、列表及面板输入；搜索/顺序/颜色、STOP/取消、权限、历史/重入、合成 composition、冻结缩放/主题/回焦 | 未运行 |

首批已执行静态检查：28 个生产/草稿/夹具文件、28 段脚本语法、5 个 Vue 模板、5 段样式、354 个本地引用和 9 次 Node 语法解析；8 份测试草稿包含 99 个声明、163 个计划场景，均未运行。19 locales × 23 keys 与既有 locale 集合一致；5 份专项文档的 98 个本地链接（其中固定上游来源 26 个）及真实文件名大小写核对通过。这些数字仅代表首批；续接为 47 文件和 15 草稿，详细静态结果与 181 声明/306 计划场景见 [新批记录](SHEET-13-checkbox-feedback-dropdown-2026-11-28.md)。固定来源的文件名是 `ListDropDown.tsx`，已纠正旧记录的大小写。限定 `git -c core.whitespace=cr-at-eol diff --check` 首批通过；此前强制关闭 autocrlf 的检查把 CRLF 错报为尾随空白，未据此改写换行。

上述结果只属于静态语法、SFC 模板/样式、本地引用、locale、来源/文档链接和 diff 检查。没有测试通过、类型通过、构建通过、浏览器通过或 Office 文件样本通过记录。浏览器草稿默认 Canvas2D；合成 composition 不代表 Windows IME 验收。

## 仍需推进

- 复选框 Canvas Ant 勾选、验证输入提示/无效标记、公开 show/hide 和共享动态/命名源已写入，均未验收，见 [续接记录](SHEET-13-checkbox-feedback-dropdown-2026-11-28.md)。checkbox 默认填值/自定义值迁移、完整普通键入/粘贴/填充、API、结构编辑及协作仍需推进。
- INDEX authoring 和候选数字格式已在 [后续专项](SHEET-13-list-source-index-numfmt-2026-11-28.md) 落盘，仍未验收；完整数组/spill 与 Worker 大源计算继续推进。共享 Facade 的规则/候选 View 掩码由主任务接线，不能由控件静态结果推断全局权限验收。
- [标签自动行高/列宽与旋转/缩进](SHEET-13-validation-tag-dimensions-2026-11-28.md)以及[内联缓存/Rust 查询/Ant 虚拟 UI](SHEET-13-validation-list-cache-query-2026-11-28.md)已在后续落盘，仍未验收；完整动态来源求值/依赖 epoch、两后端和真实大样本/取消/内存/传输基准继续推进。
- XLSX 及独立 Office 样本保真/往返，真实 Windows IME、焦点恢复、主题、权限与协作。原生快照重开草稿不能替代文件验收。

单格日期和小型命令计划仍在可测试的 TypeScript 服务域；Canvas 负责内容绘制，Ant 负责可替换控件。此首批没有新 Worker/WASM buffer/lease 协议；后续查询增加了协议 2 的必需导出，接口与预算见[查询专项](SHEET-13-validation-list-cache-query-2026-11-28.md)，首批静态结果不覆盖该追加。未新增商业 SDK 或项目外服务，未实测性能，不声称优化或达标。本批结果不能替代 Sheets/Docs/Pro 的完整范围。
