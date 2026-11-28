# SHEET-13：复选框、输入提示、无效反馈与公开下拉操作

记录日期：2026-11-27（客户端 Asia/Shanghai）；文件名沿续接批次。固定版本：Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。生产实现和测试草稿已落盘，状态为未验收；测试、类型、构建、Rust/WASM、浏览器、Office 往返和性能按 AGENTS.md 统一后置，运行次数为 0。

<a id="checkbox-feedback-dropdown-2026-11-27"></a>

## 入口与来源

| 分类 | 固定版来源与可观察契约 | 本地落点 |
| --- | --- | --- |
| 常驻 Insert / CELL | [dv.menu.ts](../../../univer/packages/sheets-data-validation-ui/src/menu/dv.menu.ts)：checkbox order 0，dropdown order 1，date order 2；三个根无 gridLayout | 沿既有三根、split 主动作及分类清除。本批没有新增 Ribbon 根 |
| 常驻 Data / RULES | [schema.ts](../../../univer/packages/sheets-data-validation-ui/src/menu/schema.ts)：管理根，管理/添加两个子项；grid r1c1 | 沿既有管理和添加命令、规则面板 owner |
| 单元格复选框上下文 | [checkbox-widget.ts](../../../univer/packages/sheets-data-validation-ui/src/views/widgets/checkbox-widget.ts)、[checkbox-validator.ts](../../../univer/packages/sheets-data-validation/src/validators/checkbox-validator.ts)：字体 × 1.6、6px 水平 margin、合并主格、空值/合法值、1/0 默认及完整 TRUE/FALSE 别名 | 一个可见合并主格对应一个 Ant Checkbox；typed 切换命令直接写实际 scalar |
| 无效内容上下文 | [dv-render.controller.ts](../../../univer/packages/sheets-data-validation-ui/src/controllers/dv-render.controller.ts)：右上角 6px 标记；[dv-alert.controller.ts](../../../univer/packages/sheets-data-validation-ui/src/controllers/dv-alert.controller.ts)：100ms hover 与错误内容 | Canvas 标记与 Ant Popover/Alert；完整 View 掩码；滚动/缩放/切换清理 hover |
| 显式下拉 UI operations | [data-validation.operation.ts](../../../univer/packages/sheets-data-validation-ui/src/commands/operations/data-validation.operation.ts)、[dropdown-manager.service.ts](../../../univer/packages/sheets-data-validation-ui/src/services/dropdown-manager.service.ts) | 同名 show/hide，目标可与当前选区不同；显式 show 支持隐藏箭头，仍检查完整目标 View/Edit |
| 列表弹层 | [ListDropDown.tsx](../../../univer/packages/sheets-ui/src/views/dropdown/ListDropDown.tsx)、[list-validator.ts](../../../univer/packages/sheets-data-validation/src/validators/list-validator.ts) | 搜索、单/多选、清除/关闭、编辑规则、动态候选更新，共用现有 popup owner |

分类计数：数据验证共有 4 个常驻根，管理根含 2 个普通菜单项；本批新增常驻、插件扩展、溢出根均为 0。复选框上下文 1 类；日期/列表触发上下文 2 类，当前目标互斥；prompt 与 invalid 反馈各 1 类。显式 show/hide 是 2 个公开 UI operation，不计为 Ribbon 根。每个可见复选框、N 个列表候选和弹层内部按钮不充当功能完成数量。classic/grid 共用命令和能力条件，窄屏沿既有溢出归属。

prompt 的固定上游依据目前只核实了模型/侧栏字段，没有找到与本地相同的选中提示 UI。已实际读取 Microsoft [Apply data validation to cells](https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249)（2026-11-27）：`Show input message when cell is selected`，配置 Title/Input message，选中有输入要求的单元格时显示。这里用于本地可观察行为依据；不声称固定 Univer prompt UI 已核实。关闭提示、Ant 浮层外观和与无效 hover 互斥是本地交互选择，尚待浏览器验收。

## 真实链路

[validation-checkbox.ts](../../packages/office-studio/src/features/sheet-validation/validation-checkbox.ts) 集中解析 checked/unchecked：缺省 1/0；只有完整 TRUE/FALSE 词转换比较别名，自定义大小写、空白及类型保持原 scalar。规则公式在保存 anchor 或首范围求值，使用共享公式 reader。空 scalar 可显示 unchecked，invalid text、公式错误、富文本和图片保留原 Canvas 内容，不伪装成 checkbox。默认候选查询现已为 `1` / `0`，原 TRUE/FALSE 查询默认差异已修正。

[validation-checkbox-input.ts](../../packages/office-studio/src/features/sheet-validation/validation-checkbox-input.ts) 的 `sheet.command.commit-checkbox-input` 复用 [validation-widget-input.ts](../../packages/office-studio/src/features/sheet-validation/validation-widget-input.ts) 与 [sheet-validation-input-commands.ts](../../packages/office-studio/src/api/sheet-validation-input-commands.ts)。pointer、Checkbox Space/Enter、Canvas Space 共用单主格 typed 内容命令：规则、用户、工作簿/表、revision/history、完整合并范围 View/Edit 与输入会话均重查；成功一次内容历史和 SheetValueChanged。清除被替换公式/富文本缓存，保留其它 style/metadata；取消或旧回调不修改值。

[use-sheet-validation-checkbox.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-validation-checkbox.ts) 捕获可见主格及规则基线，阻止旧 pointer/Space、观察者重入和失活后的写入/回焦；输入、composition、其它编辑器弹层和保护禁用。回焦仅在仍属同一工作台及内容版本时执行。

[validation-checkbox-layout.ts](../../packages/office-studio/src/rendering/spreadsheet/validation-checkbox-layout.ts) 按真实可见稀疏轴与冻结 panes 汇集合并主格 fragments，每个主格只求值一次。[SheetCheckboxValidationOverlay.vue](../../packages/office-studio/src/components/spreadsheet/SheetCheckboxValidationOverlay.vue) 仍只创建一个 Ant Checkbox，CSS clip-path 合并实际 fragments，修复只取最大 pane 会丢掉跨冻结线方框大部分的问题。checkbox 内容尺寸、align/vertical、zoom 与 Canvas geometry 共用。浏览器 clip-path、焦点及两种后端尚未验证。

[validation-cell-feedback.ts](../../packages/office-studio/src/features/sheet-validation/validation-cell-feedback.ts) 为请求的可见格和选中主格派生 prompt/issues；合并格去重，整批候选缓存共用验证引擎。DTO 仅含目标位置和规则 ID/提示类型/文案，不向 UI 带值、完整规则或源范围。实际公式通过共享求值链读取，旧 cached scalar 不冒充当前结果；View-denied 目标/完整 merge 不产生反馈。

[SheetValidationFeedbackPopup.vue](../../packages/office-studio/src/components/spreadsheet/SheetValidationFeedbackPopup.vue) 使用 Ant Popover、Alert、Button。选中满足 showInputMessage 的规则显示 prompt；dismiss 对该选中 owner 生效，换格恢复，composition 期间禁止关闭。无效 hover 延迟 100ms，在 Canvas [validation-cell-markers.ts](../../packages/office-studio/src/rendering/spreadsheet/validation-cell-markers.ts) 绘制按 zoom 缩放的 6px 标记；只显示当前有 issues 的可见位置。无效内容仍保留原值，selected gridcell 标注 aria-invalid。error 关闭时不显示自定义错误文案，仍提供一般无效原因。所有 UI 文案沿统一 [validation-input.ts](../../packages/office-studio/src/i18n/sheets/locales/validation-input.ts)，19 locales × 37 keys 与共享 i18n 合并。

[sheet-validation-dropdown-operations.ts](../../packages/office-studio/src/api/sheet-validation-dropdown-operations.ts) 在 Studio 注册/释放：

- `sheet.operation.show-data-validation-dropdown` 接收 unitId/subUnitId/row/column，解析完整 merge 和日期/列表规则；有效宿主可 reveal 非当前目标而不修改选区。与 [use-sheet-validation-cell.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-validation-cell.ts) 共用 owner，同目标重复 show 保留搜索及日期草稿。
- `sheet.operation.hide-data-validation-dropdown` 支持可选工作簿/表过滤；同一活动 popup 的 composition 或 blocked scope 不能被另一 scope 抢占，已关闭目标 hide 幂等，不产生内容历史。
- reveal false/throw、准备期间 revision/用户/历史变化和旧 owner 开始 composition 均保留原有效草稿；候选完成和 reveal 成功后才发布新 owner。命令取消在 UI entry 返回 false，API 仍保留标准 canceled 异常契约。

动态列表源已移除 validation 自有区域正则 parser，使用共享 `evaluateFormulaReference`：名称、Unicode/local scope/qualified names、IF/CHOOSE、INDIRECT A1/R1C1、OFFSET 与嵌套 COUNTA、INDEX 以及 whole-axis 区域。scalar/reference 同时传共享 canReadCell/canReadRange，source reader 扩展完整合并范围，检查嵌套依赖；未选择分支不泄露其隐藏内容。候选按真实稀疏 cells 行优先、去重、过滤空值/错误；修改源/selector 后更新同一 popup。前段创建仍沿固定版四函数过滤；后续已允许纯 INDEX authoring，并接源数字格式/地域/日期系统与颜色，差异和未运行闭环见 [INDEX 与 numfmt 专项](SHEET-13-list-source-index-numfmt-2026-11-28.md)。

## 草稿与实际检查

新增/扩展的草稿全部未运行：

| 草稿 | 声明 / 计划场景 | 行为断言 |
| --- | --- | --- |
| [sheet-checkbox-input.contract.test.mjs](../../packages/office-studio/tests/sheet-checkbox-input.contract.test.mjs) | 14 / 25 | 默认/自定义/boolean/公式、invalid/blank、single-cell history、merge/源权限、取消/撤权/重开 |
| [sheet-checkbox-host.contract.test.mjs](../../packages/office-studio/tests/sheet-checkbox-host.contract.test.mjs) | 8 / 16 | captured owner、旧回调、composition/input、blocked/失活、standalone、观察者重入与回焦 |
| [sheet-validation-feedback.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-feedback.contract.test.mjs) | 11 / 21 | 九类无效、any prompt、required blank、merge 去重、DTO 保密、实际公式/完整源 View、batch cache |
| [sheet-validation-dropdown.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-dropdown.contract.test.mjs) | 15 / 17 | 显式隐藏 trigger、非选中目标、幂等、过滤、reveal false/throw、取消/revision/composition、dispose/replacement |
| [sheet-validation-source.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-source.contract.test.mjs) | 14 / 33 | 名称、动态 reference/selector/anchor、隐藏分支/空范围、嵌套依赖、非法/循环源，以及 INDEX 创建、numfmt/region/date1904/颜色 |
| [spreadsheet-checkbox-feedback.contract.test.mjs](../../packages/office-studio/tests/rendering/spreadsheet-checkbox-feedback.contract.test.mjs) | 10 / 12 | 字体/align/zoom、四 pane 裁剪、hidden axes、6px marker/主题、invalid 原文/选区/View 掩码 |
| [sheet-checkbox-feedback.browser.test.mjs](../../packages/office-studio/tests/sheet-checkbox-feedback.browser.test.mjs) / [fixture](../../packages/office-studio/tests/fixtures/sheet-validation-cell.html) | 19 / 42 | classic/grid 的真实 Studio/Canvas/Ant checkbox、history、保护、旧 pointer/Space、输入合成 composition、四 pane、prompt/dismiss/hover delay、theme、公开动态 show/hide；后续新增联合规则/cells 事务、真实 Worker/WASM 及取消，以及 INDEX/numfmt 单选/多选文字 |

现有 [sheet-validation-cell-host.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-cell-host.contract.test.mjs) 增加动态 IF 分支更新、实际提交与两条历史 undo/redo。前段静态时点的日期/列表/规则/feedback 草稿合计 **15 文件、181 声明、306 计划场景，运行 0**；联合规则/cells 与 Worker 首次快照为 **21 文件、256 声明、435 计划场景，运行 0**，详见 [联合事务专项](SHEET-13-checkbox-rule-cells-worker-2026-11-28.md)。INDEX/numfmt 与 bulk workbench 续段快照为 **22 文件、271 声明、467 计划场景，运行 0**，详见 [后续专项](SHEET-13-list-source-index-numfmt-2026-11-28.md)。

前段实际只读静态检查为 **47 文件、47 段脚本语法、7 个 Vue 模板、7 段样式、450 个本地引用和 16 次 Node 语法解析**；前段 19 locales × 37 keys、7 文档/159 本地链接/43 固定来源及限定 diff 检查通过。联合事务、Worker 与批量 UI 首次静态快照为 **67 文件、67 scripts、7 templates、7 styles、577 imports、22 Node parses**；bulk 文案附加 5 keys/locale。INDEX/numfmt 与 workbench 续段快照为 **68 文件、68 scripts、7 templates、7 styles、592 imports、23 Node parses**，后续来源/链接/计数见 INDEX/numfmt 专项。Node --check 和 SFC 编译仅解析语法/模板；未启动测试 runner 或浏览器，不能从草稿数量推断功能验收。

## 仍需实施及验收

- Checkbox 新增/扩展空格填值和自定义操作数迁移已进入联合规则/cells 事务；有界 Worker/Rust、异步 owner 和取消/内存已落盘，仍须真实大范围、性能和同步 API 策略验收。
- INDEX authoring 和候选数字格式已在 [后续专项](SHEET-13-list-source-index-numfmt-2026-11-28.md) 落盘且未验收；完整数组/spill 和大列表 Worker 计算仍待推进。
- 标签自动行高/列宽、旋转/缩进等完整布局。
- 真实 Windows IME、焦点/键盘/clip-path、明暗主题、权限/协作、结构变换、两渲染后端及独立 Office 样本往返；内部快照 reopen 草稿不替代 Office 验收。
- 首屏/滚动/编辑/zoom、传输次数与字节、内存和取消的同数据集大样本实测。前段未新增 Worker/WASM buffer/lease 协议，后续联合规则/cells 事务已新增并复用共享 bridge，详见 [联合事务专项](SHEET-13-checkbox-rule-cells-worker-2026-11-28.md)；尚未实测，不声称优化或达标。

SHEET-13 及 Sheets/Docs/Pro 全量范围均保持未验收；这一批实现不能代替完整交付。
