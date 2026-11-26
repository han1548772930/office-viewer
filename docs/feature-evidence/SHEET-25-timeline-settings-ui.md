# SHEET-25：正式 Timeline 设置与日期筛选 UI

记录日期：2026-11-26。状态：**UI、提交服务、夹具和验证草稿已落盘，未验收**。

本批接通独立 `FSheetTimeline` 的十项显示设置、真实来源日期边界、精确半开范围和八种固定条件。按“先实现、最后统一测试”的当前要求，仅静态读取与编辑，没有执行测试、类型检查、构建、Rust/WASM、dev、服务、浏览器、基准、计数脚本、链接检查或 `git diff --check`。计划场景不是通过数，也不代表完整 SHEET-25 已验收。

## 固定来源与入口清点

固定 OSS：`E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。1.6.0 发布包仅作为本地静态研究资料，没有安装或执行商业 SDK。

| 精确来源 | 可核对的契约与本地差异 |
| --- | --- |
| `.research.local/univer-1.6.0/slicer/lib/types/definition.d.ts` 第 94–114 行 | `ITimelineRange`、八种 `TimelineDateFilterOperator`、`ITimelineDateCondition`；本地保留真正的 Excel serial 范围或条件，不转换成成员快照 |
| 同文件第 115–142 行 | cache range/condition 互斥；`ITimelineViewConfig` 的 name、level、selectionLevel、caption、style、scrollPosition 和四个显示开关属于 view；本地模型仍使用字符串 grain，公共设置使用 0 年、1 季度、2 月、3 日 |
| `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/facade/f-timeline.d.ts` 第 89–98、153–168 行 | `getSettings/updateSettings`，防御性 retained 配置和 boolean acceptance；本地设置 Drawer 用同名正式 Facade，不增加 `setSettings` |
| 同 Facade 第 403–443 行 | `getDateRange/getDateCondition` 与 `setDateRange`；getter 不因当前来源失败丢失作者状态，setter `[startDate,endDate)`，undefined 清除，独立粒度不变 |
| 同 Facade 第 444–460 行 | async `getDateBounds` 返回 inclusive minDate/maxDate、Date1900/Date1904；本地独立来源查询，不从 viewport、成员列表或 selection end 推导 |
| 同 Facade 第 463–501 行及后续显示 setter | 显示粒度、选择粒度和 serial 导航互相独立；本地选择器使用 numeric levels，导航使用日期输入或精确 serial |
| `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/types.d.ts` 的 SheetTimelineSettings、SetDateRange、SetDateCondition、PatchView | 正式 Facade 只有条件 getter，条件写入属于内部 intent；本地日期 Modal 通过已有共享 Pivot 日期命令提交，没有新造公开 `setDateCondition` |
| `.research.local/univer-1.6.0/sheets-slicer-ui/lib/es/index.js` 的 getSettings、selectionLevel 控件与 Timeline dirty 比较 | caption??name、四个显示开关缺省为 true，显示与选择粒度分开；本地 Vue/Ant Form 保留 optional 作者值，不把显示缺省写成作者覆盖 |
| 同 ES 的 lr/ur 菜单及 `E:/rust_test/univer/packages/ui/src/views/components/ribbon/Ribbon.tsx`、RibbonGrid.tsx | 格式分类中“样式”先于“选项”，Timeline/Slicer 设置互斥，两个布局共用入口；本批复用既有上下文根 |

完整常驻、插件、上下文与溢出集合继续见[顶部来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)和[连接 UI 清点](SHEET-25-slicer-connections-ui.md)。本批入口逐项如下：

| 入口 / 标签 | 分组、顺序和位置 | 显隐 / 禁用条件 | 命令、模型与历史 |
| --- | --- | --- | --- |
| `sheet.menu.timeline-settings` / 时间线设置 | `slicer-format.options` order 1，根 order 1；grid r1c1；classic 第二组首项，与普通 Slicer 设置互斥 | 单个可访问 Timeline 被选中时显示；完整对象 footprint 无编辑权限时禁用 | `sheet.operation.open-slicer-settings-dialog` → Timeline Drawer → `getTimeline().getSettings()`；纯显示 `sheet.command.update-timeline-settings`，混合几何 `sheet.command.update-slicer-presentation`，同一文档内容历史 |
| `sheet.operation.open-timeline` / 时间线 | 既有“数据”页分析组，grid r2c2；沿用固定清点的 classic 位置 | 当前工作簿宿主可用；来源与字段由同一插入 Form 校验，Table 不适用 | 原始日期字段 → `resolveSheetSlicerSource` / 正式 Info → `FWorksheet.insertTimeline` → `sheet.command.insert-timeline`，一次插入历史 |
| Drawer / 日期筛选 | 设置 Form 的上下文动作区；Timeline 分支，报表连接之后 | 仅 Timeline；busy、IME 或无当前对象编辑权限时禁用；独立日期草稿不被未提交显示草稿覆盖 | range：正式 `setDateRange`；condition：`sheet.command.set-timeline-date-condition` 内部共享 intent；none：`sheet.command.set-timeline-date-range` 的 own undefined，同一模型、权限、保存与撤销链路 |
| Drawer / 还原显示默认值 | 显示设置组，四个开关之后 | Timeline 显示草稿可修改且非 busy/IME/stale 时可用 | 只改草稿；保存时一次 settings/presentation 命令，保留 name，清七个 optional，两个 grain 回到 Month |
| 日期 Modal / 重新读取、确定、取消 | Form 底部依次排列；填入来源整日范围位于 range 子组 | 重新读取/取消受 IME 保护；确定需有效 dirty、当前 owner、权限和可靠日期边界，none 不要求日期边界 | 读取与取消不写内容历史；确定提交实际条件/范围；原子提交失败保留作者状态 |

可复查分类：本批新增常驻 Ribbon 根 **0**、插件常驻根 **0**、独立 More 根 **0**；既有上下文根仍 **4** 个声明，Timeline 当前适用快速样式、Timeline 设置、连接 **3** 个。本批新增一个 Timeline 专用的嵌套“日期筛选”入口，复用“还原显示默认值”动作；它们不是顶部根。窄屏溢出仍由共享 Ribbon 收纳原入口，实际溢出未运行。

十项显示属性、三种筛选模式和八个条件操作符分别属于配置/数据契约，不能加到根入口数。样式项来自文档资源及内置目录，动态 tile 数也不作为完成证明。Canvas 右键/浮动设置根继续进入同一 Drawer。

## 模块与实际行为

按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，Vue/Ant UI 位于 `components/spreadsheet` 与 `features/slicer`；防御性读取、diff、owner 与提交编排可独立测试。权威日期 predicate、报告聚合、缓存、命令与历史复用主任务的服务。没有新增第二套日期引擎或 bulk WASM 协议。[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)继续约束来源查询和大样本计算。

- 设置读取真实 `getTimeline(id).getSettings()`；numeric levels 仅在 UI 边界转成本地字符串草稿。name、caption、两个粒度、style、serial scrollPosition、四开关分别 diff。caption fallback 是 `caption ?? name`，使用名称时保留 authored caption 缺省。
- 纯设置使用 `updateSettings` 的一次正式 settings 命令；显示与布局、锚定、纵横比、local lockedPosition 混合使用一次 `updatePresentation`。未变投影几何不回写作者 affine 数据。设置不更改共享日期筛选、来源或其他 view 的粒度和导航。
- “还原显示默认值”只改草稿，保留 name，level/selectionLevel=Month，清 caption、style、scrollPosition 和四开关的 authored overrides。own undefined 与 omitted/false 分别保留；无变化确定按钮禁用。
- 插入使用 `resolveSheetSlicerSource` 的稳定字段和显式 Pivot/Table/owned PivotChart reference；正式 `dataSource/settings/anchor|position/size` Info 进入 `insertTimeline/insertSlicer`，不再把 legacy `{type,pivotTableId,...}` 送入正式接口。`SlicerPanelInsert` 使用正式 CSS/anchor placement；原始 drawing/cache 模型的调用入口为本地 `insertSlicerControl`。
- 日期来源能力判断复用 `sheetCellRecordValue`。formatted/stored Date 和 typed cached Date 可提供 Timeline 字段；仅有“日期”标题、普通 number 或 ISO text 不冒充 typed Date。字段列表仍使用现有主线程扫描，未声称改进了性能。
- 同步 retained range/condition 与异步 inclusive bounds 分开读取。来源失败或没有可靠 min/max 时仍显示作者筛选；新 range/condition 禁止提交。清除不依赖日期边界，但仍接受实际来源/报告权限和内容命令校验，不能绕过来源权限。
- 范围编辑两个真实 serial operands，结束标签明确“不含”。`useWholeDay` 保留作者 fractions，生效比较由既有引擎 floor(start)/ceil(end)；不会做 end−epsilon。“填入来源日期的整日范围”是明确用户动作，以 floor(min) 和 floor(max)+1 创建范围，不改变 inclusive 元数据。
- 固定条件有八个 operator，单值或双值 operands 原样保留；between/notBetween 包含两端。条件使用同一 Pivot 日期引擎和内部 shared intent，没有公开 Facade 条件 setter，没有以枚举当前成员代替条件。
- range/condition/none 共用捕获的 Studio、unit、sheet、user、revision、history generation、view/cache 和 UI generation。真实 BeforeCommand 以本次 payload 对象身份匹配；发布事件按进入执行的请求与 defensive cloned DTO 等价识别。clear 的 undefined 没有对象身份，因此以完整 params 对象匹配，旧 clear 不冒充新 clear。
- 来源/历史/用户/宿主变化、关闭和卸载撤销旧提交；bounds 请求有独立 generation，旧结果/旧 finally 不结束新查询或清新 busy。明确重新读取才重建 DateInput key。自己的 publish 可接受一次内容 revision，外部同命令/revision 不自动算自己的完成。
- 日期 Modal 保存使未提交父显示草稿 stale，保留草稿等待重新读取；仅更新父面板的 retained filter 摘要。取消、打开、读取、模式切换和主题切换均不直接修改内容历史；保存、undo/redo 和 JSON 重开沿用唯一历史源。

## UI、主题与输入

本批所有可替换控件使用 `antdv-next`：Drawer、Modal、Form/Form.Item、Input、InputNumber、DatePicker、Select、Checkbox、Button、Popover、Segmented、Tooltip 和 Alert。没有原生 form controls 或新增整页 contenteditable。画布绘制、命中、选区与拖动继续 Canvas 内核。

日期输入默认 Ant DatePicker，精确值方式使用 Ant InputNumber；未修改 raw fractions 不经过日历文本转换。Excel Date1900 的 serial day 60 使用最小 Ant Input，原因是 Dayjs/DatePicker 不能表示不存在于公历的 `1900-02-29`；它仍复用既有 Office calendar converter。compatibility 文本在无效输入时保持本地 pin，不因值变 undefined 卸载输入。明确选择其他日期才退出兼容日桥。

ConfigProvider 与 `theme.useToken()` 覆盖 Drawer、Modal、反馈、DatePicker/Select 弹层、文字和焦点状态。关闭日期 Modal 只向当前有效 owner、同一宿主且仍连接的 Ant 触发按钮归还焦点；顶层 Drawer 继续工作台 Canvas 焦点恢复协议。

Form 捕获 composition；IME 期间阻止 Enter/Escape、确定、重新读取和非文本动作。外部来源变化时保留正在组合的输入，compositionend 后过期草稿等待明确重新读取。浏览器草稿的 synthetic events 不代替真实 Windows IME、Ant 关闭动画或系统焦点验收。

## 文件与接口

| 文件 | 本批责任 |
| --- | --- |
| [SlicerPanel.vue](../../packages/office-studio/src/components/spreadsheet/SlicerPanel.vue) | Timeline 正式显示设置、numeric grains、日期导航、默认动作、独立日期 Modal 和 retained 摘要 |
| [slicer-panel-state.ts](../../packages/office-studio/src/features/slicer/slicer-panel-state.ts) | typed Date 字段能力、正式 Info 插入和稳定来源解析 |
| [slicer-panel-command-owner.ts](../../packages/office-studio/src/features/slicer/slicer-panel-command-owner.ts) | patch/range/condition 或完整 params 身份与两个实际命令边界的 owner guard |
| [timeline-settings-panel-state.ts](../../packages/office-studio/src/features/slicer/timeline-settings-panel-state.ts) | 十属性 diff、防御性设置、校验、纯/混合单事务 |
| [timeline-range-panel-state.ts](../../packages/office-studio/src/features/slicer/timeline-range-panel-state.ts) | retained filter、真实 bounds、草稿/operand 校验、半开范围与条件/清除提交 |
| [use-timeline-date-filter.ts](../../packages/office-studio/src/features/slicer/use-timeline-date-filter.ts) | 独立查询和草稿 generation、真实 publish、stale/busy、composition、取消 |
| [TimelineDateInput.vue](../../packages/office-studio/src/features/slicer/TimelineDateInput.vue) | Ant calendar/serial/compatibility 输入适配与未修改 fractions 保留 |
| [TimelineDateFilterDialog.vue](../../packages/office-studio/src/features/slicer/TimelineDateFilterDialog.vue) | 三模式、八条件、inclusive metadata、主题与 owner 安全焦点 |
| [timeline-settings-ui-fixture.mjs](../../packages/office-studio/tests/fixtures/timeline-settings-ui-fixture.mjs) | 真 typed/cached Date、raw colliders、共享 views、样式、optional/default/权限样本 |
| [timeline-settings-workbench.html](../../packages/office-studio/tests/fixtures/timeline-settings-workbench.html) | 真实 Ant/ExcelWorkbench/Canvas2D/Facade/history，独立 retained state 和受控等待 gate |
| [timeline-settings-ui.contract.test.mjs](../../packages/office-studio/tests/timeline-settings-ui.contract.test.mjs)、[timeline-settings-ui.browser.test.mjs](../../packages/office-studio/tests/timeline-settings-ui.browser.test.mjs) | 新批未运行验证草稿 |
| 本文件、[Slicer 设置 UI](SHEET-25-slicer-settings-ui.md)、[连接 UI](SHEET-25-slicer-connections-ui.md) | 固定来源、接口衔接、分类与本批未验收记录 |

正式 `FSheetTimeline`、Info/Builder/exports、canonical 生命周期、命令/权限、源 query、引擎、渲染和 Worker 由主任务维护。本批新增服务接受 captured snapshot、设置/几何或日期 draft 和当前 UI 回调；日期 prepared 回调只标识真实待提交 DTO。正式 settings/date setter 返回 boolean acceptance；mixed presentation 返回 `Promise<this>`，拒绝通过异常传递。本批按这两类实际结果处理。

## 未运行草稿与样本

按 `rg -n '^test'` 静态读取声明，人工展开固定 each 参数，没有运行计数脚本：

| 新草稿 | 声明数 | 计划场景 | 状态 |
| --- | ---: | ---: | --- |
| `timeline-settings-ui.contract.test.mjs` | 27 | 97 | 全未运行 |
| `timeline-settings-ui.browser.test.mjs` | 21 | 59 | 全未运行 |
| 合计 | **48** | **156** | 不是通过数 |

合同覆盖真实 getter、十项显示/七个 optional、纯/混合/默认/no-op、正式来源/Info 插入、numeric grains、exact fractional bounds、whole-day、八条件、清除、Date1900/Date1904/serial60、typed cached Date、raw collision、无日期/权限、保存/undo/redo/JSON 重开、真实 Before listener、旧 clear/bounds/finally、host/history/user/source/close/disposal 与 composition。

浏览器草稿覆盖实际来源插入 Form 与 Canvas publish、classic/grid 两个 Timeline view、正式设置及样式、三种筛选、八条件、inclusive metadata、Date1900/Date1904、serial60 Ant bridge、未修改导航 precision、父草稿/焦点、旧查询与旧提交、synthetic composition、主题弹层、readonly disabled entry 与 API refusal。

真实合成样本复用 [timeline-filter-fixture.mjs](../../packages/office-studio/tests/fixtures/timeline-filter-fixture.mjs)：两个 worksheet reports、一个 owned report、linked chart、两个共享 Timeline 和一个不同源字段的普通 Slicer，十一条原始记录。`[Jan1+.5,Jan1+.75)` 的静态期望为三个报告各 80，只有 typed cached Date；end Date 金额 20 不含，raw number 1000、ISO text 2000、Error 3000 不进入 typed 日期谓词。whole-day 或 `[Jan1+.25,Jan1+.75]` inclusive between 的静态期望各 110。这些是草稿期望，未执行。JSON 保存重开不是 XLSX 或独立 Office 验收。

本批只统计上述两份新草稿，不重复计入[Timeline 核心批](SHEET-25-timeline-settings-and-date-filters.md)、[Slicer 设置 UI](SHEET-25-slicer-settings-ui.md)、[连接 UI](SHEET-25-slicer-connections-ui.md)、cache 选项、ΣValues 或主任务 Facade/Worker 草稿。

## 未验收与剩余缺口

1. 所有测试、类型、前端/Rust/WASM 检查、dev、服务、浏览器和基准均未运行；正式 Facade/Studio/exports 与旧 raw API 测试合并后仍需统一验证。
2. 真实 Windows IME、Ant DatePicker 时间输入/locale/时区、serial60 关闭动画和焦点、用户切换的生产响应式界面、窄屏与完整 Timeline 拖动/键盘/显示粒度闭环仍未验收。
3. XLSX timeline/cache/style parts、author fractional values、日期系统、affine 几何、条件的 parser/writer 保真及独立 Office 文件重开未验收。不能以本批 JSON 证明 Office 全面互操作。
4. 未公开 Pro 演示行为、更多日期/日历条件、完整来源诊断与文档自定义样式编辑仍按完整范围继续追踪；本批没有以按钮/提示或 permanent disabled 代替这些能力。
5. 来源扫描、bounds、报告 staging、Worker/Rust/WASM 取消、TypedArray leases、复制/传输字节、内存、首屏/滚动/编辑/缩放/帧时间与长任务没有真实大样本前后基准，保持未验收，不声称优化、达标或端到端零拷贝。

上述边界继续属于完整 Sheets / Docs / Pro 目标，没有缩减交付范围。
