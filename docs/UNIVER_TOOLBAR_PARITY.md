# Univer 顶部工具栏对齐核对

核对日期：2026-11-21。固定 OSS 参考：`E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

本记录按 [完整功能清单](UNIVER_FEATURE_PARITY.md)、[源码规则](UNIVER_REFERENCE_RULES.md)核对顶部入口与实际行为。它记录部分实现及缺口，**不表示顶部所有功能或 Pro 能力已经对齐**。所有完整功能 ID 保持“未验收”；单个按钮、模型、Facade 或测试通过均不能代替整个功能的编辑、历史、保存、Office 往返、权限和性能验收。

源码核对覆盖 Tab → Group → Item → 命令/服务，而非仅比较分类名称。上游 ID 与本项目独立实现的 ID 可不同；表中记录可观察行为的对应关系。Pro 在固定开源检出中没有菜单实现时，明确记为“公开菜单 ID 未核实”，不编造闭源 ID，也不把未知行为排除出最终范围。

## 1. 当前装配与证据边界

### 按屏幕顺序核对（顶行，待统一测试）

第一行从左到右核对 `examples/index.html`、`examples/src/global.css`、`examples/src/main.ts` 和 `examples/src/workbench-settings-control.tsx`。品牌只展示，不负责切换产品；品牌位宽 68px、高 20px，随后是 Sheets、Docs。顶行高 36px，导航高 35px；当前产品使用 2px 内侧下划线。示例保留 Office Studio 自有品牌和本地化产品名。产品导航现在反映 `#sheets` / `#docs`，支持直达、未知片段回退及浏览器历史；现有 Office 文件菜单是本项目扩展，不插在两个产品标签之间。

最右侧设置已接通语言、LTR/RTL、明暗外观和 grid/classic 布局偏好，方向独立于语言。控件使用 antdv-next Dropdown/Menu/Button，Provider 主题派生顶栏颜色。与上游相比，设备、地区、主题色、simple/collapsed、UI chrome、缩放及重置偏好仍需逐项对齐；不将设置菜单已出现视为完成。浏览器用例已改用新菜单，并增加产品顺序、路由与前进/后退契约，尚未运行。旧记录中的独立“双行工具栏”按钮为历史实现。

### 2026-11-22 入口位置修正（待统一测试）

Sheets 开始/格式的边框面板已接入，15 个位置项和 12 个线型均不另计 Ribbon 根项；源码、Canvas 行为和文件边界见 [SHEET-04 边框记录](feature-evidence/SHEET-04-borders.md)。

以下源码路径均相对固定的 Univer v1.6.0 检出。此批没有运行测试，不改变功能验收状态。

| 区域 | 上游依据 | 本地变更与边界 |
| --- | --- | --- |
| Sheets 开始 | `packages/sheets-ui/src/menu/schema.ts` 的 LAYOUT、OTHERS | classic 中缩小字体填充排在文字方向前；保护范围恢复 `r1c1` 跨两行，额外的清除内容移到其后。命令不变。 |
| Sheets 插入 | `packages/sheets-data-validation-ui/src/menu/schema.ts` 的 CELL；`packages/ui/src/views/components/ribbon/ribbon-grid-layout.ts` | 复选框、下拉列表、日期按列排列为 `r1c1/r2c1/r1c2`；已接入分类清除、下拉预设和创建后定位规则的编辑入口。[SHEET-13 待验证记录](feature-evidence/SHEET-13-data-validation.md)。 |
| Sheets 公式 | `packages/sheets-formula-ui/src/menu/schema.ts` 的 BASIC、OTHERS，`menu/menu.ts` 的 `AllFunctionsMenuItemFactory` | 常用分类占 `r1c1` 两行，打开完整函数面板的入口收进其菜单；其余分类自动排列。计算模式占工具组两行两列，本地编辑公式、立即计算保留在其右侧，不占用分类位置。本地函数目录、计算设置 UI 和上游并非完全等价。 |
| Docs 插入 | `packages/docs-thread-comment-ui/src/menu/schema.ts` 的 OTHERS | 评论面板与添加评论恢复 `r1c1/r2c1`，均单行显示文字，沿用 `docs.operation.toggle-comment-panel` 与 `docs.operation.start-add-comment` 的状态和派发；不据此验收线程协作。 |

共享 Ribbon 继续使用 `antdv-next` Button/Menu/Popover/Tooltip。跨行菜单允许 schema 关闭标签；色块边框继承 ConfigProvider 派生的主题色。统一测试时检查双行/经典和窄屏 More 中的命令、按键焦点、明暗主题及文档选区保留。

2026-11-22，COMMON-01：固定提交的 `packages/ui/src/views/workbench/Workbench.tsx` 缺省为 classic，`examples/src/workbench-settings.ts:getDefaultWorkbenchSettings` 显式选用 grid。本地嵌入组件保留 classic 缺省；可运行工作台首次使用 grid，顶栏“双行工具栏”可切换，偏好保存在浏览器而非文档。`tests/browser/univer-workbench.spec.ts` 的 Edge 用例验证两种产品首次使用同一 grid、两行各 24px、标签区 32px、工具区 69px，以及换产品与刷新后保留布局。未据此验收各分组、Pro 注入及上下文入口。

- 应用壳：[App.vue](../examples/workbench/src/App.vue)；共享 Ribbon：[WorkbenchRibbon.vue](../packages/office-studio/src/platform/ribbon/WorkbenchRibbon.vue)、[RibbonItem.vue](../packages/office-studio/src/platform/ribbon/RibbonItem.vue)、[schema/状态/溢出算法](../packages/office-studio/src/platform/ribbon/ribbon-schema.ts)。
- Sheets：[sheet-ribbon.ts](../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) → [ExcelWorkbench.vue](../packages/office-studio/src/components/ExcelWorkbench.vue) 的 `executeRibbon`。分类为开始、插入、公式、数据、视图；顺序来自 `order`。
- Docs：[document-ribbon.ts](../packages/office-studio/src/components/document/document-ribbon.ts) → [DocumentWorkbench.vue](../packages/office-studio/src/components/DocumentWorkbench.vue) 的 `onRibbonExecute` / `execute`。分类为开始、插入、布局、视图；传统分页与现代文档的大纲默认值分别处理。
- 公共 Ribbon 只展示当前模型/选区计算的 schema 并派发命令。按钮、选择器、下拉及主按钮加下拉复用同一控件；`hidden`、`disabled`、`activated`、`value` 独立。控件层防止禁用项派发，但业务命令权限仍须由命令服务验证。
- 普通格式操作、部分结构操作仍由工作台控制器映射到既有模型服务；排序、图表、透视及内容控件等使用对应 Facade。**全量 API、快捷键、右键和 Ribbon 是否共用同一命令注册及相同权限，尚须逐项核验，不能仅由 `commandId` 文本一致推断。**
- Pro 图表、透视分析、切片器、内容控件及其他功能以功能清单和对应当前证据记录跟踪；完整 Pro 范围仍未验收。

## 根入口统计与布局

历史 desktop OSS preset 计数基于前一参考提交：Sheets 62 个、Docs 38 个，合计 100 个。入口数量按顶层 schema 节点统计，split/menu/select 的子项不重复计数；宽度引起的 More 溢出不另计入口。Sheets 分别为 START 30、INSERT 9、FORMULAS 12、DATA 6、VIEW 5；Docs 分别为 Home 25、Insert 10、View 3。v1.6.0 已确认这些清单引用的 OSS 源码路径存在，但尚未重新合并全部插件 schema 并复算根入口总数，因此这些数字不应当作 v1.6.0 的最终统计。

历史 Pro 统计来自 Univer 1.5.2 官方发布包：当时识别 20 个 Docs/Sheets 菜单 schema 包和 24 个 Pro Ribbon 根节点，其中 Sheets 15 个、Docs 9 个，至少 6 个包另导出 contextual Ribbon schema。此计数只属于 1.5.2，不代表 1.6.0。1.6.0 registry 元信息及 `sheets-chart-ui`、`sheets-pivot-ui`、`sheets-slicer-ui` 三个发布 bundle 已静态审阅。通过同一 AST 计数脚本只对这三个 bundle 得到 5 个常驻 Ribbon 根项：图表 Insert > Media 1 项、透视表 Data > Analysis / Organization 各 1 项、切片器 Data > Analysis 2 项。统计脚本没有把 contextual groups 计入这 5 项。其余 1.6.0 Docs/Sheets Pro 包及上下文 schema 未完成清点，因此不推导全量 1.6.0 Pro 总数。已审阅的 1.5.2 包记录见 [research manifest](../.research.local/univer/resources.json)，1.6.0 只读缓存位于 `.research.local/univer-1.6.0`。

经典布局和网格布局共享同一套入口。Classic 依序呈现分类标签、单行工具区；`order` 决定同组顺序。Grid 使用每组两行网格，`gridLayout.row` / `column` 从 1 开始，支持 `rowSpan`、`columnSpan`、`width`、`minWidth`、`iconSize` 和 `showLabel`。上下文标签按对象/选区条件显隐，宽度不足时按上游逻辑整组进入 More；入口总数不因当前布局、溢出或重复呈现增加。各项位置和交互仍须由对应 schema、命令状态及动态演示逐项核实。

官方 Docs/Sheets 功能目录分别收录 73 和 83 项能力。这是产品功能数，不是工具栏入口数。Microsoft Office 补充能力另列为后续，不加入上述 Univer/Pro 统计。

## 2. 共享工具栏与跨产品能力

| 功能 ID | 上游依据 / 入口 | 本地入口与真实行为 | 缺口与验收边界 |
| --- | --- | --- | --- |
| COMMON-01、COMMON-03 | [Workbench/UIParts][workbench]；[Ribbon][ribbon]；[菜单类型][menu]；[菜单注册与分组][menu-manager] | classic 分类及工具组居中，36px/40px；工作台示例的 grid 使用 32px/69px 双行布局，支持网格占位/跨行、密度压缩和整组溢出 | 两种布局共享命令；simple/collapsed、完整插件生命周期和所有产品上下文仍未验收 |
| COMMON-03 | [ToolbarItem → ICommandService][toolbar-item] | 格式选择器由当前选区反推；混合值独立表示；相同 item 在原工具行与更多菜单派发相同 ID/参数；禁用按钮/选项不能派发；下拉 Escape 归还触发器焦点 | 功能命令权限不能由控件禁用代替；所有跨入口语义仍须核验 |
| COMMON-02、COMMON-04 | 命令/历史服务与各产品 schema | 顶部撤销/重做依赖文档历史状态；模型由当前应用服务保存 | 完整 Facade 兼容、协作 mutation/逆 mutation、跨文档会话和 savepoint 回归未整体验收 |
| COMMON-05 | 共享主题/locale/方向公开规格 | Ribbon 有方向参数、RTL 键盘支持；中文可访问名称 | 完整主题/暗色/语言切换、所有弹层与编辑器主题同步未验收 |
| COMMON-06、COMMON-07 | 权限、协作公开规格；产品权限菜单 | 目前未发现满足服务端持久权限/多人协作验收的顶部工作流 | 不能把隐藏按钮、只读 UI 或本地通信记为完整能力 |
| COMMON-08、COMMON-09 | 持久历史、版本对比/恢复公开规格 | 撤销/重做与持久历史分开，当前顶部尚无完整版本服务面板 | 版本列表、语义对比、恢复后广播、服务重启后可查仍待实现/验证 |
| COMMON-10 | 跨文档嵌入公开规格 | 无已验收顶部嵌入流程 | 父子模型、引用/快照语义、权限/焦点/保存未验收 |
| COMMON-11 | [本项目 Canvas/WebGPU/WASM 标准](RENDERING_AND_COMPUTE_STANDARD.md) | Ribbon 外编辑内容由 Canvas；存在真实 WebGPU/2D 后端。渲染任务已覆盖部分冻结、缩放、隐藏轴、命中与输入桥契约 | 工具栏和布局测试不能证明大文件性能；本轮无完整前后性能测量，全部性能结论保留未验收 |
| COMMON-12 | 图片/资源公开模型及 drawing 插件 | Docs 图片插入可接 URL/本地资源；有现有资源模型 | 图片复制、裁剪、旋转、引用计数、所有 Office 关系和跨产品往返未验收 |

## 3. Sheets 顶部功能逐项映射

上游分类由共享 START / INSERT / FORMULAS / DATA / VIEW / OTHERS 与插件实际菜单共同形成。[Sheets 实例装配][sheets-mount] 已注册筛选、排序、数据验证、条件格式、结构化表格、超链接、评论等；因此只读取 `sheets-ui` 本身不能代表完整工具栏。

| 功能 ID | 上游菜单 / 命令依据 | 本地 item / 真实行为 | 尚缺行为或验证 |
| --- | --- | --- | --- |
| SHEET-01 | [Sheets schema][sheet-schema] 的 `CopySheetCommand`、`RenameSheetOperation`、`SetWorksheetHideCommand`、`SetTabColorCommand` 及 SheetBar | 工作表新增、激活、重命名、删除、隐藏/恢复主要位于组合底栏；不另在顶部重复一套 | 复制/重排/标签颜色、veryHidden 与所有关联引用事务仍未整体验收 |
| SHEET-02 | [Sheets schema][sheet-schema] 的合并/取消合并与编辑/选区控制器 | `sheet.command.merge` / `unmerge`；Canvas 范围选择、名称框与公式栏 | 合并横向/纵向变体、多选、填充及完整编辑边界未全部覆盖 |
| SHEET-03 | `InsertRowBeforeCommand`、`InsertColBeforeCommand`、`RemoveRowConfirmCommand`、`RemoveColConfirmCommand`、行列尺寸与显隐菜单 | 插入页：`sheet.command.insert-row` / `insert-column` / `delete-row` / `delete-column` 可修改真实模型 | 移动/复制、隐藏/显示、手动尺寸/自适应与所有引用资源变换仍不完整 |
| SHEET-04 字体 | `sheet.command.set-range-font-family` / `set-range-fontsize` / `set-range-bold` / `set-range-italic` / `set-range-underline` | 字体、字号、字号增减、加粗/斜体/下划线/删除线；选区反推激活态，删除线及下划线按实际文字度量绘制 | 完整继承/主题、富文本范围样式及样式 Office 独立往返未验收 |
| SHEET-04 颜色 | `sheet.command.set-range-text-color`、`sheet.command.set-background-color` 及重置子项 | `sheet.menu.text-color`、`sheet.menu.fill` → `set-color` / `set-background`，有限颜色列表可执行 | 主题颜色、完整色板/自定义颜色、重置/继承语义与范围主题未完整实现 |
| SHEET-04 对齐/样式 | `sheet.command.set-horizontal-text-align` / `set-vertical-text-align` / `set-text-wrap` / `set-shrink-to-fit` / `set-text-rotation` / `set-border-basic` | 文字方向菜单含水平、左右斜 45°、竖排、左右转 90°；wrap 与 shrink-to-fit 互斥；角度、逐字竖排和按度量压缩已接入 view/model、Canvas2D/WebGPU scene 与历史 | SHEET-04 仍未验收；DOM 编辑输入桥不随文字方向旋转，混合范围只通过契约测试，第三方 XLSX 重开及设备级 WebGPU 像素验证未完成；完整边框样式/颜色、默认样式/范围主题仍缺 |
| SHEET-05 | 富文本/Docs 编辑桥 | 当前工具栏作用于单元格样式 | 单格多运行、部分文字格式、富文本复制/链接及 Office 往返未验收 |
| SHEET-06 | [数字格式 schema][numfmt-schema]：`sheet.operation.open.numfmt.panel`、百分比/货币/增减小数位命令 | 常规/数字/货币/会计/百分比/日期/时间/日期时间/文本共 9 类，小数位增减；保留 raw 类型和精确格式、1900/1904，数值型文本不擅自转型 | 完整自定义格式面板、全部 locale/会计布局、导入自定义格式重新编辑及 Office 往返未整体验收 |
| SHEET-07 | 冻结工具菜单、`ToggleGridlinesCommand`、`ZOOM_RATIO_MENU_ID` | `toggle-gridlines`；冻结选区/首行/首列/取消；放大/缩小/100%；各表选区与缩放会话 | 全部冻结分区/编辑/图形/打印组合仍待浏览器与文件样本验收；不由基础命中测试推断全部通过 |
| SHEET-08 | [Sheets schema][sheet-schema]：`PASTE_RIBBON_MENU_ID` 与 PasteValue/Format/ColWidth/BesidesBorder；[公式 schema][formula-schema] 的公式粘贴 | `sheet.command.copy` / `cut` / `paste` 有真实模型操作 | 仅值/仅公式/仅格式/列宽/排除边框等选择性粘贴、HTML/TSV/引用与权限闭环未全部连通 |
| SHEET-09 | [公式 schema][formula-schema]：`formula-ui.operation.insert-function.{common,financial,logical,text,date,lookup,math,statistical,engineering,information,database}`、`formula-ui.operation.more-functions` | 公式页新增真实函数选择/参数对话框，当前 30 个函数逐项通过引擎求值测试；同时保留编辑公式/重新计算，提交进统一内容历史 | 全部函数分类/目录、定义名称管理、计算选项及完整依赖/结构引用仍未验收；30 个函数不等于完整公式引擎 |
| SHEET-10（Pro） | 高级公式公开函数与行为目录；具体商业菜单 ID 未核实 | 有既有公式引擎，顶部目前仍是编辑/重算 | 不能用两个入口或 SUM 等函数样例声称全部高级公式已对齐；目录/数组/引用/错误/性能逐项待验收 |
| SHEET-11 | [筛选 schema][filter-schema]：`sheet.command.smart-toggle-filter` / `clear-filter-criteria` / `re-calc-filter` | 数据页筛选面板：按值多选、搜索分页、15 种条件、明确比较值类型、多列 AND、清列/清全部/移除；一次命令历史，Canvas/命中/键盘/统计使用独立 filteredRows | 列头下拉、颜色/完整日期规则、结构变换/权限/大样本与完整 XLSX 重开仍未验收 |
| SHEET-12 | [排序 schema][sort-schema]：`sheet.menu.sheets-sort` → `sheet.command.sort-range-{asc,asc-ext,desc,desc-ext,custom}` | 数据页升降序及自定义多关键字面板，复用 `FWorksheet.sortRange`；标题行/比较语言/取消进度/过期结果，单次历史；降序空白末尾，引用转换失败不部分写回 | 扩展选区交互、全部比较/locale/关联资源变换、统一权限和大样本仍未整体验收 |
| SHEET-13 | [验证 schema][validation-schema]：`sheet.menu.data-validation` / `quick-dropdown` / `quick-date`、`data-validation.operation.open-validation-panel` | 数据页新增“数据验证”实际规则面板；模型/命令/API、键入/TSV 粘贴/Facade 写入校验、错误级别反馈、内联下拉输入提示、XLSX writer 和 parser 规则读回已接入；[进度记录](feature-evidence/SHEET-13-data-validation.md) | 日期快捷选择、复选框控件绘制、列表搜索/彩色标签及完整公式覆盖未接；parser 目前以本地 XLSX fixture 验证，第三方 Office 往返尚未验收；SHEET-13 仍未验收 |
| SHEET-14 | [条件格式 schema][cf-schema]：`sheet.operation.open.conditional.formatting.panel`、AddDataBarConditionalRule / AddIconSetConditionalRule | 当前无对应本地顶部 item | 规则管理、优先级/公式、色阶/数据条/图标集、动态渲染和保存待实现/验收 |
| SHEET-15 | 查找替换插件 schema | `sheet.operation.find`：真实匹配/定位 | 当前主要为查找；替换、作用域/完整选项、取消和大样本尚缺 |
| SHEET-16 | sheets-hyper-link-ui schema 与插入/编辑弹层 | 当前无完整本地超链接顶部工具 | URL/内部范围/富文本链接、编辑/移除、引用变换和安全跳转未验收 |
| SHEET-17 | sheets-note-ui、sheets-thread-comment-ui | 当前无已验收顶部批注/持久线程评论流程 | Note/ThreadComment 分模型、锚点/权限/回复/解决/服务同步未验收 |
| SHEET-18 | [结构化表格 schema][table-schema]：`sheet.operation.open-table-selector` | 当前无对应本地顶部 item | 列/表头/汇总行、样式、自动扩展、结构化引用、筛选排序与导出未验收 |
| SHEET-19 | drawing/watermark 插件 | 当前表格顶部插入对象以图表为主，非图片/水印的完整 UI | 图片插入/替换/裁剪/锚点、文字/图片水印、打印与文件资源保真待验收 |
| SHEET-20 | sheets-crosshair-highlight schema | 当前无对应顶部十字高亮控件 | 颜色/冻结/滚动兼容及纯视图 operation 待验收 |
| SHEET-21（含 Pro） | drawing 基础 + 公开形状目录；完整商业菜单 ID 未核实 | 当前无完整本地形状顶部流程 | 连接器/文字/分组/旋转/层级/命中/Office 映射待实现/验收 |
| SHEET-22（Pro） | 公开图表规格；完整商业菜单 ID 未核实 | 插入 `sheet.operation.insert-chart`；可创建既有 bar/line 图表模型并绑定数据；已有图层选择/移动/缩放 | 全部图型、系列/分类/轴/图例/标签/主题属性面板、透视图/打印及可编辑 Office 往返未验收 |
| SHEET-23（Pro） | 公开透视表规格；商业菜单 ID 未核实 | 插入/数据页共用 `sheet.operation.open-pivot` 和真实配置面板；model/engine/字段配置、刷新/下钻已有局部契约，切片器浏览器用例操作真实透视结果 | 透视面板完整配置验收、全部汇总/分组/小计/下钻、Office Cache/Parts 和大规模性能仍未完成 |
| SHEET-24（Pro） | 公开透视图规格；商业菜单 ID 未核实 | 当前无透视图顶部/编辑流程 | 透视与图表绑定、字段联动、筛选刷新、持久化/Office 输出待实现/验收 |
| SHEET-25（Pro） | 公开切片器/时间线规格；商业菜单 ID 未核实 | 插入/数据页共用切片器/时间线面板；已用浏览器验证透视绑定、多选/清除、1900 日期边界、历史和内部快照重开 | 结构化表格/多个目标/透视图绑定、全部日期语义、样式布局、浮动画布对象及 XLSX parts 往返仍未整体验收 |
| SHEET-26（Pro） | 公开迷你图规格；商业菜单 ID 未核实 | 当前无对应本地顶部 item | 完整图型、空值/负值/比例、分组、引用更新和 Office 往返待实现/验收 |
| SHEET-27（Pro） | 公开分组规格；商业菜单 ID 未核实 | 当前无对应本地顶部 item | 行列嵌套层级、折叠展开、汇总方向与隐藏原因合并未验收 |
| SHEET-28（Pro） | 公开打印/PDF 规格；商业菜单 ID 未核实 | 当前没有完整 Sheets 页面设置/分页预览/输出工作流 | 打印区域、适页、重复标题、页眉页脚、图形和隐藏处理、实际输出待实现/验收 |
| SHEET-29 | 本项目 Office parser/writer 与公开导入导出规格 | 应用壳负责打开/保存/导出；真实 XLSX writer 已接入，JSON 不冒充 Office 输出 | XLS/XLSX/CSV/TSV 公开契约、保留 XLSB/ODS 读取、全部资源/样式/公式独立重开仍未验收 |
| SHEET-30 | [Sheets schema][sheet-schema]：`AddRangeProtectionFromToolbarCommand` 等保护菜单 | 当前无完整顶部权限流程 | 服务端工作簿/表/范围权限、可执行操作、失败原子性/协作及恢复待验收 |

## 4. Docs 顶部功能逐项映射

[Docs schema][doc-schema] 的开始/插入/视图由各插件注入；[示例装配][docs-mount] 还包含 drawing、hyper-link、thread-comment 和 TOC。正文目录与左侧导航是两个能力；页眉页脚和 TOC 都有专用上下文分类，不能把普通插入按钮当作已覆盖。

| 功能 ID | 上游菜单 / 命令依据 | 本地 item / 真实行为 | 尚缺行为或验证 |
| --- | --- | --- | --- |
| DOC-01 | 文档模型/排版与 [Docs schema][doc-schema] | 传统分页 Canvas、现有段落/文字运行模型；工具栏不以 HTML 作为权威源 | 节、复杂自动排版/跨页对象、所有模式差异与大文档性能未整体验收 |
| DOC-02 字符 | 上游 `SetInlineFormat*Command`，包括 `doc.command.set-inline-format-subscript` / `superscript` | 样式/字体/字号、bold/italic/underline/strike、颜色/高亮/上下标及清除字符格式；上下标有独立 run 属性、命令历史、Canvas 字号/基线、DOM 语义 HTML 和 DOCX `w:vertAlign` 写出 | 上下标实现切片见 [DOC-02](feature-evidence/DOC-02-vertical-align.md)；完整字符样式/主题/继承及生产 WASM/Office 往返仍未验收 |
| DOC-02 段落 | `SetParagraphNamedStyleCommand`、`AlignOperationCommand`、段落设置面板 | 正文/H1–H6、四对齐、缩进/列表、行距；`doc.command.format-paragraph` | 段前/后间距面板正在补充；Title/Subtitle 独立命名样式、混合/全段落语义及 Office 往返未完整验收 |
| DOC-03 | `doc.menu.breaks`；`doc.menu.section-break.{next-page,continuous,next-column,even-page,odd-page}`；`docs.operation.insert-column-break` / `insert-section-break` | 当前普通 `doc.command.insert-page-break`；布局纸张 A4/Letter、纵横方向；边距设置正在补充 | 分节各类型、节级分栏、列断点均未接顶部；普通分页符不能代替分节 |
| DOC-03 页眉页脚 | `doc.command.open-header-footer-panel`；[headerFooter schema][header-footer]：`docs-ui.ribbon.header-footer`、`DocHeaderFooterRibbonControl.{insertion,navigation,options,position}`、`doc.command.close-header-footer` | 当前无已实现本地页眉页脚上下文分类 | segment 切换、节关联、首页/奇偶页、页码/字段、距离、导航和结束编辑全部未验收 |
| DOC-04 左侧大纲 | DocSideMenu/SideMenu；固定模式默认值 | `doc.operation.toggle-outline`；真实标题/层级/activeId、跳转、停靠/浮动、导航键盘、传统默认关闭/现代默认开启 | 宽表与浮动大纲实际裁剪交集尚未接；正文标题折叠、表格内标题和 Title/Subtitle 保真未验收 |
| DOC-04 正文 TOC | [TOC schema][toc-schema]：`doc.menu.insert-table-of-contents`；`doc.operation.open-table-of-contents-dialog`；`ribbon.table-of-contents` 上下文 | 当前没有正文 TOC 插入、更新或上下文工具 | 不能以左侧大纲替代正文目录；目录内容/页码更新、删除、保存/Office 输出待实现/验收 |
| DOC-05（含 Pro） | `doc.command.bullet-list` / `order-list` / `check-list` / `change-list-nesting-level`，`CheckListCommand` 属于 START/LAYOUT | 本地 bullet/numbered 和缩进可执行；本轮将查找映射到 OTHERS 语义组、段落/列表映射到 LAYOUT；复选列表由 `doc.command.set-block-type` 创建、独立 `doc.command.toggle-checklist-item` 勾选，Canvas 点击和 Enter 共享文档历史；DOCX 使用原生编号符号 | 清单切片证据见 [DOC-05](feature-evidence/DOC-05-checklist.md)；嵌套/升降级、编号重启/连续、增强样式、重排、富剪贴板及真实 Office 应用重开尚未验收 |
| DOC-06（含 Pro） | 表菜单 `doc.menu.table`；上下文 `DocTableInsertRowAbove/Bellow/ColumnLeft/RightCommand`、`doc.table.delete-{rows,columns,table}` | 实际插表；表内显示本项目“表格工具”上下文，前后插行/左右插列/删行列/删表，更新控件锚点及一次文档历史，Tab 末格追加复用命令 | 不冒称上游存在同名同形 tab；合并拆分/移动/尺寸边框/整表对齐/跨页嵌套、公开 Facade 和 Office 往返仍未完整覆盖 |
| DOC-07 | docs-drawing-ui image schema | `doc.operation.open-image`，本地图片和 URL 输入，插入真实模型 | 环绕/锚点/偏移/裁剪/旋转/层级/替换、分页随对象更新及所有关系重开未验收 |
| DOC-08 | docs-hyper-link-ui：`ShowDocHyperLinkEditPopupOperation` | `doc.operation.open-link` 可插入链接并保留范围 | 完整编辑/删除、书签/内部目标、位置变换和 Office 往返未验收 |
| DOC-09 | docs-find-replace 与共享 find-replace schema | `doc.operation.toggle-find`，正文匹配/定位当前实现 | 完整替换、选项、表格/页眉页脚等作用域、可撤销全部替换未验收 |
| DOC-10 | docs-thread-comment-ui 浮动工具/绘图上下文 | 当前无完整本地顶部/浮动持久评论流程 | 锚点变换、回复/解决、用户/权限/服务同步未验收 |
| DOC-11（Pro） | 内容控件公开规格；完整商业菜单 ID 未核实 | 内容控件 split 与六种实际类型直达，真实面板/命令/历史；已验证锁、段落拆分锚点、撤销/重做与内部重开 | 图片/分组/重复节/文档部件完整 UI、跨段 SDT、全部编辑变换及 DOCX 往返仍未验收 |
| DOC-12（Pro） | 脚注尾注公开规格；商业菜单 ID 未核实 | 当前无脚注/尾注顶部编辑流程；writer 现有 notes 写出不等于编辑器能力 | 规范模型/引用编号/注释区排版/导航/命令历史/导入重开待实现/验收 |
| DOC-13（Pro） | 代码块公开规格；商业菜单 ID 未核实 | 当前无代码块顶部 item | 语言/源文本/高亮/块编辑/复制与保存待实现/验收 |
| DOC-14（Pro） | 多栏组公开规格；商业菜单 ID 未核实 | 当前无多栏组顶部 item | 块级 2–5 栏、比例/间距、选择/导航、增删、保存；与节级分栏分开验收 |
| DOC-15（Pro） | LaTeX 公开规格；商业菜单 ID 未核实 | 当前无公式插入/编辑顶部 item | 源文本、行内/块级形式、布局/错误/持久化/Office 映射待实现/验收 |
| DOC-16（Pro） | 引用块公开规格；商业菜单 ID 未核实 | 命令转换允许 quote 类型，但当前没有已验收引用块顶部工作流 | 容器内部富文本、进入/退出/删除/复制/持久化待验收；不能以单独左边框计完成 |
| DOC-17（Pro） | Callout 公开规格；商业菜单 ID 未核实 | 当前无 Callout 顶部 item | 图标/颜色/内容/容器属性、块编辑/布局/历史/保存待实现/验收 |
| DOC-18（含 Pro） | drawing 基础：`doc.command.menu-insert-shape` 及矩形/椭圆命令；完整商业/SmartArt 菜单 ID 未核实 | 当前无已连通形状/SmartArt 顶部流程 | 形状文字/锚点/层级/连接器/完整目录/SmartArt 布局和 Office 映射未验收 |
| DOC-19（Pro） | 共享/Docs 图表公开规格；商业菜单 ID 未核实 | 当前没有 Docs 图表顶部/属性面板 | 共享可编辑图表模型、内嵌数据、环绕/分页、属性与 Office 关系待实现/验收 |
| DOC-20 | watermark 插件与公开规格 | 当前无已验收水印顶部流程 | 文本/图片、透明度/定位/页范围、打印与 Office 输出未验收 |
| DOC-21（Pro） | 打印/PDF 公开规格；商业菜单 ID 未核实 | `doc.command.print` 当前通过 `documentHtml` 单 section 与 CSS 打印文档内容；不是应用壳截图，也不是 Canvas 逐页布局输出 | 真实分页布局与页范围、字体/对象/页眉页脚/脚注一致性、打印/PDF 独立文件仍未完成，不能记为 Pro 打印已达标 |
| DOC-22 | 本项目 duke-pages parser / DOCX writer | 应用壳打开/保存/导出真实 DOCX；本地 JSON 快照与 Office 输出分开 | DOC/DOCX/ODT 格式/节/列表/表格/图形资源的独立重开与完整保真未验收 |
| DOC-23 | DocFooter/DocStatistics/DocLayoutProgress | 组合底栏：当前/总页、字符/选区字素统计、缩放、内容修改状态；不把 UI 状态伪称云保存成功 | 全部上游统计口径、真实布局任务进度、大文档与模式差异未完整验收 |

公开基础插入项的符号/emoji 选择器现已接入，支持保留选区插入、分类/最近使用与自定义 Unicode。剪贴板新增剪切、源格式/纯文本粘贴，顶部与原生 paste 共用归一化，并拒绝过期异步结果；copy/cut 当前写出纯文本。完整富文本/媒体剪贴板、目标格式及最近粘贴模式切换仍未完成，目标继续保留。

## 5. 本次已经执行的共享 Ribbon 检查

```powershell
node --experimental-strip-types --test packages/office-studio/tests/ribbon-component.browser.test.mjs packages/office-studio/tests/ribbon-schema.contract.test.mjs
pnpm --filter @office-studio/office-studio check:types
```

核对结果：共享 Ribbon 契约和浏览器测试 **24/24 通过**，包括 12 条 schema 契约和 12 条实际 Edge/Chromium 组件浏览器测试；workspace 类型检查通过。覆盖 classic 和显式 grid 变体的定位、密度压缩、整组溢出、选择器参数及 split 主动作/子项派发。浏览器没有截图或像素比较。

- 契约：顺序稳定且不修改输入、隐藏/空分类、禁用项保留展示、方向/Home/End/RTL、上下文恢复、宽度/分组间隔/更多按钮预留、溢出每项仅出现一次、禁用和过期选项不派发、split 主动作与参数边界。
- 浏览器：36px/40px 实际几何；隐藏分类不出现；roving 焦点；字体选择真实命令参数；加粗激活态与禁用撤销；窄屏“更多功能”里的真实命令、嵌套颜色菜单与字体 Select；Escape 归焦；无文档水平溢出；上下文移除回退；菜单暂空重建后的 ResizeObserver 与宽窄切换弹层状态。无 `pageerror`。
- 浏览器暴露并修复了真实可访问性问题：所用 Select 的 `aria-label` 没有转发到内部 combobox，现用明确 input ID 与原生 `label for` 关联。隐藏测量树为 `aria-hidden` + `inert` + 零尺寸裁剪，不额外占文档滚动宽度。

这些测试验证共享控件及派发契约，**不证明表中每个业务功能已经实现**。Sheets/Docs 集成、透视/切片器面板、Office 重开及性能结果由对应验收记录另列。没有同数据集的设备/帧时间/长任务/内存前后对照，本记录不作性能达标或优化完成声明。

## 6. 更新方法

后续每项接入须同时更新本表中的本地入口、真实命令/模型行为和测试证据。仅添加菜单、API 名称、弹层或禁用入口时，继续记“接入中/未验收”。接口存在与真实可操作、撤销/重做、保存重开、Office 往返分别核验。商业菜单 ID/模式限制待核实时保留“未核实”，功能目标不删除。

[workbench]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/views/workbench/Workbench.tsx
[ribbon]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/views/components/ribbon/Ribbon.tsx
[menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/services/menu/menu.ts
[menu-manager]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/services/menu/menu-manager.service.ts
[toolbar-item]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/views/components/ribbon/ToolbarItem.tsx
[sheet-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/schema.ts
[formula-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-formula-ui/src/menu/schema.ts
[filter-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-filter-ui/src/menu/schema.ts
[sort-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-sort-ui/src/menu/schema.ts
[validation-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/menu/schema.ts
[cf-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/menu/schema.ts
[numfmt-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-numfmt-ui/src/menu/schema.ts
[table-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-table-ui/src/menu/schema.ts
[doc-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts
[header-footer]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/header-footer-ribbon.ts
[toc-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-toc-ui/src/menu/schema.ts
[sheets-mount]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/examples/src/sheets/mount.ts
[docs-mount]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/examples/src/docs/mount.ts
