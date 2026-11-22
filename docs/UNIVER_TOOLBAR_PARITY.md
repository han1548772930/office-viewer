# Univer 顶部工具栏对齐核对

核对日期：2026-11-21。固定 OSS 参考：`E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

本记录按 [完整功能清单](UNIVER_FEATURE_PARITY.md)、[源码规则](UNIVER_REFERENCE_RULES.md)核对顶部入口与实际行为。它记录部分实现及缺口，**不表示顶部所有功能或 Pro 能力已经对齐**。所有完整功能 ID 保持“未验收”；单个按钮、模型、Facade 或测试通过均不能代替整个功能的编辑、历史、保存、Office 往返、权限和性能验收。

2026-11-22 续作仅做实现、静态调用链审阅和编写用例，按用户安排尚未运行统一验证。第 5 节保留既有共享 Ribbon 检查的历史结果，不代表当前未提交改动已通过。

源码核对覆盖 Tab → Group → Item → 命令/服务，而非仅比较分类名称。上游 ID 与本项目独立实现的 ID 可不同；表中记录可观察行为的对应关系。Pro 在固定开源检出中没有菜单实现时，明确记为“公开菜单 ID 未核实”，不编造闭源 ID，也不把未知行为排除出最终范围。

## 1. 当前装配与证据边界

### 按屏幕顺序核对（顶行，待统一测试）

第一行从左到右核对 `examples/index.html`、`examples/src/global.css`、`examples/src/main.ts` 和 `examples/src/workbench-settings-control.tsx`。品牌只展示，不负责切换产品；品牌位宽 68px、高 20px，随后是 Sheets、Docs。顶行高 36px，导航高 35px；当前产品使用 2px 内侧下划线。示例保留 Office Studio 自有品牌和本地化产品名。产品导航现在反映 `#sheets` / `#docs`，支持直达、未知片段回退及浏览器历史；现有 Office 文件菜单是本项目扩展，不插在两个产品标签之间。

固定版分层：`examples/src/workbench-settings-control.tsx`、`workbench-settings.ts` 和 `main.ts` 管设置入口、当前值展示与存储；`examples/src/mount-example.ts` 调产品 API，不实现 Ribbon 或数字格式。这里也只在示例中保留宿主设置、品牌、路由、文件操作。当前有语言、地区、方向、七种主题色、外观、四种 Ribbon、界面显隐、八档缩放与重置，设备切换仍缺。设置样式按 `examples/src/global.css` 的 240px 菜单、184px 双栏标签与 112px 当前值宽度适配 antdv-next。

产品包对应：`packages/core/src/services/region/region.service.ts` 与 `packages/core/src/facade/f-univer.ts` 对应本地 `i18n/regional-settings.ts`、`api/office-studio.ts`、`api/facade.ts`。地区默认跟随语言，显式设置后独立；Facade 接受 `deDE` 等上游标识，Vue / Intl 使用 `de-DE`。Sheets 画布和无障碍文本、统计、格式预览、时间线刻度读同一地区；不改保存的单元格标量，不把 Docs 日期字段的语言改为地区。上游货币按钮还按 `packages/sheets-numfmt/src/commands/commands/set-currency.command.ts` 选择地区格式，本地对此尚未完整对齐。四种 Ribbon 的呈现属于 `platform/ribbon`，Docs/Sheets 显隐和缩放属于各自工作台，不是示例 CSS。产品切换、跨表缩放、RTL 焦点、设备和真实文件往返未统一验收；旧记录中的独立“双行工具栏”按钮为历史实现。

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

2026-11-22，COMMON-01：固定提交的 `packages/ui/src/views/workbench/Workbench.tsx` 缺省为 classic，`examples/src/workbench-settings.ts:getDefaultWorkbenchSettings` 显式选用 grid。本地嵌入组件保留 classic 缺省；可运行工作台首次使用 grid，顶栏“双行工具栏”可切换，偏好保存在浏览器而非文档。`tests/browser/office-workbench.spec.ts` 的 Edge 用例验证两种产品首次使用同一 grid、两行各 24px、标签区 32px、工具区 69px，以及换产品与刷新后保留布局。未据此验收各分组、Pro 注入及上下文入口。

- 应用壳：[App.vue](../examples/workbench/src/App.vue)；共享 Ribbon：[WorkbenchRibbon.vue](../packages/office-studio/src/platform/ribbon/WorkbenchRibbon.vue)、[RibbonItem.vue](../packages/office-studio/src/platform/ribbon/RibbonItem.vue)、[schema/状态/溢出算法](../packages/office-studio/src/platform/ribbon/ribbon-schema.ts)。
- Sheets：[sheet-ribbon.ts](../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) → [ExcelWorkbench.vue](../packages/office-studio/src/components/ExcelWorkbench.vue) 的 `executeRibbon`。分类为开始、插入、公式、数据、视图；顺序来自 `order`。
- Docs：[document-ribbon.ts](../packages/office-studio/src/components/document/document-ribbon.ts) → [DocumentWorkbench.vue](../packages/office-studio/src/components/DocumentWorkbench.vue) 的 `onRibbonExecute` / `execute`。分类为开始、插入、布局、视图；传统分页与现代文档的大纲默认值分别处理。
- 公共 Ribbon 只展示当前模型/选区计算的 schema 并派发命令。按钮、选择器、下拉及主按钮加下拉复用同一控件；`hidden`、`disabled`、`activated`、`value` 独立。控件层防止禁用项派发，但业务命令权限仍须由命令服务验证。
- 普通格式操作、部分结构操作仍由工作台控制器映射到既有模型服务；排序、图表、透视及内容控件等使用对应 Facade。**全量 API、快捷键、右键和 Ribbon 是否共用同一命令注册及相同权限，尚须逐项核验，不能仅由 `commandId` 文本一致推断。**
- Pro 图表、透视分析、切片器、内容控件及其他功能以功能清单和对应当前证据记录跟踪；完整 Pro 范围仍未验收。

## 根入口统计与布局

2026-11-22 通过 `scripts/audit-univer-oss-ribbon.mjs` 静态复算固定 v1.6.0 OSS schema：Sheets 常规根声明 63 项，按 START/INSERT/FORMULAS/DATA/VIEW 为 30/9/12/6/6；Docs 为 38 项，按 START/INSERT/VIEW 为 25/10/3，合计 101 项。Sheets VIEW 的第六项来自 `packages/sheets-crosshair-highlight/src/menu/schema.ts:ToggleCrosshairHighlightOperation.id`，不能继续列为本地独有扩展。Docs 另有页眉页脚 5 项、目录 2 项上下文根声明，Sheets 本次 OSS schema 未发现 contextual tab 根声明。此前 62/38 是历史计数。两产品常规路径未发现重复，但 factory 状态与动态菜单没有执行，静态声明不等于当前可见入口。

脚本另外记录 Sheets 130、Docs 143 个右键/浮动菜单层级节点；该数包含父菜单和子项，不是根入口数或功能数量。Docs 还有 3 个颜色菜单展开表达式，须继续追踪 factory。More 仅改变同一入口的展示位置，不增加计数。本地 schema 静态读取结果为 Sheets 68 项（29/6/11/17/5），Docs 50 项（33/14/3），包含 `hidden` 条件项；不作为实际可见数量。完整差异见[顶部入口表](feature-evidence/TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)。

历史 Pro 统计来自 Univer 1.5.2 官方发布包：20 个菜单 schema 包、24 个根节点（Sheets 15、Docs 9）只属于 1.5.2。2026-11-22 使用 `scripts/count-univer-pro-ribbon-schemas.mjs .research.local/univer-1.6.0/packages --bundle-only --inventory`，已静态读取缓存中的 60 个 v1.6.0 ES 包，未导入执行商业 SDK。当前已识别根声明的包包括 Docs callout/chart/code/history/latex/print/quote/reference/sdt/shape/table，以及 Sheets chart/history/outline/pivot-chart/pivot/print/shape/slicer/sparkline；共享 shape-editor 还声明形状文本的粘贴入口。

同一功能在组级 schema、完整 Ribbon schema 或替代挂载分支中重复注册，不能将声明次数相加为可见根入口总数。`docs-list-ui`、`docs-formula-ui` 等未识别 `mergeMenu` 根项，仍须追踪组件、标记菜单、浮动编辑器和替换注册；不是无入口结论。embed、pivot、latex、code、shape-editor 仍有未解析表达式。既有 [research manifest](../.research.local/univer/resources.json) 的历史记录与早期三个 1.6.0 包复核不代表本次全部缓存清点；目前不推导完整 Pro 运行态总数。

| v1.6.0 Pro 上下文来源（各包 `lib/es/index.js`） | 静态根路径数 | 统计边界 |
| --- | ---: | --- |
| `docs-sdt-ui` | 13 | 内容控件类型与属性条件须逐项核对 |
| `docs-shape-ui` | 12 | 形状对象与文字上下文分支须逐项核对 |
| `docs-table-ui` | 17 | 排除子菜单；重复路径合并，模式与选区状态未执行 |
| `sheets-shape-ui` | 13 | 不与常规 Insert 根项相加 |
| `sheets-slicer-ui` | 4 | 切片器/时间线对象条件仍待运行核实 |

经典布局和网格布局共享同一套入口。Classic 依序呈现分类标签、单行工具区；`order` 决定同组顺序。Grid 使用每组两行网格，`gridLayout.row` / `column` 从 1 开始，支持 `rowSpan`、`columnSpan`、`width`、`minWidth`、`iconSize` 和 `showLabel`。上下文标签按对象/选区条件显隐，宽度不足时按上游逻辑整组进入 More；入口总数不因当前布局、溢出或重复呈现增加。各项位置和交互仍须由对应 schema、命令状态及动态演示逐项核实。

官方 Docs/Sheets 功能目录分别收录 73 和 83 项能力。这是产品功能数，不是工具栏入口数。Microsoft Office 补充能力另列为后续，不加入上述 Univer/Pro 统计。

## 2. 共享工具栏与跨产品能力

| 功能 ID | 上游依据 / 入口 | 本地入口与真实行为 | 缺口与验收边界 |
| --- | --- | --- | --- |
| COMMON-01、COMMON-03 | [Workbench/UIParts][workbench]；[Ribbon][ribbon]；[菜单类型][menu]；[菜单注册与分组][menu-manager] | classic 分类及工具组居中，36px/40px；工作台示例的 grid 使用 32px/69px 双行布局，支持网格占位/跨行、密度压缩和整组溢出 | 两种布局共享命令；simple/collapsed、完整插件生命周期和所有产品上下文仍未验收 |
| COMMON-03 | [ToolbarItem → ICommandService][toolbar-item] | 格式选择器由当前选区反推；混合值独立表示；相同 item 在原工具行与更多菜单派发相同 ID/参数；禁用按钮/选项不能派发；下拉 Escape 归还触发器焦点 | 功能命令权限不能由控件禁用代替；所有跨入口语义仍须核验 |
| COMMON-02、COMMON-05 | 命令/历史服务与各产品 schema | 顶部撤销/重做依赖文档历史状态；模型由当前应用服务保存 | 完整 Facade 兼容、协作 mutation/逆 mutation、跨文档会话和 savepoint 回归未整体验收 |
| COMMON-04 | 共享主题/locale/方向公开规格 | Ribbon 有方向参数、RTL 键盘支持；中文可访问名称 | 完整主题/暗色/语言切换、所有弹层与编辑器主题同步未验收 |
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
| SHEET-02 | [Sheets schema][sheet-schema] 的合并/取消合并与编辑/选区控制器 | 全部/水平/垂直合并、跨多个选区取消合并、分类清除已接共享命令和 Facade；保留首项非空内容，确认捕获原目标/范围/版本，检查完整受影响合并的权限和数据验证；Canvas 范围选择、名称框与公式栏继续保留 | [合并/清除证据](feature-evidence/SHEET-02-merge-clear.md)含源码差异与未运行用例；完整输入/填充、旧模型诊断、XLSX 独立重开和大范围内存未验收 |
| SHEET-03 | `InsertRowBeforeCommand`、`InsertColBeforeCommand`、`RemoveRowConfirmCommand`、`RemoveColConfirmCommand`、行列尺寸与显隐菜单 | 插入页：`sheet.command.insert-row` / `insert-column` / `delete-row` / `delete-column` 可修改真实模型 | 移动/复制、隐藏/显示、手动尺寸/自适应与所有引用资源变换仍不完整 |
| SHEET-04 字体 | `sheet.command.set-range-font-family` / `set-range-fontsize` / `set-range-bold` / `set-range-italic` / `set-range-underline` | 字体、字号、字号增减和字符属性进入共享 feature 命令、Facade、选区权限及内容历史；字号增减按活动格 ±1，控件读混合态；删除线及下划线按实际文字度量绘制 | [字体证据](feature-evidence/SHEET-04-font-family-size.md)中的契约/浏览器用例尚未运行；完整默认/主题继承、富文本范围格式及 Office 独立往返仍未验收 |
| SHEET-04 颜色 | `sheet.command.set-range-text-color`、`sheet.command.set-background-color` 及重置子项 | `SheetColorControl` 使用 antdv-next 色板、自定义、最近颜色和独立文字/填充重置，共享命令与 Facade 保存显式颜色/清除标记 | [颜色证据](feature-evidence/SHEET-04-color-reset.md)记录默认/重置差异；主题继承、范围主题、系统 IME/焦点和全部 Office 色彩字段仍未验收 |
| SHEET-04 对齐/样式 | `sheet.command.set-horizontal-text-align` / `set-vertical-text-align` / `set-text-wrap` / `set-shrink-to-fit` / `set-text-rotation` | 方向菜单含水平、左右斜 45°、竖排、左右转 90°；wrap 与 shrink-to-fit 互斥；统一减小字号和整块文字旋转进入共享布局命令、view/model、Canvas2D/WebGPU scene 与历史 | [旋转/缩小适应证据](feature-evidence/SHEET-04-text-rotation-shrink-to-fit.md)保留未验收；编辑输入桥不随文字方向旋转，复杂 shaping/富文本、主题、IME、真实 XLSX 重开和大样本均待统一验证 |
| SHEET-04 边框 | `sheet.command.set-border-basic` 与 `border.menu.ts` | 15 个位置项、12 个线型、颜色/最近颜色、主按钮重用上次样式，模型、Facade、权限、历史和 Canvas 共用边框 feature | [边框证据](feature-evidence/SHEET-04-borders.md)记录用例与文件边界；隐藏轴空洞、完整导入样本及大范围 Worker/内存仍未验收 |
| SHEET-05 | 富文本/Docs 编辑桥 | 当前工具栏作用于单元格样式 | 单格多运行、部分文字格式、富文本复制/链接及 Office 往返未验收 |
| SHEET-06 | [数字格式 schema][numfmt-schema]、`number-format-options.ts:MENU_OPTIONS`：格式菜单与百分比/货币/增减小数位共 5 个根项 | 格式菜单为 13 个格式项、4 个分隔、1 个“更多格式”；面板另有 6 类。当前已接精确格式码、地区货币、负值/小数位预览、自定义历史及完整编辑文本；地区和 1900/1904 进入解析及渲染，值与格式一次历史事务 | 入口与编辑来源、实现差异和待验证用例见[数字格式输入](#sheet-number-input)；全部地区、真实系统 IME、权限/焦点/主题、稀疏大范围格式与 Office 往返仍未整体验收 |
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

## 5. 既有共享 Ribbon 检查（历史结果）

```powershell
node --experimental-strip-types --test packages/office-studio/tests/ribbon-component.browser.test.mjs packages/office-studio/tests/ribbon-schema.contract.test.mjs
pnpm --filter @office-studio/office-studio check:types
```

核对结果：共享 Ribbon 契约和浏览器测试 **24/24 通过**，包括 12 条 schema 契约和 12 条实际 Edge/Chromium 组件浏览器测试；workspace 类型检查通过。覆盖 classic 和显式 grid 变体的定位、密度压缩、整组溢出、选择器参数及 split 主动作/子项派发。浏览器没有截图或像素比较。

- 契约：顺序稳定且不修改输入、隐藏/空分类、禁用项保留展示、方向/Home/End/RTL、上下文恢复、宽度/分组间隔/更多按钮预留、溢出每项仅出现一次、禁用和过期选项不派发、split 主动作与参数边界。
- 浏览器：36px/40px 实际几何；隐藏分类不出现；roving 焦点；字体选择真实命令参数；加粗激活态与禁用撤销；窄屏“更多功能”里的真实命令、嵌套颜色菜单与字体 Select；Escape 归焦；无文档水平溢出；上下文移除回退；菜单暂空重建后的 ResizeObserver 与宽窄切换弹层状态。无 `pageerror`。
- 浏览器暴露并修复了真实可访问性问题：所用 Select 的 `aria-label` 没有转发到内部 combobox，现用明确 input ID 与原生 `label for` 关联。隐藏测量树为 `aria-hidden` + `inert` + 零尺寸裁剪，不额外占文档滚动宽度。

这些测试验证共享控件及派发契约，**不证明表中每个业务功能已经实现**。Sheets/Docs 集成、透视/切片器面板、Office 重开及性能结果由对应验收记录另列。没有同数据集的设备/帧时间/长任务/内存前后对照，本记录不作性能达标或优化完成声明。

### 本轮按位置复核（未统一测试）

- 共享“更多”：固定版 `packages/ui/src/views/components/ribbon/RibbonOverflow.tsx` 渲染同一 `ToolbarItem`，不是把入口展开成菜单命令。对应本地 `platform/ribbon/RibbonOverflowPanel.vue`，classic/grid 复用实际按钮、选择器、拆分按钮与弹层，分组折行；simple 的测宽键区分同一命令在不同标签的位置。弹层嵌套、键盘焦点、窄屏尺寸仍待浏览器核验。
- Sheets Home 第一组：`packages/ui/src/menu/schema.ts` 给出撤销/重做的 18px 图标与无文字的格式刷/清除格式；粘贴跨行来自 `packages/sheets-ui/src/menu/schema.ts`。对照 `ToolbarItem.tsx`、`format-painter.menu.ts` 和 `format-painter-session.service.ts`，本地支持单击启用或取消、双击升级连续模式、成功后单次退出及 Escape；修改目标前检查大小、工作表与选区权限。`tests/browser/sheet-toolbar.spec.ts` 补连续应用及撤销用例，尚未执行。Docs 及图形的格式刷适配器不在这一条完成记录内。

- Docs Home 第一组按网格读：首行撤销 → 格式刷 → 跨行粘贴，下行重做 → 清除格式。固定版 `packages/docs-ui/src/controllers/format-painter.controller.ts`、`commands/commands/format-painter.command.ts`、`services/selection/word-boundary.ts` 对应本地 `features/document/services/document-format-painter.ts` 与 `executeDocumentCommand`。正文/表格文字按首字符或光标前字符取样；包含段落标记才携带段落/列表格式；点词与拖选均在指针释放时应用，保留目标文字与链接。清除格式走同一目标范围及权限检查。`packages/docs-ui/src/menu/context-menu.ts:PasteRibbonMenuFactory`、`services/clipboard/paste-options.ts` 对应本地 Ribbon 与右键菜单共用的三种 `{ value }` 命令参数。形状编辑、富文本页眉页脚、导入文档默认样式和复杂编号定义没有完成验证；新增契约和浏览器用例未运行。

### Home 字体组第一行：实现和待验收项

| 网格位置（从左到右） | v1.6.0 证据（相对 `packages/`） | 产品实现及边界 |
| --- | --- | --- |
| Docs r1c1–2 段落样式 | `docs-ui/src/menu/schema.ts`、`ui/src/views/font-size/interface.ts`、`ui/src/views/HeadingItem.tsx` | 正文、标题、副标题、H1–H5 的 8 项菜单按命名样式预览；本地 H6、引用移到段落工具的扩展菜单，命令保留。 |
| Docs r1c3–6 / Sheets r1c1–4 字体 | `ui/src/services/font.service.ts`、`ui/src/views/font-family/FontFamily.tsx`、`docs-ui/src/menu/menu.ts` | 共享有序字体目录，支持按实例追加或覆盖，保留当前导入字体。输入按显示名不区分大小写匹配，Enter 确认。字体探测提示、不同文字脚本的独立字体字段和 DOCX 往返仍未验收。 |
| Docs r1c7–8 / Sheets r1c5 字号 | `ui/src/views/font-size/{interface.ts,FontSize.tsx}`、`sheets-ui/src/menu/font.menu.ts` | 共用 15 个预设，输入草稿与选区分离。Docs 确认范围 1–400，Sheets 为 6–400；Escape、失焦放弃。 |
| Sheets r1c6、c7 增大、减小字号 | `sheets-ui/src/commands/commands/inline-format.command.ts` | 以活动单元格字号 ±1 后应用于范围，不再逐格跳预设档位；工作表与选区保护在修改前检查。单元格内部富文本选择与继承样式仍待闭环。 |

上述入口在产品包，示例无覆盖。`docs-ui/src/utils/east-asian-text.ts` 与 `commands/commands/inline-format.command.ts` 是中西文分槽的行为依据：文档归一化读取 `fontFamilyAscii/fontFamilyEastAsia`，Canvas 按字形选字体，光标前为 CJK 字母时显示东亚字体；包含 CJK 的选区指定字体会同时更新两个槽。写出层独立写 `w:rFonts/@w:eastAsia`，标题和副标题写入 Word 命名样式。复杂脚本、主题字体、标点分类与 Office 独立样本仍待核查。`ribbon-fonts.contract.test.mjs`、`document-toolbar.browser.test.mjs`、Rust 写出再解析用例等已补，按统一测试安排未运行。各布局的视觉、IME、主题与保存重开均不据此标为通过。

Docs 同行向右，段落组 r1c1 对照 `docs-ui/src/menu/menu.ts:AlignMenuItemFactory`：左、中、右、两端对齐的顺序、按钮图标、受保护选区禁用与选中态走真实段落状态，不由菜单记忆上次点击。r1c2 无序列表及 r2c1 有序列表已使用拆分按钮、各六种四行预览与九级列表定义；连续/重启命令与 DOCX 编号往返继续按[列表证据](feature-evidence/DOC-05-checklist.md)补齐。增强列表的全部模式、混合上下文、富剪贴板与真实 Office 应用重开仍未验收。

<a id="sheet-number-input"></a>

### Sheets 数字格式入口与编辑（待统一验证）

固定版 `packages/sheets-numfmt-ui/src/menu/schema.ts` 的 NUMBER 组顺序为格式菜单（order 9、r1c1 跨 4 列、160px），再百分比/货币/增小数位/减小数位（order 9.1–9.4，r2c1–4）。`menu/number-format-options.ts:MENU_OPTIONS` 的顺序是常规、文本；数字、百分比、科学计数；会计专用、财务数值、货币、整数货币；日期、时间、日期时间、持续时间；更多格式。分号在这里表示 4 个分隔区：13 个格式项和 1 个 More 不另计 Ribbon 根项。`views/SheetNumfmtPanel.tsx` 的 6 个面板分类与此菜单数量分开；classic 与 grid 复用相同入口，混合值与保护选区来自当前模型和命令状态。

`packages/sheets-numfmt-ui/src/utils/temporal-editor-value.ts`、`controllers/numfmt.editor.controller.ts`、`packages/core/src/shared/numfmt/{parse-value,serial-date,date-time-edit-pattern,run-part}.ts` 及对应 `__tests__` 是编辑与日期来源。本地 `sheet-number-format-edit.ts` 提供 `sheetNumberEditText` / `parseSheetNumberEdit`，`sheet-cell-edit.ts` 负责类型归一化与一次内容/格式提交，`sheet-cell-edit-session.ts` 捕获会话基线；公式栏与 Canvas 的 `antdv-next` 输入复用该会话。百分比展开完整十进制意图，F2 光标在百分号前；地区与当前年在开始时捕获，校验和提交使用相同设置。`shared/decimal-input.ts` 同时保护新输入和无 rawValue 的旧数值文本，无法保持完整十进制意图时保留文本。

有效日期系统统一按工作表显式值、工作簿继承值、1900 缺省解析，Canvas、公式栏、F2、粘贴与验证候选采用相同值。编辑会话记录 documentId/sheetId/key、源值、公式、类型和数字格式；目标内容、类型、数字格式或日期系统外部改变时拒绝提交，保留草稿与选区。无关单元格、字体/颜色及公式缓存重算不使会话失效；提交保留这些最新字段。切换文档/工作表或目标越界时清理会话和 composition 状态；行列头选择、表切换/删除/隐藏、查找定位、Enter/Tab 和保存先检查提交结果。这是本地避免旧草稿覆盖新内容的策略；固定 `sheets-ui/src/controllers/editor/editing.render-controller.ts` 可核实目标存在、原始值比较、无改动不写入及拒绝提交不移动，不能据此推断其实现相同冲突协议。没有 documentId 的旧宿主仍按既有 sheet 身份判断文档，稳定身份与跨用户并发协议另按 COMMON-05/07 验收。

日期格式显示完整可编辑日期，时间补齐秒及所需秒小数；日期编辑保留原隐藏时间，时间编辑保留原隐藏日期。上游 date 分支对显式 datetime 输入同样保留原隐藏时间，datetime 分支使用新时间；本地维持此已知契约，Microsoft Office 在同类输入下的效果尚未独立核实。持续时间本地按实际 `[h]` / `[m]` / `[s]` 单位解释，秒小数最多 10 位；这比上游固定小时序列和毫秒编辑文本更精细，但尚无真实 Office 样本验收。普通数字不通过裸小时持续时间推断。规范整秒只在原 serial 等于其规范编码时修正浮点误差，真实小秒小数保留；显示舍入可以跨日，零日历值按选定日期系统解码。

所有可替换编辑控件仍由 `antdv-next` 承载。其 AutoComplete 的公开 `focus()` 没有文本选区参数或 `nativeElement`，因此 F2 在既有 Canvas 宿主内查找库内 input 后调用 `setSelectionRange`；没有另建原生输入控件。主题、输入法、焦点归还和保护拒绝需最后统一浏览器验证。`number-format-input.contract.test.mjs`、`sheet-cell-input.contract.test.mjs`、`sheet-cell-edit-session.contract.test.mjs` 与 `sheet-number-format.browser.test.mjs` 已补地区、百分比、两日期系统、elapsed 单位、精度/舍入、外部更新冲突、缓存刷新、历史及重开用例，**本轮均未运行**；Rust/WASM、Office 独立重开和全部地区样本也未运行。格式仍有枚举坐标和 100,000 格上限，稀疏范围样式、Worker 及真实大样本性能未验收，没有优化达标结论。

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
