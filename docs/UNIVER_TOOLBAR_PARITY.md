# Univer 顶部工具栏对齐核对

核对日期：2026-11-21。固定 OSS 参考：`E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

本记录按 [完整功能清单](UNIVER_FEATURE_PARITY.md)、[源码规则](UNIVER_REFERENCE_RULES.md)核对顶部入口与实际行为。它记录部分实现及缺口，**不表示顶部所有功能或 Pro 能力已经对齐**。所有完整功能 ID 保持“未验收”；单个按钮、模型、Facade 或测试通过均不能代替整个功能的编辑、历史、保存、Office 往返、权限和性能验收。

2026-11-22 起的续作仅做实现、静态调用链审阅和编写用例，按用户安排尚未运行统一验证。2026-11-25–26 新增富文本/默认来源/范围主题、稀疏复制填充、共享局部移位/确认、图表标题/绘图区/对象上下文及面板路由、Docs Canvas/词法/段落分页、LaTeX，以及 Sheets/Docs 图片和共享绘图工作台证据，所有功能仍未验收。第 5 节保留既有共享 Ribbon 检查的历史结果，不代表当前在途改动已通过。

源码核对覆盖 Tab → Group → Item → 命令/服务，而非仅比较分类名称。上游 ID 与本项目独立实现的 ID 可不同；表中记录可观察行为的对应关系。Pro 在固定开源检出中没有菜单实现时，明确记为“公开菜单 ID 未核实”，不编造闭源 ID，也不把未知行为排除出最终范围。

## 1. 当前装配与证据边界

### 按屏幕顺序核对（顶行，待统一测试）

第一行从左到右核对 `examples/index.html`、`examples/src/global.css`、`examples/src/main.ts` 和 `examples/src/workbench-settings-control.tsx`。品牌只展示，不负责切换产品；品牌位宽 68px、高 20px，随后是 Sheets、Docs。顶行高 36px，导航高 35px；当前产品使用 2px 内侧下划线。示例保留 Office Studio 自有品牌和本地化产品名。产品导航现在反映 `#sheets` / `#docs`，支持直达、未知片段回退及浏览器历史；现有 Office 文件菜单是本项目扩展，不插在两个产品标签之间。

固定版分层：`examples/src/workbench-settings-control.tsx`、`workbench-settings.ts` 和 `main.ts` 管设置入口、当前值展示与存储；`examples/src/mount-example.ts` 调产品 API，不实现 Ribbon 或数字格式。这里也只在示例中保留宿主设置、品牌、路由、文件操作。当前有语言、地区、方向、七种主题色、外观、四种 Ribbon、界面显隐、八档缩放与重置，设备切换仍缺。设置样式按 `examples/src/global.css` 的 240px 菜单、184px 双栏标签与 112px 当前值宽度适配 antdv-next。

产品包对应：`packages/core/src/services/region/region.service.ts` 与 `packages/core/src/facade/f-univer.ts` 对应本地 `i18n/regional-settings.ts`、`api/office-studio.ts`、`api/facade.ts`。地区默认跟随语言，显式设置后独立；Facade 接受 `deDE` 等上游标识，Vue / Intl 使用 `de-DE`。Sheets 画布和无障碍文本、统计、格式预览、时间线刻度读同一地区，快捷货币及高级面板的差异见 [SHEET-06 地区货币](feature-evidence/SHEET-06-regional-currency.md)。四种 Ribbon 的呈现属于 `platform/ribbon`，Docs/Sheets 显隐和缩放属于各自工作台。产品切换、跨表缩放、RTL 焦点、设备和真实文件往返未统一验收。

<a id="top-settings-and-feature-search"></a>

### 2026-11-26 顶栏语言与开始区搜索

- 固定版 `examples/src/{workbench-settings.ts,main.ts,mount-example.ts}` 的语言菜单采用该语言默认方向，地区可独立保持；`core/src/facade/f-univer.ts::setDirection` 只更新界面方向。当地语言菜单、方向菜单和 `FOffice.setDirection` 已连到同一个实例状态及现有偏好，偏好在重载后恢复。顶栏及两个产品的 ConfigProvider 使用 19 个 Ant 语言包，日期组件同时注册 Day.js 语言数据，依据项目 [ConfigProvider](llm.md#configprovider) 与日期组件说明。窗口方向与 Canvas 的坐标增量独立，画布滚动容器继续使用左起的物理坐标。设备切换会涉及固定版 MobileWorkbench、产品移动工具和触摸输入，仍待实现。此批未增加顶栏设置项。
- 固定版共享 `packages/ui/src/menu/schema.ts` 在 START/OTHERS 注册 `ui.operation.open-feature-search`，classic order 2，两产品静态 grid 均为第二行第二列，同组快捷键面板在第一行第二列。`feature-search.controller.ts` 注册 Ctrl/Cmd+Shift+P，`FeatureSearch.tsx` 读取当前 Ribbon 及上下文菜单、过滤禁用/隐藏、对命令与参数去重。当前 Sheets、Docs 各接入 1 个常驻搜索根，按钮和编辑区快捷键打开相同的 Ant Modal/Select，实时读取当前 Ribbon 及其上下文标签，按名称、分组路径和命令 ID 搜索，执行时回到当前产品原命令。更多菜单沿用同一根。专属弹窗中的仅组件操作、独立右键菜单、快捷键面板和从 Facade 打开搜索仍待接入。
- [生产工作台草稿](../tests/browser/office-workbench.spec.ts) 覆盖 RTL、地区与内容保持、偏好重载，以及搜索执行、焦点和撤销。[组件语言草稿](../packages/office-studio/tests/document-content-control-workbench.browser.test.mjs) 覆盖法语日期控件及日期保真。代码、文档和静态清点已核对，运行验证后置。

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

脚本另外记录 Sheets 130、Docs 143 个右键和浮动菜单层级节点，包含父菜单和子项。Docs 还有 3 个颜色菜单展开表达式。More 只改变已有根的展示位置。当前本地主标签、Pro 扩展和上下文根分开记录，包含条件声明，详见[顶部入口表](feature-evidence/TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)。P16 图表对象四根、LaTeX 上下文标签与 71 个模板、范围主题九区段、图表线型与共享绘图工作台接线沿用各自口径。相关来源见[对象菜单](feature-evidence/SHEET-22-chart-object-context-actions.md)、[公式模板](feature-evidence/DOC-15-latex-templates-ribbon.md)和[绘图入口](feature-evidence/SHEET-21-drawing-workbench-affine.md#entry-inventory)。

历史 Pro 统计来自 Univer 1.5.2 官方发布包：20 个菜单 schema 包、24 个根节点（Sheets 15、Docs 9）只属于 1.5.2。2026-11-22 使用 `scripts/count-univer-pro-ribbon-schemas.mjs .research.local/univer-1.6.0/packages --bundle-only --inventory`，已静态读取缓存中的 60 个 v1.6.0 ES 包，未导入执行商业 SDK。当前已识别根声明的包包括 Docs callout/chart/code/history/latex/print/quote/reference/sdt/shape/table，以及 Sheets chart/history/outline/pivot-chart/pivot/print/shape/slicer/sparkline；共享 shape-editor 还声明形状文本的粘贴入口。

2026-11-26 以同一脚本复算，当前缓存中 64 个 ES 包可从 `package.json` 核实为 1.6.0。`engine-pivot` 缓存缺少清单，报告单列待核实，其他包继续扫描。OSS 结果仍为 Sheets 63、Docs 38 和 Docs 上下文 7。64 统计的是版本可核实的包，静态注册、根入口去重、运行时显隐和功能验收各自保留证据。

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
| COMMON-03 | [ToolbarItem → ICommandService][toolbar-item]；Sheets/Docs selection/editor | 格式选择器反映实际范围；Sheets 完整多选区与 Tab 起点进入共享 store/operation，Canvas/公式栏和 Facade 复用单格输入命令；Docs caret marks 按文档隔离，body/story 校验 revision/generation，成功输入才推进范围 | [共享输入/选区/历史证据](feature-evidence/COMMON-03-input-selection-axis-history.md)记录静态接线和未运行草稿；真实系统 IME、全部快捷键/剪贴板、主题/焦点与所有跨入口语义仍未验收 |
| COMMON-02、COMMON-05 | 命令/历史服务、`core/facade/f-univer.ts` 与各产品 schema | 顶部/全局键盘/Facade 通过共享历史命令注册与文档隔离的 Pinia 内容历史；生产适配器检查来源 unit；FOffice 全局异步取消返回 false，同步 unit API 保留既有返回/取消契约；输入/选区通知延后到栈/revision 提交；同 unit 历史重入拒绝；BeforeUndo/Redo 改变目标/内容/历史后停止旧请求；保存点包含 history generation | [共享事务证据](feature-evidence/COMMON-03-input-selection-axis-history.md#5-历史观察事件与取消)含别名、生产 App、取消与观察者重入草稿，尚未运行；完整 Facade 兼容、协作 mutation/逆 mutation、持久历史、跨模块 savepoint 和大文件快照内存仍未整体验收 |
| COMMON-04 | 共享主题/locale/方向公开规格 | Ribbon 有方向参数、RTL 键盘支持；中文可访问名称 | 完整主题/暗色/语言切换、所有弹层与编辑器主题同步未验收 |
| COMMON-06、COMMON-07 | 权限、协作公开规格；产品权限菜单 | 目前未发现满足服务端持久权限/多人协作验收的顶部工作流 | 不能把隐藏按钮、只读 UI 或本地通信记为完整能力 |
| COMMON-08、COMMON-09 | 持久历史、版本对比/恢复公开规格 | 撤销/重做与持久历史分开，当前顶部尚无完整版本服务面板 | 版本列表、语义对比、恢复后广播、服务重启后可查仍待实现/验证 |
| COMMON-10 | 跨文档嵌入公开规格 | 无已验收顶部嵌入流程 | 父子模型、引用/快照语义、权限/焦点/保存未验收 |
| COMMON-11 | [本项目 Canvas/WebGPU/WASM 标准](RENDERING_AND_COMPUTE_STANDARD.md) | Ribbon 外编辑内容由 Canvas；存在真实 WebGPU/2D 后端。渲染任务已覆盖部分冻结、缩放、隐藏轴、命中与输入桥契约 | 工具栏和布局测试不能证明大文件性能；本轮无完整前后性能测量，全部性能结论保留未验收 |
| COMMON-12 | 图片/资源公开模型及 drawing 插件 | Sheets 浮动/单元格图片与背景、Docs 图片编辑/多选使用真实资源；作者 HTML 保留 affine/size/crop，PNG 编码有取消/释放，混合绘图用同一 native owner | [Sheets 图片 66 条草稿](feature-evidence/SHEET-19-images.md)、[Docs 批量 57 条新草稿](feature-evidence/DOC-07-image-editing.md#image-batch)、[共享工作台](feature-evidence/SHEET-21-drawing-workbench-affine.md)均未运行；所有 Office 关系、跨产品往返、系统剪贴板和性能未验收 |

## 3. Sheets 顶部功能逐项映射

上游分类由共享 START / INSERT / FORMULAS / DATA / VIEW / OTHERS 与插件实际菜单共同形成。[Sheets 实例装配][sheets-mount] 已注册筛选、排序、数据验证、条件格式、结构化表格、超链接、评论等；因此只读取 `sheets-ui` 本身不能代表完整工具栏。

| 功能 ID | 上游菜单 / 命令依据 | 本地 item / 真实行为 | 尚缺行为或验证 |
| --- | --- | --- | --- |
| SHEET-01 | [Sheets schema][sheet-schema] 的 `CopySheetCommand`、`RenameSheetOperation`、`SetWorksheetHideCommand`、`SetTabColorCommand` 及 SheetBar | 工作表新增、激活、重命名、删除、隐藏/恢复主要位于组合底栏；不另在顶部重复一套 | 复制/重排/标签颜色、veryHidden 与所有关联引用事务仍未整体验收 |
| SHEET-02 | [Sheets schema][sheet-schema] 的合并/取消合并与编辑/选区；AutoFill/Refill/CopyDown/CopyRight/AutoClear；四核心移位+四 confirm ID | 合并/清除、多选区/输入/填充走共享 planner 和内容历史；四方向移位新增同一完整候选/资源权限/验证，右键按固定子项次序走 confirm，纵向确认拒绝筛选尾部；FRange 正确使用 COLUMNS=0/ROWS=1，Ant 合并确认捕获内容/历史/宿主版本，observer 重入不发布旧 view | [合并/清除](feature-evidence/SHEET-02-merge-clear.md)、[共享输入/选区](feature-evidence/COMMON-03-input-selection-axis-history.md)、[自动填充](feature-evidence/SHEET-02-auto-fill.md)、[局部移位与 56 个草稿](feature-evidence/SHEET-02-cell-shift-transactions.md)均未运行；standalone refill receipt、完整 drawing/media/所有范围资源、大范围 Worker、全部序列/引用、IME/两后端和 XLSX 独立重开仍未验收 |
| SHEET-03 | 插删/移动/复制命令；`set-row-visible` / `set-col-visible`；行列尺寸与显隐菜单 | 结构、显隐、手动/自动尺寸与拖拽增量通过公共 API 注册；三个轴 controller 完成编辑后捕获目标，detached 准备失败不发布；结构 planner 联动可识别资源，隐藏采用 spans 并保留过滤/折叠/零尺寸原因；公开 h/ah/ia/w/hd 导入真实状态、来源存档随结构变换；观察者导航/重入后停止旧宿主投影 | [轴命令与事务证据](feature-evidence/COMMON-03-input-selection-axis-history.md#3-行列命令与-ui)、[公开轴与 17 个草稿](feature-evidence/SHEET-03-public-axis-import.md)未运行；完整 auto-height、全部资源/复杂公式变换、字形度量、Office 单位/独立重开和真实大样本仍未验收 |
| SHEET-04 字体 | `sheet.command.set-range-font-family` / `set-range-fontsize` / `set-range-bold` / `set-range-italic` / `set-range-underline` 等固定别名 | 整格字体同步对应 rich run 覆盖；编辑态按实际文字范围/typing 偏好进入草稿。有效字体读取默认/行列/theme/引用/direct，null 清作者属性并恢复继承；公开读回输出 ff/fs 等固定短字段与本地 nested 兼容字段；最终一次内容提交 | [字体](feature-evidence/SHEET-04-font-family-size.md)、[局部字体/Canvas](feature-evidence/SHEET-04-05-rich-text-font-editing.md)、[来源/引用/reset 与 58 个草稿](feature-evidence/SHEET-04-default-style-inheritance.md)、[短字段读回与 13 个草稿](feature-evidence/SHEET-04-public-style-readback.md)均未运行；全部范围主题/字体、复杂 shaping、真实 IME 和 Office 独立往返仍未验收 |
| SHEET-04 颜色 | `sheet.command.set-range-text-color`、`sheet.command.set-background-color` 及重置子项 | `SheetColorControl` 使用 antdv-next 色板、自定义、最近颜色和独立文字/填充重置，共享命令与 Facade 保存显式颜色/清除标记 | [颜色证据](feature-evidence/SHEET-04-color-reset.md)记录默认/重置差异；主题继承、范围主题、系统 IME/焦点和全部 Office 色彩字段仍未验收 |
| SHEET-04 对齐/样式 | `sheet.command.set-horizontal-text-align` / `set-vertical-text-align` / `set-text-wrap` / `set-shrink-to-fit` / `set-text-rotation` | 方向菜单含水平、左右斜 45°、竖排、左右转 90°；wrap 与 shrink-to-fit 互斥；统一减小字号和整块文字旋转进入共享布局命令、view/model、Canvas2D/WebGPU scene 与历史 | [旋转/缩小适应证据](feature-evidence/SHEET-04-text-rotation-shrink-to-fit.md)保留未验收；编辑输入桥不随文字方向旋转，复杂 shaping/富文本、主题、IME、真实 XLSX 重开和大样本均待统一验证 |
| SHEET-04 边框 | `sheet.command.set-border-basic` 与 `border.menu.ts` | 15 个位置项、12 个线型、最近颜色及主区重用；plan 读有效继承边，真实邻格权限及 border 作者组共享模型/Facade/历史/Canvas | [边框](feature-evidence/SHEET-04-borders.md)、[继承空格投影](feature-evidence/SHEET-04-default-style-inheritance.md)记录未运行用例；继承清除同步限 100,000 个位置，大范围 Worker、隐藏轴和真实 Office 未完成或未验收 |
| SHEET-04 范围主题 | 固定 `RangeThemeStyle` / `SheetRangeThemeModel` / Facade 与四个 worksheet-range-theme-style 命令 | 32 内置/九区段、目录注册、稀疏规则/隐藏条纹、区域快照 blocker、轴/工作表复制、共享 set/remove 与真实 Ant Modal；九区段 Ant 表单保存目录/规则为一次事务，覆盖所有同名引用权限与内容/历史/宿主 guards；局部移位与内部 cut/all 保留逻辑片段/作者来源；新增本地开始/工具第 3 根 | [基础 46 个未运行草稿](feature-evidence/SHEET-04-range-themes.md)、[九区段编辑与 45 个新草稿](feature-evidence/SHEET-04-range-theme-editor.md)、[局部结构 32 个草稿](feature-evidence/SHEET-04-range-theme-structure.md)、[共享移位/确认](feature-evidence/SHEET-02-cell-shift-transactions.md)；重叠采用后应用优先，全套固定 snapshot/Office、未知样式/完整混合态、两后端/IME/性能仍待实施或验收 |
| SHEET-05 | 富文本/Docs 编辑桥 | F2/公式栏按 UTF-16 输入轨迹重放文字和局部字体，保留运行身份、空标记、链接和扩展；覆盖模式清旧运行；Canvas 运行文字/选区/光标/滚动由实际布局驱动，越界拖选持续按帧更新，换行边界保留视觉行 affinity，antd 输入承接 native/IME；FRange.setRichTextStyle 仅改运行 | [输入轨迹](feature-evidence/COMMON-03-input-selection-axis-history.md#2-sheets-的共享模型与事务)与[55 个富文本草稿](feature-evidence/SHEET-04-05-rich-text-font-editing.md)均未统一验证；富剪贴板、局部链接、BiDi/复杂 shaping、混合字号 IME、两后端和 Office 往返仍未验收 |
| SHEET-06 | [数字格式 schema][numfmt-schema]、`number-format-options.ts:MENU_OPTIONS`：格式菜单与百分比/货币/增减小数位共 5 个根项 | 格式菜单为 13 个格式项、4 个分隔、1 个“更多格式”；面板另有 6 类。当前已接精确格式码、地区货币、负值/小数位预览、自定义历史及完整编辑文本；地区和 1900/1904 进入解析及渲染，值与格式一次历史事务 | 入口与编辑来源、实现差异和待验证用例见[数字格式输入](#sheet-number-input)；全部地区、真实系统 IME、权限/焦点/主题、稀疏大范围格式与 Office 往返仍未整体验收 |
| SHEET-07 | 冻结工具菜单、`ToggleGridlinesCommand`、`ZOOM_RATIO_MENU_ID` | `toggle-gridlines`；冻结选区/首行/首列/取消；放大/缩小/100%；各表选区与缩放会话 | 全部冻结分区/编辑/图形/打印组合仍待浏览器与文件样本验收；不由基础命中测试推断全部通过 |
| SHEET-08 | [Sheets schema][sheet-schema]：`PASTE_RIBBON_MENU_ID` 与 PasteValue/Format/ColWidth/BesidesBorder；[公式 schema][formula-schema] 的公式粘贴 | `sheet.command.copy` / `cut` / `paste` 有真实模型操作 | 仅值/仅公式/仅格式/列宽/排除边框等选择性粘贴、HTML/TSV/引用与权限闭环未全部连通 |
| SHEET-09 | [公式 schema][formula-schema]：`formula-ui.operation.insert-function.{common,financial,logical,text,date,lookup,math,statistical,engineering,information,database}`、`formula-ui.operation.more-functions` | 公式页新增真实函数选择/参数对话框，当前 30 个函数逐项通过引擎求值测试；同时保留编辑公式/重新计算，提交进统一内容历史 | 全部函数分类/目录、定义名称管理、计算选项及完整依赖/结构引用仍未验收；30 个函数不等于完整公式引擎 |
| SHEET-10（Pro） | 高级公式公开函数与行为目录；具体商业菜单 ID 未核实 | 有既有公式引擎，顶部目前仍是编辑/重算 | 不能用两个入口或 SUM 等函数样例声称全部高级公式已对齐；目录/数组/引用/错误/性能逐项待验收 |
| SHEET-11 | [筛选 schema][filter-schema]：`sheet.command.smart-toggle-filter` / `clear-filter-criteria` / `re-calc-filter` | 数据页筛选面板：按值多选、搜索分页、15 种条件、明确比较值类型、多列 AND、清列/清全部/移除；一次命令历史，Canvas/命中/键盘/统计使用独立 filteredRows | 列头下拉、颜色/完整日期规则、结构变换/权限/大样本与完整 XLSX 重开仍未验收 |
| SHEET-12 | [排序 schema][sort-schema]：`sheet.menu.sheets-sort` → `sheet.command.sort-range-{asc,asc-ext,desc,desc-ext,custom}` | 数据页升降序及自定义多关键字面板，复用 `FWorksheet.sortRange`；标题行/比较语言/取消进度/过期结果，单次历史；降序空白末尾，引用转换失败不部分写回 | 扩展选区交互、全部比较/locale/关联资源变换、统一权限和大样本仍未整体验收 |
| SHEET-13 | [验证 schema][validation-schema]：`sheet.menu.data-validation` / `quick-dropdown` / `quick-date`、`data-validation.operation.open-validation-panel` | 数据页新增“数据验证”实际规则面板；模型/命令/API、键入/TSV 粘贴/Facade 写入校验、错误级别反馈、内联下拉输入提示、XLSX writer 和 parser 规则读回已接入；[进度记录](feature-evidence/SHEET-13-data-validation.md) | 日期快捷选择、复选框控件绘制、列表搜索/彩色标签及完整公式覆盖未接；parser 目前以本地 XLSX fixture 验证，第三方 Office 往返尚未验收；SHEET-13 仍未验收 |
| SHEET-14 | [条件格式 schema][cf-schema]：DATA 管理、INSERT 数据条/图标集快插与 selector 子项 | 数据 / 规则中的 `sheet.operation.open-conditional-formatting` 打开 antdv-next 真实规则面板，`sheet.command.set-conditional-formats` 原子提交；模型/priority/公式基准、平均值六种比较与自动数据条边界接入计算和来源读写预算 | [格式刷/规则编辑](feature-evidence/SHEET-04-format-painter.md)及[Office 来源往返](feature-evidence/SHEET-14-conditional-format-office-roundtrip.md)记录实际接口；完整 selector/快插、所有公式/vendor 条件、Office 独立互操作、资源/峰值内存和浏览器草稿仍未验收 |
| SHEET-15 | 查找替换插件 schema | `sheet.operation.find`：真实匹配/定位 | 当前主要为查找；替换、作用域/完整选项、取消和大样本尚缺 |
| SHEET-16 | sheets-hyper-link-ui schema 与插入/编辑弹层 | INSERT/MEDIA 有 `sheet.operation.insert-hyperlink`，实际 Ant 弹窗及共享 set/remove 命令处理已有接线 | URL/内部范围/富文本链接、编辑/移除、引用变换和安全跳转未验收 |
| SHEET-17 | sheets-note-ui、sheets-thread-comment-ui | 当前无已验收顶部批注/持久线程评论流程 | Note/ThreadComment 分模型、锚点/权限/回复/解决/服务同步未验收 |
| SHEET-18 | [结构化表格 schema][table-schema]：`sheet.operation.open-table-selector` | 当前无对应本地顶部 item | 列/表头/汇总行、样式、自动扩展、结构化引用、筛选排序与导出未验收 |
| SHEET-19 | 固定 drawing 图片/背景 schema、clipboard/reference/placement 与 watermark 插件 | INSERT/MEDIA 图片 r1c1 跨两行，VIEW/DISPLAY 背景 r1c2 跨两行，各两子项；图片/背景命令与 Ant picker/editor/crop 已接；完整 affine HTML/PNG、Position/Both/None 结构与最终轴投影进入统一模型/事务 | [图片入口、作者保真与 66 条草稿](feature-evidence/SHEET-19-images.md)、[最终 filter/auto-height 和共享工作台](feature-evidence/SHEET-21-drawing-workbench-affine.md)均未运行；水印、Office、系统剪贴板、两后端/IME 和真实性能仍有缺口 |
| SHEET-20 | 固定 sheets-crosshair-highlight schema | VIEW/DISPLAY 十字高亮 split 及八种颜色子项已有 handler；纯视图 operation | 颜色/冻结/滚动兼容及视图/内容历史边界待验收 |
| SHEET-21（含 Pro） | 固定 drawing group/arrange/align schema、operations、共享 DrawingUpdateController、Sheets 持久化监听、MoveDrawingsCommand/UndoRedoGroupService 与公开形状目录 | typed 选区、组合/取消组合、混合排列/对齐/变换/删除、Canvas 浮条/右键及七项 FWorksheet wrapper 已接；固定 move-drawing 使用实际 skeleton 边界、1000ms 快照历史；真实 image/chart/shape owners 保留完整 affine，前后 footprint 权限与一次内容历史 | [原组合 45 条草稿](feature-evidence/SHEET-21-drawing-groups-and-mixed-order.md)、[工作台当前 82 条草稿及八个 operation ID 修正](feature-evidence/SHEET-21-drawing-workbench-affine.md)均未运行；[切片器/时间线 affine 与 Ant 控件、98/182 草稿](feature-evidence/SHEET-25-slicer-timeline-affine.md)已补入但未运行；shape 完整作者/渲染、组 clipboard、跨 workbook 依赖、极端边界、Office 与大样本仍缺 |
| SHEET-22（Pro） | 固定发布包类型、ES 与公开图表规格；完整商业动态入口仍未核实 | 45 图型目录复用真实模型/配置；标题/绘图区同源命中与对象四根已有接线；P17 的 tab/section/axisId/seriesId/titleKey 定位真实 Ant Drawer；P18 的 Line/Smooth/Step 接二维/3D 共享几何、空值及 Builder/FChart；同图表保留草稿，IME 期间延后路由；导航不提交内容 | [审计](feature-evidence/SHEET-22-chart-type-audit.md)、[资源](feature-evidence/SHEET-22-relation-word-cloud-resources.md)、[图例](feature-evidence/SHEET-22-legend-selection-wrap.md)、[标题配置](feature-evidence/SHEET-22-title-subtitle-contract.md)、[标题直接编辑](feature-evidence/SHEET-22-title-direct-editing.md)、[绘图区](feature-evidence/SHEET-22-plot-area-direct-editing.md)、[层级提示](feature-evidence/SHEET-22-hierarchy-ant-tooltip.md)、[P16 对象与 34 个草稿](feature-evidence/SHEET-22-chart-object-context-actions.md)、[P17 route 与 33 个草稿](feature-evidence/SHEET-22-chart-edit-panel-routing.md)、[P18 线型与 34 个新草稿](feature-evidence/SHEET-22-line-and-area-styles.md)均未验收；全部 element 命中/直接编辑、上游曲率/组合最终效果、跨 workbook 依赖、真实 chart-sheet/外部 drawing clipboard、XLSX 线型/Office、两后端及性能仍缺 |
| SHEET-23（Pro） | 公开透视表规格；固定 v1.6.0 发布 UI 的 15 项 Show Values As 及字段参数 | 插入/数据页共用 `sheet.operation.open-pivot` 和真实配置面板；[共享显示变换](feature-evidence/SHEET-23-24-pivot-show-data-as.md)与 [命令/Facade/linked 图格式](feature-evidence/SHEET-23-24-pivot-show-data-as-facade.md)已落盘，source/output 权限和 content/history/user/host 复用内容事务 | 15 项为字段枚举，不增加 Ribbon 根；全部草稿未运行，完整面板/展开折叠/日期条件/Pro API、Office Cache/Parts 与大规模性能仍待闭环 |
| SHEET-24（Pro） | v1.6.0 `sheets-pivot-chart` 来源/编辑类型、Facade 和发布 UI factories/schema/ES，具体标签/位置/条件见 [独立 UI 清点](feature-evidence/SHEET-24-pivot-chart-ui.md) | 既有根支持 [owned/linked 创建、字段+外观一次配置、Facade/事件](feature-evidence/SHEET-24-owned-pivot-charts.md)；[四字段区/11 汇总/标签与报告选项 UI](feature-evidence/SHEET-24-pivot-chart-ui.md)和 [15 项显示变换](feature-evidence/SHEET-23-24-pivot-show-data-as-facade.md)已落；open-pivot-chart-panel 为同一 UI operation 别名 | 上游两条互斥 Ribbon 根/两个创建子项、四个轴子项与动态层级单列；本地常规根增量 0。核心29/45、引擎30/51、UI43/70等草稿均未运行；同时创建表图/完整轴浮层与 Pro API/Office/真实 UI/大样本继续闭环 |
| SHEET-25（Pro） | 固定发布 `f-slicer.d.ts`、连接/来源模型类型、四根格式 schema 和 open-dialog operations，见[多目标连接证据](feature-evidence/SHEET-25-slicer-connections.md) | Table/Pivot affine/五类叶与 Ant 画布控件保留；共享缓存、替换/去重 linked/owned 连接、多报告筛选、Facade/权限/一次历史、复制/删除/结构生命周期已落盘；四条件格式根和连接 Ant 弹窗接入，常规根增量 0 | 原 98/182 及本批核心 25/31、UI/lifecycle 草稿均未运行；detach filter 语义待运行核实，完整交叉筛选成员状态、日期/Pro API、缺依赖跨 workbook、Office、两后端/IME/性能仍缺 |
| SHEET-26（Pro） | 固定 `sheets-sparkline-ui` Insert/Media，`sheet.operation.open-sparkline-selector` | 迷你图插入、跨表数据范围、四图型和分组设置可从实际工作台触发；上下文编辑、删除与公开模型使用同一数据 | [入口和设置](feature-evidence/SHEET-26-sparkline-settings-and-groups.md)、[XLSX 投影](feature-evidence/SHEET-26-sparkline-office-roundtrip.md)，文件独立重开、移动与真实性能未验收 |
| SHEET-27（Pro） | 固定 `sheets-outline-ui` Data/Organization，`sheet.toolbar.dimension-outline.group` | 数据整理、行列头、画布层级按钮和公开 Worksheet API 已接入；手动、筛选和大纲折叠分别保留 | [入口与模型](feature-evidence/SHEET-27-outline.md)、[XLSX 投影](feature-evidence/SHEET-27-outline-office-io.md)，真实文件与界面未验收 |
| SHEET-28（Pro） | 公开打印/PDF 规格；商业菜单 ID 未核实 | 当前没有完整 Sheets 页面设置/分页预览/输出工作流 | 打印区域、适页、重复标题、页眉页脚、图形和隐藏处理、实际输出待实现/验收 |
| SHEET-29 | 本项目 Office parser/writer 与公开导入导出规格 | 应用壳负责打开/保存/导出真实 XLSX；条件格式来源、dxf/theme、base/linked x14 与整本模型/输出资源预算接入本地读写链路 | [条件格式 Office 专项](feature-evidence/SHEET-14-conditional-format-office-roundtrip.md)的本轮草稿未执行；XLS/XLSX/CSV/TSV 公开契约、保留 XLSB/ODS 读取、全部资源/样式/公式独立重开仍未验收 |
| SHEET-30 | [Sheets schema][sheet-schema]：`AddRangeProtectionFromToolbarCommand` 等保护菜单 | 当前无完整顶部权限流程 | 服务端工作簿/表/范围权限、可执行操作、失败原子性/协作及恢复待验收 |

## 4. Docs 顶部功能逐项映射

[Docs schema][doc-schema] 的开始/插入/视图由各插件注入；[示例装配][docs-mount] 还包含 drawing、hyper-link、thread-comment 和 TOC。正文目录与左侧导航是两个能力；页眉页脚和 TOC 都有专用上下文分类，不能把普通插入按钮当作已覆盖。

| 功能 ID | 上游菜单 / 命令依据 | 本地 item / 真实行为 | 尚缺行为或验证 |
| --- | --- | --- | --- |
| DOC-01 | 文档模型/排版与 [Docs schema][doc-schema] | 传统分页 Canvas、现有段落/文字运行模型；工具栏不以 HTML 作为权威源 | 节、复杂自动排版/跨页对象、所有模式差异与大文档性能未整体验收 |
| DOC-02 字符 | 上游 `SetInlineFormat*Command`、`SetInlineFormatCommand`，包括 subscript / superscript | 字符格式改实际 run 范围；caret 格式保留共享待输入 marks，body/story 成功后受 revision/generation 保护；CJK 字体槽、明确清色、高亮、混合缓存运行与来源清除标记进入服务/写出；成功撤销清旧偏好，取消保留 | [上下标](feature-evidence/DOC-02-vertical-align.md)、[局部字符/清除格式](feature-evidence/DOC-02-clear-formatting.md)和[共享输入](feature-evidence/COMMON-03-input-selection-axis-history.md)记录草稿；系统 IME、全部字符样式/主题/继承、复杂字段/锁与生产 WASM/Office 独立往返仍未验收 |
| DOC-02 段落 | `SetParagraphNamedStyleCommand`、`AlignOperationCommand`、段落设置面板 | 正文/H1–H6、四对齐、缩进/列表、行距；`doc.command.format-paragraph` | 段前/后间距面板正在补充；Title/Subtitle 独立命名样式、混合/全段落语义及 Office 往返未完整验收 |
| DOC-03 | `doc.menu.breaks`；`doc.menu.section-break.{next-page,continuous,next-column,even-page,odd-page}`；`docs.operation.insert-column-break` / `insert-section-break` | 当前普通 `doc.command.insert-page-break`；布局纸张 A4/Letter、纵横方向；边距设置正在补充 | 分节各类型、节级分栏、列断点均未接顶部；普通分页符不能代替分节 |
| DOC-03 页眉页脚 | `doc.command.open-header-footer-panel`；[headerFooter schema][header-footer]：`docs-ui.ribbon.header-footer`、`DocHeaderFooterRibbonControl.{insertion,navigation,options,position}`、`doc.command.close-header-footer` | 当前无已实现本地页眉页脚上下文分类 | segment 切换、节关联、首页/奇偶页、页码/字段、距离、导航和结束编辑全部未验收 |
| DOC-04 左侧大纲 | DocSideMenu/SideMenu；固定模式默认值 | `doc.operation.toggle-outline`；真实标题/层级/activeId、跳转、停靠/浮动、导航键盘、传统默认关闭/现代默认开启 | 宽表与浮动大纲实际裁剪交集尚未接；正文标题折叠、表格内标题和 Title/Subtitle 保真未验收 |
| DOC-04 正文 TOC | [TOC schema][toc-schema]：`doc.menu.insert-table-of-contents`；`doc.operation.open-table-of-contents-dialog`；`ribbon.table-of-contents` 上下文 | 当前没有正文 TOC 插入、更新或上下文工具 | 不能以左侧大纲替代正文目录；目录内容/页码更新、删除、保存/Office 输出待实现/验收 |
| DOC-05（含 Pro） | `doc.command.bullet-list` / `order-list` / `check-list` / `change-list-nesting-level`，`CheckListCommand` 属于 START/LAYOUT | 本地 bullet/numbered 和缩进可执行；本轮将查找映射到 OTHERS 语义组、段落/列表映射到 LAYOUT；复选列表由 `doc.command.set-block-type` 创建、独立 `doc.command.toggle-checklist-item` 勾选，Canvas 点击和 Enter 共享文档历史；DOCX 使用原生编号符号 | 清单切片证据见 [DOC-05](feature-evidence/DOC-05-checklist.md)；嵌套/升降级、编号重启/连续、增强样式、重排、富剪贴板及真实 Office 应用重开尚未验收 |
| DOC-06（含 Pro） | 表菜单 `doc.menu.table`；上下文 `DocTableInsertRowAbove/Bellow/ColumnLeft/RightCommand`、`doc.table.delete-{rows,columns,table}` | 实际插表；表内显示本项目“表格工具”上下文，前后插行/左右插列/删行列/删表，更新控件锚点及一次文档历史，Tab 末格追加复用命令 | 不冒称上游存在同名同形 tab；合并拆分/移动/尺寸边框/整表对齐/跨页嵌套、公开 Facade 和 Office 往返仍未完整覆盖 |
| DOC-07 | 固定 docs-drawing-ui image schema、shared drawing align 与 Docs 模型/布局 | 图片/inline/crop/替换已接真实模型；多选、批量拖动/均匀缩放/旋转、六种对齐/两种分布/排列/复制剪切/批删、Ant 浮条/右键及四个 FDocument wrapper 进入同一文档事务和真实布局 provider | [图片及 inline](feature-evidence/DOC-07-image-editing.md)、[批量编辑与 57 条新草稿](feature-evidence/DOC-07-image-editing.md#image-batch)均未运行；跨页 anchor、混合对象/非均匀 affine、Office、原生 IME/主题/两后端与大样本未验收 |
| DOC-08 | docs-hyper-link-ui：`ShowDocHyperLinkEditPopupOperation` | `doc.operation.open-link` 可插入链接并保留范围 | 完整编辑/删除、书签/内部目标、位置变换和 Office 往返未验收 |
| DOC-09 | docs-find-replace 与共享 find-replace schema | `doc.operation.toggle-find`，正文匹配/定位当前实现 | 完整替换、选项、表格/页眉页脚等作用域、可撤销全部替换未验收 |
| DOC-10 | docs-thread-comment-ui 浮动工具/绘图上下文 | 当前无完整本地顶部/浮动持久评论流程 | 锚点变换、回复/解决、用户/权限/服务同步未验收 |
| DOC-11（Pro） | 内容控件公开规格；完整商业菜单 ID 未核实 | 内容控件 split 与六种实际类型直达，真实面板/命令/历史；已验证锁、段落拆分锚点、撤销/重做与内部重开 | 图片/分组/重复节/文档部件完整 UI、跨段 SDT、全部编辑变换及 DOCX 往返仍未验收 |
| DOC-12（Pro） | 固定 `docs-reference-ui` 引用标签，脚注和尾注各有插入、上一条、下一条、设置 | 传统分页下 8 个固定根和 1 个本地管理根已接实际命令与编辑宿主，返回引用另有上下文入口 | [脚注和尾注](feature-evidence/DOC-12-footnotes-endnotes.md)记录真实模型、权限与分页边界，DOCX 独立重开、主题、焦点和性能未验收 |
| DOC-13（Pro） | 固定 docs-code / -ui 发布包，`docs-code.command.insert/update` 与语言目录 | INSERT/MEDIA 1/7；21 项真实 Prism 词法、整块字体、语言/More 四字段、源编辑/滚动/caret、桌面 Popover/移动 Drawer 与结构转换进入共享内容命令/Facade | [结构](feature-evidence/DOC-13-14-16-17-structured-blocks.md)、[Canvas/host](feature-evidence/DOC-13-14-16-17-canvas-interactions.md)、[词法/字体/段落入口与 56 个新草稿](feature-evidence/DOC-13-14-16-17-code-popmenus.md)均未运行；子选择不另计根，复杂语言/字体、复制、Office 与性能未验收 |
| DOC-14（Pro） | 固定 docs-column / -ui 发布包，paragraph context schema 与公开组/栏 Facade | 2–5 栏、比例/间距、增删移动/逐栏文本；Canvas 调宽 preview→释放一次命令，px/公开单位来源、段落及表后入口进入真实结构历史 | [结构](feature-evidence/DOC-13-14-16-17-structured-blocks.md)、[Canvas/单位](feature-evidence/DOC-13-14-16-17-canvas-interactions.md)、[段落/表后](feature-evidence/DOC-13-14-16-17-code-popmenus.md)保留未验收；本地常规根/分组差异、全部模式/故事/嵌套和 Office 仍缺 |
| DOC-15（Pro） | 固定 docs-latex/-ui 发布类型和 ES；四个 `docs-latex.command.*-formula`、Insert/Media、右键与上下文 schema | 常驻插入、三右键动作和 Canvas 双击接同一稳定 atom/命令；71 模板/11 类与 formula 上下文 11 根接同一 source planner、operation 和 unit editor owner；Ant source、公开范围/segment aliases、named DOC_LATEX_PLUGIN、历史与内部重开已有接线；公开仅 inline，本地 block 是扩展 | [LaTeX 编辑与 43 个未运行草稿](feature-evidence/DOC-15-latex-editing.md)、[模板/上下文与 47 条新声明（计划 117 用例）](feature-evidence/DOC-15-latex-templates-ribbon.md)；其他段落/浮动/移动入口、完整 TeX/OMML、source 内部多样式、Worker/Rust/WASM、真实 IME/主题/两后端和性能仍有缺口 |
| DOC-16（Pro） | 固定 docs-quote / -ui 发布包，`docs-quote.command.insert`、keepLines/keepNext/widowControl 与 context 挂载 | 真实子段落/属性/退出/unwrap/remove；Canvas/辅助布局及保 ID/内容转换；四项段落分页、Quote 自动 keep/作者恢复、能容纳时整体移页与超高拆页、Ant 设置和共享 Facade 命令 | [结构](feature-evidence/DOC-13-14-16-17-structured-blocks.md)、[Canvas](feature-evidence/DOC-13-14-16-17-canvas-interactions.md)、[段落入口/转换](feature-evidence/DOC-13-14-16-17-code-popmenus.md)、[分页与 37 个未运行草稿](feature-evidence/DOC-16-quote-pagination.md)均未验收；复杂对象/表格分页、常规分组差异、复制/Office 与性能仍缺 |
| DOC-17（Pro） | 固定 docs-callout / -ui 发布包，`docs-callout.command.insert` 及 context 挂载 | 图标/颜色/容器和真实子段落；Canvas/辅助 host、段落/表后/移动入口、可恢复配置的结构转换进入内容历史/Facade | [结构](feature-evidence/DOC-13-14-16-17-structured-blocks.md)、[Canvas](feature-evidence/DOC-13-14-16-17-canvas-interactions.md)、[完整结构入口](feature-evidence/DOC-13-14-16-17-code-popmenus.md)未运行；全部菜单组合/触控、Office、分组差异、历史传输及性能仍缺 |
| DOC-18（含 Pro） | drawing 基础：`doc.command.menu-insert-shape` 及矩形/椭圆命令；完整商业/SmartArt 菜单 ID 未核实 | 当前无已连通形状/SmartArt 顶部流程 | 形状文字/锚点/层级/连接器/完整目录/SmartArt 布局和 Office 映射未验收 |
| DOC-19（Pro） | 共享/Docs 图表公开规格；商业菜单 ID 未核实 | 当前没有 Docs 图表顶部/属性面板 | 共享可编辑图表模型、内嵌数据、环绕/分页、属性与 Office 关系待实现/验收 |
| DOC-20 | watermark 插件与公开规格 | 当前无已验收水印顶部流程 | 文本/图片、透明度/定位/页范围、打印与 Office 输出未验收 |
| DOC-21（Pro） | 打印/PDF 公开规格；商业菜单 ID 未核实 | `doc.command.print` 当前通过 `documentHtml` 单 section 与 CSS 打印文档内容；不是应用壳截图，也不是 Canvas 逐页布局输出 | 真实分页布局与页范围、字体/对象/页眉页脚/脚注一致性、打印/PDF 独立文件仍未完成，不能记为 Pro 打印已达标 |
| DOC-22 | 本项目 duke-pages parser / DOCX writer | 应用壳打开/保存/导出真实 DOCX；DOC/ODT 从真实全局定义取默认格式，原样式资源经专用 legacy-sources part 保存；DOCX 当前 Word 资源与历史来源分开；清除后缓存的局部混合格式接入 writer | [默认样式/历史来源/字段缓存证据](feature-evidence/DOC-02-clear-formatting.md)记录未运行内部/原生两次重开草稿；应用专用 part 在独立 Office 另存中的保留、原格式全部语义/关系及完整 DOC/DOCX/ODT 保真仍未验收 |
| DOC-23 | DocFooter/DocStatistics/DocLayoutProgress | 组合底栏：当前/总页、字符/选区字素统计、缩放、内容修改状态；不把 UI 状态伪称云保存成功 | 全部上游统计口径、真实布局任务进度、大文档与模式差异未完整验收 |

公开基础插入项的符号/emoji 选择器现已接入，支持保留选区插入、分类/最近使用与自定义 Unicode。Docs 复制/剪切可写文本与 HTML，功能区、右键及原生粘贴共用保留源格式、匹配目标格式和纯文本逻辑。最近一次模式切换、系统文件型图片及跨结构资源保真仍待补齐。见 [Docs 剪贴板静态证据](feature-evidence/COMMON-03-docs-clipboard.md)。

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
