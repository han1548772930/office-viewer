# Docs 与 Sheets 顶部功能区对齐清单

审阅日期：2026-11-21  
静态入口复算快照：2026-11-22；2026-11-27 按当前 TypeScript 声明复算 Docs/Sheets 常规根，并与 HEAD 的同口径结果比较，未执行产品、factory 或商业 SDK。
固定 Univer 版本：`eaf4cfa04fabdbcfc922c692f3450459f8a29546`
Microsoft 资料审阅日期：2026-11-21，主要以 Microsoft 365 桌面版公开说明为准。

本清单先核对固定版 Univer 的顶部入口、插件扩展和上下文菜单，再核对本项目的入口、网格布局和命令处理器。Microsoft Word / Excel 的公开功能用于发现可补充能力，不代替 Univer 固定版入口对照。功能 ID 的验收状态仍以[完整功能对齐表](../UNIVER_FEATURE_PARITY.md)为准。

## 统计口径

- 一个 Ribbon 根入口是合并菜单 schema 后位于某个标签和分组中的顶层项。split、menu、select 的子项不增加根入口计数，动态菜单按其根菜单计一次。Docs 的 10×10 表格尺寸格是 Table 的子菜单，不另算 100 个按钮。
- 插件 schema 注入的常驻 Ribbon 项计入固定版 tab 根入口。依选择状态才出现的 contextual tab、右键菜单、浮动对象工具栏、工作表标签菜单和页脚菜单分开统计。
- 本地数值统计 `createSheetRibbon` / `createDocumentRibbon` 的常规根声明，包含条件隐藏项；Sheets 动态函数分类按当前 catalog 的 9 类静态展开。嵌入组件缺省 classic，示例缺省 grid，两种变体共享 schema。额外入口仍计入本地数量，但不因此认为与上游功能等价；整组进入 More 不增加入口数。
- Univer 商业 Pro 插件实现源码不在固定检出中。Pro 清单依据项目的公开功能规格和已有 feature ID，不把推测的闭源按钮登记为上游源码事实。

## Univer 固定入口统计

### 全量功能 ID 统计（含 Pro）

功能 ID 的统计范围取自[完整功能与 Pro 对齐清单](../UNIVER_FEATURE_PARITY.md)，不等同于 Ribbon 按钮数。当前共有 **65 项**：共享能力 12 项、Sheets 30 项、Docs 23 项。三个域均纳入统计，表内的 Pro/Server 项和增强目标不从总量中剔除。

| 范围 | ID 数 | 覆盖范围 | 其中明确标为 Pro / Server / Pro 增强目标 |
| --- | ---: | --- | ---: |
| Shared | 12 | COMMON-01–12：工作台、命令、剪贴板、i18n/主题、权限、协作、历史、对比、嵌入、渲染计算、媒体 | 3（COMMON-07、08、10） |
| Sheets | 30 | SHEET-01–30：工作簿/网格、格式、公式、数据工具、对象、图表、透视分析、打印、Office I/O、保护 | 10（SHEET-10、21–29） |
| Docs | 23 | DOC-01–23：排版、格式、节/目录/列表/表格、媒体、评论、控件、脚注、扩展块、图表、水印、打印、Office I/O、统计 | 11（DOC-11–19、21–22），另有 DOC-05/06 的 Pro 增强目标 |
| 合计 | **65** | 全部状态仍按功能对齐总表逐项验收 | **24 个明确 Pro/Server ID，另有 2 个混合基础与增强项** |

此处的“明确 Pro/Server”按总表中明确写出的标记统计，不表示其上游每个子能力都仅存在于商业发行，也不表示已完成。SHEET-29、DOC-22 是 Pro/Server 对齐目标，仍要求使用本项目自己的读写链路。完整 ID 名称、规格和验收状态以总表为准。

### Docs

v1.6.0 OSS 常规 Ribbon 静态根声明为 **38 项**，来自 `scripts/audit-univer-oss-ribbon.mjs`。2026-11-27 本地三个主标签为 **54 项（35/16/3）**，包含条件声明与原位插件，子项和上下文分别记录；HEAD 的同口径结果为 53 项，本批新增共享快捷键面板根。传统模式另有「引用」9 根，其中 8 根来自固定 Pro 契约，1 根为本地管理。Pro v1.6.0 缓存注册及未解析入口见 [Pro Ribbon 统计](../UNIVER_TOOLBAR_PARITY.md#根入口统计与布局)。

| 标签 | 固定版分组和入口数 | 本地当前根项 | 对齐状态 |
| --- | --- | ---: | --- |
| Home / 开始 | HISTORY 5、FORMAT 11、LAYOUT 5、OTHERS 4，共 25 | 35 | 本地 history 5、format 11、layout 5、others 4，另有段落工具 5、其他工具 5。搜索与快捷键面板位于 others 的 r2c2 / r1c2；搜索读取现有右键动作并转发完整命令与参数，实例入口绑定既有 workbenchUiHost。运行验证后置。 |
| Insert / 插入 | MEDIA 7、CELL 1、OTHERS 2，共 10 | 16 | 本地 media 9、cell 2、structured 3、extensions 2；代码块、引用/Callout/多栏、目录及内容控件已有接线，全部 Pro 仍按对应包追踪；结构块差异见[证据](DOC-13-14-16-17-structured-blocks.md)。 |
| View / 视图 | DISPLAY 3，共 3 | 3 | 页面设置、缩放、全屏三个根项；大纲、专注和打印现在位于 Home 扩展组。数量相同不代表行为已验收。 |
| 主标签合计 | **38** | **54** | 条件显隐、Pro 替换、独立浮动和上下文入口继续按路径核对。 |

固定版 Insert 的 MEDIA 网格由多个插件共同注册。Breaks 在 `r1c1` 跨两行，Table 在 `r1c2`，Emoji 在 `r1c3`，Symbol 在 `r2c3`，Image 在 `r1c4`，Hyperlink 在 `r1c5`，Horizontal Line 在 `r2c5`。本地之前没有显式 `gridLayout`，列优先 fallback 导致 Image 掉到第二行；并且 Breaks、Horizontal Line、Symbol、Emoji 被放入 OTHERS。现已按上述位置修正，定向 Ribbon 契约测试 20/20 通过。页面布局统一浏览器验收尚未执行。

固定版另有目录上下文 Ribbon 两项，仅选中目录时显示；页眉页脚上下文为 5 个控件，合计 7 个 OSS 上下文根声明。右键与浮动工具不合并进常规 38 项。目录是独立 opt-in 插件，插入位置为 Insert > CELL，更新和删除位于目录 contextual tab 与主区域 Layout 右键菜单。本地目录入口、模型/命令、配置和上下文控制器已接入；仍按 [DOC-04 证据](DOC-04-table-of-contents.md) 保留未验收。Pro 内容控件、形状、增强表格等上下文另列，不受 OSS 七项数量限制。

代码块在 OSS 固定源码中有模型和剪贴板支持；v1.6.0 Pro 发布包已静态识别 `docs-code-ui`、`docs-latex-ui`、`docs-sdt-ui` 与 `docs-reference-ui` 的根声明。它们的模式、状态与命令闭环须按对应包及公开规格继续核对，不能把静态注册当成最终效果。段落设置和节设置是段落右键 Layout 组入口，不属于 Ribbon 项。

### Sheets

v1.6.0 OSS 常规 Ribbon 静态根声明为 **63 项**，2026-11-27 当前本地五个主标签为 **71 项（32/9/12/12/6）**，含范围主题、图片与背景、数据条、图标集及 Pro 扩展。HEAD 的同口径结果为 69 项；快捷键面板批增加 1 根至 70，本次 NA 批再增加“信息”1 根至 71。此前错误的 71 项合计已在快捷键面板批纠正，本次数量依据新接入的真实分类根，不能沿用旧合计。子项与上下文分别计数，More 复用原入口。完整逐组、上下文、处理器和 Pro 边界见 [Sheets 源码入口清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)、[图片入口](SHEET-19-images.md)及 [Pro Ribbon 统计](../UNIVER_TOOLBAR_PARITY.md#根入口统计与布局)。

| 标签 | 固定版根项 | 本地当前根项 | 主要差异 |
| --- | ---: | ---: | --- |
| START / 开始 | 30 | 32 | 本地 clipboard 5、font 11、layout 6、number 5、tools 5。搜索功能与快捷键面板位于工具组 r2c2 / r1c2，范围主题在 r1c4 跨两行。查找单元格保留在 Data；面板读取 Ribbon 元数据和当前编辑器的真实快捷键。 |
| INSERT / 插入 | 9 | 9 | 本地 media 4、cells 5。图片、超链接、图表、迷你图走媒体组，三项快速验证和数据条/图标集走单元格组。线程评论和全部动态行为仍待核对。 |
| FORMULAS / 公式 | 12 | 12 | 固定版是 11 类函数菜单与 1 个计算选项；本地为 9 个 catalog 根分类、计算模式/编辑公式/重算 3 项。[参数表单、NA 信息根与选区禁用](SHEET-09-formulas-entry-review.md)已接；根数相同仍缺工程、数据库及完整函数目录，自动范围/批量插入和迭代设置仍未齐。 |
| DATA / 数据 | 6 | 12 | 本地 organization 5（排序、筛选和 3 个大纲根）、rules 2、pivot 4、data-tools 1。按实际 items 数组及 outlineItems 展开复算，纠正旧合计。 |
| VIEW / 视图 | 6 | 6 | 本地 display 4、visibility 2；背景根在冻结后、order 0.5、r1c2 跨两行，设置/删除两子项接实际模型；十字高亮属于固定 OSS 入口。背景历史/资源重开、两后端和所有动态状态仍未验收。 |
| 合计 | **63** | **71** | 本地包含扩展项，常规数量、上下文数量与功能完成度分别记录。 |

Sheets 固定源码的组顺序为 START HISTORY/FORMAT/LAYOUT/NUMBER/OTHERS；INSERT EDIT/MEDIA/CELL/OTHERS；FORMULAS BASIC/OTHERS；DATA FORMULAS/RULES/ORGANIZATION/ANALYSIS/OTHERS；VIEW DISPLAY/VISIBILITY。零根项的组不进入数量；根声明按菜单 schema 的产品归属静态投影。固定 desktop 装配与 factory 决定最终显隐，尚未执行。本地公式按当前 118 项 catalog 的 9 类静态展开；没有执行求值引擎。

2026-11-26–27 顶部复核另补 [Sheets 快捷条件格式](SHEET-14-quick-formatting.md)、[Sheets 上下文搜索](COMMON-02-sheet-context-feature-search.md)、[Docs 命名段落样式](DOC-named-paragraph-styles.md)与两产品快捷键面板接线。搜索的 `contextCommand` 现在转发 key、commandId 和完整 params；Sheets Canvas 与右键共用实际上下文，Docs 保留结构/字段目标及既有可用性判断。新增的快捷键面板根为每产品 1 个；条件格式子项、搜索结果及溢出副本均不新增根数。

样式复核覆盖共享按钮/选择器、两产品颜色控件、数字格式、地区货币字形、More、文件/设置按钮及页眉页脚。共享尺寸、禁用 Select hover 和大 split 布局差异见[顶部控件记录](../UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing)，专属证据见 [Docs 页眉页脚](DOC-03-header-footer-ribbon-sizing.md)、[Sheets 自定义弹层](SHEET-RIBBON-custom-popup-sizing.md)和[地区货币](SHEET-06-regional-currency.md#number-group-review-2026-11-27)。图标、箭头、长标题和内层 Ant 控件纳入真实矩形与 hover 回归草稿；全部运行验证后置。

2026-11-27 续核已修正 Grid 靠左主按钮误作用于箭头、Classic 的 24px 规则压过 Grid 16px 规则，以及 More 箭头宽度/flex basis；两产品各 45 个预设按 Ant 透明边框计入 padding，窄弹层色块等比收缩。页眉页脚导航按下保留原焦点；公式栏展开箭头填满实际栏高，数字格式 classic/grid 字号分别为 14/12px。Sheets 边框颜色/线型使用同一 small 高度，快捷下拉动作和格式菜单约束长标签；Docs 代码块浮动工具栏限制长字体/语言按钮与右侧边界，并用 Ant Tooltip 显示完整值。Data 单行 split、颜色预设、导航焦点、公式栏按钮边界、格式字号及 Code 长值已补未运行草稿；本批新增常规/扩展/溢出/上下文根均为 0。

前轮语法、引用、文档链接/锚点和 diff 空白检查已完成，具体范围与结果见[顶部控件静态收尾](../UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing)。这些检查不执行产品或浏览器；所有运行验证仍后置，入口数量和功能验收状态不变。

本次续核修正 classic split 单侧边框的 0.5px 箭头中心偏移，Grid/More 保留原宽度；[符号与表情](DOC-03-header-footer-ribbon-sizing.md#doc-special-character)接固定公共命令、完整 value 与事件，空格/NBSP 有效。[筛选重应用](SHEET-11-12-sort-filter-entry-review.md)复用 operation 保留内容历史，[大纲详情与设置](SHEET-27-outline.md)读取真实候选/权限，Mac 快捷键补回 Shift。两个特殊字符根、大纲三个根及其子项沿用现有位置，本次常规/扩展/溢出/上下文根增量均为 0，实际 hover、焦点、主题和编辑闭环仍待统一验收。

本次第一轮统一静态收尾通过：84 个 TS/JS、25 个 Vue、26 个 CSS 文件/样式块、13 段夹具脚本、1378 处本地引用；30 份 Markdown 的 907 条链接、66 处锚点，以及 diff 与 23 个新增文本文件的空白检查。检查没有执行产品或运行验证，完整 Sheets / Docs / Pro 仍按专项证据保留未验收。

随后专有控件追加统一两产品颜色箭头的边框/分隔线和禁用文字色，Sheets 快捷下拉五个按钮与 Docs 列表预设/编号表单接收原有 disabled；两产品颜色预设迁移到 Ant Tooltip，支持 hover 和 focus。颜色按钮还受库中对称透明边框影响，具体偏移量保留待实测。以上追加小修不属于前述静态快照，浏览器运行验证仍后置，根数和功能验收状态不变。

Docs 表格覆盖层的禁用 hover/focus 已同步修正；[完整 Picker 大改方案](DOC-03-header-footer-ribbon-sizing.md#doc-character-picker-plan)尚未实施，生成目录数量未知。方案、目录项数和根数均不替代完整功能闭环验收。

专有控件最终静态收尾通过：86 个 TS/JS、27 个 Vue、28 个 CSS/样式块、15 段夹具脚本、1420 处本地引用；30 份 Markdown 的 926 条链接、71 处锚点，以及 diff 与 23 个新增文本文件尾空白。检查覆盖两子任务稳定后的本批小修和主线 Tooltip/草稿补充，没有运行产品或功能验收。

随后 NA/分隔符批补齐[信息分类与 RIGHT 零长度回归](SHEET-09-formulas-entry-review.md)、[Docs 分隔符范围状态及辅助宿主执行守卫](DOC-03-section-breaks.md#doc-view-break-scope)，并将数字格式按钮的零边框、打开态背景、禁用文字和 Tooltip hover/focus 接到原有控件。当前常规根为 Sheets 71、Docs 54；本批仅 Sheets 常驻根 +1、函数子项 +1，其他入口口径增量为 0。上述历史静态快照不覆盖本批；[自动引用/批量插入方案](../UNIVER_TOOLBAR_PARITY.md#formula-insertion-plan)仍待确认，全部运行验证后置。

专有控件续修补齐边框/切片器提示的 focus 触发和 Docs 列表预览的固定标记列，未增加入口或用例声明；原浏览器草稿追加真实提示、长标记裁切与 hover 边界检查。切片器主 gallery 与设置字段 popup 的列数/尺寸/预览差异已分开记录，保留未对齐与未核实项。当前收尾累计八份生产文件、八份既有草稿，静态与运行结果单独记录，当前根数不变。

后续批复用真实样式投影补齐[切片器/时间线预览和默认选中](SHEET-RIBBON-custom-popup-sizing.md#slicer-style-preview-continuation)，保持 14/12 内置目录、作者样式及额外默认重置动作；样式块改为 72px/4px padding，本地主 gallery 三列与上游七列继续单列差异。[快捷下拉预设](SHEET-13-quick-dropdown-date.md#快捷预设尺寸复核)与[页眉页脚 Form row](DOC-03-header-footer-ribbon-sizing.md#页眉页脚面板与插表对话框续核)也已小修，后者仍无 Modal 生产打开路径。四份生产文件、两份既有草稿均无常驻、扩展、溢出或上下文根增量，Sheets 71 / Docs 54 保持不变；全部运行验证后置。

本轮最终静态检查通过：16 个 TS/JS、5 个 Vue 模板、5 个样式块、3 段内嵌脚本、179 处本地引用；7 份证据/清单的 538 条链接、71 处锚点及 23 个文本文件空白/冲突标记检查通过。该结果只覆盖本轮限定文件，不执行产品或运行验收，当前所有功能仍未验收。

固定版还有 MAIN_AREA、COL_HEADER、ROW_HEADER、FOOTER_TABS、绘图对象上下文菜单。脚本记录 Sheets 130 个右键/浮动层级节点，包含父子层级，不当作根数。本地行列结构/显隐右键传 sheetId/ranges，尺寸入口传命中 sheetId/index；标签复制、颜色与保护亦有实现切片，关联资源、veryHidden 和全部权限/保存闭环仍未验收。具体来源和差异见 Sheets 清点文档。

填充柄外观使用 1 CSS px 选区/预览边框，2026-11-21 历史描边测试覆盖 0.75、1 和 1.5 缩放，不覆盖后续修改。2026-11-25 已补共享 AutoFill/Refill/CopyDown/CopyRight/AutoClear、FRange 和 Ctrl/Cmd+D/R；Canvas 使用 antdv-next 四模式选项，拖动和双击共用推断默认、单向范围和内缩清尾。序列新增数值线性/等比、文本数字/补零、中文数字/星期、英文星期/月份及其他循环；筛选行、合并、CF/DV、权限与验证进入同一候选事务。重填修订同一历史，等价内容只更新选区/菜单；边缘自动滚动沿用现有调度。见 [SHEET-02 自动填充证据](SHEET-02-auto-fill.md)。本批草稿未执行，standalone 重填 receipt、超过同步预算的 Worker、完整日期/公式边界、真实 IME/冻结/主题/滚动和 Office 往返仍未验收。

2026-11-26 [局部移位事务](SHEET-02-cell-shift-transactions.md)复用四核心+四确认命令、两个父项/四个移位子项，右键子项顺序按固定 schema；新增常规/扩展/溢出/独立上下文根均为 0。确认、Facade Dimension、完整候选/资源、权限/验证、内容与选择历史、宿主重入保护已落盘，56 个草稿未运行；完整 drawing/media、全部资源组合、Office、IME/两后端和性能仍有缺口。

同日[共享绘图工作台](SHEET-21-drawing-workbench-affine.md#entry-inventory)接组合/排列/对齐三个右键父项（2/4/8 叶）及本地删除，Canvas 浮条四个控件复用同一命令。它们属于上下文展示位置，本批常规/顶部扩展/独立 More 根增量均为 0。八个对齐 operation ID 已按固定 `sheet.operation` 前缀修正；后续 `sheet.command.move-drawing` 方向键/API 入口、实际尺寸边界与 1000ms 连续历史不新增 Ribbon 根。主任务当前 82 条、Sheets 图片批 66 条及[Docs 图片批](DOC-07-image-editing.md#image-batch)57 条草稿均未运行，[slicer/timeline affine 与 Ant 控件](SHEET-25-slicer-timeline-affine.md)已补入独立 98/182 未运行草稿；[Pro 多目标连接](SHEET-25-slicer-connections.md)另批已落盘未验收。完整 shape 作者、交叉筛选成员状态/Pro API、组 clipboard、Office 与性能仍缺。

## 顶部入口对应功能闭环

逐项补齐以下链路后才可更新功能状态：入口可见/禁用规则 → 共用 command → model/service mutation → 权限检查 → 撤销/重做 → 保存/重开与文件互操作 → 错误反馈。仅有按钮、面板或 command handler 不算完成。

### Docs

- [ ] DOC-01 文档模型、渲染与分页布局。
- [ ] DOC-02 字符格式、段落格式和混合选区状态；[段距/缩进、当前宿主首段状态](DOC-02-paragraph-spacing-indent.md)、[分隔线真实命令/边框/Canvas](DOC-02-horizontal-line.md)及[特殊字符公共命令/整值/事件](DOC-03-header-footer-ribbon-sizing.md#doc-special-character)已落盘，草稿未运行；六个上下文子项挂载、结构 gap、非空段落下边框、旧 flat story 和完整 Picker 仍缺。
- [ ] DOC-03 分页、分栏、分节及页眉页脚。
- [ ] DOC-04 大纲、标题折叠和正文目录；目录入口、上下文更新/删除、跳转、页码与前导符完整闭环。
- [ ] DOC-05 列表和增强列表。
- [ ] DOC-06 表格、表格上下文命令、尺寸、合并拆分和样式；[顶部 create-table 操作、3×5 尺寸弹窗、公开内容命令及默认整页正文宽](DOC-06-enhanced-tables.md)已接既有插表事务，Studio 操作注册转发、结构 gap、完整 story 插表和 headless 连续分节物理页差异仍缺，草稿未运行。
- [ ] DOC-07 图片、绘图、锚点及文字环绕；[单图/inline 与多选、六对齐/两分布/均匀变换/批删、四 Facade 和 57 条新草稿](DOC-07-image-editing.md#image-batch)，未验收。跨页 anchor、混合对象/非均匀 affine、Office 和真实 UI/性能仍缺。
- [ ] DOC-08 超链接和书签；[链接领域/顶部/快捷键、生成引用链接限制、公开 operation、Canvas 来源 View 与草稿](DOC-08-document-hyperlinks.md)已落盘，未运行。链接信息浮层、浮动/右键挂载、跨文档注册转发/激活、页面定位与独立书签仍缺。
- [ ] DOC-09 查找替换及范围/选区作用域。
- [ ] DOC-10 线程评论入口、锚点、回复/编辑/解决、权限和持久同步。
- [ ] DOC-11 内容控件及锁定/重复结构。
- [ ] DOC-12 脚注和尾注的编辑、编号、分页与 DOCX 映射。
- [ ] DOC-13 代码块；[结构](DOC-13-14-16-17-structured-blocks.md)、[Canvas/host](DOC-13-14-16-17-canvas-interactions.md)、[21 项词法/字体/More/桌面移动入口](DOC-13-14-16-17-code-popmenus.md)，未验收。
- [ ] DOC-14 多栏组与节级分栏；[组/栏](DOC-13-14-16-17-structured-blocks.md)、[Canvas 调宽/单位](DOC-13-14-16-17-canvas-interactions.md)、[段落/表后入口](DOC-13-14-16-17-code-popmenus.md)，未验收。
- [ ] DOC-15 LaTeX 公式；[四公开命令、稳定 source 范围/资源、Canvas/Ant 编辑与 43 个未运行草稿](DOC-15-latex-editing.md)、[71 模板/11 类、1 个上下文 tab/11 根与 47 条新声明（计划 117 用例）](DOC-15-latex-templates-ribbon.md)。常规根增量 0；公开 inline/本地 block 区分；其他段落/浮动/移动入口、完整 TeX/OMML、Worker 和真实 UI/性能仍未验收。
- [ ] DOC-16 引用块；[容器/来源](DOC-13-14-16-17-structured-blocks.md)、[Canvas](DOC-13-14-16-17-canvas-interactions.md)、[入口/保内容转换](DOC-13-14-16-17-code-popmenus.md)、[四项段落分页/Quote keep 与作者恢复、37 个草稿](DOC-16-quote-pagination.md)，未验收。
- [ ] DOC-17 Callout 标注块；[属性/历史](DOC-13-14-16-17-structured-blocks.md)、[Canvas](DOC-13-14-16-17-canvas-interactions.md)、[段落/表后/移动入口](DOC-13-14-16-17-code-popmenus.md)，未验收。
- [ ] DOC-18 形状、连接器和 SmartArt。
- [ ] DOC-19 文档图表。
- [ ] DOC-20 水印。
- [ ] DOC-21 打印和 PDF。
- [ ] DOC-22 DOC/DOCX 导入、真实写出与独立重开。
- [ ] DOC-23 文档统计、页信息及真实布局进度。

### Sheets

- [ ] SHEET-01 工作簿和工作表生命周期。
- [ ] SHEET-02 范围、选区、填充和网格编辑；[四方向共享移位/确认、资源与 56 个未运行草稿](SHEET-02-cell-shift-transactions.md)，未验收。
- [ ] SHEET-03 行列结构、尺寸、隐藏/显示和自动适应；[公开 h/ah/ia/w/hd 与来源变换、17 个草稿](SHEET-03-public-axis-import.md)，未验收。
- [ ] SHEET-04 样式、默认样式、范围主题及混合态；[富文本/Canvas](SHEET-04-05-rich-text-font-editing.md)、[默认来源/引用/reset/空白复制填充与 58 个草稿](SHEET-04-default-style-inheritance.md)、[短字段读回与 13 个草稿](SHEET-04-public-style-readback.md)、[32 内置/九区段/共享命令/Ant 弹窗与 46 个草稿](SHEET-04-range-themes.md)、[九区段 Ant 表单/原子保存应用与 45 个新草稿](SHEET-04-range-theme-editor.md)、[局部移位/剪切与 32 个草稿](SHEET-04-range-theme-structure.md)，未验收。
- [ ] SHEET-05 富文本；[局部格式/typing/API、连续拖选/换行定位与 55 个草稿](SHEET-04-05-rich-text-font-editing.md)，未验收。
- [ ] SHEET-06 数字格式和 1900/1904 日期系统。
- [ ] SHEET-07 冻结、网格线、视口和缩放；[冻结图表作者输入与共享绘图命中](SHEET-21-drawing-workbench-affine.md)，未验收。
- [ ] SHEET-08 剪贴板和选择性粘贴；[图片作者 affine/HTML/PNG 取消与资源](SHEET-19-images.md)，未验收。
- [ ] SHEET-09 公式、函数和定义名称；[固定 12 根与本地 11 根、参数表单/实时禁用及历史重开草稿](SHEET-09-formulas-entry-review.md)，完整目录、自动范围/批量插入和迭代计算仍未完成，草稿未运行。
- [ ] SHEET-10 高级公式（Pro 目标）。
- [ ] SHEET-11 筛选；[条件状态、派生行重置、重应用 operation/主题及两根/八子项证据](SHEET-11-12-sort-filter-entry-review.md)已落盘，单格/单行连续区域扩展待确认，草稿未运行。
- [ ] SHEET-12 排序；[真实活动列、历史/重开草稿及连续区域方案](SHEET-11-12-sort-filter-entry-review.md)已补，扩展和自定义范围选择待确认，未验收。
- [ ] SHEET-13 数据验证；[快捷下拉/日期入口、全部选区一次规则与列表转义](SHEET-13-quick-dropdown-date.md)、[复选框候选校验与 Canvas 勾选缺口](SHEET-13-checkbox-entry-review.md)、[多选单元格 CSV 校验与未完成 UI](SHEET-13-list-multiple-csv.md)已复核，草稿未运行；[Canvas 日期控件方案](SHEET-13-date-picker-entry-plan.md)待确认。
- [ ] SHEET-14 条件格式。
- [ ] SHEET-15 查找替换。
- [ ] SHEET-16 超链接。
- [ ] SHEET-17 批注与线程评论。
- [ ] SHEET-18 结构化表格；[固定 selector 根与 Studio 创建调用缺口](SHEET-18-table-selector-entry-review.md)已核对，创建方案尚未实施。
- [ ] SHEET-19 图片和水印；[图片/背景两常规根、完整矩阵剪贴板、锚点投影与 66 条草稿](SHEET-19-images.md)、[最终 filter/auto-height 与组框生命周期](SHEET-21-drawing-workbench-affine.md#model-lifecycle)，未验收；水印及 Office 仍缺。
- [ ] SHEET-20 十字高亮。
- [ ] SHEET-21 形状及绘图（Pro 目标）；[组合/混合层级与 45 条草稿](SHEET-21-drawing-groups-and-mixed-order.md)、[选区/对齐/移动/Canvas/Facade、作者 affine 与 82 条草稿](SHEET-21-drawing-workbench-affine.md)，未验收；[slicer/timeline affine、五类叶/组与 Ant 控件](SHEET-25-slicer-timeline-affine.md)及 [Pro 多目标连接](SHEET-25-slicer-connections.md)已另批落盘但未运行。完整 shape 作者、交叉筛选成员状态/Pro API 和组 clipboard/Office 仍缺。
- [ ] SHEET-22 可编辑图表（Pro 目标）；[逐型审计](SHEET-22-chart-type-audit.md)、[Relation/词云资源](SHEET-22-relation-word-cloud-resources.md)、[图例](SHEET-22-legend-selection-wrap.md)、[标题配置](SHEET-22-title-subtitle-contract.md)、[标题直接编辑](SHEET-22-title-direct-editing.md)、[绘图区大小/恢复](SHEET-22-plot-area-direct-editing.md)、[层级图 Ant Tooltip](SHEET-22-hierarchy-ant-tooltip.md)、[四个对象上下文根/剪贴板/PNG/删除与 34 个草稿](SHEET-22-chart-object-context-actions.md)、[真实 Drawer 路由/草稿保护与 33 个草稿](SHEET-22-chart-edit-panel-routing.md)、[Line/Smooth/Step 二维/3D 与 34 个新草稿](SHEET-22-line-and-area-styles.md)，均未验收。27 route 分组是导航分类，不是完成数量；本批线型常规根增量为 0。
- [ ] SHEET-23 数据透视表（Pro 目标）；[固定 15 项值显示方式](SHEET-23-24-pivot-show-data-as.md)与 [命令/Facade/linked 图格式](SHEET-23-24-pivot-show-data-as-facade.md)已落盘，枚举不计 Ribbon 根，所有草稿未运行；完整 UI/展开折叠/日期条件/Office/大样本继续闭环。
- [ ] SHEET-24 透视图（Pro 目标）；[owned/linked 创建与一次作者配置](SHEET-24-owned-pivot-charts.md)、[字段引擎](SHEET-24-pivot-chart-engine.md)及 [四字段区/11 汇总/标签筛选 UI](SHEET-24-pivot-chart-ui.md)已落盘；上游两条互斥根、两个创建子项、四个轴子项与动态层级单列，本地常规根增量 0；全部草稿未运行，同时创建表图/展开折叠/全部 Pro API/Office/真实 UI/性能继续闭环。
- [ ] SHEET-25 切片器和时间线（Pro 目标）；[原 affine/Ant 与 98/182 草稿](SHEET-25-slicer-timeline-affine.md)、[共享多目标连接/Facade/历史](SHEET-25-slicer-connections.md)、[四格式根与连接 UI](SHEET-25-slicer-connections-ui.md)、[复制/结构生命周期](SHEET-25-slicer-connections-lifecycle.md)已落盘，常规根增量 0、上下文四根分开统计；全部新草稿未运行，完整交叉筛选成员状态、日期/Pro API、缺依赖跨 workbook、Office 与真实 UI/性能仍缺。
- [ ] SHEET-26 迷你图（Pro 目标）；[一个媒体根、三个上下文挂载、四图型与删除默认范围修复](SHEET-26-sparkline-settings-and-groups.md#media-sparkline-entry-review-2026-11-27)，API/右键/搜索仅删选中成员、显式整组删除保留，全部草稿未运行。
- [ ] SHEET-27 大纲分组（Pro 目标）；[三个 Organization 根、真实候选/权限、无变化状态与 Mac Shift 快捷键](SHEET-27-outline.md)已接线，草稿未运行，文件/界面未验收。
- [ ] SHEET-28 打印和 PDF（Pro 目标）。
- [ ] SHEET-29 Office 导入与导出。
- [ ] SHEET-30 工作簿、工作表和范围保护。

SHEET-02 的早期填充柄预览描边有历史定向测试结果；双击、序列推断、填充选项、重填和视口外拖动已有后续代码及未运行草稿。standalone 重填 receipt、大范围 Worker、完整序列/文件样本与浏览器验收仍未完成，不能因为填充柄可命中/可拖动就关闭该 ID。

## Microsoft Office 可补充功能

Microsoft 的标签和可用项会随桌面/Web、版本、许可、视口和当前选区变化。以下按 Microsoft 365 桌面版常见 Ribbon 结构盘点，File/Backstage、选区上下文标签、溢出菜单与状态栏单列，不和 Univer 常驻根入口计数混合。功能类别是对照范围，不表示微软的每个命令都必须照搬 UI；标为候选的能力先归入功能模型与 feature ID，再接入共享命令体系。

### Word

| Office 标签/区域 | 功能盘点 | 本项目补充清单 |
| --- | --- | --- |
| File / Backstage | 新建/打开/保存副本/打印/共享/导出/账户/选项 | [ ] 补齐文件级入口和导入导出状态；真实保存、打印、导出能力映射 COMMON-05、DOC-21/22。 |
| Home | 剪贴板、字体、段落、样式、编辑/查找替换；Microsoft 365 还可能有 Editor、Dictate | [ ] 核查剪切/复制/粘贴类型、格式刷、清除格式、样式、列表、对齐缩进、Find 的位置和命令；语音输入、拼写/语法检查列为新增候选。 |
| Insert | 页面元素、表格、图片/图标/形状、链接、页眉页脚/页码、文本框/艺术字、公式/符号、评论 | [ ] 逐组对应 DOC-06–08、DOC-10、DOC-15、DOC-18/19；补文本框/图标/艺术字等对象模型候选并核实 DOCX 写回。 |
| Draw | 笔/荧光笔、橡皮擦、套索/墨迹选择、墨迹转形状/数学 | [ ] 新增墨迹对象、笔迹输入、转换与选择行为候选；验证 Canvas 命中、撤销和 DOCX/打印表示。 |
| Design | 文档格式集、主题色/字体、段落间距、页面颜色和边框 | [ ] 增加主题/样式集/配色/字体主题/页面背景与边框的级联模型候选，区别于单个段落直接格式。 |
| Layout | 页边距、方向、纸张、分栏、分页/分节、缩进、段落间距、对象排列 | [ ] 对照 DOC-03/14；补分页符/分节符、节级页面设置、对齐/环绕/层级排列的布局和写出。 |
| References | 目录、脚注/尾注、引文/来源管理、书目、题注、图表目录、索引、法律引证 | [ ] 接通 DOC-04、DOC-12；新增来源库/引用样式/书目、题注/交叉引用、索引功能候选。 |
| Mailings | 收件人/数据源、地址块/问候语/合并字段、规则、预览、完成合并、信封/标签 | [ ] 新增邮件合并工作流候选，覆盖数据源绑定、字段校验、预览、批量输出与打印。 |
| Review | 拼写/语法、同义词/字数、评论/@提及、修订跟踪、接受/拒绝、比较/合并、保护、无障碍检查 | [ ] 完成 DOC-10 评论；增加修订变更模型、逐条审阅、文档比较、校对与 Accessibility Checker 候选。 |
| View | 阅读/打印/草稿/大纲视图、导航窗格、显示标尺/网格、缩放、窗口/并排 | [ ] 核对视图状态、导航/大纲、标尺和缩放；支持多窗口属于桌面能力，先明确本地多文档工作区对应方式。 |
| Help / 搜索 | 帮助、搜索命令/支持内容、反馈 | [ ] 将 Ribbon 命令搜索和帮助入口列为共享工具候选，搜索结果调用已注册命令，不另造执行路径。 |
| Developer（可选） | 内容控件、宏/VBA、加载项、XML 映射、模板 | [ ] DOC-11 覆盖内容控件；宏、XML 映射和模板工具需有权限与文件保真设计后再纳入。 |
| 上下文标签 | 表格布局/设计、图片格式、形状格式、页眉页脚、绘图对象等选中态工具 | [ ] 为表格/图片/形状/页眉页脚建立选区驱动标签、对象命令、键盘焦点和持久化。 |

### Excel

| Office 标签/区域 | 功能盘点 | 本项目补充清单 |
| --- | --- | --- |
| File / Backstage | 新建/打开/保存/另存/导入/导出/打印/共享/账户/选项 | [ ] 补文件级入口和操作状态；SHEET-28/29 的真实打印、文件转换与错误反馈闭环。 |
| Home | 剪贴板、字体/对齐/数字格式、样式、单元格插入删除格式、自动求和、填充/清除/排序筛选/查找；Analyze Data 可能在此出现 | [ ] 核查格式刷、全部清除与仅内容/格式清除、粘贴特殊、填充系列/重填、分析入口分组；Analyze Data 作为分析建议面板候选。 |
| Insert | 表格/PivotTable、推荐/常见图表、迷你图、筛选器、图片/形状/图标、链接、文本框/页眉页脚、加载项 | [ ] 对照 SHEET-16–26；补图片、链接、结构化表格、图标、文本框/页眉页脚入口；清理透视/切片器/时间线重复入口，加入对象上下文工具。 |
| Draw | 笔、荧光笔、橡皮擦、套索、墨迹形状/数学转换 | [ ] 新增单元格批注之外的墨迹/绘图对象、笔迹层、转换和命中候选，映射 SHEET-21 并验证写回。 |
| Page Layout | 主题、页边距、方向、纸张、打印区域/分页符/重复标题、缩放适页、网格线/标题打印、对象排列 | [ ] 增加打印区域、分页符/分页预览联动、重复打印标题、页面缩放、打印网格线和对象排列命令，映射 SHEET-28。 |
| Formulas | 函数库/插入函数、名称管理、公式审核追踪、错误检查、计算选项/重算 | [ ] 补依赖追踪/移除追踪箭头、错误检查/监视窗口、名称管理器；完整函数目录、数组/动态数组求值归 SHEET-09/10 验收。 |
| Data | 获取/转换数据、查询/连接、全部刷新、排序筛选、高级筛选、分列、快速填充、删除重复值、数据验证、预测/假设分析 | [ ] 增加 Power Query 连接/查询/转换/加载/刷新候选、快速填充/分列/删除重复值/假设分析；补齐现有筛选切换/清除/重应用与表选择器。 |
| Review | 拼写、检查、翻译/智能查找、批注与 Note、显示所有评论、保护工作表/工作簿/范围 | [ ] 明确区分 Note 与线程评论；补审阅入口、显示/导航评论、工作簿结构保护和密码保护范围；校对/无障碍检查列为候选。 |
| View | 普通/分页预览/页面布局视图、显示公式栏/网格线/标题、缩放、冻结窗格、拆分/并排/窗口、宏 | [ ] 补分页预览和页面布局视图、标题显隐、拆分窗口/并排、每工作表视图状态和背景图；对应 SHEET-07/19/28。 |
| Automate（许可/平台相关） | Office Scripts 新建/查看/录制/运行/共享，Power Automate 工作流 | [ ] 新增脚本/自动化能力候选，先定义脚本权限、沙箱、API、撤销和共享边界；按可用许可/平台显示。 |
| Help / 搜索 | 帮助、命令搜索、反馈 | [ ] 加入共享命令搜索入口，命令调用复用注册表和权限上下文。 |
| Developer（可选） | VBA 宏、加载项、表单/ActiveX 控件、XML、映射/设计模式 | [ ] 宏和脚本需隔离执行；评估 XML 映射和表单控件模型及 XLSX 保真后再登记子 ID。 |
| 上下文标签 | 表格设计、图表、透视表字段/分析、图片/形状格式、切片器 | [ ] 建立 Table / Chart / Pivot / Drawing 选区驱动的 contextual toolbar；保证目标对象绑定、字段操作和命令状态一致。 |
| 状态栏/快速访问工具栏 | 选区统计、视图快捷切换、缩放、保存/撤销/重做快捷入口 | [ ] 按工作表选择动态计算 Sum/Count/Min/Max 等统计；核对状态栏视图按钮和快速访问命令，不计入 Ribbon 根入口数。 |

官方来源：[Word Ribbon 标签和功能说明](https://support.microsoft.com/en-us/accessibility/word/keyboard-shortcuts-in-word)、[Word 功能区结构与自定义](https://support.microsoft.com/en-us/word/customize-the-ribbon-in-word)、[Word 邮件合并](https://support.microsoft.com/en-us/word/use-mail-merge-for-bulk-email-letters-labels-and-envelopes)、[Word 引文和书目](https://support.microsoft.com/en-us/word/add-or-change-sources-citations-and-bibliographies)、[Word 修订跟踪](https://support.microsoft.com/en-us/word/training/track-changes-in-word)、[Word 文档比较](https://support.microsoft.com/en-us/word/compare-document-differences-using-the-legal-blackline-option)、[Word 口述输入](https://support.microsoft.com/en-us/word/dictate-your-documents-in-word)、[Word 无障碍检查](https://support.microsoft.com/en-us/accessibility/word/make-your-word-documents-accessible-to-everyone-with-accessibility-assistant)。

官方来源：[Excel Ribbon 标签和功能说明](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel)、[Excel 数据导入与 Power Query](https://support.microsoft.com/en-us/excel/power-query-for-excel-help)、[Excel Office Scripts / Automate](https://support.microsoft.com/en-us/excel/office-scripts/introduction-to-office-scripts-in-excel)、[Excel 记录 Office Scripts](https://support.microsoft.com/en-us/excel/office-scripts/record-your-actions-as-office-scripts)、[Excel Analyze Data](https://support.microsoft.com/en-us/excel/analyze-data-in-excel)、[Excel 页面设置](https://support.microsoft.com/en-us/excel/page-setup)、[Excel 批注与 Note](https://support.microsoft.com/en-us/accessibility/excel/use-a-screen-reader-to-work-with-comments-in-excel)、[Excel 工作簿保护](https://support.microsoft.com/en-us/excel/protect-a-workbook)、[Excel 排序与筛选](https://support.microsoft.com/en-gb/excel/sort-data-in-a-range-or-table-in-excel)。

## 后续 Microsoft 365 补充项

以下 Word/Excel 差异在 Univer 与 Pro 功能完整闭环之后单独评估，不计入当前 Univer 功能范围、入口统计或阶段验收。对应候选项以可验证的 Microsoft 365 行为为依据。
