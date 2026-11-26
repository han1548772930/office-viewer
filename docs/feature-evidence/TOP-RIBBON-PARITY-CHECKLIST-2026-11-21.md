# Docs 与 Sheets 顶部功能区对齐清单

审阅日期：2026-11-21  
静态入口复算快照：2026-11-22；2026-11-25–26 人工更新 Docs/Sheets 常规根声明与续作证据，未执行产品、factory 或商业 SDK。
固定 Univer 版本：`eaf4cfa04fabdbcfc922c692f3450459f8a29546`
Microsoft 资料审阅日期：2026-11-21，主要以 Microsoft 365 桌面版公开说明为准。

本清单先核对固定版 Univer 的顶部入口、插件扩展和上下文菜单，再核对本项目的入口、网格布局和命令处理器。Microsoft Word / Excel 的公开功能用于发现可补充能力，不代替 Univer 固定版入口对照。功能 ID 的验收状态仍以[完整功能对齐表](../UNIVER_FEATURE_PARITY.md)为准。

## 统计口径

- 一个 Ribbon 根入口是合并菜单 schema 后位于某个标签和分组中的顶层项。split、menu、select 的子项不增加根入口计数，动态菜单按其根菜单计一次。Docs 的 10×10 表格尺寸格是 Table 的子菜单，不另算 100 个按钮。
- 插件 schema 注入的常驻 Ribbon 项计入固定版 tab 根入口。依选择状态才出现的 contextual tab、右键菜单、浮动对象工具栏、工作表标签菜单和页脚菜单分开统计。
- 本地数值统计 `createSheetRibbon` / `createDocumentRibbon` 的常规根声明，包含条件隐藏项；Sheets 动态函数分类按当前 catalog 的 8 类静态展开。嵌入组件缺省 classic，示例缺省 grid，两种变体共享 schema。额外入口仍计入本地数量，但不因此认为与上游功能等价；整组进入 More 不增加入口数。
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

v1.6.0 OSS 常规 Ribbon 静态根声明为 **38 项**，来自 `scripts/audit-univer-oss-ribbon.mjs` 读取共享、Docs 和插件 schema；未执行 factory 的动态条件。2026-11-25 人工静态回读当前 document-ribbon，本地常规根声明为 **52 项（33/16/3）**，含 hidden、排除 children/context；原 50（33/14/3）为 11-22 快照。Pro v1.6.0 缓存注册及未解析入口见[Pro Ribbon 统计](../UNIVER_TOOLBAR_PARITY.md#根入口统计与布局)，不再用 1.5.2 的 9 项作为当前总数。

| 标签 | 固定版分组和入口数 | 本地当前根项 | 对齐状态 |
| --- | --- | ---: | --- |
| Home / 开始 | HISTORY 5、FORMAT 11、LAYOUT 5、OTHERS 4，共 25 | 33 | 本地 history 5、format 11、layout 5、others 2，另有段落工具 5、其他工具 5。列表、颜色、格式刷与粘贴已有实现切片；快捷键/功能搜索和全部分组状态仍须核对。 |
| Insert / 插入 | MEDIA 7、CELL 1、OTHERS 2，共 10 | 16 | 本地 media 9、cell 2、structured 3、extensions 2；代码块、引用/Callout/多栏、目录及内容控件已有接线，全部 Pro 仍按对应包追踪；结构块差异见[证据](DOC-13-14-16-17-structured-blocks.md)。 |
| View / 视图 | DISPLAY 3，共 3 | 3 | 页面设置、缩放、全屏三个根项；大纲、专注和打印现在位于 Home 扩展组。数量相同不代表行为已验收。 |
| 合计 | **38** | **52** | 这是静态常规声明数量，条件隐藏、Pro 替换与浮动入口仍需追踪，不表示功能完成。 |

固定版 Insert 的 MEDIA 网格由多个插件共同注册。Breaks 在 `r1c1` 跨两行，Table 在 `r1c2`，Emoji 在 `r1c3`，Symbol 在 `r2c3`，Image 在 `r1c4`，Hyperlink 在 `r1c5`，Horizontal Line 在 `r2c5`。本地之前没有显式 `gridLayout`，列优先 fallback 导致 Image 掉到第二行；并且 Breaks、Horizontal Line、Symbol、Emoji 被放入 OTHERS。现已按上述位置修正，定向 Ribbon 契约测试 20/20 通过。页面布局统一浏览器验收尚未执行。

固定版另有目录上下文 Ribbon 两项，仅选中目录时显示；页眉页脚上下文为 5 个控件，合计 7 个 OSS 上下文根声明。右键与浮动工具不合并进常规 38 项。目录是独立 opt-in 插件，插入位置为 Insert > CELL，更新和删除位于目录 contextual tab 与主区域 Layout 右键菜单。本地目录入口、模型/命令、配置和上下文控制器已接入；仍按 [DOC-04 证据](DOC-04-table-of-contents.md) 保留未验收。Pro 内容控件、形状、增强表格等上下文另列，不受 OSS 七项数量限制。

代码块在 OSS 固定源码中有模型和剪贴板支持；v1.6.0 Pro 发布包已静态识别 `docs-code-ui`、`docs-latex-ui`、`docs-sdt-ui` 与 `docs-reference-ui` 的根声明。它们的模式、状态与命令闭环须按对应包及公开规格继续核对，不能把静态注册当成最终效果。段落设置和节设置是段落右键 Layout 组入口，不属于 Ribbon 项。

### Sheets

v1.6.0 OSS 常规 Ribbon 静态根声明为 **63 项**，2026-11-26 人工回读本地常规 schema 为 **71 项（30/7/11/17/6）**，含范围主题、图片与背景根；69/68 为此前清单快照。根声明与子项分开，More 不增加数量。Pro 注册中的重复挂载、替换与上下文条件仍待继续追踪，不将历史 1.5.2 的 15 项加到当前 OSS 数字。完整逐组、上下文、处理器和 Pro 边界见[Sheets Ribbon 源码入口清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)、[图片入口](SHEET-19-images.md)及[Pro Ribbon 统计](../UNIVER_TOOLBAR_PARITY.md#根入口统计与布局)。

| 标签 | 固定版根项 | 本地当前根项 | 主要差异 |
| --- | ---: | ---: | --- |
| START / 开始 | 30 | 30 | 本地 clipboard 5、font 11、layout 6、number 5、tools 3；工具第 3 项为本地范围主题，grid r1c3 跨两行。数量相同不表示入口语义一致；查找在 Data，快捷键/功能搜索和全部动态状态仍须核对。 |
| INSERT / 插入 | 9 | 7 | 本地 media 4、cells 3；图片根在媒体首项 r1c1 跨两行，浮动/单元格图片两子项；超链接、图表、迷你图及快速验证已有接线。线程评论和快捷数据条/图标集仍缺完整入口闭环；图片保真/Office 未验收。 |
| FORMULAS / 公式 | 12 | 11 | 固定版是常用加十类函数菜单与计算选项；本地 8 个 catalog 根分类、计算模式/编辑公式/重算 3 项，完整函数目录与语义未齐。 |
| DATA / 数据 | 6 | 17 | 本地 organization 2、rules 2、pivot 4、outline 8、data-tools 1。筛选切换/清除/重应用已接入；表格选择器及完整 Pro 分组/分析仍须核对。 |
| VIEW / 视图 | 6 | 6 | 本地 display 4、visibility 2；背景根在冻结后、order 0.5、r1c2 跨两行，设置/删除两子项接实际模型；十字高亮属于固定 OSS 入口。背景历史/资源重开、两后端和所有动态状态仍未验收。 |
| 合计 | **63** | **71** | 本地包含扩展项，常规数量、上下文数量与功能完成度分别记录。 |

Sheets 固定源码的组顺序为 START HISTORY/FORMAT/LAYOUT/NUMBER/OTHERS；INSERT EDIT/MEDIA/CELL/OTHERS；FORMULAS BASIC/OTHERS；DATA FORMULAS/RULES/ORGANIZATION/ANALYSIS/OTHERS；VIEW DISPLAY/VISIBILITY。零根项的组不进入数量；根声明按菜单 schema 的产品归属静态投影。固定 desktop 装配与 factory 决定最终显隐，尚未执行。本地公式按当前 117 项 catalog 的 8 类静态展开；没有执行求值引擎。

固定版还有 MAIN_AREA、COL_HEADER、ROW_HEADER、FOOTER_TABS、绘图对象上下文菜单。脚本记录 Sheets 130 个右键/浮动层级节点，包含父子层级，不当作根数。本地行列结构/显隐右键传 sheetId/ranges，尺寸入口传命中 sheetId/index；标签复制、颜色与保护亦有实现切片，关联资源、veryHidden 和全部权限/保存闭环仍未验收。具体来源和差异见 Sheets 清点文档。

填充柄外观使用 1 CSS px 选区/预览边框，2026-11-21 历史描边测试覆盖 0.75、1 和 1.5 缩放，不覆盖后续修改。2026-11-25 已补共享 AutoFill/Refill/CopyDown/CopyRight/AutoClear、FRange 和 Ctrl/Cmd+D/R；Canvas 使用 antdv-next 四模式选项，拖动和双击共用推断默认、单向范围和内缩清尾。序列新增数值线性/等比、文本数字/补零、中文数字/星期、英文星期/月份及其他循环；筛选行、合并、CF/DV、权限与验证进入同一候选事务。重填修订同一历史，等价内容只更新选区/菜单；边缘自动滚动沿用现有调度。见 [SHEET-02 自动填充证据](SHEET-02-auto-fill.md)。本批草稿未执行，standalone 重填 receipt、超过同步预算的 Worker、完整日期/公式边界、真实 IME/冻结/主题/滚动和 Office 往返仍未验收。

2026-11-26 [局部移位事务](SHEET-02-cell-shift-transactions.md)复用四核心+四确认命令、两个父项/四个移位子项，右键子项顺序按固定 schema；新增常规/扩展/溢出/独立上下文根均为 0。确认、Facade Dimension、完整候选/资源、权限/验证、内容与选择历史、宿主重入保护已落盘，56 个草稿未运行；完整 drawing/media、全部资源组合、Office、IME/两后端和性能仍有缺口。

同日[共享绘图工作台](SHEET-21-drawing-workbench-affine.md#entry-inventory)接组合/排列/对齐三个右键父项（2/4/8 叶）及本地删除，Canvas 浮条四个控件复用同一命令。它们属于上下文展示位置，本批常规/顶部扩展/独立 More 根增量均为 0。八个对齐 operation ID 已按固定 `sheet.operation` 前缀修正；后续 `sheet.command.move-drawing` 方向键/API 入口、实际尺寸边界与 1000ms 连续历史不新增 Ribbon 根。主任务当前 82 条、Sheets 图片批 66 条及[Docs 图片批](DOC-07-image-editing.md#image-batch)57 条草稿均未运行，[slicer/timeline affine 与 Ant 控件](SHEET-25-slicer-timeline-affine.md)已补入独立 98/182 未运行草稿；[Pro 多目标连接](SHEET-25-slicer-connections.md)另批已落盘未验收。完整 shape 作者、交叉筛选成员状态/Pro API、组 clipboard、Office 与性能仍缺。

## 顶部入口对应功能闭环

逐项补齐以下链路后才可更新功能状态：入口可见/禁用规则 → 共用 command → model/service mutation → 权限检查 → 撤销/重做 → 保存/重开与文件互操作 → 错误反馈。仅有按钮、面板或 command handler 不算完成。

### Docs

- [ ] DOC-01 文档模型、渲染与分页布局。
- [ ] DOC-02 字符格式、段落格式和混合选区状态。
- [ ] DOC-03 分页、分栏、分节及页眉页脚。
- [ ] DOC-04 大纲、标题折叠和正文目录；目录入口、上下文更新/删除、跳转、页码与前导符完整闭环。
- [ ] DOC-05 列表和增强列表。
- [ ] DOC-06 表格、表格上下文命令、尺寸、合并拆分和样式。
- [ ] DOC-07 图片、绘图、锚点及文字环绕；[单图/inline 与多选、六对齐/两分布/均匀变换/批删、四 Facade 和 57 条新草稿](DOC-07-image-editing.md#image-batch)，未验收。跨页 anchor、混合对象/非均匀 affine、Office 和真实 UI/性能仍缺。
- [ ] DOC-08 超链接和书签。
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
- [ ] SHEET-09 公式、函数和定义名称。
- [ ] SHEET-10 高级公式（Pro 目标）。
- [ ] SHEET-11 筛选。
- [ ] SHEET-12 排序。
- [ ] SHEET-13 数据验证。
- [ ] SHEET-14 条件格式。
- [ ] SHEET-15 查找替换。
- [ ] SHEET-16 超链接。
- [ ] SHEET-17 批注与线程评论。
- [ ] SHEET-18 结构化表格。
- [ ] SHEET-19 图片和水印；[图片/背景两常规根、完整矩阵剪贴板、锚点投影与 66 条草稿](SHEET-19-images.md)、[最终 filter/auto-height 与组框生命周期](SHEET-21-drawing-workbench-affine.md#model-lifecycle)，未验收；水印及 Office 仍缺。
- [ ] SHEET-20 十字高亮。
- [ ] SHEET-21 形状及绘图（Pro 目标）；[组合/混合层级与 45 条草稿](SHEET-21-drawing-groups-and-mixed-order.md)、[选区/对齐/移动/Canvas/Facade、作者 affine 与 82 条草稿](SHEET-21-drawing-workbench-affine.md)，未验收；[slicer/timeline affine、五类叶/组与 Ant 控件](SHEET-25-slicer-timeline-affine.md)及 [Pro 多目标连接](SHEET-25-slicer-connections.md)已另批落盘但未运行。完整 shape 作者、交叉筛选成员状态/Pro API 和组 clipboard/Office 仍缺。
- [ ] SHEET-22 可编辑图表（Pro 目标）；[逐型审计](SHEET-22-chart-type-audit.md)、[Relation/词云资源](SHEET-22-relation-word-cloud-resources.md)、[图例](SHEET-22-legend-selection-wrap.md)、[标题配置](SHEET-22-title-subtitle-contract.md)、[标题直接编辑](SHEET-22-title-direct-editing.md)、[绘图区大小/恢复](SHEET-22-plot-area-direct-editing.md)、[层级图 Ant Tooltip](SHEET-22-hierarchy-ant-tooltip.md)、[四个对象上下文根/剪贴板/PNG/删除与 34 个草稿](SHEET-22-chart-object-context-actions.md)、[真实 Drawer 路由/草稿保护与 33 个草稿](SHEET-22-chart-edit-panel-routing.md)、[Line/Smooth/Step 二维/3D 与 34 个新草稿](SHEET-22-line-and-area-styles.md)，均未验收。27 route 分组是导航分类，不是完成数量；本批线型常规根增量为 0。
- [ ] SHEET-23 数据透视表（Pro 目标）；[固定 15 项值显示方式](SHEET-23-24-pivot-show-data-as.md)与 [命令/Facade/linked 图格式](SHEET-23-24-pivot-show-data-as-facade.md)已落盘，枚举不计 Ribbon 根，所有草稿未运行；完整 UI/展开折叠/日期条件/Office/大样本继续闭环。
- [ ] SHEET-24 透视图（Pro 目标）；[owned/linked 创建与一次作者配置](SHEET-24-owned-pivot-charts.md)、[字段引擎](SHEET-24-pivot-chart-engine.md)及 [四字段区/11 汇总/标签筛选 UI](SHEET-24-pivot-chart-ui.md)已落盘；上游两条互斥根、两个创建子项、四个轴子项与动态层级单列，本地常规根增量 0；全部草稿未运行，同时创建表图/展开折叠/全部 Pro API/Office/真实 UI/性能继续闭环。
- [ ] SHEET-25 切片器和时间线（Pro 目标）；[原 affine/Ant 与 98/182 草稿](SHEET-25-slicer-timeline-affine.md)、[共享多目标连接/Facade/历史](SHEET-25-slicer-connections.md)、[四格式根与连接 UI](SHEET-25-slicer-connections-ui.md)、[复制/结构生命周期](SHEET-25-slicer-connections-lifecycle.md)已落盘，常规根增量 0、上下文四根分开统计；全部新草稿未运行，完整交叉筛选成员状态、日期/Pro API、缺依赖跨 workbook、Office 与真实 UI/性能仍缺。
- [ ] SHEET-26 迷你图（Pro 目标）。
- [ ] SHEET-27 大纲分组（Pro 目标）。
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
