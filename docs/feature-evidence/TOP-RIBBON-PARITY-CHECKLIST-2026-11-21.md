# Docs 与 Sheets 顶部功能区对齐清单

审阅日期：2026-11-21  
固定 Univer 版本：`6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68`  
Microsoft 资料审阅日期：2026-11-21，主要以 Microsoft 365 桌面版公开说明为准。

本清单先核对固定版 Univer 的顶部入口、插件扩展和上下文菜单，再核对本项目的入口、网格布局和命令处理器。Microsoft Word / Excel 的公开功能用于发现可补充能力，不代替 Univer 固定版入口对照。功能 ID 的验收状态仍以[完整功能对齐表](../UNIVER_FEATURE_PARITY.md)为准。

## 统计口径

- 一个 Ribbon 根入口是合并菜单 schema 后位于某个标签和分组中的顶层项。split、menu、select 的子项不增加根入口计数，动态菜单按其根菜单计一次。Docs 的 10×10 表格尺寸格是 Table 的子菜单，不另算 100 个按钮。
- 插件 schema 注入的常驻 Ribbon 项计入固定版 tab 根入口。依选择状态才出现的 contextual tab、右键菜单、浮动对象工具栏、工作表标签菜单和页脚菜单分开统计。
- 本地数值统计当前生产工作台实际使用的 Grid Ribbon schema。额外入口仍计入本地数量，但不因此认为它们与上游功能等价。整组进入 More 仍然属于该标签入口。
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

固定版常规 Ribbon 根入口共 **38 项**，分组如下。计数包括 OSS 插件扩展注册到默认 Ribbon 的入口，不含子菜单。

| 标签 | 固定版分组和入口数 | 本地当前根项 | 对齐状态 |
| --- | --- | ---: | --- |
| Home / 开始 | HISTORY 5、FORMAT 11、LAYOUT 5、OTHERS 4，共 25 | 28 | 有额外复制/剪切/普通粘贴等入口；缺固定格式刷、专用粘贴特殊项和其他固定入口；部分 FORMAT/LAYOUT 网格坐标未齐。 |
| Insert / 插入 | EDIT 0、MEDIA 7、CELL 1、OTHERS 2，共 10 | 11 | MEDIA 分组和坐标刚按源码修正；CELL 的目录插入入口尚未接线；本地代码块、内容控件、脚注和 LaTeX 是额外扩展入口。 |
| View / 视图 | DISPLAY 3，共 3 | 8 | 本地加入纸张/方向、文档大纲、专注模式和打印等入口，超出固定版三个根项；需逐项定组并核验功能路径。 |
| 合计 | **38** | **47** | 数量差异包括本地增强项和错放项，不能据此判定能力完整度。 |

固定版 Insert 的 MEDIA 网格由多个插件共同注册。Breaks 在 `r1c1` 跨两行，Table 在 `r1c2`，Emoji 在 `r1c3`，Symbol 在 `r2c3`，Image 在 `r1c4`，Hyperlink 在 `r1c5`，Horizontal Line 在 `r2c5`。本地之前没有显式 `gridLayout`，列优先 fallback 导致 Image 掉到第二行；并且 Breaks、Horizontal Line、Symbol、Emoji 被放入 OTHERS。现已按上述位置修正，定向 Ribbon 契约测试 20/20 通过。页面布局统一浏览器验收尚未执行。

固定版另有目录上下文 Ribbon 两项，仅选中目录时显示；Docs 插件清点还发现页眉页脚上下文工具栏的 5 个控件。右键、浮动工具栏和插件注入入口不合并进常规 38 项。目录是独立 opt-in 插件，固定版插入位置为 Insert > CELL，更新和删除位于目录 contextual tab 与主区域 Layout 右键菜单。本地目录模型/命令与配置组件已有定向契约，但工作台入口和上下文控制器尚未接通。

代码块在固定源码中有模型和剪贴板支持，但未发现默认 Insert Ribbon 注册；脚注、内容控件和 LaTeX 的商业 UI 注册也未在固定检出中确认。它们在本地按扩展功能记录，不能伪称固定版入口。段落设置和节设置是段落右键 Layout 组入口，不属于 Ribbon 项。

### Sheets

固定版常规 Ribbon 根入口共 **62 项**，本地当前 schema 共 **78 项**。按标签统计见下表。完整逐组、上下文、处理器和 Pro 边界记录在[Sheets Ribbon 源码入口清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)。

| 标签 | 固定版根项 | 本地当前根项 | 主要差异 |
| --- | ---: | ---: | --- |
| START / 开始 | 30 | 32 | 查找和条件格式错放到 Start；格式刷、清除格式/粘贴特殊类型、单独清除文本色/底色等仍需核对。 |
| INSERT / 插入 | 9 | 12 | 图片、超链接、线程评论和快捷条件格式入口缺失；行列插删及透视类入口错放或重复。 |
| FORMULAS / 公式 | 12 | 12 | 固定版为 11 个函数分类加计算选项；本地是插入函数、8 个 catalog 分类及自定义计算项，分类语义不同。 |
| DATA / 数据 | 6 | 18 | 分组和顺序差异大；缺表格选择器、独立规则入口和固定版筛选命令组合。 |
| VIEW / 视图 | 5 | 4 | 缺工作表背景图入口及添加/删除处理器。 |
| 合计 | **62** | **78** | 本地包含额外能力和重复入口，计数不代表完整对齐。 |

Sheets 固定源码的根组顺序为 START HISTORY/FORMAT/LAYOUT/NUMBER/OTHERS；INSERT EDIT/MEDIA/CELL/OTHERS；FORMULAS BASIC/OTHERS；DATA FORMULAS/RULES/ORGANIZATION/ANALYSIS/OTHERS；VIEW DISPLAY/VISIBILITY。固定版上游根数按 MenuManager、desktop preset 和 OSS 插件注册合并计算。本地公式动态分类按当前 117 项 catalog 生成的 8 个分类计数。

固定版还有 MAIN_AREA、COL_HEADER、ROW_HEADER、FOOTER_TABS、绘图对象上下文菜单。当前本地行列右键菜单未先把命中行/列映射为操作目标，表格标签右键也缺复制、标签颜色和工作表保护等入口。具体差异、涉及的 SOURCES、handler 状态和 SHEET-01 至 SHEET-30 Pro 目标见 Sheets 清点文档。

填充柄外观现已统一到 1 CSS px 选区/预览边框，测试覆盖 0.75、1 和 1.5 缩放。当前已接入双击自动推断向下填充范围，并在计划层支持 COPY/SERIES/ONLY_FORMAT/NO_FORMAT；定向契约记录见 [SHEET-02 自动填充证据](SHEET-02-auto-fill.md)。填充选项尚无画布菜单，画布当前只触发 SERIES；序列规则仅覆盖规则数值和数值日期序列，缺工作日/月份/中文及文本数字扩展、Refill 重应用和拖出视口后的自动滚动。浏览器交互与 Office 文件往返尚未验证，SHEET-02 仍未验收。

## 顶部入口对应功能闭环

逐项补齐以下链路后才可更新功能状态：入口可见/禁用规则 → 共用 command → model/service mutation → 权限检查 → 撤销/重做 → 保存/重开与文件互操作 → 错误反馈。仅有按钮、面板或 command handler 不算完成。

### Docs

- [ ] DOC-01 文档模型、渲染与分页布局。
- [ ] DOC-02 字符格式、段落格式和混合选区状态。
- [ ] DOC-03 分页、分栏、分节及页眉页脚。
- [ ] DOC-04 大纲、标题折叠和正文目录；目录入口、上下文更新/删除、跳转、页码与前导符完整闭环。
- [ ] DOC-05 列表和增强列表。
- [ ] DOC-06 表格、表格上下文命令、尺寸、合并拆分和样式。
- [ ] DOC-07 图片、绘图、锚点及文字环绕。
- [ ] DOC-08 超链接和书签。
- [ ] DOC-09 查找替换及范围/选区作用域。
- [ ] DOC-10 线程评论入口、锚点、回复/编辑/解决、权限和持久同步。
- [ ] DOC-11 内容控件及锁定/重复结构。
- [ ] DOC-12 脚注和尾注的编辑、编号、分页与 DOCX 映射。
- [ ] DOC-13 代码块。
- [ ] DOC-14 多栏组与节级分栏。
- [ ] DOC-15 LaTeX 公式。
- [ ] DOC-16 引用块。
- [ ] DOC-17 Callout 标注块。
- [ ] DOC-18 形状、连接器和 SmartArt。
- [ ] DOC-19 文档图表。
- [ ] DOC-20 水印。
- [ ] DOC-21 打印和 PDF。
- [ ] DOC-22 DOC/DOCX 导入、真实写出与独立重开。
- [ ] DOC-23 文档统计、页信息及真实布局进度。

### Sheets

- [ ] SHEET-01 工作簿和工作表生命周期。
- [ ] SHEET-02 范围、选区、填充和网格编辑。
- [ ] SHEET-03 行列结构、尺寸、隐藏/显示和自动适应。
- [ ] SHEET-04 样式、默认样式、范围主题及混合态。
- [ ] SHEET-05 富文本。
- [ ] SHEET-06 数字格式和 1900/1904 日期系统。
- [ ] SHEET-07 冻结、网格线、视口和缩放。
- [ ] SHEET-08 剪贴板和选择性粘贴。
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
- [ ] SHEET-19 图片和水印。
- [ ] SHEET-20 十字高亮。
- [ ] SHEET-21 形状及绘图（Pro 目标）。
- [ ] SHEET-22 可编辑图表（Pro 目标）。
- [ ] SHEET-23 数据透视表（Pro 目标）。
- [ ] SHEET-24 透视图（Pro 目标）。
- [ ] SHEET-25 切片器和时间线（Pro 目标）。
- [ ] SHEET-26 迷你图（Pro 目标）。
- [ ] SHEET-27 大纲分组（Pro 目标）。
- [ ] SHEET-28 打印和 PDF（Pro 目标）。
- [ ] SHEET-29 Office 导入与导出。
- [ ] SHEET-30 工作簿、工作表和范围保护。

SHEET-02 的填充柄预览描边契约已通过定向测试，但双击填充、序列推断、填充选项、重填和视口外拖动仍为未实现项，不能因为填充柄可命中/可拖动就关闭该 ID。

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

## 本轮已完成与仍待完成

- [x] 写入项目级规则，要求主任务和全部子任务先清点各标签、插件扩展、分组/行列、条件、命令和统计范围，再实现功能闭环。
- [x] 固定版 Sheets 插件菜单合并和本地入口/handler 差异已写入独立清点文件。
- [x] 固定版 Docs 顶层菜单计数和 Insert MEDIA 网格位置已源码复核；本地媒体网格契约测试通过。
- [ ] 将 Docs Insert CELL 目录入口、contextual 更新/删除与本地 TOC dialog/commands 接通。
- [ ] 接通 Docs 评论入口、线程模型/服务及保存同步；当前面板契约组件未接入 Workbench。
- [ ] 根据本清单逐项修正 Docs 和 Sheets 的错组、错序、缺入口、错 handler 与实际功能差异。
- [ ] 实现或如实保留所有 Univer Pro 目标和 Microsoft 补充候选项，按功能 ID 提供契约/集成及 Office 文件验证。
- [ ] 按用户要求，等整批入口和功能实现稳定后统一进行浏览器验收；此清单的源码对照及契约测试不代替该步骤。

本清单的源码与类型/契约检查不能代替可编辑工作流、权限、撤销重做、持久化、独立 Office 文件重开和浏览器布局验收。当前未修改功能对齐表中任何功能的验收状态。
