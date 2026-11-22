# 完整功能与 Pro 对齐清单

基线日期：2026-11-21
源码：`dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`
公开规格：当日官方 Sheets / Docs 功能导航、功能说明、API 与产品对照页。  
执行依据：[项目标准](../PROJECT_EXECUTION_STANDARD.md)和[源码布局规则](UNIVER_REFERENCE_RULES.md)。

技术路径遵守[Canvas/WebGPU 与 Rust/WASM 计算标准](RENDERING_AND_COMPUTE_STANDARD.md)。WebGPU 为优先绘制后端，Canvas2D 为功能一致的回退；大型计算优先交给 Worker 中的 Rust/WASM。

## 1. 范围与状态

用户目标是 Sheets / Docs 工作台的完整功能与界面对齐，包括 Pro。Pro 仅描述上游发行方式，不改变本项目的交付等级；没有公开实现的部分按公开规格独立实现。这里不把 Slides、Bases、Boards 或独立 AI SDK 整个产品线自动加入本次 Sheets / Docs 范围。

表中来源分为“公开源码”“公开规格”与二者组合。Pro 可以通过公开接口、可访问演示和标准 Office 样本反推可观察行为；记录已核实规格和待核实推断。所有 Sheets / Docs Pro 能力都在范围内，可以分批独立实现。

SHEET-29 和 DOC-22 对齐 [Univer Exchange 的公开行为](UNIVER_REFERENCE_RULES.md#8-rustwasm保真与服务边界)，Office 文件转换由本项目读写链路承担。独立快照转换与协作文件的服务端版本、权限和持久化分别验收。

**当前所有条目均为未验收。** 已有组件草稿和 Rust 解析接口不作为完整验收证据。各项可有部分进度，只有补齐第 5 节的证据后才可标为已验收。

清单验收状态仅使用“未验收 / 已验收”；开发进度另用“待核实、待开发、开发中、待验证、待修复”记录。未提供完整证据时保持未验收；永久禁用、按钮提示、静态样例、仅保存 JSON 不能记为功能完成。

## 2. 共享工程与工作台

| ID | 能力 | 公开依据 | 本项目交付与验收要点 | 当前 |
| --- | --- | --- | --- | --- |
| COMMON-01 | 应用壳与工作台区域 | 公开源码：ui/workbench、parts | 文件操作、类型切换、Ribbon、产品 Header、编辑区、面板、单条产品 Footer 和统一浮层；无重复工具区 | 未验收；[工作台 Grid/Classic 切换验证](UNIVER_TOOLBAR_PARITY.md#1-当前装配与证据边界) |
| COMMON-02 | 命令、菜单和插件 | 公开源码：core/command、ui/menu、presets | command/mutation/operation 分层；各入口同命令；按 feature 注册、排序、上下文显隐/禁用、正确销毁 | 未验收 |
| COMMON-03 | 选择、输入、快捷键与剪贴板 | 公开源码：Sheets/Docs selection、editing、clipboard | 中文输入法、光标/选区保留、系统剪贴板；Docs 源格式/目标格式/纯文本粘贴及最近粘贴切换；工具栏和右键不破坏范围 | 未验收 |
| COMMON-04 | 主题、语言、字体与配置 | 公开源码 ui/config、examples/workbench-settings；[主题规格][shared-themes] | 中文默认、多语言资源、明暗主题、字体加载/度量、LTR/RTL、布局变体与配置保存；不污染文档内容 | 未验收 |
| COMMON-05 | 统一快照、公开 API 与内部保存 | 公开源码 IWorkbookData、IDocumentData、Facade | 稳定模型版本、增量变更、公共 API 与事件；保存失败/dirty/savepoint 正确；重开保持模型 | 未验收；日期系统与编辑快照的实现及待验证边界见[数字格式输入](UNIVER_TOOLBAR_PARITY.md#sheet-number-input) |
| COMMON-06 | 权限与只读 | 公开源码 permission；[公开规格][shared-permissions] | 单元、工作表/范围、段落等权限；查看/编辑/复制/评论/导出/打印分开；服务端与命令入口执行权限 | 未验收 |
| COMMON-07 | 实时多人协作（Pro/Server 目标） | [协作规格][shared-collaboration]、[Sheets][sheet-collaboration]、[Docs][doc-collaboration] | 实际服务端、并发协议、用户与远端范围、确认/重试/去重、断线重连、离线队列与多用户一致性 | 未验收 |
| COMMON-08 | 持久历史版本（Pro/Server 目标） | [共享历史][shared-history]、[Sheets][sheet-history]、[Docs][doc-history] | 版本列表、快照+变更日志、版本内容、对比、恢复和恢复后广播；重启后可查；区别于本地撤销 | 未验收 |
| COMMON-09 | 语义对比与恢复 | [公开对比规格][shared-compare]及历史规格 | 内容/样式/结构差异、范围定位与只读历史视图；恢复形成新版本而非清空日志 | 未验收 |
| COMMON-10 | 跨文档嵌入（Pro 目标） | [公开嵌入规格][shared-embed] | Docs 嵌 Sheets、Sheets 对象/标签嵌 Docs；父子独立模型、焦点/权限/保存/生命周期；明确引用与快照语义 | 未验收 |
| COMMON-11 | Canvas/WebGPU、大文件与 Worker/WASM | 公开源码引擎边界 + 用户技术要求 | Canvas 内容渲染、WebGPU 优先/2D 回退；大计算优先 Rust/WASM；进度/取消/过期结果、稀疏/视口、内存与真实基准 | 未验收；基准方法见 [COMMON-11 基准记录](feature-evidence/COMMON-11-compute-benchmark.md) |
| COMMON-12 | 图片/媒体资源管理 | 公开源码 drawing、ImageIo | 资源 ID、二进制/关系、锚点/尺寸/裁剪/旋转/层级、引用计数；复制粘贴和保存后资源不丢失 | 未验收 |

协作、版本历史、持久评论与权限需要本项目内的真实服务及协议。BroadcastChannel 或多个本地数组只能作为辅助测试，不能替代跨用户服务验收。公开的“collaboration: true”、命令常量或数据类型并不构成能力实现。

## 3. Sheets 全量目标

顶部 Ribbon 根入口、插件扩展、上下文菜单与 Microsoft Excel 补充候选项见[Docs 与 Sheets 顶部功能区对齐清单](feature-evidence/TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)及[Sheets 固定源码入口清点](feature-evidence/SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)。

| ID | 能力 | 公开依据 | 本项目交付与验收要点 | 当前 |
| --- | --- | --- | --- | --- |
| SHEET-01 | 工作簿、工作表生命周期 | 公开源码 sheets、sheets-ui；[核心][sheet-core]、[API][sheet-api] | 新建/复制/删除/重命名/重排、颜色、隐藏/veryHidden、活动表、校验和可撤销事务 | 未验收 |
| SHEET-02 | 范围、选区和网格编辑 | 公开源码 selection/editor；[范围规格][sheet-selection] | 单元格/区域/行列/全选/多选、拖选、填充、合并、F2/直接输入、Enter/Tab/Escape；双向视口处理 | 未验收；[合并、取消合并与分类清除](feature-evidence/SHEET-02-merge-clear.md) |
| SHEET-03 | 行列结构与尺寸 | [行列规格][sheet-row-col]、公开源码 commands | 插删/移动/复制/隐藏/显示/自动适应、行高列宽；更新公式与所有范围资源 | 未验收 |
| SHEET-04 | 样式、默认样式、范围主题 | 公开源码 style/numfmt；[默认样式][sheet-style]、[范围主题][sheet-range-theme] | 字体、填充、边框、对齐、旋转、换行、缩小适应、继承/主题与混合态；按真实模型生效 | 未验收；[字体与字号](feature-evidence/SHEET-04-font-family-size.md)、[颜色与重置](feature-evidence/SHEET-04-color-reset.md)、[边框](feature-evidence/SHEET-04-borders.md)、[旋转/竖排/缩小适应](feature-evidence/SHEET-04-text-rotation-shrink-to-fit.md)、[分类清除](feature-evidence/SHEET-02-merge-clear.md) |
| SHEET-05 | 富文本 | [富文本规格][sheet-rich-text]、公开源码 docs/editor bridge | 单格多运行样式、部分文本选择和编辑、复制粘贴、链接与导出；不扁平化成一个字符串样式 | 未验收 |
| SHEET-06 | 数字格式与日期系统 | [数字格式][sheet-numfmt]、公开源码 sheets-numfmt | 数字/货币/会计/百分比/日期/自定义、小数位、locale、1900/1904；原始值与显示值分离 | 未验收；[入口、编辑语义与待统一验证](UNIVER_TOOLBAR_PARITY.md#sheet-number-input) |
| SHEET-07 | 冻结、网格线、视口和缩放 | [冻结][sheet-freeze]、[网格线][sheet-gridlines]、公开源码 render | 多分区真实冻结、坐标命中/编辑一致、每表视口保存、组合底栏缩放 | 未验收 |
| SHEET-08 | 复制粘贴与选择性粘贴 | [剪贴板][sheet-clipboard]、公开源码 clipboard | 值/公式/格式/列宽/排除边框、剪切移动、外部 HTML/TSV、引用调整、合并冲突与权限 | 未验收 |
| SHEET-09 | 公式与定义名称 | [公式][sheet-formula]、[命名][sheet-names]、公开源码 engine-formula/sheets-formula | 词法/AST/求值、函数目录、相对绝对/跨表引用、名称作用域、依赖与重算、错误传播和循环依赖 | 未验收 |
| SHEET-10 | 高级公式（Pro 增强目标） | [高级公式规格][sheet-advanced-formula] | 按公开函数/行为清单实现高级求值、依赖和性能；补缓存/数组/引用等公开场景，不能以 SUM 示例代表引擎 | 未验收 |
| SHEET-11 | 筛选 | [筛选][sheet-filter]、公开源码 filter | 值/文本/数值/日期/颜色等公开条件、筛选结果与可见行、恢复、结构编辑后的范围变换；不删除源数据 | 未验收；现有入口、实现与遗留验证见[顶部功能对照](UNIVER_TOOLBAR_PARITY.md#3-sheets-顶部功能逐项映射) |
| SHEET-12 | 排序 | [排序][sheet-sort]、公开源码 sort | 多关键字、类型/locale 比较、范围与整表、关联单元格/样式/引用、撤销；不只排序显示字符串 | 未验收；现有入口、实现与遗留验证见[顶部功能对照](UNIVER_TOOLBAR_PARITY.md#3-sheets-顶部功能逐项映射) |
| SHEET-13 | 数据验证 | [验证][sheet-validation]、公开源码 data-validation | 全部公开规则、范围/公式依赖、下拉、提示与无效状态；键入/粘贴/填充/API 均受验证 | 未验收；[数据验证切片证据](feature-evidence/SHEET-13-data-validation.md) |
| SHEET-14 | 条件格式 | [条件格式][sheet-conditional]、公开源码 conditional-formatting | 规则/范围/优先级、公式、高亮/色阶/数据条/图标集、动态重算、样式叠加；不直接写死底色 | 未验收 |
| SHEET-15 | 查找替换 | [查找替换][sheet-find]、公开源码 find-replace | 值/公式/格式化文本等公开作用域与选项、结果定位、全部替换与历史；大文件可取消 | 未验收 |
| SHEET-16 | 超链接 | [超链接][sheet-links]、公开源码 hyper-link | 普通/富文本链接、URL/内部范围/工作表目标、编辑/删除、引用变换和安全跳转 | 未验收 |
| SHEET-17 | 批注与线程评论 | [批注][sheet-notes]、[评论][sheet-comments]、公开源码 note/thread-comment | 普通 Note 与 ThreadComment 分模型；锚点随行列变化、回复/解决、用户与持久同步 | 未验收 |
| SHEET-18 | 结构化表格 | [表格][sheet-table]、公开源码 table | 表范围/列/表头/汇总行、样式、自动扩展、筛选排序联动、结构化引用、转换回范围及导出 | 未验收 |
| SHEET-19 | 图片与水印 | [图片][sheet-images]、[水印][sheet-watermark]、公开源码 drawing/watermark | 插入/替换/移动/裁剪/缩放、锚点、文字/图片水印；打印及文件写回资源保真 | 未验收 |
| SHEET-20 | 十字高亮 | [公开规格][sheet-crosshair]、公开源码 crosshair-highlight | 活动行列提示、颜色配置、冻结/滚动兼容；纯视图 operation 不改文档内容 | 未验收 |
| SHEET-21 | 形状及绘图（含 Pro 目标） | [形状规格][sheet-shapes]、公开 drawing 基础 | 几何、箭头/连接线、填充/边框/文本、移动/调整/旋转/层级/组合、命中与 Office 映射 | 未验收 |
| SHEET-22 | 可编辑图表（Pro 目标） | [图表规格][sheet-charts] | 图表类型目录、系列/分类/区域绑定、轴/图例/标签/主题、数据变更刷新、属性面板、可编辑导出 | 未验收 |
| SHEET-23 | 数据透视表（Pro 目标） | [透视表规格][sheet-pivot] | 数据源/cache、字段行列筛选值、公开汇总方式、小计总计、排序筛选、日期/数字/手工分组、下钻、刷新 | 未验收 |
| SHEET-24 | 透视图（Pro 目标） | [透视图规格][sheet-pivot-chart] | 依赖透视模型和图表模型、字段联动、筛选刷新、编辑与持久化；不生成独立静态图 | 未验收 |
| SHEET-25 | 切片器和时间线（Pro 目标） | [公开规格][sheet-slicer] | 表格/透视目标绑定、多选/清除、交叉筛选、日期粒度与时间范围、样式/布局及保存 | 未验收；现有入口与绑定、日期、Office 往返边界见[顶部功能对照](UNIVER_TOOLBAR_PARITY.md#3-sheets-顶部功能逐项映射) |
| SHEET-26 | 迷你图（Pro 目标） | [公开规格][sheet-sparkline] | 公开图型、源范围、分组/比例/空值/负值样式、源更新与复制移动引用、导出 | 未验收 |
| SHEET-27 | 大纲分组（Pro 目标） | [分组规格][sheet-outline] | 行列嵌套组、层级、折叠/展开、汇总方向；与筛选/手动隐藏原因组合，不互相覆盖 | 未验收 |
| SHEET-28 | 打印和 PDF（Pro 目标） | [打印规格][sheet-print] | 纸张/方向/边距/打印区域/缩放适页、分页预览、重复标题、页眉页脚、网格与隐藏处理、真实输出 | 未验收 |
| SHEET-29 | 导入和 Office 导出（Pro/Server 对齐目标） | [导入导出规格][sheet-io]、duke-pages/rxls 的实际契约 | XLS/XLSX/CSV/TSV 对齐公开规格；保留已有 XLSB/ODS 读取；真实 XLSX/CSV/TSV 写出及重导入 | 未验收 |
| SHEET-30 | 工作簿/工作表/范围保护 | [权限规格][sheet-permission]、公开源码 permission | 保护粒度、可执行操作、只读反馈与撤销/协作边界；服务端验证，不能仅隐藏菜单 | 未验收 |

每个模块的“全部公开”指固定基线对应的具体行为/函数/图型目录，不代表可以用一条笼统说明代替细项。开发前把目录展开为可测试子项；未核实的子项记录为待核实，不能自行选择一小部分后将整个模块标完成。

## 4. Docs 全量目标

顶部 Ribbon 根入口、插件扩展、classic/grid 布局及上下文工具栏清点见[Docs 与 Sheets 顶部功能区对齐清单](feature-evidence/TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)。Microsoft 365 补充功能单列在 Univer/Pro 完整闭环之后，不计入当前 Univer 入口数和实现阶段。目录入口和闭环缺口见[DOC-04 目录证据](feature-evidence/DOC-04-table-of-contents.md)。

| ID | 能力 | 公开依据 | 本项目交付与验收要点 | 当前 |
| --- | --- | --- | --- | --- |
| DOC-01 | 文档模型、运行与排版 | [核心][doc-core]、[API][doc-api]、公开源码 docs/engine-render | 文本运行/段落/节、样式/主题、页面尺寸边距、换行/自动分页、光标与布局映射；传统/现代 flavor 分开 | 未验收 |
| DOC-02 | 字符和段落格式 | 核心 API、公开源码 docs-ui/menu | 字体/字号/颜色/高亮/上下标/字符样式、标题级别、对齐/缩进/行距/段前后、混合选区；格式与模型一致 | 未验收；[上下标实现切片](feature-evidence/DOC-02-vertical-align.md) |
| DOC-03 | 分页、分栏、分节与页眉页脚 | 公开源码 header-footer、insert-break、布局引擎 | 连续/下一页/奇偶页等公开分节语义、节级布局、首页/奇偶页、与前节关联、页码与字段更新 | 未验收；[分节符切片证据](feature-evidence/DOC-03-section-breaks.md) |
| DOC-04 | 左侧大纲、标题折叠和正文目录 | [目录规格][doc-toc]、公开源码 DocSideMenu/docs-toc/doc-heading-collapse/docs-toc-ui | 大纲导航；现代标题折叠保留子状态、隐藏选区迁回可见标题、历史/保存/同步；正文目录支持整体/页码更新与删除 | 未验收；[正文目录源码对照与本地差异](feature-evidence/DOC-04-table-of-contents.md) |
| DOC-05 | 列表与增强列表（含 Pro 目标） | [列表规格][doc-list]、公开基础列表 | 有序/无序/复选、嵌套/升降级、编号连续/重启、样式与重排；复制和导出保留语义 | 未验收；[清单列表切片证据](feature-evidence/DOC-05-checklist.md) |
| DOC-06 | 表格与增强表格（含 Pro 目标） | [表格规格][doc-table]、公开基础表格 | 行列/单元格模型、插删移动/尺寸、合并拆分、边框背景、文字对齐与整表对齐分别实现；上游整表对齐限传统模式，跨页/嵌套细项待核实并保留目标 | 未验收 |
| DOC-07 | 图片和文字环绕 | [图片规格][doc-images]、公开源码 docs-drawing | 行内/环绕、锚点、偏移、尺寸/裁剪/旋转/层级、替换及 Office 关系；输入/分页随对象更新 | 未验收 |
| DOC-08 | 超链接和书签/锚点 | [超链接规格][doc-links]、公开源码 docs-hyper-link | 范围目标、插入/编辑/删除、位置变换、跳转、序列化与文件映射 | 未验收 |
| DOC-09 | 查找替换 | [查找替换规格][doc-find]、公开源码 docs-find-replace | 遍历与作用域明确、匹配高亮/定位、可撤销替换；正文表格、页眉页脚等覆盖分别记录，不虚报上游一期限制 | 未验收 |
| DOC-10 | 线程评论 | [评论规格][doc-comments]、公开源码 docs-thread-comment | 范围锚点、编辑后的定位、回复/解决、权限、用户与持久化同步；不以本地数组验收多人评论 | 未验收 |
| DOC-11 | 内容控件（Pro 目标） | [内容控件规格][doc-controls] | 富/纯文本、复选、日期、列表、图片、分组/重复节等公开类型；占位、属性、内容锁/控件锁、导出映射 | 未验收 |
| DOC-12 | 脚注与尾注（Pro 目标） | [公开规格][doc-footnotes] | 上游 API 适用传统分页；引用与注释内容同一历史事务、编号/重排、跳转、分页排版、DOCX 结构保持；不冒称现代模式已支持 | 未验收 |
| DOC-13 | 代码块（Pro 目标） | [公开规格][doc-code] | 语言、源文本、语法高亮、块级选择/编辑、复制粘贴及保存；高亮 HTML 不作为唯一模型 | 未验收 |
| DOC-14 | 多栏组（Pro 目标） | [分栏规格][doc-columns] | 块级 2–5 栏、比例/间距/增删、选择/导航和统一命令；区分节级分栏，支持模式/嵌套规则按公开契约核验，不擅自限定现代模式 | 未验收 |
| DOC-15 | LaTeX（Pro 目标） | [公开规格][doc-latex] | 公式源文本、行内/块级公开形式、编辑/测量/渲染、错误反馈、保存与可行的 Office 映射 | 未验收 |
| DOC-16 | 引用块（Pro 目标） | [引用规格][doc-quote] | 独立块样式、内部富文本/段落、进入/退出/删除、复制与持久化；不是仅加一个左边框 | 未验收 |
| DOC-17 | 标注块 Callout（Pro 目标） | [公开规格][doc-callout] | 图标/颜色/内容/容器属性、块编辑和布局、选择、撤销与保存 | 未验收 |
| DOC-18 | 形状、连接器和 SmartArt（Pro 目标） | [形状规格][doc-shapes]、固定提交 README 商业扩展说明 | 对象/文本/布局/锚点/层级和连接关系；SmartArt 依公开图型、布局与样本另展开规格，不假称有源码 | 未验收 |
| DOC-19 | 文档图表（Pro 目标） | [共享图表规格][shared-charts]及官方 Docs 图表 API | 复用可编辑图表模型、内嵌数据/源引用、尺寸/环绕/分页、属性编辑及 Office 图表关系 | 未验收 |
| DOC-20 | 水印 | [水印规格][doc-watermark]、公开源码 watermark | 文本/图片、样式/透明度/定位、页面范围、打印与导出保留 | 未验收 |
| DOC-21 | 打印和 PDF（Pro 目标） | [打印规格][doc-print] | 使用真实排版、页范围/纸张/方向/边距、分页/页眉页脚/图表/脚注正确；真实打印/PDF 输出 | 未验收 |
| DOC-22 | 导入与 DOCX 写出（Pro/Server 对齐目标） | [导入导出规格][doc-io]、本项目 duke-pages 适配 | DOC/DOCX 导入、DOCX 真正写出；保留已有 ODT 读取；样式/节/列表/表格/对象/关系重开验证 | 未验收 |
| DOC-23 | 文档统计与排版状态 | 公开源码 DocFooter/DocStatistics/DocLayoutProgress | 文档/选区统计、页信息、缩放、排版任务真实进度；与内容历史、保存状态分开 | 未验收 |

传统分页和现代连续模式中，上游标明尚未支持/计划中的行为，应记录真实成熟度；用户完整目标仍保留。不能将未知细节写成“已经完整实现”，也不能将单个模式限制扩大为全部文档能力限制。

## 5. 每个功能的完成证据

把每个 ID 的细项和证据写到对应 feature 的验收记录中，清单的状态必须链接到记录。每项至少包含：

1. **规格**：源码路径与符号，或公开文档、API、演示和文件样本及核实日期；列出输入、动作、结果、上下文、边界和错误。反推的行为要有可复现用例，待核实项单列。
2. **模型**：节点/字段/schema、读写路径、引用变换、迁移/保留策略；明确权威数据源。
3. **实现**：命令 ID、mutation/逆 mutation、service/controller/store/view 边界及已修改文件。
4. **行为**：用户可通过页面实际完成；按钮、快捷键、右键/API 的效果一致；只读与错误可理解。
5. **闭环**：编辑 → 撤销/重做 → 保存 → 重开；涉及 Office 资源时再导出并用独立读取器/Office 兼容程序核对。
6. **界面**：在对应的视口、字体、缩放、主题及布局变体下验证布局、命中、键盘交互和控件状态。只有用户明确要求视觉回归时才执行截图或像素检查。
7. **回归**：适用的模型/命令/协议/浏览器测试、执行命令、结果及仍未通过的场景。模型、服务、命令、算法和 Worker 协议等可独立测试逻辑必须有有意义的单元或契约测试；编辑闭环、跨 feature 和权限行为补集成/浏览器测试。

8. **性能**：性能相关功能记录同一数据集的前后对照，包括设备、浏览器/Canvas2D 或 WebGPU 后端、数据规模、首屏和任务耗时、P50/P95 交互或帧时间、主线程长任务、Worker/WASM 跨边界传输和 JS/WASM/GPU 内存。没有可复现测量时保持“未验收”，不能以小样本、静态截图或构建通过代替性能证据。

不能用“pnpm build 通过”或“Cargo test 通过”填满以上八项。只读调研本身也不算功能实现。

## 6. 必须补齐的独立实现模块

| 模块 | 核心工作 | 不可替代的验收 |
| --- | --- | --- |
| office-io | 格式检测、duke/rxls 适配、源资源保留、DTO 校验、XLSX/DOCX writer、CSV/TSV | 真实二进制输出，样式/公式/关系/对象的导出重开 |
| formula-engine | 函数目录、解析/引用、依赖图、求值、增量重算与 Worker | 公开函数夹具、结构变更引用、错误传播与确定性 |
| chart-engine | 类型/系列/轴/主题、数据绑定、布局绘制、命中/编辑 | 编辑后重算刷新，图表模型与 Office 输出一致 |
| pivot-engine | cache、字段、分组聚合、排序筛选、小计/总计/下钻 | 明细核对、源更新、切片器与透视图联动 |
| document-layout | 字体度量、段落/表格/对象布局、节/分页/页眉页脚、注释和索引映射 | 跨页编辑、重排光标正确、PDF/打印与编辑布局一致 |
| document-extensions | 控件、注释、代码、分栏、LaTeX、引用、Callout、形状/SmartArt | 每种真实节点均可编辑、选择、撤销、保存并说明 Office 映射 |
| print-pipeline | 页设置、分页、预览、字体/媒体准备、PDF/打印输出 | 不是只调用 window.print 打印应用壳 |
| collaboration-server | 身份/权限、changeset、OT/CRDT 策略、版本确认、持久化、重连与 presence | 两个独立客户端并发、断网重连、重启服务后数据一致 |
| version-service | 快照/增量、版本查询、语义比较与恢复 | 跨会话可查，恢复产生新版本且同步到其他用户 |

这些模块是项目架构责任，不意味着必须一个模块一个 crate 或引入微服务。Rust 原生服务和浏览器 WASM 可共享纯模型/算法 crate，禁止把后端网络职责塞进 WASM 解析 API。

## 7. 实施顺序与不缩水原则

1. 统一 DTO、真实模型与导入闭环，消除内容丢失；完成 Pinia/命令/历史/Worker 的边界。
2. 以固定布局实现共享壳、Sheets 和 Docs 的真实基础行为，补充真实文件写出。
3. 在同一架构中接入公式/数据功能、绘图图表/透视、文档高级节点、打印。
4. 接入实际协作/持久版本服务，完成权限与并发边界。
5. 按 ID 覆盖跨 feature 的插删/粘贴/保存/重开/并发回归、界面行为及性能验证。
6. Univer 与全部 Pro 能力完成后，再评估 Microsoft 365 专属补充项；这些候选能力不替代或挤占前五阶段的完整目标。

以上是实施顺序，不是把第 1 阶段改叫“最终完成”。子任务可以并行，但文件归属、共享模型和命令协议先明确；交付清单与未完成项始终保留。

## 9. 官方公开规格索引

链接指向官方当日公开规格，内容会变化；复核时记录日期。开源实现的固定路径另见[源码索引](UNIVER_REFERENCE_RULES.md#10-固定源码索引)。

[shared-themes]: https://docs.univer.ai/zh-CN/overview/features/shared/themes-and-locales
[shared-permissions]: https://docs.univer.ai/zh-CN/overview/features/shared/permissions
[shared-collaboration]: https://docs.univer.ai/zh-CN/overview/features/shared/collaboration
[shared-history]: https://docs.univer.ai/zh-CN/overview/features/shared/history
[shared-compare]: https://docs.univer.ai/zh-CN/overview/features/shared/compare
[shared-embed]: https://docs.univer.ai/zh-CN/overview/features/shared/embed
[shared-charts]: https://docs.univer.ai/zh-CN/overview/features/charts
[sheet-core]: https://docs.univer.ai/zh-CN/guides/sheets/features/core
[sheet-api]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/sheets-api
[sheet-selection]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/range-selection
[sheet-row-col]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/row-col
[sheet-style]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/default-style
[sheet-range-theme]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/range-theme
[sheet-rich-text]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/rich-text
[sheet-numfmt]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/numfmt
[sheet-freeze]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/freeze
[sheet-gridlines]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/gridlines
[sheet-clipboard]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/clipboard
[sheet-formula]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/formula
[sheet-names]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/defined-names
[sheet-advanced-formula]: https://docs.univer.ai/zh-CN/guides/sheets/features/advanced-formula
[sheet-filter]: https://docs.univer.ai/zh-CN/guides/sheets/features/filter
[sheet-sort]: https://docs.univer.ai/zh-CN/guides/sheets/features/sort
[sheet-validation]: https://docs.univer.ai/zh-CN/guides/sheets/features/data-validation
[sheet-conditional]: https://docs.univer.ai/zh-CN/guides/sheets/features/conditional-formatting
[sheet-find]: https://docs.univer.ai/zh-CN/guides/sheets/features/find-replace
[sheet-links]: https://docs.univer.ai/zh-CN/guides/sheets/features/hyper-link
[sheet-notes]: https://docs.univer.ai/zh-CN/guides/sheets/features/notes
[sheet-comments]: https://docs.univer.ai/zh-CN/guides/sheets/features/comments
[sheet-table]: https://docs.univer.ai/zh-CN/guides/sheets/features/table
[sheet-images]: https://docs.univer.ai/zh-CN/guides/sheets/features/images
[sheet-watermark]: https://docs.univer.ai/zh-CN/guides/sheets/features/watermark
[sheet-crosshair]: https://docs.univer.ai/zh-CN/guides/sheets/features/crosshair-highlighting
[sheet-shapes]: https://docs.univer.ai/zh-CN/guides/sheets/features/shapes
[sheet-charts]: https://docs.univer.ai/zh-CN/guides/sheets/features/charts
[sheet-pivot]: https://docs.univer.ai/zh-CN/guides/sheets/features/pivot-table
[sheet-pivot-chart]: https://docs.univer.ai/zh-CN/guides/sheets/features/pivot-chart
[sheet-slicer]: https://docs.univer.ai/zh-CN/guides/sheets/features/slicer
[sheet-sparkline]: https://docs.univer.ai/zh-CN/guides/sheets/features/sparkline
[sheet-outline]: https://docs.univer.ai/zh-CN/guides/sheets/features/outline
[sheet-print]: https://docs.univer.ai/zh-CN/guides/sheets/features/print
[sheet-io]: https://docs.univer.ai/zh-CN/guides/sheets/features/import-export
[sheet-permission]: https://docs.univer.ai/zh-CN/guides/sheets/features/core/permission
[sheet-collaboration]: https://docs.univer.ai/zh-CN/guides/sheets/features/collaboration
[sheet-history]: https://docs.univer.ai/zh-CN/guides/sheets/features/edit-history
[doc-core]: https://docs.univer.ai/zh-CN/guides/docs/features/core
[doc-api]: https://docs.univer.ai/zh-CN/guides/docs/features/core/docs-api
[doc-toc]: https://docs.univer.ai/zh-CN/guides/docs/features/table-of-contents
[doc-list]: https://docs.univer.ai/zh-CN/guides/docs/features/list
[doc-table]: https://docs.univer.ai/zh-CN/guides/docs/features/table
[doc-images]: https://docs.univer.ai/zh-CN/guides/docs/features/images
[doc-links]: https://docs.univer.ai/zh-CN/guides/docs/features/hyper-link
[doc-find]: https://docs.univer.ai/zh-CN/guides/docs/features/find-replace
[doc-comments]: https://docs.univer.ai/zh-CN/guides/docs/features/comments
[doc-controls]: https://docs.univer.ai/zh-CN/guides/docs/features/content-controls
[doc-footnotes]: https://docs.univer.ai/zh-CN/guides/docs/features/footnotes
[doc-code]: https://docs.univer.ai/zh-CN/guides/docs/features/code
[doc-columns]: https://docs.univer.ai/zh-CN/guides/docs/features/columns
[doc-latex]: https://docs.univer.ai/zh-CN/guides/docs/features/latex
[doc-quote]: https://docs.univer.ai/zh-CN/guides/docs/features/quote
[doc-callout]: https://docs.univer.ai/zh-CN/guides/docs/features/callout
[doc-shapes]: https://docs.univer.ai/zh-CN/guides/docs/features/shape
[doc-watermark]: https://docs.univer.ai/zh-CN/guides/docs/features/watermark
[doc-print]: https://docs.univer.ai/zh-CN/guides/docs/features/print
[doc-io]: https://docs.univer.ai/zh-CN/guides/docs/features/import-export
[doc-collaboration]: https://docs.univer.ai/zh-CN/guides/docs/features/collaboration
[doc-history]: https://docs.univer.ai/zh-CN/guides/docs/features/edit-history

