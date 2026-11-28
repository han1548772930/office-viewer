# 完整功能与 Pro 对齐清单

基线日期：2026-11-21
源码：`dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`
公开规格：当日官方 Sheets / Docs 功能导航、功能说明、API 与产品对照页。  
执行依据：[项目标准](../PROJECT_EXECUTION_STANDARD.md)和[源码布局规则](UNIVER_REFERENCE_RULES.md)。

技术路径遵守[Canvas/WebGPU 与 Rust/WASM 计算标准](RENDERING_AND_COMPUTE_STANDARD.md)。WebGPU 为优先绘制后端，Canvas2D 为功能一致的回退；大型计算优先交给 Worker 中的 Rust/WASM。

全部功能组的责任模块、跨产品共享引擎和实现依赖见 [完整工作台架构](../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)。库调用、主线程 WASM DOM 适配和零拷贝边界见 [WASM 平台标准](RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)。这些是设计与实施契约，不能作为产品功能完成证据。

## 1. 范围与状态

用户目标是 Sheets / Docs 工作台的完整功能与界面对齐，包括 Pro。Pro 仅描述上游发行方式，不改变本项目的交付等级；没有公开实现的部分按公开规格独立实现。这里不把 Slides、Bases、Boards 或独立 AI SDK 整个产品线自动加入本次 Sheets / Docs 范围。

表中来源分为“公开源码”“公开规格”与二者组合。Pro 可以通过公开接口、可访问演示和标准 Office 样本反推可观察行为；记录已核实规格和待核实推断。所有 Sheets / Docs Pro 能力都在范围内，可以分批独立实现。

SHEET-29 和 DOC-22 对齐 [Univer Exchange 的公开行为](UNIVER_REFERENCE_RULES.md#8-rustwasm保真与服务边界)，Office 文件转换由本项目读写链路承担。独立快照转换与协作文件的服务端版本、权限和持久化分别验收。

**当前所有条目均为未验收。** 已有组件草稿和 Rust 解析接口不作为完整验收证据。各项可有部分进度，只有补齐第 5 节的证据后才可标为已验收。

2026-11-28 续接以现有提交及在途代码为基线，继续使用 Sheets / Docs 两个子任务。最新修改与缺口见[本次源码续接记录](feature-evidence/WORKBENCH-CONTINUATION-2026-11-28.md)。其中修复填充后网格缩到目标范围、编辑态填充柄命中、Sheets 初始化依赖和 Docs 缺省 metadata 初始化；追加结构化表格原生范围保真与目录命令闭环。历史段落中的“待确认/未实施”应结合最新代码和本次记录阅读，不能重复把既有字符选择器、页面设置、表格创建或页眉页脚公共接口当作空入口；新旧链路均须实际验收。

同日恢复线程 `01a220b5-b017-7723-b522-04979a516e34` 后，继续补[函数自动引用和一次批量插入](feature-evidence/SHEET-09-function-insertion-2026-11-28.md)、[筛选表头/颜色/合并整组](feature-evidence/SHEET-11-filter-header-color-2026-11-28.md)、[颜色筛选原生 Office 边界](feature-evidence/SHEET-11-29-color-filter-office-2026-11-28.md)及[目录完整布局/单历史、手动更新/删除与公开对话框](feature-evidence/DOC-04-toc-layout-history-2026-11-28.md)。筛选 API 同步检查旧/新范围及相交合并尾部的权限和发布上下文。历史“自动范围/批量插入、目录迭代、flat update/delete 未实施”由这些记录中的新链路替代；Worker 大任务、复杂数组/原生样本及真实运行验收仍保留缺口。唯一函数目录仍为215项/8类，常规Ribbon根增量0；所有65组继续未验收。用户追加的 `pnpm dev` 要求已改为先依次构建 parser/compute/writer/bridge，任一步失败即中止启动；该入口修改批先做静态检查，后续用户构建报错的实际验证另列。

2026-11-29继续补[AND/OR复合筛选、通配符及跨模块Office读写](feature-evidence/SHEET-11-29-compound-filter-office-2026-11-29.md)和[Docs现代标题折叠、Canvas可见投影、输入及历史](feature-evidence/DOC-04-heading-collapse-2026-11-29.md)，日期树沿用原始值选择模型。用户明确要求修复构建报错后，四WASM聚合release/wasm-opt已实际通过且Rust错误/警告均0；Vite重复HistorySavepoint导入已修复并经实际OXC/现有dev模块HTTP200核对，见[COMMON-11定向例外](feature-evidence/COMMON-11-wasm-build-and-vite-import-2026-11-29.md)。测试、类型、前端完整构建、浏览器、Office与性能仍后置；入口和函数数量不因这些子行为增加，全部65组保持未验收。

清单验收状态仅使用“未验收 / 已验收”；开发进度另用“待核实、待开发、开发中、待验证、待修复”记录。未提供完整证据时保持未验收；永久禁用、按钮提示、静态样例、仅保存 JSON 不能记为功能完成。

当前顶层目录是 COMMON 12、SHEET 30、DOC 23，共 65 个功能组。详细行为、公式函数、图表类型和界面入口分别展开统计；65 不是功能细项总数。v1.6.0 OSS 常规 Ribbon 根声明 101（Sheets 63、Docs 38）、Docs 上下文根声明 7（页眉页脚 5、目录 2）、Pro 图表目录选择项 45 均有各自口径，不能相加成为功能总数。缓存 60 个 Pro ES 包是研究覆盖范围，不能代替 Pro 入口清点或功能完成数量；未解析的动态入口继续登记为待核实。

## 2. 共享工程与工作台

| ID | 能力 | 公开依据 | 本项目交付与验收要点 | 当前 |
| --- | --- | --- | --- | --- |
| COMMON-01 | 应用壳与工作台区域 | 公开源码：ui/workbench、parts | 文件操作、类型切换、Ribbon、产品 Header、编辑区、面板、单条产品 Footer 和统一浮层；无重复工具区 | 未验收；[工作台 Grid/Classic 切换验证](UNIVER_TOOLBAR_PARITY.md#1-当前装配与证据边界)、[生产历史接线及 Form/反馈迁移静态记录](feature-evidence/COMMON-03-input-selection-axis-history.md#7-ui-与性能边界)、[顶部控件尺寸、箭头与 hover 修正](UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing) |
| COMMON-02 | 命令、菜单和插件 | 公开源码：core/command、ui/menu、presets | command/mutation/operation 分层；各入口同命令；按 feature 注册、排序、上下文显隐/禁用、正确销毁 | 未验收；[共享输入/轴命令与历史静态接线](feature-evidence/COMMON-03-input-selection-axis-history.md)、[Ribbon 功能搜索](UNIVER_TOOLBAR_PARITY.md#top-settings-and-feature-search)、[Sheets 右键目标、搜索参数与权限](feature-evidence/COMMON-02-sheet-context-feature-search.md) |
| COMMON-03 | 选择、输入、快捷键与剪贴板 | 公开源码：Sheets/Docs selection、editing、clipboard | 中文输入法、光标/选区保留、系统剪贴板；Docs 源格式/目标格式/纯文本粘贴及最近粘贴切换；工具栏和右键不破坏范围 | 未验收；[完整选区、body/story 输入及未运行草稿](feature-evidence/COMMON-03-input-selection-axis-history.md)、[Docs 富格式剪贴板和三种粘贴模式](feature-evidence/COMMON-03-docs-clipboard.md)、[快捷键面板与真实编辑器清单](UNIVER_TOOLBAR_PARITY.md#top-settings-and-feature-search) |
| COMMON-04 | 主题、语言、字体与配置 | 公开源码 ui/config、examples/workbench-settings；[主题规格][shared-themes] | 中文默认、多语言资源、明暗主题、字体加载/度量、LTR/RTL、布局变体与配置保存；不污染文档内容 | 未验收；[语言、方向和组件 locale](UNIVER_TOOLBAR_PARITY.md#top-settings-and-feature-search)、[顶部控件主题 token 与未运行布局草稿](UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing) |
| COMMON-05 | 统一快照、公开 API 与内部保存 | 公开源码 IWorkbookData、IDocumentData、Facade | 稳定模型版本、增量变更、公共 API 与事件；保存失败/dirty/savepoint 正确；重开保持模型 | 未验收；[数字格式输入](UNIVER_TOOLBAR_PARITY.md#sheet-number-input)、[候选准备、事件重入与内容历史](feature-evidence/COMMON-03-input-selection-axis-history.md#5-历史观察事件与取消) |
| COMMON-06 | 权限与只读 | 公开源码 permission；[公开规格][shared-permissions] | 单元、工作表/范围、段落等权限；查看/编辑/复制/评论/导出/打印分开；服务端与命令入口执行权限 | 未验收 |
| COMMON-07 | 实时多人协作（Pro/Server 目标） | [协作规格][shared-collaboration]、[Sheets][sheet-collaboration]、[Docs][doc-collaboration] | 实际服务端、并发协议、用户与远端范围、确认/重试/去重、断线重连、离线队列与多用户一致性 | 未验收 |
| COMMON-08 | 持久历史版本（Pro/Server 目标） | [共享历史][shared-history]、[Sheets][sheet-history]、[Docs][doc-history] | 版本列表、快照+变更日志、版本内容、对比、恢复和恢复后广播；重启后可查；区别于本地撤销 | 未验收 |
| COMMON-09 | 语义对比与恢复 | [公开对比规格][shared-compare]及历史规格 | 内容/样式/结构差异、范围定位与只读历史视图；恢复形成新版本而非清空日志 | 未验收 |
| COMMON-10 | 跨文档嵌入（Pro 目标） | [公开嵌入规格][shared-embed] | Docs 嵌 Sheets、Sheets 对象/标签嵌 Docs；父子独立模型、焦点/权限/保存/生命周期；明确引用与快照语义 | 未验收 |
| COMMON-11 | Canvas/WebGPU、大文件与 Worker/WASM | 公开源码引擎边界 + 用户技术要求 | Canvas 内容渲染、WebGPU 优先/2D 回退；大计算优先 Rust/WASM；进度/取消/过期结果、稀疏/视口、内存与真实基准 | 未验收；基准方法见 [COMMON-11 基准记录](feature-evidence/COMMON-11-compute-benchmark.md)；[依赖升级、XML/ZIP 适配与获准 wasm32 编译检查](feature-evidence/COMMON-11-wasm-dependency-compatibility.md)不代替浏览器/Office/性能验收 |
| COMMON-12 | 图片/媒体资源管理 | 公开源码 drawing、ImageIo | 资源 ID、二进制/关系、锚点/尺寸/裁剪/旋转/层级、引用计数；复制粘贴和保存后资源不丢失 | 未验收；[Sheets affine/HTML/PNG 与资源保真草稿](feature-evidence/SHEET-19-images.md)、[Docs 图片批量编辑与资源](feature-evidence/DOC-07-image-editing.md#image-batch) |

协作、版本历史、持久评论与权限需要本项目内的真实服务及协议。BroadcastChannel 或多个本地数组只能作为辅助测试，不能替代跨用户服务验收。公开的“collaboration: true”、命令常量或数据类型并不构成能力实现。

2026-11-25 用户要求先完成前端，后端暂缓；ntex、同步作业、协作/权限/持久历史等需求保存在[后端需求](BACKEND_REQUIREMENTS.md)，全部范围和未验收状态保留。

2026-11-26 [绘图工作台接线](feature-evidence/SHEET-21-drawing-workbench-affine.md)扩展 COMMON-02/03/05/06/12：共享 typed 选区、主动空选区与 owner 生命周期、混合对象命令及 Facade、内容/选择历史与前后 footprint 权限、图片作者 HTML 和 PNG 取消/释放均已落盘。后续已接固定移动命令、实际 skeleton 边界、selection generation 和 1000ms 连续快照历史。主任务当前 82 条草稿（初批 53 条）、Sheets 图片批 66 条、Docs 批量图片 57 条均未运行；不据此验收系统剪贴板、Windows IME、两后端、Office 或性能。

## 3. Sheets 全量目标

顶部 Ribbon 根入口、插件扩展、上下文菜单与 Microsoft Excel 补充候选项见[Docs 与 Sheets 顶部功能区对齐清单](feature-evidence/TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)及[Sheets 固定源码入口清点](feature-evidence/SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)。

| ID | 能力 | 公开依据 | 本项目交付与验收要点 | 当前 |
| --- | --- | --- | --- | --- |
| SHEET-01 | 工作簿、工作表生命周期 | 公开源码 sheets、sheets-ui；[核心][sheet-core]、[API][sheet-api] | 新建/复制/删除/重命名/重排、颜色、隐藏/veryHidden、活动表、校验和可撤销事务 | 未验收；[复制资源与命令闭环](feature-evidence/SHEET-01-worksheet-copy.md)、[生命周期／引用／标签交互](feature-evidence/SHEET-01-worksheet-lifecycle.md) |
| SHEET-02 | 范围、选区和网格编辑 | 公开源码 selection/editor；[范围规格][sheet-selection] | 单元格/区域/行列/全选/多选、拖选、填充、合并、F2/直接输入、Enter/Tab/Escape；双向视口处理 | 未验收；[合并、取消合并与分类清除](feature-evidence/SHEET-02-merge-clear.md)、[完整选区与共享单格输入](feature-evidence/COMMON-03-input-selection-axis-history.md#2-sheets-的共享模型与事务)、[四模式自动填充、重填与未运行草稿](feature-evidence/SHEET-02-auto-fill.md)、[四方向共享移位/确认、事件订阅修复与 57 个草稿](feature-evidence/SHEET-02-cell-shift-transactions.md) |
| SHEET-03 | 行列结构与尺寸 | [行列规格][sheet-row-col]、公开源码 commands | 插删/移动/复制/隐藏/显示/自动适应、行高列宽；更新公式与所有范围资源 | 未验收；[轴命令、资源变换与宿主发布保护](feature-evidence/COMMON-03-input-selection-axis-history.md#3-行列命令与-ui)、[公开 h/ah/ia/w/hd 导入、来源变换与 17 个草稿](feature-evidence/SHEET-03-public-axis-import.md) |
| SHEET-04 | 样式、默认样式、范围主题 | 公开源码 style/numfmt；[默认样式][sheet-style]、[范围主题][sheet-range-theme] | 字体、填充、边框、对齐、旋转、换行、缩小适应、继承/主题与混合态；按真实模型生效 | 未验收；[范围主题 32 内置/九区段/命令/Ant 弹窗与 46 个草稿](feature-evidence/SHEET-04-range-themes.md)、[九区段 Ant 编辑/原子保存应用及 45 个新草稿](feature-evidence/SHEET-04-range-theme-editor.md)、[局部移位/剪切规则与 32 个草稿](feature-evidence/SHEET-04-range-theme-structure.md)、[默认组合/目录引用/reset/稀疏复制填充](feature-evidence/SHEET-04-default-style-inheritance.md)、[公开样式短字段读回与 13 个草稿](feature-evidence/SHEET-04-public-style-readback.md)、[字体与字号](feature-evidence/SHEET-04-font-family-size.md)、[富文本字体与 Canvas 编辑](feature-evidence/SHEET-04-05-rich-text-font-editing.md)、[颜色与重置](feature-evidence/SHEET-04-color-reset.md)、[边框](feature-evidence/SHEET-04-borders.md)、[当前边框预设图标与弹层尺寸](feature-evidence/SHEET-RIBBON-custom-popup-sizing.md)、[布局六根清点与主格翻转](feature-evidence/SHEET-04-text-rotation-shrink-to-fit.md#2026-11-27-布局组续核)、[分类清除](feature-evidence/SHEET-02-merge-clear.md)、[格式刷与规则模型](feature-evidence/SHEET-04-format-painter.md) |
| SHEET-05 | 富文本 | [富文本规格][sheet-rich-text]、公开源码 docs/editor bridge | 单格多运行样式、部分文本选择和编辑、复制粘贴、链接与导出；不扁平化成一个字符串样式 | 未验收；[范围轨迹与内容历史](feature-evidence/COMMON-03-input-selection-axis-history.md#2-sheets-的共享模型与事务)、[局部字体、typing、API 与 Canvas 选区](feature-evidence/SHEET-04-05-rich-text-font-editing.md) |
| SHEET-06 | 数字格式与日期系统 | [数字格式][sheet-numfmt]、公开源码 sheets-numfmt | 数字/货币/会计/百分比/日期/自定义、小数位、locale、1900/1904；原始值与显示值分离 | 未验收；[入口与编辑语义](UNIVER_TOOLBAR_PARITY.md#sheet-number-input)、[区域货币、会计与格式预览](feature-evidence/SHEET-06-regional-currency.md) |
| SHEET-07 | 冻结、网格线、视口和缩放 | [冻结][sheet-freeze]、[网格线][sheet-gridlines]、公开源码 render | 多分区真实冻结、坐标命中/编辑一致、每表视口保存、组合底栏缩放 | 未验收 |
| SHEET-08 | 复制粘贴与选择性粘贴 | [剪贴板][sheet-clipboard]、公开源码 clipboard | 值/公式/格式/列宽/排除边框、剪切移动、外部 HTML/TSV、引用调整、合并冲突与权限 | 未验收；[格式粘贴目标字体与真实内容事件](feature-evidence/SHEET-08-clipboard-ui-2026-11-22.md#sheets-format-paste-run-style-review)已落盘，权限、历史与内部保存草稿未运行 |
| SHEET-09 | 公式与定义名称 | [公式][sheet-formula]、[命名][sheet-names]、公开源码 engine-formula/sheets-formula | 词法/AST/求值、函数目录、相对绝对/跨表引用、名称作用域、依赖与重算、错误传播和循环依赖 | 未验收；[公式入口、208 项静态目录与当前缺口](feature-evidence/SHEET-09-formulas-entry-review.md)、[单格/数值文本/UTF-16](feature-evidence/SHEET-09-formulas-entry-review.md#scalar-input-continuation)、[五个目录补项及 VALUE 转换](feature-evidence/SHEET-09-formulas-entry-review.md#value-input-continuation)、[六个三角标量分支](feature-evidence/SHEET-09-formulas-entry-review.md#trigonometric-scalar-continuation)、[常用七项投影与真实分类](feature-evidence/SHEET-09-formulas-entry-review.md#common-function-projection)、[PI 与角度转换](feature-evidence/SHEET-09-formulas-entry-review.md#pi-angle-conversion-continuation)、[八个反三角标量分支](feature-evidence/SHEET-09-formulas-entry-review.md#inverse-trigonometric-scalar-continuation)、[双曲/指数对数/ATAN2](feature-evidence/SHEET-09-formulas-entry-review.md#hyperbolic-log-atan2-scalar-continuation)、[奇偶取整/整商/SQRTPI](feature-evidence/SHEET-09-formulas-entry-review.md#even-odd-quotient-sqrtpi-scalar-continuation)、[阶乘/组合数](feature-evidence/SHEET-09-formulas-entry-review.md#factorial-combination-scalar-continuation)、[SUMX/SUMPRODUCT/SERIESSUM](feature-evidence/SHEET-09-formulas-entry-review.md#pairwise-sumproduct-seriessum-continuation)、[MDETERM](feature-evidence/SHEET-09-formulas-entry-review.md#mdeterm-matrix-continuation)、[ERROR.TYPE/N/TYPE/ISREF](feature-evidence/SHEET-09-formulas-entry-review.md#information-types-reference-continuation)、[奇偶/公式身份/工作表信息](feature-evidence/SHEET-09-formulas-entry-review.md#information-parity-sheet-continuation)、[邮件/网址](feature-evidence/SHEET-09-formulas-entry-review.md#information-email-url-continuation)、[ISBETWEEN](feature-evidence/SHEET-09-formulas-entry-review.md#information-interval-continuation)；完整目录、自动范围/批量插入和迭代计算仍有缺口，草稿未运行 |
| SHEET-10 | 高级公式（Pro 增强目标） | [高级公式规格][sheet-advanced-formula] | 按公开函数/行为清单实现高级求值、依赖和性能；补缓存/数组/引用等公开场景，不能以 SUM 示例代表引擎 | 未验收 |
| SHEET-11 | 筛选 | [筛选][sheet-filter]、公开源码 filter | 值/文本/数值/日期/颜色等公开条件、筛选结果与可见行、恢复、结构编辑后的范围变换；不删除源数据 | 未验收；[固定入口、重应用与主题](feature-evidence/SHEET-11-12-sort-filter-entry-review.md)、[连续区域与原子命令](feature-evidence/SHEET-11-12-continuous-range-2026-11-28.md)、[表头/颜色](feature-evidence/SHEET-11-filter-header-color-2026-11-28.md)、[AND/OR、通配符及Office边界](feature-evidence/SHEET-11-29-compound-filter-office-2026-11-29.md)已落盘；日期树、面板与IO草稿均未运行，原生复杂谓词与大型Worker仍有缺口 |
| SHEET-12 | 排序 | [排序][sheet-sort]、公开源码 sort | 多关键字、类型/locale 比较、范围与整表、关联单元格/样式/引用、撤销；不只排序显示字符串 | 未验收；[活动列排序与原来源](feature-evidence/SHEET-11-12-sort-filter-entry-review.md)、[折叠大纲行保留原位](feature-evidence/SHEET-11-12-sort-filter-entry-review.md#sort-collapsed-outline-rows-review)；[连续区域、保持/扩展/取消与异步 owner](feature-evidence/SHEET-11-12-continuous-range-2026-11-28.md)已接面板，全部草稿未运行 |
| SHEET-13 | 数据验证 | [验证][sheet-validation]、公开源码 data-validation | 全部公开规则、范围/公式依赖、下拉、提示与无效状态；键入/粘贴/填充/API 均受验证 | 未验收；[数据验证切片证据](feature-evidence/SHEET-13-data-validation.md)、[列表规则修改保留颜色](feature-evidence/SHEET-13-data-validation.md#list-color-preservation-review)、[快捷下拉/日期入口与列表转义](feature-evidence/SHEET-13-quick-dropdown-date.md)、[复选框候选校验与 Canvas 勾选缺口](feature-evidence/SHEET-13-checkbox-entry-review.md)、[多选单元格 CSV 校验与未完成 UI](feature-evidence/SHEET-13-list-multiple-csv.md)；[Canvas 日期控件方案](feature-evidence/SHEET-13-date-picker-entry-plan.md)待确认 |
| SHEET-14 | 条件格式 | [条件格式][sheet-conditional]、公开源码 conditional-formatting | 规则/范围/优先级、公式、高亮/色阶/数据条/图标集、动态重算、样式叠加；不直接写死底色 | 未验收；[规则编辑与格式刷](feature-evidence/SHEET-04-format-painter.md)、[Office 来源、整本预算与往返草稿](feature-evidence/SHEET-14-conditional-format-office-roundtrip.md)、[快捷规则、十项管理菜单与整表清除](feature-evidence/SHEET-14-quick-formatting.md)；[工作表/选区规则列表方案](feature-evidence/SHEET-14-quick-formatting.md#conditional-scope-list-review)待确认 |
| SHEET-15 | 查找替换 | [查找替换][sheet-find]、公开源码 find-replace | 值/公式两种模式、格式化显示样本与公开范围/方向/匹配选项、结果定位、全部替换与历史；大文件可取消 | 未验收；[单次替换的方向与完整工作表顺序](feature-evidence/SHEET-15-find-replace.md#single-replace-target-order-review)、[Ant Form 与主题](feature-evidence/SHEET-15-find-replace.md#find-form-review)已接线，原契约与 classic/grid 历史/重开草稿未运行；value 模式的格式化文本样本仍待核实，完整结果交互与大文件取消仍缺 |
| SHEET-16 | 超链接 | [超链接][sheet-links]、公开源码 hyper-link | 普通/富文本链接、URL/内部范围/工作表目标、编辑/删除、引用变换和安全跳转 | 未验收；[活动范围的可用性与编辑目标](feature-evidence/SHEET-16-cell-hyperlink.md#hyperlink-active-selection-review)已接线，固定首范围禁用与最后选区写入差异分别记录，classic/grid 历史/内部重开草稿未运行；局部富文本、多链接、tooltip、Office 与权限样本仍缺 |
| SHEET-17 | 批注与线程评论 | [批注][sheet-notes]、[评论][sheet-comments]、公开源码 note/thread-comment | 普通 Note 与 ThreadComment 分模型；锚点随行列变化、回复/解决、用户与持久同步 | 未验收 |
| SHEET-18 | 结构化表格 | [表格][sheet-table]、公开源码 table | 表范围/列/表头/汇总行、样式、自动扩展、筛选排序联动、结构化引用、转换回范围及导出 | 未验收；[固定根与 selector 来源](feature-evidence/SHEET-18-table-selector-entry-review.md)；现有 selector 已接 Studio 创建命令，本次补初始化依赖与挂载草稿；[原生表/列身份、公式/扩展、筛选与排序范围](feature-evidence/SHEET-18-29-native-table-identity-sort-2026-11-28.md)继续补齐，全部运行验证后置 |
| SHEET-19 | 图片与水印 | [图片][sheet-images]、[水印][sheet-watermark]、公开源码 drawing/watermark | 插入/替换/移动/裁剪/缩放、锚点、文字/图片水印；打印及文件写回资源保真 | 未验收；[图片/背景入口、affine 剪贴板与 66 条草稿](feature-evidence/SHEET-19-images.md)、[最终尺寸/filter/auto-height 投影与共享工作台](feature-evidence/SHEET-21-drawing-workbench-affine.md#model-lifecycle)；水印、Office 和真实性能仍有缺口 |
| SHEET-20 | 十字高亮 | [公开规格][sheet-crosshair]、公开源码 crosshair-highlight | 活动行列提示、颜色配置、冻结/滚动兼容；纯视图 operation 不改文档内容 | 未验收 |
| SHEET-21 | 形状及绘图（含 Pro 目标） | [形状规格][sheet-shapes]、公开 drawing 基础 | 几何、箭头/连接线、填充/边框/文本、移动/调整/旋转/层级/组合、命中与 Office 映射 | 未验收；[组合/混合层级与 45 条草稿](feature-evidence/SHEET-21-drawing-groups-and-mixed-order.md)、[共享选区/对齐/移动/Canvas/Facade、作者 affine 与 82 条草稿](feature-evidence/SHEET-21-drawing-workbench-affine.md)；[slicer/timeline affine、五类叶/组和 Ant 控件、98/182 草稿](feature-evidence/SHEET-25-slicer-timeline-affine.md)已落盘但未运行；完整 shape 作者、组剪贴板、跨 workbook 依赖和 Office 仍缺 |
| SHEET-22 | 可编辑图表（Pro 目标） | [图表规格][sheet-charts] | 图表类型目录、系列/分类/区域绑定、轴/图例/标签/主题、数据变更刷新、属性面板、可编辑导出 | 未验收；[图型与入口审计](feature-evidence/SHEET-22-chart-type-audit.md)、[Relation/词云资源](feature-evidence/SHEET-22-relation-word-cloud-resources.md)、[图例选择/换行](feature-evidence/SHEET-22-legend-selection-wrap.md)、[主/副标题配置与几何](feature-evidence/SHEET-22-title-subtitle-contract.md)、[层级图 Ant Tooltip](feature-evidence/SHEET-22-hierarchy-ant-tooltip.md)、[四个对象上下文根、复制粘贴/PNG/删除与 34 个草稿](feature-evidence/SHEET-22-chart-object-context-actions.md)、[真实 Drawer 路由/草稿保护与 33 个草稿](feature-evidence/SHEET-22-chart-edit-panel-routing.md)、[Line/Smooth/Step、二维/3D 与 34 个新草稿](feature-evidence/SHEET-22-line-and-area-styles.md) |
| SHEET-23 | 数据透视表（Pro 目标） | [透视表规格][sheet-pivot]、固定 v1.6.0 发布 UI/类型 | 数据源/cache、字段行列筛选值、公开汇总方式、小计总计、排序筛选、日期/数字/手工分组、下钻、刷新 | 未验收；[15 种值显示方式共享引擎](feature-evidence/SHEET-23-24-pivot-show-data-as.md)及 [命令/Facade/关联图与格式投影](feature-evidence/SHEET-23-24-pivot-show-data-as-facade.md)已落盘；完整字段 UI、展开折叠/日期条件、Office Cache/Parts 与真实 UI/大样本仍待闭环 |
| SHEET-24 | 透视图（Pro 目标） | [透视图规格][sheet-pivot-chart]、固定 v1.6.0 Range/PivotTable source/types/UI ES | 依赖透视模型和图表模型、字段联动、筛选刷新、编辑与持久化；不生成独立静态图 | 未验收；[独立/关联创建、一次配置与 Facade](feature-evidence/SHEET-24-owned-pivot-charts.md)、[字段引擎](feature-evidence/SHEET-24-pivot-chart-engine.md)、[创建/四字段区 UI](feature-evidence/SHEET-24-pivot-chart-ui.md)和 [15 项值显示方式](feature-evidence/SHEET-23-24-pivot-show-data-as-facade.md)已落盘，全部草稿未运行；同时创建表图/展开折叠/完整 Pro API、Office 与真实 UI/性能继续实现验收 |
| SHEET-25 | 切片器和时间线（Pro 目标） | [公开规格][sheet-slicer] | 表格/透视目标绑定、多选/清除、交叉筛选、日期粒度与时间范围、样式/布局及保存 | 未验收；[affine/Ant 画布控件与原 98/182 草稿](feature-evidence/SHEET-25-slicer-timeline-affine.md)、[共享多目标连接、四格式根、Facade/历史与新草稿](feature-evidence/SHEET-25-slicer-connections.md)、[连接 UI](feature-evidence/SHEET-25-slicer-connections-ui.md)、[复制/结构生命周期](feature-evidence/SHEET-25-slicer-connections-lifecycle.md)已落盘但未运行；完整交叉筛选成员状态、日期/Pro API、跨 workbook 缺依赖迁移与 Office/真实 UI/性能仍缺 |
| SHEET-26 | 迷你图（Pro 目标） | [公开规格][sheet-sparkline] | 公开图型、源范围、分组/比例/空值/负值样式、源更新与复制移动引用、导出 | 未验收；[设置与分组](feature-evidence/SHEET-26-sparkline-settings-and-groups.md)、[媒体/上下文入口与删除默认范围](feature-evidence/SHEET-26-sparkline-settings-and-groups.md#media-sparkline-entry-review-2026-11-27)、[XLSX 读写](feature-evidence/SHEET-26-sparkline-office-roundtrip.md) |
| SHEET-27 | 大纲分组（Pro 目标） | [分组规格][sheet-outline] | 行列嵌套组、层级、折叠/展开、汇总方向；与筛选/手动隐藏原因组合，不互相覆盖 | 未验收；[分组命令、真实菜单候选/权限和 Mac Shift 快捷键](feature-evidence/SHEET-27-outline.md)、[XLSX 原生投影](feature-evidence/SHEET-27-outline-office-io.md)，本次草稿未运行 |
| SHEET-28 | 打印和 PDF（Pro 目标） | [打印规格][sheet-print] | 纸张/方向/边距/打印区域/缩放适页、分页预览、重复标题、页眉页脚、网格与隐藏处理、真实输出 | 未验收 |
| SHEET-29 | 导入和 Office 导出（Pro/Server 对齐目标） | [导入导出规格][sheet-io]、duke-pages/rxls 的实际契约 | XLS/XLSX/CSV/TSV 对齐公开规格；保留已有 XLSB/ODS 读取；真实 XLSX/CSV/TSV 写出及重导入 | 未验收；[条件格式来源与资源预算、Office 往返草稿](feature-evidence/SHEET-14-conditional-format-office-roundtrip.md)、[原生图表/辅助公式绑定与往返草稿](feature-evidence/SHEET-29-native-charts.md) |
| SHEET-30 | 工作簿/工作表/范围保护 | [权限规格][sheet-permission]、公开源码 permission | 保护粒度、可执行操作、只读反馈与撤销/协作边界；服务端验证，不能仅隐藏菜单 | 未验收；[顶部新建意图、原规则保留与未运行草稿](feature-evidence/SHEET-30-range-protection-toolbar.md#range-protection-toolbar-create-intent-review)，完整输入/绘图/创建权限、多选区与服务/Office仍缺 |

每个模块的“全部公开”指固定基线对应的具体行为/函数/图型目录，不代表可以用一条笼统说明代替细项。开发前把目录展开为可测试子项；未核实的子项记录为待核实，不能自行选择一小部分后将整个模块标完成。

同日绘图接线还涉及 SHEET-03/07/08/22/25/29：最终轴投影在 filter/auto-height 后一次发布、冻结图表作者事件复用 local↔world affine、完整叶顺序决定输入 owner、图片 HTML/PNG 与七项 FWorksheet 绘图 API 已接入。来源、布局、命令、差异与未运行草稿统一见[绘图工作台证据](feature-evidence/SHEET-21-drawing-workbench-affine.md)。固定 move-drawing 与连续微移历史已落盘；[slicer/timeline 与五类叶/组变换、Ant 控件](feature-evidence/SHEET-25-slicer-timeline-affine.md)已补入但未运行；极端边界及真实 UI/文件样本仍未验收，不以入口接线或内部 JSON 关闭这些 ID。

## 4. Docs 全量目标

顶部 Ribbon 根入口、插件扩展、classic/grid 布局及上下文工具栏清点见[Docs 与 Sheets 顶部功能区对齐清单](feature-evidence/TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)。Microsoft 365 补充功能单列在 Univer/Pro 完整闭环之后，不计入当前 Univer 入口数和实现阶段。目录入口和闭环缺口见[DOC-04 目录证据](feature-evidence/DOC-04-table-of-contents.md)。

| ID | 能力 | 公开依据 | 本项目交付与验收要点 | 当前 |
| --- | --- | --- | --- | --- |
| DOC-01 | 文档模型、运行与排版 | [核心][doc-core]、[API][doc-api]、公开源码 docs/engine-render | 文本运行/段落/节、样式/主题、页面尺寸边距、换行/自动分页、光标与布局映射；传统/现代 flavor 分开 | 未验收；[body/story 输入会话与 Facade 索引桥](feature-evidence/COMMON-03-input-selection-axis-history.md#4-docs-的输入会话) |
| DOC-02 | 字符和段落格式 | 核心 API、公开源码 docs-ui/menu | 字体/字号/颜色/高亮/上下标/字符样式、标题级别、对齐/缩进/行距/段前后、混合选区；格式与模型一致 | 未验收；[上下标](feature-evidence/DOC-02-vertical-align.md)、[局部格式、清除与来源保留](feature-evidence/DOC-02-clear-formatting.md)、[输入偏好与取消/历史](feature-evidence/COMMON-03-input-selection-axis-history.md)、[八项命名段落样式、列表身份与整段权限](feature-evidence/DOC-named-paragraph-styles.md)、[段距/缩进与宿主首段状态](feature-evidence/DOC-02-paragraph-spacing-indent.md)、[分隔线真实命令、边框读写与 Canvas](feature-evidence/DOC-02-horizontal-line.md)、[特殊字符固定命令/整值/公共事件](feature-evidence/DOC-03-header-footer-ribbon-sizing.md#doc-special-character)；上下文挂载、结构 gap、非空段落下边框、旧 flat story 和完整 Picker 仍缺，全部草稿未运行 |
| DOC-03 | 分页、分栏、分节与页眉页脚 | 公开源码 header-footer、insert-break、布局引擎 | 连续/下一页/奇偶页等公开分节语义、节级布局、首页/奇偶页、与前节关联、页码与字段更新 | 未验收；[分节符切片证据](feature-evidence/DOC-03-section-breaks.md)、[页眉页脚 Ribbon 表单、箭头和主题修正](feature-evidence/DOC-03-header-footer-ribbon-sizing.md)、[富内容宿主导航与正文范围恢复草稿](feature-evidence/DOC-03-header-footer-ribbon-sizing.md#rich-story-navigation) |
| DOC-04 | 左侧大纲、标题折叠和正文目录 | [目录规格][doc-toc]、公开源码 DocSideMenu/docs-toc/docs-ui heading-collapse service/docs-toc-ui | 大纲导航；现代标题折叠保留子状态、隐藏选区迁回可见标题、历史/保存/同步；正文目录支持整体/页码更新与删除 | 未验收；[目录源码差异](feature-evidence/DOC-04-table-of-contents.md)、[统一插入与对话框](feature-evidence/DOC-04-toc-insertion-command-dialog-2026-11-28.md)、[真实布局/单历史与手动更新删除](feature-evidence/DOC-04-toc-layout-history-2026-11-28.md)、[现代标题折叠与选区历史](feature-evidence/DOC-04-heading-collapse-2026-11-29.md)已接主路径；复杂原生样本、折叠DOCX新flag写出、同步服务及运行验收仍缺 |
| DOC-05 | 列表与增强列表（含 Pro 目标） | [列表规格][doc-list]、公开基础列表 | 有序/无序/复选、嵌套/升降级、编号连续/重启、样式与重排；复制和导出保留语义 | 未验收；[清单列表切片证据](feature-evidence/DOC-05-checklist.md)、[列表切换保留命名样式与独立深度](feature-evidence/DOC-named-paragraph-styles.md)、[原生编号资源映射及未运行草稿](feature-evidence/DOC-22-native-numbering-resources.md) |
| DOC-06 | 表格与增强表格（含 Pro 目标） | [表格规格][doc-table]、公开基础表格 | 行列/单元格模型、插删移动/尺寸、合并拆分、边框背景、文字对齐与整表对齐分别实现；上游整表对齐限传统模式，跨页/嵌套细项待核实并保留目标 | 未验收；[17 根清点、候选事务/续格/富文本粘贴、create-table 顶部/公开命令与默认整页正文宽、未运行草稿](feature-evidence/DOC-06-enhanced-tables.md)、[插表弹窗的实时权限与组合输入禁用](feature-evidence/DOC-06-enhanced-tables.md#table-insert-live-disabled)、[原生表格来源与读写草稿](feature-evidence/DOC-22-native-tables.md) |
| DOC-07 | 图片和文字环绕 | [图片规格][doc-images]、公开源码 docs-drawing | 行内/环绕、锚点、偏移、尺寸/裁剪/旋转/层级、替换及 Office 关系；输入/分页随对象更新 | 未验收；[图片/inline 编辑](feature-evidence/DOC-07-image-editing.md)、[多选/六对齐/两分布/均匀变换/批删、四 Facade 与 57 条新草稿](feature-evidence/DOC-07-image-editing.md#image-batch)；跨页 anchor、混合对象/非均匀变换、Office 和真实 UI/性能仍缺 |
| DOC-08 | 超链接和书签/锚点 | [超链接规格][doc-links]、公开源码 docs-hyper-link | 范围目标、插入/编辑/删除、位置变换、跳转、序列化与文件映射 | 未验收；[三个编辑命令、生成引用链接限制、公开 operation、Canvas 来源 View 与多 run 公开 ID、未运行草稿](feature-evidence/DOC-08-document-hyperlinks.md)、[弹窗组合输入期间保留链接草稿](feature-evidence/DOC-08-document-hyperlinks.md#hyperlink-popup-composition)；信息浮层、浮动/右键挂载、跨文档注册转发/激活、页面定位与独立书签仍缺 |
| DOC-09 | 查找替换 | [查找替换规格][doc-find]、公开源码 docs-find-replace | 遍历与作用域明确、匹配高亮/定位、可撤销替换；正文表格、页眉页脚等覆盖分别记录，不虚报上游一期限制 | 未验收 |
| DOC-10 | 线程评论 | [评论规格][doc-comments]、公开源码 docs-thread-comment | 范围锚点、编辑后的定位、回复/解决、权限、用户与持久化同步；不以本地数组验收多人评论 | 未验收 |
| DOC-11 | 内容控件（Pro 目标） | [内容控件规格][doc-controls] | 富/纯文本、复选、日期、列表、图片、分组/重复节等公开类型；占位、属性、内容锁/控件锁、导出映射 | 未验收；[十类控件与重复节接线](feature-evidence/DOC-11-content-controls.md)、[自定义复选符号显示](feature-evidence/DOC-11-content-controls.md#sdt-custom-checkbox-ribbon)、[日期格式与锁](feature-evidence/DOC-11-content-controls.md#sdt-date-format-lock-state)、[多行/临时属性编辑状态](feature-evidence/DOC-11-content-controls.md#sdt-multiline-temporary-edit-state)、[行为组禁用与外壳锁别名](feature-evidence/DOC-11-content-controls.md#sdt-behavior-edit-state-and-lock-alias)、[自定义 XML 资源/绑定与原生 SDT 边界](feature-evidence/DOC-22-native-custom-xml.md) |
| DOC-12 | 脚注与尾注（Pro 目标） | [公开规格][doc-footnotes] | 上游 API 适用传统分页；引用与注释内容同一历史事务、编号/重排、跳转、分页排版、DOCX 结构保持；不冒称现代模式已支持 | 未验收；[引用 Ribbon、模型、权限与分页](feature-evidence/DOC-12-footnotes-endnotes.md) |
| DOC-13 | 代码块（Pro 目标） | [公开规格][doc-code] | 语言、源文本、语法高亮、块级选择/编辑、复制粘贴及保存；高亮 HTML 不作为唯一模型 | 未验收；[结构块入口/命令/Facade](feature-evidence/DOC-13-14-16-17-structured-blocks.md)、[Canvas 滚动/输入/辅助宿主](feature-evidence/DOC-13-14-16-17-canvas-interactions.md)、[21 项词法/字体/More/桌面移动入口](feature-evidence/DOC-13-14-16-17-code-popmenus.md) |
| DOC-14 | 多栏组（Pro 目标） | [分栏规格][doc-columns] | 块级 2–5 栏、比例/间距/增删、选择/导航和统一命令；区分节级分栏，支持模式/嵌套规则按公开契约核验，不擅自限定现代模式 | 未验收；[组/栏模型与公开偏移](feature-evidence/DOC-13-14-16-17-structured-blocks.md)、[Canvas 调宽/单位/辅助宿主](feature-evidence/DOC-13-14-16-17-canvas-interactions.md)、[段落/表后入口](feature-evidence/DOC-13-14-16-17-code-popmenus.md) |
| DOC-15 | LaTeX（Pro 目标） | [公开规格][doc-latex] | 公式源文本、行内/块级公开形式、编辑/测量/渲染、错误反馈、保存与可行的 Office 映射 | 未验收；[四公开命令、稳定公式身份、Canvas/Ant 编辑、source 范围/资源与 43 个草稿](feature-evidence/DOC-15-latex-editing.md)、[71 模板/11 上下文根、数学布局分支与 47 条新声明（计划 117 用例）](feature-evidence/DOC-15-latex-templates-ribbon.md)；其他段落/浮动/移动入口、完整 TeX/OMML、Worker 和真实 UI/性能仍有缺口 |
| DOC-16 | 引用块（Pro 目标） | [引用规格][doc-quote] | 独立块样式、内部富文本/段落、进入/退出/删除、复制与持久化；不是仅加一个左边框 | 未验收；[结构命令与边框来源](feature-evidence/DOC-13-14-16-17-structured-blocks.md)、[Canvas/辅助布局](feature-evidence/DOC-13-14-16-17-canvas-interactions.md)、[完整结构入口/保内容转换](feature-evidence/DOC-13-14-16-17-code-popmenus.md)、[四项段落分页/Quote keep 与作者恢复、37 个未运行草稿](feature-evidence/DOC-16-quote-pagination.md)；复杂对象/表格分页及 Office 仍有缺口 |
| DOC-17 | 标注块 Callout（Pro 目标） | [公开规格][doc-callout] | 图标/颜色/内容/容器属性、块编辑和布局、选择、撤销与保存 | 未验收；[属性/历史/快照](feature-evidence/DOC-13-14-16-17-structured-blocks.md)、[Canvas/辅助布局](feature-evidence/DOC-13-14-16-17-canvas-interactions.md)、[段落/表后/移动入口与转换](feature-evidence/DOC-13-14-16-17-code-popmenus.md) |
| DOC-18 | 形状、连接器和 SmartArt（Pro 目标） | [形状规格][doc-shapes]、固定提交 README 商业扩展说明 | 对象/文本/布局/锚点/层级和连接关系；SmartArt 依公开图型、布局与样本另展开规格，不假称有源码 | 未验收 |
| DOC-19 | 文档图表（Pro 目标） | [共享图表规格][shared-charts]及官方 Docs 图表 API | 复用可编辑图表模型、内嵌数据/源引用、尺寸/环绕/分页、属性编辑及 Office 图表关系 | 未验收 |
| DOC-20 | 水印 | [水印规格][doc-watermark]、公开源码 watermark | 文本/图片、样式/透明度/定位、页面范围、打印与导出保留 | 未验收 |
| DOC-21 | 打印和 PDF（Pro 目标） | [打印规格][doc-print] | 使用真实排版、页范围/纸张/方向/边距、分页/页眉页脚/图表/脚注正确；真实打印/PDF 输出 | 未验收 |
| DOC-22 | 导入与 DOCX 写出（Pro/Server 对齐目标） | [导入导出规格][doc-io]、本项目 duke-pages 适配 | DOC/DOCX 导入、DOCX 真正写出；保留已有 ODT 读取；样式/节/列表/表格/对象/关系重开验证 | 未验收；[默认样式与历史来源](feature-evidence/DOC-02-clear-formatting.md)、[正文原生字段与段落保留](feature-evidence/DOC-22-native-fields.md)、[原生表格/续格/legacy merge 来源与草稿](feature-evidence/DOC-22-native-tables.md)、[原生编号/媒体依赖映射](feature-evidence/DOC-22-native-numbering-resources.md)、[自定义 XML 原生资源与绑定](feature-evidence/DOC-22-native-custom-xml.md) |
| DOC-23 | 文档统计与排版状态 | 公开源码 DocFooter/DocStatistics/DocLayoutProgress | 文档/选区统计、页信息、缩放、排版任务真实进度；与内容历史、保存状态分开 | 未验收 |

2026-11-26 [LaTeX 编辑批](feature-evidence/DOC-15-latex-editing.md)也改变 COMMON-03/05 与 DOC-01/03/12 的共享边界：公开段落 text 读取真实 dataStream source，页眉页脚 segment ID 复用保存分配与 aliases，named `DOC_LATEX_PLUGIN` 优先并按 live range 重建。它们是已落盘的接口变化，辅助 hosts、锁/历史和重开草稿尚未运行，不能据此验收其他功能组。

同日后续[公式模板批](feature-evidence/DOC-15-latex-templates-ribbon.md)增加 unit/studio 运行时 editor owner、template operation 和本地 Facade bind/insert；[图表线型批](feature-evidence/SHEET-22-line-and-area-styles.md)增加线型/空值枚举及 Builder/FChart set/clear/reset；[范围主题编辑批](feature-evidence/SHEET-04-range-theme-editor.md)扩展既有 set DTO 与 history/host guards。临时导航和草稿不进入内容历史，作者结果继续走各自文档事务；全部接口与跨模块草稿均未运行。

2026-11-26–27 顶部批补齐[共享控件尺寸和 hover 记录](UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing)，以及 [Sheets 边框与切片器样式弹层](feature-evidence/SHEET-RIBBON-custom-popup-sizing.md)的 SHEET-04/06/13/25 证据。该批结束时本地常规 Ribbon 根为 Sheets 70、Docs 54，同口径 HEAD 为 69、53，该批各增加一个快捷键面板根；这些是入口清点，不是功能完成数量。所有该批测试草稿未运行，实际布局、主题、焦点恢复、Windows IME、文件重开和性能均保留未验收。

2026-11-27 前轮续批修正 Grid 主按钮/箭头的对齐与宽度级联、两产品颜色预设内边距和窄弹层等比色块，并保留页眉页脚导航按下前的焦点；公式栏展开按钮随实际栏高伸展，数字格式 classic/grid 分别使用 14/12px 字号。Sheets 边框颜色和线型控件采用同一 small 高度，快捷下拉动作及数字格式菜单约束长标签；Docs 代码块浮动工具栏限制桌面字体/语言按钮及右侧可用宽度，Tooltip 保留完整值。SHEET-11/12 已接真实活动列、筛选条件状态和独立 mutation 派生行重置，DOC-06 已接顶部物理页正文宽及 headless 目标节默认宽；本批根增量为 0。[静态收尾记录](UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing)区分各轮范围，运行验证后置，不提升功能验收状态。

该次续核修正 classic split 单侧边框造成的箭头中心偏移，保留 Grid/More 实际宽度；特殊字符入口接固定公共 ID，空格、NBSP 和完整 UTF-16 字符串进入既有输入事务。重应用筛选改用 operation，大纲详情和设置读取真实候选/权限；Mac 分组快捷键补回 Shift。具体来源、接口、草稿与未完成范围见上述专项证据。该次收尾时本地常规根为 Sheets 70、Docs 54，增量为 0；完整选择器、表格创建、连续区域及所有运行验收继续保留缺口。

专有控件续修已统一两产品颜色箭头的边框/分隔线和禁用文字 token，并将 disabled 传到 Sheets 快捷下拉与 Docs 列表弹层的实际控件；两产品颜色预设改用 Ant Tooltip，支持 hover 和 focus。颜色按钮还受库中对称透明边框影响，具体偏移量保留待实测。沿用既有命令、权限和保存链路，未增加常规根或功能验收项。具体修改与未运行用例见上述专项证据。

Docs 表格覆盖层同步排除禁用 hover/focus 高亮。[完整符号与表情选择器方案](feature-evidence/DOC-03-header-footer-ribbon-sizing.md#doc-character-picker-plan)已具体到文件和行数，尚待较大改动确认；完整目录、19 locale 数据、搜索、family/肤色、持久最近项和 keepOpen/虚拟列表仍未实施。生成数量未知，不能把方案算作完成项。

本次语法、引用、文档链接/锚点与 diff 空白检查已通过，范围和数量记录在[顶部控件静态收尾](UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing)。静态检查不提升功能验收状态，后续改动与运行验证继续单独记录。

专有控件追加小修已完成最后统一静态检查：86 个 TS/JS、27 个 Vue、28 个 CSS/样式块、15 段夹具脚本、1420 处本地引用，以及 30 份 Markdown 的 926 条链接、71 处锚点均通过。空白和冲突标记检查通过；测试、类型、构建、Rust/WASM、浏览器及基准后置，完整功能验收状态不变。

随后 NA/分隔符批补上[NA 信息分类及 RIGHT 零长度分支](feature-evidence/SHEET-09-formulas-entry-review.md)、[数字格式零边框/打开态/Tooltip](feature-evidence/SHEET-RIBBON-custom-popup-sizing.md#number-format-open-state)及 [Docs 分隔符范围状态和辅助宿主执行守卫](feature-evidence/DOC-03-section-breaks.md#doc-view-break-scope)。该批结束时本地常规根为 Sheets **71（32/9/12/12/6）**、Docs **54（35/16/3）**，函数 catalog 当时为 **118 项、9 类**；该批常驻根仅 Sheets +1、函数子项 +1，Docs 与其他入口口径增量为 0。公式 12 根与固定版同数但分类及计算行为不同，[自动引用和批量插入方案](UNIVER_TOOLBAR_PARITY.md#formula-insertion-plan)仍待较大改动确认；完整目录、表格分页符、Docs 页面设置事务及全部运行验收继续保留缺口。前述静态快照不覆盖该批追加，全部草稿未运行，65 组功能均保持未验收。

专有控件续修再补[边框/切片器 Tooltip 的 hover 与 focus](feature-evidence/SHEET-RIBBON-custom-popup-sizing.md)及 [Docs 列表预览固定标记列](feature-evidence/DOC-03-header-footer-ribbon-sizing.md#doc-list-color-controls)，复用原浏览器用例检查实际裁切、名称、提示与 hover 矩形，未运行。切片器主 Ribbon gallery 与设置字段的不同布局和预览已区分，当前本地三列样式弹层未记为主 gallery 对齐；宽度 token 仍有未核实项。本轮专有续修未新增入口、接口或验收项，当前数量保持不变。

本轮限定静态收尾通过：八份生产文件、八份既有草稿的语法/引用，以及七份证据/清单的链接/锚点和文本空白检查。没有执行产品、测试或构建；本轮静态快照与上述历史记录分开，功能验收状态不变。

随后专有控件批补上[切片器/时间线作者样式预览、默认选中、72px 样式块及原始模型边界](feature-evidence/SHEET-RIBBON-custom-popup-sizing.md#slicer-style-preview-continuation)、[快捷下拉预设尺寸](feature-evidence/SHEET-13-quick-dropdown-date.md#快捷预设尺寸复核)和[页眉页脚面板的 Form row 布局](feature-evidence/DOC-03-header-footer-ribbon-sizing.md#页眉页脚面板与插表对话框续核)。四份生产文件、两份既有浏览器草稿属于新批，前述静态结果不代替本批检查；新草稿未运行，常规根、目录和所有功能验收状态保持不变。切片器三列主 gallery 与上游七列仍有差异，页眉页脚 Modal 入口未接入，实际主题/命中/焦点/IME/文件闭环继续未验收。

2026-11-27 前批续核补上[字体参数拒绝顺序与受控 AutoComplete 输入](feature-evidence/SHEET-04-font-family-size.md)、[富内容页眉页脚实际宿主导航](feature-evidence/DOC-03-header-footer-ribbon-sizing.md#rich-story-navigation)和[文本取整/WEEKDAY 语义](feature-evidence/SHEET-09-formulas-entry-review.md#scalar-function-continuation)。八个信息谓词及草稿在续作前已存在，当时目录为 **126 项、9 类**；该批常驻、插件扩展、上下文、溢出根与函数子项增量均为 0，常规根保持 Sheets **71（32/9/12/12/6）**、Docs **54（35/16/3）**。新增回归草稿未运行，全部 65 组仍未验收；完整函数/数组/定义名称、自动范围/批量插入、迭代计算、完整字符选择器、页眉页脚 Modal、真实 Office 和性能缺口继续保留。

前批续修[继承颜色的稀疏显式重置](feature-evidence/SHEET-04-color-reset.md#2026-11-27-继承颜色重置续修)、[正文/flat/rich story 的字符格式组合输入守卫](feature-evidence/DOC-02-clear-formatting.md#home-format-composition)及[公式单格引用/文本转换/UTF-16](feature-evidence/SHEET-09-formulas-entry-review.md#scalar-input-continuation)，随后接回 LEFT/MID/RIGHT/WEEKDAY/VALUE 的五个目录缺项并修正 VALUE 的格式化文本转换。该批唯一函数目录为 **131 项、9 类**（文本 25、日期 27、信息 9）；目录子项 +5，常驻、扩展、上下文和溢出根增量均为 0，Sheets 71 / Docs 54 不变。内容历史、模型/命令 ID 和 Worker/WASM 协议沿用既有链路，草稿未运行，全部 65 组继续未验收。

后续[四个文本函数单格/INDEX 与 TRIM 空白规则](feature-evidence/SHEET-09-formulas-entry-review.md#scalar-text-reference-continuation)小修复用同一求值链路，保留已有错误与 Tab/换行/NBSP，补真实求值和 Facade 重算/历史/快照重开草稿。无目录或入口增量；统一测试、Office 与性能验收继续后置。

同批补齐[Sheets Ctrl/Cmd+Shift+句号/逗号字号快捷键](feature-evidence/SHEET-04-font-family-size.md#2026-11-27-字号快捷键续修)和[Docs Ctrl/Cmd+句号/逗号上下标快捷键](feature-evidence/DOC-02-clear-formatting.md#fixed-vertical-shortcuts)，登记既有 Ribbon 项的快捷键元数据。字符格式继续走活动编辑器、组字状态、权限、历史及保存链路；Docs 内部 toggle 事件的 field 增加两个取值，公共 API、持久模型和 Worker/WASM 协议未变。[该批静态收尾](UNIVER_TOOLBAR_PARITY.md#font-story-formula-static-close)覆盖当时 27 个修改文件，全部草稿未运行，当时 131 项函数、Sheets 71 / Docs 54 常规根及 65 组未验收状态保持。

前批[三角函数标量续接](feature-evidence/SHEET-09-formulas-entry-review.md#trigonometric-scalar-continuation)补 SIN/COS/TAN/COT/SEC/CSC 六个算法分支及六个数学菜单子项。当时唯一目录 **137 项、9 类**，数学 15、统计 16、常用 6；这是修正 common 前的历史快照。实际菜单、公式输入、权限、历史和 native 保存链路沿用原实现，真实求值/Facade/Ribbon/浏览器草稿已补且未运行；Sheets 71 / Docs 54 常规根及全部 65 组未验收状态不变。

同批[Docs 七条快捷键](feature-evidence/DOC-02-clear-formatting.md#home-format-layout-shortcuts)补删除线、四种对齐与两种列表，复用正文及 flat/rich story 输入、命令、权限/锁、历史与保存；本组 12 条绑定均有静态接线，新增 6 个浏览器计划用例、该阶段累计 20 例且运行 0，相同对齐幂等与固定 UNSPECIFIED 的差异保留。[Sheets 前五根](feature-evidence/SHEET-08-clipboard-ui-2026-11-22.md#sheets-start-history-five-roots-review)继续追到命令与历史，[格式刷](feature-evidence/SHEET-04-format-painter.md#sheets-format-painter-history-run-style-review)修复扩大选区的撤销重做和目标 run 字体继承，合同/浏览器新增 4/1 条草稿且未运行；[共享清除派发](feature-evidence/SHEET-02-merge-clear.md#sheets-clear-formatting-dispatch-review)仍缺，格式粘贴 run 字体随后按本批证据修复。内部 Docs toggle 事件增加 strike，公共 API/模型/命令/保存/Worker/WASM 协议及各根入口口径没有新增，全部 65 组继续未验收。

[前批静态收尾](UNIVER_TOOLBAR_PARITY.md#trig-shortcut-painter-static-close)覆盖当时累计 34 个修改文件：23 个代码/草稿文件的 28 段脚本、3 个模板、3 个样式块与 353 处本地引用通过，目录当时137项/9类无重复；十一份 Markdown 链接/锚点、文本与 diff、固定 HEAD 和参考工作树均已检查。该快照不覆盖随后改动，仅为语法、路径与文档证据，未执行产品或运行验收。

常用投影批[常用七项投影与真实分类](feature-evidence/SHEET-09-formulas-entry-review.md#common-function-projection)复用同一函数目录和原 Ant 参数表单：SUM 归数学，AVERAGE/COUNT/COUNTA/MIN/MAX 归统计。当时唯一目录 **137 项、8 类**，数学16、统计21；函数根 **9**，菜单函数出现项 **144**、插入函数子项另计1。菜单重复出现 +7 不计作新算法或唯一功能，Sheets71/Docs54各根口径和65组未验收状态保持；三份既有草稿补真实求值、分类派发和原编辑/历史/native重开，全部未运行。完整目录/数组、自动引用/批量插入、迭代计算与真实运行验收仍有缺口。

同批[NA 参数数目错误](feature-evidence/SHEET-09-formulas-entry-review.md#na-argument-count)恢复固定 #N/A 身份，原算法分支和两份草稿同步；[格式粘贴](feature-evidence/SHEET-08-clipboard-ui-2026-11-22.md#sheets-format-paste-run-style-review)复用 run 字体清除，保留正文/链接/未知字段，并将真实内容事件范围与完整权限范围分开，recent 切换恢复值/公式仍发事件。新增4条合同与1条浏览器草稿、fixture及原历史/native重开均未运行，ONLY_FORMAT 的共用影响已记录。[Docs 显式0与继承方案](feature-evidence/DOC-02-clear-formatting.md#paragraph-alignment-unspecified)涉及23个生产文件、345–580行及220–360行草稿，待较大改动确认；本批只新增证据，未修改模型、命令参数或Office来源协议。所有根口径及65组未验收状态保持。

[本批静态收尾](UNIVER_TOOLBAR_PARITY.md#common-clipboard-static-close)覆盖累计41个修改文件：30个代码/草稿文件的35段脚本、4个模板、4个样式块、537处本地引用和14份mjs语法检查通过；目录137项/8类、常用7项与9个函数根分开核对。十一份Markdown的链接/锚点、文本/diff及固定参考版本检查通过，均未执行产品或运行验收。

前一批续作[PI 与角度转换](feature-evidence/SHEET-09-formulas-entry-review.md#pi-angle-conversion-continuation)和[八个反三角函数](feature-evidence/SHEET-09-formulas-entry-review.md#inverse-trigonometric-scalar-continuation)共增加11个数学标量分支及子项，当时唯一目录148项/8类、数学27、函数根9；函数菜单出现项155、插入函数另1，常用七项仍为重复投影。原 antdv-next 表单/公式输入、权限、历史和保存链路复用，公式/Facade/Ribbon/浏览器草稿未运行。Docs [列表编号上下文](feature-evidence/DOC-02-clear-formatting.md#home-list-numbering-context)修正活动正文/页眉页脚与 composition 守卫，模型/API/协议沿用，24个浏览器计划均未运行。根数量及65组状态保持未验收。

同批 [values / forced-string](feature-evidence/SHEET-08-clipboard-ui-2026-11-22.md#sheets-values-forced-string-review)补显式强制字符串识别及旧公式目标的源正文/富文本与图片清理；原权限、验证、内容事件和历史复用，非公式目标源 run 字体保留，4条合同/1条浏览器草稿均未运行。六个既有三角函数按固定严格单参数返回 #N/A，原算法/目录数量保持。上游完整 rich 分类、非文本 body、数组/spill、Office 与性能仍有缺口。

[前一批统一静态收尾](UNIVER_TOOLBAR_PARITY.md#math-numbering-values-static-close)覆盖当时累计41个修改文件的35段脚本、4个模板、4个样式块及541处本地引用；目录148项/8类、常用投影、函数根9与菜单出现项155分别复核。文档链接/锚点、文本/diff、固定HEAD和参考工作树检查通过，全部运行验证后置。

此前[双曲、指数对数与 ATAN2](feature-evidence/SHEET-09-formulas-entry-review.md#hyperbolic-log-atan2-scalar-continuation)、[EVEN/ODD/QUOTIENT/SQRTPI](feature-evidence/SHEET-09-formulas-entry-review.md#even-odd-quotient-sqrtpi-scalar-continuation)和[FACT/FACTDOUBLE/COMBIN/COMBINA](feature-evidence/SHEET-09-formulas-entry-review.md#factorial-combination-scalar-continuation)共新增20个数学标量分支及目录子项。该批唯一168项/8类、数学47；常用七项、函数根9、菜单出现项175与插入函数子项1分别记录。沿用同一求值器、Ant 表单和编辑/权限/历史/native 保存链路；四份既有草稿补真实求值与依赖重算/历史/重开、名称派发和表单提交/撤销，未运行。完整目录/数组、非有限缓存的 Office/JSON 往返、迭代计算、自动引用/批量插入与性能继续未验收；Sheets71/Docs54根数量及全部65组状态不变。

同批 Docs [缩进/行距/段距](feature-evidence/DOC-02-clear-formatting.md#home-paragraph-indent-spacing)补缩进执行的活动输入/composition 守卫；原 Ant 控件、活动 story、历史和保存复用，新增4例后 Home/story累计28个浏览器计划、运行0，缩进/单位/默认与继承差异仍在。Sheets [公式粘贴/列宽/recent](feature-evidence/SHEET-08-clipboard-ui-2026-11-22.md#sheets-special-paste-modes-review)修普通值/空格、cut 引用及隐藏/零宽列；合同55、浏览器26、专项合计115，运行0，默认宽 DTO、权限粒度、完整资源/Office仍有差异或未验收。两子任务无新模型/API/历史或Worker/WASM协议，根增量均为0。

既有[TRUE/FALSE/NOW/TODAY 零参数数目错误](feature-evidence/SHEET-09-formulas-entry-review.md#zero-argument-function-arity)恢复固定 #N/A 与参数错误优先级，正常逻辑值/时钟/日期系统沿用。公式草稿和原 NA Facade 七格编辑/历史/native重开未运行；没有新增算法名称或入口。

Sheets [普通整段字体 run](feature-evidence/SHEET-08-clipboard-ui-2026-11-22.md#sheets-uniform-rich-text-paste-review)沿原 cell 工厂归一完整普通字体，external HTML 继续解析正文并写源 cell 样式，values 保留目标字体/独立链接；局部/链接/未知/换行/publicBody 和 forced-string 沿原行为。新增2 UI/1浏览器草稿，当前57/27/专项118，运行0；前述55/26/115为特殊模式批快照。固定 forced-string 路径差异、非图片 publicBody/完整资源/默认宽/Office/IME/性能仍有缺口，API/历史/协议和根增量0。

Docs [表格工作台会话边界](feature-evidence/DOC-06-enhanced-tables.md#table-workbench-input-session)在已有 table context 和 execute 守卫加入失活/composition，并通过原生命周期版本保证恢复重算。该批草稿补4例，当时共12个声明/18个计划、运行0；原领域命令、权限、历史和保存不变，新增入口为0。

随后[表格属性与 UI 操作](feature-evidence/DOC-06-enhanced-tables.md#table-ui-operation-sessions)补属性打开、create/open-table 的原会话/权限判断和属性四处显式 disabled；浮动偏好只拒绝失活回调，活动只读/IME 切换不写内容历史。既有草稿新增2声明/12计划，当前14声明/30计划、运行0。API/模型/保存/协议及入口增量0；创建/通用右键目标绑定、完整辅助 host、主题/焦点/原生IME/Office/性能仍未验收。

[前批统一静态收尾](UNIVER_TOOLBAR_PARITY.md#math-table-uniform-static-close)覆盖当时累计45个修改文件：39段脚本、5个模板、5个样式块、591处本地引用与15份mjs语法通过；AST确认当时168项/8类、数学47、常用七项、函数根9及175个函数菜单出现项。文档链接/锚点、固定源路径、文本/diff与固定参考版本通过；全部运行验证仍后置，65组继续未验收。

随后[ROUND 系列与 TRUNC](feature-evidence/SHEET-09-formulas-entry-review.md#rounding-scalar-continuation)新增四个数学标量名称，并修原 ROUND 的精度/误差/参数错误；当时唯一172项/8类、数学51、常用七项、函数根9、菜单出现项179与 picker 子项1。三条公式及一条五依赖 Facade 草稿、原 Ribbon/Ant 参数表单循环未运行，原求值/输入/权限/历史/native 保存复用，完整数组/Office/性能未验收。

Docs [Home 字符活动会话](feature-evidence/DOC-02-clear-formatting.md#home-character-input-session)在专用字符禁用值加入失活，并复用既有生命周期版本，10个原字符根和独立能力保持。该批 Home 草稿新增2例，单文件当时29声明/37计划，跨 Home/story 续核累计30；运行0，模型/API/保存/协议及入口增量0。完整 story/锁/权限、主题/焦点/原生IME/Office/性能仍未验收。

Sheets [格式刷输入会话](feature-evidence/SHEET-04-format-painter.md#sheets-format-painter-input-session-review)复用真实编辑/组字状态，阻止 UI 新采样/应用/确认，绕过原 Ribbon 自动提交并保留样本与输入草稿；取消确认补工作簿归属，Ant 确定按钮与关闭焦点同步状态。内部 hook/可选组件属性增量已记录，公开 API/模型/保存/协议及根增量0。合同61、浏览器57、合计118声明、运行0，实际弹窗同时输入/原生IME/Office/性能仍未验收。

[前批统一静态收尾](UNIVER_TOOLBAR_PARITY.md#rounding-character-painter-static-close)覆盖当时累计47个修改文件：18份生产、17份草稿/夹具和12份文档。41段脚本、6个模板、5个样式块、604处本地引用及15份mjs语法通过；AST确认当时172项/8类、数学51、常用七项、函数根9、179个函数菜单出现项及picker子项1，无重名。文档链接/锚点、固定源路径、文本/diff与固定参考版本通过；未执行产品或运行验证，全部65组继续未验收。

继续 [CEILING / FLOOR 七个变体](feature-evidence/SHEET-09-formulas-entry-review.md#ceiling-floor-scalar-continuation)：旧版/MATH/PRECISE/ISO 的步长、模式、默认、零值和错误顺序接原同步引擎，保留固定 epsilon 与非有限结果。该批唯一179项/8类、数学58、常用七项、函数根9、菜单出现项186和picker子项1，新增根0。三条公式、一条七依赖 Facade 与原 Ribbon/Ant 表单循环草稿未运行；模型/命令/权限/历史/native保存沿用，完整数组/自动插入/Office/IME/性能仍未验收。

Docs [Home 混合列表预设](feature-evidence/DOC-02-clear-formatting.md#home-list-preset-mixed-state)新增内部可选 listPresetMixed，修正有序/无序根误激活，保留清单、旧编号、未传字段、类别toggle和预设命令。两生产文件净增4行，新增4契约/4浏览器计划；当前 toolbar-state19声明/21计划、Home30声明/41计划、跨Home/story续核累计34，运行0。API/领域/保存/协议/根增量0，自定义Office编号和完整story/主题/IME/Office/性能未验收。

Sheets [清除格式输入会话](feature-evidence/SHEET-02-merge-clear.md#sheets-clear-formatting-input-session-review)复用 canQuickFormatSelection 与可选内部 clearFormatEditable，第4根及两个真实sheet格式清除UI命令在自动提交前按输入/组字/权限状态禁用和拒绝。两生产文件+4/-1；原命令、模型、过滤/合并/run字体、历史/native保存保持。原草稿各新增1声明，当前合同25、浏览器12声明/预计13用例，运行0；共享clear-formatting、clear-range formatOnly及真实IME/主题/焦点/Office/性能缺口保留，API/领域/协议/根增量0。

[前批统一静态收尾](UNIVER_TOOLBAR_PARITY.md#ceiling-list-clear-static-close)覆盖当时累计49个修改文件：19份生产、18份草稿/夹具、12份文档。44段脚本、6个模板、5个样式块、621处本地引用及16份mjs语法通过；AST确认当时179项/8类、数学58、常用七项、函数根9、186个函数菜单出现项和picker子项1，无重名。草稿声明/计划口径、本地文档链接/锚点、固定源路径、文本/diff与固定版本均复核；全部运行验证后置，65组继续未验收。

续接 [GCD / LCM / MULTINOMIAL / MROUND](feature-evidence/SHEET-09-formulas-entry-review.md#integer-aggregate-mround-continuation)：三个整数聚合复用真实稀疏区域、数字文本和原错误，保留2^53/170边界；MROUND 使用固定原始 Number 转换、单格约束、符号/零步长及 epsilon。唯一目录179→183、数学58→62、函数菜单出现项186→190、含 picker 子项187→191；函数根9和新增根0保持。三条公式、一条四依赖 Facade 及原名称派发/Ant 表单循环草稿未运行；非有限欧几里得终止差异、完整数组/自动插入/Office/IME/性能继续未核实或未验收。

Sheets [合并输入会话](feature-evidence/SHEET-02-merge-clear.md#sheets-merge-input-session-review)新增可选内部 mergeEditable，父菜单/四子项及全部11个 UI 合并命令共用已有 canQuickFormatSelection，输入/组字时在 finishEdit 前拒绝，保留草稿和确认请求；生产 +9/-5，原激活/领域/API/历史/保存保持，原两草稿各增1声明，运行0，入口增量0。

Docs [页面设置权限/草稿/成功关闭](feature-evidence/DOC-03-section-breaks.md#doc-page-settings-ui-transaction)在三份生产 UI 文件补文档 Edit 能力、父层成功发布后关闭及模型 ID/API 切换清理；沿用 Ant Form/Modal，错误反馈改用 Ant Alert 和主题 token。原契约新增5计划例、浏览器新增4计划例，运行0；公开页面设置命令/统一事务、全部节尺寸、表格缩放与同 ID 重开仍有缺口，七文件160–240行方案待确认。三项均不提升65组验收状态。

[前批统一静态收尾](UNIVER_TOOLBAR_PARITY.md#integer-merge-page-static-close)覆盖累计52个修改文件：20份生产、19份草稿/夹具、13份文档；47段脚本、7模板、6样式、632处本地引用和17份mjs语法通过。AST确认183项/8类、数学62、常用七项、函数根9、190个函数菜单出现项和picker子项1，无重名；草稿声明/计划已区分并纠正Docs外层for漏计。文档链接/锚点、固定源路径、文本/diff和参考版本通过，全部运行验证后置，65组继续未验收。

本批 [ABS/SQRT/INT/MOD/POWER 与 SUMSQ](feature-evidence/SHEET-09-formulas-entry-review.md#abs-sqrt-int-mod-power-sumsq-continuation)修正五个原函数并新增平方和；固定转换顺序、INT精度、MOD/POWER边界和引用/标量区别沿原权威引擎。当前184项/8类、数学63、函数根9、菜单出现项191、picker子项1、全部子项192。四条公式、两条Facade与原派发/Ant参数表单循环草稿未运行，完整数组/其他引用/迭代计算/Pro和Office仍未完成或未验收。

Sheets [数字组输入会话](feature-evidence/SHEET-06-regional-currency.md#number-format-input-session-review)复用 canQuickFormatSelection，五根禁用和8个既有定义ID/operation提交前守卫保护草稿，生产 +5/-1；原合同/浏览器各增1声明。Docs [全屏退出与缩放焦点](feature-evidence/DOC-03-section-breaks.md#doc-view-fullscreen-focus)在已有其他根全屏时先退出，并通过原会话守卫取消旧文档缩放焦点，浏览器增3计划；文档级缩放持久化和跨document仍有缺口。入口增量0，领域/API/计算协议无新增。

[本批统一静态收尾](UNIVER_TOOLBAR_PARITY.md#scalar-numfmt-view-static-close)覆盖累计54文件：20生产、20草稿/夹具、14文档；49段脚本、7模板、6样式、639本地引用与18份mjs语法通过。声明/计划与目录重复项已复核，运行测试/类型/构建/浏览器/Rust/WASM/Office/基准全部后置，65组状态不提升。

本批 [SUMX / SUMPRODUCT / SERIESSUM](feature-evidence/SHEET-09-formulas-entry-review.md#pairwise-sumproduct-seriessum-continuation)新增五个数学名称及真实求值；SUMX 保留等数量不同形状、原始 Number 转换和左侧略过顺序，SUMPRODUCT 保留行列一致、后续标量广播及单/多参数错误差异，SERIESSUM 保留分阶段转换、稀疏系数索引与最终 NUM 判定。引用/INDEX、typed text/@/@@@ 与非有限引用复用共享读取。目录189项/8类、数学68、函数根9、菜单出现项196、picker子项1、全部子项197；5公式、2Facade与原名称派发/Ant参数循环草稿全部未运行。

Sheets [边框输入会话](feature-evidence/SHEET-04-borders.md#border-input-session-review)复用 canQuickFormatSelection：父/子控件禁用、绘图焦点隐藏与六个既有定义 ID 在 finishEdit 前拒绝，保持草稿、领域/API/历史与笔状态；原两草稿新增1声明及4计划浏览器场景，全部未运行。Docs [全屏 ownerDocument 生命周期](feature-evidence/DOC-03-section-breaks.md#doc-fullscreen-owner-document)绑定实际宿主、卸载原监听、重激活同步宿主和失活退出，增加2浏览器草稿，公共接口与缩放焦点保持；文档级缩放持久化、真实跨 Document 浏览器行为仍未验收。

Sheets [布局输入会话](feature-evidence/SHEET-04-text-rotation-shrink-to-fit.md#layout-input-session-review)复用 canQuickFormatSelection，五根/15子项禁用，十二个既有布局ID在 finishEdit 前拒绝；原主格、wrap/shrink、模型/API和内容历史保持，旋转合同修正两处旧参数预期。Docs [分隔线 CODE 与输入状态](feature-evidence/DOC-02-horizontal-line.md#doc-horizontal-line-code-input)覆盖完整选区相交并给三个既有ID增加失活/composition守卫；新增仅为内部派生状态，公共模型/协议不变。两项草稿未运行，根增量0。

[本批统一静态收尾](UNIVER_TOOLBAR_PARITY.md#pairwise-border-layout-divider-static-close)覆盖累计60个修改文件：20生产、23草稿/夹具、17文档；54段脚本、7模板、6样式、666处本地引用和21份mjs语法通过。目录189项/8类、数学68、常用七项、函数根9及菜单出现项196分别复核；草稿声明与计划已区分。文档链接/锚点、固定源路径、文本/diff与参考版本通过；全部运行验证后置，65组继续未验收。

续接 [罗马与进制转换](feature-evidence/SHEET-09-formulas-entry-review.md#roman-radix-conversion-continuation)和 [RAND/RANDBETWEEN](feature-evidence/SHEET-09-formulas-entry-review.md#rand-randbetween-continuation)：新增六个数学名称，真实求值、原错误/转换顺序、手动缓存、编辑/历史/native重开与 Ant 参数表单草稿已写，未运行。该批唯一目录195项/8类、数学74、函数根9、菜单出现项202、picker子项1、全部子项203；完整数组/spill、volatile依赖、迭代/批量插入、Office与大样本仍未验收。该批固定数学 map 82项尚缺8项，另列未注册LET，不缩减目标。

Sheets [清除内容输入会话](feature-evidence/SHEET-02-merge-clear.md#sheets-clear-content-input-session-review)扩到全部7个既有清除ID，在 finishEdit 前保留草稿；[保护范围顶部新建](feature-evidence/SHEET-30-range-protection-toolbar.md#range-protection-toolbar-create-intent-review)传原mode:add并保留旧规则。Docs [链接弹窗组合输入](feature-evidence/DOC-08-document-hyperlinks.md#hyperlink-popup-composition)拒绝重开覆盖草稿；原命令/模型/内容历史/保存沿用，无公共协议增量。Ribbon30/30、range-editor15/16、链接浏览器16/24，均为声明/计划且运行0。各根数量不变，全部65组仍未验收。

[转换/随机、清除/保护/链接统一静态收尾](UNIVER_TOOLBAR_PARITY.md#conversion-clear-protection-link-static-close)覆盖累计63个文件（20生产、24草稿/夹具、19文档）：56脚本、7模板、6样式、681处本地引用和22份mjs语法通过，包含新SHEET-30证据。目录、草稿声明/计划、文档链接/锚点、固定源码路径和diff已核对；没有执行运行验证，全部65组继续未验收。

续接 [MDETERM](feature-evidence/SHEET-09-formulas-entry-review.md#mdeterm-matrix-continuation)及[信息类型与引用](feature-evidence/SHEET-09-formulas-entry-review.md#information-types-reference-continuation)，补 RAND 空参数接线。该批目录200项/8类：数学75、信息13，其余类别保持；函数根9、菜单出现项207、picker子项1、全部子项208。当时固定数学82项仍缺7项，信息23项仍缺10项；ERROR.TYPE 数组广播、TYPE 对象/数组、ISREF 越界名称及其他完整目录/Office/性能保留缺口。原公式草稿75/75、calculation23/23、sheet-toolbar14/14，均为声明/计划且运行0，模型/API/历史/协议保持。

Sheets [保护新建的编辑状态](feature-evidence/SHEET-30-range-protection-toolbar.md#range-protection-toolbar-create-intent-review)及[范围主题状态与 Modal 生命周期](feature-evidence/SHEET-04-range-themes.md#range-theme-input-session-review)继续复用原能力和历史；Docs [Code 组合输入与栏数只读](feature-evidence/DOC-13-14-16-17-code-popmenus.md#code-panel-composition-and-column-readonly)和[结构属性 Modal 生命周期](feature-evidence/DOC-13-14-16-17-code-popmenus.md#structured-modal-lifecycle)补实际接线。全部运行验证后置，根增量0；专属权限、真实 IME/主题/动画回焦、文件样本及全部65组仍未验收。

[矩阵/信息、结构/主题静态收尾](UNIVER_TOOLBAR_PARITY.md#mdeterm-information-structured-theme-continuation)覆盖累计70文件（23生产、26草稿/夹具、21文档）：62脚本、10模板、8样式、715处本地引用和24份mjs语法通过。目录200项/8类及草稿声明/计划分别复核，文档/固定源路径/diff通过；检查仅读取语法、引用与文档，全部运行结果为0，完整目标和65组未验收状态保持。

随后[奇偶、公式身份与工作表信息](feature-evidence/SHEET-09-formulas-entry-review.md#information-parity-sheet-continuation)、[邮件/网址](feature-evidence/SHEET-09-formulas-entry-review.md#information-email-url-continuation)及 [ISBETWEEN](feature-evidence/SHEET-09-formulas-entry-review.md#information-interval-continuation)共补八个信息名称；[既有八个信息判断](feature-evidence/SHEET-09-formulas-entry-review.md#information-predicate-arity-text-reference)修正参数错误、文本格式引用与 ISNA 文本行为。当前唯一208项/8类，数学75、信息21；常用七项投影、函数根9、菜单出现项215、picker子项1、全部子项216分别计数，各类新增根0。固定信息23项尚缺CELL/ISDATE，数学82项尚缺七项；完整日期格式身份、数组/spill、共享公式、迭代/批量插入、Pro/Office与大样本继续未完成或未验收。

Sheets [折叠大纲行排序](feature-evidence/SHEET-11-12-sort-filter-entry-review.md#sort-collapsed-outline-rows-review)保留隐藏记录原位置，[列表编辑](feature-evidence/SHEET-13-data-validation.md#list-color-preservation-review)保留原颜色且不继承非列表第二操作数，[单次替换](feature-evidence/SHEET-15-find-replace.md#single-replace-target-order-review)补实际方向和完整工作表顺序。Docs [自定义复选显示](feature-evidence/DOC-11-content-controls.md#sdt-custom-checkbox-ribbon)、[日期格式](feature-evidence/DOC-11-content-controls.md#sdt-date-format-lock-state)、[多行/临时属性](feature-evidence/DOC-11-content-controls.md#sdt-multiline-temporary-edit-state)与[行为组](feature-evidence/DOC-11-content-controls.md#sdt-behavior-edit-state-and-lock-alias)补真实值和编辑状态，保留只读属性查看、内容锁下卸壳与外壳锁拒绝。两产品复用现有 antdv-next 控件、命令/历史/内部保存；全部新增草稿未运行，不提升65组验收状态。条件格式范围列表和前述较大方案继续待确认。

续修[查找表单](feature-evidence/SHEET-15-find-replace.md#find-form-review)将两行面板迁移到 Ant Form，七个原控件使用 FormItem，提交阻止、原 Enter/Escape 和主题 token 保持；[超链接活动范围](feature-evidence/SHEET-16-cell-hyperlink.md#hyperlink-active-selection-review)统一决定可用性、显示文字及写入目标。原 range-editor 草稿为19声明/23计划、查找契约7/7，运行0。Docs [内容控件记录](feature-evidence/DOC-11-content-controls.md)已纠正既有同范围父导航：公开导入/保存保留 univerBodySdtOrder，显式 parentId 权威，唯一来源顺序用于无显式父的同范围选择；不计本轮新增能力。完整父区属性浮层、共享选择 view 权限、真实主题/IME/焦点、Office 和性能仍有缺口或未验收。

[信息/内容控件/排序/查找统一静态收尾](UNIVER_TOOLBAR_PARITY.md#information-sdt-sort-find-continuation)覆盖该批累计81文件（25生产、29草稿/夹具、27文档），68段脚本、11模板、9样式、739处本地引用与27份mjs语法通过；208唯一函数/8类和各草稿声明/计划分别复核。文档链接/锚点、固定源码路径、文本/diff及参考HEAD通过。全部运行验证后置，65组继续未验收；此快照不把文件/菜单数量当作功能完成证明。

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

