# 工作台源码续接：2026-11-28

固定源码：`E:\rust_test\univer`，Univer v1.6.0，`eaf4cfa04fabdbcfc922c692f3450459f8a29546`。恢复线程 `01a22041-ba0f-7a42-865b-ce9e64183005` 及现有提交 `7bdfc93`，复用已有清点和实现。本次按 AGENTS.md 分别使用 Sheets / Docs 两个子任务，主任务维护共享模型、API、IO 和全局记录。全部 65 功能组保持未验收；后端暂缓与完整目标不变。

## 用户报告与本次修改

| 问题或入口 | 已落盘修改 | 证据与剩余验证 |
| --- | --- | --- |
| Sheets 初载 `painterRequest` 尚未初始化 | 格式刷 composable 初始化前移到同步读取它的表格 host watcher 之前，保留即时状态订阅 | `ExcelWorkbench.vue`、表格 selector 真实挂载草稿；尚未浏览器验收 |
| Docs 初载读取缺省 `metadata.documentPageOrientation` 崩溃 | 先归一化文档，按显式方向或真实页几何统一读取；页面设置/节布局使用同一缺省模式 | [DOC-03 初载与方向](DOC-03-page-orientation-initialization-2026-11-28.md) |
| 填充后右侧/下方的空白网格整体消失 | 初始模型未设行列数、视图使用默认值，而填充按 `?? 1` 物化为目标范围。按固定源码统一缺省 1000 行/20 列及 API/parser/view/planner，明确尺寸保持 | [SHEET-02 输入与填充](SHEET-02-edit-fill-guard-2026-11-28.md)、[维度/网格草稿](../../packages/office-studio/tests/sheet-grid-extents.contract.test.mjs)；截图指向绘制边界缩减，不是其他单元格内容被清空 |
| 输入时填充柄无法抓取 | 用户明确要求保留输入中的同指针拖动；撤回曾临时加入的隐藏限制，继续修复可视柄自身的命中及提交后的拖动续接 | 同上；组字、验证拒绝和选区失效继续保留草稿/取消语义，实际 Windows 鼠标/IME 未验收 |
| 自定义排序/筛选的连续区域、取消与上下文 | 接通原工作台传给面板的 owner/主格/权限，复用统一排序和筛选命令；筛选命令增加候选和发布边界检查 | [排序/筛选来源](SHEET-11-12-sort-filter-entry-review.md)；真实面板、取消、事件重入和历史草稿后置 |
| 原生表格移动/换列后的导出范围 | 按列原生 ID 追踪筛选和排序，保留公式/样式/扩展，区别当前汇总行与历史显示标志 | [SHEET-18/29 原生表格](SHEET-18-29-native-table-identity-sort-2026-11-28.md)；XLSX 独立重开与复杂扩展仍未验收 |
| 页眉页脚首个编辑器 owner | 无预建输入会话时也绑定实际工作台，执行阶段保持会话、历史、版本与权限检查 | [DOC-03 初载与方向](DOC-03-page-orientation-initialization-2026-11-28.md)；desktop 固定 operation 控制上下文 Ribbon，移动端 Sidebar 分别记录 |
| 自动/自定义目录插入 | 将组件的构建后粘贴改接统一 `doc.command.insert-table-of-contents`，typed/公开 flat 参数共用正文/模式/选区/锁/历史；自定义不强加标题，失败保留对话框草稿 | [DOC-04 本次插入与对话框](DOC-04-toc-insertion-command-dialog-2026-11-28.md)；完整排版迭代、统一历史分组、真实页码和 Office 验收继续保留 |

所有可替换 UI 继续使用 `antdv-next`，Canvas 填充柄及指针命中属于编辑器内核。没有安装或依赖 Univer 商业 SDK。填充、目录和筛选保持原有内容历史；临时弹窗状态、拖动预览和取消不成为内容编辑。

## 数量与状态

静态读取当前目录：唯一公式函数 **215 项、8 类**，数学 75、统计 26、逻辑 12、查找与引用 11、财务 17、文本 25、日期与时间 27、信息 22；常用七项是已有对象的重复投影，不另计函数。该目录数量来自现有提交，本次未增加公式算法。

本次展开 `outlineItems` 与函数分类投影后，AST 重新清点常规 Ribbon 根声明为 Sheets **72**（开始 32、插入 9、公式 12、数据 13、视图 6）及 Docs **54**（开始 35、插入 16、视图 3）。相同方法读取 `HEAD` 和当前工作树得到同一结果，本次根增量 **0**。此前 Sheets 71 / 数据 12 是历史漏计，不能当作本次新增功能。函数分类投影九根与插入函数子项分别统计；上下文标签、插件扩展、溢出呈现和子菜单不混入常规根。新目录命令复用已有目录入口，不因为新增命令 ID 就新增入口计数。固定版 OSS 的 Sheets 63 / Docs 38 与 Pro 清点口径保持原表，不把本地扩展总数当作功能等价证明。

## 恢复线程后的续批

继续读取 `01a220b5-b017-7723-b522-04979a516e34` 的本地记录及摘要，原子任务句柄不在当前线程，因此依照 AGENTS.md 新建 Sheets / Docs 两个子任务并明确文件归属。复用此前57文件的改动及清点，没有重置工作区或缩减完整目标。

| 功能 ID | 本批真实生产变更 | 未完成与证据 |
| --- | --- | --- |
| SHEET-09 / COMMON-02/03/05 | 固定函数operation与batch command、上方/左方引用、行/列结果位置、相对共享公式展开、完整候选验证/权限/单历史；真实编辑器owner、分类菜单和快速求和；参数Form作为独立picker保留 | [函数插入证据](SHEET-09-function-insertion-2026-11-28.md)。20个主线模型/API声明及浏览器草稿未运行；大任务Worker、完整数组/函数与Office仍缺或未验收 |
| SHEET-11/12 / COMMON-06 | 实际筛选首行各列Ant按钮、冻结/缩放/裁剪及16px文字避让；颜色条件、其他列后的候选、整组合并派生；6个API入口的旧/新范围与merge尾部权限及版本/用户/参数守卫 | [前端筛选证据](SHEET-11-filter-header-color-2026-11-28.md)。子任务新增20个声明，主线另有5个API权限草稿，均未运行；日期树、复合条件、表格专属控件及大计算仍缺 |
| SHEET-11/18/29 | 真实colorFilter/dxf读写、工作表/结构化表完整项目快照、当前表格过滤行写出、主题/调色板依赖指纹、外部颜色修改使旧快照失效、缺资源时保留来源及实际隐藏行；跨ref合并缓存边界同步 | [Office颜色边界](SHEET-11-29-color-filter-office-2026-11-28.md)。Rust草稿未运行；多颜色的标准Office reapply、图案/渐变与未知资源重导出不宣称已验收 |
| DOC-04 / COMMON-03/05/11 | Canvas真实布局/字体版本发布、owner/provider及取消；插入后最多三次更新并amend同一历史；手动全文/仅页码、删除及固定公开对话框operation，独立story输入与typing marks保留 | [Docs目录证据](DOC-04-toc-layout-history-2026-11-28.md)。typed/flat/API/UI与模型历史草稿均未运行；复杂原生目录/页码、标题折叠和Worker排版仍按后续实现推进 |
| COMMON-11 构建入口 | 新增 `build:wasm:all` 串行构建四个cdylib模块，根 `dev` 等全部成功再启动workbench，README同步四个产物目录 | [package.json](../../package.json)、[README](../../README.md)。已读取JSON及脚本依赖并做限定diff检查；没有实际打包或启动 |

当前只读静态快照（在后续标题折叠改动之前）：累计工作树114文件、77段脚本（包含HTML夹具）、6份Vue模板、6份样式、1589处本地静态/类型/动态import与export引用，以及15份Markdown的723个本地链接，错误0。catalog AST重新读取为唯一215项/8类，各分类与上表原基线相同；主线函数模型/API和筛选权限草稿分别10/10/5个test声明。先前compiler-sfc从根直接require不可用，改用已安装Vue的 `vue/compiler-sfc` 完成检查，没有新增依赖。

上述是累计静态范围，包含前轮改动，不能当作本轮新增文件数；不覆盖快照之后的改动。没有运行测试、类型、前端构建、Rust/WASM、dev、浏览器、Office或基准。所有65组均未验收，继续推进后续入口；当前原生多颜色映射诊断及Docs主线程布局等差异均保持明确记录。

用户随后实际执行dev时报告 `wasm-pack` 不可识别。环境只读检查确认Rust和`wasm32-unknown-unknown`目标已存在，而PATH及Cargo bin没有wasm-pack。已在根package.json固定开发依赖`wasm-pack: 0.15.1`，在pnpm-workspace.yaml为该依赖启用必要的安装脚本，并执行`pnpm.cmd install --no-frozen-lockfile`，由pnpm生成锁文件，安装成功。`pnpm.cmd exec wasm-pack --version`实际输出`wasm-pack 0.15.1`，四份构建配置文件的限定diff检查通过。没有全局安装工具、实际打包WASM或启动dev；工具版本确认不作为Rust/WASM产品编译验证。README同步项目内工具的安装方式。

旧记录中的字符选择器、页面设置、表格 selector 和页眉页脚 editor 方案已经有生产实现；仍须按入口追踪权限、命令、模型、历史与重开，不将“已有代码”改成“已验收”。完整函数/数组、复杂 Docs 排版、Pro 深层行为、协作/服务、Office 和性能目标继续推进。

## 检查记录

运行验证继续后置。新增模型/命令/IO 契约和浏览器测试均为未运行草稿；没有执行测试、类型检查、前端构建、Rust/WASM 检查、dev、浏览器、Office 文件验收或性能基准。以下记录只读语法、引用、源码路径、计数、文档链接与 diff 的实际范围，不以这些静态检查提升验收状态。

本次已读取固定版 [Rectangle.equals](../../../univer/packages/core/src/shared/rectangle.ts) 对省略 `rangeType` 与 `RANGE_TYPE.NORMAL` 的等价处理；本地 [sheet-fill-gesture.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-fill-gesture.ts) 按同一语义修复提交后误取消。捕获生命周期同时区分子元素冒泡的释放与画布宿主实际失效；仅指针结束/宿主失活对应的 DOM 异常视为正常取消，其他异常继续传播。

两子任务交接后的最终静态快照共 57 份修改/新增文件（27 份生产、21 份草稿/夹具、9 份文档）。41 份代码/夹具的 41 段脚本、5 份 Vue 模板、1047 处本地 import/export 引用与 16 份 mjs 语法检查通过；仅读取语法和路径，没有导入或执行业务模块。9 份 Markdown 的 413 处本地链接、文本编码、冲突标记和新增文件空白检查通过，git diff 空白检查通过。该快照覆盖本批 DOC-04 扁平参数、末轮续拖与共享历史捕获修正；不替代后续改动的检查。参考仓库 HEAD 与固定提交相符，参考工作树干净。

生产改动文件按职责分组如下，详细行为和草稿见上表的专项记录：

- 共享初始化、默认尺寸与命令边界：[model.ts](../../packages/office-studio/src/domain/model.ts)、[model-utils.ts](../../packages/office-studio/src/api/model-utils.ts)、[office-studio.ts](../../packages/office-studio/src/api/office-studio.ts)、[document-editor-commands.ts](../../packages/office-studio/src/api/document-editor-commands.ts)、[document-toc-commands.ts](../../packages/office-studio/src/api/document-toc-commands.ts)、[index.ts](../../packages/office-studio/src/index.ts)、[workbookNormalizer.ts](../../packages/office-studio/src/parser/workbookNormalizer.ts)、[excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts)。
- Sheets：[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue)、[SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue)、[SheetDataPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataPanel.vue)、[sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts)、[use-sheet-sort-filter.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-sort-filter.ts)、[sheet-auto-fill.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-auto-fill.ts)、[sheet-fill.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-fill.ts)、[sheet-fill-gesture.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-fill-gesture.ts)。
- Docs：[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue)、[DocumentTableOfContentsDialog.vue](../../packages/office-studio/src/components/document/DocumentTableOfContentsDialog.vue)、[document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts)、[document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts)、[content-control-editing.ts](../../packages/office-studio/src/features/document/services/content-control-editing.ts)、[document-toc-command.ts](../../packages/office-studio/src/features/document/toc/document-toc-command.ts)。
- Rust IO：[sheet_table_source.rs](../../crates/office-io-common/src/sheet_table_source.rs)、[table_sidecar.rs](../../crates/office-parser/src/table_sidecar.rs)、[xlsx.rs](../../crates/office-writer/src/xlsx.rs)、[xlsx_table.rs](../../crates/office-writer/src/xlsx_table.rs)、[xlsx_table_sort.rs](../../crates/office-writer/src/xlsx_table_sort.rs)。
