# 续作：工具栏、真实操作与 Office 链路（2026-11-20）

本轮从干净工作区 `f62be17` 继续，按 AGENTS.md 阅读四份执行规范。用户随后明确指出顶部工具栏缺项，并要求 UI 和 Pro 能力对齐 Univer，因此本轮同时推进实际工具栏和编辑闭环。**这些是完整目标中的部分实现；所有完整功能 ID 继续保持未验收。**

本机不存在原约定的 `D:\rust_test\univer-reference`，同名 E 盘路径也不存在。本轮没有升级参考版本：按既有固定提交 `6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68` 的 GitHub 原始文件核对菜单和 classic 布局，并使用官方公开规格核对 Pro 可观察行为。没有把无法核实的商业菜单 ID 写成已知事实。

## Ribbon layout

- 对应 COMMON-01/02/03。参考固定版本 `packages/ui/src/views/components/ribbon/{Ribbon,ToolbarItem,ToolbarButton,RibbonOverflow}.tsx`、`ribbon-menu/ClassicMenu.tsx`、`views/custom-label/CustomLabel.tsx`。
- `WorkbenchRibbon.vue`、`RibbonItem.vue`、`ribbon-schema.ts` 对齐 classic 分类/工具区居中、36/40px 两行、28px 分类和选择器、24px 普通工具按钮、16px 图标、组内及组边距 8px、相邻组 1px 分隔。溢出计算同步包含实际边距和分隔。
- 修复空扩展槽破坏居中、窄屏“更多功能”弹层越界、Select small 样式覆盖高度。分类、嵌套菜单、下拉及更多菜单继续派发同一 schema 的命令/参数。
- `ribbon-schema.contract.test.mjs` 9/9、`ribbon-component.browser.test.mjs` 9/9；覆盖实际几何、LTR/RTL、210/360/620/1200px 可达性、唯一入口与 Escape 归焦。没有截图或像素比较，不能据此宣称整个工作台像素一致。

## Sheet toolbar

- 对应 SHEET-04/06/09。固定源码依据为 `sheets-ui/src/menu/schema.ts`、`sheets-numfmt-ui/src/menu/schema.ts`、`sheets-formula-ui/src/menu/schema.ts`。
- 开始页新增删除线、字号增减、垂直对齐、8 种边框操作、清除格式、货币/会计/日期/时间等 9 种数字类别及小数位增减；布局和 Canvas 绘制实际使用这些样式。`sheet-toolbar-format.ts`、`sheet-number-format.ts` 与 `excelViewAdapter.ts` 保留原始数值、公式缓存及精确数字格式；数字型文本不被擅自变成数值。
- 公式页新增真实函数选择和参数输入，`sheet-function-catalog.ts` 列出 30 个已用当前引擎求值验证的函数，提交走已有公式输入/内容历史。这个目录不代表完整普通/高级公式目录。
- 插入和数据页均可打开已有图表、透视配置、切片器/时间线；这些 Pro 入口共用各自真实模型、命令和面板。图表当前仅 bar/line，不能据入口存在称为全图型完成。
- 已有专项 `sheet-toolbar-format.contract.test.mjs` 和 `sheet-toolbar.spec.ts` 验证样式、原始类型、数字显示、Canvas、历史和内部快照重开；最终联合结果见收尾验证。
- 仍缺完整自定义数字格式/locale、所有旋转/缩小适应操作、全部函数/高级公式、样式 Office 独立往返和所有跨入口统一权限核验。

## Sort filter

- 对应 SHEET-11/12。固定参考为 `sheets-filter-ui/src/menu/schema.ts` 与 `sheets-sort-ui/src/menu/schema.ts`。
- 新增 `SheetDataPanel.vue`、`sheet-data-panel-state.ts`；顶部“筛选”“自定义排序”打开真实面板。支持按值多选/全不选、搜索/分页、15 种条件、明确的数值/文本/逻辑值类型、清除此列/全部条件和移除区域。不同列条件按 AND 组合，手动隐藏原因独立保留。
- `sheet.command.configure-filter` 将整组范围和列条件作为一次内容事务，校验工作簿 revision，取消/失败/过期不部分写入。排序复用 `FWorksheet.sortRange`，支持多关键字、标题行、比较语言、取消/进度以及已有 Worker/WASM 委派与过期结果检查。
- 修复降序把空白排到前面；公式引用转换在全部成功后才写回，失败不再清空或半写范围。纠正旧回归夹具中的 AND 条件与 payload 所属行预期，并保留/增强全部行为断言。
- 视图新增独立 `filteredRows`，Canvas、命中、统计、名称跳转和键盘导航共同使用筛选+手动隐藏的并集。视图回传不把筛选隐藏混入手动隐藏，并在编辑后的真实单元格上重算条件。
- `sheet-data-panel.contract.test.mjs` 6/6、`sheet-sort-filter-engine.contract.test.mjs` 9/9；另有真实 UI 和 filter/view/历史契约。XLSX 条件写出与重导入的边界单独记录，不能用隐藏行计数代替 Canvas 或文件验证。
- 列头筛选菜单、完整日期/颜色规则、结构变更的全部资源联动和大样本性能仍待验收。

## Calendar input

- 对应 COMMON-05、SHEET-06。`api/model-utils.ts` 修复部分工作簿输入丢失 `date1904`；初始化与 `createWorkbook` 都保留指定日历。
- 两条新增公共 API 契约覆盖日期公式结果、按文档隔离、撤销/重做以及内部快照重开；1904/1900 日历没有因为缺少 `kind` 字段而混淆。

## Slicer timeline

- 对应 SHEET-25。新增 `features/slicer/timeline-date.ts`，面板和引擎共用严格日期转换；处理 1900/1904、serial 0/负数/1–59/60/61，保留 Excel 虚构的 1900-02-29，不把非法 2 月 30 日滚动成另一日期。无时区的 ISO 日期时间按日历值处理。
- 8 条专用契约覆盖多选去重、多字段交叉筛选、空选与清除的区别、删除、历史、保存重开、取消和错误原子性。`feature-panels.spec.ts` 从顶部执行切片器和时间线面板，验证原始透视结果与历史。
- 仍缺结构化表格/多个目标/透视图绑定、画布浮动布局、切片器 XLSX parts 与大样本性能，不能标为完整 Pro 对齐。

## Content controls

- 对应 DOC-11。`content-controls.ts`、`document-facade.ts`、`document-commands.ts` 修复普通回车、CRLF 粘贴、插表/分隔符拆段的锚点迁移，覆盖 UTF-16、表格单元格、父子及空控件。
- 跨越单段控件的拆分在修改前明确失败；变换先校验后提交，文本、锚点和历史不被污染。修复 `controlLocked` 别名、子控件绕过祖先内容锁，以及未变化表单属性妨碍解锁。
- 新增 `content-control-boundaries.contract.test.mjs` 14 条，联合既有控件/文档边界/Facade 共 41/41。真实面板与 Canvas 回车迁移另有浏览器用例。
- 插入工具栏提供当前六种可编辑控件的类型直达，仍不代表图片/分组/重复节/文档部件的完整 UI。跨段 SDT、完整 RichText/重复项结构、导入保真与 DOCX 往返仍未验收。

## Document toolbar

- 对应 COMMON-03、DOC-02/06/11。参考固定 `docs-ui/src/menu/schema.ts`、菜单上下文与公开表格/内容控件规格。
- 开始页补剪切、粘贴及仅文本粘贴；输入桥复用剪贴板归一化，保留模型选区。修复 HTML 文本节点被 marks 中的空 `text` 覆盖而消失的问题，移除剪贴板中非正文的 meta/link 元素。
- 插入页补符号/emoji 实际插入及内容控件类型直达。表格内部出现专用上下文分类，支持前后插行、左右插列、删行/列/表，每次作为一组模型命令进入现有历史。删除前检查所有后代控件锁；对暂不支持的合并/不规则网格组合原子报错，保留完整增强表格目标。
- 当前表格上下文不等同于完整 Pro 增强表格：移动、合并拆分、边框/尺寸、跨页嵌套和完整文件映射仍有缺口。源格式/目标格式切换、富剪贴板对象/媒体及完整 Office 往返仍需继续实现。

## Worksheet viewport

- 对应 SHEET-07。`SpreadsheetCanvas.vue` 修复 KeepAlive 将原生滚动位置归零后覆盖缓存的问题；停用保存逻辑坐标、忽略停用 scroll、激活和切表先恢复两轴再绘制。
- 原工作台浏览器用例增强为 T45，检查水平和垂直位置精确恢复，随后仍可编辑该单元格。

## Office IO

- 对应 SHEET-29、DOC-22。本轮独立模块 Worker 复现真实保存的 `RuntimeError: unreachable`，根因是 ZIP 的 `time` feature 在 WASM 写出时读取不可用的宿主 `SystemTime`。
- `office-writer/Cargo.toml` 移除此可选 feature，保留原 OPC/OOXML 写出链路；Cargo 自动更新锁文件。新增固定 DOS 时间戳/字节确定性 Rust 测试和 `office-worker-smoke.spec.ts`，对真正的 DOCX/XLSX 二进制验证 ZIP 签名与重复写出。
- 生产构建已实际执行 DOCX 导入→Canvas 编辑→下载→重导入、XLSX 导入→提交公式栏待编辑内容→保存→重导入，以及真实 compute WASM。此类样本只证明被测字段和动作，不证明复杂节/图形/全部 Pro 对象的保真。

## Verification boundaries

执行环境：Windows、Node 24.18.1、pnpm 12.4.1、Rust 1.100.0；本轮使用 wasm-pack 0.15.1 标准 release 构建及 wasm-opt，工具放在被忽略的 `node_modules/.cache/office-tools`。解析、计算与写出 WASM 均在本项目生成；没有引入商业 SDK 或外部转换服务。

浏览器使用 headless Edge，计算专项使用 Chrome；Windows 受限进程清理会使 Playwright webServer 收尾滞留，因此最终浏览器命令以允许正常子进程回收的权限运行。compute/docs/charts 等配置新增独立 outputDir，避免并行测试清理 `test-results` 根目录而删除彼此的 trace、下载和报告。

本轮没有进行性能优化验收或像素对比。全部功能 ID 维持未验收；完整图表/透视/切片器、权限、协作、持久版本、打印、Docs 脚注/代码/分栏/LaTeX/引用块/Callout/形状/SmartArt/图表以及其它 Pro 目标均继续保留，不能以当前工具栏已扩充替代其实现和验收。
