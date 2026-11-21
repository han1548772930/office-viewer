# 续作：工具栏、真实操作与 Office 链路（2026-11-20）

本轮从干净工作区 `f62be17` 继续，按 AGENTS.md 阅读四份执行规范。用户随后明确指出顶部工具栏缺项，并要求 UI 和 Pro 能力对齐 Univer，因此本轮同时推进实际工具栏和编辑闭环。**这些是完整目标中的部分实现；所有完整功能 ID 继续保持未验收。**

本机固定参考仓库为 `E:\rust_test\univer`，固定提交 `6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68` 可直接读取。参考仓库当前检出为 `4767b3b309467c283bbd52b725536ca73861e0fc`，与固定提交不同，因此源码对照通过 `git show <固定提交>:<path>` 完成，避免跟随当前分支漂移。Pro 可观察行为继续按官方公开规格核对，没有把无法核实的商业菜单 ID 写成已知事实。

## Ribbon layout

- 对应 COMMON-01/02/03。参考固定版本 `packages/ui/src/views/components/ribbon/{Ribbon,ToolbarItem,ToolbarButton,RibbonOverflow}.tsx`、`ribbon-menu/ClassicMenu.tsx`、`views/custom-label/CustomLabel.tsx`。
- `WorkbenchRibbon.vue`、`RibbonItem.vue`、`ribbon-schema.ts` 对齐 classic 分类/工具区居中、36/40px 两行、28px 分类和选择器、24px 普通工具按钮、16px 图标、组内及组边距 8px、相邻组 1px 分隔。溢出计算同步包含实际边距和分隔。
- 后续 grid 变体复核发现跨两行大按钮缺少上游 56px 最小宽度，导致组宽测量和溢出边界偏小；`GRID_RIBBON.largeButtonSize` 与 `RibbonGridToolbar.vue` 已补齐该约束，浏览器契约验证无显式宽度的大按钮为 56px。该变体仍为 opt-in，生产 Sheets/Docs 默认 classic。
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

## Chart type coverage

- 对应 SHEET-22、SHEET-24。图表模型及透视图面板现在提供簇状柱形、折线、面积、饼、圆环和堆积柱形六类。`chart-engine.ts` 实现面积零线交叉分段、正负堆积值域/基线及饼/圆环扇区；饼/圆环约束为一个非负且总和为正的序列，堆积柱形要求线性值轴。缺失值不会被折线或面积图跨越连接。
- 图表绘制现在生成分类标签、线性/对数数值刻度、水平网格线及模型中已有的 x/y 轴标题；折线/面积点和分类标签使用一致的分类中心位置。新增渲染契约覆盖这些命令及对数刻度。固定参考包的 `sheet-chart-render-object.config.ts` 只定义通用绘图对象外框配置，未找到 Sheets 图表引擎实现可供逐项比照。
- 普通 XLSX 图表 writer 写出 area、pie、doughnut 及 stacked-column OOXML，并保留既有 bar/line 写出。PivotChart writer 已生成 cache definition/records、pivot table、workbook/worksheet/cache/table relationships、`pivotTableParts` 和 chart `c:pivotSource`。缺失关联 pivot definition 仍会明确返回 `UNSUPPORTED_EXPORT`。
- `office-writer` 当前全量 47/47、`office-parser` 17/17 通过。同一生成 XLSX 的 Writer→Parser 契约验证 PivotCache/PivotTable/PivotChart 部件、关系、来源/输出范围、字段配置和图表关联能恢复到 canonical pivot model。Studio normalizer 定向契约 2/2、类型检查通过。Excel/LibreOffice 实际重开未验收；复杂筛选、分组等配置的导入保真仍有缺口。
- 按本轮安排未运行浏览器测试。轴属性编辑、图例位置完整布局、数据标签、主题编辑和真实 Office 重开仍待验收，SHEET-22/24 继续标为未验收。

## Office IO

- 对应 SHEET-29、DOC-22。本轮独立模块 Worker 复现真实保存的 `RuntimeError: unreachable`，根因是 ZIP 的 `time` feature 在 WASM 写出时读取不可用的宿主 `SystemTime`。
- `office-writer/Cargo.toml` 移除此可选 feature，保留原 OPC/OOXML 写出链路；Cargo 自动更新锁文件。新增固定 DOS 时间戳/字节确定性 Rust 测试和 `office-worker-smoke.spec.ts`，对真正的 DOCX/XLSX 二进制验证 ZIP 签名与重复写出。
- 生产构建已实际执行 DOCX 导入→Canvas 编辑→下载→重导入、XLSX 导入→提交公式栏待编辑内容→保存→重导入，以及真实 compute WASM。此类样本只证明被测字段和动作，不证明复杂节/图形/全部 Pro 对象的保真。

## Toolbar audit and Pro source boundary

- 按固定提交的 OSS 菜单 schema 和当前 Ribbon schema 逐分类清点，Sheets 根入口计数为固定版 START/INSERT/FORMULAS/DATA/VIEW 30/9/12/6/6，本地为 28/8/3/8/4；Docs 固定版 Start/Insert/View 25 项，本地常规 Ribbon 42 项并有 7 个表格上下文项。统计范围是各分类直接注册项，不包括子菜单、动态隐藏入口或运行时上下文，不能把数量差直接等同于能力差。
- Sheets 已确认的源码入口缺口包括公式分类和计算选项、特殊粘贴、格式刷、范围保护、快速数据验证/条件格式、结构化表格选择器、工作表背景图和十字高亮。Docs 已确认目标格式粘贴、权限面板及上下文 Ribbon 缺失。页眉页脚现已具备按节编辑、布局显示和 DOCX story 导入/写出，机器文档验证通过；浏览器和 Word/LibreOffice 重开仍未验收。对应的入口并不自动等于完整功能，仍需接入共享命令、领域状态和持久化。
- 当前轮新增选择性粘贴菜单与内部复制数据的模式应用，含公式引用平移、值/格式/列宽/保留目标边框；新增工作表范围保护面板入口，使用既有范围保护 API。格式刷现在在 mouseup 后应用最终单元格/行列标题选区，并以单次历史事务提交；新增拖选、撤销和重做浏览器用例，尚未运行。范围保护与验证面板已接入 feature-owned Vue I18n 资源。尚未运行浏览器闭环测试，整批完成后再统一运行。
- Docs 页眉/页脚现由 parser 保留 DOCX 节引用、首页不同与奇偶页设置，normalizer 映射到各节模型；Writer→Parser 同一 DOCX 契约覆盖 default/first/even story 和继承。Rust parser 17/17、writer 47/47、Studio 类型检查通过。尚未在 Word/LibreOffice 重开，也未浏览器验收。
- 脚注/尾注已有文档模型、插入引用与编辑/删除/跳转服务、管理面板、Insert Ribbon 入口、单快照历史和 18 locale 资源。DOCX parser 暴露定义及 inline 引用，writer→parser 契约通过。前端服务契约 3/3。真实分页脚注定位、跨页续注版式与浏览器闭环仍未验收。
- Vue I18n catalogs 接入 core/sheets/docs feature locales，并按固定版 19 locales 做目录覆盖契约；生产 Sheets/Docs Ribbon 与已注册 panel keys 对照通过。生成的翻译初稿标记为待人工审核，当前机器翻译内容尚未全量人工校订。
- 公式目录差异与固定版 function-map 逐项比对，并补入多批公式。当前差异表显示 date 24/27、text 26/46、logical 11/21、lookup 7/37、financial 17/55。`TEXTSPLIT` 等数组结果函数因公式值模型缺少数组传播仍明确未实现；完整公式兼容仍未验收。
- 固定源码仓库在 README 中明确区分 Apache-2.0 开源核心与单独发行、授权的商业 Web SDK 插件。已检查本地固定版本仓库，存在通用权限、菜单、工作簿/文档模型等源码，没有 `@univerjs-pro` 商业插件实现源码。本项目不安装商业 SDK。闭源能力按官方公开规格和 API 独立反推，来源包括 [Sheets 切片器与时间线](https://docs.univer.ai/zh-CN/guides/sheets/features/slicer)、[Sheets 透视图](https://docs.univer.ai/zh-CN/guides/sheets/features/pivot-chart)、[Docs LaTeX](https://docs.univer.ai/zh-CN/guides/docs/features/latex) 及官方 Sheets 功能比较页。已实现的切片器/时间线仍缺表格源、多目标连接、完整浮动布局和 Office parts 往返，保持 SHEET-25 未验收。
- 根项计数和源码差异来自固定版本本地源码检查，非 Pro 菜单清单。Pro 插件的界面细节、持久化编码和未公开边界仍须按官方规格、演示与 Office 样本分别记录证据；本轮没有将此类行为标记为已核实。

## Verification boundaries

执行环境：Windows、Node 24.18.1、pnpm 12.4.1、Rust 1.100.0；本轮使用 wasm-pack 0.15.1 标准 release 构建及 wasm-opt，工具放在被忽略的 `node_modules/.cache/office-tools`。解析、计算与写出 WASM 均在本项目生成；没有引入商业 SDK 或外部转换服务。

浏览器使用 headless Edge，计算专项使用 Chrome；Windows 受限进程清理会使 Playwright webServer 收尾滞留，因此最终浏览器命令以允许正常子进程回收的权限运行。compute/docs/charts 等配置新增独立 outputDir，避免并行测试清理 `test-results` 根目录而删除彼此的 trace、下载和报告。

本轮没有进行性能优化验收或像素对比。全部功能 ID 维持未验收；完整图表/透视/切片器、权限、协作、持久版本、打印、Docs 脚注/代码/分栏/LaTeX/引用块/Callout/形状/SmartArt/图表以及其它 Pro 目标均继续保留，不能以当前工具栏已扩充替代其实现和验收。

## 工具栏细节续查（2026-11-21）

- 对照固定提交 `6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68` 复核 `packages/ui/src/views/components/ribbon/RibbonGrid.tsx`。该实现工具区高度为 `calc(4.25rem + 1px)`，网格两行各 24px，行列 gap 为 4px，组内 padding 为 8px，溢出区宽 72px。本项目原参数为 73/28/8/64，现已改为 69/24/4/72，垂直 padding 改为 8px、项间距改为 4px。此前本节曾将生产 Ribbon 记为默认 classic/opt-in，实际 `ExcelWorkbench.vue` 与 `DocumentWorkbench.vue` 都明确传入 `variant="grid"`，现以实际生产入口为准。
- `RibbonGridToolbar.vue`、`WorkbenchRibbon.vue` 的“更多”提示/无障碍名称及 `RibbonItem.vue` 的“混合”占位已改用 Vue I18n 插件 translator 和 `common.more`/`common.mixed` keys。定向 i18n 与 Ribbon schema 契约通过。菜单项的 menuitem ARIA 角色和嵌套左右键行为仍需补验。
- 普通组件控件审计将页眉页脚字段迁移到 `Input.TextArea`，条件格式面板迁移规则 `Button`、`ColorPicker` 和 `Checkbox`，工作表标签与右键命令迁移到 `Button`，Docs 颜色字段迁移到 `ColorPicker`，Docs 缩放滑杆迁移到 `Slider`。画布编辑桥和隐藏文件选择器仍使用原生输入元素。Vue 类型检查通过；右键及标签键盘行为尚待统一浏览器验证。
- 活动单元格选区描边宽度与 Univer 默认 `genNormalSelectionStyle().strokeWidth` 同为 1 CSS px。Canvas 活动框和输入框边缘过去重叠且绘制模型不同，现编辑态隐藏同一单元格的重复 Canvas 边框，输入层外扩半个像素并拥有单一边框；右下角选择柄以覆盖层保留。布局契约覆盖缩放、冻结窗格和合并单元格。
- 多格选区原先只有填充色和活动单元格框，缺少范围外沿。现按可见范围的四条真实外沿绘制，冻结窗格内部不额外补边，柄落在末端选中单元格。选区/编辑专项 8 项契约与 `vue-tsc` 通过。仍未进行浏览器像素或 WebGPU/Canvas2D 对照，截图视觉验收待整批功能稳定后统一运行。
- 本次只读 Ribbon 审计确认 Docs 尚缺上游明确入口或行为：查找替换（DOC-09）、权限管理面板、评论（DOC-10）、正文目录（DOC-04）、单独清除字体颜色/高亮；段落 named styles 与字体/字号选择行为也未对齐。Docs 表格上下文 Ribbon 的单元格/行列对齐、合并拆分、边框底色和尺寸等仍是 DOC-06 缺口。
- Sheets 尚缺上游插件注册的插入图片/背景图、超链接、评论和结构化表格入口；查找缺替换/全部替换（SHEET-15）；公式 UI 仍使用有限本地 catalog，不能代表固定版完整类别与函数集合。Sheets/Docs Ribbon 使用 Grid 并非其本身是错误变体，但须继续对照上游布局和产品上下文，不能仅以入口数量判定 parity。
- `COMMON-01/02/03`、`SHEET-15`、`DOC-02/04/06/09/10` 继续未验收。以上缺口不因入口显示或本轮控件替换而关闭。

## Test runner migration (2026-11-21)

- JavaScript contract/unit tests now use Vitest 5.1.1. The default unit config excludes `.browser.test.mjs` files, which remain under the existing separate component command; Playwright browser suites continue to use their Playwright configs.
- The compute benchmark smoke test is included only by `vitest.benchmark.config.ts` and runs with `pnpm test:benchmark:smoke`, keeping benchmark execution out of the default unit suite. The smoke test passed once after migration. Full test results are recorded only after the final combined run.
- This changes the JavaScript test runner and adds no product feature coverage by itself. Office feature IDs remain subject to their existing evidence and acceptance requirements.
