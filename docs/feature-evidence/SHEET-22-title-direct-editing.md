# SHEET-22 / SHEET-24：标题真实命中、Ant 浮栏与拖动提交

2026-11-25，接续[标题对象配置与四边绘制](SHEET-22-title-subtitle-contract.md)。本批代码和测试草稿已落盘；所有场景保持 **未验收**。按用户当前“先实现、后统一验证”安排，本批没有执行测试、类型检查、构建、开发服务、浏览器、Rust/WASM、性能基准、`git diff --check` 或自动链接检查。

## 固定来源与分类清点

固定仓库 `E:\rust_test\univer` 为 v1.6.0、commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。以下字符 offset 属于已静态读取的固定 ES 文件，不是新版文档或运行证据：

| 来源 | 已核实的入口、状态和数据契约 |
| --- | --- |
| [chart-ui ES](../../.research.local/univer-1.6.0/packages/chart-ui/lib/es/index.js)，`hs` 91523、`ql` 182316 | title/subtitle hit 映射为 `scope=title` 与独立 `titleKey`；解析到这两个标题目标时显示浮动菜单。 |
| 同文件 `Su` 约 189900–191150、`bu`、`ku` 198850 | 第一行文字；主标题第二行五个位置值；字号（含恢复→null）、颜色、主标题对齐、粗体、斜体。副标题浮栏没有位置/对齐控件，完整属性面板仍有副标题对齐。onChange 调用 adapter.setState 的对象 patch。 |
| 同文件 `ju` 199650；[浮栏类型](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/types.d.ts)、[元素选择类型](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/chart-element-selection.d.ts) | host 提供 view state、尺寸适配、close/openFullPanel 协议。存在 ACTION.MORE 常量，但 `ku` 没有渲染 More 按钮，本批不据常量创建入口。 |
| [engine-chart 归档](../../.research.local/univer-1.6.0/archives/engine-chart-1.6.0.tgz)，`package/lib/es/index.js`：`S5` 1644853、`C5` 1645955、`w5` 1647031 | 主/副 graphic 带 `info.chartElementHit.kind`、silent=false 和 pointer cursor；左右逐字纵排，上下为水平文字；共同 graphic 替代 native title。 |
| 同一 ES：`C5`、`Z6` 约 1532700；归档 `package/lib/types/types.d.ts:642` 的 IChartTitleStyle/ILabelStyle | 水平 manualLayout 使用 left/top × chart width/height，手动文字按 left 对齐，不再用自动 wrap width。类型包含 manualLayout 和继承 rotate，但 C5 rotation 固定为 0，侧向 S5 没有读取它们。 |
| 同一 ES：`g7`/`v7` 1672300–1673150、`x7` 1673950 | 事件提供 hit、真实边界及 localPoint；edit mode 有悬停/固定边框，点击固定选择。整个 ES 中 manualLayout 的 3 个命中都属于读取绘制，未查到标题拖动写入协议。 |

这是上下文标题组的清点，不新增常驻 Ribbon、插件扩展或溢出根项。两个命中目标不是两个按钮，文字行数也不是入口数。固定浮栏共 12 个值/切换控件（主标题 7、副标题 5），另有 2 个字号恢复子动作；它们分别如下，保持分组顺序：

| 目标 | 文字行 | 第二行与字体组 | 显隐/禁用与提交 |
| --- | --- | --- | --- |
| 主标题 | 文字 | 位置 → 字号/恢复 → 颜色 → 对齐 → 粗体 → 斜体 | 命中已有可绘制标题后显示；Hide 移除整个主/副标题块；作者能力受工作表权限和文档版本约束。 |
| 副标题 | 文字 | 字号/恢复 → 颜色 → 粗体 → 斜体 | 独立 subtitle 目标，浮栏无位置/对齐；使用相同权限、版本与作者命令。 |

本地增加的浮栏值控件为主/副各一个旋转输入，共 2 个；自动位置恢复各一个，共 2 个。文字 Apply/关闭是各两个宿主动作，共 4 个；Ant ColorPicker 的颜色清除是 2 个字段子动作，其上游等效入口未核实。属性面板另增 rotate、manual 开关、left、top 四个控件/标题，共 8 个本地属性控件。拖动、双击聚焦和 Shift+Enter 输入是交互行为，单独记录，不计作上游已证实新增 Ribbon 入口；这些数值均不证明功能完成。

## 模块、计算域与边界

对照[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)与[WASM 平台契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)：作者字段属于 charts 领域模型；文字槽/对象 patch/命中多边形属于 charts 的可测试逻辑；ECharts/ZRender 度量属于浏览器 Canvas 适配；浮栏和输入属于 Vue/Ant 视图；内容提交仍由既有 command/service/history 编排。

本批没有新增 Worker/WASM DTO、内存 lease、文件转换或另一套权威计算引擎。小型指针命中留在前端，本批没有端到端基准来宣称它优于 Worker/WASM。大图表准备继续走现有统计与作者事务链路；Office 写回和复杂布局保真仍待后端及文件样本验收。

## 实际前端行为

- [chart-titles.ts](../../packages/office-studio/src/features/charts/chart-titles.ts) 继续保持 title/subtitle 字符串与独立 style；对象 patch 接受 manualLayout、rotate，显式 undefined 恢复对应作者字段。水平手动标题不自动换行、采用左对齐；自动标题仍保留既有有界文字槽，手动条目不占自动标题边带。四边、纵排和多行文字都按一个锚点旋转并裁剪到 chart frame。
- [chart-title-interaction.ts](../../packages/office-studio/src/features/charts/chart-title-interaction.ts) 接收明确的标题元数据和引擎测量结果，裁剪真实旋转多边形、解析 title/subtitle 优先级、生成 Canvas 选中线条和归一化拖动锚点。命中不按显示字符串猜测，也不按字符宽度估计。owner/workbook/sheet/chart/revision/snapshot 共同限制作者会话。
- [echarts-option.ts](../../packages/office-studio/src/features/charts/echarts-option.ts) 给两个标题目标生成稳定 graphic ID、kind/info 和可命中的 graphic；native 内容与 shared DrawCommand 内容均使用它。标题旋转沿用共享渲染协议的顺时针角度；ECharts 的负角度与 ZRender 变换约定一致。旋转标题的 maxWidth 处理与既有 Canvas2D/WebGPU 文本协议一致，仍按 frame 裁剪。
- [echarts-renderer.ts](../../packages/office-studio/src/features/charts/echarts-renderer.ts) 按稳定 ID 从 ZRender 的实际 Text 获取 getBoundingRect/getComputedTransform，随后生成裁剪多边形；标题选择层和悬停层使用相同测量。完整重绘与 shared 图例 rebuild 都重新测量，内容局部重绘保留标题。公开 getTitleHit/getTitleTarget 返回 chart-local CSS 坐标，不包含模型修改。
- [ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) 在标题 → 图例 → Sankey 内容 → 整图移动的顺序命中；单击选择并画 Canvas 边框，双击或选中后 Enter 聚焦标题 Ant 输入。标题拖动超过 4 个屏幕像素才产生暂态 preview；指针松开后将最后一个 manualLayout patch 提交一次，提交期间保留 preview，失败回到作者状态。图表 position/size 不随标题拖动修改。
- [ChartTitleFloatMenu.vue](../../packages/office-studio/src/features/charts/views/ChartTitleFloatMenu.vue) 使用 Ant Form/FormItem、Input.TextArea、InputNumber、ColorPicker、Select、Button、Tooltip、Alert 和宿主 ConfigProvider/theme token。标题字号/颜色/粗斜体由现有字体继承计算；副标题浮栏保持无位置/对齐。Apply、Enter、有效失焦或格式动作合并待提交文字、数值和颜色为一个对象 patch。数值逐键输入只更新暂态值，字号/旋转 Enter 或步进动作才提交；ColorPicker 以 changeComplete 提交最终颜色，避免一次拖色发出半途内容命令。Shift+Enter 保留换行。
- Input 的组合输入期间只保留文字草稿，不提交 Enter，也不因格式按钮生成半个词；compositionend 读 Ant 现有 TextArea 节点的最终值。Esc、换目标、外部指针、滚动、缩放、主题/字体变化、失去窗口焦点、宿主卸载/切换等结束当前选择并丢弃未提交草稿。关闭后只向仍存在的原宿主 chart Button 恢复焦点；已被移除输入的迟到事件捕获旧选择 token，不能修改另一标题或复用的 chart ID。
- 工具栏只读时可查看已有文字/状态，作者控件禁用；权限失效和文档/图表变化取消未提交拖动。已经松开/Apply 的内容请求复用既有准备事务的权限/版本检查，视图关闭不另造一套 Promise 撤销协议。Canvas 完全裁掉已选手动文字时，浮栏沿作者锚点保持自动恢复按钮可访问；这不会在不可见位置生成标题命中。
- 标题删除键走标题组清除，不删除整个图表；空文字/Hide 提交后关闭浮栏。恢复自动位置只清当前 manualLayout，保留文字、字体、旋转与另一标题。

统一作者路径为 `FOffice.getWorkbook(unitId) → FWorkbook.getSheetBySheetId → FWorksheet.getChartById → FChart.update → sheet.command.update-chart`，复用现有准备事务、工作表权限、内容历史和内部 JSON 保存/重开。浮栏、对象 API、属性面板和拖动共用 chartTitlePatch/作者命令；hover/selection 不写 revision/history。没有以提示消息或空操作替代这条路径。

## 接口与文件

- [model.ts](../../packages/office-studio/src/domain/model.ts) 在标题 style 增加 `manualLayout?: { left:number; top:number }` 和 `rotate?:number`；[校验](../../packages/office-studio/src/features/charts/chart-option-validation.ts)要求完整、有限的 0..1 锚点对和 -360..360 角度，非法对象在候选提交前拒绝。
- [chart-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-settings-controls.ts) 增加两标题的旋转/手动位置/坐标与独立恢复；选择主标题预设位置同时清其手动锚点，作为同一次属性 patch。
- [index.ts](../../packages/office-studio/src/index.ts) 补导出 ChartTitleManualLayout、ChartTitleKind、ChartTitleObjectSpec；既有 Builder/live FChart spec 类型自动接受新增字段，live Promise<this> 返回语义保留。本批无需改 Facade 源码、主工作台或 OfficeStudio。
- `chartTitleCommands` 返回带 chart-local 识别元数据的 TextCommand 子类型；共享 DrawCommand、跨 Worker 协议、ChartOverlay 的既有 emit payload 和文件 DTO 没有改动。

本批共涉及 16 个文件：上述 10 个代码文件（含新 interaction helper/Ant 浮栏），下列 4 个测试/fixture 文件，以及本文与[上批证据](SHEET-22-title-subtitle-contract.md)。中央功能清单由主任务协调更新。

## 差异与未核实项

1. 已核实固定标题命中、浮栏、水平 manualLayout 读取。固定 ES 未证明标题拖动写入或 rotate 绘制；本批标题拖动、旋转和侧向 manualLayout 是本地 Office 行为扩展，仍待公开行为/Office 样本核实，不计作已证实上游新入口。
2. 上游文字 onChange 立即 setState；本地采用组合输入/文字编辑会话，再经一次作者命令提交。上游字号恢复传 null；本地传 undefined 恢复继承。TextArea/Shift+Enter 为保留本地多行标题的输入扩展。
3. 本地限定手动锚点为 0..1、拖动归一化并夹紧；固定类型没有声明这一区间。手动水平标题不自动换行会改变多行自动居中标题的对齐/分行，这是明确手动布局语义；模型中的完整文字仍保留。
4. 字体默认值、自动断行/code point 估算、复杂脚本 shaping、grapheme/RTL、轴标题/PlotArea 碰撞延续上批差异。真实命中使用引擎测量不表示整个字体布局或所有碰撞已经验收。
5. 45 图型/全部变体、Docs 嵌入图表、协作 ACL、公共事件的所有交互路由、真实 OS 输入法、触屏/多指、窄屏浮栏/弹层、WebGPU/Canvas2D 的实际输入闭环、Office 导入写回/独立重开均未验收。内部 JSON 草稿不能替代 XLSX/DOCX 保真。
6. 新建标题元素的入口、本批没有核实的 full-panel/More 路由以及其他图表元素的浮栏仍由完整图表清点继续追踪；没有将它们从 Pro 目标删除，也没有把当前局部实现标为完整 SHEET-22 交付。
7. 标题扫描、DrawCommand/metadata、ZRender 度量及多边形有 JS 分配；没有真实大样本、复制/传输、主线程长任务或内存测量，性能状态仍为未验收。

## 测试草稿与依据

新增 16 个场景，均未运行：

| 文件 | 新增数量 | 计划验证的行为 |
| --- | --- | --- |
| [chart-title-interaction.contract.test.mjs](../../packages/office-studio/tests/chart-title-interaction.contract.test.mjs) | 8 | 手动锚点/旋转校验与解除引用、独立恢复、手动/自动边带与完整文字、多行旋转/稳定 ID/native/shared 元数据、真实标签识别、裁剪命中/重叠优先级、Canvas 线条/拖动边界、作者上下文隔离。 |
| [chart-titles-facade.contract.test.mjs](../../packages/office-studio/tests/chart-titles-facade.contract.test.mjs) | 2（文件总计 5） | Builder 双标题字段解除引用/原子失败/独立恢复；live 一次 text/manual/rotate 作者提交、Undo/Redo、JSON 重开和恢复自动位置。 |
| [chart-title-interaction.browser.test.mjs](../../packages/office-studio/tests/chart-title-interaction.browser.test.mjs) 与 [fixture](../../packages/office-studio/tests/fixtures/chart-title-interaction.html) | 6 | 实际 Canvas 命中与 Ant 输入/格式同次提交、IME/Esc/焦点、拖动 preview/单次历史/JSON、只读/权限丢失/public veto、视图/宿主隔离、侧向/旋转/多行及图例 rebuild 后命中。 |

fixture 使用真实 OfficeStudio、Facade、ChartOverlay、Ant、ECharts 和 Canvas2D driver；preparedData 仅隔离绘制准备，作者命令仍需要后置的真实 Worker/WASM。这些是浏览器场景草稿，不是已经执行的完整工作台或两种渲染后端验收。组合事件草稿不能替代操作系统输入法实测。

Input/TextArea 的 composition 事件、nativeElement/focus 来自 Context7 获取的 [Antdv Next 官方当前源码文档](https://github.com/antdv-next/antdv-next/blob/main/packages/antdv-next/src/input/TextArea.tsx)，具体签名以本地 [TextArea.d.ts](../../node_modules/antdv-next/dist/input/TextArea.d.ts) 为准；颜色提交使用同次 Context7 查阅的 [ColorPicker changeComplete/valueFormat 契约](https://github.com/antdv-next/antdv-next/blob/main/docs/src/pages/components/color-picker/index.en-US.md)与本地 [interface.d.ts](../../node_modules/antdv-next/dist/color-picker/interface.d.ts)。只使用已有 Ant 节点的焦点/文本选区能力，没有新增原生表单控件。真实边界依据本地 [ZRender 6.1.0 Text.js](../../node_modules/.pnpm/zrender@6.1.0/node_modules/zrender/lib/graphic/Text.js) 的 getBoundingRect（112，实际 child rect union）与 [Transformable.js](../../node_modules/.pnpm/zrender@6.1.0/node_modules/zrender/lib/core/Transformable.js) 的负旋转约定（113）。依赖文档和静态源码均不是运行证据。

前端 Ant 迁移、主题、焦点与 IME 已在本批接线；真实行为及类型/前端/Rust/WASM 检查全部留待统一验证。本批没有重新执行检查，也没有声称性能、Office 保真或完整 Sheets/Docs/Pro 目标已完成。
