# SHEET-22 / SHEET-24：P15 绘图区选择、拖动、大小与独立恢复

2026-11-25，接续[绘图区既有配置审计](SHEET-22-chart-type-audit.md)和[标题直接编辑](SHEET-22-title-direct-editing.md)。本批前端代码、15 个新增测试场景草稿及本文已落盘，状态均为 **未验收**。按用户当前“先实现、后统一验证”安排，本批未运行测试、类型检查、构建、dev server、浏览器、Rust/WASM、性能基准、`git diff --check` 或自动链接检查。

## 固定来源与证据边界

固定只读仓库为 `E:\rust_test\univer`，Univer v1.6.0、commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。以下字符 offset 属于固定发布 ES 的静态阅读位置；归档内路径指归档成员，不冒称已检出的独立源码文件。

| 来源 | 静态核实结果与边界 |
| --- | --- |
| [chart-ui ES](../../.research.local/univer-1.6.0/packages/chart-ui/lib/es/index.js)，`Wi` 40058、`Yi` 40630 | PlotArea 能力映射包含 Column、Bar、Line、Area、Scatter、Bubble、Combination、Pie、Radar 九族；映射到本地既有目录为 31 项。族/目录数量不是 UI 按钮数量，也不是图型验收数量。 |
| 同文件 `Kr.PlotArea` 19105、`Ru/Bu` 201621、`Ku/qu/Ju` 207206 | 确认 plot-area 配置键、样式状态观测/投影以及共用 `commit({ [key]: value })` 的作者组路径。该通用路径不能单独证明存在四坐标输入、独立恢复或拖动作者入口。 |
| 同文件 `hs` 91523、`ql` 182316 | 已阅读的 chart element hit / 浮栏映射没有 plot-area kind；本轮静态搜索没有定位实际四坐标 section 或 plot 拖动/resize 写入。不能据此断言上游不存在这些能力，也不把本地新增入口记为已核实 Univer 入口。 |
| [engine-chart 1.6.0 归档](../../.research.local/univer-1.6.0/archives/engine-chart-1.6.0.tgz)，`package/lib/es/index.js::Y6` 1532157 | 确认实际矩形校验：四项为有限数字；left/top ≥ 0；width/height > 0；left + width ≤ 1；top + height ≤ 1。因此“矩形范围约束未核实”已由此次固定 ES 阅读补齐；本地仍有普通对象校验和 `1e-12` 边界容差差异。 |
| 同一 ES `X6` 1532382 | 四个归一化字段乘整个图表 width/height；缺省图表尺寸为 468 × 369。定位基准是完整外框，不是先扣标题/轴后的内框。 |
| 同一 ES `gke/$6/e8` 约 1534375–1535900、`_ke` 1536839 及随后 `r8` | 自动布局包含标题、图例、轴留边和 adaptive/RTL 分支；有效 manual 矩形原样复制。自动 Pie 有 best-fit 及历史 .55、部分隐藏图例/标签 .925 分支，Doughnut .5、Radar .65；其他从自动 grid 归一化。manual grid 使用 `containLabel:false`、`outerBoundsMode:'auto'`。本地自动 seed 不宣称复制这些像素算法。 |
| 同一 ES `g7/v7` 1672236 / 1672682 | 事件/overlay 包含真实 hit bounds 和 localPoint；本轮已阅读映射没有 plot kind。上下文选择/拖动协议尚无固定版直接对应入口证据。 |
| 同归档的 `package/lib/types/types.d.ts:867::IChartPlotArea`、`package/lib/types/chart-builder/chart-types.d.ts:502::IChartPlotAreaSpec`、`package/lib/types/models/chart-render/converter/render-spec-operators/plot-area.d.ts`、`package/lib/types/models/chart-render/converter/render-spec-operators/grid-operator.d.ts`、`package/lib/types/chart-element-interaction/chart-element-hit-layout.d.ts` | 确认矩形类型及转换、grid、命中模块职责。类型声明与上述 ES 阅读一起构成配置依据；不是运行证据。 |

本批边缘选择、显式按钮、浮栏、正文拖动、八向调整大小和键盘操作属于本地 Office 语义交互实现；对应 Microsoft Office 样本/公开演示尚未实测。源事实、实现选择及待核实项分别保留，不能以按钮名称推断等价或将 Pro 范围缩减为本批切片。

## 入口分类、顺序与计数

本批不新增常驻 Ribbon、插件扩展根项或 More 溢出根项。上游只核实九族/31 项能力映射、配置键和作者组路径；实际 PlotArea 输入/恢复按钮总数、classic/grid 位置与运行显隐未定位，仍待继续清点。

| 本地分类 | 顺序与布局 | 数量 | 显隐、禁用与调用 |
| --- | --- | ---: | --- |
| 既有图表设置区值入口 | 手动开关 → left → top → width → height；四项均为完整图表外框的百分比 | 5 | 限上述九族；坐标限手动模式显示。保留设置面板既有草稿/应用作者路径。 |
| 既有设置区新增恢复子动作 | 开关恢复整组；left/top/width/height 分别恢复自动值 | 5 | 按是否有作者值及与 automatic 的差异判断；单项恢复保留其余三项，不兼容组合仍是可纠正草稿。 |
| 选中图表上下文入口 | 图表对象右上方“绘图区”Ant Button | 1 | 仅选中且支持 PlotArea 的图表显示；只读可进入查看。不可见 plot 仍可通过此入口打开恢复表单，不制造不可见命中。 |
| 新上下文浮栏值入口 | 顶部手动开关；两列坐标，首行 left/top，次行 width/height | 5 | 自动模式坐标显示 automatic 值；手动模式编辑 local draft。只读、作者提交中、组合输入中禁用相应动作。 |
| 新浮栏恢复子动作 | 四坐标行内恢复 → 自动位置 → 自动大小 → 整组自动布局 | 7 | 四项独立恢复保留其他字段；自动位置只改 left/top，自动大小只改 width/height；整组清除 plotArea。全部先改草稿，由 Apply/Enter 提交。 |
| 新浮栏宿主动作 | 底部“应用修改”、顶部关闭 | 2 | 有效且有差异时可应用；关闭/Esc 丢弃未提交草稿并向仍存活的原图表宿主恢复焦点。 |
| Canvas 覆盖层大小 handles | nw → n → ne → e → se → s → sw → w | 8 | 都是 Ant Button；按实际可见 plot bounds 定位。只读或提交中禁用；对边固定，指针松开写一次完整矩形。 |

边缘/正文命中、悬停轮廓、方向键、Shift 加速和 Delete/Backspace 单独作为交互行为记录；不把它们与控件数量相加推导上游入口或功能完成数量。恢复、Apply/关闭是动作，九族/31 项是能力目录，二者统计口径不同。

## 模块归属、计算域与协议

对照[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)与[WASM 平台/内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)：作者矩形沿用 charts 领域模型 `ChartPlotAreaOptions`；归一化、命中、移动/大小算法是独立可测试 helper；自动 seed 复用既有 chart layout；ECharts 是 Canvas 适配器；浮栏/handles 属于 Vue/Ant 视图；内容修改由既有 Facade、统一命令、候选准备、权限、内容历史与内部保存链路编排。

`ChartCanvasPlotAreaTarget` 是主线程 surface 暂态 DTO，包含 `frame/plot/bounds?/area/automatic/manual`。它不进入持久化模型或跨 Worker/WASM 协议。命中和四字段交互是有界小操作，本批保留前端路径，尚无同数据集基准证明它优于 WASM；图表重准备仍沿现有 Worker/WASM 作者链路。本批没有新增计算引擎、buffer lease 或文件转换协议，也没有零拷贝/性能达标主张。

## 已落盘的前端行为

- [chart-plot-area-interaction.ts](../../packages/office-studio/src/features/charts/chart-plot-area-interaction.ts) 使用同一 `chartPlotFrame` / shared layout 提供的绘图区矩形。自动 seed 复用 `chartPlotAreaFromFrame`；manual 由原矩形直接归一化，避免把低于自动 seed 下限的既有作者宽高放大。target 的矩形与字段解除引用，不由整图外框、文字或 series marks 猜测 plot。
- 绘图区轮廓/handles 剪裁到 chart frame。完全不可见时不生成命中、线条或 handles；显式入口可打开表单恢复。未选中仅命中边缘，容差为 `5 / zoomScale()`，即 5 个屏幕像素；选中后正文可拖动。优先级为标题 → 图例 → 绘图区 → Sankey 内容 → 整图。
- 拖动超过 4 个屏幕像素才创建 preview；preview 不改作者矩形、图表 position/size 或内容历史。指针松开提交最后一份完整矩形一次，提交期间保留 preview，失败后回作者状态。移动保留宽高并限制整个矩形在 0..1；八向大小固定对边，最小拖动尺寸为 8 个图表逻辑像素，已有更小 manual 矩形保留原下限。
- [ChartPlotAreaFloatMenu.vue](../../packages/office-studio/src/features/charts/views/ChartPlotAreaFloatMenu.vue) 的手动开关、四个百分比及全部恢复只修改 local draft。Apply 或数字 Enter 合并完整矩形提交；无效组合显示 Ant Alert，不发布内容、不悄悄夹紧其他字段。清空坐标时，提交候选取该字段的 automatic 值。整组恢复写 `plotArea:undefined`，保留标题及其他图表选项。
- [chart-plot-area-controls.ts](../../packages/office-studio/src/features/charts/chart-plot-area-controls.ts) 为既有开关/四坐标添加恢复动作，复用本地自动 seed；单项恢复不连带修改其他坐标。无效恢复组合仍由设置面板保留供纠正，作者候选校验拒绝无效提交。
- 选中绘图区后，图表宿主方向键移动矩形；handle 的方向键调整对应边，Shift 为 10 个逻辑像素。Delete/Backspace 恢复 plot 自动布局，不删除图表。Canvas 键盘计算使用最新作者 plotArea，避免下一帧重绘前重复使用旧 frame 的矩形。
- 只读可以选择/查看，作者字段和 handles 禁用。Escape、外部指针、滚动、zoom/theme/font 变化、窗口 blur/resize、DPR、ResizeObserver、工作表/宿主切换和卸载结束编辑会话。关闭后只向仍存在的原宿主恢复焦点。

上述为代码接线效果，实际交互、命令历史及保存/重开仍待运行；未以空操作或提示消息充当图表实现。

## 作者路径与异步隔离

统一路径为 `FOffice.newAPI(studio) → getWorkbook(workbookId) → getSheetBySheetId(sheetId) → getChartById(chartId) → FChart.setOptions({plotArea}) → sheet.command.update-chart`。本批不创建新命令，也不改 Facade 源码或 ChartOverlay 的既有 emit payload；浮栏、拖动、键盘与设置区沿用图表内容历史/准备事务。

[ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) 复用 `ChartTitleAuthorContext` / `chartTitleAuthorContextMatches`，检查 owner/workbook/sheet/chart/revision/snapshot 的作者上下文；查询 revision 前先确认 workbook 仍存在。通用检查命名为 `chartElementSelectionIsCurrent`，原标题 wrapper 保留。

浮栏 apply callback 捕获选中对象，旧输入不会动态写入新选区。真正拖动、Canvas 键盘编辑及删除恢复更新 generation，淘汰旧未提交草稿；标题/绘图区作者请求互斥，拖动期间拒绝浮栏作者更新。作者提交成功等待 `nextTick` 再更新 selection context；菜单同步草稿使用作者矩形，不复制尚未重绘的 preview target。失败时必要由 Overlay Alert 报告，preview 在 finally 清除。

已 Apply/松开的内容请求继续遵守现有候选事务的版本/权限与 Promise 语义；关闭视图不另造撤销协议。hover/selection/reset draft 不写 revision；真实作者变更的事件、历史及 JSON 重开使用既有机制，仍需后置测试。

## 内部接口与修改文件

- `ChartEChartsOption.officeChartPlotArea?:ChartCanvasPlotAreaTarget`：绘制准备传递精确 plot/automatic geometry；在 ECharts `setOption` 和 signature 生成前拆除。
- `ChartCanvasLegendFrameOption.plotArea?:ChartCanvasPlotAreaTarget`：shared legend rebuild 更新其 target；完整 render、缓存命中同步 target，legend 失败恢复旧 target。
- `EChartsCanvasRenderer.getPlotAreaTarget/getPlotAreaHit`：仅返回 chart-local surface 几何/命中，不作者写入。
- [index.ts](../../packages/office-studio/src/index.ts) 增加导出既有领域类型 `ChartPlotAreaOptions`。领域模型本批未改；共享 DrawCommand、live Promise、内容历史、保存模型和 Worker/WASM 协议沿用既有契约。

本批共 13 个文件，其中 8 个代码文件、4 个测试/fixture 和本文：

| 分类 | 文件 |
| --- | --- |
| 纯交互与既有设置 | [chart-plot-area-interaction.ts](../../packages/office-studio/src/features/charts/chart-plot-area-interaction.ts)（新增）、[chart-plot-area-controls.ts](../../packages/office-studio/src/features/charts/chart-plot-area-controls.ts) |
| Canvas 适配 | [echarts-option.ts](../../packages/office-studio/src/features/charts/echarts-option.ts)、[echarts-legend.ts](../../packages/office-studio/src/features/charts/echarts-legend.ts)、[echarts-renderer.ts](../../packages/office-studio/src/features/charts/echarts-renderer.ts) |
| Ant 浮栏与宿主 | [ChartPlotAreaFloatMenu.vue](../../packages/office-studio/src/features/charts/views/ChartPlotAreaFloatMenu.vue)（新增）、[ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) |
| 图表类型导出 | [index.ts](../../packages/office-studio/src/index.ts) |
| 新测试草稿 | [chart-plot-area-interaction.contract.test.mjs](../../packages/office-studio/tests/chart-plot-area-interaction.contract.test.mjs)、[chart-plot-area-interaction.browser.test.mjs](../../packages/office-studio/tests/chart-plot-area-interaction.browser.test.mjs) |
| 扩展测试/fixture | [chart-gradient-layout-facade.contract.test.mjs](../../packages/office-studio/tests/chart-gradient-layout-facade.contract.test.mjs)、[chart-title-interaction.html](../../packages/office-studio/tests/fixtures/chart-title-interaction.html) |
| 独立证据 | 本文 |

`chart-plot-area.ts` 本批未改；Workspace/OfficeStudio/Facade 非图表段、后端及中央清单由主任务持有。工作区还有标题等既有批次改动，本表只统计本批文件，不把整个工作区计为本批成果。

## 测试草稿与静态审阅

新增 15 个场景草稿，全部未执行：

| 文件 | 新增数量 | 计划验证 |
| --- | ---: | --- |
| [chart-plot-area-interaction.contract.test.mjs](../../packages/office-studio/tests/chart-plot-area-interaction.contract.test.mjs) | 8 | target/31 项能力映射与解除引用；裁剪/边缘/正文/不可见；八 handles/Canvas 轮廓；移动与整矩形边界；八向固定对边及小 manual 保留；独立恢复/无效草稿/整组恢复；native/shared metadata 与 legend rebuild；标题作者与其他选项保留。 |
| [chart-gradient-layout-facade.contract.test.mjs](../../packages/office-studio/tests/chart-gradient-layout-facade.contract.test.mjs) | 1（文件共 4） | helper 生成 resize 矩形后调用既有 live setOptions；一次 command/revision；Undo/Redo；position/size/source/title 保留。 |
| [chart-plot-area-interaction.browser.test.mjs](../../packages/office-studio/tests/chart-plot-area-interaction.browser.test.mjs) | 6 | 真实 plot 边缘/上下文按钮与四字段单次提交；正文 preview/旧输入/Undo/Redo/JSON；八 handles；无效草稿及全部恢复；只读/权限/public veto/视图/宿主；键盘/Delete、标题优先、native/shared legend rebuild 和 metadata 拆除。 |

复用真实标题 fixture，并扩展 plot target/point/driver。使用 OfficeStudio、Facade、ChartOverlay、Ant、ECharts 和 Canvas2D；`preparedData` 仅隔离 paint 准备，作者提交仍需要统一验证阶段生成匹配的真实 Worker/WASM。浏览器草稿等待 target.manual/area 与作者一致，避免只见 busy 结束而误认为 Canvas 已重绘；拖动 preview 另等几何变化。旧宿主迟到数值使用合法矩形候选，标题切换也先创建合法未提交 plot draft，不依赖无效范围偶然阻断旧输入。

静态阅读复核了源路径/版本、target 与 paint 的同源关系、恢复对其他作者字段的保留、preview 到作者调用、generation/宿主/revision 隔离、legend rebuild 与 metadata 拆除；未执行自动检查或运行断言。控件签名依据此前 Context7 查阅的 Antdv Next InputNumber `update:value/pressEnter/composition`、Switch `checked` 文档及本地 [InputNumber 类型](../../node_modules/antdv-next/dist/input-number/index.d.ts)。这些资料均不是运行验收。

## UI、差异与剩余缺口

本批新增/修改的可替换 UI 使用 `antdv-next` 的 Form/FormItem、InputNumber、Switch、Button、Tooltip、Alert，handles 也为 Ant Button；Canvas 轮廓/命中/preview 属于编辑器内核。主题来自现有统一 ConfigProvider 和 `theme.useToken`，覆盖浮栏背景/边框/阴影和 handle 状态；没有新增原生表单控件。InputNumber composition 期间阻止 Enter/作者动作，未创建另一套 DOM 模型。主题、实际焦点和操作系统输入法均尚未运行核验。

1. 固定版四坐标输入/恢复入口和 plot hit/拖动/大小写入协议尚未定位；本地新增交互的 Office 行为样本尚未核实，不能标为完全复刻 Univer。
2. 本地 automatic seed 包含标题、图例和活动次轴留边，但 fixed adaptive/RTL、Pie .55/.925、Doughnut .5、Radar .65 等布局分支仍存在算法差异。恢复取本地同源 automatic geometry，不宣称上游像素一致。
3. 最小 drag 尺寸、4/5 屏幕像素阈值、完整矩形 0..1 clamp、按 Apply/Enter 批量提交、无效草稿保留和独立恢复是本地交互选择。有限范围已获 ES 依据；`1e-12` 校验容差仍是明确本地差异。
4. 真实 OS IME、触屏/多指、窄屏浮栏/弹层、焦点归还、尺寸/DPR/字体变化和两种渲染后端的实际交互闭环仍未验收。
5. 所有 45 图型/全部变体、Docs 嵌入、协作 ACL、公共事件全路由、XLSX/DOCX 写回/独立重开仍需完整实现与验证；本批没有删去目标，也不将 JSON 重开草稿等同于 Office 保真。
6. target、DrawCommand 与指针 preview 有主线程 JS 分配；大样本首屏/滚动/编辑/缩放、主线程长任务、跨 Worker/WASM 复制/传输及内存没有测量，性能保持未验收。

中央 P15/标题状态与后端需求由主任务协调更新。本批只补本切片证据，完整 Sheets/Docs/Pro 工作台仍按原清单继续推进。
