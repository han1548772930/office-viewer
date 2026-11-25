# SHEET-22 / SHEET-24：Treemap / Sunburst 的 Ant Tooltip 与原生祖先悬停

2026-11-25。接续[主、副标题批](SHEET-22-title-subtitle-contract.md)，处理既有图表属性清点中的悬停提示缺口。本批只落盘前端实现、测试草稿和证据；全部 **未验收**。按当前工作约束，没有运行测试、类型检查、构建、开发服务、浏览器、Rust/WASM、基准或新增自动检查。

## 来源、入口与可观察语义

- 固定 Univer 为 `E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；商业包只作静态参考。[图表审计 P26 / P27](SHEET-22-chart-type-audit.md) 已清点 Treemap / Sunburst 的公开 API、数据/样式页面、层级契约、权限和内容历史。本批没有新增常驻、扩展根或溢出按钮，也没有新增作者配置；两种图型沿用既有 Canvas 上下文悬停路径，事件名称和提示行数不作为功能完成数量。
- [engine-chart 固定归档](../../.research.local/univer-1.6.0/archives/engine-chart-1.6.0.tgz) 的 `package/lib/es/index.js`：`qce`（字符 offset 337583）对 Treemap `1048576`、Sunburst `2097152` 使用默认开启的 item 提示；专属 converter 保留实际 formatter。`eke`（约 1519990）给树图的正父自身值生成残差，并在残差中保留 `semanticNode`；`tke`（约 1521110）显示原父路径、语义总量、自身数值及子节点合计，不把残差 ID 当类别。
- 同一归档的 `XOe`（1517208）明确以当前节点的 `value / 根语义合计` 显示一位小数百分比，路径来自 semanticPath；self / child 同时非零才追加两项。`QOe`（紧随 XOe）保持 `nodeClick=false`、`sort=null` 与 ancestor emphasis。Treemap 对应 converter 同样使用 ancestor emphasis；悬停不生成作者 mutation。
- A1:C8 的既有有符号层级样本：[Treemap fixture](../../packages/office-studio/tests/fixtures/chart-treemap-fixtures.mjs) 与 [Sunburst fixture](../../packages/office-studio/tests/fixtures/chart-sunburst-fixtures.mjs)。North 语义总量 13、自身 −5、后代 18；South 总量 9、自身 3、后代 6；全部根总量 22。树图 South 残差提示仍报告父总量 9；旭日图 South 残差报告自身 3、占总量 13.6%，原父路径和 self=3 / child=6 保留。两种面积/角度语义没有互换。
- ECharts 当前项目锁定 6.1.0。本批通过 Context7 查阅 [tooltip.showContent](https://github.com/apache/echarts-doc/blob/master/en/option/component/tooltip.md)、[实例 on/off](https://github.com/apache/echarts-doc/blob/master/en/api/echarts-instance.md) 与 [鼠标事件的 params.data](https://github.com/apache/echarts-doc/blob/master/en/api/events.md)。查询指定 6.1.0，返回 master 文档，只作 API 参考，不能当固定版运行证据。静态读取本地锁定 `echarts/lib/core/echarts.js::_initEvents` 可见其沿真实事件 target 的父链取 dataIndex，并通过 getDataParams 发出 raw data；未调用该内部方法或复制其命中算法。

## 实际前端改动与接口

- [chart-content-tooltip.ts](../../packages/office-studio/src/features/charts/chart-content-tooltip.ts) 提供纯 presentation：ChartContentTooltip、层级上下文、两种语义 formatter 和文本投影。路径与权重从已准备数据的真实 native raw item 读取，不根据标签或像素坐标重建层级。非法/非有限字段不生成提示；树图子合计或旭日图百分比溢出也不生成虚假有限值。
- [echarts-treemap-option.ts](../../packages/office-studio/src/features/charts/echarts-treemap-option.ts) 与 [echarts-sunburst-option.ts](../../packages/office-studio/src/features/charts/echarts-sunburst-option.ts) 各自生成轻量 `contentTooltip` runtime。闭包只保留字段名及必要的总量/数字格式/日期系统，沿用既有 semanticNode 或 flat Sunburst facts，不再额外复制一棵语义子树。原有 formatter 仍调用同一 presentation；native tooltip 为 `showContent=false`、`triggerOn=none`，其 richText 配置不会显示浮层。
- [echarts-hierarchy-tooltip.ts](../../packages/office-studio/src/features/charts/echarts-hierarchy-tooltip.ts) 以每实例的公开 on/off 订阅 mouseover、mousemove、mouseout、globalout。只接受对应 seriesType 的真实项事件与有效 dataIndex；renderer 每次转发新点前清旧 capture，空白项或其他组件事件不能沿用上个节点。销毁只解绑本 binding 的四个函数，先释放 capture，不清除其他订阅；一个 off 抛错后仍尝试全部自有句柄，构造回滚报告原订阅错误，dispose 报告首个解绑错误。renderer 即使解绑报错也继续释放该 native 实例和画布，然后报告首个错误。
- [echarts-option.ts](../../packages/office-studio/src/features/charts/echarts-option.ts) 将这两个 runtime 放进既有 officeChartPointer；它们在 setOption / bitmap signature 前移除。[echarts-renderer.ts](../../packages/office-studio/src/features/charts/echarts-renderer.ts) 保留原生 ZRender pointer 转发、锁定 frame listener 的 pending emphasis 应用及 surface version 更新，移除这两种图型的 showTip/hideTip 动作。getContentTooltip 读取已采样同一 CSS 点的 capture，不再触发第二次事件或内容命令。
- 既有 Chord / Relation 的 getContentTooltip 对象、kind、id 和 index 仍保留；层级新增 `kind=hierarchy-node`、native dataIndex 与 view ID。它们是 renderer 的暂态身份，不能作为作者节点或保存模型。没有修改 Facade、domain、命令 ID、Worker DTO、WASM lease 或原生 Office 文件协议。
- [ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) 以共同 presentation 类型接现有受控 `antdv-next Tooltip`。多行文本继续使用实际 container 语义样式，统一 ConfigProvider / token / 明暗主题；Canvas 内容绘制与按钮命中区保留。hitbox 添加 data-chart-id 诊断属性，支持多个图表独立定位；没有新增可替换原生输入。
- view-only 工作表仍可悬停，编辑沿原 canEditWorksheetProtection 判断；提示不改 revision、内容历史、Undo/Redo 或保存点。原生实例的 DPR 替换、setOption 失败、retain、dispose 先解绑 capture，再释放画布；类型变成其他图型时解绑层级监听。pointer 转发/刷新失败会清 capture 并使内容缓存失效，后续重画前不读旧提示。

## 模块与计算归属

层级权重继续由既有 ChartStatisticsService / Worker / Rust 的 hierarchy job 提供，本批不增加聚合算法。纯 presentation 属于 charts 的模型/视图适配；浏览器 ECharts adapter 拥有真实图形命中和原生祖先高亮；Vue / antdv-next 拥有可替换提示浮层。坐标沿用现有 CSS 像素、对象原点、zoom 和 scroll 契约；DPR 仅用于 raster surface，不再次缩放 native hit。捕获的数据和 popup 状态不持久化，不写领域对象。

## 草稿与未验收范围

- [chart-hierarchy-tooltip.contract.test.mjs](../../packages/office-studio/tests/chart-hierarchy-tooltip.contract.test.mjs)：7 个未运行草稿，覆盖残差/有符号语义/格式、旭日图占比、非法事实、真实项事件的契约筛选、空白/leave 清理、两个实例及精确解绑、native 隐藏配置与共同 formatter、订阅回滚与多个 off 失败时全部尝试及原/首个错误。事件 port 的替身只验证 binding 契约，不证明 ECharts 运行时命中。
- [chart-hierarchy-tooltip.browser.test.mjs](../../packages/office-studio/tests/chart-hierarchy-tooltip.browser.test.mjs) 与 [实际 Overlay fixture](../../packages/office-studio/tests/fixtures/chart-hierarchy-tooltip.html)：4 个未运行草稿，计划覆盖双图型实际父/残差图形、Ant role=tooltip、只读和 revision 隔离、空白/leave/Escape/合成 window blur、主题/zoom/scroll、prepared facts 刷新及卸载/同 ID 新宿主。fixture 使用真实 OfficeStudio、公开数据投影、ECharts Canvas、ChartOverlay、Ant ConfigProvider；以 preparedData 隔离统计准备并用 Canvas2D 复制现有位图，不是完整工作台/WebGPU 或性能验收。倍增 prepared facts 是下一准备结果的替身，不验证真实源编辑到 Worker 的链路。
- 改写[既有 Treemap renderer 草稿](../../packages/office-studio/tests/chart-treemap-renderer.browser.test.mjs)和[既有 Sunburst renderer 草稿](../../packages/office-studio/tests/chart-sunburst-renderer.browser.test.mjs)各 1 个：检查提示只由 Ant presentation 输出、native Canvas 无重复提示字、ancestor / leave / reentry 保留。另在[Sunburst 既有契约草稿](../../packages/office-studio/tests/chart-sunburst.contract.test.mjs)追加 native hidden 断言。这些旧草稿本轮也没有运行。
- Treemap 固定 `tke` 使用原始数字，本地继续保留此前的 source formattedValue / formattedSelfValue 优先行为；旭日图继续 label.format 优先、source display 回退。自己的/子节点合计文案仍为现有中文，与上游完整多语言策略不同。纯文本 Ant 内容替代上游 HTML / 本地旧 Canvas richText 的排版，不复刻字号、边距、marker、HTML dir 或像素细节；完整 RTL / 混合方向 / 复杂脚本与窄视口尚未核实。
- 深树/大文本/极小扇区、透明边框与叠加标签命中、祖先 banner / 残差真实命中、空源到有源、只读 ACL 全部上下文、触摸/键盘可访问悬停、焦点/真实窗口 blur/IME、类型切换与失败恢复、多个宿主同时存在、实际 DPR / 两后端 / Docs 宿主均未验收。悬停提示迁移不代表标题/系列选择、浮动工具栏、拖动或所有图型 tooltip 入口已经完成；其余图型 native 可替换提示仍待逐项接同一 Ant host。
- 没有新增后端或文件写回能力，没有运行 XLSX/DOCX 或 DrawingML 样本。没有耗时、帧时间、长任务、传输字节、峰值内存或真实释放测量；事件 capture 与缓存策略只是代码实现，不能称性能达标或整链路零拷贝。
- 本批仅作源读取与局部静态代码阅读；没有执行任何新增检查，测试/类型/浏览器/文件/性能全部继续 **未验收**。完整 Sheets / Docs / Pro 工作台目标保持有效。
