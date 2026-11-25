# SHEET-22 / SHEET-24：主、副标题对象配置与四边绘制

2026-11-25。接续[图例批](SHEET-22-legend-selection-wrap.md)的相邻属性入口。本轮只落盘前端代码、测试草稿并作静态审阅，全部保持 **未验收**；没有运行测试、类型检查、构建、开发服务、浏览器、Rust/WASM 或性能基准。不将配置行数或测试草稿数量当成功能验收。

## 固定来源与入口清点

- 固定 Univer 为 `E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；商业包只作本地静态参考。
- [公开配置类型](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/chart-builder/chart-types.d.ts) 声明 `ChartTitleSpec = string | IChartTitleOptions` 和 `ChartSubtitleSpec = string | IChartSubtitleOptions`。主标题对象包含 text、visible、position、alignment 及扁平 color/fontSize/bold/italic/family；副标题没有独立 visible/position，包含 text、alignment 和字体。[Builder](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/f-chart-builder-base.d.ts) 与 [FChart](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/f-chart.d.ts) 均有 setTitle/clearTitle/setSubtitle/clearSubtitle。
- [Enum](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/enum.d.ts) 的 TitlePositionEnum 包含 Top/Bottom/Left/Right/Hide，LabelAlignEnum 包含 Left/Right/Center；[Facade Enum](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/f-enum.d.ts) 对应 `ChartTitlePositionEnum`、`ChartLabelAlignEnum`。
- [ChartTitleSections 声明](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/editor/sections/ChartTitleSections.d.ts) 与 [UI ES](../../.research.local/univer-1.6.0/packages/chart-ui/lib/es/index.js) 的 `Zc`（字符 offset 143662）、`Qc`（144787）明确给出主标题文字 → 五个位置值 → 字体/对齐，副标题文字 → 字体/对齐；Hide 通过 visible=false 表达，选择其他位置同时恢复 visible=true。`Jc` 的主标题字体默认字号 16、居中，副标题字号 14、UI 对齐默认 left。
- [engine-chart 归档](../../.research.local/univer-1.6.0/archives/engine-chart-1.6.0.tgz) 中 `package/lib/es/index.js` 的 `rne`（约 55914）区分字符串与对象，将 visible=false 映射为 `titles.titlePosition=hide`，对象的未给出字段不覆盖原值；`nne` 清主标题和 titlePosition。`Die`/`Oie`（82820/82841）记录 set/clear；Facade ES 的 Builder 与 live FChart 都调用这些操作。
- 同一归档的 `gke`（1534375）按四边预留标题空间，`w5`（1647031 附近）隐藏 native title 并生成共同 graphic：Hide 隐藏主、副标题整个块。`S5`（1644853）在左右位置逐字纵排，主标题靠外，副标题向内；`C5`（1645955）在上下位置按独立对齐绘制，Bottom 的副标题在主标题上方。`x5` 提供副标题的内向偏移。运行源码的水平默认对齐为 center，侧向默认继承 textStyle.align=left，与部分 UI 默认显示不同；最终公开效果尚未运行核实。
- 本地现有标题/副标题文字只有字符串 API，字体已有表单但无组清除；本轮新增 **3 个上下文配置行**（主标题位置、主标题对齐、副标题对齐），给 2 个既有文字行和 10 个字体叶行补恢复动作，共 15 个恢复入口，其中 2 个整组、13 个字段目标。位置/对齐的选择值不是功能数。没有新增常驻 Ribbon、扩展或溢出根入口，完整清点仍受[图表审计](SHEET-22-chart-type-audit.md)约束。

## 作者模型、命令、历史与重开

- [model.ts](../../packages/office-studio/src/domain/model.ts) 添加 `ChartTitlePosition`、`ChartTextAlignment`、`ChartTitleStyleOptions`、`ChartSubtitleStyleOptions`，只扩展既有 options.titleStyle/subtitleStyle；旧 title/subtitle 字符串和旧字体快照保持兼容，没有改成另一份嵌套文本权威模型。
- [chart-titles.ts](../../packages/office-studio/src/features/charts/chart-titles.ts) 集中公开 spec → 本地 patch、有效字体、纯几何与 TextCommand。字符串只改文字，对象按给出的字段更新，显式 text=undefined 只清文字；clear/undefined 清该标题的文字和样式整组，保留另一标题和 appearance。扁平字体归入对应 style，候选先校验再替换并解除输入引用。继续采用模型既有的 100000 字符边界，没有缩减旧文本范围。
- [chart-option-validation.ts](../../packages/office-studio/src/features/charts/chart-option-validation.ts) 校验 visible、位置与主/副 alignment；非法字体/颜色/位置在候选提交前拒绝。副标题公开 API 不接受独立 visible/position；保存模型仍按既有规则保留可序列化的未知字段，没有把未识别数据丢弃后声称保真。
- [facade.ts](../../packages/office-studio/src/api/facade.ts) 的 Builder/live FChart 扩展 setTitle/setSubtitle 并增加 clearTitle/clearSubtitle，复用原有 `sheet.command.update-chart`、权限、按文档内容历史与准备事务。空对象没有内容提交；非法对象在调用内容命令前拒绝。Builder 同步返回自身，live 继续遵循本项目 Promise<this> 契约，不声称与上游同步 live 返回类型相同。
- [index.ts](../../packages/office-studio/src/index.ts) 导出 ChartTitlePositions/ChartTextAlignments、spec 和领域 style 类型；FOffice.Enum 使用固定上游名称。Builder build/getInfo 继续是本地字符串 + style 结构；describe 也沿用既有本地投影，未改为上游嵌套 title DTO。
- 属性编辑继续 detached draft → 一次 Apply → 既有内容命令；恢复文字会清文字/字体/对齐/位置整组，恢复字号/字体/颜色/粗体/斜体/对齐/位置只清对应覆盖。图例 selected/page 和 Canvas 指示器仍属于 renderer view，标题配置不写入这些 transient 状态。清主标题后副标题采用默认 Top，符合共享位置恢复的本地契约。
- 内部 JSON 保存/重开可以表达全部新增作者字段，其真实历史/重开行为目前只有未运行草稿。XLSX/DOCX/DrawingML 的写回映射未在本轮扩展或运行，不能把内部 JSON 保留当作原生 Office 等价。

## 表单、Canvas 与计算归属

- [chart-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-settings-controls.ts) 将主标题文字/位置/对齐、副标题文字/对齐接在现有图例之前。字号/字体等读有效继承值，恢复只清作者覆盖。[ChartSettingsFields.vue](../../packages/office-studio/src/features/charts/views/ChartSettingsFields.vue) 使用这些配置及恢复目标，继续由 `antdv-next` Form.Item/Input/InputNumber/Select/Switch/ColorPicker/Button 承载，统一 ConfigProvider/token；没有新增原生可替换输入。图例字体面板同时复用 chartHasLegend，WordCloud 与其他无通用图例图型一致隐藏该面板。
- [chart-layout.ts](../../packages/office-studio/src/features/charts/chart-layout.ts) 用共同标题几何向 Top/Bottom/Left/Right 预留空间；手动 PlotArea 的绝对作者比例保持原值。左右标题逐 Unicode code point 纵排；上下标题按有界行宽换行。主/副文字各自对齐、字体与颜色，副标题沿对应边向内部偏移；Hide 不画标题块，也不预留其区域。
- [chart-draw.ts](../../packages/office-studio/src/features/charts/chart-draw.ts) 与 [echarts-option.ts](../../packages/office-studio/src/features/charts/echarts-option.ts) 共用 TextCommand，不再由两份 title/subtext 配置决定几何。native 内容也追加共同 graphic，保留轴标题、日期轴、边框和 Heatmap 色阶 graphic；默认标题颜色统一，主题只改变无作者颜色的有效值。
- [chart-legend.ts](../../packages/office-studio/src/features/charts/chart-legend.ts) 的槽位可用宽度扣除左右标题带；native 与纯 DrawCommand 图例按对应标题边向内排列。它延续[图例批](SHEET-22-legend-selection-wrap.md)的选择/分页与故障恢复契约，没有将标题重绘变成 view mutation 或重复计算统计。
- [chart-heatmap-draw.ts](../../packages/office-studio/src/features/charts/chart-heatmap-draw.ts) 在原图表 frame 内派生扣除标题带的色阶 frame，更新 continuous bar、piecewise swatch 与相应 hit 矩形；外部 frame/裁剪协议保持原图表坐标。绘制、指示器和 native Heatmap 都消费同一个色阶布局。布局预留按扣除侧向标题后的宽度计算分段色阶行数，避免仅移动画面而留下旧命中区域。
- 这些逻辑属于 charts 模型适配/纯布局/Canvas chrome。Worker 可复用同一纯几何，浏览器平台层负责实际 Canvas 位图；未改 Worker DTO、WASM 平台或内存 lease 协议，没有增加第二套权威求解器。字体几何不访问 DOM，不在 Vue 组件中解析文件或求解公式。

## 已知差异与未验收范围

- 为旧快照保留本地 14px 主标题/11px 副标题及主标题默认 bold；它与固定版水平 16/14px、侧向不同默认字号有差异，字体作者字段与全局继承可明确覆盖。侧向未作者指定对齐时本地仍居中，与 `S5` 的默认 left 有差异；本地表单显示值与其绘制保持一致。
- 采用纯 code point 宽度估算、字间换行和 frame 裁剪，不能宣称复制了上游 native 字体测量、单词换行或像素外观。viewport 放不下时以有界行数/省略号绘制，原文字完整保留在模型。已处理 surrogate pair 的末尾截断；grapheme/ZWJ/复杂脚本整形、字号极大/极小画布、轴标题碰撞和完整 RTL 未验收。
- 本批配置阶段未包含 graphic title/subtitle 命中与浮栏。2026-11-25 的[标题交互续作](SHEET-22-title-direct-editing.md)已接实际文字边界、Ant 浮栏、直接输入和拖动单次提交，全部仍未验收。固定版已核实命中/浮栏及水平 manualLayout 读取；没有找到标题拖动写入或 rotate 绘制，续作将拖动/旋转和侧向手动位置明确记为本地扩展。手动 PlotArea 保持作者几何。
- 跨宿主 Docs 图表、协作/ACL 的全部上下文、原生 Office 样本和 45 个图型逐变体仍缺证据。没有安装商业 SDK、调用 Univer Server 或缩减 Pro 范围。
- 没有性能优化实测。文字扫描、换行、标题带和 DrawCommand 都有 JS 分配；大样本首屏、滚动、编辑、缩放、Worker/WASM 任务、传输与内存指标保持未验收，不声称零拷贝或帧时间达标。

## 草稿与静态检查

- [chart-titles.contract.test.mjs](../../packages/office-studio/tests/chart-titles.contract.test.mjs)：7 个草稿，覆盖 spec 原子校验/解除引用、旧字符串局部更新/清除、四边与 Hide、继承/独立恢复、native/shared 标题一致与 Heatmap graphic 保留、文本边界/code point/手动 PlotArea、两种色阶的矩形/命中避让。
- [chart-titles-facade.contract.test.mjs](../../packages/office-studio/tests/chart-titles-facade.contract.test.mjs)：本批原有 3 个草稿，覆盖固定 Enum、Builder 对象/字符串/独立清除，live 一次内容操作的 Undo/Redo/内部保存重开和非法/空配置的原子边界；[交互续作](SHEET-22-title-direct-editing.md)另增 2 个，总计 5 个，均未运行。后续运行需要正常 OfficeStudio/WASM 平台。
- [chart-titles.browser.test.mjs](../../packages/office-studio/tests/chart-titles.browser.test.mjs) 与 [fixture](../../packages/office-studio/tests/fixtures/chart-titles.html)：3 个草稿，计划覆盖真实 ECharts Canvas 文本、实际 ChartSettingsFields 的 Ant 控件→Apply→内容历史→JSON 重开、主题/DPR 后新位图与 Input 焦点。preview 使用 preparedData 隔离统计准备，Apply 使用真实 public chart update；不是完整 Workbench/WebGPU、真实 Office 文件或性能验收。
- 上批执行时，tracked 改动的 `git diff --check` 未报告空白错误，当时 27 个本地链接检查找到目标；本次交互续作没有重跑，结果不覆盖新增文件或链接。草稿、类型、前端、Rust/WASM、浏览器/焦点/真实 IME、文件样本与性能均未运行，全部继续 **未验收**。
- 完整 Sheets/Docs/Pro 工作台目标仍有效；本轮标题局部前端实现、已知差异和后端后置不构成最终范围缩减。
