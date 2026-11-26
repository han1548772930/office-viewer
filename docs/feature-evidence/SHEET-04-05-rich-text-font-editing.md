# SHEET-04 / SHEET-05 富文本字体与 Canvas 编辑接线

记录日期：2026-11-25。固定参考根 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。状态：**前端已补代码，未验收**。按用户安排，测试、类型检查、构建、Rust/WASM、浏览器、dev 和性能基准均后置；本记录没有任何运行通过结论。

本记录续接[字体与字号](SHEET-04-font-family-size.md)、[颜色与重置](SHEET-04-color-reset.md)和[输入范围轨迹](COMMON-03-input-selection-axis-history.md)，不将入口数量、内部快照或测试草稿当作全部功能完成。

## 固定来源与入口计数

以下路径均相对固定参考根，未加入 Univer 商业 SDK 或改动参考仓库。

| 来源 | 核实的可观察契约 |
| --- | --- |
| `packages/sheets-ui/src/menu/schema.ts` 的 START / FORMAT | 字体组 11 个常驻根声明：第一行字体、字号、增大、减小共 4 项；第二行加粗、斜体、下划线、删除线、文字颜色、填充、边框共 7 项。classic order 与 grid 位置沿既有证据保持。 |
| `packages/sheets-ui/src/menu/font.menu.ts`、`menu/menu.ts`、`menu/utils.ts` | 字体/字号控件与字符开关、颜色状态，范围或编辑器光标的有效字体，工作簿/工作表/范围能力判断。 |
| `packages/sheets-ui/src/commands/commands/inline-format.command.ts` | 同一 Range 字体命令按输入上下文分派到整格或行内格式；`set-range-font-family`、`set-range-fontsize`、`set-range-font-increase/decrease`、`set-range-bold/italic/underline/stroke/text-color`，文字色恢复为 `reset-range-text-color`。 |
| `packages/sheets/src/basics/cell-style.ts` 的 `handleStyle()`、`mergeRichTextStyle()` | 整格字体变更也处理可见运行覆盖，不能让原 rich run 覆盖新整格样式；无关内容/扩展继续保留。 |
| `packages/sheets/src/commands/commands/set-style.command.ts`、`commands/mutations/set-range-values.mutation.ts` | 样式提交进入模型、筛除行和按文档隔离的内容历史。 |
| `packages/sheets-ui/src/controllers/editor/editing.render-controller.ts`、`packages/engine-render/src/components/sheets/extensions/font.ts` | Sheets 编辑器与 Docs 文字运行/布局的桥接，显示文字、输入、字体度量与绘制分层。 |
| `packages/docs-ui/src/services/selection/doc-selection-render.service.ts`、`packages/engine-render/src/scroll-timer.ts` | 捕获文字拖选在边界外持续按帧滚动，更新实际选区，结束手势时释放 timer；本地按编辑 clip 和 rich glyph extent 独立实现。 |

本批覆盖字体组中 **9 个可作用于行内文字的根入口**：第一行 4 项、第二行字符开关 4 项、文字颜色 1 项。填充与边框继续作用于整格，因此全组仍为 11 个根入口；文字颜色重置、下拉候选和固定别名不另计根入口。插件扩展新增 0 项、独立上下文根新增 0 项，overflow 重用原入口；本地局部 API 和草稿 operation 不是新增 Ribbon 按钮。上标/下标虽在固定命令文件出现，不属于本组 11 根项，完整扩展清点仍须继续。

## 模型、命令、API 与历史

- [sheet-cell-rich-text-style.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-rich-text-style.ts) 读取嵌套 `font` 和公开顶层别名，保留显式 false、链接、空运行和未知扩展。整格字符/字体/文字色变更同步相应运行覆盖；null 仅清对应作者字段，填充不覆盖运行字体。
- [sheet-cell-inline-format.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-inline-format.ts) 以原生输入的 UTF-16 范围轨迹顺序重放部分格式和后续文字插入。拆分运行时保留来源身份及无关数据；光标格式作为 typing 偏好，只应用到其后实际插入的文字。草稿撤销/重做跟随输入 journal，分支编辑丢弃旧未来。
- [sheet-cell-inline-format-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-inline-format-command.ts) 注册 `sheet.operation.format-cell-input` 与 `sheet.command.format-cell-rich-text`。编辑态 operation 核对工作簿/工作表、合并内容主格、保护、源会话、trace、选区和字体值，不直接发布内容；最终提交由[共享单格输入命令](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-input-command.ts)合成一次内容事务。
- [sheet-cell-edit.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-edit.ts) 的 `editSheetCellInlineText()` 保持已有 raw string 的文本类型，例如 General 下的文本 `123` 不因局部字体变更重新解析为数字。真正数字、Boolean 和公式不接受局部文字格式。原数字格式继续保留。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的 Ribbon、快捷键、公式栏与单元格输入共享命令及选区，读取运行混合态，保留弹层期间的输入会话并恢复原生选区。escape apostrophe 只在临时 Canvas view 前置一个无来源 ID 的运行，不写入领域内容。
- [Facade](../../packages/office-studio/src/api/facade.ts) 的 `FRange.setRichTextStyle(start, end, style)` 使用 raw-cell UTF-16 offsets，要求单格或完整合并范围，校验字段/代理对边界后仅修改字体运行；保留 value、dataType、display、numberFormat 和 metadata，不通过数字/日期输入解析。返回单格命令结果，经共享注册、权限和内容历史提交。
- 纯字体输入提交在原值、公式和规范化类型不变时不重新阻止已经存在的无效数据；真实内容变化仍按 stop/warning 验证。非法、取消和过期会话不能部分写入，草稿保留。输入 trace、typing、选区、焦点和滚动属于编辑视图，不应写入持久内容。

### 2026-11-27 增大/减小字号续核

本轮核对字体组第一行第 3、4 项：增大字号、减小字号。两项是既有常驻根，grid 位于第 1 行第 6、7 列；未新增插件、上下文或独立 overflow 根。完整字体组仍为 11 根，本记录的行内入口范围仍为 9 根。

固定版 `packages/sheets-ui/src/commands/commands/inline-format.command.ts` 的 `SetRangeFontIncreaseCommand` / `SetRangeFontDecreaseCommand` 在编辑态读取 `menu/utils.ts` 的 `getFontStyleAtCursor()`；后者要求一个 textRun 包含完整选区。匹配到运行时按实际字号做单侧 ±1；跨运行选区匹配不到时，两项都把初始化的 `DEFAULT_STYLES.fs=11` 交给行内格式命令。这也适用于字号相同但属于不同运行的选区，不等同于给每段独立步进。

本地 `ExcelWorkbench.executeRibbon()` 已在 `finishEdit()` 之前交给 `runInlineFontFormatCommand()`，通过 `sheet.operation.format-cell-input` 将既有按钮/别名 ID、会话、UTF-16 选区、输入 trace 和格式草稿送入真实 planner。权限、过期源校验、光标 typing 和最后一次输入提交沿用既有实现；本轮未改工作台、Canvas 或注册协议。

修正 [sheet-cell-inline-format.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-inline-format.ts) 的现有字号步进分支：多个运行与选区相交时直接应用固定默认 11，替代原来的首个相交运行 ±1。单个运行内仍按有效运行字号步进；光标优先采用同位置的 typing 字号，连续增减继续累积 typing 偏好。整格字号、mark、颜色、模型和内容历史接口保持原有契约。

`font.menu.ts` 的字号状态在跨运行匹配不到时不更新；本地继续按全部选中运行显示混合态，应用 11 后显示共同字号。这项状态呈现差异保留并记录；绘图内文字、所有继承/主题字体和纯文本未生成 run 的编辑态仍须完整核验，不能据本轮分支推导全部等价。

补 3 个真实 planner 契约草稿，覆盖按钮和固定别名的局部范围、字号相同/不同的跨运行默认 11，以及连续光标增减和后续插入。另补 classic/grid 两个浏览器用例，读取真实活动字号，核对库输入焦点/原生选区、暂存期间模型与内容历史不变、最终一次提交及撤销重做。全部未运行，未新增运行通过结论。

## 布局、Canvas 与库输入桥

[sheet-cell-rich-text-layout.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-rich-text-layout.ts) 按源区间切运行，使用实际运行字体、宽度和可变行高，供普通显示、旋转/缩小适应及[自动适应尺寸](../../packages/office-studio/src/features/sheet-workbench/sheet-auto-fit.ts)共用。[cell-text.ts](../../packages/office-studio/src/rendering/spreadsheet/cell-text.ts) 输出两渲染后端共用 DrawCommand；justify 分词保留前导空白。可选 glyphHeight 的非法数值回退到行高。

[rich-text-editor.ts](../../packages/office-studio/src/rendering/spreadsheet/rich-text-editor.ts) 产生编辑背景、运行文字、选区、光标及命中几何。按 grapheme 记录 UTF-16 offsets，行内 x 坐标懒测并设 128 项缓存上限；二分命中、视觉行 Home/End/上下移动及 contentWidth/contentHeight 共用此几何。该变更仅说明算法和缓存结构，不是实测性能优化结论。

隐式换行边界新增临时 `upstream/downstream` caret affinity：前一视觉行末尾与后一行开头可以共享 raw UTF-16 offset，pointer hit、光标绘制、Home/End、上下移动和光标可见性保留它们实际所在的行。连续 End 不跳到下一行，Home 返回当前行起点；native 文字/水平移动或会话重置清 view hint。该 hint 不进入领域、输入内容事务或持久历史，公开 offset-only hit 接口保留；复杂脚本和系统选择语义仍待实测。

[SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) 接入临时 richTextEditorCell/editSelection，在原场景末尾画编辑层，按可变行高扩大输入区域并保留冻结 pane 裁剪。填充柄命中优先；富文本鼠标拖选、双击选词、视觉行导航、wheel 与拖选滚动通过库持有的 native selection 回到共享输入桥。会话、工作表、zoom、取消、失活和卸载清理手势；焦点恢复后重新同步 selection。

[rich-text-editor-scroll.ts](../../packages/office-studio/src/rendering/spreadsheet/rich-text-editor-scroll.ts) 集中内容边界和拖选步进：捕获的鼠标停在可见编辑 clip 外时，独立 requestAnimationFrame 继续按实际帧时间滚动并重命中选区，达到内容边界或鼠标回到 clip 内停止；松开、pointercancel、失焦、会话替换、composition、zoom、失活和卸载取消帧。速度随越界距离增加，封顶 1200 CSS px/s，晚帧最多应用 64ms 位移，非法 extent/坐标得到有限结果；这些是本地选择，与上游 ScrollTimer 的边缘 padding/按帧 offset 有差异，需真实交互验收。该循环不依赖 GPU 初始化完成，也不修改内容历史。

编辑器仍使用 `antdv-next` AutoComplete / Input.TextArea，公式栏仍使用库输入。非 composing 时库输入文字、原生选区和 caret 透明，Canvas 绘制运行；select 内层背景也透明。组合输入时 Canvas 只画背景，库输入显示 native 文字/IME。Canvas 命中/选区/拖拽属于编辑器内核，不引入原生可替换表单、Table 或整页 contenteditable。颜色来自现有 ConfigProvider/token，显式运行颜色保留；字体弹层、明暗主题、焦点和真实输入法尚缺浏览器证据。

本批领域/命令/纯布局属于 Sheets feature 与 rendering；Vue 只装配输入、视图及命令。没有改 Worker/WASM 控制 DTO、lease、linear-memory 或 ntex 网络边界；未来大文本布局、批量样式与取消仍须按[计算标准](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)落到对应计算域并实测。

## 用例草稿与未验收范围

| 草稿 | 数量 | 计划覆盖 |
| --- | --- | --- |
| [sheet-cell-rich-text-font.contract.test.mjs](../../packages/office-studio/tests/sheet-cell-rich-text-font.contract.test.mjs) | 7 | 别名/继承、未知字段、显式 false、整格运行更新、颜色恢复、筛除与保护。 |
| [sheet-cell-inline-format.contract.test.mjs](../../packages/office-studio/tests/sheet-cell-inline-format.contract.test.mjs) | 14 | 运行拆分/身份、重置、无效轨迹、时间重放、typing、apostrophe、journal 分支、文本类型、验证和字号步进。 |
| [sheet-cell-rich-text-rendering.contract.test.mjs](../../packages/office-studio/tests/sheet-cell-rich-text-rendering.contract.test.mjs) | 15 | 运行宽度/行高、重复文字、grapheme、绘制/缩放、auto-fit、justify、命中/导航、滚动、主题/IME、懒测及共享偏移的视觉行 affinity。 |
| [sheet-cell-inline-format-api.contract.test.mjs](../../packages/office-studio/tests/sheet-cell-inline-format-api.contract.test.mjs) | 7 | 注册/API、取消/no-op、目标隔离、合并、权限、内容历史、内部快照重开与过期会话。 |
| [sheet-cell-rich-text-scroll.contract.test.mjs](../../packages/office-studio/tests/sheet-cell-rich-text-scroll.contract.test.mjs) | 5 | clip/内容界限、方向/距离/速度上限、静止鼠标连续步进/反向、帧时间/晚帧、非法值和隐藏 clip。 |
| [sheet-cell-inline-format.browser.test.mjs](../../packages/office-studio/tests/sheet-cell-inline-format.browser.test.mjs) | 12 | 真实库输入/Ribbon、公式栏、弹层焦点、Canvas 拖选/导航、数字外观文本、主题/替换宿主、静止鼠标滚动/取消后重入、重复 End/Home 行边界及 classic/grid 字号步进。 |

共 **60 个用例草稿，全部未运行**。[浏览器样本](../../packages/office-studio/tests/fixtures/sheet-rich-text-input.html) 保留 A1–D1 原样本，新增 E1 `A🙂BC` 混合字号、F1 `aa\nbb` 可变行高、E2 General 文本 `123`、G1 两段混合字号长文本及 H1 隐式换行样本。仅静态阅读了本批代码和草稿；没有声称这些用例通过。

- 默认/行列/theme/direct 组合、目录引用和选择性 reset、空白格来源与缓存重开基础已续作，见[默认样式证据](SHEET-04-default-style-inheritance.md)及其 58 个未运行草稿；family/size/mark null 恢复有效继承，无上层来源才回退 Aptos/11。本地默认与固定 Arial/11、完整范围主题/主题字体及 Office 映射差异仍保留。
- 富剪贴板、局部链接编辑、全部扩展字体属性、绘图内文字和跨宿主权限需要继续闭环，不以运行字段保留替代功能验收。
- 编辑态暂取消 justify、rotation、shrink，超长 center/right 行临时左对齐；复杂 shaping/BiDi、混合字号 baseline、wrap caret affinity、连续拖选滚动、popup 焦点和真实中文 IME 候选定位尚未完整核验。
- 两个渲染后端、实际浏览器输入、历史重开、真实 XLSX/Excel 独立往返及来源字体/扩展保真均未运行；内部快照能力不证明 Office 互操作。
- 长文本布局、整格大范围格式、Worker/WASM 取消/传输/内存、首屏/滚动/编辑/缩放没有同数据集测量，性能保持未验收。

后端及原生文件增量需求登记到[后端需求](../BACKEND_REQUIREMENTS.md)；Sheets / Docs / 全部 Pro 完整目标保持有效。

