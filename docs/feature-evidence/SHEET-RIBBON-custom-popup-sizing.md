# Sheets 自定义 Ribbon 弹层尺寸

更新日期：2026-11-27。涉及 COMMON-01/04、SHEET-04/06/13/25；状态：未验收，浏览器测试草稿未运行。

本批修正边框斜线预览和切片器快速样式预览的尺寸约束，让边框主按钮显示已接受的预设图标，并补齐公式栏展开按钮高度与数字格式 classic/grid 字号。续批统一边框颜色和线型字段的 small 高度，约束快捷下拉动作及数字格式菜单的长标签；后续静态核对移除颜色箭头的单侧占位边框，补齐颜色与快捷下拉的禁用状态。保留既有 `antdv-next` 控件、主题 token、入口注册和命令转发。

## 来源与清点

固定上游：Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地 `E:\rust_test\univer`。

| 来源 | 本批依据 |
| --- | --- |
| `packages/ui/src/views/components/ribbon/RibbonGrid.tsx` | 两行 grid 按真实 placement 分配空间；窄宽度时先调整密度，再把无法容纳的分组放进 More |
| `packages/sheets-ui/src/views/border-panel/BorderPanel.tsx` | 边框预设在紧凑按钮中呈现，边框颜色/样式共用状态，明暗主题分别处理 hover |
| `packages/sheets-ui/src/menu/border.menu.ts`、`packages/sheets-ui/src/views/border-panel/interface.ts` | `CellBorderSelectorMenuItemFactory` 在边框命令执行后按 pen type 更新主图标；15 个预设与 12 个线型沿用现有清点 |
| [FormulaBar.tsx](../../../univer/packages/sheets-ui/src/views/formula-bar/FormulaBar.tsx) | 公式栏分别为 h-7/h-20（28/80px），箭头容器使用 h-full 并居中 |
| [MoreNumfmtType.tsx](../../../univer/packages/sheets-numfmt-ui/src/views/components/MoreNumfmtType.tsx)、[ToolbarItem.tsx](../../../univer/packages/ui/src/views/components/ribbon/ToolbarItem.tsx) | 格式名称为 text-sm（14px）；grid selector 使用 text-xs（12px），箭头为 size-3（12px） |
| [DropdownPresets.tsx](../../../univer/packages/sheets-data-validation-ui/src/views/components/DropdownPresets.tsx)、[dv.menu.ts](../../../univer/packages/sheets-data-validation-ui/src/menu/dv.menu.ts) | 360px 面板、三列、6px gap、80px 最小卡片高度；预设文字允许换行，明暗主题均有 hover。快捷入口复用 workbook、worksheet 与 range 的编辑权限状态 |
| 本项目 [docs/llm.md](../llm.md) | Button 支持 classes/styles 的语义节点定制；默认 size 为 medium；嵌套 ConfigProvider 中未改动的 token 继承父主题 |
| 本项目 [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) | 自定义 inline/popup 根、classic/grid 位置和命令的实际注册 |
| 本项目 [RibbonItem.vue](../../packages/office-studio/src/platform/ribbon/RibbonItem.vue)、[RibbonOverflowPanel.vue](../../packages/office-studio/src/platform/ribbon/RibbonOverflowPanel.vue) | 自定义组件与真实工具栏/更多菜单共用弹层入口，读取和转发既有命令 |

本轮静态清点 Sheets 实际注册的自定义控件：3 个 inline 根、3 个 popup 根；其中常规分类根 5 个，上下文根 1 个。更多菜单复用这 6 个根，没有额外业务根；生产工作台未另外提供 Ribbon item slot。公式栏另计 1 个产品 Header 区域，不计为 Ribbon 根。字体和字号使用共享 Ribbon Select。

| 入口 | 组件/位置 | 命令与状态 |
| --- | --- | --- |
| `sheet.menu.text-color` | `SheetColorControl`，开始/字体，inline，grid 第 2 行第 7 列 | `sheet.command.set-color`；选区颜色和上次颜色 |
| `sheet.menu.fill` | `SheetColorControl`，开始/字体，inline，grid 第 2 行第 6 列 | `sheet.command.set-background`；选区填充和上次填充 |
| `sheet.command.set-number-format` | `SheetNumberFormatControl`，开始/数字，inline，grid 第 1 行第 1 列，跨 4 列 | `sheet.command.set-number-format`；格式值与 pattern |
| `sheet.menu.borders` | `SheetBorderSelector`，开始/字体，popup，grid 第 2 行第 5 列 | `sheet.command.set-borders`；preset、颜色、线型与格式权限 |
| `sheet.operation.quick-data-validation.dropdown` | `QuickDropdownPicker`，插入/单元格，popup | `sheet.operation.quick-data-validation`；预设与校验能力 |
| `sheet.menu.slicer-quick-style` | `SlicerQuickStylePicker`，切片器格式/样式，上下文 popup，grid 第 1 行第 1 列，跨 2 行 | `sheet.command.update-slicer-presentation`；对象 ID、styleOwner、revision 与配置能力 |
| 公式栏 | `SheetFormulaBar`，产品 Header | 既有名称导航、输入事件、确认/取消与编辑会话 |

完整入口清点仍以 [Sheets 固定源码入口清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)和 [顶部工作区清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)为准。这些计数只说明本次样式审阅范围，不证明 Sheets/Pro 功能闭环已完成。

## 实际修改

- [SheetBorderIcon.vue](../../packages/office-studio/src/components/spreadsheet/SheetBorderIcon.vue)：复用原预览的边线与主题样式，按 `size` 明确外框尺寸。五种单斜线/组合斜线统一为 block 与 100% 宽高的 SVG，填满默认 18×18px 外框的 16×16px 内容区；单斜线不再依赖固定像素伪元素。
- [SheetBorderSelector.vue](../../packages/office-studio/src/components/spreadsheet/SheetBorderSelector.vue) 和 [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 共用该图标；主按钮从既有 `borderPreset` 读取已接受的预设，保持原 repeat 参数。颜色字段高度从固定 28px 改为当前 `token.controlHeightSM`，与既有 small 线型 Select 一致。共享 [RibbonItem.vue](../../packages/office-studio/src/platform/ribbon/RibbonItem.vue) 的主图标尺寸规则只作用于直接 SVG 子节点，避免把预览内部 SVG 拉到外框尺寸。
- [SlicerQuickStylePicker.vue](../../packages/office-studio/src/components/spreadsheet/SlicerQuickStylePicker.vue)：样式列表和预览新增 `box-sizing: border-box`。列表的 3px padding 计入视口宽度约束，预览的 4px padding 与 1px border 计入 100% 尺寸，避免侵入按钮的 8px padding。
- [SheetFormulaBar.vue](../../packages/office-studio/src/components/spreadsheet/SheetFormulaBar.vue)：`.formula-expand` 从固定 27px 改为 100% 高度。栏高从 28px 展开到 80px 后，Button 随实际内容区填满，箭头仍用既有 flex 居中；输入、展开状态、确认/取消与焦点回调不变。
- [SheetNumberFormatControl.vue](../../packages/office-studio/src/components/spreadsheet/SheetNumberFormatControl.vue)：`.sheet-number-format-trigger` 的 classic 字号/行高改为 14/20px，`.is-grid` 明确保留 12/16px。菜单首列增加 `min-width:0` 和单行截断，preview 使用 `flex:0 0 auto` 保留样例宽度；原 28/24px 控件高度、12px 箭头、caption 截断、弹层和命令路径不变。共享 Ribbon 和两个 ColorControl 的在途改动由主线负责。
- [QuickDropdownPicker.vue](../../packages/office-studio/src/components/spreadsheet/QuickDropdownPicker.vue)：两个动作 Button 通过 `classes.content` 为实际文本节点设置收缩和 ellipsis，显式 `aria-label` 保留完整名称；SVG 固定为不收缩的 16×16px。预设卡片、insert 参数及保留焦点的 mousedown 行为不变。

下表记录专用控件的当前约束、保留差异及待验证项：

| 文件/选择器 | 当前实现 | 保留差异与待验证项 |
| --- | --- | --- |
| [QuickDropdownPicker.vue](../../packages/office-studio/src/components/spreadsheet/QuickDropdownPicker.vue)：`.quick-presets / .quick-preset / .preset-values / .quick-action` | 面板上限 360px、三列 6px gap、卡片最小高 80px 与固定来源一致；预设文字已有 normal/anywhere。动作的实际文本节点已允许收缩并截断，完整名称和固定 16px 图标保留 | 卡片 padding 本地 10px、上游 12px 的差异继续记录；长标签、图标边界及 hover 前后尺寸草稿未运行 |
| 数字格式菜单：`.sheet-number-format-row / .sheet-number-format-preview` | 保留 28px 菜单行和 24px gap；首列已设 min-width:0/ellipsis，预览已禁止 shrink | 长标签截断、预览完整和两列边界的真实测量后置 |
| [SheetBorderSelector.vue](../../packages/office-studio/src/components/spreadsheet/SheetBorderSelector.vue)：`.border-preset / .border-color-field / .border-style-field` | 24×24px 预设在 grid 中居中；颜色字段读取同一 controlHeightSM，线型 Select 保持 small。选中背景与库 hover 分别控制，hover 时库状态规则优先 | 保持现有颜色与线型命令；明暗主题下 24/30px token 对应的实际字段矩形草稿未运行 |
| [SlicerQuickStylePicker.vue](../../packages/office-studio/src/components/spreadsheet/SlicerQuickStylePicker.vue)：`.slicer-quick-styles / .slicer-quick-style / .slicer-quick-sample` | Button 高 66px、padding 8px；样例已有 border-box，100% 尺寸对应内容区；3px 列表 padding 容纳 2px outline+1px offset，完整样式名由 Tooltip 承载 | 本轮没有新的明确尺寸修项；default Button 的 hover/disabled 仍需实际主题测量，Pro 交互验收沿用 SHEET-25 |

已静态核对安装的 antdv-next 1.8.2 `dist/button/buttonHelper.js`：文本节点会生成带 classes.content 的 span，普通预览 VNode 保持原节点。数字格式 caption/箭头、边框预览和切片器样例按真实子节点布局，不存在统一的 `.ant-btn-content` 包装。五个专用组件均从当前 ConfigProvider token 取得颜色；没有发现额外明暗主题选择器覆盖它们。按钮的自定义背景/颜色仍与库状态选择器参与同一 CSS 级联，实际 disabled/hover 结果未作浏览器验证。

本项目用 `antdv-next Button/Form/Select/Tooltip` 和 Canvas 输入桥承载现有功能，预览为本地 CSS/SVG；不复制上游 React 控件或声称像素一致。切片器 Pro 的完整功能证据、模型/历史/保存协议沿用既有 SHEET-25 记录，本批未新增功能完成结论。

模块归属为 Vue 视图层，计算保持主线程同步；没有改变模型、命令、内容历史、解析/写出、Worker/WASM 协议、计算域或内存所有权。

## 检查与待验收

新增 [sheet-ribbon-popup-layout.browser.test.mjs](../../packages/office-studio/tests/sheet-ribbon-popup-layout.browser.test.mjs)，classic/grid × toolbar/More 共 4 个尺寸计划用例；后续禁用状态续核增加 1 个真实组件草稿，当前共 5 个计划用例。夹具使用完整 `createSheetRibbon`、`WorkbenchRibbon`、`ConfigProvider`，通过宿主宽度触发实际溢出；全部未运行。

草稿测量全部 5 个斜线 SVG、预览宿主内容区、真实 Button 矩形，以及切片器列表视口宽度和样式预览。预览断言扣除宿主实际 computed padding/border，hover 前后比较按钮尺寸；明暗主题分别读取实际 token。续批仍复用这 4 个用例：切换 controlHeightSM 为 24/30px，比较边框颜色字段与实际 small Select 高度；通过真实翻译覆盖注入长中文标签，检查快捷动作的完整可访问名称、文本截断、16×16px 图标和 hover 前后内容区边界，并检查数字格式首列截断、preview 完整及两列不重叠。全部断言未运行。UI 命令记录用于检查仅打开/悬浮未触发编辑，不替代真实模型、撤销重做或保存验证。[sheet-border-format.browser.test.mjs](../../packages/office-studio/tests/sheet-border-format.browser.test.mjs) 的真实 Studio 草稿补 classic/grid 主图标从 all 到 outside、API 同步与保护邻格拒绝后保留 all 的断言。

复用 [sheet-number-format.browser.test.mjs](../../packages/office-studio/tests/sheet-number-format.browser.test.mjs) 的已有真实 Studio 草稿：在公式栏区域编辑用例中检查展开为 80px、Button 高度等于扣除实际边框后的内容区、箭头纵向中心、靠近按钮底部的点击能折叠回 28px，随后重新展开到 80px，再继续原有展开态 caret/IME/确认链路；classic/grid 的格式状态用例增加实际 computed 字号与行高断言。这些新增断言未运行。

前批两个修改组件的限定 `git diff --check` 与引用检查已通过。本轮共享图标、主按钮接线及三个组件小修已纳入主线统一语法、引用、文档链接/锚点与空白检查，范围和结果见[顶部控件静态收尾](../UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing)；测试、类型检查、前端构建、Rust/WASM 检查、dev、浏览器和基准全部后置。

前轮公式栏、数字格式及对应浏览器草稿按仓库 core.autocrlf=true 的限定 `git diff --check` 通过。Sheets 子任务收尾时复查的 24 个文件链接均存在，指向 19 个不同文件；追加的小修、草稿和证据变更随后已完成主线统一静态收尾。没有运行测试、类型、构建或浏览器。

待统一验证：这些草稿的真实矩形与 hover 命中、全部入口箭头/标签居中、popup 视口边界、主题切换、禁用态、键盘焦点与关闭后的恢复、Canvas/公式栏输入法、两种渲染后端和文件独立重开。性能与大样本未实测，保留未验收，不声称性能改善。

## 2026-11-27 下拉箭头续核

本轮按现有分组顺序及组内行列核对 3 个 inline 根和 3 个 popup 根，公式栏另计；字体、字号与其他常驻入口的共享控件由主线核对。

| 实际位置 | 专有控件核对结果 |
| --- | --- |
| 开始 / 字体，第 2 行第 5 列：边框 | 顶部箭头由共享 RibbonItem 渲染；15 个预设、颜色字段、small 线型 Select 保留既有尺寸修复和 token。预设与字段均接收 disabled |
| 开始 / 字体，第 2 行第 6、7 列：填充与文字颜色 | 两个根共用 SheetColorControl。16px 箭头 Button 原样式指定了分隔侧 border-inline-start，安装包 text Button 另有对称透明边框，不能据此前者单独认定 0.5px 偏移。本轮统一 border:0 与不占布局、不参与命中的绝对定位分隔线；grid 和边框颜色字段的 select 模式隐藏分隔线。主按钮与箭头禁用时明确读取 --ribbon-disabled；实际中心与 hover 仍待测 |
| 开始 / 数字，第 1 行第 1 列、跨 4 列：数字格式 | 保留 border-box、28/24px 高度、12px 不收缩箭头、caption 与菜单长标签截断；disabled 已传入 Button/Menu 并在状态变化时关闭弹层。实际样式级联与命中仍待浏览器测量 |
| 插入 / 单元格：快捷下拉 | 顶部箭头由共享 RibbonItem 渲染。本轮为 QuickDropdownPicker 声明可选 disabled，将已传入的状态送到 3 个预设及编辑/清除共 5 个 Button；保留文字截断、完整可访问名称和命令值 |
| 切片器格式 / 样式，第 1 行第 1 列、跨 2 行：快速样式 | 顶部箭头由共享 RibbonItem 渲染；66px Button、border-box 预览、Tooltip 完整名称和 disabled 保留。样式颜色继续由主题与已有 slicerPalette 投影 |
| Ribbon 下方：公式栏 | 保留 28/80px 栏高及 100% 高度展开 Button，箭头 flex 居中；编辑确认/取消、焦点和 IME 桥保持既有行为 |

固定 [ToolbarItem.tsx](../../../univer/packages/ui/src/views/components/ribbon/ToolbarItem.tsx) 的紧凑箭头为 16px 容器、12px 图标并居中；本项目共享 [RibbonItem.vue](../../packages/office-studio/src/platform/ribbon/RibbonItem.vue) 已使用不占布局的分隔线，本轮专有颜色箭头采用同一处理。颜色的深层普通文字选择器比库 `.ant-btn:disabled` 更具体，新增禁用规则使用 [ribbon-theme.ts](../../packages/office-studio/src/platform/ribbon/ribbon-theme.ts) 已提供的 colorTextDisabled token，避免普通文字颜色覆盖禁用态。

库的 variant 样式对各侧应用 lineWidth / lineType，text 模式只将颜色改为 transparent；旧颜色箭头的分隔侧规则不等于其他各侧宽度为零。前述共享 Ribbon 在其自身 border:0 上增加单侧 1px 的静态中心差，不能直接当作颜色控件的实测结果。本次没有运行 CSS 或 DOM 测量，颜色箭头的实际中心仍由统一浏览器用例验收。

快捷下拉的 disabled 原来作为未声明属性落在根 div 上，未到达内部 Button。[RibbonItem.vue](../../packages/office-studio/src/platform/ribbon/RibbonItem.vue) 继续统一传递 disabled、阻止不可用命令并关闭失效弹层；本轮只补专有组件的接收与控件状态，没有增加第二套权限判断。安装的 antdv-next 1.8.2 [Button](../../node_modules/antdv-next/dist/button/button.js) 与 [variant 样式](../../node_modules/antdv-next/dist/button/style/variant.js) 确认了 disabled 优先级及 colorTextDisabled 的使用。Context7 查询因月度配额不可用，未取得在线库文档，依据来自本地文档和安装源码。

既有弹层草稿新增 1 个真实 QuickDropdownPicker 用例，在明暗主题下检查全部 5 个原生 Button 的 disabled 属性，调用原生 click 时没有 insert，解除禁用后选择仍发布原 action。颜色箭头继续由既有 [Ribbon 组件草稿](../../packages/office-studio/tests/ribbon-component.browser.test.mjs) 检查 LTR/RTL、classic/grid、toolbar/More 的中心与命中；水平中心容差为 0.05px。两个草稿文件均未运行，颜色禁用态的实际样式仍未验收。

本轮实际修改 2 个生产组件、1 个既有草稿及本证据，视图接口只为 QuickDropdownPicker 增加可选 disabled；模型、命令、内容历史、保存/重开及 Worker/WASM 协议不变。此批追加改动不属于前述主线静态快照：限定两个生产组件的 git diff --check、四个文件的尾部空白检查、草稿文件及其内嵌模块的 node --check 均通过，本文件 33 个本地链接均存在；没有运行测试、类型检查、构建、Rust/WASM、dev、浏览器或基准。连续区域、表格创建等较大方案继续保留，未实施。真实箭头中心、可见内容边界、hover 矩形、明暗主题、禁用态、焦点恢复和输入法保持 **未验收**。

主线随后回读颜色预设，将 Button 上的原生 title 迁移到现有 Ant Tooltip，支持 hover 和 focus；保留原 Button、aria-label、鼠标按下与 apply 路径，预设继续使用原九列网格。共享 Ribbon 的既有 4 个颜色预设用例同步检查两产品无原生 title、明暗主题下的 Ant Tooltip hover/focus 与原尺寸边界，声明数不变，全部未运行。此项追加改动将在两子任务文件稳定后的统一静态收尾中检查。

主线最后统一静态收尾已完成并覆盖上述追加小修：86 个 TS/JS、27 个 Vue、28 个 CSS/样式块、15 段夹具脚本、1420 处本地引用，及 30 份 Markdown 的 926 条本地链接、71 处锚点通过；diff 与新增文本尾空白检查通过。没有执行测试、类型、构建、Rust/WASM、dev、浏览器或基准，真实交互和完整功能仍未验收。

## 2026-11-27 颜色按钮 hover/active 边框续核

按数字格式、边框面板、颜色、快捷下拉的顺序回读四个专有控件及上述来源，未重复前批的箭头分隔线、字段高度、预设 Tooltip 与 disabled 修复。本批仍复用同一组根，常驻、扩展、溢出和上下文根增量均为 **0**；没有增加模型、业务命令或内容历史接口。

安装版 [button/style/variant.js](../../node_modules/antdv-next/dist/button/style/variant.js) 明确在非禁用按钮的 hover/active 状态重新赋完整 `border`，包括 `lineWidth / lineType`；text variant 只将边框颜色设为透明。颜色主区/箭头原有 `border:0` 位于普通 scoped 规则，颜色 hover 规则只写背景，不能阻止库的状态规则恢复边框。这与前批“分隔侧是否造成 0.5px 中心差”的问题不同；本批依据是已读的 CSS 状态赋值，尚未浏览器测出旧矩形的具体变化值。

[SheetColorControl.vue](../../packages/office-studio/src/components/spreadsheet/SheetColorControl.vue) 只为主按钮和箭头按钮增加 `styles.root: { border: 0 }`。安装版 [Button.js](../../node_modules/antdv-next/dist/button/Button.js) 将该语义样式合并到真实按钮的 inline style，固定静止、hover、active 和 disabled 时的零边框，保留原尺寸、padding、绝对分隔线、12px 图标、主题 token、输入法及命令路径。没有修改共享 Ribbon 或其他颜色组件。

数字格式的 28/24px border-box、12px 箭头与菜单两列约束，边框的 24px 居中预设和同 token 高度字段，快捷下拉的三列/80px 最小高度、10px padding 及动作文本截断均保持现状；这些按钮未在普通状态移除库边框，未发现同一边框恢复导致的新尺寸修项。快捷卡片与固定上游 12px padding 的差异继续保留。打开/禁用的控件状态与已读监听、事件判断按前批记录，明暗主题的实际 CSS 级联、内容矩形、命中及焦点仍未验收。

[既有 popup 布局草稿](../../packages/office-studio/tests/sheet-ribbon-popup-layout.browser.test.mjs) 的 classic/grid × toolbar/More 四个用例追加真实颜色主区/箭头检查：明暗主题下分别进入 hover 和 pointer active，确认四侧边框为 0，比较前后完整 Button 矩形；按下后移出再释放，避免执行颜色命令。原用例和 QuickDropdown disabled 用例数量不变，合计仍为 **5 个计划用例**，全部未运行。

本批限定静态检查通过：4 个 Vue 的 SFC 结构与 TypeScript 脚本共 0 条语法诊断；既有浏览器草稿及其内嵌模块 `node --check` 通过；本文件 37 个本地链接及既有 `top-ribbon-control-sizing` 锚点存在，限定 diff 与证据尾空白检查通过。没有运行测试、类型检查、构建、Rust/WASM、dev、浏览器或基准。箭头中心、真实 hover/active 矩形、disabled 外观、主题切换、输入法和保存重开继续保持 **未验收**，未提出新的功能完成或性能结论。

<a id="number-format-open-state"></a>

## 2026-11-27 数字格式打开态与边框续核

[SheetNumberFormatControl.vue](../../packages/office-studio/src/components/spreadsheet/SheetNumberFormatControl.vue) 复用前述安装版 Button 的 `styles.root`，为唯一触发按钮固定 `border:0`；保留 classic/More 的 28px 与 grid 的 24px 高度、6px 对称 padding 和 12px 不收缩箭头。打开态按真实 `aria-expanded` 使用现有激活背景与主色，并明确覆盖 hover/active；disabled 使用统一文字 token。Tooltip 支持 hover/focus，延迟 0.6 秒，打开弹层时关闭提示。原 Button/Menu、命令、模型、选区、历史和保存路径保持现有接口。

固定 [ToolbarButton.tsx](../../../univer/packages/ui/src/views/components/ribbon/ToolbarButton.tsx) 的激活状态及 [ToolbarItem.tsx](../../../univer/packages/ui/src/views/components/ribbon/ToolbarItem.tsx) 的 selector 居中/尺寸是交互依据；透明边框恢复的依据仍为安装版 [variant.js](../../node_modules/antdv-next/dist/button/style/variant.js)。这些源码赋值与本地修正尚不能证明旧按钮的具体偏移量或最终真实矩形。

[既有 popup 布局草稿](../../packages/office-studio/tests/sheet-ribbon-popup-layout.browser.test.mjs) 的四个 classic/grid × toolbar/More 用例追加数字格式按钮检查：明暗主题下打开态背景、四侧零边框、hover 前后尺寸不变、箭头位于按钮纵向中心及内容不越界。仍共 5 个计划用例，未运行。常驻、扩展、溢出和上下文根增量为 0；本次静态检查另行收尾，实际 hover/active、禁用外观、主题、焦点恢复、IME 与文件重开保持未验收。

## 2026-11-27 边框与切片器提示续核

安装版 [tooltip/index.js](../../node_modules/antdv-next/dist/tooltip/index.js) 的 mergedTrigger 默认仅 hover，当前 ConfigProvider 没有统一 focus 配置。[SheetBorderSelector.vue](../../packages/office-studio/src/components/spreadsheet/SheetBorderSelector.vue) 和 [SlicerQuickStylePicker.vue](../../packages/office-studio/src/components/spreadsheet/SlicerQuickStylePicker.vue) 各补一行 `trigger=['hover','focus']`，使现有完整名称同时响应键盘聚焦。主线在原四个 popup 布局用例中追加 hover → 移出 → focus → blur 的提示检查，并比较按钮/预览真实边界，仍未运行。没有改变预设、布局、接口或入口数量。

复查 [v1.6.0 发布 slicer-ui](../../.research.local/univer-1.6.0/sheets-slicer-ui/lib/es/index.js) 时区分主 Ribbon gallery 与设置字段：gn 的 quickStyles 为 7 列、8px gap、12px padding，_n tile 为 72px 高/4px padding；主 gallery 使用 w-md，实际宽度未核实，不能按 maxWidth.md 推断为 448px。wn 的 settings 样式字段 popup 才是 w-80/3 列/70vh。本地 320px 上限/3 列/66px 高/8px tile padding 与两类型共用三条带预览仍属于现有差异；公开 vn 分别展示 Slicer 的 selected/unselected 与 Timeline 的 label/time-block。此次只补提示触发，没有调整 gallery 或宣称布局/预览已对齐，完整样式行为、主题、焦点和重开仍待验收。

<a id="slicer-style-preview-continuation"></a>

## 切片器与时间线的实际样式预览续修

2026-11-27 后续复查同一发布包的 `vn / yn`：切片器预览包含 header、selectedItemWithData 和两行 unselectedItemWithData；时间线包含 header、标签、五段交替 periodLabel1/2 和三段选中、两段未选中 time-block。主 gallery 的 activeStyle 使用对象 style、文档默认样式、首个内置样式的顺序，不将缺省 style 当作无选中项。该包 `lib/index.css` 未找到 `univer-w-md` 的宽度声明；固定 [共享 Tailwind 配置](../../../univer/common/shared/tailwind/tailwind.config.ts) 也未补此 width，宿主最终宽度继续未核实，不使用 maxWidth.md 替代依据。

[SlicerQuickStylePicker.vue](../../packages/office-studio/src/components/spreadsheet/SlicerQuickStylePicker.vue) 复用 [getSlicerStyle / listSlicerStyleNames](../../packages/office-studio/src/features/slicer/slicer-style-catalog.ts) 和 [resolveSlicerRenderProfile / slicerRenderRegionCss](../../packages/office-studio/src/features/slicer/slicer-style-rendering.ts)，按实际类型绘制上述预览结构。容器、标题、成员、时间标签/块分别使用现有区域的填充、文字、边框和对齐投影，包含作者自定义样式；预览不再只消费三个 palette 颜色。默认值按钮显示文档默认预览，当前生效的命名样式独占选中态；默认值按钮仍可清除显式 style。

样式块改为 border-box、72px 高、4px 内边距和满列宽，预览使用有界 flex 区域，保持实际 Button、Tooltip 和原 selected outline。主 gallery 继续保留本地 320px 上限、3 列与 3px 外内边距，和上游 7 列/12px padding 有明确差异；没有把这部分标为完全对齐。时间线的可访问组名改为“时间线快速样式”。14 个切片器内置样式、12 个时间线内置样式、作者目录去重和额外默认重置动作保持原覆盖，常规、扩展、溢出及上下文根增量均为 0。

[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的 Canvas 与 Ribbon 两处 styleOwner 适配使用现有 `toRaw(workbook.value)`，将响应式 UI 状态转换到纯模型投影边界；现有目录会对作者 definition 使用 structuredClone，不向它传 Vue Proxy。没有新增领域引擎、复制工作簿、改变 DTO、内容事务、权限、历史或 Worker/WASM 协议；仍由既有 `sheet.command.update-slicer-presentation` 携带 slicerId、expectedRevision 和 style patch。

[既有 popup 浏览器草稿](../../packages/office-studio/tests/sheet-ribbon-popup-layout.browser.test.mjs) 将原四个布局用例的样式块高度改为 72px，并补快捷下拉预设的 12px padding、6px 行距、零边框及 hover 前后真实边界；新增两项类型用例核对作者区域色、默认选中、内置/作者/重置数量、明暗主题、内容边界和真实命令参数。[既有真实工作台草稿](../../packages/office-studio/tests/sheet-slicer-drawing.browser.test.mjs) 新增两项类型用例，覆盖响应式工作簿上的作者默认样式、顶部选择、一次命令历史、撤销/重做、保存重开与默认重置。未新增草稿文件，全部未运行；静态收尾另记，实际 hover/active、主题、焦点、IME、Office 互操作及完整样式功能继续未验收。
