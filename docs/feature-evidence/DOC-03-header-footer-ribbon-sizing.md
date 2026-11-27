# Docs 页眉页脚 Ribbon 尺寸与主题

功能：DOC-03，关联 COMMON-01 / COMMON-04。更新日期：2026-11-27。
固定基线：`dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地 `E:\rust_test\univer`，本轮已核对 HEAD。
状态：未验收；组件改动与浏览器测试草稿已落盘，运行验证统一后置。

## 来源与入口口径

| 来源 | 本轮核对的契约 |
| --- | --- |
| [header-footer-ribbon.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/header-footer-ribbon.ts) | 上下文分类 order 90；插入、导航、选项、位置四个自定义控件和关闭入口，共五个上下文根；四个控件 grid 均占两行，label 的 hoverable 为 false |
| [DocHeaderFooterRibbonControl.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/header-footer/panel/DocHeaderFooterRibbonControl.tsx) | 位置、选项、导航的紧凑阈值分别为 1100 / 900 / 650；插入不收成紧凑标题；导航使用 text 按钮，两个按钮的 onMouseDown 都调用 preventDefault；位置按标签与数字输入两列排列 |
| [ribbon-schema.ts](../../packages/office-studio/src/platform/ribbon/ribbon-schema.ts) | 本地 classic 按钮高 24px、选择器高 28px；grid 行高 24px、行间距 4px，两行合计 52px |
| [ribbon-theme.ts](../../packages/office-studio/src/platform/ribbon/ribbon-theme.ts) | `useRibbonTheme` 从 `antdv-next ConfigProvider` 取得文字、禁用、主色和 hover token，可应用于独立弹层 |

计数按上下文分类下的顶层根项统计，不把根内按钮、页码菜单、日期对话框或溢出副本重复计为根入口。本轮涉及四个控件根，关闭根不变；未新增常驻或插件扩展根。紧凑弹层和 More 是这些入口的呈现位置，不构成功能完成证明。

本地四个控件通过 `doc.operation.header-footer-control` 转发既有 action 与 `requestScope`；关闭通过 `doc.operation.close-header-footer`。沿用 `HeaderFooterRibbonState`、`HeaderFooterAction` 和既有能力判断，没有变更命令、模型、内容历史或保存协议。

## 静态问题与改动

- `antdv-next 1.8.2` 的 FormItem 在内部 `.ant-form-item-row` 承载标签和输入区。旧样式把两列 grid 放在 FormItem 外壳，内部仍采用 vertical 布局；数字输入区还保留普通表单的最小高度，无法按 Ribbon 的两行约束排列。
- [DocumentHeaderFooterRibbonControl.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonControl.vue) 改用 horizontal / small / 无 colon 的 Form，在真实 row 上设置标签与 92px 输入两列；行、标签和数字输入高 24px，grid 两行间距 4px，classic 在一行中排列。同步调整输入区最小高度及输入文字高度。
- 页码与日期按钮的图标、文字和箭头使用同一个 inline-flex caption，SVG 按块显示；导航改为 text 按钮。按钮与 Checkbox 采用对应行高、字号和行高，hover、禁用与 focus-visible 使用共享 Ribbon token。
- 后续焦点小修为“转到页眉”“转到页脚”两个现有 Ant Button 补 `@mousedown.prevent`，对齐固定来源的按下事件，阻止浏览器先将焦点移到导航按钮。点击仍转发原 navigate action、目标 area 和 requestScope；既有只读可导航能力与当前区域按钮禁用规则保留。
- [DocumentHeaderFooterRibbonGroup.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonGroup.vue) 区分 classic、grid 与 overflow 宿主：classic 高 28px，grid 填满分配高度，overflow 随完整内容计算高度。紧凑标题与 12px 箭头使用同一个居中 caption；hover 背景仅绘制在实际按钮，禁用保持透明，焦点轮廓收在按钮内部。
- 紧凑 caption 内的标题允许收缩并显示省略号，箭头保持 12px；长名称受实际按钮宽度约束，完整可访问名称仍保留。
- 两个组件都应用 `useRibbonTheme`，紧凑弹层中的 Control 自身取得主题 token。所有涉及的按钮、勾选、表单、数字输入、选择、菜单、弹层和对话框继续使用 `antdv-next`，没有新增原生控件或 Canvas 输入桥。

与上游的实现差异：本地通过 `antdv-next Form / FormItem` 承载位置表单，数字输入宽 92px；上游使用标签 grid 和 5rem 输入列。本地 classic / grid 使用仓库既有尺寸 token，紧凑标题保留本地 ChevronDown。上游导航按钮同时受 disabled 限制，本地沿用 `headerFooterActionDisabled` 将导航作为读取 story 的操作，允许只读文档导航；本次只对齐按下事件。这里只记录布局与事件契约，不声称像素相同。

本次归属为 Vue 视图与共享主题适配，计算保持主线程同步；没有新增 Worker/WASM 任务、跨线程消息或缓冲区。领域计算与 Office 写出仍由既有 feature 和 Rust 链路承担。

## 浏览器测试草稿

新增 [document-header-footer-ribbon.browser.test.mjs](../../packages/office-studio/tests/document-header-footer-ribbon.browser.test.mjs)，使用真实 `WorkbenchRibbon`、`createDocumentRibbon` 与 `ConfigProvider`，续修后共九个计划用例，均未运行：

| 用例 | 计划验证 |
| --- | --- |
| classic 完整控件 | 四个控件的真实宿主边界、位置行的标签/输入中心；text 按钮 idle/hover/active/禁用保持零边框；按下页脚导航时保留数字输入焦点，只读时按下页眉导航保留已有 Tab 焦点，抬起后各转发一次命令 |
| grid 完整控件 | 52px 两行边界、24px 行高、28px 行起点差；相同的零边框、两个导航按钮按下焦点与只读导航断言 |
| 紧凑标题 | 标题和箭头居中，鼠标离开按钮后不再命中按钮 hover，宿主背景透明 |
| 紧凑位置弹层 | 两行表单、ArrowUp 转发一次带 requestScope 的边距修改并关闭弹层 |
| 长名称紧凑标题 | 68px/最小 60px 宽度下标题省略、箭头居中且保持在按钮内；移出按钮后 hover 消失 |
| classic/light 与 grid/dark 紧凑状态（两个用例） | 权限、requestScope 和紧凑/完整布局切换关闭旧弹层；恢复权限或紧凑布局不自动重开，重新打开后只对新 scope 提交一次 |
| More 溢出 | 显示完整位置表单，宿主随 52px 内容伸展 |
| 主题与禁用 | 明暗主题 hover 随 token 更新；紧凑按钮 idle/hover/active/禁用保持零边框，禁用透明且不触发命令 |

焦点小修直接扩展现有 classic/grid 用例，复用同一个真实 WorkbenchRibbon 夹具；夹具收到 navigate 后只更新显示 area，使两个导航按钮均可检查。草稿使用 Playwright 的真实鼠标 down/up，按下期间断言焦点没有转移，并在抬起后核对带 requestScope 的导航事件。没有新增夹具、测试文件或计划用例。

这些测试全部未运行。草稿中的命令记录用于检查 UI 转发，不替代真实文档的命令执行、内容历史与持久化验证，也不证明 Canvas 输入桥或 Windows IME 已通过。

## 其余 Docs 自定义顶部控件的补充审阅

本节记录特殊字符公共命令接线前的尺寸静态审阅，当时未发现由已读源码可确认的尺寸、箭头、hover 或文字越界问题，没有增加尺寸代码或测试草稿。后续字符入口、Ant Tooltip 与空值判断的小修见下节。

先读取 `document-ribbon.ts` 与 table / note / structure / latex 子 Ribbon 的真实注册项：除本轮已处理的页眉页脚和另行处理的字体/表格颜色控件外，其余本地自定义 popup 有三个组件，承载四个常规根和十一个 LaTeX 上下文根。弹层内选项及 classic / grid / More 副本不重复计为根入口。

| 审阅组件 / 宿主 | 注册与既有尺寸约束 |
| --- | --- |
| [DocumentListPicker.vue](../../packages/office-studio/src/components/document/DocumentListPicker.vue) | 有序、无序列表两个常规 split 根；弹层通过 `doc.command.change-list-type` 执行。三列 72px 卡片、80px 高度，预览在卡片内；重新编号表单独占整行 |
| [DocumentCharacterPicker.vue](../../packages/office-studio/src/components/document/DocumentCharacterPicker.vue) | 表情、符号两个常规菜单根；本轮改为 `doc.command.insert-special-character`，保留旧字符串调用兼容，接线与缺口见下节。字符格仍使用 minmax(42px, 1fr)、42px 按钮和纵向滚动；分类、自定义字符沿用 Form 与库中输入控件 |
| [DocumentLatexTemplateGallery.vue](../../packages/office-studio/src/components/document/DocumentLatexTemplateGallery.vue) | `document-latex-ribbon.ts` 注册十一类上下文菜单，均转发 `docs-latex.operation.insert-formula-template`。gallery 宽 440px、max-width 100%，桌面三列、≤640px 两列；标签有省略约束。已读 `DocumentLatexPreview.vue`，Canvas 使用 width / max-width 100%、height auto |
| [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 tab-extra | 只读查找替换浮层：查找 Input 可收缩，图标按钮 24px，替换行 ≤650px 换行。示例 `main.ts` 已导入 `antdv-next/dist/reset.css`；已核对其通配 border-box，不能据未局部声明 box-sizing 就认定 padding / border 导致宽度越界 |

固定 OSS 另核对了 `packages/docs-ui/src/menu/schema.ts`、`menu.ts` 的 OrderList / BulletList / EmojiPicker / SymbolPicker 注册与 factory，以及 `views/list-type-picker/Picker.tsx` 的三列预览。LaTeX 十一类发布来源沿用 [既有模板证据](DOC-15-latex-templates-ribbon.md)。表格、note、structure 子 Ribbon 没有发现其他本地自定义 inline / popup 组件；共享 Ribbon 与颜色控件仅核对注册归属，不在本节修改。

仍未核实：实际弹层宽度和滚动、中文/英文及长标签在不同字体下的边界、字符与组合 emoji 显示、模板预览在两列布局下的实际尺寸、触发按钮箭头及 hover 命中、明暗主题、键盘焦点恢复与 IME。这些静态约束不构成浏览器验收证据。

<a id="doc-special-character"></a>

## 特殊字符公共命令与选择器

2026-11-27 补充。关联 DOC-01 / DOC-02 / DOC-22 与 COMMON-03；状态仍为**已落盘、未验收**。本批只完成命令接线和现有草稿，不把完整符号、表情选择器标为完成。

| 固定来源 | 已核对的契约 |
| --- | --- |
| [schema.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts)、[menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts) | Insert 媒体组内表情 order 5 / r1c3，符号 order 6 / r2c3；两个常规菜单根，共用插入命令。未新增插件扩展、上下文或溢出根，More 呈现副本不重复计数 |
| [insert-special-character.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/insert-special-character.command.ts) | ID 为 `doc.command.insert-special-character`，参数 `{ value: string }`；空串返回 false，空格及 NBSP 属于非空输入。读取当前选区、segmentId 和字符格式，以 UTF-16 `value.length` 插入整个值，并继承当前位置 customRange / customDecorations，再复用 InsertTextCommand |
| [symbols.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/views/symbol-picker/symbols.ts)、[SymbolPicker.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/views/symbol-picker/SymbolPicker.tsx) | 固定目录按数组项静态计数：mathematics 70、greek 48、common 32，共 150 项，包含重复 `%` 和 `|a|` / `kg` 等多字符值；三组、10 列、32px 单元、420×340 宿主，有 activeSymbol / aria-pressed 与 keepOpen |
| [EmojiPicker.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/views/emoji-picker/EmojiPicker.tsx)、[emoji-picker-utils.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/views/emoji-picker/emoji-picker-utils.ts)、[generate-emojis.mts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/scripts/generate-emojis.mts) | 七个目录类别加最近项；搜索、肤色、家庭及混合肤色变体、持久最近项。最近项上限 11，持久键为 `univer.ui.recent-emojis` / `univer.ui.emoji-skin-tone`。虚拟列表 10 列、32px 行、24px 分组标题、overscan 3。生成器指定 Unicode Emoji 17.0.0 / CLDR JSON 48.2.0；当前固定工作树没有 `emojis.generated.ts`，未执行生成或统计全部 Emoji 项数 |

实际接线：

- [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 的两个 popup 改用固定 ID。共用 adapter 将 `{ value }` 翻译成现有 `insert-text` 领域对象，保留整个字符串，拒绝空串；旧 `insert-character` 字符串入口仍可用，也拒绝空串，避免空值删除已有选区。
- [DocumentCharacterPicker.vue](../../packages/office-studio/src/components/document/DocumentCharacterPicker.vue) 发出 `{ value }`，函数和自定义插入按钮改按 length 判空，保留纯空格、NBSP 和组合字符。字符格采用 Ant Tooltip 的 hover / focus，Form、Input、Select 和 Button 响应禁用状态；没有新原生控件。
- [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue)、[facade.ts](../../packages/office-studio/src/api/facade.ts) 沿现有 execute 流程透传可选 `sourceCommandId`，没有复制选区、宿主、权限或焦点流程。`DocumentCommand` 没有新增变体。
- 主线已合入并经静态重读的 [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 保留外层固定 ID 和 `params.value`；公共注册直接进入 `commitDocumentEditorCommand('doc.command.insert-text', request)`。它没有再次 dispatch insert-text，因此 UI 和固定 API 进入同一注册及事务，泛型 BeforeCommandExecute / CommandExecuted 的实际 ID 为 special-character。旧私有入口继续使用原 insert-text 事件。空值 flat API 返回 false，空值 envelope 校验后返回 unchanged / 原选区。

本地复用既有输入格式、内容控件锁、Code 输入、表格坐标、历史和保存链路。归属仍为 Vue 交互与命令边界；没有新计算引擎、Worker/WASM 协议、缓冲区或性能优化结论。原注册中的请求 schema 仍按 insert-text 校验，不用 special-character ID 校验其领域请求。

宿主范围：flat `{ value }` 使用当前正文或已激活块宿主输入会话，包括表格单元格；Workbench 沿既有流程把 legacy 页眉页脚范围激活为 story block host。旧 flat API 的 `location.kind === 'story'` 尚未接通，当前返回 false；不能据此验收固定版 segmentId 的全部行为。

新增的五个计划用例全部未运行：

| 现有草稿 | 本批新增覆盖意图 |
| --- | --- |
| [document-toolbar-state.contract.test.mjs](../../packages/office-studio/tests/document-toolbar-state.contract.test.mjs) | 1 项：两个根的固定 ID / 布局、整值 adapter、旧字符串兼容、空值 / 参数拒绝、空格 / NBSP 与禁用 |
| [document-editor-commands.contract.test.mjs](../../packages/office-studio/tests/document-editor-commands.contract.test.mjs) | 3 项：真实 Studio / Facade 的固定事件、输入偏好、UTF-16 光标、一次历史与保存重开；flat 当前选区、empty envelope、不匹配命令、取消、输入代数失效与内容锁；表格单元格坐标、撤销重做和保存重开 |
| [document-structured-workbench.browser.test.mjs](../../packages/office-studio/tests/document-structured-workbench.browser.test.mjs) | 1 项、八个组合场景：classic / light 正文 Code 与 grid / dark 页眉 Code，各插入家庭 Emoji、多字符符号 / 组合字、纯空格及纯 NBSP；从真实 popup 检查固定事件、classic 28px / grid 24px 菜单按钮与箭头中心、hover 背景、42px 字符按钮及 Ant Tooltip、焦点恢复、一次历史、宿主光标及保存重开 |

字符接线子任务收尾时只静态重读共享注册与相关修改，当时未执行语法检查。随后主线已完成 84 个 TS/JS、25 个 Vue、26 个 CSS/样式块、13 段夹具脚本及 1378 处本地引用的静态检查，并核对 30 份 Markdown、907 条本地链接、66 处锚点及空白，见[顶部控件静态收尾](../UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing)。这些数字覆盖此前字符与尺寸批次，不覆盖下节的新改动。测试、类型、构建、浏览器/dev、Rust/WASM、服务和基准仍未运行；实际 hover、Tooltip、主题、布局、Code / story / 表格输入、IME、历史和保存重开继续保留未验收。

完整选择器仍有缺口：本地符号六组和 Emoji 六组是既有有限目录；组件内最近项最多 32，未持久化。搜索、完整目录、肤色及家庭变体、mixed skin tone、activeSymbol / keepOpen / 虚拟列表和实际本地化接线尚未完成；自定义粘贴不替代完整目录。源 customRange / customDecorations 的继承、更多辅助宿主、Office 文件独立重开与 Unicode 字体显示也未核实。完整目录批按另行确认的多文件方案推进，本批没有展开。

<a id="doc-list-color-controls"></a>

## Docs 列表、颜色与表格覆盖层的小修

2026-11-27 续核。来源仍为固定 [menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts)、[ListTypePicker](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/list-type-picker/Picker.tsx)、[EditorColorPicker](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/EditorColorPicker.tsx) 与 [EditorFloatingToolbar](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/components/editor-toolbar/EditorFloatingToolbar.tsx)。有序/无序列表与字体颜色/文本高亮各有两个常驻 split 根；表格颜色包装器与浮动工具栏复用颜色组件。没有新增常驻、扩展或上下文根。

- [DocumentListPicker.vue](../../packages/office-studio/src/components/document/DocumentListPicker.vue) 接收 `RibbonItem` 已传入的 disabled，使六个预设卡片、Form、起始编号及重新编号/继续编号按钮同步禁用；重新编号提交也遵守总禁用状态。卡片 hover 排除禁用按钮，保留三列 72×80px、四行预览、活动项和编号边界。
- [DocumentColorControl.vue](../../packages/office-studio/src/components/document/DocumentColorControl.vue) 将参与布局的单侧分隔边框改为逻辑边缘 pseudo 分隔线，grid 隐藏分隔线，箭头继续为 16px 按钮内的 12px SVG。安装包 text Button 本身含透明边框，本次没有测量旧偏移，不将 0.5px 推断写成 Docs 实测结果。专有深层 color 规则比库的 disabled color 更强，现对主按钮及箭头明确使用 `--ribbon-disabled`，禁用背景透明。
- 色块移除原生 title，改用现有 Ant Tooltip 的 hover/focus；45 项、五行九列和 32px 色块不变。沿用颜色草稿、确认、取消、IME、能力判断与焦点流程，没有新增颜色命令或存储协议。
- [DocumentTableOverlay.vue](../../packages/office-studio/src/components/document/DocumentTableOverlay.vue) 的 handle hover 及 resize hover/focus-visible 排除禁用按钮，避免组合输入期间显示可操作的主色边框或拖拽高亮。未修改表格选择、命中、手势、预览或命令算法。

续读了 Image / Table / ContentControl 覆盖层的真实 template、状态及样式，和 CodeControls / StructuredPopMenus 的现有表单。Image 的按钮/菜单受 disabled 与图片几何能力约束；Code 的 Form 与应用按钮遵守现有能力状态；内容控件禁用或权限变化会取消值弹层。未确认这些组件的新箭头或尺寸问题，没有据猜测修改。所有本次可替换控件仍为 `antdv-next`，主题由 ConfigProvider 和现有 token 承载；没有新增输入桥、Worker/WASM 或性能优化结论。

[document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) 复用真实 Workbench / Studio / ConfigProvider，新增 classic/grid 箭头用例（明色 LTR 与暗色 RTL，两个颜色根，中心容差 0.05px）和权限切换后的列表禁用用例；扩展原颜色权限、Tooltip/色块边界及 More 箭头断言。列表草稿验证六个 72×80px 卡片、关闭后保留的禁用 DOM、零历史与零命令。测试不把隐藏弹层的 DOM 断言当成禁用 hover 的实测。

[document-table-workbench.browser.test.mjs](../../packages/office-studio/tests/document-table-workbench.browser.test.mjs) 新增组合输入期间的明暗主题禁用 handle/resize hover 用例，检查 idle 色和透明背景、模型/历史/命令不变；使用合成 composition，不替代 Windows IME 验收。本小批三个 Vue、两个 JS 草稿、三个 CSS/样式块、两段夹具脚本、44 处本地引用静态解析通过；本文 25 条本地链接及两处引用锚点可解析，限定 diff 空白、六个文本文件尾空白及冲突标记检查通过。没有执行产品模块或运行浏览器草稿；最终文档与草稿的后续小修仍需主线统一收尾。

主线已在两子任务文件稳定后完成最终统一静态收尾，覆盖上述三个控件、两份草稿及主线颜色 Tooltip 补充：86 个 TS/JS、27 个 Vue、28 个 CSS/样式块、15 段夹具脚本、1420 处本地引用，及 30 份 Markdown 的 926 条本地链接、71 处锚点通过；diff 与新增文本尾空白检查通过。所有测试、类型、构建、Rust/WASM、dev、浏览器和基准继续后置，真实 hover、主题、焦点、IME 与编辑闭环仍未验收。

<a id="doc-control-state-border"></a>

## Docs 自定义控件的边框与弹层状态续核

2026-11-27 继续按固定版来源与实际安装的 `antdv-next 1.8.2` 审阅五个组件。库的 `dist/button/style/variant.js` 在启用按钮 hover/active 时重新写入完整 border，优先级高于本地部分 scoped `border: 0`；`dist/button/Button.js` 将 `styles.root` 合入原生按钮 inline style。[本地组件文档](../llm.md) 的 Button styles 与 Popover open API 和安装源码一致。这是源码确认的样式冲突，未执行浏览器测量，不把旧边框描述为已实测的像素偏移。

- [DocumentHeaderFooterRibbonControl.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonControl.vue) 的四个 text 按钮、[DocumentHeaderFooterRibbonGroup.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonGroup.vue) 的紧凑触发器及 [DocumentColorControl.vue](../../packages/office-studio/src/components/document/DocumentColorControl.vue) 的主按钮/箭头通过 `styles.root` 固定零边框，删除对应冗余 scoped border。尺寸、padding、caption、逻辑边缘分隔线及共享主题 token 保留，色块按钮仍使用库的正常边框。
- Group 同步监视 `requestScope`、`visible`、对应控件能力与 `compact`，目标或呈现失效时关闭外层弹层。原先 open 只在转发时清空，权限变化不会让受控 Popover 自动关闭，目标切换会让旧位置弹层读取新 state，紧凑/完整切换还会保留 open。复用 `headerFooterControlDisabled`，导航继续按读取操作允许只读使用。
- [DocumentListPicker.vue](../../packages/office-studio/src/components/document/DocumentListPicker.vue) 与 [DocumentCharacterPicker.vue](../../packages/office-studio/src/components/document/DocumentCharacterPicker.vue) 保留正常库边框，未确认新的尺寸或居中冲突。本批不改变列表预设、编号、自定义整值输入或字符目录；没有据合成事件假设增加禁用守卫。

既有页眉页脚草稿新增两个紧凑状态用例，并扩展完整控件及主题用例的真实 CSS 边框断言。[document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) 复用两个颜色根的 classic/grid、明色 LTR/暗色 RTL 用例，检查主按钮与箭头 idle/hover/active 零边框；权限禁用及 More 箭头也检查零边框。按下/抬起在真实按钮上完成，未提交的状态检查在移出按钮后抬起，保留零模型变化和零历史断言。

本批仍属于 Vue 交互与主题适配，根入口增量为零；命令 ID、模型、内容历史、保存与 Worker/WASM 协议不变。所有可替换控件仍使用 `antdv-next`。静态检查通过：七段脚本（含两段夹具）、三个 Vue 模板、三个样式块、22 处本地导入引用、本文 32 条本地链接及两处锚点；六个文本文件尾空白/冲突标记与限定 diff 检查通过。未运行测试、类型检查、构建、dev、浏览器、Rust/WASM、服务或基准；真实 hover/active、弹层关闭、主题、焦点恢复、IME、编辑及保存重开继续未验收。完整 Picker 多文件方案仍待确认，本批没有实施。

## 列表预览标记列续核

2026-11-27 补充 DOC-05 静态小修。固定版 [Picker.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/list-type-picker/Picker.tsx) 第 99 行将预览标记设为 `w-3 shrink-0 overflow-hidden text-center`。本地原先只有 `min-width: 12px`，多级编号按内容占宽，不能保证同级正文预览线具有相同起点。

[DocumentListPicker.vue](../../packages/office-studio/src/components/document/DocumentListPicker.vue) 仅将标记列固定为 12px，使用 `flex: 0 0 12px` 并裁切溢出文字。`ORDER_LIST_2` 的完整 `1.1.1.` 值继续由既有模型生成并保留在卡片可访问名称中；裁切只作用于预览。六个有序预设、六个无序预设、72×80px 卡片、四行预览及 8px 级别缩进保持现有契约，根入口增量为零。

已读取 [document-toolbar.browser.test.mjs](../../packages/office-studio/tests/document-toolbar.browser.test.mjs)，其中没有列表弹层检查。主线复用 [document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) 的 `an open numbering popup disables every preset and numbering action when edit permission changes`：它在真实 Workbench 中打开编号弹层，已有六张卡片和 72×80px 边界检查。追加 ORDER_LIST_2 的完整多级标记/可访问名称、实际裁切与非收缩、正文预览线边界及 hover 前后矩形不变断言；没有增加用例声明，所有浏览器用例均未执行。

本轮只改上述预览 CSS 和本节证据。控件仍使用 `antdv-next Button`，沿用 ConfigProvider 主题及原有焦点/命令/历史/保存链路；没有接口或 Worker/WASM 协议变化。标记列的真实布局、主题、hover、焦点、IME 及编辑重开未验收。完整符号与表情选择器方案仍待确认，本轮没有实施。

## 页眉页脚面板与插表对话框续核

2026-11-27 静态复查实际组件 [DocumentHeaderFooterPanel.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterPanel.vue) 和 [DocumentTableInsertDialog.vue](../../packages/office-studio/src/components/document/DocumentTableInsertDialog.vue)。本地没有独立的 `DocumentHeaderFooterControls.vue` 或 `DocumentTableSelector.vue`，插表入口使用既有菜单与尺寸对话框，不新增替代控件。

固定版 [DocHeaderFooterPanel.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/header-footer/panel/DocHeaderFooterPanel.tsx) 按编辑区域展示选项，[DocHeaderFooterOptions.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/header-footer/panel/DocHeaderFooterOptions.tsx) 提供节关联、首页/奇偶页及 0–200、精度一位的距页边输入；上游以侧栏呈现，本地另有 560px Modal、六类 story 内容和字段操作，不声称两者布局相同。

安装的 `antdv-next 1.8.2` 在 `dist/form/FormItem/ItemHolder.js` 第 55 行将 label/control 放入内部 `.ant-form-item-row`。本地旧样式把 100px + 1fr 网格放在 Form.Item 外壳，其唯一 row 被分配到第一列，不能形成预期标签/输入两列。[组件文档](../llm.md) 的 Form layout、labelAlign 和 colon API 与安装源码一致。本轮仅把 Form 改为 horizontal/start/无 colon，并把两列网格及 ≤560px 单列规则移到实际 row；保持原列宽、间距、控件、只读/关联条件及事件。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 中 `headerFooterPanelOpen` 仅声明为 false、读取并绑定到面板，当前没有设为 true 的路径；顶部 `doc.operation.open-header-footer` 打开的是既有 story 编辑区。本次面板布局修正不代表顶部面板入口完成，未新增入口、命令或页面设置事务。真实面板打开路径尚未接入，对应浏览器布局验证尚未编写。

插表依据为固定版 [TableCreate.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/table/create/TableCreate.tsx)：默认 3×5、行列数字输入最小 1、精度整数，上游最大 20。本地对话框按已有模型允许更多行列，并提供宽度、行高、表头、标题及 TSV 初值。本轮确认 Form 的 disabled、确定按钮和 insert 边界守卫均使用现有状态，未修改其尺寸或业务逻辑。[document-toolbar.browser.test.mjs](../../packages/office-studio/tests/document-toolbar.browser.test.mjs) 的既有 `the table Ribbon operation opens the size dialog and retains an undoable table after snapshot reopening` 仍用于该闭环；本轮未追加插表用例，全部运行验证后置。

<a id="rich-story-navigation"></a>

## 富内容页眉页脚的导航续核

2026-11-27 补充 DOC-03 / COMMON-03 静态小修。固定版 [DocHeaderFooterRibbonControl.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/header-footer/panel/DocHeaderFooterRibbonControl.tsx) 的 navigation 组分别以 HEADER / FOOTER 参数执行 [OpenHeaderFooterPanelCommand](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/doc-header-footer.command.ts)，后者使用当前页、目标区域及节引用恢复实际编辑 segment。上下文导航仍是既有两个按钮，常规、插件扩展、上下文及溢出根增量均为 0。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 `handleStoryAction` 已校验活动 story 与 requestScope；当 `editingHost.kind === 'story'` 时转交 `handleBlockStoryAction`。旧 block 处理器只覆盖字段、选项和位置，没有处理 `navigate`，因此富内容页眉的“转到页脚”以及富内容页脚的“转到页眉”会落入无操作分支。旧 flat story 处理器已有 `openStory` 调用，不能把该分支正常行为当作富内容导航完成。

本批仅在 block 处理器增加 `navigate` 早返回，复用同一个 `openStory({ page, area })`。模式、补页、正文选区保存、目标 story block 宿主、输入会话及 Canvas reveal/focus 继续由该入口处理；内容编辑仍沿用既有 story/section 能力判断，导航按现有本地读取操作语义执行。该读取选择与上游按 Edit 禁用导航的差异继续保留，不因本修正消除。没有新增菜单、持久字段、内容 mutation、历史来源或 Worker/WASM 协议。

[document-story-input.browser.test.mjs](../../packages/office-studio/tests/document-story-input.browser.test.mjs) 复用真实 Studio / FDocument / Workbench / ConfigProvider 夹具，通过既有 `projectDocumentStoryBlocks` 构造两个富内容 story。新增 classic 明色与 grid 暗色两个计划用例，检查页眉→页脚→页眉实际 input host、当前区域按钮禁用、输入焦点、模型与内容历史不变、关闭后恢复正文 2–6 范围，以及内部保存重开后继续导航。新增草稿未运行；这些断言不是已通过证据。

所有可替换控件继续使用既有 `antdv-next Button`，主题使用既有 ConfigProvider/token；未改样式或输入桥。真实导航、焦点、主题、Windows IME、文件互操作及性能仍未验收。页眉页脚 Modal 的生产打开路径和表格分页符较大方案继续保留缺口。

本批静态检查通过：3 段 TS/JS（含夹具脚本）的语法、1 个 Vue SFC 解析、125 处本地引用、本文 42 条本地文件链接，以及 3 个修改文件的尾空白/冲突标记与限定 diff 检查。仅加载解析工具，没有执行产品模块或上述浏览器草稿；测试、类型、构建、Rust/WASM、dev、浏览器和基准继续后置。

字符选择器继续使用库中 Select/Form/Input/Button 和 42px 字符格，按钮的禁用由实际 Button props、表单输入的禁用由 Form 上下文传递；未确认新的箭头/hover/尺寸冲突。本轮不实施完整目录、搜索、family/肤色、持久最近项或 keepOpen 方案。全部新增/涉及控件仍使用 `antdv-next`，没有输入桥、接口、Worker/WASM 或性能改动；真实尺寸、主题、焦点、IME、打开/禁用交互与编辑重开均未验收。

<a id="doc-character-picker-plan"></a>

## 完整符号与表情选择器的待确认方案

本方案仅供较大改动确认，尚未实施，也未创建下列计划新增文件。命令保持 `{ value: string }`，已有空格、NBSP、组合字符、宿主、历史及保存链路继续复用。目录和选择器属于独立批次，不能由本节方案或生成数量认定功能完成。

| 计划文件（相对项目根） | 预计手写改动 | 具体工作 |
| --- | --- | --- |
| `scripts/generate-document-emojis.mjs`（新增） | 约 220 行 | 固定 Unicode Emoji 17.0.0 / CLDR JSON 48.2.0，复核固定版分类与 family 分组，生成下面两个数据文件及版本/来源/计数信息；不生成到 Univer 工作树 |
| `packages/office-studio/src/features/document/model/document-emojis.generated.ts`（新增，生成） | 0 行；生成行数未知 | 七类完整 base family、完整 Unicode 序列、默认 recent、统一和 mixed skin tone 变体映射；Emoji/family/变体的数量尚未知 |
| `packages/office-studio/src/features/document/model/document-emoji-locales.generated.ts`（新增，生成） | 0 行；生成行数未知 | 项目 19 个 locale 的 CLDR 标题/搜索词及固定 UI 类别文案；记录 locale 映射和缺失项，条目数/字节尚未知 |
| `packages/office-studio/src/features/document/services/document-special-characters.ts`（新增） | 约 190 行 | 保存上游三组 150 个符号数组项及顺序，保留重复 `%` 与多字符值；将本地箭头/标点/货币等额外值列为扩展。实现目录查询、搜索、family/肤色选择、最近项归族去重与虚拟行计算，供 UI 与合同测试共用 |
| `packages/office-studio/src/features/document/stores/document-character-preferences-store.ts`（新增） | 约 80 行 | Pinia 管理 recent/肤色/活动符号；外部存储读入一次校验，最近项按 family 去重、最多 11，使用本项目命名的存储键。偏好不进入文档快照或内容历史 |
| `packages/office-studio/src/components/document/DocumentCharacterPicker.vue` | 约 360 行 | 现有组件改用完整目录；Ant Form/Input 搜索、类别导航、肤色及 family 变体弹层、随机选择、active/aria-pressed、32px 单元与虚拟滚动；保留自定义整值输入与总 disabled。布局按 10 列、32px 行、24px 分组标题、overscan 3 和可用宽度实现 |
| `packages/office-studio/src/components/document/document-ribbon.ts` | 约 12 行 | 两个现有根透传同一偏好状态与呈现选项；固定命令、分组和入口计数保留 |
| `packages/office-studio/src/platform/ribbon/RibbonItem.vue` | 约 12 行 | insert 事件另传 UI 的 keepOpen 选项；默认沿用菜单关闭，明确要求 keepOpen 的调用保留弹层。选项不混入 `{ value }` 命令参数，关闭/焦点仍共用现有 scope |
| `packages/office-studio/src/components/DocumentWorkbench.vue` | 约 24 行 | API 与独立 modelValue 两种既有用法提供同一偏好 store；重复插入时沿当前选区/宿主推进光标，协调 keepOpen 的选择器焦点与 Canvas 输入焦点 |
| `packages/office-studio/src/api/office-studio.ts` | 约 8 行 | 在现有实例 Pinia 中建立偏好 store，不改变特殊字符事务 |
| `packages/office-studio/src/api/facade.ts` | 约 6 行 | 为 Workbench 提供实例的偏好状态/操作；内容 API 和命令事件仍只使用原参数 |
| `packages/office-studio/src/i18n/docs/locales/document-characters.ts`（新增） | 约 110 行 | 补齐 19 个 locale 的本地扩展/自定义输入文案，复用生成的目录标题与搜索词 |
| `packages/office-studio/src/i18n/office-i18n.ts` | 约 4 行 | 合并字符文案，与既有 locale/direction 注入同用 |
| `package.json` | 约 2 行 | 增加 pnpm 的显式目录生成入口，不在普通构建时临时下载数据 |

固定七类为 people / nature / foods / activity / places / objects / symbols，recent 是额外呈现组，不能算第八个目录类。当前本地六组 Emoji 与六组符号不是这套目录；保留本地有用值时标成扩展，生成后再记录扩展与上游重叠/新增数量。本方案没有生成 Emoji 数据，不承诺总数量，也不把肤色序列简单拼接到任意 Emoji。

搜索同时使用实际字符、固定标题和 CLDR 词，按当前 locale 更新；肤色选择按生成 family 映射选择真实序列，不支持该肤色的 family 保持 base。family 弹层列出全部实际变体，包括 mixed skin tone。recent 保存实际选择值并按 family 去重，换文档/重开工作台可读取偏好，但撤销文档内容不撤销偏好。符号 active 按完整值呈现；`keepOpen` 是 popup 回调的呈现选项，固定来源的 embedded onChange 与 popup onSelect 分开，本地默认菜单关闭行为保持，显式 keepOpen 路径连续插入且保持同一宿主。

| 计划草稿/证据文件 | 预计改动 | 验证目标 |
| --- | --- | --- |
| `packages/office-studio/tests/document-character-catalog.contract.test.mjs`（新增） | 约 200 行 | 直接测真实目录/服务/Pinia：150 符号顺序与重复/多字符、七类、版本、完整 family/肤色映射、中文/英文/RTL 搜索、recent 去重/11 上限、序列化读回及虚拟行边界 |
| `packages/office-studio/tests/document-toolbar-state.contract.test.mjs` | 约 30 行 | 两根、fixed ID、偏好透传和 UI keepOpen 不进入领域参数 |
| `packages/office-studio/tests/document-structured-workbench.browser.test.mjs` | 约 180 行 | 复用真实 Workbench/Studio：完整目录搜索、统一/混合肤色、最近项持久化、随机项属于目录、active/keepOpen 连续插入、虚拟滚动尾部、classic/grid/More、主题/RTL/焦点、锁与一次历史、Code/story/table 保存重开；同步原 42px 草稿为实际 32px 目录规格 |
| `docs/feature-evidence/DOC-03-header-footer-ribbon-sizing.md` | 约 45 行 | 记录生成来源、分类计数、实现差异、实测和未核实项 |
| `docs/UNIVER_FEATURE_PARITY.md` | 约 3 行 | 更新完整 Picker 实际完成/未验收口径 |
| `docs/UNIVER_TOOLBAR_PARITY.md` | 约 3 行 | 链接证据，保持根入口统计 |
| `docs/feature-evidence/TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md` | 约 5 行 | 更新菜单与上下文可观察行为，不把目录项数当功能完成证明 |

目录加载、搜索和虚拟行计算先作为同步可测试模块；测试、浏览器和基准按用户约定统一后置。在统一验证阶段，用同一完整数据测加载/搜索/滚动、主线程长任务与内存，再按实测选择计算路径；本方案没有预报达标或优化。Office 样本独立重开、真实 IME、字体显示及其他辅助宿主继续按完整验收标准检查。

## 检查与剩余验证

此前尺寸批次已记录两个修改组件的限定 `git diff --check`、两个新增文件尾部空白/冲突标记及本文本地链接检查通过，固定上游 HEAD 与记录一致。导航按下事件和焦点草稿的小修在子任务收尾时未重新执行这些检查；随后主线已完成本轮统一语法、引用、文档链接/锚点及空白检查，范围和结果见[顶部控件静态收尾](../UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing)。

前批导航焦点小修只修改对应控件、现有浏览器草稿和本文；未运行测试、类型检查、前端构建、Rust/WASM 检查、dev、浏览器或性能基准。

待统一验证：实际布局与鼠标命中、键盘焦点和关闭弹层后的焦点恢复、明暗主题切换、IME、两个渲染后端中的工作台交互，以及实际文档的节/页眉页脚编辑、撤销重做、保存与 Office 文件独立重开。未新增性能优化结论，未把现有 DOC-03 或 Pro 能力标为已验收。
