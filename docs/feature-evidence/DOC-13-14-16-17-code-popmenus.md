# DOC-13 / DOC-14 / DOC-16 / DOC-17 Code 字体、语法与段落弹层实施记录

记录日期：2026-11-25。状态：**未验收，前端开发中**。这是 [初批结构块](DOC-13-14-16-17-structured-blocks.md) 和 [第二批 Canvas 交互](DOC-13-14-16-17-canvas-interactions.md) 的第三批补充。按当前用户安排先实现和编写草稿，后端与统一运行验证后置；本批没有运行测试、类型检查、构建、浏览器/dev、Cargo/WASM 或性能基准。

本批落盘固定 21 项语言的真实词法、Code 字体命令与 Ant 属性弹层、六种 paragraph 挂载及表格下方入口。入口和草稿数量只描述清点口径；不代表 Docs、全部 Pro、Office 文件或性能验收已完成。

## 1. 固定依据与实现差异

OSS 固定为 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 来源为本项目 `.research.local/univer-1.6.0/packages` 的公开发布包类型与可读 ES，未将这些商业包加入产品依赖。动态公开规格继续见初批记录，本批没有新增可访问演示或 Office 文件验证。

| 可复查来源 | 观察到的契约 | 本地实现 / 差异 |
| --- | --- | --- |
| [set-font-family.command.d.ts](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/types/commands/set-font-family.command.d.ts)；[Code ES](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/es/index.js) 中 `Vt` / `docs-code.command.set-font-family` | 只接受 FontService 的 MONOSPACE；按 unitId/blockId 校验 Code 和权限；给所属全部段落的 `paragraphStyle.textStyle.ff` 和 textRuns 写真实字体，通过内容 mutation 提交 | `doc.command.set-code-font-family` 设置每个段落、Latin/EA run 字体和段落默认，沿用候选、锁、权限及内容历史。增加 EA 字体为本地适配；不是单独的高亮外观配置 |
| OSS [font.service.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/services/font.service.ts)，`FontCategory.MONOSPACE`（60、111、119、120、138、202、205、216、220 行） | 固定九种默认等宽字体；宿主可登记分类字体 | 本地九种目录加已有 Windows 默认 Consolas/Cascadia Code/monospace，并提供显式等宽注册 API。登记不加载字体文件，真实字体可用性和度量仍需验收 |
| [DocsCodeFontFamilyPicker](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/types/views/font-family-picker/DocsCodeFontFamilyPicker.d.ts)、[桌面 render controller](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/types/controllers/docs-code-render.controller.d.ts)、[mobile controller](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/types/controllers/mobile/docs-code-render.controller.d.ts)、[MobileDocsCodeLanguagePicker](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/types/views/language-picker/MobileDocsCodeLanguagePicker.d.ts) | 真实块命中、字体/语言选择回调，桌面 popup 与移动弹层分别挂载，受内容权限约束 | 桌面使用 Ant Popover/Dropdown，宽度不大于 640 的宿主使用底部 Ant Drawer。More 提供语言、换行、行号、缩进和真实内容动作；具体弹层布局与上游有差异，全部触控/运行态溢出仍未核实 |
| Code ES 的固定语言目录、`Wt` / `Kt` / `Jt` 与 [languages.d.ts](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/types/common/languages.d.ts) | 21 项；上游实际调用 Prism.tokenize，展平嵌套 Token 为源码范围，按类型取色；未知 grammar 回退 JavaScript | 本地固定 Prism 1.30.0，装载所有非纯文本 grammar，嵌套展平为 UTF-16 分区。保留既有七种 Canvas 色类，未知 grammar 在视图中按纯文本显示，作者 language 原值保留；这两项为明确差异 |
| OSS [ParagraphMenu.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/ParagraphMenu.tsx)（412–435、897–909 行）；[paragraph-menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/paragraph-menu.ts)；[初批六挂载及 Column schema](DOC-13-14-16-17-structured-blocks.md#22-quote--callout--code-的-paragraph-context-挂载) | 空段快速插入、普通/结构块快速编辑、内容/下方插入、表格菜单区分目标；Callout 隐藏 Code/Quote 转换，Code/Quote 隐藏 Callout；Column 编辑有四个独立操作 | 本地按相同分支生成当前/下方菜单，保留上述转换显隐和组/栏身份；普通右键另保留已有“将选段设为”的选区语义。表格下方 Column 有直接上游依据；本地 Quote/Callout/Code 表后入口是便利扩展，未认定为上游相同注册 |

Context7 曾查询 Ant 与 Prism API；Prism 查询返回的新版文档不作为本批版本依据。实际接口以本地 `prismjs@1.30.0` 源码/components.json、`@types/prismjs@1.26.5` 和 `antdv-next` 声明为准。

## 2. 入口分类、顺序和状态

常规 Ribbon 根仍为 **52（home 33 / insert 16 / view 3）**，人工静态计数包含 hidden、排除 children/context。本批没有新增常规根。结构上下文 Ribbon 沿用四根；当前 Code 结构菜单由第二批七根变为九根（语言、wrap、行号、字体、More，加属性/unwrap/exit/remove），Quote/Callout 各四根，Column 七根。

以下只清点本批四类结构功能在挂载中的贡献，不是完整 Docs 菜单总数。classic/grid 常规位置未变；这些入口属于 paragraph、上下文或局部弹层，不能计入常驻根。

| 本地挂载 | 实际宿主 | 分组顺序 | 声明计数与状态 |
| --- | --- | --- | --- |
| `layout-below` | 文档右键“在下方插入” | 分栏、Callout、Quote、Code | 4 根 / 7 叶；显式 afterBlockId，不能在已有栏内嵌套新组 |
| `content-insert` | 当前内容末尾 `＋`，包含辅助 host | 分栏、Callout、Quote、Code | 4 根 / 7 叶；定位最终结构分片或段落末行 |
| `empty-paragraph` | 空段右键“插入内容” | 分栏、Callout、Quote、Code | 4 根 / 7 叶；只对空文本段落显示 |
| `quick-insert` | 空段 `＋` / 移动“段落” | 分栏、Code、Quote、Callout | 4 根 / 7 叶；包裹明确的当前段落，移动端同一目录 |
| `quick-edit` | 普通/结构段落 `T` / 移动“段落” | Code、Quote、Callout；栏内追加添加栏、删栏、删组、等宽 | 3 个结构根，最多追加 4 个栏操作；按当前结构隐藏转换项，当前类型 activated；Code 只允许完整文本段落 |
| `quick-insert-below` | `T` 的“在段落下方插入” | 分栏、Code、Quote、Callout | 4 根 / 7 叶；下方容器能力独立于当前光标是否在栏内，组后允许另一个组 |
| `table-below` | 最终可见 table fragment 的 `＋`；表格选区右键 | 分栏、Callout、Quote、Code | 4 根 / 7 叶；Column 为上游贡献，其余为本地扩展；按 clip 后矩形定位，不显示空 clip 或 continuedAfter 分片 |

每个分栏 submenu 有 2/3/4/5 四个选择。quick-edit 的栏操作绑定真实 `columnGroupId`/`columnId`，不会把内层 Code ID 传给栏命令；5 栏禁添加、2 栏禁删栏，保留既有菜单禁用差异。所有动作仍由执行时权限/内容锁确认，生成目录不等于授予能力。

Code 桌面浮动区为字体、语言、More 三个 Ant 触发控件，移动端相同三个触发控件进入 Drawer；Quote/Callout/Column 保持属性和 More 两控件。语言目录固定 21 项，旧 plain 或作者未知值额外显示；字体目录固定 12 项，未登记的当前作者字体额外保留为禁用选项，避免无声替换。物理分片、触发控件和菜单子项分别计数，不重复当作产品功能。

## 3. 命令、事务与公开接口

新增两个实际编辑命令，已同时纳入 `DocumentStructureCommand`、ID 分类和 [document-editor-commands.ts](../../packages/office-studio/src/api/document-editor-commands.ts) 标签/注册路径：

| 操作 | 共用命令 | 实际数据与历史 |
| --- | --- | --- |
| Code 字体 | `doc.command.set-code-font-family` | 整块段落/run 字体与默认；检查等宽登记、文本锁及实体/段落权限，一次内容历史 |
| Code More | `doc.command.update-code-settings` | language/wrap/showLineNumbers/tabSize 四字段同一 patch；菜单解析不再丢掉第二及后续字段 |
| 完整 Code 属性 | `doc.command.update-code-block`，新增可选 fontFamily | 文本、设置、字体同一候选；现代容器与旧 paragraph metadata 路线都转发字体；失败不发布部分文本/运行 |
| 结构转换 | `doc.command.convert-structured-block` | 保留容器/子段落 ID、内容和资源；Code/Callout 配置保存在 `metadata.documentStructuredTypeSources`，再次转换可恢复；离开 Quote 清理可确认的生成边框 |
| 当前/下方插入与栏操作 | wrap/insert structured、wrap/insert column、add/remove column、remove structured、update column | 显式 blockIds/afterBlockId/组/栏身份进入已有命令，不直接在 Vue 中修改模型 |

`FDocumentCode` 新增 `getFontFamily()` / `setFontFamily()`，describe 返回 fontFamily；`DocumentCodeInsertOptions` 和 `DocumentCodeBlockData` 增加可选 fontFamily。insert/wrap Code 同样接受字体，非 Code 拒绝该字段。[document-code-fonts.ts](../../packages/office-studio/src/features/document/code-blocks/document-code-fonts.ts) 提供公开 `registerDocumentCodeFontFamilies()`（返回可重复调用的解除函数）和 `DocumentCodeFontFamilyOption`；[index.ts](../../packages/office-studio/src/index.ts) 同时导出固定语言目录及其类型。

字体权威数据仍在段落/run，不把 fontFamily 塞入仅四字段的 Code config。明确设置当前首 run 字体时也会统一其它混合 runs；只改语言/More 时不抹平混合字体。已撤销登记的作者字体继续可读和保存，但重新设置要先登记或选择可用字体。

[document-body-paragraph.ts](../../packages/office-studio/src/api/document-body-paragraph.ts) 写真实 `paragraphStyle.textStyle.ff` 并合并其它作者来源字段；既有 textRuns 编解码写 ff/EA 字体。快照和 body 往返有草稿，Office 二进制与独立重开仍未验收。新增 metadata 字段及字体 DTO 的后端/协作传输需求由主任务维护。

## 4. 真实语法与源码保真

新增 [document-code-tokenizer.ts](../../packages/office-studio/src/features/document/code-blocks/document-code-tokenizer.ts)，替换旧九语言共同 regex。Prism 默认 markup/css/clike/javascript 加其余 grammar，按 TS→JS、Markdown→markup、C++→C、PHP→markup-templating 的实际依赖顺序装载，目录中 plaintext 返回纯文本。

递归访问 Token 的字符串/数组/嵌套内容，并处理 alias/插值。输出为完整、不重叠的 UTF-16 start/end/text 分区，保留 LF 规范化、emoji、组合字、tab 与转义；既有 Canvas 从同一源码范围取七类颜色，不生成或保存高亮 HTML。别名仅用于 grammar 选择，查询只接受目录自身的键，`constructor` / `toString` / `__proto__` 等作者值不会误用对象原型。`normalizeCodeLanguage()` 现在保留非空作者字符串及大小写/空白，旧 metadata 更新合并未知字段。

[document-code-prism-config.ts](../../packages/office-studio/src/features/document/code-blocks/document-code-prism-config.ts) 在 lexer 导入前设置 manual/disableWorkerMessageHandler，模块也列入包的 sideEffects，防止构建器丢掉初始化。只使用 tokenize，不接管 DOM 高亮或本项目 Worker 消息。本批通过 pnpm.cmd 安装固定 prismjs/types，使用 `--ignore-scripts`；依赖安装不是测试/构建结果。

语法仍是当前 TypeScript/主线程前端路径，没有新 Worker/WASM token 任务，也没有大样本词法性能、内存或 bundle 体积实测。

## 5. Ant 控件、目标与焦点

新增 [DocumentCodeControls.vue](../../packages/office-studio/src/components/document/DocumentCodeControls.vue) 和 [DocumentStructuredPopMenus.vue](../../packages/office-studio/src/components/document/DocumentStructuredPopMenus.vue)；修改 [DocumentCodeBlockPanel.vue](../../packages/office-studio/src/components/document/DocumentCodeBlockPanel.vue)、[DocumentStructuredOverlay.vue](../../packages/office-studio/src/components/document/DocumentStructuredOverlay.vue)、[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 和 [document-structure-ribbon.ts](../../packages/office-studio/src/components/document/document-structure-ribbon.ts)。可替换控件全部来自 antdv-next Form/FormItem、Select、Input/Input.TextArea、InputNumber、Switch、Button、Popover、Dropdown、Menu、Drawer、Modal、Tooltip、Divider；输入桥沿用 Ant Input.TextArea。

新弹层继承现有统一 ConfigProvider，局部布局使用 Overlay 的主题 token；Canvas 绘制/命中、调宽 separator/guide 保留编辑器内核职责。Code 全属性面板可按打开意图聚焦字体、语言或缩进字段，已静态确认 Modal 的 afterOpenChange 和 Drawer 的 close/focusable 接口。

2026-11-27 补充桌面 Code 浮动触发器边界小修，状态仍为**未验收**。固定 [Code ES](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/es/index.js) 中 `_shouldDrawLanguageButton` 分支的字体、语言均调用 `pn`：宽度 70–118px，文本绘制也限宽。本地沿用 Ant small Button 的尺寸，仅将这两个桌面触发器上限设为 118px，并允许按工具栏剩余空间收缩；通过 Button `classes.content` 约束实际文本 span 的省略，Ant Tooltip 在悬浮或聚焦时显示完整值。字体、未知 language 作者值与命令载荷保持完整，移动 Drawer 入口沿用原布局。

[DocumentStructuredOverlay.vue](../../packages/office-studio/src/components/document/DocumentStructuredOverlay.vue) 复用原 left 定位结果，maxWidth 改为 viewport.width − left − 8px，修正原来预留 290px、却允许整个 viewport 宽度的矛盾；top、目标捕获和焦点恢复逻辑未变。既有 [document-structured-workbench.browser.test.mjs](../../packages/office-studio/tests/document-structured-workbench.browser.test.mjs) 新增 1 个未执行草稿：真实登记的长字体、作者未知语言与右侧代码块样本，覆盖 classic/light 与 grid/dark 的工具栏/按钮/文本边界、hover 尺寸、完整 Tooltip、真实 Popover 取消与输入焦点，以及保存重开保留原值。浏览器夹具补导入与产品相同的 Ant reset；新增草稿不计入第三批原有 56 个用例，也不代表字体实际加载或运行态验收。

打开时捕获 source、选区、host、revision/generation；提交、取消恢复焦点和内容动作再核对。换模型/选区/host、禁用、失活、缩放或 viewport 尺寸改变会取消未提交弹层。移动 Drawer 禁用自动返回旧触发按钮，由当前有效目标恢复 Canvas 输入焦点；旧目标不恢复焦点。实际焦点顺序、Popover/Select 事件顺序、触控、明暗主题和真实 Windows IME 仍需浏览器验收。

## 6. 草稿、静态核对与剩余缺口

新增 **2 份契约草稿**，扩展 **4 份既有草稿**，本批新增 **56 个用例**；另调整旧 Code 返回值/语言保真预期和旧语言弹层浏览器选择器。全部未执行。

| 文件 | 本批新增草稿 |
| --- | --- |
| [document-code-tokenizer.contract.test.mjs](../../packages/office-studio/tests/document-code-tokenizer.contract.test.mjs) | 27：21 种真实语言样例；嵌套插值、嵌入脚本/CSS、多行/转义、UTF-16/emoji/空输入、alias/未知作者值与来源保真 |
| [document-structured-popmenus.contract.test.mjs](../../packages/office-studio/tests/document-structured-popmenus.contract.test.mjs) | 13：7 挂载实际执行，分组/显隐、原容器转换、嵌套 Code 的栏目标、下方容器能力、完整 More patch 和权限/不存在目标原子性 |
| [document-structured-facade.contract.test.mjs](../../packages/office-studio/tests/document-structured-facade.contract.test.mjs) | 8：实际 OfficeStudio 注册/事件/一次历史、font/More/source 联合事务、内容锁/权限/取消/无效参数、自定义登记解除、legacy 混合 runs、转换/undo/redo/body/保存重开 |
| [document-structured-canvas.contract.test.mjs](../../packages/office-studio/tests/document-structured-canvas.contract.test.mjs) | 1：字体改变真实 glyph 度量、命中和 Canvas scene.font，逻辑 ID/源码不变；使用可复现人工度量，不是浏览器字体或性能验收 |
| [document-structured-workbench.browser.test.mjs](../../packages/office-studio/tests/document-structured-workbench.browser.test.mjs) | 第三批 7：实际 Canvas/Ant 桌面字体/More、移动 Drawer、取消/旧目标、T/空段插入、表格后真实 Quote/Column、辅助 header 内容插入；新增 table/blank 样本及 cellBlockId 几何定位。2026-11-27 另增 1 个桌面触发器边界草稿，见第 5 节 |
| [document-code-blocks.contract.test.mjs](../../packages/office-studio/tests/document-code-blocks.contract.test.mjs) | 更新旧 fontFamily 返回值和未知 language 不再丢失的预期，没有新增用例 |

静态核对仅为 UTF-8 源码阅读、rg、git status/diff 和声明/来源核对；未执行测试、类型、构建或浏览器。修正了命令标签缺项、legacy 字体漏转发、混合字体同首值无操作、table clip 锚点、Drawer 默认焦点竞争及语言别名原型键误识别；原型键断言并入既有语言保真草稿，56 个新增用例的口径不变。草稿数量不是已通过记录。

本次补充前的三批证据中，105 个本地相对链接（初批 45 / 第二批 32 / 第三批 28）与其中 7 个 Markdown 锚点曾按文件存在性及标题/显式锚点静态核对，无缺失。2026-11-27 子任务收尾时没有重新执行整表链接核对；随后主线已完成本轮统一语法、引用、文档链接/锚点和空白检查，范围与结果见[顶部控件静态收尾](../UNIVER_TOOLBAR_PARITY.md#top-ribbon-control-sizing)。远程链接、在线演示和运行态行为没有验证。

剩余工作保持原完整范围：

1. **Quote 分页**：[第四批](DOC-16-quote-pagination.md) 已落盘四项段落属性、Quote keepLines/widowControl/相邻 keepNext、作者恢复、共享布局/命令/历史/公共 body 和 Ant 面板；新增 37 个草稿全部未执行。原生 DOCX / Office 样本、复杂对象/多栏/notes 分页兼容及性能仍未验收，未标完整完成。
2. **入口和模式**：全部 paragraph/PopMenu 的真实显示、触控、More 运行态溢出、与基础标题/列表/表格菜单的完整组合，以及 Column modern/traditional、继承故事/分页/嵌套场景尚未验收；常规 Quote/Callout 分组位置差异沿用初批记录。
3. **语法和字体**：所有语言复杂/超长样例、原生字体/自定义字体加载与混合 Latin/EA/RTL、七色类和上游细分类的差异、真实 IME、缩放/滚动后编辑仍需样本。
4. **内容与 Office**：完整字段/控件/批注/note/表格/资源组合、结构复制粘贴、DOC/DOCX 解析写出、Office 独立重开、权限/协作与历史传输仍未验收。
5. **平台与性能**：规定前端/Rust/WASM/浏览器检查，WebGPU/Canvas2D、首屏/滚动/编辑/缩放、Worker/WASM 复制/取消与内存、大样本前后基准均未执行。本批不声称性能优化或整条链路零拷贝。

中央对齐表与 BACKEND_REQUIREMENTS 由主任务维护。本证据只记录前端真实落盘、接口和未验收项，不把临时后置工作改为交付范围缩减。
