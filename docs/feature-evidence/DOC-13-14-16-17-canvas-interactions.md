# DOC-13 / DOC-14 / DOC-16 / DOC-17 Canvas 交互与辅助内容区实施记录

记录日期：2026-11-25。状态：**未验收，前端开发中**。这是 [结构块初批记录](DOC-13-14-16-17-structured-blocks.md) 的第二批补充。按用户安排，前端和测试草稿先推进，后端及运行验证后置；本批没有运行测试、类型检查、构建、浏览器/dev、Cargo/WASM 或性能基准。

本批接通结构块浮动菜单、段落插入菜单、多栏调宽、尺寸配置、代码块水平滚动，以及页眉/页脚顶层结构块和表格内结构块的布局路径。代码接线与草稿不是产品验收证据；全部 Pro 能力、Office 互操作、主题/IME 和真实性能仍在交付范围内。

本文为第二批实施快照。后续的 21 语言真实词法、Code 字体/More、六种段落及表格下方/mobile 入口见 [第三批记录](DOC-13-14-16-17-code-popmenus.md)，全部仍未验收。

## 1. 固定依据与可复查差异

OSS 固定参考为 `E:\rust_test\univer` 的 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 依据为本项目 `.research.local/univer-1.6.0/packages` 发布包的类型、公开命令及可读 ES 分发代码，没有将这些包加入生产依赖。

| 依据 | 已核对的可观察契约 | 本批实现 / 差异 |
| --- | --- | --- |
| OSS [column.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/block/column.ts)，默认值与 `createColumnGroupSkeleton`（49–83 行）、`calculateColumnGroupLayout`（225 行） | 默认水平内边距 8、最小高度 72；子栏有真实左右 margin，组高度至少达到 minHeight；宽度与 SHRINK/STACK 分开计算 | 默认和显式尺寸接入同一布局度量；超大正文 minHeight 占用后续物理 frame。采用本项目树与分页引擎，运行行为未验收 |
| Pro [resize-handle.d.ts](../../.research.local/univer-1.6.0/packages/docs-column-ui/lib/types/services/resize-handle.d.ts)、[resize-interaction.d.ts](../../.research.local/univer-1.6.0/packages/docs-column-ui/lib/types/services/resize-interaction.d.ts)、[resize render controller](../../.research.local/univer-1.6.0/packages/docs-column-ui/lib/types/controllers/render-controllers/column-resize.render-controller.d.ts) | 以组 ID、分界索引、left/top/height 命中真实分界；拖动 delta 转成 resize 命令；上游最小可见高度常量为 72 | 本地分界来自布局容器；移动只显示 guide，释放提交 widthRatios。保留相邻比例总和与其他栏权重，使用同一宽度算法反求比例；最低拖动宽度为 24 px 并尊重 minWidth。显式 minHeight 小于 72 时，手柄按实际容器高度绘制，与上游最小可见高度规则有差异 |
| Pro [FDocumentColumnGroup](../../.research.local/univer-1.6.0/packages/docs-column/lib/types/facade/f-document-column-group.d.ts) | 组有 2–5 栏；宽度修改走同一内容 mutation/undo/collaboration；声明仅支持 modern，traditional 的读写有显式限制 | 本项目接同一内容命令和历史。padding/minHeight 的 pt getter/setter 是本地 Facade 扩展。当前允许已有组进入本地传统分页与辅助 host；这属于实现差异，模式限制和 Office 语义还需验证 |
| Pro [Code render controller](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/types/controllers/docs-code-render.controller.d.ts)、[Code menu](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/types/views/menu.d.ts) 与初批 schema 清点 | 有块/控件命中、语言和字体弹层，受编辑权限约束；语言命令为 `docs-code.command.update`；固定目录 21 项 | 第二批浮动菜单修改语言、wrap、行号；第三批已接真实字体命令和桌面/mobile 弹层，见后续记录。本地长行滚动使用视图投影，上游滚动 DTO 和运行行为仍未核实 |
| 初批 [六种 paragraph/context 挂载](DOC-13-14-16-17-structured-blocks.md#22-quote--callout--code-的-paragraph-context-挂载) 与 Column `F` schema | 当前段落、下方插入、空段、快速插入/编辑以及表格下方入口分别挂载，不能把重复 schema 计成常规根 | 第二批提供真实段落 `＋` 与共用命令解析；第三批已补六种挂载、表格下方和移动 Drawer，实际行为仍未验收 |

公开功能规格继续见初批记录的 Code、Column、Quote、Callout 文档链接。本批没有新增在线演示或 Microsoft Office 文件样本验证。

## 2. 本地入口计数与命令闭环

本批没有增加常规 Ribbon 根。当前 [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 常规根人工静态计数为 **52（home 33 / insert 16 / view 3）**，包含 hidden，不包含 children 或 context；该数字不是功能完成数量。

下表保留第二批入口快照；第三批 Code 结构菜单改为九根，字体/语言/More 为三个触发控件，段落挂载分类和顺序以第三批记录为准。

| 类别 | 当前本地入口 | 计数口径 |
| --- | --- | --- |
| 结构上下文 Ribbon | 属性、保留内容并移除容器、在块后继续输入、删除块和内容 | 4 根，沿用初批；按当前结构块显示 |
| 浮动结构区 | 属性按钮、更多菜单按钮 | 2 个控件；每个可见物理实例不计成新功能 |
| Quote / Callout 结构菜单 | 上述四动作 | 各 4 个根 |
| Code 结构菜单 | 语言 submenu、wrap、行号，再加上述四动作 | 7 个根；语言选择另计 |
| Column 结构菜单 | 当前栏右侧添加、删除当前栏和内容、等宽，再加上述四动作 | 7 个根；5 栏禁用添加，2 栏禁用删栏 |
| 固定 Code 语言目录 | plaintext、javascript、typescript、json、markup、css、markdown、yaml、sql、bash、python、java、c、cpp、csharp、go、rust、ruby、php、swift、kotlin | 21 项；当前来源的未知语言或旧 `plain` alias 会额外保留一项 |
| 段落 `＋` | “将选段设为”“在段落下方插入” | 2 个 submenu 根；每个含 Quote/Callout/Code 3 个动作和 2–5 栏 4 个选择，共 14 个叶动作。分栏分组不当成常规根 |
| 调宽 / Code 滚动 | 每个实际栏边界的 separator、每个可见 nowrap Code 分片的 Slider | 动态几何实例，不作为静态菜单根或功能数量 |

[document-structure-ribbon.ts](../../packages/office-studio/src/components/document/document-structure-ribbon.ts) 集中生成上下文和段落目录，`documentStructureCommandFromMenu()` 被 Ribbon、右键及浮动菜单共用。

| 用户操作 | 统一本地命令 | 数据效果 / 历史 |
| --- | --- | --- |
| 当前选段设为结构块 / 多栏 | `wrap-structured-block` / `wrap-column-group` | 包裹真实子块，保留来源运行；沿用命令权限、锁、字段和内容历史 |
| 在段落下方插入 | `insert-structured-block` / `insert-column-group`，显式 `afterBlockId` | 创建真实块和输入段落 |
| Code 语言、换行、行号 | `update-code-settings` | 修改实际 Code 配置，源段落保留 |
| 添加、删除当前栏、等宽 | `add-column` / `remove-column` / `update-column-group` | 添加真实段落，删除明确使用 `keepContents:false`，等宽提交全 1 比例 |
| 指针调宽 | `update-column-group` | pointermove 只有 guide；pointerup 一次真实命令。Escape/pointercancel、模型/选区/host/缩放/失活变化撤销未提交手势 |
| 键盘调宽 | `update-column-group` | 左右键 8 px，Shift 加大为 24 px；每次进入同一命令 |
| Code 水平滚动 | 无内容命令，局部视图状态 | 改变绘制/命中/输入投影；不修改模型、不建立内容历史、不写入保存结果 |

表内短名均为 `doc.command.<短名>`。属性 operation 只打开有明确 blockId 的真实面板；菜单和手柄没有空操作、提示消息替代功能或第二套内容历史。

## 3. 尺寸、公共 API 与 body 保真

领域 `DocumentColumnGroup.horizontalPadding?` / `minHeight?` 使用 CSS px；默认常量位于 [document-structural-blocks.ts](../../packages/office-studio/src/features/document/model/document-structural-blocks.ts)，分别为 8 和 72。[document-structural-layout.ts](../../packages/office-studio/src/features/document/layout/document-structural-layout.ts) 继续重导出，既有导入兼容。

[DocumentStructuredBlockPanel.vue](../../packages/office-studio/src/components/document/DocumentStructuredBlockPanel.vue) 使用 `FormItem` / `InputNumber` 提供尺寸字段，提交经 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 进入 `wrap-column-group` 或 `update-column-group`。正文、多栏嵌套和表格相对测量读同一 `documentColumnMetrics()`。

[document-columns.ts](../../packages/office-studio/src/api/document-columns.ts) 新增插入 options 的两个字段，及 `getHorizontalPadding` / `setHorizontalPadding`、`getMinHeight` / `setMinHeight` 和 `describe()` 对应字段。Facade 使用 pt，控制 DTO/领域/Canvas 使用 CSS px，只转换一次（乘 4/3）；有效默认读值为 6 pt / 54 pt。非法插入 gap 仍回退既有默认；非法 gap 更新和两个新尺寸的显式非法值拒绝。

[document-body-structures.ts](../../packages/office-studio/src/api/document-body-structures.ts) 按既有 NumberUnit 归一化协议读取尺寸；[document-body-read.ts](../../packages/office-studio/src/api/document-body-read.ts) 保留来源及解析基线。尺寸未改变时保留作者数值/单位和附加来源字段，改变后写 `{ v: <CSS px>, u: 3 }`；新字段显式提供 null 或无效单位会拒绝，避免静默丢弃。字号/参考长度变化后的全部相对单位组合仍需验证。

这些接口是前端控制与领域数据；后端 DTO、Rust/WASM、协作和二进制读写需求由主任务维护 [BACKEND_REQUIREMENTS.md](../BACKEND_REQUIREMENTS.md)，本批没有后端实现或 Office 独立重开证据。

## 4. Code 的源布局、投影、命中与输入

[document-code-scroll.ts](../../packages/office-studio/src/features/document/layout/document-code-scroll.ts) 从权威布局读取 Code 容器、内容 clip 和所有物理分片的最大横向 extent。key 区分 body 与 story/note/decoration host；同一块的分片共用 offset 和最大滚动量。

[DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 保存 `sourceLayout` 和局部 `codeScrollOffsets`，通过 `projectDocumentCodeScroll()` 产生绘制/命中/输入布局。向父级 emit 的 layout 仍是 sourceLayout；Overlay 的 Slider 最大值来自 `sourceEditingLayout`，避免从已滚动投影再计算导致最大值缩小。

投影复制需要变化的 line/glyph/marker 坐标，使用同一 Map 让主 lines、pageContent、auxiliary、note 和 header/footer 索引共享同一投影 line。源布局和模型不被移动。行号使用固定 `gutterX` / `gutterClip`，[document-scene.ts](../../packages/office-studio/src/features/document/rendering/document-scene.ts) 在文字全滚出时仍能画可见行号；[document-layout.ts](../../packages/office-studio/src/features/document/layout/document-layout.ts) 的可见命中矩形保留短/空 Code 行的可点击区域。

光标可达性由 `documentCodeCaretScrollOffset()` 读取原始 glyph 坐标，先滚入 Code 内容，再进入既有文档滚动路径。`caretRect` 的 clip 钳制继续使用同一投影。nowrap Code 留 24 px 底部空间承载 `antdv-next Slider`，wrap Code 保持 8 px；Slider 配置 `aria-label-for-handle` 与 `changeComplete` 完成事件。

第二批 tokenizer 仅补 plaintext→plain、markup→html alias；第三批已替换为固定 Prism 1.30.0 的 21 项对应词法，并保留真实容器和 legacy paragraph 的作者 language 原值。未知 grammar 在本地视图按纯文本回退，七种 Canvas 色类继续保留；复杂语法、字体和源码组合仍未验收，见第三批记录。

## 5. 辅助 host 的真实结构块布局

静态回读发现旧 `layoutDocumentAuxiliaryBlocks()` 把顶层 Quote/Callout/Code/Columns 落入普通空 runs 分支，顶层表格也未收集 `measured.containers`。本批已替换这条路径。

新增 [document-block-measurement.ts](../../packages/office-studio/src/features/document/layout/document-block-measurement.ts)，从既有表格内容测量提取 `measureDocumentBlockCollections()`：共享段落、结构递归、Code nowrap/tab/行号、Column padding/minHeight、图片输入锚点、嵌套表格和 clip 的相对几何。调用方提供真实 body/cell target 与表格测量回调；领域数据不依赖 DOM，不增加第二套表格权威引擎。

[document-table-measurement.ts](../../packages/office-studio/src/features/document/layout/document-table-measurement.ts) 复用该内核，保留 table cell margin、owner、rowSpan/exact row、continuation、分页切口和 `tableDepth`。原 `intersectTableRect` / `translateTableLine` / `translateTableRect` 以兼容导出继续可用；共享 line 平移还处理 Code gutter 坐标。

[document-auxiliary-layout.ts](../../packages/office-studio/src/features/document/layout/document-auxiliary-layout.ts) 直接测量顶层真实块，收集所有 containers 和表格 fragment；统一转换 page、host、section、line/glyph/marker、clip 和 table row 坐标。footer 根据实际内容高度底部对齐；PAGE/NUMPAGES 显示仍在副本中刷新并映回缓存的源 UTF-16 坐标。`documentLayoutForBlockHost()` 返回该 host 的 lines 和 containers，并同步 pageContent，供浮动操作、调宽、命中和输入共用。

正文超大 Column minHeight 会占用真实后续页面/栏，包括无文字的中间 frame。表格内与辅助故事的溢出、裁剪、重复表头、继承页眉/页脚和独立 inspection 的运行行为仍待验收；本批不宣称这些场景全部已验证。

## 6. 会话、控件与主题

[document-structured-canvas.ts](../../packages/office-studio/src/features/document/controllers/document-structured-canvas.ts) 的目标捕获 root source 身份、克隆选区/host 和输入 session；standalone 保存源快照。提交前核对 modelRevision、generation、有效状态、位置和 host。Overlay、Viewport、Workbench 三端都检查当前目标，再进入既有 `execute()`；旧源/输入会话、选区变化、composition 和失活不沿用旧操作。

Workbench 的代码属性面板保存明确 `codeDialogBlockId`，右键结构动作也保存打开时的目标。嵌套结构属性按 params.blockId 找到真实块，避免提交时临时选中另一个 Code 或父容器。

新增浮动区使用 `antdv-next Button / Dropdown / Tooltip / Slider`，属性面板使用既有 `Modal / Form / FormItem` 和对应输入组件。控件继承统一 ConfigProvider，布局装饰与手柄使用 `theme.useToken()` 的主题变量；没有新增原生表单控件。DOM separator/guide 是 Canvas 编辑器的命中和拖拽几何。正文/辅助输入继续由既有 `antdv-next Input.TextArea` 桥承载。

本批静态核对了主题 token、禁用条件、取消监听和焦点恢复调用。明暗弹层、焦点、真实 Windows IME、触控、窄屏和低缩放 Slider 均未运行验证；合成 composition 草稿不构成真实输入法验收。

## 7. 测试草稿与未验收项

新增 **2 份契约草稿**，扩展 **3 份既有草稿**，本批新增 **26 个用例**，全部未执行。

| 草稿 | 本批新增范围 |
| --- | --- |
| [document-structured-canvas.contract.test.mjs](../../packages/office-studio/tests/document-structured-canvas.contract.test.mjs) | 9 个：尺寸/超大正文与表格分页、源不变/索引共享/命中、光标与 clamp、host/分片 key、调宽比例与 minWidth、陈旧会话、共享菜单解析、语言 alias |
| [document-structured-auxiliary-layout.contract.test.mjs](../../packages/office-studio/tests/document-structured-auxiliary-layout.contract.test.mjs) | 6 个：顶层结构真实文字、表格 containers/clip/cell target、footer 底部对齐、物理 host 命中和光标、结构内 PAGE 源坐标、gutter/marker 平移 |
| [document-structured-body-layout.contract.test.mjs](../../packages/office-studio/tests/document-structured-body-layout.contract.test.mjs) | 2 个：原 NumberUnit 保持、改变转 px；非法尺寸失败且输入不变 |
| [document-structured-facade.contract.test.mjs](../../packages/office-studio/tests/document-structured-facade.contract.test.mjs) | 2 个：pt/px 默认/插入/更新、一次历史与 undo/redo/保存重开；非法插入/更新无发布 |
| [document-structured-workbench.browser.test.mjs](../../packages/office-studio/tests/document-structured-workbench.browser.test.mjs) | 7 个：拖动 preview/释放一次/Escape、键盘调宽、主题和尺寸持久、Slider 最大值稳定与末端真实输入、真实菜单、旧浮动动作、页眉调宽/属性/重开；挂实际 Canvas + ConfigProvider + antdv-next |

本批仅通过 UTF-8 源码阅读、`rg`、`git status` / `git diff` 做静态核对，并修正了 Slider handle 标签、辅助空 runs 路径和容器遗漏。没有测试通过、类型通过、构建通过或浏览器成功记录。

仍需继续完成或验收：

1. 第三批已接六种 paragraph/PopMenu、表格下方、mobile Drawer、Code 字体/语言/More；全部基础菜单组合、真实触控和运行态 More 溢出仍未验收。Quote/Callout 的常规分组位置差异保持初批记录。
2. 第三批已接固定 21 项对应词法；复杂语言和全部源/配置组合、手动滚动后的光标显示、拖动与触控、低缩放、重复/嵌套物理片段和真实 IME 仍未验收。
3. Column modern/traditional 模式契约、超大 minima 的完整分页/裁剪、继承辅助故事、表格 continuation/repeated header、全部数据资源与权限组合；上游手柄 72 px 最小可见高度差异。
4. Quote keepLines/widowControl/相邻 keepNext 的生成、作者恢复及共享分页已在 [第四批](DOC-16-quote-pagination.md) 落盘，未运行验收；复杂分页兼容、复制粘贴结构组合、相对单位参考变化、DOCX/DOC 解析/写出和 Office 独立重开仍待补或验收。
5. 规定的前端、Rust/WASM 与浏览器检查；同一数据集的首屏、滚动、编辑、缩放、主线程长任务、传输复制和内存测量。当前共享测量和投影是 TypeScript 前端路径，没有新增 Worker/WASM 任务或零拷贝/性能达标证据。

中央对齐表和后端需求由主任务维护。本记录保留真实开发状态及未验收项，不能以入口或草稿数量替代完整功能验收。
