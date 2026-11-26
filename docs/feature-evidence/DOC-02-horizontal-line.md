# DOC-02 分隔线命令、边框与保存

记录日期：2026-11-27。状态：**已落盘、未验收**。关联 DOC-02 段落格式、DOC-03 story 和 DOC-22 读写。本批补生产实现与契约草稿，只做静态阅读、限定 diff 和路径核对；未运行测试、类型检查、构建、浏览器/dev、Rust/WASM 或基准。

## 固定来源与入口

参考只读仓库 `E:\rust_test\univer`，Univer v1.6.0，固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。没有新增线上演示或 Office 文件样本证据。

| 固定来源 | 已读契约 |
| --- | --- |
| [menu/schema.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts) | INSERT → MEDIA 的分隔线为 order 4、grid r2c5；classic 按组内 order 排序。正文段落、内容 gap、空段落、表格块和段落工具菜单另有上下文挂载。 |
| [menu/menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts) | `HorizontalLineFactory` 为 Button、ReduceIcon、`docs-ui.toolbar.horizontalLine`；无有效 Docs range 时禁用，非 Docs 单元时隐藏。 |
| [menu/paragraph-menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/paragraph-menu.ts) | 空段落的 `EmptyParagraphHorizontalLineMenuItemFactory` 执行 horizontal-line；`InsertHorizontalLineBellowMenuItemFactory` 执行 bellow。两个 factory 本身没有额外 disabled/hidden observable；实际位置由父菜单上下文决定。 |
| [doc-horizontal-line.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/doc-horizontal-line.command.ts) | `doc.command.horizontal-line` 转发 BreakLineCommand，默认 width 1、padding 5、RGB #CDD0D8、SOLID。`doc.command.insert-horizontal-line-bellow` 优先消费 DocContentInsertService 的结构 gap，否则使用当前 paragraph.startIndex + 1。 |
| [break-line.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/break-line.command.ts) | 复用 RichTextEditingMutation 和既有选区/段落拆分；水平线参数进入段落生成。InsertGap 与文字插入是不同路径。 |
| [docs/utils/paragraphs.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/utils/paragraphs.ts) | 参数最终写入 `paragraphStyle.borderBottom`，并设置 `spaceBelow: { v: 10 }`；没有独立的 horizontalLine 段落字段。新段继承原 styleId、bullet 和段落属性，先删除旧 borderBottom。 |

按 schema 的静态挂载分别记录，不能把这些数量解释为功能验收或运行态可见总数：

| 分类 / 标签 | 父位置与顺序 | 命令 | 当前本地接线 |
| --- | --- | --- | --- |
| 常驻顶部 / 分隔线 | INSERT → MEDIA，order 4，r2c5 | horizontal-line | [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 的 `doc-divider` 已由主任务改为真实 ID；根、分组和位置保留 |
| 段落右键 / 下方插入 → 分隔线 | PARAGRAPH → LAYOUT → INSERT_BELLOW_MENU_ID，子项 order 3 | bellow | 命令已注册；对应完整右键子项仍待接线 |
| 内容 gap / 分隔线 | DOC_CONTENT_INSERT_MENU_ID → LAYOUT，order 3 | bellow | 结构 gap 的 consumeInsertRange 路径待接线 |
| 空段落菜单 / 分隔线 | EMPTY_PARAGRAPH_MENU_ID → LAYOUT，order 4 | horizontal-line | 命令已注册；专用菜单待接线 |
| 表格块右键 / 下方插入 → 分隔线 | DOC_TABLE_BLOCK_MENU_ID → LAYOUT → INSERT_BELLOW_MENU_ID，子项 order 3 | bellow | 表格结构 gap 和专用入口待接线 |
| 段落浮动插入 / 分隔线 | DOC_PARAGRAPH_T_INSERT_MENU_ID → quickBottom，order 5 | horizontal-line | 专用浮动入口待接线 |
| 段落浮动下方插入 / 分隔线 | DOC_PARAGRAPH_T_INSERT_BELOW_MENU_ID → quickBottom，order 5 | bellow | 专用浮动入口待接线 |

本切片为 **1 个常驻顶部根、0 个插件扩展顶部根、0 个独立溢出根、6 个上下文子项挂载**。同一根的 classic/grid 和共享 More 呈现不重复计数。当前三个本地命令 ID 包括一个保留的旧别名，不代表三个功能。完整顶部与 Pro 范围继续按[工具栏台账](../UNIVER_TOOLBAR_PARITY.md)和[完整功能清单](../UNIVER_FEATURE_PARITY.md)核对。

## 本地模型、事务与数据

[document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) 保留 `doc.command.insert-divider`，新增真实 `doc.command.horizontal-line` 和 `doc.command.insert-horizontal-line-bellow`；[document-editor-commands.ts](../../packages/office-studio/src/api/document-editor-commands.ts) 统一注册，Facade 和顶部复用同一文档候选与内容历史。

horizontal-line 复用既有 `insertBlocks`，分隔线置于拆分后的 head 与可编辑 tail 之间，作者文字、样式、字段和内容控件走既有变换。bellow 在公开命令入口将 focus 转为当前作用域段落结尾的折叠范围，再执行 horizontal-line；权限、内容锁和控件 affinity 检查使用这个实际插入范围。API 事件与单次内容历史保留调用的 bellow ID。结构表格选区需要合法文字位置，当前明确拒绝；没有用空操作代替结构 gap。

主任务同步补齐 [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 的 `documentCommandFromRibbon`，适配 horizontal-line、bellow 和原别名；顶部按钮同步 `editingDisabled`。这保证真实 Ribbon ID 能进入既有执行路径，不能仅以 schema 中的 ID 更新判断接线完成。

[document-horizontal-line.ts](../../packages/office-studio/src/features/document/model/document-horizontal-line.ts) 提供 `DocumentHorizontalLineBorder` 和默认边框。沿用已有 `DocumentBlock.type = 'divider'` 与 `metadata.documentDividerBorder`，没有第二套领域树。border getter 供解析、写出和两条布局路径复用；新增 `LayoutDivider.style` 为可选的 solid/dot/dash，没有更改 DrawCommand 或 Worker/WASM 协议。段后 10 在本地 CSS px 模型中记为 `10 × 4/3`，写回公开默认 NumberUnit 仍为 `{ v: 10 }`。

[document-body-paragraph.ts](../../packages/office-studio/src/api/document-body-paragraph.ts) 将空 borderBottom 段落归一为 divider，保留 width、padding、color、dashStyle 和来源未知属性。归一发生在标题/列表解析之后，避免继承状态覆盖分隔线；列表的公开来源保留，divider 不新增列表求值引擎。非空 borderBottom 段落保持真实文字；即使旧 Office 来源扩展声明 divider，也按当前段落文字与类型恢复，避免写出时丢弃新文字。入口拒绝负 width/padding 和非 0–3 的 dashStyle；内部不重复校验。

[document-body-write.ts](../../packages/office-studio/src/api/document-body-write.ts) 的 divider 分支复用 paragraph writer，写当前 borderBottom，并保留源边框未知字段。默认 SOLID=1、DOT=2、DASH=3；公开 UNSPECIFIED=0 在本地渲染为 solid，写出归一为 1，这项枚举差异尚未用独立 Office 样本核实。[document-body-source.ts](../../packages/office-studio/src/api/document-body-source.ts) 将 documentDividerBorder 列为派生字段，不在公开 Office 来源扩展中另存一份边框状态。

## 布局、Canvas 与界面

[document-block-layout.ts](../../packages/office-studio/src/features/document/layout/document-block-layout.ts) 与 [document-block-measurement.ts](../../packages/office-studio/src/features/document/layout/document-block-measurement.ts) 共用边框，线条厚度来自 width，纵向位置来自 padding，占位高度为 `max(18, padding × 2 + width)`。正文/story 分页与表格测量使用同一数据；已有 note separator 的宽度比例、位置和颜色规则保留。

[document-scene.ts](../../packages/office-studio/src/features/document/rendering/document-scene.ts) 按缩放后的厚度绘制实线，按同一 DrawCommand 分段绘制点线和虚线；width 0 不绘制。Canvas2D/WebGPU 继续消费同一场景，没有新增渲染后端、传输或性能声明。

本批没有新增 UI 控件。主任务的顶部按钮仍使用统一 Ribbon/`antdv-next` 路径和 ConfigProvider 主题。明暗主题、classic/grid 溢出、点击后的焦点恢复、真实输入法与画布交互尚未执行验收。

## 契约草稿与剩余验收

[document-horizontal-line.contract.test.mjs](../../packages/office-studio/tests/document-horizontal-line.contract.test.mjs) 新增 **9 个 test 声明，全部未运行**。草稿调用真实 OfficeStudio/FDocument、codec、content controls、story host、布局与场景；覆盖两个插入模式及旧别名、作者文字和样式、单次历史/undo/redo/save、默认边框、非默认边框和未知字段、标题/列表继承、非空段落来源恢复、边界拒绝、正文/单元格线型与厚度、内容锁/权限拒绝和 header story 隔离。主任务在 [document-toolbar-state.contract.test.mjs](../../packages/office-studio/tests/document-toolbar-state.contract.test.mjs) 另补真实 Ribbon → 适配 → 领域插入、旧别名/bellow 和禁用态草稿，也未运行。

以下保留为缺口和未验收项：

- 上下文右键、空段落、浮动工具和结构 gap 仍需完整接线。bellow 的段落结尾路径不能作为 DocContentInsertService.consumeInsertRange 的验收。
- 本地显式 divider 加可编辑 tail 与固定版带 borderBottom 的段落拆分、继承样式/列表以及 InsertGap 语义存在结构差异；所有位置、编号标记与编辑体验仍需浏览器及公开样本核对。
- 非空段落的下边框当前保留数据，完整段落边框绘制与编辑仍待实现；不以数据保留宣称渲染完成。
- 本批只补公开 body 和快照路径，真实 DOCX 解析/写出、独立 Office 重开、段落边框未知属性和线型保真尚未验收。
- 正文、表格、story、note 的生产输入闭环、主题、焦点、Windows IME 与大样本性能/内存未测量。没有运行结果，不能把契约草稿或静态接线标为已验收。

限定 `git diff --check` 通过；新增文件无尾随空白。静态核对 15 个文档本地链接、6 个固定源码链接、两个新源文件的 8 个相对 import 和两个契约的 19 个 SSR 模块路径，均存在。静态检查只证明本批 diff、导入和引用路径可复查；整个 DOC-02、Docs/Sheets 及 Pro 功能的完整交付目标保持不变。
