# DOC-12 脚注、尾注与分隔故事

状态：**未验收**。更新日期：2026-11-27。本记录覆盖已有实现、本批契约草稿与仍须完成的边界，不能证明 Docs、DOC-12 或 Pro 能力整体完成。本轮按用户已确认的“先实现，后统一检查”顺序编写，未运行测试、类型检查、前端构建、Rust/WASM、dev、浏览器、Office 文件重开或基准。

关联：[功能清单](../UNIVER_FEATURE_PARITY.md)、[引用规则](../UNIVER_REFERENCE_RULES.md)、[执行标准](../../PROJECT_EXECUTION_STANDARD.md)、[渲染与计算标准](../RENDERING_AND_COMPUTE_STANDARD.md)、[DOC-11 控件证据](DOC-11-content-controls.md)。

## 固定来源与入口清点

固定 OSS 来源为 `E:/rust_test/univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。已查阅 `packages/engine-render/src/components/docs/layout/{footnote-layout,endnote-layout,note-numbering}.ts`。Pro 入口查阅本项目只读缓存 `.research.local/univer-1.6.0/packages/docs-reference-ui/lib/es/index.js`、`lib/types/menu/{schema,menu}.d.ts` 与 `lib/types/commands/operations/{footnote,navigate-note}.operation.d.ts`；发行包与固定 OSS 的构建提交是否一致未核实，没有安装商业 SDK。菜单按压缩产物中的公开命令与 schema 字符串复查，不把内部变量作为稳定 API。

References 标签顺序 3；Footnote / Endnote 分组顺序分别 1 / 2。每组四个根入口，共八个 Pro 扩展根入口；插入与设置对话框、菜单子项和上下文操作须单独追踪，八个按钮不是八个已完成功能。

| 每类入口 | 公开命令 ID | grid 位置 / 顺序 | 当前能力判断与行为 |
| --- | --- | --- | --- |
| 插入脚注 / 尾注 | `docs.operation.insert-footnote` / `docs.operation.insert-endnote` | r1c1，rowSpan 2，order 0 | 传统正文、真实 UTF-16 文字位置、权限与锁检查；非空选区在逻辑末端插入引用，保留选中文字，返回真实 note host 的正文选区。 |
| 上一条 | `docs.operation.previous-footnote` / `docs.operation.previous-endnote` | r1c2，order 1 | 按真实引用位置严格向前寻找，不循环；跨 kind 位置记忆。 |
| 下一条 | `docs.operation.next-footnote` / `docs.operation.next-endnote` | r1c3，order 2 | 按真实引用位置严格向后寻找，不循环。 |
| 设置 | `docs.operation.open-footnote-settings` / `docs.operation.open-endnote-settings` | r2c2，colSpan 2，order 3 | 文档与当前节作用域、继承/显式设置、位置/编号/起始值/重启/脚注列、尾注延后；草稿未应用时阻止隐式切作用域。 |

分类计数：固定版扩展 Ribbon 根入口 8，分组 2；本项目增加管理根入口 1（`doc.operation.open-notes`），辅助检查/管理视图单独记账；上下文返回引用命令 1（`docs.operation.navigate-footnote`），上游挂在 MAIN_AREA / PARAGRAPH 的 OTHERS 两处，不计入常驻根入口。2026-11-27 已回读固定发行包的 schema 与菜单工厂，此分类补充不计本轮新增能力。当前没有另外声称已盘点 Docs 全部插件、溢出入口或 Ribbon 功能完成。classic 布局使用相同分组/顺序，实际尺寸、显隐、键盘和最终交互未作浏览器验收。

`docs.operation.navigate-footnote` 的公开参数为 `{unitId?,noteId?,toReference?}`；指定 ID 默认进入注释，无 ID 默认回到当前引用，显式 `toReference` 覆盖。自定义标记不消耗自动编号；`eachPage` 使用实际物理页而非印刷页码。现代文档 flavor、链接故事、嵌套表格路径与无引用孤立定义都保留来源；不靠集合索引猜引用地址。

## 模型、命令、布局与持久化

权威内容为 `DocumentStoryContent.blocks`；runs/paragraphs 为派生桥。领域、服务、命令、布局与 Vue 控件分层；Canvas 编辑器复用段落、表格、图像内核，UI 使用统一 ConfigProvider / `antdv-next`。本批布局仍为主线程派生计算，没有声称已实现或优化 Worker/Rust/WASM 注释排版。

插入、完整内容替换和分隔设置通过 detached candidate 原子发布。`contentControls` 与对应 note content 一起复制；独立 `decorationContentControls` 仅为同次 settings 请求的三个分隔故事提供 SDT 所有权，不作为持久化设置属性。实际复制映射同时处理块、run、表格行/格、字段、父控件、blockRange、tableRowRange 与图片锚点；外部 parent/range/live reference、非法 UTF-16、临时 story lock 控件、后续无效故事均拒绝整批候选。未知 nested 来源 XML/codec metadata 保留，不把 provenance 中的旧标识改成实时地址。完整替换/删除遵守故事锁、SDT wrapper/content 锁、字段锁，并仅清理原内容区控件与批注段。

三个活动或默认分隔故事，以及每个休眠特殊定义，均为独立 `noteDecoration` host；它们不分配普通注释编号或正文引用。公开来源为 `metadata.univerNoteDecorationSource={version:1,kind,type,noteId:string}`；额外定义保存在 `metadata.officeDormantNoteDecorations={version:1,contents:[...]}`。公开身份不伪造 DOCX 部件/XML；原生 `metadata.decorationSource` 保持数字 ID、partName、sourceXml 的严格要求。两个来源同时存在时分别验证、类型须一致；活动与休眠条目分别展示公开与有效原生 ID，原 native 记录保留。替换保持旧来源身份；设置重置保留 dormant/source provenance，额外故事通过自身 host 编辑。根任务负责公开 codec 与原生 IO，不以领域接线声称文件往返已通过。

独立检查画布展示真实 note/decoration 的完整段落、嵌套表格和图像，并将所有几何标为辅助宿主；正文根几何、正文物理分页、原页眉页脚与编号不被检查页替换。默认空 continuationNotice 是稳定、可编辑的 detached paragraph，打开检查不写模型，首次真实编辑命令才落盘。返回正文使用捕获的选区/缩放/滚动，并在发布后复核 host/revision/generation/选区/epoch，避免观察者重入抢焦点。设置与标记草稿提供显式应用/还原/清空，不能因进入检查或改变 scope 静默丢弃。

scoped view 与 edit 独立判断。preview、note panel、Canvas 导航与分隔检查检查实际实体/段落/节目标；权限收回清预览与对应 UI 草稿。权限面板已增加 View 及 note/default/active/dormant/story 的实体与段落目标；辅助目标标签不投影内容文字。这些是本地能力判断，整画布内容遮蔽与协作服务 ACL 尚未实现，不能宣称已提供完整权限保密隔离。

权限目标的 View、六类页眉/页脚、脚注/尾注、三类分隔与保留来源标签已补齐现有 19 套 locale；翻译与实际面板仍未运行验证。Notes / Settings Modal 使用 `antdv-next` 的 `afterClose` callback 和 `focusTriggerAfterClose=false`；Workbench 在动画结束时复核文档、输入 generation/revision/location、UI epoch、host、选区、权限与其他弹窗，再恢复当前 Canvas 输入。返回正文也拒绝覆盖观察者后来的选区/修订；Canvas 显式恢复滚动偏移后用 `viewportRestoreVersion` 跳过同次 host/selection/zoom 已排队的自动定位，后续用户选区继续正常定位。项目当前 `antdv-next@1.8.2` 的 `dist/input/TextArea.js` 中 exposed `focus: () => textAreaRef.value?.focus?.()` 不转发参数，因此优先使用组件 exposed `nativeElement`，只对库管理的 textarea 使用原生 `focus({preventScroll:true})`；没有添加原生表单控件。这些是静态实现依据，焦点归还与 IME 尚未浏览器验收。

普通脚注使用物理页底部预留与多列流，尾注占用后续正文同一物理流。小分隔故事的测量也使用真实 `noteDecoration` 地址；固定预留路径的文字、表格、图片和 divider 几何标注同一真实 host，legacy runs/paragraph 的测量与编辑使用一致来源 ID。超页 separator / continuationSeparator / continuationNotice 的调用方现已转入完整 block 分页流：脚注整组内容与装饰统一预留/平移，装饰单栏而 note 保留实际列；尾注使用与后续正文相同的物理流。`DocumentNoteStoryPagination` 只在当前页已有真实 note 几何后添加续页装饰；装饰自身跨页时禁止递归重复，装饰独占满页后下一页不再次插分隔符，从而允许 note 前进。真实 paragraph/table/image 与 `continuedBefore/continuedAfter`、host、source span、clip 均保留，不造普通 note 或引用；物理页上限 40,000 是预算而非性能达标证明。

大 continuationNotice 作为实际顺序故事流从前页剩余空间延续，随后完整 continuationSeparator 再进入 note；它可能占据只有装饰的页面。此策略与最终 Word/Univer 的续页位置和分隔显示语义尚未核实，布局发出 `DOCUMENT_NOTE_DECORATION_PAGINATED` 记录真实状态。单个不可拆超页 glyph/对象仍保留完整源与排版对象，但视觉裁剪及固定点未收敛需独立诊断；不能以本批接线声称所有极端内容均已可见或达标。

分隔字段的来源与几何需要分开验收。固定 OSS `layout/footnote-layout.ts::defaultSeparatorBody`（205）使用 `U+FFFC` 及覆盖其单字符范围的 `CustomRangeType.FIELD`，fieldType 为 `FOOTNOTE_SEPARATOR` / `FOOTNOTE_CONTINUATION_SEPARATOR`；`layout/block/paragraph/shaping.ts::createNoteGlyph`（332–350）将首分隔宽度设为 `min(192,textWidth)`、续页分隔为 `textWidth`。`extensions/font-and-base-line.ts`（124–133）以 `em/20` 厚度绘制，底边在 baseline 上方 `em/4`。本地此前统一以可用文字宽度 `widthFraction` 绘制 1px divider，位置为首行底部减 2px；默认与明确 FIELD 的新实现见[分隔线字段与几何](#note-separator-field-geometry)。未标记自定义分隔内容仍保留原行为。公开字段映射、空段锚点和本次几何实现均不能证明 Office 显示已经验收。

## 未运行的验证草稿与缺口

| 草稿文件 | 断言范围 | 状态 |
| --- | --- | --- |
| `packages/office-studio/tests/document-note-controls.contract.test.mjs` | 同图复制、block/row/field/picture/SDT map、跨 kind ID、来源隔离、foreign/UTF-16/临时锁拒绝、note/settings 原子回滚、批注外部段、完整替换锁、settings no-op/source/dormant | 未运行。 |
| `packages/office-studio/tests/document-note-inspection.contract.test.mjs` | 真实 table/image/orphan 检查、body 页保持、活动/默认/原生/公开 dormant 区分、默认首次注册编辑、撤销重做/内部保存重开、settings/mark 草稿能力、scoped view、陈旧会话 | 未运行。 |
| `packages/office-studio/tests/document-note-pagination.contract.test.mjs` | 默认及公开 FIELD 分隔线宽度/厚度/baseline/缩放、公开 Body 编码重开与来源身份、自定义规则保留；小型 legacy 分隔测量/编辑 ID 与真实 host、separator 恰好占满页不递归、foot/end 完整 notice 与 continuation source、table cell 全部 UTF-16 span、完整 image 比例/clip、连续 table strip、多栏 note/单栏装饰、parity blank、sectEnd 后续正文、模型/引用不变 | 未运行。 |
| `packages/office-studio/tests/document-notes.contract.test.mjs` | 插入真实 note body、跨段保留文字、两位引用、UTF-16、源定义冲突、nested 删除后重读权威模型 | 修改草稿；未运行。 |
| `packages/office-studio/tests/document-permissions.contract.test.mjs` 与 `document-permission-rules.contract.test.mjs` | view/edit 独立、deny/default-deny 与多目标、note/decoration 目标清单 | 修改草稿；未运行。 |
| `packages/office-studio/tests/document-note-workbench.browser.test.mjs` | 真实 Workbench/FDocument/ConfigProvider；纯图片来源空段锚点与实际 Canvas 像素、separator-only/public dormant、首次 notice 编辑/撤销重做/saveModel 重开、scope/mark 草稿、view/edit、动画关闭焦点、正文 scroll/zoom、编辑/返回观察者重入、明暗 classic/grid、合成 composition 的单提交/取消/旧节点 | 新增草稿；未运行。合成 composition 不代表 Windows IME 验收。 |

已补 Notes 设置/管理/独立检查的真实 Workbench 浏览器草稿，但尚未运行浏览器闭环；明暗主题、弹层、焦点归还、滚动/缩放、Windows IME、观察者重入和历史跳转均未验收。公开 Body/native DOCX 包含控件、字段、图片、复杂表格、额外特殊定义的保存重开由根任务继续处理；浏览器草稿里的本地领域模型保存使用 `saveModel()`，不把内部快照重开等同于公开 `save()` / `getSnapshot()` 或 Office 文件往返。纯图片/分隔故事必须由 Body/native 适配保留来源中真实 `w:p` 或公开 paragraph；UI 不生成临时 paragraph 充当保真输入。没有 Office 文件样本就保持未验收。缺少完整 nested table DocumentPoint 路径的引用不可直接导航，需独立检查；性能、大样本内存、首屏/滚动/编辑/缩放/长任务、跨 Worker/WASM 字节与取消均无实测。

本记录不缩减 Sheets/Docs/Pro 的完整目标，也不将未完成项改写为永久禁用、展示入口或提示消息。

<a id="note-insert-selection-permission"></a>

## 2026-11-27 非空选区插入的权限目标

状态：**未验收**。固定 v1.6.0 的 `docs-reference-ui/lib/es/index.js` 中插入操作使用 `Math.max(startOffset, endOffset)`；八个根入口的 ID、顺序和 grid 位置已按同文件 schema 回读。上游设置菜单只按传统 flavor 显隐，本项目设置入口仍允许打开只读 Form，由作用域能力决定字段和提交是否可用，不能把该入口当成遗漏编辑禁用。

[document-note-command-state.ts](../../packages/office-studio/src/features/document/commands/document-note-command-state.ts) 和 [document-note-editing.ts](../../packages/office-studio/src/features/document/notes/document-note-editing.ts) 已按非空选区的逻辑末端判断权限、字段/内容锁并插入引用，保留原选中文字；此前 [document-note-mutations.ts](../../packages/office-studio/src/features/document/notes/document-note-mutations.ts) 的入站权限仍检查整个选区，导致起点段落无编辑权限而末端可编辑时，按钮可用、命令却拒绝。本次生产只改该文件两行：复用 `orderedSelection` 与 `collapsedSelection`，将 insert 权限目标收窄到真实插入点。候选末尾对实际变更的块、注释、编号和设置仍执行原权限判断，若重编号确实修改其他被拒绝区域，仍拒绝整批提交；没有新增守卫、权限服务、API/DTO、历史或 Worker/WASM 协议。

[document-notes.contract.test.mjs](../../packages/office-studio/tests/document-notes.contract.test.mjs) 增加一个声明，内部覆盖脚注/尾注及正反选区四个场景。草稿直接调用原能力、候选规划与 OfficeStudio 注册命令，检查起点 paragraph edit deny 和 richText contentLocked 下的正文/控件不变、末端引用、真实 note 输入 host、一次历史、撤销重做与 saveModel 重开；末端 edit deny 时，按钮能力与命令均拒绝，模型、revision、历史和输入会话不变。模型比较捕获真实快照，不依赖 getModel 返回的活动对象身份。

[document-note-workbench.browser.test.mjs](../../packages/office-studio/tests/document-note-workbench.browser.test.mjs) 复用原 fixture，仅增加调用既有 `FDocument.setInputSelection` 的选区入口；新增一个声明展开四个计划案例，分别检查两类注释的 classic/light 正向与 grid/dark 反向选区。通过真实 Ribbon 按钮检查取消后模型/历史不变、末端引用与原选中文字保留、单次 CommandExecuted/历史、note host/Canvas 输入焦点、撤销重做/saveModel 重开及末端拒绝时的禁用状态。现有 ConfigProvider、Ant 按钮与 Input.TextArea 输入桥不变，未新建 fixture、测试文件或 UI 控件。

仅执行源码/diff 回读和 Babel AST 计数：两个草稿文件本次前均为 **11 个声明 / 11 个计划案例**；当前契约为 **12 / 12**，浏览器为 **12 / 15**，运行结果均为 **0**。契约声明内部的 kind/方向循环不额外计注册案例。统一静态检查由主线收尾；测试、类型、构建、dev/服务、浏览器、Rust/WASM、公开 Body/native DOCX 样本和基准均后置。新增 saveModel 草稿只覆盖本地模型重开，不能代替公开 save/getSnapshot 或 Office 文件保真；真实主题、弹层、焦点和 Windows IME 仍未验收。

<a id="note-separator-field-geometry"></a>

## 2026-11-27 分隔线字段与几何

状态：**未验收**。本节记录默认分隔和明确公开 FIELD 的实现，替代上文对这两类内容使用 1/3 宽度、1px 厚度、首行底部减 2px 及默认段后距的历史说明；未标记自定义 `noteSeparator` 继续使用原规则。

固定来源仍为 `E:/rust_test/univer` 的 v1.6.0 / `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。[默认来源](../../../univer/packages/engine-render/src/components/docs/layout/footnote-layout.ts)、[字段 shaping](../../../univer/packages/engine-render/src/components/docs/layout/block/paragraph/shaping.ts) 和[字体基线绘制](../../../univer/packages/engine-render/src/components/docs/extensions/font-and-base-line.ts) 分别给出真实空段及 FIELD、首分隔 `min(192,textWidth)` / 续页全宽、段前后零、厚度 `em/20` 与底边 `baseline-em/4`。

[document-note-decorations.ts](../../packages/office-studio/src/features/document/notes/document-note-decorations.ts) 为默认首分隔与续页分隔增加准确 `fieldType`，段后距改为零，保留原宽度分数与颜色。主任务配套修改 [document-body-ranges.ts](../../packages/office-studio/src/api/document-body-ranges.ts) 和 [document-body-read.ts](../../packages/office-studio/src/api/document-body-read.ts)：公开 `CustomRangeType.FIELD` 与实际分隔 token 的解码结果携带 `FOOTNOTE_SEPARATOR` / `FOOTNOTE_CONTINUATION_SEPARATOR`，保留 `univerNoteSeparatorRangeSource` 等原来源；不根据 `widthFraction === 1/3` 推测字段身份。

[document-paragraph-layout.ts](../../packages/office-studio/src/features/document/layout/document-paragraph-layout.ts) 使用已形成的真实 `LayoutLine.baseline`：首分隔宽度为 `min(192,column.width)`，续页宽度为 `column.width`，`em=fontSize*4/3`，厚度为 `em/20`。本地场景采用中心线 stroke，因此分隔线 y 为 `baseline-em/4-thickness/2`，使绘制底边对应上游位置。小分隔测量和分页分隔继续复用同一 paragraph/block 内核，保留真实空段输入地址。未标记自定义内容继续使用 `geometry.width(0)*widthFraction`、原行底位置、1px 厚度与原颜色。

[document-layout-types.ts](../../packages/office-studio/src/features/document/layout/document-layout-types.ts) 的 `LayoutDivider` 新增可选派生 `noteSeparator` 标记；已有布局平移和裁剪保留该标记。[document-scene.ts](../../packages/office-studio/src/features/document/rendering/document-scene.ts) 对标记分隔线直接使用 `height*scale`，保留缩小后的分数像素宽度；其他 divider 仍按原规则绘制。持久化来源 metadata 增加 `fieldType`，公开 API、命令、历史及 Worker/WASM 协议没有新增接口。本批未改变 Vue 控件或输入桥，未新增原生表单；既有 `antdv-next` 主题、焦点和 IME 仍待浏览器验收。

[document-note-pagination.contract.test.mjs](../../packages/office-studio/tests/document-note-pagination.contract.test.mjs) 新增三个声明草稿：默认两类注释在 200/360/800px 页宽下的首分隔与续页几何、零段后距和真实空段；未标记 1/3、1/2、1 自定义规则的原行为；两种公开 FIELD 通过真实 `decodeDocumentBody → encodeDocumentBody → decodeDocumentBody` 保留单字符范围、range ID、fieldType、空段及 12/18pt 字体，再进入真实 note settings、layout 与 scene 检查宿主、来源和几何。默认及公开草稿都检查 50/100% 缩放与布局不修改模型。这些是公开 Body 往返契约草稿，不代表已验证整文档 `save/getSnapshot`、DOCX 文件往返或浏览器 Canvas 像素。

本批只做源码与限定 diff 回读、空白差异和文档来源链接检查；测试、类型、前端构建、Rust/WASM、dev/服务、浏览器、Office 样本和性能/内存基准均未运行。真实 Office 与浏览器显示、特殊字体/缩放、极端分隔内容以及 Worker/Rust/WASM 排版仍保留未验收状态。
