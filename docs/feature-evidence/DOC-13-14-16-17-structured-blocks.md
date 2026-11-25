# DOC-13 / DOC-14 / DOC-16 / DOC-17 结构块静态实施记录

记录日期：2026-11-25。状态：**未验收，前端开发中**。本阶段按用户要求先实现、后统一验证，后端暂缓。本轮没有执行测试、类型检查、构建、Cargo/WASM、浏览器、dev 或性能基准。

本记录覆盖代码块、多栏组、引用块和 Callout 的前端模型、命令、属性面板、Facade、body 编解码及测试草稿。入口计数说明清点口径，不作为完整功能证明；Pro 全部能力、Office 互操作和真实性能要求保持在交付范围内。

## 1. 依据与模块边界

- OSS：只读仓库 `E:\rust_test\univer`，Univer v1.6.0，固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。
- Pro：项目研究缓存 `.research.local/univer-1.6.0/packages/{docs-code,docs-column,docs-quote,docs-callout}` 及相应 `-ui` 包。采用发布包的类型、公开命令和可读 ES 分发代码作行为依据，不将缓存包加入生产依赖。
- 公开规格：[代码块](https://docs.univer.ai/zh-CN/guides/docs/features/code)、[多栏](https://docs.univer.ai/zh-CN/guides/docs/features/columns)、[引用](https://docs.univer.ai/zh-CN/guides/docs/features/quote)、[Callout](https://docs.univer.ai/zh-CN/guides/docs/features/callout)。链接属于动态规格；本轮静态核对本地发布包，未新增在线演示或 Office 样本验证。
- 执行规则：[项目架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)、[参考规则](../UNIVER_REFERENCE_RULES.md)、[功能清单](../UNIVER_FEATURE_PARITY.md)、[渲染/WASM 契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)。

领域权威数据为 `DocumentModel` 的真实块树、段落运行、控件/字段/注释及来源 metadata。Vue/`antdv-next` 负责 UI；结构命令和服务负责候选事务；纯 TypeScript 布局产生 Canvas 场景。已有渲染后端接收同一布局；本次未增加 Rust/WASM 布局引擎，也未证明主线程布局或缓冲传输满足大样本标准。

可直接复查的 Pro 来源：

| 能力 | 入口源码 / 声明 | Facade 声明 |
| --- | --- | --- |
| Code | [ES schema 与 factory](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/es/index.js)、[schema 导出](../../.research.local/univer-1.6.0/packages/docs-code-ui/lib/types/menu/schema.d.ts) | [FDocumentCode](../../.research.local/univer-1.6.0/packages/docs-code/lib/types/facade/f-document-code.d.ts) |
| Column | [ES schema 与 factory](../../.research.local/univer-1.6.0/packages/docs-column-ui/lib/es/index.js)、[schema 导出](../../.research.local/univer-1.6.0/packages/docs-column-ui/lib/types/menu/schema.d.ts) | [FDocumentColumnGroup](../../.research.local/univer-1.6.0/packages/docs-column/lib/types/facade/f-document-column-group.d.ts)、[FDocumentColumn](../../.research.local/univer-1.6.0/packages/docs-column/lib/types/facade/f-document-column.d.ts) |
| Quote | [ES schema 与 factory](../../.research.local/univer-1.6.0/packages/docs-quote-ui/lib/es/index.js)、[schema 导出](../../.research.local/univer-1.6.0/packages/docs-quote-ui/lib/types/menu/schema.d.ts) | [FDocumentQuote](../../.research.local/univer-1.6.0/packages/docs-quote/lib/types/facade/f-document-quote.d.ts) |
| Callout | [ES schema 与 factory](../../.research.local/univer-1.6.0/packages/docs-callout-ui/lib/es/index.js)、[schema 导出](../../.research.local/univer-1.6.0/packages/docs-callout-ui/lib/types/menu/schema.d.ts) | [FDocumentCallout](../../.research.local/univer-1.6.0/packages/docs-callout/lib/types/facade/f-document-callout.d.ts) |

## 2. 固定版 Pro 入口清点

### 2.1 常规 Ribbon、扩展及浮动挂载

以下符号来自相应 `-ui/lib/es/index.js`。ES 分发文件为单行，复查时使用表内符号、导出名和属性片段。`lib/types/menu/schema.d.ts` 可核对 schema 导出名称。

| ID | 包与符号 | 标签 key / 分组 | classic 顺序 / grid 位置 | 命令 ID | 状态来源 |
| --- | --- | --- | --- | --- | --- |
| DOC-17 | `docs-callout-ui`：`st` / `ct` / `lt` | `docs-callout-ui.menu.callout`；INSERT → MEDIA | order 5；row 2 / column 6 | `docs-callout.command.insert` | `V`：Docs 类型显隐、页眉页脚编辑禁用；浮动 `at` 还隐藏 block-range 选区 |
| DOC-16 | `docs-quote-ui`：`Ge` / `Ke` / `qe` | `docs-quote-ui.menu.quote`；INSERT → MEDIA | order 6；row 1 / column 6 | `docs-quote.command.insert` | `z`：Docs 类型显隐、页眉页脚编辑禁用；浮动 `We` 还隐藏 block-range 选区 |
| DOC-13 | `docs-code-ui`：`Pt` / `Ft` / `It` | `docs-code-ui.menu.code`；INSERT → MEDIA | order 7；row 1 / column 7 | `docs-code.command.insert` | `M`：Docs 类型显隐、页眉页脚编辑禁用；浮动 `kt` 还隐藏 block-range 选区 |
| DOC-14 | `docs-column-ui`：`F`，导出 `DocsColumnUIMenuSchema` | `docs-column-ui.menu.column` | 仅确认 paragraph context schema，无该包常规 Ribbon 根声明 | `doc.command.menu-insert-column-group` → `doc.command.insert-column-group` | `Ie` / `Ve`：有效文档、可解析插入偏移、当前位置不处于已有栏内 |

Quote / Callout / Code 的 `rootMenuSchema` 分别另在 `FLOAT_TOOLBAR_MENU_POSITION → FLOAT_TEXT_STYLE_MENU_ID` 挂载插入项，order 10 / 11 / 12。合并 schema 中出现相同片段不重复计为新常规根。Column 桌面和 mobile 插件各注册相同 `F`，不能双计为两个入口集合。

### 2.2 Quote / Callout / Code 的 paragraph context 挂载

每个包均有下列六个插入挂载。表中的顺序依次为 Quote / Callout / Code。

| 挂载路径 | 动作 / 命令后缀 | 顺序 | factory 状态 |
| --- | --- | --- | --- |
| `PARAGRAPH → LAYOUT → INSERT_BELLOW_MENU_ID` | 在下方插入；`insert-below` | 6 / 5 / 7 | Docs 类型显隐；页眉页脚编辑禁用 |
| `PARAGRAPH → DOC_CONTENT_INSERT_MENU_ID → LAYOUT` | 在下方插入；`insert-below` | 6 / 5 / 7 | 同上 |
| `PARAGRAPH → EMPTY_PARAGRAPH_MENU_ID → LAYOUT` | 插入；`insert` | 6 / 5 / 7 | 同上 |
| `PARAGRAPH → DOC_PARAGRAPH_T_INSERT_MENU_ID → quickBottom` | 插入；`insert` | 3 / 4 / 2 | 同上 |
| `PARAGRAPH → DOC_PARAGRAPH_T_EDIT_MENU_ID → quickBottom` | 插入；`insert` | 4 / 5 / 3 | 同上 |
| `PARAGRAPH → DOC_PARAGRAPH_T_INSERT_BELOW_MENU_ID → quickBottom` | 在下方插入；`insert-below` | 3 / 4 / 2 | 同上 |

Code 另有 `PARAGRAPH → LAYOUT → docs-code-ui.menu.language` submenu 根，order 0；`At` / `jt` 要求当前 paragraph target 是 CODE blockRange。`Nt` 按目录 `j` 注册 21 个选择，命令为 `docs-code.command.update`，参数 `{ config: { language } }`：plaintext、javascript、typescript、json、markup、css、markdown、yaml、sql、bash、python、java、c、cpp、csharp、go、rust、ruby、php、swift、kotlin，顺序 0–20。语言选择自身按 Docs 类型显隐。

### 2.3 Column 的完整 `F` schema

| `PARAGRAPH` 下的挂载路径 | 标签 / 命令 | 顺序与状态 |
| --- | --- | --- |
| `LAYOUT → INSERT_BELLOW_MENU_ID` | 分栏；menu insert 的 `.below` 入口 | order 3；`P` 解析下方插入偏移 |
| `EMPTY_PARAGRAPH_MENU_ID → LAYOUT` | 分栏；menu insert | order 3；`N` 解析当前插入偏移 |
| `DOC_CONTENT_INSERT_MENU_ID → LAYOUT` | 分栏；menu insert 的 `.below` 入口 | order 3；`P` |
| `DOC_PARAGRAPH_T_INSERT_MENU_ID → insert` | 分栏；menu insert | order 1；`N` |
| `DOC_PARAGRAPH_T_INSERT_BELOW_MENU_ID → insert` | 分栏；menu insert 的 `.below` 入口 | order 1；`P` |
| `DOC_TABLE_BLOCK_MENU_ID → LAYOUT → INSERT_BELLOW_MENU_ID` | 分栏；menu insert 的 `.below` 入口 | order 3；`P` |
| `DOC_PARAGRAPH_T_EDIT_MENU_ID → quickBottom` | `docs-column-ui.menu.addColumn`；`doc.command.add-column` | order 6；`Xe` 隐藏无栏内光标或已 5 栏；在当前栏右侧添加 |
| 同上 | `docs-column-ui.menu.deleteColumn`；`doc.command.delete-column` | order 7；`Ze` 隐藏无栏内光标或不足 3 栏 |
| 同上 | `docs-column-ui.menu.deleteColumnGroup`；`doc.command.delete-column-group` | order 8；`Qe` 隐藏无栏内光标 |
| 同上 | `docs-column-ui.menu.equalWidth`；`doc.command.resize-column-group` | order 9；`$e` 隐藏无栏内光标；提交全 1 的 widthRatios |

`N` 是 BUTTON_SELECTOR，picker `M` 的 `qe` 为 `[2,3,4,5]`，主按钮默认 2 栏。`A` / `Ge` 从 explicit offset、contentInsertRange、paragraph target、collapsed selection 依次取得偏移，再归一化到段落末尾并检查可插入性。factory 未单独声明 disabled 不表示绕过命令权限；命令结果仍须校验。

### 2.4 可复查计数与未清点范围

| 分类 | 此四包静态计数 | 口径 |
| --- | ---: | --- |
| 常规 Pro Ribbon 根 | 3 | Quote / Callout / Code 各 1；Column 0 |
| 浮动 text-style 插入挂载 | 3 | 与常规 Ribbon 分开 |
| paragraph/context 插入挂载 | 24 | 三种结构块各 6，Column 6 |
| Column context 结构操作 | 4 | 添加、删除栏、删除组、等宽 |
| Code context 语言 submenu 根 | 1 | 子选择另计 |
| Code 语言选择 | 21 | 目录 `j`；不是 21 个 Ribbon 根 |
| Column picker 栏数 | 4 | 2–5；同 picker 的重复挂载不倍增目录数 |

上述计数覆盖导出 schema，不包含每个 Canvas PopMenu / 移动浮动组件内部的全部动作、动态实例数量和运行态 More 溢出。相关 `views/menu.d.ts`、PopManager 注册和颜色/语言浮层仍需逐项追踪并验收，不能据此标记全部 Pro 入口完成。

## 3. 本地入口和真实命令

[document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 保留常规代码块入口（MEDIA，order 7，grid 1/7），在 INSERT 的“高级块”组提供引用、Callout、多栏三根：order 0 / 1 / 2，grid 1/1、2/1、1/2 跨两行。多栏含 2–5 栏和“设置比例和间距”五个子项，START 的块样式子菜单还有引用入口。这是本地当前布局差异：引用/Callout 尚未移到上游 MEDIA 1/6 与 2/6，多栏提供了额外常规入口；不能称布局完全对齐。

[document-structure-ribbon.ts](../../packages/office-studio/src/components/document/document-structure-ribbon.ts) 为命中的真实块注册四个上下文 Ribbon 动作，并以同一命令注册四个右键动作：属性、保留内容并移除容器、在块后继续输入、删除块和内容。Ribbon grid 顺序为属性 1/1、unwrap 1/2、exit 2/2、remove 2/1。禁用来自编辑权限、活动输入会话和 composition；命令还执行自身校验。

本地此四项能力对应 4 个常规根、5 个多栏子项、1 个既有 START 块样式子项、4 个结构上下文 Ribbon 动作和 4 个结构右键动作，按类别分别记录。没有执行运行态 More 溢出统计。

| 用户动作 | 本地命令 / 服务 | 实际数据效果 |
| --- | --- | --- |
| 包裹引用/Callout/代码 | `doc.command.wrap-structured-block` | 移动真实段落/子块进入容器；半开区间不吞入下一段 offset 0；保持内部运行 |
| 插入结构块 | `doc.command.insert-structured-block` | 创建真实子段落，CRLF/CR 归一化；Facade 支持明确公共偏移 / paragraph ID |
| 移除容器 / 删除内容 / 退出 | `unwrap-structured-block` / `remove-structured-block` / `exit-structured-block` | 保留并迁出内容、移除内容资源、或创建/定位块后真实段落；不存在空回调 |
| 引用线条/文字颜色 | `doc.command.update-quote-style` | 容器线颜色进 metadata；文字颜色修改真实 runs，受锁/权限校验 |
| Callout 属性/文字 | `doc.command.update-callout` | 图标、显示、颜色、边框、圆角、padding/尺寸配置；未明确改变文字色时保留混合运行颜色 |
| 代码源/配置 | `update-code-block` / `update-code-settings` | 同一源编辑服务更新真实段落、语言、行号、tabSize、wrap |
| Tab / Shift+Tab | `doc.command.indent-code-block` | UTF-16 范围的降序 edits；一份外层候选和一条内容历史；保留反向选区和运行样式 |
| 创建/配置多栏 | `insert-column-group` / `wrap-column-group` / `update-column-group` | 2–5 栏、稳定组/栏 ID、比例、间距、最小宽度、fixed/auto、shrink/stack |
| 添加/移除/移动栏 | `add-column` / `remove-column` / `move-column` | 添加真实输入段落；最少保留 2 栏；删除可保留内容或一并删除；不允许跨栏字段被反转 |
| 栏文字/插段 | `set-column-text` / `insert-column-paragraph` | LF 分隔的真实多个段落；插入已有嵌套段落前保留其真实容器；第一段保留请求 ID |

表中没有域前缀的短名均为 `doc.command.<短名>`。UI operation 只打开面板或将意图转入上述统一命令，Facade / Ribbon / 右键 / Tab 不分别修改模型。

## 4. 事务、权限、输入会话和 UI

[document-structure-commands.ts](../../packages/office-studio/src/features/document/commands/document-structure-commands.ts) 对候选模型完成控件范围、字段端点/锁、TOC 父关系、批注锚点、注释故事与重新编号的处理；验证 story 和权限后发布。删除完整块控件由整体结构 reconcile 判断锁，inline 控件在删除资源阶段检查。有保留子块时清除其祖先派生 HTML；失败不发布部分候选。代码源编辑使用既有 [document-code-editing.ts](../../packages/office-studio/src/features/document/services/document-code-editing.ts) 与内容绑定服务。

命令通过 `DocumentCommand` / `applyDocumentEditorCommand` 进入 OfficeStudio 既有候选、事件、输入会话和按文档隔离的内容历史。新 `remove-structured-block`、`update-quote-style`、`indent-code-block`、`set-column-text`、`insert-column-paragraph` 已纳入注册和标签。模型级函数不另建历史栈；保存仍使用既有 `FDocument.save()` / savepoint。本地快照闭环的运行验证尚未执行。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 打开面板时捕获 API 身份、unit、model revision、history/input generation、selection、block host。提交和关闭后的焦点恢复都重新验证上下文，拒绝旧 revision、重开后的 API、切换选区、失活和 composition。standalone 还检查模型对象、内容快照、父 modelValue 身份和输入 epoch，允许无文档 ID；同内容新对象也不能沿用旧设置目标。纯检查函数在 [document-structure-panel-state.ts](../../packages/office-studio/src/components/document/document-structure-panel-state.ts)。

新增 [DocumentStructuredBlockPanel.vue](../../packages/office-studio/src/components/document/DocumentStructuredBlockPanel.vue) 和修改过的 [DocumentCodeBlockPanel.vue](../../packages/office-studio/src/components/document/DocumentCodeBlockPanel.vue) 均使用 `antdv-next` Modal / Form / FormItem / Input / InputNumber / Select / Switch / ColorPicker / Checkbox / Button / Tooltip / Alert。样式使用 `theme.useToken()`，由应用统一 ConfigProvider 继承主题。画布输入继续使用 `antdv-next Input.TextArea`；本次没有新增原生表单控件或整页 contenteditable。明暗弹层、真实 Windows IME、焦点恢复和窄屏行为尚未运行验证。

## 5. Facade、公共偏移与持久化

- [document-structured-blocks.ts](../../packages/office-studio/src/api/document-structured-blocks.ts)：`FDocumentStructuredBlock` / `FDocumentQuote` / `FDocumentCallout` / `FDocumentCode`，提供 ID、查询、范围、逐段落文字、配置/颜色、插入、unwrap/remove/exit。包装只持有文档与 ID，每次重新解析当前模型，拒绝跨文档段落。
- [document-columns.ts](../../packages/office-studio/src/api/document-columns.ts)：`FDocumentColumnGroup` / `FDocumentColumn` / `DocsColumnPositions`，提供组/栏查询、配置、比例/间距、添加删除移动、逐栏文字/格式/段落和公开范围。Facade gap 用 pt，领域/Canvas gap 用 CSS px（乘 4/3）；非法插入 gap 使用默认，非法 setGap 拒绝。Facade deleteColumn 使用 `keepContents:false`，对应公开破坏性删除；UI 可选择保留合并。
- [facade.ts](../../packages/office-studio/src/api/facade.ts) 的 Docs 方法接入这些包装；[index.ts](../../packages/office-studio/src/index.ts) 导出包装及相关配置、查询、范围和结构命令类型。未复制第二套 OfficeStudio / 历史实现。
- [document-body-write.ts](../../packages/office-studio/src/api/document-body-write.ts) / [document-body-codec.ts](../../packages/office-studio/src/api/document-body-codec.ts) 继续输出/读取真实结构标记、`blockRanges`、`columnGroups`、Code/Callout 配置及 officeStudio 来源。Column 范围扫描跳过栏间非打印 rangeStart/rangeEnd，并匹配嵌套 column 标记。

`DocumentBodyParagraphPosition.spans` 的 **source 是公开 stream 偏移，target 是段落局部 UTF-16 偏移**；writer 内部暂存 map 与输出 position 方向不同。公共 range 的文字相交判断使用 source，flat → public 使用 target。注释标签在本地可占多字符、公开占一个 U+FFFC，按 affinity 映射两端，不能混用线性等长映射。方向已在 [document-body-types.ts](../../packages/office-studio/src/api/document-body-types.ts) 注释，调用链位于 [document-public-body.ts](../../packages/office-studio/src/api/document-public-body.ts)。

逐栏 `getText()` 返回用 LF 拼接的普通段落文字；`getTextRange().getText()` 遵守既有公开契约，返回原始 dataStream 切片，包含 CR 段落标记和可能的非打印结构标记。两种 API 不共用文字预期。

Quote 左边框来源由 [document-structured-appearance.ts](../../packages/office-studio/src/features/document/model/document-structured-appearance.ts) 管理。实际导入来源为 `metadata.univerBodyParagraphSource.paragraph.paragraphStyle`，兼容旧直接结构；修复此前漏读 `.paragraph` 的颜色问题。writer 写真实 `paragraphStyle.borderLeft`，同时在导出 paragraph 的 officeStudio.metadata 写 `documentQuoteBorderSource: { version:1, quoteId, generated, original? }`，不修改输入模型。重复保存/重开仍保留引用前的作者边框；unwrap 仅在 quoteId 与当前 generated 边框都匹配时恢复 original 或移除 generated。没有来源、作者独立改过边框或来源不匹配时保留原边框，不能推断所有上游 borderLeft 都由 Quote 产生。

本轮保存/重开只涉及前端 body 和领域快照草稿。DOCX / DOC 二进制来源映射、真实写出、Office 独立重开不在本轮已证实能力内，仍需 DOC-22 适配及样本验收。

## 6. 布局与代码编辑边界

[document-structural-layout.ts](../../packages/office-studio/src/features/document/layout/document-structural-layout.ts) 和 [document-table-measurement.ts](../../packages/office-studio/src/features/document/layout/document-table-measurement.ts) 保留内层 Code 的 containerId / 语言 / 行号标注，外层容器不覆盖它们。`wrap:false` 保持长逻辑行，tab 用测得的空格宽度建立 tab stop，源中仍是一个 UTF-16 tab。渲染使用既有 Code tokenizer 和 Canvas 容器装饰，未把高亮 HTML 变为唯一内容模型。

Column SHRINK 保持声明的 gap 与最小宽度，只压缩超过最小宽度的部分；不足容纳所有 minima 时允许实际溢出，不能无声改 gap。STACK 在不足容纳 minima 时纵向排列并使用整个宿主宽度。算法依据固定 OSS [column.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/block/column.ts) 的 `calculateColumnGroupLayout` / `compressToFit`，本地保持自己的树/节流布局；不声称已经完成性能优化。

## 7. 测试草稿与本轮静态检查

四份新增草稿均**未执行**：

| 文件 | 有意义的预期范围 |
| --- | --- |
| [document-structured-blocks.contract.test.mjs](../../packages/office-studio/tests/document-structured-blocks.contract.test.mjs) | 半开/反向 wrap、颜色与混合 runs、字段/控件锁失败原子性、批注/note 删除与迁移、2–5 栏/ID/嵌套多行插入/字段反转、Tab/Shift+Tab/emoji |
| [document-structured-body-layout.contract.test.mjs](../../packages/office-studio/tests/document-structured-body-layout.contract.test.mjs) | body 往返与输入不变、Quote 边框来源恢复、嵌套栏/SDT 标记、SHRINK/STACK、nowrap/tab、嵌套 Code、公共偏移/单 token note、standalone 会话身份 |
| [document-structured-facade.contract.test.mjs](../../packages/office-studio/tests/document-structured-facade.contract.test.mjs) | 实际 OfficeStudio/FDocument 命令与撤销、快照保存重开、颜色/配置、pt/px、逐栏段落/范围、取消/权限/跨文档错误 |
| [document-structured-workbench.browser.test.mjs](../../packages/office-studio/tests/document-structured-workbench.browser.test.mjs) | 实际 Canvas + ConfigProvider + antdv-next、Tab 历史、旧 revision/重开拒绝、无 ID standalone/相同父快照替换、明暗控件/焦点、合成 composition 保护 |

本轮逐份静态阅读了以上草稿与相关调用链，并修正 TextRange 的 CR dataStream 预期。使用 `git status` / `git diff`、UTF-8 源码阅读、`rg` 和只读 Python 提取 Pro schema 进行检查；这些不是测试通过记录。浏览器草稿的合成 composition 不能作为真实 Windows 输入法验收；尚未补专门的“内容绑定改变缩进计划”失败草稿。

## 8. 剩余缺口

1. **入口与交互**：Quote/Callout 本地组位置与上游 MEDIA 有差异；六种 paragraph 挂载、Pro Canvas PopMenu / mobile / 语言与比例浮层没有全部补齐。当前四动作 Ribbon/右键和属性面板不代表全部 Pro 入口对齐。
2. **代码块**：既有 tokenizer 为 plain、javascript、typescript、json、css、html、python、sql、rust 九种；真实 Code 容器与属性面板保留未知语言配置，旧 paragraph metadata 适配仍将未知语言归一化为 plain。与上游 21 项语言目录及全部语法行为仍有差异。长行的容器内水平滚动、裁剪和光标可达性未完成或未验收。
3. **引用分页**：公开插入规格要求 keepLines / widowControl 和相邻 keepNext，在传统分页能容纳时保持整体；本地尚未实现完整 keep 规则及其历史/文件映射。
4. **多栏**：上游 `column.ts` 默认 minHeight 72、horizontalPadding 8 未对齐；拖拽调宽、完整比例/等宽上下文、嵌套与 traditional/modern 模式限制的逐行为验证仍需补齐。不得把当前可读/可布局推断成所有模式闭环。
5. **保真**：上游导入的 Quote 边框若没有本地来源记录，无法确认引用前边框，保留为未核实作者格式；Office 二进制映射、字段/控件/note/表格混合文件、独立重开、复制粘贴全部结构组合还需样本。
6. **验证**：前端类型、构建、四份草稿、真实浏览器/IME、主题/焦点、WebGPU/Canvas2D、Office 文件和大样本性能均未运行。首屏、滚动、编辑、缩放、主线程长任务、Worker/WASM 复制与内存无本轮实测，继续未验收。

中央功能清单由主任务维护；本证据记录真实前端改动和待完成项，不将开发进度改写为产品验收。
