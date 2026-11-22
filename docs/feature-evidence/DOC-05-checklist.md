# DOC-05 列表、编号与 DOCX 保留字段证据

更新日期：2026-11-22（Asia/Shanghai，按本批主任务客户端日期记录）。状态：**未验收**。本记录覆盖清单列表和本轮编号模型、命令、DOCX 读写保真切片；DOC-05 的完整增强列表及 Pro 范围继续保留。下面的代码行为与新增测试均不能代替浏览器、文件样本及独立 Office 重开验收。

## 固定源码与入口契约

参考版本为 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地只读检出位于 `E:\rust_test\univer`。本轮按该版本复读列表命令、菜单状态和预设。全工作台的 Sheets、Docs、插件扩展、溢出和上下文入口清点继续以[顶部功能区清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)与[工具栏对照](../UNIVER_TOOLBAR_PARITY.md)为依据；本切片不重新定义全局入口计数。

| 常驻列表根入口 | 固定 schema 的组内顺序 / grid 位置 | 上游命令与菜单状态 | 本项目入口 |
| --- | --- | --- | --- |
| 有序列表 | `RibbonStartGroup.LAYOUT`，order 7；row 2 / column 1 | `doc.command.order-list`；按钮带选择器；显隐按 Docs 上下文，禁用按有效文档范围，激活态来自选区列表类型 | `doc-numbered` 主按钮切换列表；预设选择执行 `doc.command.change-list-type` |
| 无序列表 | 同组，order 8；row 1 / column 2 | `doc.command.bullet-list`；按钮带选择器；同样由文档范围与当前列表类型驱动状态 | `doc-bullet` 主按钮切换列表；预设选择执行 `doc.command.change-list-type` |
| 清单 | 同组，order 9；row 2 / column 2 | `doc.command.check-list`；普通按钮；`CHECK_LIST` / `CHECK_LIST_CHECKED` 驱动激活态 | `doc-checklist` 执行 `doc.command.set-block-type`，Canvas 勾选执行 `doc.command.toggle-checklist-item` |

这张表只清点三个既有列表根入口。classic 按 schema 的 order 保持该顺序；grid 是同一入口集合的位置变体。固定预设提供六个常规有序、六个常规无序选项，另有快速输入预设和清单状态；预设子项、快速输入、Canvas 勾选及本地编号动作分别统计，不能与 Ribbon 根入口混算。本项目有序列表弹层内的“重新编号”和“继续前一编号”是两个本地动作，起始值是表单输入，不作为新增上游根入口。

可复查的固定来源：

- [menu/schema.ts][source-schema]：LAYOUT 中列表根入口的分组、顺序和 grid 配置。
- [menu/menu.ts][source-menu]：`OrderListMenuItemFactory`、`BulletListMenuItemFactory`、`CheckListMenuItemFactory` 与 `listValueFactory$`；多段落类型一致时才给出当前列表值。
- [commands/commands/list.command.ts][source-command]：`OrderListCommand` / `BulletListCommand` 的默认切换经 `ListOperationCommand`，指定预设经 `ChangeListTypeCommand`；`ChangeListNestingLevelCommand` 使用 `type: 1 | -1` 修改嵌套级别。命令作用于文档 segment 的实际段落范围，最终派发 `RichTextEditingMutation`，不能只改变按钮激活态。
- [preset-list-type.ts][source-presets]：常规与快速预设、嵌套定义及 U+2610 / U+2611 清单符号。
- [doc-checklist.render-controller.ts][source-checklist]：命中清单符号后执行 `ToggleCheckListCommand`。
- [官方增强列表规格](https://docs.univer.ai/zh-CN/guides/docs/features/list)：动态规格索引，完整 Pro 目标仍须逐行为核验。本轮没有重新观测公开演示，也没有据开源入口推断闭源内部实现。

上游 `ChangeListNestingLevelCommand` 参数名为 `type`，本地命令使用 `direction`。本地清单切换命令 ID 与上游不同，领域语义均为段落列表成员关系与勾选状态。重启、继续和下面的 SEQ 保真方案是本项目按 Office 编号语义实现的契约；上述开源命令文件不能证明 Pro 内部采用相同模型或字段。

## 本地模型、命令与历史

相关文件：

- [document-numbering.ts](../../packages/office-studio/src/features/document/model/document-numbering.ts)：编号定义、实例 ID、连续计数分组和列表层级。
- [document-lists.ts](../../packages/office-studio/src/features/document/model/document-lists.ts)：有效级别、显示标签和跨正文 / 表格段落的计数。
- [document-office-numbering-marker.ts](../../packages/office-studio/src/features/document/model/document-office-numbering-marker.ts)：导入字段的编号与可见前缀基线契约。
- [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts)：列表切换、嵌套、继续 / 重启及保留前缀替换事务。
- [documentNormalizer.ts](../../packages/office-studio/src/parser/documentNormalizer.ts)：解析 DTO 到可编辑编号模型和保留字段基线的归一化。

编号以 `metadata.numberingDefinition` 的 `abstract` / `instance` / `levels` 保存原始定义、实例覆盖和有效级别，成员关系使用 `list.id` / `metadata.numberingListId`。起始值零、`levelStartOverrides`、`levelOverrides`、`lvlRestart`、`isLgl`、缩进、字符格式、补零宽度和未识别字段保留在可序列化模型中。原始编号定义、当前可编辑模型、布局标签和 DOCX 字段缓存各有明确边界。

标题深度使用标题段落的 `level`；其编号深度单独使用 `metadata.numberingLevel` 或解析的 `numbering.ilvl`。调整编号和重启 / 继续不把标题深度当成列表深度。普通列表段落仍以 `level` 表达列表深度。

| 本地命令 / 接口 | 当前实现行为与边界 |
| --- | --- |
| `doc.command.change-list-type` / `doc.command.set-block-type` | 对真实选区段落应用预设或切换类型，按模型维护成员关系；明确替换导入编号时先处理缓存前缀。非法预设拒绝，不能变成空操作。 |
| `doc.command.change-list-nesting-level` | 校验 `direction: -1 | 1` 和有效级别；编号深度与标题深度独立，相关缩进随有效级别调整。 |
| `doc.command.restart-list-numbering` | 起始值须为 0–2147483647 的安全整数，缺省为 1。对当前段落及后续同实例段落建立新的实例 / 起始覆盖；正文、标题及表格内同实例段落使用统一遍历，连续计数实例重启后脱离原计数分组。 |
| `doc.command.continue-list-numbering` | 查找前方兼容编号序列，将当前段落及后续同实例段落接入该序列；保留各段落自己的编号级别。不存在兼容序列时拒绝；已经属于同一序列时不新增内容变更。 |
| `canRestartDocumentNumbering` / `canContinueDocumentNumbering` | 由当前编号级别、前方兼容序列和全部受影响段落的编辑权限计算能力；执行命令再次验证。 |

导入缓存前缀的基线是复制的快照，包含编号契约和可见前缀的可写出字符属性；忽略 run ID 与仅改变运行切分的差异。明确替换编号仅删除仍匹配的标签文本；已编辑成普通正文的标签保留为文本。替换会同步变换正向或反向选区、跨标签 / 正文批注与正文内容控件范围。多段落替换先计算全部内容控件变换并验证锁定，后续段落拒绝时不得先删掉前面段落的标签。

前缀提取拒绝在 UTF-16 代理对中间截断，也拒绝消费前缀中的脚注、尾注或 LaTeX 对象。前缀后的空脚注引用保留原 run ID 与 note 关联。当前文档命令与应用命令总线共用内容历史；新增浏览器用例通过公开 `FDocument` 路径验证命令事件、撤销 / 重做和快照重开，未建立独立组件历史栈。

## 清单切片的既有实现

- `DocumentBlock.type === 'checklist'` 使用 `checklist.listId` / `checklist.checked` 保存成员关系和勾选值。它是列表语义，与内容控件复选框及内联字符分开。
- `doc.command.set-block-type` 对选中段落创建、切换或移除清单；相邻项目复用 list ID。Enter 生成未勾选的新项目，空项目 Enter 退出清单。
- Canvas 布局产出清单命中矩形，scene 绘制状态；点击通过 `doc.command.toggle-checklist-item` 和文档历史执行。
- 导入把 Duke 的 `details.listLabel` 中 `☐` / `☑` 映射为清单状态，并保留源元数据。写出用原生 Word numbering 定义、对应符号和 `Segoe UI Symbol`。
- 内部 HTML 快照按 list ID 分组，通过 `data-office-list-type`、`data-office-list-id` 和 `data-office-checked` 保存清单状态。JSON / HTML 快照不是 DOCX 文件互操作验收。

## DOCX 编号与保留字段读写

相关文件：[document_numbering_extension.rs](../../crates/office-parser/src/document_numbering_extension.rs)、[document_format_sidecar.rs](../../crates/office-parser/src/document_format_sidecar.rs)、[document_comment_sidecar.rs](../../crates/office-parser/src/document_comment_sidecar.rs)、[docx_numbering.rs](../../crates/office-writer/src/docx_numbering.rs)、[docx_numbering_sequence.rs](../../crates/office-writer/src/docx_numbering_sequence.rs)、[docx_numbering_marker.rs](../../crates/office-writer/src/docx_numbering_marker.rs)、[docx_source_xml.rs](../../crates/office-writer/src/docx_source_xml.rs)、[docx_source_styles.rs](../../crates/office-writer/src/docx_source_styles.rs) 和 [docx.rs](../../crates/office-writer/src/docx.rs)。解析仍经过 Duke 0.6.0，写出经过项目自身 `office-writer` / OPC 包链路，不依赖 Univer Server、Exchange 或私有协议。

普通有序 / 无序 / 清单及适合原生表达的多级编号写入 `word/numbering.xml`，段落使用 `w:numPr`；标题大纲级别与编号级别分别写入。十进制补零以原生 `w:numFmt/@format` 表达；缺少必需的自定义格式令牌返回 `INVALID_MODEL`，不静默生成另一个编号格式。

跨层连续计数使用 `SEQ` 域和可见缓存，`office_list_v1_` 书签标识本项目编号区间，并通过 `customXml/office-numbering.xml` 中 `urn:office-studio:numbering:1` / version 1 映射保存实例、字段契约及原生定义对照。该映射服务于本项目语义恢复；SEQ 域不是 Word 原生列表实例。独立 Office 的字段更新、复制、重排与列表操作效果仍待样本验收，不能把本项目重导入成功等同于原生列表完全等价。

解析恢复必须同时匹配编号映射、实际书签唯一性、字段指令 / 字面文本、相应字符格式以及适用的原生编号定义。映射缺失、被外部修改或契约不一致时，保留当前可见标签及其各自运行样式，并把原编号区间写入 `details.officeNumberingMarker`。归一化在 `metadata.officeNumberingMarker` 中保存原 XML、命名空间、源编号和当前编号 / 前缀基线；保留源数据不等于已经完整解释所有 Office 行为。

标签被批注覆盖时不能直接变成没有文本锚点的语义标签，因此保留字段和可见文本。批注上下文追踪覆盖前方正文 / 表格段落；前一段开始、下一段标签结束的批注不会因只检查当前段落而丢失标签。只覆盖正文的批注仍允许恢复语义编号，同时把 source XML 投影和可见 children 对齐，避免隐藏字段缓存造成批注偏移漂移。

写出保留字段时再次校验基线和 XML：

- 正文编辑与运行重新切分不使未改动的前缀失效；前缀字符格式、源编号或 XML 不再满足契约时，不重放旧字段，保留当前文本并给出 `DOCX_NUMBERING_MARKER_NOT_REPLAYED`。
- 明确更换编号时，移除未编辑的缓存标签并按当前编号写出，给出 `DOCX_NUMBERING_MARKER_REPLACED`。这条替换分支不把已经改写的标签文本再次删除。
- 有效保留 XML 中的旧批注标记先去除，再按当前文档批注模型插入当前 ID 和范围；跨字段可见结果的锚点按 UTF-16 偏移切分，不允许分割代理对。
- 重放和复制后的书签 ID / 名称重新分配，新增编号书签同时避让。SEQ 标识符只从以 SEQ 开始的实际指令提取，按大小写无关匹配并识别引号；IF 的普通字符串中出现 SEQ 不预占计数器名。
- XML 校验覆盖闭合、命名空间、实体、重复属性和控制字符，并限制大小 / 深度 / 元素数。尚未保留的关系或扩展导致有诊断的文本保留；原 XML 仍在快照模型中。
- 对保留字段实际引用的字符样式，导入原定义及 `basedOn` / `link` / `next` 依赖闭包；原 ID 与生成样式冲突时分配 `OfficeSourceStyleN`，同步改写字段引用与样式依赖。缺失、无效或循环定义、未保留的关系资源、无法安全恢复的上下文给出 `DOCX_NUMBERING_MARKER_STYLE_CONTEXT`，保留文本与模型中的原资源；该诊断不代表全部字符样式已恢复。

解析 DTO 的 [DocumentPayload](../../crates/office-parser/src/model.rs) 新增可选 `officeSource`，version 1，包含 `styles` / `theme` 的 `partName` 和原始 XML。资源优先按主文档内部关系定位，缺少关系时按内容类型和标准路径回退，单个 XML 限制为 16 MiB；原字符串由归一化后的 `metadata.source` 保留。源 part 路径不会直接成为新包的写出路径，最终资源通过项目 OPC 包接口写入。

写出把源 `docDefaults`、默认或显式段落的样式 ancestry、字符样式 ancestry 和原直接 `rPr` 合成为字段运行的直接属性；按样式层处理切换属性，按单个属性合并 `rFonts` / `lang`，直接格式覆盖继承结果。导入依赖定义与生成样式并存，导入定义的默认标记移除，必要的 `DefaultParagraphFont` 引用补定义。缺省字号与固定 Duke 0.6.0 `src/parse/styles.rs::{with_word_application_defaults,parse_style_sheet_selected}` 的规则保持一致：未声明运行默认属性使用 12 pt 和可用的 minor theme 字体，声明了运行默认属性但缺少字号时使用 10 pt。该适配的 Office 独立重开效果仍待验证。

字体或颜色依赖主题时，检查主题根节点及所需字体 / 颜色方案，写入 `word/theme/theme1.xml`、内容类型和文档关系。XML 安全校验由字段与资源共用的 `docx_source_xml.rs` 执行，覆盖 UTF-8 XML 1.0、命名空间、实体、重复属性、控制字符和大小 / 深度 / 元素数限制。样式数限制为 4096，继承及依赖深度限制为 64。主题方案中更深层定义的完整性、原字体表 / 嵌入字体、表格条件样式和任意扩展资源仍需核实和适配。

同段落的正文批注与内容控件使用共享事件游标，在递归 SDT 范围和运行切片中写出，不再因两者同时存在而跳过批注。点批注按开始 → 结束 → 引用排列；零文本脚注 / 尾注及同位置批注事件归属其后区间，段尾对象和事件只由最终区间写出。附着在非空运行起点的引用在后续文本切片中不重复写出。无效、同级相交或不支持的控件包装仍有诊断，正文和当前批注继续写出。这些行为已加入真实 DOCX 用例，尚未运行。

明确替换完全被批注覆盖的缓存标签后，批注范围变为显式点锚点，消息保留。Duke 没有收集到文本目标、但正文 `comment.reference` 带有相等的显式起止位置时，解析适配生成点目标；TypeScript 归一化保留零长度锚点，并递归解析嵌套表格段落及 `sourceRefs`。非安全整数、倒序、起点越界或代理对切分的目标不生成有效锚点，消息仍保留，锚点状态报告为 orphaned。[thread-comments.ts](../../packages/office-studio/src/features/document/comments/thread-comments.ts) 也使用递归段落查找与范围校验。

Duke 的 typed SDT 聚合为单个 FormField，而 review walker 计数内部运行，可能产生批注偏移差异。`document_comment_sidecar.rs` 从段落 XML 投影可见文字和批注边界，跟踪跨正文 / 嵌套表格的活动范围；只有投影文字与生产 paragraph children 完全相等时才替换受影响正文目标。缺少 / 无效 XML、投影不一致或同时包含正文非文本对象目标时，保留既有目标并追加诊断。该模块修复的是批注位置；完整内容控件导入、模型编辑和再次写出仍需另行补齐，不能把 typed SDT 的 FormField 或 richText 的扁平文字当作完整控件模型。

## 本轮 UI 检查边界

[DocumentListPicker.vue](../../packages/office-studio/src/components/document/DocumentListPicker.vue) 的预设和动作使用 `antdv-next` Button；起始编号使用 Form / Form.Item / InputNumber，接受零并限制有效整数。输入捕获保护 composition、`isComposing` 和 keyCode 229 下的 Enter / Escape，组合输入期间不提交重启。

组件样式使用现有 Ribbon 主题变量，由共享 ConfigProvider / token 上下文承载。本轮没有引入原生表单控件，也未替换 Canvas 编辑内核。明暗主题、弹层尺寸、禁用 / 悬浮 / 焦点状态、失焦恢复以及真实中文 IME 尚未在浏览器执行验证；自动触发组合事件的用例不能代替系统输入法验证。

## 验证记录

### 历史清单验证

以下来自本文件此前的清单切片记录，未在 2026-11-22 本轮重跑，不代表修改后的编号链路已通过：

- `document-checklist.contract.test.mjs`、`document-command-boundaries.contract.test.mjs`、`document-toolbar-state.contract.test.mjs` 和 `rendering/document-scene.contract.test.mjs`：历史 26/26 通过。
- `document-toolbar.browser.test.mjs`：历史 6/6 通过，覆盖 Ribbon 创建、Canvas 勾选、撤销 / 重做、Enter 和内部 HTML 重开。此前记录的执行命令为 `node --experimental-strip-types --test packages/office-studio/tests/document-toolbar.browser.test.mjs`。
- `rustfmt --edition 2024 --check crates/office-writer/src/docx.rs`：历史通过。此前 crate 级格式检查在同期的 `src/xlsx_validation.rs` 存在差异，这不是本轮执行结果。
- `cargo test --manifest-path crates/office-writer/Cargo.toml`：历史 38/38 通过，含清单符号 writer 用例。
- `pnpm.cmd --filter @office-studio/office-studio check:types`：历史通过。
- `tests/browser/production-office-io.spec.ts` 的生产 writer → parser DOCX 用例此前未运行。此前 `wasm-pack` 已编译 writer，但环境拒绝下载 / 安装缺少的 wasm-bindgen CLI；尝试项目内临时目录仍失败。该历史环境结果不证明当前 WASM 构建或浏览器 JS glue 状态。

### 2026-11-22 新增 / 更新的待执行验证

本编号子任务只做了源码、接口、夹具和文档复查；**没有执行单元测试、契约测试、浏览器验证、类型检查、前端构建、Cargo、WASM 或格式检查**。最终统一验证由主任务执行，并须以实际结果另补记录。

| 文件 | 待执行场景 |
| --- | --- |
| [document-numbering-model.contract.test.mjs](../../packages/office-studio/tests/document-numbering-model.contract.test.mjs) | 标题 / 编号层级独立、跨标题 / 表格实例重启和继续、零起始及首个嵌套覆盖；基线复制、运行切分、代理对与内联对象；缓存替换后的反向选区、点 / 范围批注、内容控件和锁定事务；正文及嵌套表格点目标归一化、无效目标的消息保留和 orphaned 状态；脚注 ID 及编辑后的标签文本保留。 |
| [document-numbering-ribbon.contract.test.mjs](../../packages/office-studio/tests/document-numbering-ribbon.contract.test.mjs) | 既有命令映射、编号起始参数及能力状态回归；本轮未运行。 |
| [document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) | 编号 / 标题重启、继续和保持各自级别；公开命令事件、历史及 JSON 重开；缓存删除后点批注、正文批注 / 内容控件范围、焦点恢复及 Canvas 宿主。修正夹具 `[0] 原始正文` 偏移 2–8 对应的文本为 `] 原始正文`。该套夹具不执行 DOCX Worker 往返。 |
| [document_format_roundtrip_tests.rs](../../crates/office-parser/src/document_format_roundtrip_tests.rs) | 真正的 writer / parser DOCX 往返：源编号覆盖、标题 / 表格、补零、自定义格式错误；缺失 / 改动映射、编辑缓存和样式、明确新列表、旧批注 ID 替换、跨段 / 正文 / 点批注；同段嵌套 richText / plainText 控件与内含、跨边界和点批注；typed SDT 跨正文及表格的批注；空脚注 / 尾注在批注及嵌套控件边界、段尾、全空段落和非空引用不重复；字符样式 / 段落继承、ID 冲突及依赖改写、直接格式、缺省字号、非标准样式 / 主题关系路径、主题字体 / 颜色与反复保存；缺失 / 无效 / 循环样式和不安全主题的诊断。 |
| [document_comment_sidecar.rs](../../crates/office-parser/src/document_comment_sidecar.rs) 内测试 | 源 XML 缺少 / 无效、可见投影不一致及混合正文非文本对象目标时，既有批注目标保持完整并报告诊断。 |
| [document-numbering-roundtrip.spec.ts](../../tests/browser/document-numbering-roundtrip.spec.ts) | 生产 WASM 在专用 Worker 生成合成 DOCX；应用的 parser / export Worker、Canvas 正文编辑、两次真实下载和再导入保留缓存字符样式与标签批注；Ribbon 更换列表、撤销 / 重做、下载和两次再导入保留标签删除后的点批注。独立配置为 [document-numbering-roundtrip.playwright.config.ts](../../tests/browser/document-numbering-roundtrip.playwright.config.ts)，复用生产 preview 和 WASM 资源；暂未运行。 |
| [docx_numbering_marker.rs](../../crates/office-writer/src/docx_numbering_marker.rs) 内六个公开导出测试 | 复制书签和新编号冲突、SEQ 初始指令 / 引号 / 大小写、无效 XML 文本保留、有效实体与字符样式上下文诊断、源编号不匹配及明确替换、字段内代理对批注边界拒绝整个导出。 |

Rust 测试直接调用公开 DOCX 导出并检查实际 OPC XML，回读经生产 parser；辅助模型夹具不替代 TypeScript 归一化测试。新增浏览器套件通过实际应用导入和导出并在另一个 Worker 中重读下载文件，不改写生产适配器。上述测试均需由主任务执行后记录结果，仍需 Office 兼容程序独立重开和真实系统 IME 验证。

## 剩余范围与验收缺口

- 新增测试、类型 / 前端 / Rust / WASM 检查均未在本轮执行；普通和长文档、多级及表格跨段交互仍需浏览器 / 文件样本验收。
- 点批注、同段批注 / 内容控件和字符样式资源已补实现与待执行用例，尚无本轮运行结果；生产浏览器用例覆盖点 / 范围批注和缓存样式，完整内容控件导入模型与 UI 控件再次导出仍未闭环。缓存前缀与控件重叠、非文本 SDT、任意其他 story 的批注仍需展开并验收。
- 源字符样式按字段引用及依赖恢复；更深层主题定义、表格条件样式、原字体表 / 嵌入字体、祖先专属命名空间和任意替代前缀的 Duke 行为、未支持关系 / 扩展的回写仍需补适配。循环 / 缺失定义、linked style 原 numPr 和不支持的 rStyle 委托继续保留诊断，不能据资源 XML 留在模型中认定完整保真。
- 所有增强列表 / Pro 子行为、自定义样式、图片编号、重排、富列表剪贴板、跨 segment 操作及 Office 字段更新继续按完整目标展开和验收；本轮不把未核实项改成产品范围限制。
- 未进行 Word / LibreOffice 独立重开、公共演示对照或真实大样本性能 / 内存测量；没有首屏、滚动、编辑、缩放、Worker / WASM 传输与内存的前后对照，不能宣称已优化或达标。

DOC-05 保持 **未验收**；功能主清单及全局入口对照由主任务统一更新，本证据文件不以测试数量证明功能覆盖。

[source-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts
[source-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts
[source-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/list.command.ts
[source-presets]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/docs/data-model/preset-list-type.ts
[source-checklist]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/render-controllers/doc-checklist.render-controller.ts
