# Docs 命名段落样式

功能：DOC-02，关联 DOC-04 / DOC-05 / DOC-22。更新日期：2026-11-27。
固定基线：`dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地 `E:\rust_test\univer`。
状态：未验收；本轮实现与测试草稿已落盘，运行验证统一后置。

## 固定源码与入口

| 来源 | 可复查契约 |
| --- | --- |
| `packages/ui/src/views/font-size/interface.ts` 的 `HEADING_LIST` | 依次为 Normal、Title、Subtitle、Heading 1–5，共八项 |
| `packages/docs-ui/src/menu/schema.ts` 的 `SetParagraphNamedStyleCommand.id` | START 字体组，order 5.5；grid row 1 / column 1 / columnSpan 2 / width 80 / minWidth 72 |
| `packages/docs-ui/src/menu/menu.ts` 的 `HeadingSelectorMenuItemFactory`、`getParagraphStyleAtCursor` | 选择器执行命名样式命令；当前值取 active range 的 startOffset 所在段落，缺少命名样式时取 Normal；显隐按文档类型、禁用按可用范围 |
| `packages/docs-ui/src/commands/commands/set-heading.command.ts` 的 `SetParagraphNamedStyleCommand` | 命名样式与 outlineLevel 分开存储，重设段落间距/行距；段落样式 patch 不清除 bullet，编辑进入 RichTextEditingMutation；通过 segmentId 作用于 header/footer |
| `packages/docs-ui/src/commands/commands/list.command.ts` 的 `ListOperationCommand` / `ChangeListTypeCommand`；`packages/core/src/docs/data-model/text-x/build-utils/paragraph.ts` 的 bullet switch/set | 设置、切换及取消列表保留 paragraph / paragraphStyle，沿用独立的 bullet.nestingLevel；标题与命名样式不会被列表操作清除 |
| `packages/core/src/types/const/const.ts` 的 `NAMED_STYLE_MAP`、`NAMED_STYLE_SPACE_MAP` | Title 26 pt 加粗、Subtitle 15 pt 灰色、Heading 1–5 为 20/18/16/14/12 pt 加粗；正文使用全局默认样式 |

本地 `doc-block-style` 是一个常驻根入口，八个选择项均执行 `doc.command.set-paragraph-named-style`，参数为 `namedStyle`。classic 保持 120px 选择器，grid 使用以上布局元数据。标题 6 与引用保留在更多段落样式扩展入口。本轮调整既有入口的状态和行为，不新增根入口；完整分类仍见[顶部清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)。

## 本轮实际实现

- `document-model.ts` 在 `metadata.namedStyleType` 保存稳定样式键，保留段落 ID、文本、运行 ID、缩进、对齐以及无关字符属性。Normal / Title / Subtitle 保持普通段落，Heading 1–5 更新标题级别。
- 应用命名样式时，编号和项目符号保留 listId、preset、自定义 numbering definition，并将列表深度保存在独立的 `metadata.numberingLevel`，避免标题级别覆盖列表层级。旧式列表缺少列表资源时补默认 preset 身份。清单保留 `checklist` 类型、列表层级、listId 和 checked 状态。
- 设置列表、切换预设及取消列表同样保留 `metadata.namedStyleType` 和字符/段落格式。标题首次启用列表从深度 0 开始，即使共享相邻标题列表实例；已有编号标题切换预设沿用独立的 `metadata.numberingLevel`，标题 `level` 保留。取消仅移除列表属性，恢复原有标题类型/级别；普通列表继续以 `level` 表示列表深度，清单沿用现有类型并在取消时恢复命名标题。复用 `blockHasNumbering`、`blockNumberingLevelIndex`、`blockListLevel` 和已有编号复制/清理逻辑，不新增模型或协议。
- `document-commands.ts` 先收集全部可设置样式的目标段落，对每个整段检查编辑权限与内容锁，全部通过后再修改，避免最后一段拒绝时已改动前面的段落。沿用半开文本范围与真实表格结构选区。
- 导出 `documentNamedParagraphStyleState(model, selection, permissionContext?)`，返回 `{ available, reason? }`；菜单状态和命令复用同一个目标收集及整段权限/锁判断。Ribbon context 的对接字段为 `namedStyleDisabled` / `namedStyleDisabledReason`，工作台接线与共享工具区一并完成。
- `document-toolbar-state.ts` 新增 `namedStyle`，按排序后范围的起始段落显示样式；`blockStyle` 的混合态、字符混合态及清单状态保持独立。
- `document-story-formatting.ts` 在现有选区格式结果中返回 `namedStyle`，复用归一化选区和 `storyBuffer.literalOffsets`，读取起始段落的 `metadata.namedStyleType`，缺省取 Normal；`documentStoryToolbarState` 将该值传给 Ribbon。六种 header/footer story 使用同一读取路径，PAGE 缓存仍在输入桥中占一个对象位置。
- `DocumentWorkbench.vue` 对 legacy story 复用 `documentBlockHostSelectionFromLegacyRange`、`documentModelForBlockHost` 和 `documentNamedParagraphStyleState` 判断整段权限/锁；选择器不再因 legacy story 永久禁用，仍受当前可编辑状态、IME composing 和前节链接约束。执行沿用 `execute()` 已有的 `activateStoryBlockEditing()` 与同一命名样式命令。新 story block host 原本已支持该命令；首次 legacy 块编辑沿用现有迁移和兼容投影，不新增模型、命令 ID 或跨边界协议。
- `document-ribbon.ts` 读取 `namedStyle`，将 heading1–5 映射到既有 H1–H5 选项值；有序/无序列表按真实 listPreset 显示激活态，编号标题继续显示编号按钮状态。

样式选择器继续使用共享 `RibbonItem.vue` 中的 `antdv-next Select / Tooltip`，不新增原生表单或 Canvas 输入桥。主题、焦点恢复、输入法与两种渲染后端的实际交互仍待统一验证。

## 已有导入与写出边界

旧记录中的“Ribbon 未接命名样式”和“Title / Subtitle 无导入或写出”已过时：

- `packages/office-studio/src/parser/documentNormalizer.ts` 已按 styleId 识别 Title / Subtitle，存入 `metadata.namedStyleType`；普通标题级别仍由 headingLevel 映射。
- `crates/office-writer/src/docx.rs` 的 `paragraph` 已写出 Title / Subtitle 的 `w:pStyle`，并按 heading block 写出 Heading1–6；列表编号属性独立写出。
- `crates/office-parser/src/parser.rs` 已有 `writer_title_subtitle_and_east_asian_font_reopen_as_independent_properties` 草稿，检查写出重读后的 styleId 与中西文字体槽。本轮未执行该测试。

这些已有适配不代表任意自定义命名样式、继承链或清单标题均能保真往返。

## 草稿、验证与剩余项

本轮脚本声明数为：`document-named-paragraph-style.contract.test.mjs` 16 条、`document-toolbar-state.contract.test.mjs` 15 条、`document-toolbar.browser.test.mjs` 14 条。新增内容覆盖八种样式、列表身份和层级、清单状态、整段锁/权限原子性、半开范围、表格选区、反向混合选区、撤销重做、JSON 重载及 Canvas 清单切换。legacy story 新增三个合同草稿，覆盖六种 story 的反向选区/PAGE 对象偏移和第二段 caret 回读、八种样式的既有 host 命令复用及字段/正文/其他 story/末段保留，以及整段内容锁、段落权限和真实分节链接的原子拒绝。列表反向操作另新增三个合同草稿，覆盖八种命名样式在三种列表按钮下的启用/取消、共享相邻标题列表而新段从深度 0 开始、半开末段保留、十二种预设的独立深度、导入/无命名键标题取消，以及 story host 命名标题/字段缓存/正文/其他 story 保留。

`document-toolbar.browser.test.mjs` 夹具沿用已有的外部 history 数组；它只为组件交互编写草稿，不作为生产统一历史、持久保存或 Office 互操作已通过的证据。所有测试、类型检查、构建、Rust/WASM、浏览器和性能基准本轮均未运行，未启动 dev。

`document-story-input.browser.test.mjs` 补充三个命名样式计划用例：classic/grid 各一个真实 Studio 场景，覆盖选区值回读、命令迁移、正文/字段保留、生产历史撤销重做、`FDocument.save()` 后新 Studio 重开；另一个 synthetic composition 场景检查选择器禁用且不增加历史。草稿均未运行，真实 Windows IME、列表按钮与命名样式交互、生产保存闭环和 Office 独立重开仍未验收；合同中的 JSON 重载不作为保存或 Office 独立重开的验收证据。

本轮已阅读限定实现与测试 diff、复核真实 Table 构造函数和 Select/选项文案；legacy 补充复核真实分节构造、story 块迁移、字段投影和既有工作台执行路径。上述七个实现/测试文件及本记录已有 `git diff --check` 退出码 0；legacy 补充的 story formatting、toolbar state、named-style 合同和本记录四个文件的限定 `git diff --check` 及本地文档链接检查均通过。测试声明数仅作草稿统计。

列表反向操作补充阅读固定版 bullet switch/set、既有 numbering/list 模型与 scoped host 调用；命令、合同草稿及本记录三个文件的限定 `git diff --check` 与本地文档链接检查通过。运行验证仍统一后置。

历史旧文档记录过命名样式合同 3/3、Studio 类型检查及限定 diff 检查通过。测试脚本及实现已变化，该记录仅保留历史来源，不作为本轮结果。

剩余项：

- 当前运行模型未区分继承样式与显式字符格式，应用命名样式仍会写入字号/加粗，并按样式设置颜色；Normal 使用本地固定正文默认值。默认样式继承与真实字号/间距单位效果继续核实。
- 清单标题保留清单类型，其大纲/目录语义及原生 DOCX 写出尚未闭环；自定义编号标题的完整控件状态继续核实。
- 列表与命名样式的双向命令修正已落盘；编号/标题混合选区、完整增强列表模式、真实工具栏交互、生产历史和 Office 编号往返仍需运行验收。
- 上游 header/footer segment 对应的本地 story 命名样式状态与 legacy 接线已落盘；真实工作台的选区迁移、IME/焦点恢复、撤销重做、保存重开和 Office 文件往返仍未验收。
- 自定义 OOXML 样式、源样式继承、独立 Office 重开、生产历史与保存闭环、明暗主题、Windows IME、Canvas2D/WebGPU 和大样本性能保持未验收。
