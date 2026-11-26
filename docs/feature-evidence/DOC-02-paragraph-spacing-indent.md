# DOC-02 段落缩进、行距与段前后间距

记录日期：2026-11-27。状态：**已落盘、未验收**。关联 DOC-02 格式、DOC-05 列表缩进和 DOC-03 页眉页脚宿主。本批只实现、编写契约草稿及静态核对；未运行测试、类型检查、构建、浏览器/dev、Rust/WASM 或性能基准。

## 固定来源与入口清点

参考仓库 `E:\rust_test\univer`，Univer v1.6.0 固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。本批没有新增线上演示或 Office 文件证据。

| 固定来源 | 已读契约 |
| --- | --- |
| [menu/schema.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts) | START → LAYOUT 只有对齐、三种列表、页眉页脚五根；本地直接缩进、行距和段距不是固定顶部同形入口。段落设置挂 MAIN_AREA → LAYOUT order 0、PARAGRAPH → LAYOUT order 0、EMPTY_PARAGRAPH → LAYOUT order 0；后者组 order 1。 |
| [menu/context-menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/context-menu.ts)、[段落设置 operation](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/operations/doc-paragraph-setting-panel.operation.ts) | `sidebar.operation.doc-paragraph-setting-panel` 打开面板；Docs 上下文可见，Edit 权限决定禁用。 |
| [menu/menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts) | `AlignMenuItemFactory.value$` 与四个对齐按钮 `activated$` 共用 `getParagraphStyleAtCursor`：取 active range 或首个 range，按 startOffset、segmentId 找段，未求全选区共同值。 |
| [menu/paragraph-menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/paragraph-menu.ts)、[ParagraphMenu.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/ParagraphMenu.tsx) | `doc.menu.paragraph-t.align` 下的 indent 组 order 1；增加/减少顺序 0/1，ID 为 `doc.menu.paragraph-t.indent.increase/decrease`。执行转发段落设置命令，将当前段 indentFirstLine 加减 20，最小 0。 |
| [paragraph-setting/Setting.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/paragraph-setting/Setting.tsx)、[hook/utils.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/paragraph-setting/hook/utils.ts) | 面板按 segmentId 取正文或 header/footer，再取范围中的段落；对齐、缩进、段距、行距字段均读 paragraph[0]。行距 metrics 也取首段。 |
| [paragraph-setting/line-spacing.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/paragraph-setting/line-spacing.ts)、[ParagraphWordWrapSetting.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/paragraph-setting/ParagraphWordWrapSetting.tsx) | auto 倍率 1–5；atLeast/exact 数值 1–100；换规则优先实际 glyph/line metrics。wordWrap 为独立字段。 |
| [doc-paragraph-setting.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/doc-paragraph-setting.command.ts) | `doc-paragraph-setting.command` 支持 segmentId，保留 paragraph 后 patch，进入 RichTextEditingMutation；historyAction 为 FormatParagraph。snapToGrid 在参数中，本批读取的两边面板都没有对应可见控件。 |
| [core const.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/types/const/const.ts) | 默认 lineSpacing 1.5、spaceAbove 0、spaceBelow 12。 |

以下是 [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 中本批四个根的清点，归本地 Home → `doc-layout-extensions`（段落工具，group order 4）。classic 顺序由 [ribbon-schema.ts](../../packages/office-studio/src/platform/ribbon/ribbon-schema.ts) 的稳定 order 排序决定；四项默认 order 0，更多样式 order 10。

| 标签 / ID | grid | classic 顺序 | 命令与参数 | 状态 |
| --- | --- | --- | --- | --- |
| 减少缩进 / doc-indent-less | r1c1 | 1 | `doc.command.change-paragraph-indent`，direction -1 | 编辑上下文及实际变化段落 capability 决定禁用 |
| 增加缩进 / doc-indent-more | r2c1 | 2 | 同命令，direction 1 | 同上 |
| 行距 / doc-line-height | r1c2 | 3 | `doc.command.format-paragraph`，spacingRule auto、lineHeight | 六个倍率：1、1.15、1.5、2、2.5、3；exact/atLeast 显示规则与数值 |
| 段落间距 / doc-paragraph-spacing | r2c2 | 4 | 同命令，spaceBefore 或 spaceAfter | 段前/段后两个子菜单；各七项 0、6、8、12、18、24、36，显示当前值选中态 |

分类计数：本批 **0 个固定常驻顶部根、4 个本地扩展顶部根**，加上上批更多样式，该组共 5 根；行距 6 个选项，段距 2 个子菜单/14 个叶，不能把选项当根。没有新增独立溢出功能；宽度不足时仍由共享 Ribbon 对同一根做溢出。固定相关上下文为段落设置 1 个唯一根/3 个挂载，以及对齐缩进子菜单 1 根，其中本批相关缩进叶 2 个。没有把固定上下文缩进叶算作顶部常驻根。

固定基础段落面板有 10 个逻辑字段、13 个业务交互控件。本地 [DocumentParagraphSettingsPanel.vue](../../packages/office-studio/src/components/document/DocumentParagraphSettingsPanel.vue) 有 14 个逻辑字段、17 个业务交互控件：四个对齐按钮、四个缩进数字、两个段距数字、行距规则 Select、行距数字、wordWrap 和四个分页 Checkbox。关闭按钮另计。四个分页字段为 [DOC-16 已有扩展](DOC-16-quote-pagination.md)，不是新增的四个固定根。

## 模块、命令及本批修正

按 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，能力归 `features/document` 的模型、命令与服务。Vue 只读当前状态和发命令；计算仍是现有同步 TypeScript 小型交互，没有新 Worker/WASM 控制 DTO、缓冲区或内存协议。后续重计算迁移继续遵守 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)。

[document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) 本批新增内部复用 `prepareDocumentParagraphIndent`。按钮 capability 和真实执行都先计算实际变化段落，再检查其编辑权限和整段内容锁；所有检查完成后才移除导入编号缓存和修改缩进。检查复用 [document-text-formatting.ts](../../packages/office-studio/src/features/document/services/document-text-formatting.ts)，没有复制锁规则或引入新命令协议。

正文普通段缩进继续修改 indentFirstLine，步长 20、范围 0–100；带编号段继续修改自己的列表深度，并保留相对编号定义的有符号偏移。已到边界、不发生变化的锁定段不阻止其他段变化。跨段任何实际变化目标被锁时，前段样式、编号深度、导入 marker 缓存、runs 和选区都保持原值。范围继续采用现有半开段落语义，末点为后一段 offset 0 时不修改后一段。

行距和段距沿用真实 `doc.command.format-paragraph`。已有实现先检查每个整段内容锁，准备全部 style/metadata，并通过 [document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts) 应用 patch，随后统一提交；本批没有重造服务或改动该算法。规则切换复用 [document-paragraph-settings-state.ts](../../packages/office-studio/src/components/document/document-paragraph-settings-state.ts)，具体行高仍由现有 Canvas 布局消费。

## 首段状态、宿主与界面

固定顶部对齐读选区起始段；固定段落设置面板读 paragraph[0]。原本地面板基准取 selection.focus，跨段前向选择会显示末段；legacy header/footer 时还会读取休眠正文范围，并因 legacyStoryActive 禁用面板。[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 已统一 `paragraphContext`，从当前编辑 host 的首个选中段读取段落样式、行距和字号；表格选区读取首个实际文字目标。打开 legacy 段落设置时通过已有 `activateStoryBlockEditing` 迁移至 block host。面板和顶部对齐共用该起点，命令仍通过原执行路径；[document-toolbar.browser.test.mjs](../../packages/office-studio/tests/document-toolbar.browser.test.mjs) 与 [document-story-input.browser.test.mjs](../../packages/office-studio/tests/document-story-input.browser.test.mjs) 已补对应草稿，全部未运行。

命令早已支持既有 host。`documentModelForBlockHost` 将 legacy story locks 投影为当前内容区的 content control；本批缩进预备因此复用相同整段锁判断。链接到前节、权限和输入法限制仍由已有宿主/输入路径负责，没有另建 legacy 执行协议。

面板 Form/FormItem、Button、InputNumber、Select、Checkbox、Tooltip/反馈全部来自 `antdv-next`；本批没有新增或改写 UI 控件。共享 ConfigProvider 和主题 token 路径保留，明暗主题、禁用状态、焦点恢复和真实 IME 尚未运行验收。

## 数据、历史与保存差异

[document-editor-commands.ts](../../packages/office-studio/src/api/document-editor-commands.ts) 已注册缩进和段落格式命令；工作台与 [FDocument.formatParagraph](../../packages/office-studio/src/api/facade.ts) 都进入这一路径。公共 API 的 paragraphId/segmentId 转成现有 host/range；[OfficeStudio](../../packages/office-studio/src/api/office-studio.ts) 使用候选文档和按文档隔离的历史。FDocument.undo/redo 复用同一历史，save 输出公开快照并记录 savepoint；本批只追踪源码，不能视为生产历史、保存/重开或 Office 文件验收。

本批记录以下真实实现差异与未核实项：

- 固定上下文缩进取当前段数值并以一个 patch 应用；本地多段从各段自己的数值加减。固定 handler 只截断到最小 0，本地同时遵守设置面板的 100 上限。
- 固定默认段后间距为 12，本地段落设置默认状态为 8；倍率 1.5、段前 0 一致。保留既有本地默认，本批没有擅自修改所有新文档样式。
- 规则转换均优先真实 metrics；无 metrics 时固定源码使用 DEFAULT_STYLES/default spacing，本地使用当前字号 ×4/3，并截断到输入范围，精度也有差异。
- [document-body-style.ts](../../packages/office-studio/src/api/document-body-style.ts) 将公开非 auto lineSpacing 导入时乘 4/3、写出时乘 3/4；本地领域/UI 为 px。固定设置 hooks 直接读写 scalar。该公开 DTO/排版单位差异尚未用浏览器或独立 Office 样本核实，本批不声称完全一致。
- snapToGrid 参数的可见入口、完整继承、signed 导入缩进和低于 1 的导入 auto 倍率，需按来源另行验收；不能仅凭导入值绕过固定面板的现有范围。

## 草稿、静态核对与未验收

[document-paragraph-indent.contract.test.mjs](../../packages/office-studio/tests/document-paragraph-indent.contract.test.mjs) 原有 7 个 test 声明，本批新增 5 个，共 12 个，全部未运行。新增覆盖锁外 caret/部分选区、后段锁的原子拒绝及前段导入编号缓存、缩进边界的锁段不变化、正反向半开范围、真实 story host 的锁/正文与其他段隔离。草稿直接调用生产实现，没有 Mock。

本批限定 diff 检查通过；核对 12 个 test 声明、16 个本地链接和 12 个固定源路径，链接/源文件均存在，新增文档无尾随空白。尚需统一运行契约、跨模块历史/save/reopen、classic/grid 浏览器、主题/焦点/Windows IME、真实 Office 读写样本，以及普通/大样本性能和内存指标。未完成实测继续保留**未验收**，本记录不缩减 Docs 或 Pro 的完整交付范围。
