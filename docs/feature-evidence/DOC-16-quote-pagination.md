# DOC-16 Quote 与段落分页：第四批前端实现记录

记录日期：2026-11-25。状态：**已落盘、未验收，前端开发中**。这是 [结构块初批](DOC-13-14-16-17-structured-blocks.md)、[第二批 Canvas 交互](DOC-13-14-16-17-canvas-interactions.md) 和 [第三批 Code / PopMenu](DOC-13-14-16-17-code-popmenus.md) 的补充。按当前用户安排先实现、后统一验证，后端暂缓；本批没有运行测试、类型检查、构建、浏览器/dev、Cargo/Rust/WASM 或性能基准。

本批实现四项段落分页属性、Quote 自动生成规则及作者格式恢复、共享正文/notes 分页路径、Ant 段落设置控件和 Facade 命令接线。以下说明真实代码范围及草稿，不能视为全部 Docs / Pro、Office 文件或性能验收完成。

## 1. 固定上游依据与入口口径

参考 Univer v1.6.0，OSS 提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，仓库 `E:\rust_test\univer`；Pro 类型和可读 ES 位于本项目 `.research.local/univer-1.6.0/packages`。缓存只用于研究，不作为生产商业 SDK 依赖。本批未新增在线演示或 Microsoft Office 文件验证。

| 来源 | 可观察契约 | 本地处理与差异 |
| --- | --- | --- |
| Pro [FDocument Quote 插入声明](../../.research.local/univer-1.6.0/packages/docs-quote/lib/types/facade/f-document.d.ts)，71–90 行 | Quote 每段写 `keepLines=TRUE`、`widowControl=TRUE`，直接相邻段落写 `keepNext=TRUE`，末段 FALSE；能容纳时传统模式保持整体，超高可拆，Modern / Unspecified 只保留属性 | 插入、包裹和转换进入 Quote 自动生成；显式 break、table 和嵌套对象截断直接文本邻接；超高按实际行高拆分 |
| Pro [docs-quote ES](../../.research.local/univer-1.6.0/packages/docs-quote/lib/es/index.js)，`b()` / `x()`、插入 `z`、退出 `B`、解包 `V`、样式 `H` / `S()` | 三项由插入 mutation 合并；Quote style 公开项为 lineColor / textColor；退出路径清理 keep 样式 | 保留既有颜色命令；本地用明确来源恢复引用前格式，保留作者覆盖，避免删除不能确认由 Quote 生成的属性；这属于更保守的保真差异 |
| OSS [IParagraphStyle](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/types/interfaces/i-document-data.ts#L1310)，1310–1357 行 | `keepLines`、`keepNext`、`pageBreakBefore`、`widowControl`；undefined 继承、FALSE 覆盖继承；传统模式应用 | 领域使用可选 boolean；公共 body 接收 boolean / BooleanNumber，作者写出为 0 / 1；四项均纳入 patch 和默认格式校验 |
| OSS [DocumentFlavor](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/types/interfaces/i-document-data.ts) | 0 Unspecified、1 Traditional、2 Modern、3 DrawingML | 公共来源及新切换使用 1；旧本地无公共来源的 0 保留传统兼容，导出公共快照转为 1；公共 0 保留不应用 |
| OSS [linebreaking.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/block/paragraph/linebreaking.ts) 与 [quote-block-layout.spec.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/block/paragraph/__tests__/quote-block-layout.spec.ts) | keepLines / keepNext / widow 优先关系、32 段扫描预算、显式 break 边界、可容纳和超高样本 | 本地树/行测量/section flow 独立实现相关规则，保留扫描预算；未宣称复刻上游全部兼容优先级 |
| Pro [Quote UI schema](../../.research.local/univer-1.6.0/packages/docs-quote-ui/lib/es/index.js)，初批 [根及六种挂载清点](DOC-13-14-16-17-structured-blocks.md#2-固定版-pro-入口清点) | 常规 INSERT → MEDIA，order 6，grid row 1 / column 6；六种 paragraph 挂载及 desktop/mobile PopMenu | 沿用前三批入口和命令；本批没有新增 Quote 分页专属 Ribbon 根 |

本批的四个 Checkbox 属于已有通用段落设置面板的**作者便利扩展**，不是上游新增的四个 Quote 根入口。常驻、扩展、溢出、上下文入口和 classic/grid 位置仍采用前三批清点口径；不以控件数量证明功能完成。已有 Quote / Callout 的常规分组布局差异仍保留。

## 2. 模块归属、模型和公开 body

按 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，责任归 `features/document`：Vue 面板只发 patch；命令和服务管理作者内容；领域保存可序列化属性；layout 输出派生坐标。App.vue、store、解析器未新增业务。本批分页计算仍是现有主线程 TypeScript 路径；没有新增 Worker/WASM 控制或内存协议，也没有将其称为大样本性能方案。后续迁移仍受 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract) 约束。

| 文件 | 本批作用 |
| --- | --- |
| [model.ts](../../packages/office-studio/src/domain/model.ts) 的 Docs 段落定义 | 新增四个可选 boolean；不写页号或 layout 坐标 |
| [document-paragraph-pagination.ts](../../packages/office-studio/src/features/document/model/document-paragraph-pagination.ts) | 字段目录、默认值解析、传统模式判断、Quote 来源生成/作者脱离/恢复/重算 |
| [document-default-style.ts](../../packages/office-studio/src/features/document/model/document-default-style.ts)、[document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts) | 默认格式与 patch 接受四项；非法数字、字符串、null 等领域值拒绝；同值 API 写入仍记录作者格式 |
| [document-body-style.ts](../../packages/office-studio/src/api/document-body-style.ts)、[document-body-paragraph.ts](../../packages/office-studio/src/api/document-body-paragraph.ts) | 公共字段读写、直接/继承/解析基线来源；未修改的继承值不 flatten 成直接格式，作者同值写入可写出 |
| [document-body-source.ts](../../packages/office-studio/src/api/document-body-source.ts) | 将分页解析基线列为派生来源，避免冒充作者 Office 字段 |
| [document-input.ts](../../packages/office-studio/src/api/document-input.ts)、[document-output.ts](../../packages/office-studio/src/api/document-output.ts) | 公共模式来源标记及旧本地 0 → 公共 Traditional 1 兼容 |
| [index.ts](../../packages/office-studio/src/index.ts) 的 Docs 导出区域 | 导出分页字段、解析/模式 helper 及类型 |

`documentParagraphStyleToBody(style, source?, resolution?)` 增加可选第三参数。保留未知 raw 字段及未改的直接/继承表达；没有来源基线时，已知作者分页属性写 BooleanNumber。导入命名样式先解析当前值并保存 baseline；默认/命名样式后续 live 更新的完整继承闭环仍待补和验收。

新持久 metadata：

- 文档 `documentFlavorContract: 'univer'` 区分公共模式契约。
- Quote `documentQuotePaginationManaged: 1` 标记本地受管理的规则。
- 段落 `documentQuotePaginationSource` 保存 version、quoteId、每项 generated / original / originalDirect / originalBaseline / originalAuthored、overrides。
- 段落 `univerParagraphPaginationResolution` 保存 version、inherited、baseline、可选 authored，属于解析来源信息。

生成规则只归属于相应 Quote。Enter / 合并 / 插入 / 删除后在同一候选中重算相邻与末段；显式作者操作，包括同值写入，使该项脱离后续自动重写。解包、移出受管理 Quote、转换离开 Quote 时，只恢复仍与 generated 匹配且来源归属可确认的属性；独立作者改动和没有可信来源的导入 keep 属性保留。body 通过 officeStudio.metadata 传递来源，使多次保存/重开后的恢复有信息依据。协作/后端保留这些扩展字段的需求由主任务维护，本批未实现原生 DOCX 来源映射。

## 3. 命令、Facade、内容锁与历史

[document-structure-commands.ts](../../packages/office-studio/src/features/document/commands/document-structure-commands.ts) 在 `wrap-structured-block`、`insert-structured-block`、`convert-structured-block` 进入 Quote 时生成规则；离开 Quote / unwrap 清理来源并恢复作者值。新建普通段落不从默认格式复制四项成直接格式，而由解析 helper 继续继承。

[document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) 在存在受管理 Quote 时，对候选执行编辑 → 分页重算 → 内容锁/变更权限检查 → 原子发布。段落格式命令先准备全部 style 与独立 metadata；任一目标失败时不提交部分格式。没有另建内容历史源，继续使用 OfficeStudio 按文档隔离的 undo / redo / savepoint。

[document-text-formatting.ts](../../packages/office-studio/src/features/document/services/document-text-formatting.ts) 对分页属性发生变化的整段检查前/后模型的 content control 锁，覆盖“插入点在锁外，但重算会改锁定段”的情况。现有锁定字段对段落格式的语义仍待确认，不能将此描述为已覆盖全部字段锁。

[document-facade.ts](../../packages/office-studio/src/features/document/services/document-facade.ts) 使用共用 paragraph patch 校验和作者标记。[facade.ts](../../packages/office-studio/src/api/facade.ts) 新增本地 `FDocument.formatParagraph(paragraphId, patch, segmentId='')`，按稳定段落 ID 和 note host 调用真实注册 `doc.command.format-paragraph`；`FDocumentParagraph.setParagraphStyle()` 改走同一路径。

可观察接口差异：

- paragraph API 事件从旧未注册的 `doc.command.set-paragraph-style` 变为已有共享 `doc.command.format-paragraph`。
- `{ field: undefined }` 继续是 no-op，不作为清除继承格式；四项真正设置值使用 boolean。
- `getParagraphStyle()` 保持原领域 style 形状，不无条件补四个 false；UI/布局通过 helper 解析默认值。
- 多段格式现在检查整段 content control 锁；不以局部选区绕开段落属性影响范围。

## 4. 共享分页与装饰边界

新增 [document-pagination-rules.ts](../../packages/office-studio/src/features/document/layout/document-pagination-rules.ts) 和 [document-pagination-layout.ts](../../packages/office-studio/src/features/document/layout/document-pagination-layout.ts)，接入 [document-section-flow.ts](../../packages/office-studio/src/features/document/layout/document-section-flow.ts)、[document-paragraph-layout.ts](../../packages/office-studio/src/features/document/layout/document-paragraph-layout.ts)、[document-block-layout.ts](../../packages/office-studio/src/features/document/layout/document-block-layout.ts)。正文和 notes 使用同一 flow / preflight / line layout。

已写规则：

1. 传统模式 `pageBreakBefore` 在已有实际内容时转下一物理页；当前已是空页边界时复用，不制造二次空页。显式 page/column/section break 仍保持自身语义。
2. 可容纳的 keepLines 段或 keepNext 相邻链整体移到可用页/栏；扫描最多 32 段，显式 break、非文本宿主、后段 pageBreakBefore 截断。
3. keepNext 链末段没有 keepLines 时只预留其一行开头；有 widowControl 时预留两行。超高前段的最后可行分片为后段开头预留空间。
4. 超高 keepLines 段在满高栏开始后允许拆分；使用真实可变行高，尽量避免首/末片段仅一行。只能容纳一行或超高单行时允许有限推进。
5. Quote 在消费外间距和捕获装饰起点前，以缩窄后的实际宽度与上下 padding 预检；能容纳的组一起移动。只有 padding 的新栏不误判为已有内容，避免超高 Quote 再跳一空页。
6. 可容纳的组合结合 note reference floor 整体移动；后续行仍遵守 reference floor。section columns 与正文共享规则；header/footer 固定 bounds 不等同正文物理分页。

[document-block-measurement.ts](../../packages/office-studio/src/features/document/layout/document-block-measurement.ts) 复用共用段落行测量，减少预测/实际两套 wrapping；[document-structural-layout.ts](../../packages/office-studio/src/features/document/layout/document-structural-layout.ts) 集中 decorated Insets，并向最后子段传 bottom padding。共用测量保留 UTF-16、不同字体行高、列表 marker、tab/目录页号 tab、Code 配置、note label；发生实际栏宽变化时只重排未放置文本后缀。

object-led 容器的预检只保证首个可见内容能开始；没有宣称完整实现 table-cell keep/widow、所有复杂对象/row 的 Word 分页优先级。本批复用真实段落/对象测量属于代码整合，没有端到端性能测量，不能称为已优化或达标。

## 5. Ant UI、主题、焦点与输入

[DocumentParagraphSettingsPanel.vue](../../packages/office-studio/src/components/document/DocumentParagraphSettingsPanel.vue) 在已有 Form / FormItem 中加入四个 `antdv-next Checkbox`：段中不分页、与下段同页、孤行控制、段前分页。显示 mixed / aria-checked，沿用即时单字段 patch 和 composition 暂存。说明文字明确传统模式应用，其他模式仍保留作者值。没有新增原生表单控件。

[document-paragraph-settings-state.ts](../../packages/office-studio/src/components/document/document-paragraph-settings-state.ts) 补完整 UI 默认值；[document-toolbar-state.ts](../../packages/office-studio/src/components/document/document-toolbar-state.ts) 的正文及 stories mixed 比较先解析四项默认值；[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的面板状态使用同一解析，传统模式切换写公共 1；[document-settings.ts locales](../../packages/office-studio/src/i18n/docs/locales/document-settings.ts) 增加中英文标签。

控件由统一 ConfigProvider 继承 theme；新增说明颜色使用 `theme.useToken()` 的 colorTextSecondary，边界/背景继续使用既有 token。关闭面板沿用当前 Canvas 输入焦点恢复，正文输入保持 `antdv-next Input.TextArea`。明暗主题、disabled/mixed、焦点和合成 composition 有草稿；真实 Windows IME、窄屏、触控及 WebGPU/Canvas2D 实际渲染均未运行验收。

## 6. 草稿数量、静态核对与剩余缺口

本批新增 **1 份契约草稿**、扩展 **2 份既有集成/浏览器草稿**，新增 **37 个 test 声明**，全部未执行。另同步旧 settings 默认值预期和两个旧 Code→Quote 草稿，正确区分保留原 runs / ID 与新增 Quote 分页属性；这些调整不另计新增用例。

| 草稿 | 本批新增 |
| --- | --- |
| [document-quote-pagination.contract.test.mjs](../../packages/office-studio/tests/document-quote-pagination.contract.test.mjs) | 25 个：四项校验、默认/FALSE、命名来源、BooleanNumber、作者恢复、移出/相邻/break、重复 body 往返、模式、真实行高 widow、fit/oversized 段/Quote、hit offsets、keepNext 首尾、物理 break、section columns、真实 reference floor 路径、Enter、多段内容锁 |
| [document-structured-facade.contract.test.mjs](../../packages/office-studio/tests/document-structured-facade.contract.test.mjs) | 6 个：共享 command/history、Enter/merge、公共快照恢复、同值作者格式、锁/非法值/cancel、Quote→Code undo |
| [document-structured-workbench.browser.test.mjs](../../packages/office-studio/tests/document-structured-workbench.browser.test.mjs) | 6 个：真实工作台 fit/Canvas 命中、四个 Ant 控件和 mixed/中英/主题/history/focus、合成 composition、公共模式/body 重开恢复、超高装饰分片、真实 Canvas Enter/undo/redo/重开 |

浏览器草稿仍使用该文件的内联 fixture，新增 `?pagination=fitting|oversized`、`flavor` 和 `locale` 变体，没有另建不存在的 HTML fixture。fixture 对外保存/重开允许选择当前段落，避免分页场景强制选择 Code。全部草稿数量只描述可审阅代码，不是测试通过或行为覆盖完成记录。

静态工作仅为 UTF-8 阅读、rg、git status/diff、声明与本地链接/锚点核对；修正旧转换草稿预期和共享引用。未运行任何类型、测试、构建或浏览器命令。

仍需统一验证或继续实现：

1. 所有新增/修改源文件的类型、前端与契约/集成/浏览器运行验证；各控件主题、焦点恢复、真实 IME；两种渲染后端的实际 frame、hit 和编辑闭环。
2. DOCX 真二进制四项 `<w:keepLines>` / `<w:keepNext>` / `<w:widowControl>` / `<w:pageBreakBefore>` 读写、Quote 来源保真及 Word 独立重开样本。后端按用户安排暂缓，不能用 body 草稿代替 Office 文件证据。
3. table-cell keep/widow、复杂 row/嵌套对象、Word compatibility 11–14 与 15+、嵌套多栏 notes、极端 padding/超高单行、连续节不同宽度和全部分页优先级样本。
4. 默认/命名样式 live 更新继承完整闭环；全项目 notes/outline/toc 既有旧 0 判断尚未统一，本批只处理分页 helper、公共输入/输出及模式切换。
5. header/footer 固定 bounds 的保真边界；字段 locked 对整段格式的语义；所有权限、粘贴/复制、撤销组合、协作来源和历史冲突。
6. Worker/Rust/WASM 分页计算迁移与取消/内存协议、真实普通/大样本首屏/滚动/编辑/缩放/任务指标。本批没有设备、浏览器、帧时间、长任务、复制字节或内存实测，继续**未验收**。

中央功能清单和后端需求由主任务维护，本记录不缩减 Docs / Pro 完整范围。
