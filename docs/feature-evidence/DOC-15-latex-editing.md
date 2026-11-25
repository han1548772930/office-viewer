# DOC-15 LaTeX 编辑、公开范围与快照资源

记录日期：2026-11-26。状态：**实现与测试草稿已落盘，未验收**。本批按用户“前端优先，先实现、后统一测试”的要求推进；没有执行测试、类型检查、构建、dev、浏览器、Cargo/Rust/WASM 或性能基准。本记录不将完整 TeX、Office 数学公式转换或全部上游上下文入口标为完成。

后续 [段落 / 浮动 / 移动 popup 批](DOC-15-latex-popmenus-mobile.md)已落盘五个有效挂载、移动编辑 / 删除、public / local popup operation 与运行时 owner，新增 42 条声明 / 80 个计划用例，均未运行；本记录的 43 条仍按前批口径保留。

本记录保留前批公式编辑阶段的实现与缺口。后续 [LaTeX 模板与上下文 Ribbon](DOC-15-latex-templates-ribbon.md) 已落盘 11 类 / 71 个模板、上下文分类及其 11 个根入口，并补矩阵、重音和极限等目录所需解析 / 布局分支；下文“模板尚未实现”等描述属于前批时点。后续批新增 47 条静态测试声明，计划展开 117 个用例，均未运行，与本记录的 43 条草稿分开统计；完整 TeX / OMML / Worker、真实 UI 与性能缺口继续保留，状态仍为未验收。

## 1. 固定依据与来源

OSS 基线为 `E:\rust_test\univer` 的 v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本次只读复核 HEAD 与该提交一致。LaTeX Pro 的契约依据为本项目缓存的 v1.6.0 发布类型和 ES 代码，未安装商业 SDK；闭源排版算法不视为已知实现。

| 来源 | 本批核实的契约 |
| --- | --- |
| [公开 LaTeX 规格](https://docs.univer.ai/zh-CN/guides/docs/features/latex) | 功能目标；页面内容变化、完整演示和 Office 样本仍需统一验收时复核 |
| [common/type.d.ts](../../.research.local/univer-1.6.0/packages/docs-latex/lib/types/common/type.d.ts)、[common/const.d.ts](../../.research.local/univer-1.6.0/packages/docs-latex/lib/types/common/const.d.ts) | 公开 `kind` 和 `displayMode` 均只允许 `inline`；资源名 `DOC_LATEX_PLUGIN`，内容为 `{ formulas: { [rangeId]: config } }` |
| [命令声明](../../.research.local/univer-1.6.0/packages/docs-latex/lib/types/commands/latex.command.d.ts)、[ES 注册与动作实现](../../.research.local/univer-1.6.0/packages/docs-latex/lib/es/index.js) | 插入、更新、删除、转文本四命令；真实 source 保存在 `dataStream`，`CUSTOM5` / `wholeEntity=true`，原始 `startIndex/endIndex` 两端 inclusive |
| [Facade 类型](../../.research.local/univer-1.6.0/packages/docs-latex/lib/types/facade/types.d.ts)、[FDocument](../../.research.local/univer-1.6.0/packages/docs-latex/lib/types/facade/f-document.d.ts)、[FDocumentLatex](../../.research.local/univer-1.6.0/packages/docs-latex/lib/types/facade/f-document-latex.d.ts) | 查询、描述、更新、删除和转源文本；公开范围 `endOffset` exclusive，正文 `segmentId=''`，页眉页脚带 segment ID |
| [段落扩展](../../.research.local/univer-1.6.0/packages/docs-latex/lib/types/facade/f-document-paragraph.d.ts)、[范围扩展](../../.research.local/univer-1.6.0/packages/docs-latex/lib/types/facade/f-document-text-range.d.ts) | `appendLatex` / `replaceWithLatex` 共用文档范围及命令契约 |
| [UI schema](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/types/menu/schema.d.ts)、[UI ES 注册](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/es/index.js) | Insert → Media，order 8，grid row 2 / column 7；右键 order 13；另有段落 Layout 与公式上下文 Ribbon |
| [模板目录](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/types/common/formula-templates.d.ts)、[上下文 Ribbon 控制器](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/types/controllers/docs-latex-contextual-ribbon.controller.d.ts) | 前批时点尚未完整实现；后续 11 类 / 71 模板及上下文入口的逐项来源与实现见 [模板批证据](DOC-15-latex-templates-ribbon.md#1-固定依据与逐入口清点) |

11 类模板为 fractions、scripts、radicals、integrals、large-operators、brackets、functions、accents、limits-logarithms、operators、matrices。模板类别数不是独立功能完成数。

## 2. 本批入口与统一命令

本批接线的本地入口分类：常驻插入入口 1 项；公式右键菜单 3 项；Canvas 双击编辑 1 条交互路径；公开内容命令别名 4 项。classic / grid 属于同一个常驻入口的呈现变体，不重复计数。此编辑批时点，段落 Layout、公式上下文 Ribbon、模板目录、溢出和移动端入口尚未完整逐项核实及实现；后续上下文与模板增量见 [模板批入口清点](DOC-15-latex-templates-ribbon.md#1-固定依据与逐入口清点)。本批数字不能充当 DOC-15 全量入口计数。

| 动作 | 本地 canonical command | v1.6.0 公开别名 | 结果与失败边界 |
| --- | --- | --- | --- |
| 插入公式 | `doc.command.insert-latex-formula` | `docs-latex.command.insert-formula` | 插入一个公式 atom；可替换合法文本范围；继承字符样式及 pending marks；成功产生一次内容历史 |
| 更新公式 | `doc.command.update-latex-formula` | `docs-latex.command.update-formula` | 保留稳定 range / run ID、替换原 source；错误、锁或取消不提交部分内容 |
| 删除公式 | `doc.command.remove-latex-formula` | `docs-latex.command.remove-formula` | 删除目标 atom，保存时删除对应 modeled 配置 |
| 转源文本 | `doc.command.replace-latex-formula-with-text` | `docs-latex.command.replace-formula-with-text` | 用原始 source 或调用者指定文本替换 atom，重映射相关锚点 |
| 打开编辑器 | `doc.operation.open-latex` | 本地 UI operation | 打开 Ant 弹窗，不独立产生内容历史 |

Canonical 命令通过现有文档命令、候选模型与内容历史执行；Facade、右键和弹窗复用这些命令。权限判断涵盖目标范围、内容控件锁和字段锁。菜单禁用状态与实际执行前校验共用服务；禁用按钮本身不视为功能闭环证明。

## 3. 权威模型、坐标与保存

领域模型的公式为一个 U+FFFC 文本 atom，run metadata 保存 version 1、原始 `latex` 与本地 display mode。原始 source 保留首尾空白、CR/LF/tab；不以预览文本或布局数据作为持久模型。

公开 body 输出将 atom 展开为真实 source，建立 source UTF-16 范围与本地 atom 偏移映射；`CUSTOM5` 自定义范围用 inclusive `startIndex/endIndex`，Facade 用 exclusive `endOffset`。公式 source 内部不能作为插入或替换边界；source 中结构控制字符、坏 range ID、非 wholeEntity、重叠范围或配置冲突按整体输入错误处理。未支持的数学语法保留原 source，以可选择、可修复的错误公式进入编辑器。

更新保留导入身份；实际复制/粘贴使用现有 copier 分配独立身份。表格、合并续格、多栏和容器内公式按真实 text target 遍历；正文、页眉页脚与 note 通过对应 block host 使用同一命令。源文本内部多样式与复杂嵌套组合的精确保真仍未核实。

`DOC_LATEX_PLUGIN` named 资源优先于 `latexFormulas` 兼容输入 aliases。导出依据当前 live ranges 重建配置，删除后不以旧 wrapper 恢复公式；nested wrapper 保留各自未知字段并刷新对应配置，未修改输入 source。删除最后公式时，有未知字段的空 wrapper 保留。多 wrapper、资源冲突和真实保存重开仍待运行草稿验证。

本地 `block` 公式是扩展，不宣称 v1.6.0 有公开 block kind。公开插件配置始终输出 `{ latex, kind: 'inline', displayMode: 'inline' }`，本地 block mode 通过本项目 body 扩展保留。普通文字进入独立公式段后清理陈旧 block marker，HTML 只把真正单独的 block atom 作为整段公式写出，避免吞掉旁边普通文本。

## 4. Canvas、输入与布局

公式两侧单击、右键均选择完整 atom；拖选按方向保留完整 atom。命中校验 layout 的 clip、真实 run ID、target 和 source offset，拒绝陈旧布局误选。双击先发布选区，等待 Vue 更新后捕获新输入 session，同时保留选择前的内容 revision 约束，再请求弹窗编辑。

弹窗提交、菜单执行和焦点恢复均检查 source、host、selection、session generation 与 revision；standalone 模式也校验 source identity / snapshot。命令失败保留 source 草稿和可操作错误；成功后若 observer 已导航到其他位置，保留其新位置并拒绝过期焦点恢复。该路径已接线，尚未执行浏览器验收。

LaTeX 编辑器使用 `antdv-next` 的 Modal、Form / FormItem、Select、Input.TextArea 与 Alert；没有新增原生表单控件或其他组件库。弹层与预览颜色取统一 ConfigProvider 的 token。内容 Canvas 错误框颜色来自 `token.colorError`；Canvas 2D / WebGPU 使用同一 scene primitives，预览面板目前使用 Canvas 2D。

本批没有新增原生输入桥；文档输入继续复用现有 Ant 输入及会话适配。Input.TextArea composition 期间抑制 Escape 和 Ctrl/Cmd+Enter 提交，最终输入后再恢复提交。synthetic composition 草稿仅验证 guard，不能代替真实 Windows IME、焦点和主题验收。

本地数学布局仍是有限手写语法：修正空格后脚本、分组后脚本、text-group 空格、重复脚本报错位置与 code point 读取；错误偏移使用原始 UTF-16。分式、根式、脚本 primitives 按 ascent / descent 包围，混合文本行按共同 baseline 安排 glyph，绘制保留上下标偏移、underline / strike。没有公式的普通行保留原有 baseline 规则。复杂字体度量、完整 TeX 数学排版仍未实现。

## 5. 文件归属及公开接口变化

下列路径均相对 `packages/office-studio/src`；其与前批已有变更在同一工作区保留，没有覆盖或清理 Sheets / Charts 在途文件。

| 层 | 本批主要文件 |
| --- | --- |
| 公式模型与纯算法 | `features/document/latex/latex-formula.ts`，`features/document/model/document-model.ts` |
| 命令与能力 | `features/document/commands/document-latex-commands.ts`、`document-commands.ts`；`api/document-editor-commands.ts` |
| Canvas 请求与命中 | `features/document/controllers/document-latex-canvas.ts`；`features/document/layout/document-latex-hit.ts`、`document-text-layout.ts`、`document-layout-types.ts`；`features/document/rendering/document-scene.ts` |
| 公开范围、Facade 与 segment | `api/document-latex.ts`、`document-latex-commands.ts`、`document-segments.ts`、`document-public-body.ts`、`facade.ts`、`office-studio.ts`；`index.ts` |
| Body codec / 资源读写 | `api/document-body-latex.ts`，body codec/read/write/ranges/source/types 等适配文件；`document-input-common.ts`、`document-input.ts`、`document-note-body.ts`、`document-output.ts`、`document-output-notes.ts`、`document-plugin-resources.ts` |
| UI 与 locale | `components/document/DocumentLatexPanel.vue`、`DocumentCanvasViewport.vue`、`document-latex-menu.ts`；`components/DocumentWorkbench.vue`；`i18n/docs/locales/document-latex.ts`、`i18n/office-i18n.ts` |

新增公开 `FDocumentLatex`、公式参数 / range / info 类型、四个命令别名；FDocument 接入插入、查询和 segment 范围方法，FDocumentParagraph / FDocumentTextRange 接入 `appendLatex` / `replaceWithLatex`。

**共享读取行为变化：** `documentPublicParagraphs().text` 改为对应公开 `dataStream` slice，使段落文本与公开 LaTeX source offsets 一致；调用者不再得到本地 U+FFFC atom 作为公式文本。该变化涉及 `getParagraphInfos` / 段落读取，应在后续共享 Facade 回归中核验。

**Segment 变化：** 页眉页脚 public ID 使用 `writePublicDocumentStories` 的保存 ID 分配作为唯一依据，支持同原始 header ID 的独立 host 及保存生成的 `office:headerId:first:firstPageHeader` alias；public ID 优先，local host ID / alias fallback。此查询路径会编码 stories，尚未测量调用代价，不宣称性能已优化。

## 6. 未运行的测试草稿

2026-11-26 通过只读文本清点确认四份新文件包含 **43 条**顶层 `test(...)` 声明。声明数仅用于可复查计数，不能证明覆盖率或运行通过。

| 草稿 | 条数 | 预期覆盖 |
| --- | ---: | --- |
| [document-latex-editing.contract.test.mjs](../../packages/office-studio/tests/document-latex-editing.contract.test.mjs) | 16 | 原始空白、真实错误偏移、unicode、字符 marks、稳定身份 / 实际复制、精确 atom 选区、容器 / 栏 / 表格续格、锁 / 权限、锚点、block marker、布局几何、clip / stale 请求和菜单 |
| [document-latex-body.contract.test.mjs](../../packages/office-studio/tests/document-latex-body.contract.test.mjs) | 10 | 原始公开 stream / range、换行、错误 source、坏边界 / 结构标记、本地 block 扩展、删除不复活、复制 ID、named 优先及 nested wrapper 未知字段 |
| [document-latex-facade.contract.test.mjs](../../packages/office-studio/tests/document-latex-facade.contract.test.mjs) | 9 | 真命令 / 内容历史、公开范围、参数 guards、保存重开、pending marks、取消 / 锁、observer 导航、shared header aliases 与 note body |
| [document-latex-workbench.browser.test.mjs](../../packages/office-studio/tests/document-latex-workbench.browser.test.mjs) | 8 | 真实 Canvas 单击 / 双击 / 右键、弹窗失败保留、stale selection、observer、synthetic composition / 主题、错误修复 / 锁、辅助 hosts、standalone 快照替换 |

浏览器夹具为 [document-latex.html](../../packages/office-studio/tests/fixtures/document-latex.html)，挂载真实 DocumentWorkbench 与 Ant ConfigProvider，含 invalid / locked / auxiliary / standalone 场景；本批没有启动夹具或浏览器。

旧 `document-latex.contract.test.mjs` 的 6 条与旧 `document-latex.integration.contract.test.mjs` 的 1 条保持原有草稿，不计入本批新增 43 条，也未在本批执行。验收时必须运行并修正草稿假设，尤其辅助 host 激活、shared header aliases、弹层焦点、source 内部格式与资源多 wrapper 组合。

## 7. 验证结果与剩余目标

已执行的检查仅为只读源码 / diff 人工查看、固定 HEAD 复核、草稿声明计数和本记录链接 / 文件存在性核对。没有 TypeScript / Vue 解析器、测试 runner、构建器、浏览器或 Rust/WASM 执行结果；所有行为、UI、Office 与性能验收仍未通过验证。

| 剩余目标 | 当前缺口与后续边界 |
| --- | --- |
| 全部上下文及模板入口 | 公式 Ribbon 与 11 类模板见 [模板批](DOC-15-latex-templates-ribbon.md)，段落 Layout、quickBottom / insert、浮动及移动 popup 的五挂载 / 双动作见 [第三批](DOC-15-latex-popmenus-mobile.md)；入口、状态、溢出及持久化仍未运行验收 |
| 完整数学排版 | 后续批已补矩阵 / accents / limits 等目录所需分支，见 [算法与绘制边界](DOC-15-latex-templates-ribbon.md#4-算法绘制与平台边界)；完整 TeX 语法、数学字体 / shaping、样式与宏环境仍有缺口，保留公开源与可编辑模型 |
| Office 互操作 | 原生 DOCX / OMML 读取写出与真实 Office 样本保真尚缺；扩展本项目 parser / writer，不依赖 Univer Server / Exchange |
| 跨功能保真 | source 内部多样式、复杂对象组合、批注 / 字段 / 内容控件、跨容器复制与资源冲突要统一验证并补实现 |
| Worker / Rust / WASM | 数学布局仍在主线程 TypeScript；数学计算域的版本化输入 / primitive 输出、取消、字体版本、buffer ownership / lease 和复制计量未实现 |
| 真实 UI 与输入法 | 明暗主题、Modal 关闭及 observer 焦点、真实 Windows IME、两种 Canvas 后端及辅助 host 切换未验收 |
| 大样本性能 | 首屏、滚动 / 编辑 / 缩放、数学布局、主线程长任务、Worker 传输与 JS/WASM/GPU 内存均未测量 |

前端公式编辑不依赖新增服务端。本项目已有 Office 解析 / 写出链路的 OMML 能力需求及协作 / 版本资源同步仍应由中央 [后端需求](../BACKEND_REQUIREMENTS.md)保留，并按用户要求后置；此记录未改动中央台账或后端实现。

## 8. 中央清单同步条目

主任务同步 DOC-15 到本记录，保留“未验收”；入口台账记录四 canonical / public 命令、三个本地右键动作及双击路径，明确公开 inline / 本地 block 差异和未完成上下文模板。COMMON-03 / COMMON-05 与 DOC-01 / DOC-03 / DOC-12 的共享证据应注明公开段落 source 文本、segment alias 与 named plugin 资源读取行为变化。43 个未运行草稿、完整 TeX / OMML / Worker / UI / 性能缺口分别保留，不能合并成“完整公式已支持”。

以上是前批同步口径；最新中央同步还需关联 [模板批记录](DOC-15-latex-templates-ribbon.md#8-剩余工作与中央同步)，记录常规 Ribbon 根新增 0、上下文分类 1 / 根 11 / 目录项 71，以及该批 47 条声明 / 117 个计划用例。入口与草稿数量均不作为功能验收证明。
