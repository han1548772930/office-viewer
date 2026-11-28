# DOC-04 现代文档标题折叠

日期：2026-11-29 / Asia/Shanghai。状态：**已落盘、未验收**。复用既有大纲、目录和输入/历史实现，完整 Docs / Sheets / Pro 目标保持不变。

## 固定源码与入口清点

参考 `E:\rust_test\univer`，Univer v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。该能力实际位于 docs / docs-ui / engine-render；没有独立的 `docs-heading-collapse` 包，不能以包名或缺少该目录判断功能不存在。

| 来源 | 已核实契约 |
| --- | --- |
| [set-heading-collapsed.command.ts](../../../univer/packages/docs/src/commands/commands/set-heading-collapsed.command.ts) | `doc.command.set-heading-collapsed`，参数 `{unitId, paragraphId, collapsed}`；只适用 MODERN，真实唯一段落、非空可折叠范围、外层可见与目标权限判断；布尔元数据进入 RichTextEditingMutation 与内容历史。隐藏当前选择时迁回标题段尾。 |
| [heading-sections.ts](../../../univer/packages/engine-render/src/components/docs/view-model/heading-sections.ts) | Heading 1–5，以后一个同级/更高级标题或独立流末尾为边界；表格单元格和现代栏独立；保留嵌套折叠状态。 |
| [heading-sections.spec.ts](../../../univer/packages/engine-render/src/components/docs/view-model/__tests__/heading-sections.spec.ts) | 子状态、独立栏、继承样式/直接正文覆盖、传统模式与非标题 flag 不生效。 |
| [doc-heading-collapse.service.ts](../../../univer/packages/docs-ui/src/services/doc-heading-collapse.service.ts) | 只在现代正文可见且有内容的标题旁挂载控件；隐藏选区归一到外层可见标题，激活后归还编辑器焦点。 |
| [heading-collapse-shape.ts](../../../univer/packages/docs-ui/src/views/heading-collapse-shape.ts) | 上游 Canvas 控件宽 24 / 高 28、位于首行左侧，折叠向右/展开向下，hover 提示延迟 500ms；滚动/缩放/离开重置临时状态。 |
| [break-line.command.ts](../../../univer/packages/docs-ui/src/commands/commands/break-line.command.ts)、[doc-delete.command.ts](../../../univer/packages/docs-ui/src/commands/commands/doc-delete.command.ts) | 折叠标题 Enter 先展开再拆段；边界 Backspace/Delete 优先展开隐藏内容，避免直接删除不可见相邻段落。 |
| [set-heading-collapsed.command.spec.ts](../../../univer/packages/docs/src/commands/commands/__tests__/set-heading-collapsed.command.spec.ts) | 快照、子级状态、undo/redo、文字不变与远端 mutation 场景；本项目无实际协作服务验收，不由这些源测试推断本地多人功能。 |

本批常驻 Ribbon、插件扩展 Ribbon、上下文 Ribbon、溢出根和菜单叶增量均 **0**；新增的是动态正文覆盖控件：每个当前可见且有折叠内容的标题 1 个。页面中实际个数随文档/滚动/折叠变化，不能固定成 Ribbon 数量或功能完成总数。固定命令 1，按钮/鼠标/键盘/API 共用它；19 locale 的 collapse/expand 标签来自固定 docs-ui locale。

## 已落盘的真实链路

- [heading-collapse.ts](../../packages/office-studio/src/features/document/headings/heading-collapse.ts) 以既有 `metadata.headingCollapsed` 为持久标志，读取现代模式和 Heading 1–5。构建独立流边界、唯一身份索引及隐藏内容到外层可见标题的映射；只投影布局需要的结构，不删除作者模型、不清子标题 flag、不复制所有文本/媒体。表格、续格、栏和结构容器保留独立内部流。普通文字或传统模式中的 flag 保留但无折叠效果。
- 同一纯模型提供可见 textTargets、隐藏选区归一、状态/权限/内容锁检查、命令 mutation。折叠选择覆盖的隐藏内容时迁回标题段尾；其他隐藏端点按外层可见标题归一。源正文、对象和子级标志保留，错误在 mutation 前拒绝。
- [document-layout.ts](../../packages/office-studio/src/features/document/layout/document-layout.ts) 在原 `layoutDocumentWithNotes` 前使用结构共享投影，正文/表格/图片/高级块实际不进入隐藏区的布局绘制与命中；[document-note-layout.ts](../../packages/office-studio/src/features/document/layout/document-note-layout.ts) 可接作者模型的编号 marker map，折叠不会让后续可见编号被重新从 1 计数。派生布局不成为第二份持久文档模型。
- [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) 接固定命令以及 Enter/边界删除行为。Enter 的展开和拆段在同一候选事务中完成；单次边界删除只展开而不删除隐藏文字。与固定内部多次 mutation 编排不同，本地保留一次用户操作的原子性和源内容保护。
- 新 [document-heading-commands.ts](../../packages/office-studio/src/api/document-heading-commands.ts) 接固定 flat 参数；无效模式、缺失/重复/隐藏/空标题、权限/锁或幂等请求返回 false。typed 显式请求继续走原文档命令、版本、权限及错误契约。[FDocument.setHeadingCollapsed](../../packages/office-studio/src/api/facade.ts) 复用同一个注册，不能绕过只读。主任务在 [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 的文档方法区接注册/BeforeCommand owner guard 与统一内容历史。
- [document-input-commands.ts](../../packages/office-studio/src/api/document-input-commands.ts) 和 [document-input-store.ts](../../packages/office-studio/src/features/document/stores/document-input-store.ts) 在公开选区/模型重排时归一不可见正文位置。不存在的输入会话不因折叠被制造；独立 story/辅助内容输入保持原作用域。折叠的内容历史通过 [document-history-cache.ts](../../packages/office-studio/src/features/document/services/document-history-cache.ts) 可选 afterPublish 回调恢复 undo 前隐藏选区、redo 后标题选区；默认调用及 TOC retained pair/amend/字段缓存 rebase 责任保持。
- [DocumentHeadingCollapseOverlay.vue](../../packages/office-studio/src/components/document/DocumentHeadingCollapseOverlay.vue) 使用 **antdv-next Button / Tooltip**，沿用 ConfigProvider 主题、禁用/hover/focus。位置复用 Canvas 首行/表格裁剪/缩放/视口，控件基准 24×28，提示延迟 0.5 秒。按钮用 Lucide 方向图标，与上游手绘三角像素不同；这是可替换 UI 的项目技术选择。按下捕获源模型/输入代次，点击前重新核实；组字、失活、拖动离开或模型变化拒绝旧操作。Enter/Space 激活使用同命令。
- [DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 实际装配覆盖层；可见方向/文首文尾和 Ctrl+A 使用可见目标，隐藏文字仍保留在作者模型与保存数据中。[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 在 standalone/公共模型更新和大纲导航时归一隐藏位置；输入/IME/权限与焦点恢复复用现有链路，工具栏根数量不变。
- [locales.ts](../../packages/office-studio/src/features/document/headings/locales.ts) 提供固定 19 locale 的两个标签，仅在 [office-i18n.ts](../../packages/office-studio/src/i18n/office-i18n.ts) 加 feature import / messages spread。公开 snapshot 原有 `document-body-paragraph.ts` 读写 `headingCollapsed`，本批复用该契约；内部保存和公开 JSON 重开保留完整隐藏段落及 flag。

本批计算仍在 TypeScript 的模型/布局域；没有新增 Worker/WASM 协议、网络服务或第二套权威历史。没有实测性能、长文档主线程耗时、内存或两后端证据，不作优化达标声明。

## 未运行草稿与检查

| 文件 | 草稿覆盖意图 |
| --- | --- |
| [document-heading-collapse.contract.test.mjs](../../packages/office-studio/tests/document-heading-collapse.contract.test.mjs) | 父子/同级边界、独立表格/栏、源文字/对象保留、可见 Canvas 行与编号、前后向选择、非现代/非标题/空/重复/隐藏拒绝、ACL/内容锁、Enter/Delete/Backspace、Studio 单历史与 undo/redo 选区、公开快照重开/子标志和 BeforeCommand 取消/改选。10 条声明，包含循环场景，未将循环自动计入固定计划数量。 |
| [document-heading-collapse.browser.test.mjs](../../packages/office-studio/tests/document-heading-collapse.browser.test.mjs) | classic/grid、明暗主题、真实 Ant 折叠按钮/Canvas 行、隐藏光标迁回和 undo 恢复、键盘激活/子级状态/保存重开、组合输入禁用、只读和失活控件。3 条声明 / 4 个字面量 each 展开计划；合成 composition 不等于原生 Windows IME 验收。 |

截至落盘后静态快照：19 个生产/草稿文件、20 段脚本（含浏览器夹具）、3 个 Vue 模板和 999 处本地 import/export 路径，parser 错误 0；两份 mjs `node --check` 通过。后续文档与收尾检查另记。**测试、类型检查、前端构建、dev、浏览器、Office、Rust/WASM 与性能基准均未运行**，全部功能保持未验收。

## 尚未验收或继续缺口

- 真实 Windows IME、字体加载、RTL/缩放/窄视口、表格/续格复杂布局、鼠标捕获/主题焦点、同 ID 远端替换与两 Canvas 后端需统一运行验证。
- 主线静态核实 Rust 中尚无 `headingCollapsed` / `w15:collapsed` 专用 authored flag 映射；原段落 source XML 即使保留，也不证明用户新折叠/展开标志已写入原生 DOCX。原生映射是明确开发缺口，独立 Office 重开也未验收，不能用公共 JSON 保留 flag 宣称 Office 往返完整；该状态和真实服务端 changeset 并发仍按完整交付范围推进。
- 大纲列出完整标题；点击隐藏标题归一到外层可见标题，不自动清除上层或子级作者 flag。超链接、查找结果、复杂对象选择、完整范围复制/剪切及字段缓存修改后的历史光标映射还需补更深层交叉场景，不能用基础隐藏选区样例证明全部覆盖。
- 快照历史的选择回调复用原文档权限/模型替换顺序；on-open 字段缓存发生在同一选中段落并改变字符长度时的准确历史光标映射仍待专项核实。原 TOC 快照 pair/amend 行为保持，不把静态引用检查当作该回归已通过。
