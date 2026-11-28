# DOC-04 目录插入命令与对话框目标

日期：2026-11-28 / Asia/Shanghai。状态：**开发中、未验收**。本记录补充[既有目录证据](DOC-04-table-of-contents.md)，不代替完整 Sheets / Docs / Pro 清单。

## 固定源码与入口计数

固定来源：`E:\rust_test\univer`，Univer v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

| 来源 | 本批核实的可观察契约 |
| --- | --- |
| [docs-toc/src/commands/table-of-contents.command.ts](../../../univer/packages/docs-toc/src/commands/table-of-contents.command.ts) | 固定插入 ID `doc.command.insert-table-of-contents`；公开扁平参数；正文折叠选区、避开已有 TOC；级别截断并限制 1–9；插入后更新目录，后续最多三次更新共用 historyGroup。 |
| [docs-toc-ui/src/commands/table-of-contents-dialog.operation.ts](../../../univer/packages/docs-toc-ui/src/commands/table-of-contents-dialog.operation.ts) | automatic 传本地化 title，custom 不传 title；成功才关闭设置对话框。 |
| [docs-toc-ui/src/menu/menu.ts](../../../univer/packages/docs-toc-ui/src/menu/menu.ts) 与 [menu/schema.ts](../../../univer/packages/docs-toc-ui/src/menu/schema.ts) | Insert / CELL、order 7、grid row 1 / col 1 / showLabel，automatic / custom 两项；传统正文、选区和编辑权限决定状态。 |
| [docs-toc-ui/src/controllers/table-of-contents-ribbon.controller.ts](../../../univer/packages/docs-toc-ui/src/controllers/table-of-contents-ribbon.controller.ts) | contextual 目录管理页及 update/delete；原有坐标和显隐规则见既有证据。 |

本批入口根增量 **0**，新增内容命令 ID **1**。既有分类不变：常驻核心根 **0**、插件扩展根 **1**、上下文根 **2**、专用溢出根 **0**；插入子菜单叶 **2**、Layout 复用呈现 **2**、更新模式 **2**。命令、叶项、对话框控件和呈现副本分别计数，数量不代表验收。

## 已落盘的实现候选

- [document-toc-command.ts](../../packages/office-studio/src/features/document/toc/document-toc-command.ts) 新增同源插入状态与 detached plan。验证传统正文、真实折叠 UTF-16 光标、顶层段落、已有 TOC、锁定字段、目标 edit 和被复制标题 view 权限；校验类型、有限级别和正整数逻辑页码。标签按 own-property 读取，避免 `toString` 标题 ID 读取原型成员。无标题或无效上下文拒绝，源模型不变。
- [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) 接固定命令，复用 insertBlocks 与字段 / 内容控件 / 评论迁移；先在副本完成结构和最终 ACL 检查再发布。[content-control-editing.ts](../../packages/office-studio/src/features/document/services/content-control-editing.ts) 将该 ID 分类为文字输入，沿用既有 finalizer 和陈旧 inputGeneration 拒绝。
- 新 [document-toc-commands.ts](../../packages/office-studio/src/api/document-toc-commands.ts) 接公开扁平参数。省略对象按 `{}` 处理，仅从已有、有效且 revision 匹配的正文无 host 会话取得光标；不创建默认输入点。构造同一 DocumentEditorCommandRequest，带 expectedRevision / expectedHistoryGeneration / inputGeneration，复用原子 commit。typed `{command,selection}` 返回原结果对象；flat 返回 changed 布尔；缺失文档或有效正文会话返回 false。
- 共享 [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 由主任务接线：LABELS 循环跳过重复 ID、显式注册 wrapper；BeforeCommandExecute 前捕获文档 owner、history owner/generation、revision、输入 generation、用户及活动文档，拒绝观察者改选、同 ID 重开和切换。先物化首次历史再捕获身份；无文档不提前读取 revision。[index.ts](../../packages/office-studio/src/index.ts) 导出固定 ID、flat DTO 和领域 request。
- [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 automatic/custom 均执行固定命令。automatic 带本地化标题，custom 不附加标题；只有 automatic 提前检查默认 level 3 的标题，因此仅有 level 4+ 时仍能打开 custom。插入 / 更新对话框捕获文档、API、history、revision、输入和 selection；失效目标拒绝，切换或同 ID 更换 Studio 关闭旧草稿。
- [DocumentTableOfContentsDialog.vue](../../packages/office-studio/src/components/document/DocumentTableOfContentsDialog.vue) 保持 antdv-next Form / Modal / Select / Checkbox，补主题化 Alert、disabled、composition 守卫。失败或 BeforeCommand 取消保留设置与错误，父层成功后才关闭；关闭后仅向仍有效目标恢复焦点。

模块归属：领域目录计划 → 内容命令 → 既有 Studio 文档隔离事务 / 历史 → Vue / antdv-next 视图。本批计算仍在主线程 TypeScript，未新增 Worker / WASM 协议或第二套权威引擎；没有性能测量和优化声明。

## 静态检查与后置草稿

已执行只读 TS/JS parser、Vue SFC/template parser、本地 import/export 路径检查：**12 个文件、13 段脚本（含 1 段浏览器夹具）、2 个 Vue 模板、677 条本地引用，错误 0**。包含本批 Docs 文件与已接线共享 API / index。该结果仅为语法与引用检查，不是类型检查、构建或行为验收。

两份 DOC-04 证据的 **41 个本地链接**均存在；本批已跟踪文件的限定 `git diff --check` 和新文件尾部空白 / 末尾换行检查通过。外部网页未做 HTTP 验证。

| 未运行草稿 | 声明 / 计划展开用例 | 覆盖意图 |
| --- | --- | --- |
| [document-toc-insertion-command.contract.test.mjs](../../packages/office-studio/tests/document-toc-insertion-command.contract.test.mjs) | 10 / 27 | 固定命令、光标拆段、TOC/PAGEREF 父关系、标题、级别、缺失标签、原型 ID、错误 DTO、上下文、ACL、字段锁和源模型不变。 |
| [document-toc-insertion-runtime.contract.test.mjs](../../packages/office-studio/tests/document-toc-insertion-runtime.contract.test.mjs) | 9 / 13 | 真实 Studio/FDocument、typed/flat、取消/改选/同 ID 重开/活动切换、版本守卫、文档隔离历史、undo/redo、公开 JSON 保存重开；省略参数与无输入会话。 |
| [document-section-layout.browser.test.mjs](../../packages/office-studio/tests/document-section-layout.browser.test.mjs) 本批增量 | 5 / 6 | classic/grid 真实菜单、取消保留草稿、规范事件、undo/redo/快照；陈旧光标、level 4+ 修正、同 ID Studio 更换、更新拒绝后重试。 |

数量来自 AST 静态声明和字面量 test.each 展开。**运行 0**；未启动测试、类型检查、构建、Rust/WASM、dev、浏览器、Office 或性能基准。真实 IME、焦点、主题、两种 Canvas 后端和页面像素均未验收。DOCX / 独立 Office 重开未验收，JSON 草稿不代替其证据。

## 未完成项与差异

后续[完整布局与历史合并](DOC-04-toc-layout-history-2026-11-28.md)已接插入后的最多三次完整布局重建和同一历史，以及真实 Canvas/font provider；同记录续批接手动全文/只页码最多三次、公开 flat update/delete 与真实公共对话框 operation 宿主。公开 flat 插入/更新改为异步等待真实布局，没有宿主返回 false；delete 无布局依赖，typed 显式模型命令仍保留同步能力。上文 flat 返回布尔及草稿计数属于该轮历史快照，后续 wrapper 返回 Promise 的最终布尔。运行验证仍未执行，复杂原生来源、所有边界及全部公开行为等价仍有缺口。

本地 malformed DTO、源标题 view ACL、原生字段 / SDT 锁和事务保护较固定上游更严格，这是实现差异。复杂原生 TOC、更多来源 / 开关、Roman 等页码标签、旧快照缩进迁移及通用制表位继续按既有清单推进。完整交付范围不变，所有功能保持未验收。
