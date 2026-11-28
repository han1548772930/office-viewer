# DOC-04 目录插入后的完整布局与历史合并

日期：2026-11-28 / Asia/Shanghai。状态：**已落盘、未验收**。续接线程 `01a220b5-b017-7723-b522-04979a516e34`，沿用[插入命令与入口清点](DOC-04-toc-insertion-command-dialog-2026-11-28.md)。全部功能及 Pro 范围保持不变。

## 固定来源与数量

只读参考仓库 `E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

| 来源 | 可复查契约 |
| --- | --- |
| [docs-toc/src/commands/table-of-contents.command.ts](../../../univer/packages/docs-toc/src/commands/table-of-contents.command.ts) | `InsertTableOfContentsCommand` 先等待完整排版，插入后调用更新；`updateTableOfContents` 最多三次，每次修改后等待新布局，插入与后续更新共用 historyGroupId。`waitForCompleteLayout` / `waitForNextCompleteLayout` 的布局等待超时为 30 秒。 |
| [docs-toc-ui/src/commands/table-of-contents-dialog.operation.ts](../../../univer/packages/docs-toc-ui/src/commands/table-of-contents-dialog.operation.ts) | automatic/custom 复用插入命令；既有分组、显隐、布局和成功关闭语义不变。 |
| [core/src/services/undoredo/undoredo.service.ts](../../../univer/packages/core/src/services/undoredo/undoredo.service.ts) | 内容历史按 unit 隔离；本项目借用已有 exact-top amendment，保留一次用户撤销的可观察效果。 |

本批新增常驻、插件扩展、上下文、溢出根和子菜单均为 **0**，未新增公式名称。沿用既有目录插件根 1、上下文根 2、automatic/custom 子项 2；API、布局事件和内部任务状态不计作 UI 功能入口。Docs 常规根 54 / 三分类数量与上批相同；计数不作为行为验收。

## 生产链路与边界

- [DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 在现有 `layoutDocument` 权威路径发布 `layout-state`：`source / layoutGeneration / fontVersion / complete / layout / engine`。开始计算、组字、字体加载、失活和卸载发布 pending；只有当前模型的非预览、非 inspection、字体就绪结果可发布 complete。当前实际 engine 为 **main-thread**，没有新建布局 Worker，也不把同步排版描述为已完成 Worker 优化。
- 字体监听绑定实际宿主的 `FontFaceSet`，覆盖 loading/loadingdone/loadingerror、首次 ready 和卸载释放。字体代次变化重建测宽缓存，避免已缓存的后备字体宽度继续用于新字体。 [document-text-layout.ts](../../packages/office-studio/src/features/document/layout/document-text-layout.ts) 的 `createTextMeasurer` 增加可选宿主 `Document`，默认调用兼容；Canvas 的测宽与字体状态使用同一宿主。
- [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 将实际 Canvas publication 与当前 FDocument、模型 revision 和 history generation 绑定。只在源模型相同或首次归一化数据相同时映射权威模型；组字、失活、模型/宿主切换及同 ID reopen 不接受旧结果。未就绪时禁用插入并不提供旧页码。临时布局、字体代次、provider 和监听均不写入文档快照。
- 新 [document-toc-layout.ts](../../packages/office-studio/src/api/document-toc-layout.ts) 提供 `bindDocumentTocLayout`、版本化 read/subscribe、当前 owner 判断、等待与释放。等待同时检查 source 身份、unit、revision、history generation、布局代次和完整性；同步完成早于订阅的情况重新读取。等待有 30 秒上界，取消、替换、关闭和卸载释放监听/timer。
- [document-toc-commands.ts](../../packages/office-studio/src/api/document-toc-commands.ts) 保留 typed UI 的同步结果，并调度插入后的重排；公开 flat 命令通过异步 `executeCommand` 等待真实完整布局、用实际页码插入，再等待后续重排。没有布局宿主返回 false，不把调用者的数字当作渲染器；`executeCommandSync` 对 flat 参数拒绝。typed/headless 显式文档命令仍可构造模型，其后续排版结果为 unavailable。
- 新 [document-toc-pagination.ts](../../packages/office-studio/src/api/document-toc-pagination.ts) 从插入前后稳定 TOC 身份确定目标，最多执行三次完整更新。每次读取新布局后经原命令、ACL、字段/内容控件锁和候选事务发布，保持插入后光标。更晚的内容、选区、活动文档、用户、输入代次、保存、撤销重做、历史/模型替换或 owner 失效停止旧任务。BeforeCommand hooks 改变布局/字体/请求参数也不能授权旧候选。
- [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 将内部不可序列化 amendment capability 接入原文档事务；外部 JSON 不能指定任意历史条目。 [document-history-cache.ts](../../packages/office-studio/src/features/document/services/document-history-cache.ts) 合并时保留最初 undo 与同一 snapshot pair，并在成功发布后更新 after，因此后续 on-open 字段缓存 rebase 仍覆盖最终 redo。失败、取消和无变化不增加空历史。
- [FDocument.bindTocLayout](../../packages/office-studio/src/api/facade.ts) 公开真实 renderer provider 绑定；[index.ts](../../packages/office-studio/src/index.ts) 导出 provider/publication/result 类型及 `waitForDocumentTocPagination`。任务结果区分 settled/limit/canceled/unavailable/rejected；临时状态不改持久模型。

全部可替换控件沿用 `antdv-next` 与现有 ConfigProvider，本批未新增原生表单控件。新增 Canvas layout publication、字体监听与 DOM 测宽属于平台/排版边界，纯 TOC 模型与事务不依赖 DOM。未改变 Worker/WASM 协议和权威计算域，不声明端到端性能优化。

## 未运行草稿

| 文件 | 覆盖意图 |
| --- | --- |
| [document-toc-insertion-runtime.contract.test.mjs](../../packages/office-studio/tests/document-toc-insertion-runtime.contract.test.mjs) | 真实 Studio/FDocument：插入后完整布局与页码、光标、单次 undo/redo、保存重开；flat async/无宿主拒绝、pending/过期布局、三次非收敛上界、保存/改选/撤销/owner/同 ID reopen 取消、hook 内字体失效与监听释放。旧 flat 草稿同步断言已修正为 await/rejects。 |
| [document-field-cache-history.contract.test.mjs](../../packages/office-studio/tests/document-field-cache-history.contract.test.mjs) | amendment 保留原 undo，字段 cache rebase 使用最终 redo；未发布或发布抛错不污染 retained pair。 |
| [document-toc-pagination.browser.test.mjs](../../packages/office-studio/tests/document-toc-pagination.browser.test.mjs) | classic/grid、明暗主题、真实 Ant 目录菜单 → Canvas 布局 → 页码 → 单历史 undo/redo；FontFaceSet pending/complete/代次和不写内容历史。合成字体事件不等于真实网络字体、宿主浏览器和 Windows IME 验收。 |

当前阶段仅执行只读语法、引用、来源路径、文档链接与 diff 检查；草稿均未运行。测试、类型检查、前端构建、Rust/WASM、dev、浏览器、Office 往返及性能基准全部后置。最终静态检查范围与结果交接主任务记录。

本批静态快照：14 份生产/草稿文件、15 段脚本（含 1 段浏览器夹具）、2 份 Vue 模板和 841 处本地 import/export 路径通过只读 parser/引用检查；三份 mjs 的 `node --check` 通过。三份 DOC-04 证据的 60 处本地链接及限定 `git diff --check` 通过；最初 parser 因根目录没有直接导出的 `@vue/compiler-sfc` 退出，改用已安装 Vue 的 `vue/compiler-sfc` 后完成检查，没有安装新依赖。测试草稿声明/字面量 each 展开为 runtime **15/22**、字段缓存历史 **6/6**、浏览器 **2/3**；这三个数字分别属于各文件，历史字段缓存文件的原 4 条未计为本批新增。

## 继续缺口

- 手动 update / flat update-delete 已在下述续批接线；真实运行验收仍未执行，不由插入路径或静态检查推断整个 DOC-04 已完成。
- 复杂原生 TOC 来源/开关、部分段落字段、更多页码格式、通用制表位与旧快照迁移保留原缺口。当前数值页码不代表 Roman 等格式完整支持。
- 文档布局仍在主线程；长文档的 Worker/Rust 热点迁移、真实字体加载/错误、跨宿主测宽、IME/焦点、两种渲染后端、Office 独立重开及可复现性能/内存证据未验收。
- 三次之后仍不收敛返回 limit，拒绝第四次内容修改；这不是完整排版精度或所有 Office 样本已达标的声明。

## 手动更新与删除续批

继续读取固定 `IUpdateTableOfContentsCommandParams`、`IDeleteTableOfContentsCommandParams`、`updateTableOfContents`、`DeleteTableOfContentsCommand`，以及 [TOC menu](../../../univer/packages/docs-toc-ui/src/menu/menu.ts)、[上下文 Ribbon controller](../../../univer/packages/docs-toc-ui/src/controllers/table-of-contents-ribbon.controller.ts) 和 [更新对话框 operation](../../../univer/packages/docs-toc-ui/src/commands/table-of-contents-dialog.operation.ts)。固定 update DTO 为 `{ unitId?, rangeId?, mode?: 'pageNumbersOnly' | 'entireTable' }`，缺省为全文；delete DTO 为 `{ unitId?, rangeId? }`。显式 rangeId 独立于当前光标；未传 ID 则由当前正文范围起点取得目标。更新等待完整排版，删除没有该依赖。

- 新 [document-toc-action-commands.ts](../../packages/office-studio/src/api/document-toc-action-commands.ts) 单独注册两个既有 ID。typed UI 请求继续返回同步 DocumentEditorCommandResult；公开 flat 走异步 executeCommand，直接 sync 调用拒绝。rangeId 支持实际外层 TOC 字段 ID 及领域内部稳定 ID；原生外层 owner 位于目录行外时，通过其真实字符范围解析唯一目录，拒绝缺失/歧义目标。
- flat update 将 `entireTable` 映射到既有领域 `all`，`pageNumbersOnly` 保持只更新页码，初次执行后最多再更新两次，总计至多三次；插入后的重排仍保留三次上界。每个手动更新只创建自己的初始历史，后续 amendment 不并入更早的插入或其他用户编辑。已准确的 page-only 返回 true，不增加空历史、不清除输入格式。
- flat delete 复用同一领域锁、权限、字段/SDT/评论与选区映射事务，无需 renderer；取消同文档旧页码任务。显式目标更新/删除保留独立正文光标或 story 输入及 typing marks；没有输入会话时不制造默认会话。上述源目标、用户、活动文档、历史、输入与布局版本均在异步等待后复核。
- [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) 的 delete 请求增加可选 `preserveInputSelection`，沿用 update 同一显式目标语义；权限与锁由完整目标目录、变更块和字段实体检查，不因光标位于别处跳过。`office-studio.ts` 原事务为这类显式正文目录动作保留独立 story/辅助 host 输入；两个 ID 从通用注册循环移到专用 adapter，BeforeCommand owner 保护扩到三条目录命令。
- `document-toc-layout.ts` 在等待返回与下一微任务之间再次检查 exact publication，拒绝已换字体/布局的旧页码。插入与手动更新复用此检查；提交观察者已关闭文档时，不会为了后台任务重新物化已关闭的 history。
- `DocumentWorkbench.vue` 在布局 pending 时禁用两种更新；删除仍读取实际目录锁/权限而不等待布局。 [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 的上下文和右键更新入口改用固定 `doc.operation.open-table-of-contents-dialog`，本地旧 operation 仍兼容到同一对话框。没有增加根项、叶项或占位按钮。

本续批的公共 **内容命令** DTO 和异步结果已接线。新增 [document-toc-dialog-editors.ts](../../packages/office-studio/src/api/document-toc-dialog-editors.ts) 将固定 `doc.operation.open-table-of-contents-dialog` 实际注册到 Studio，`FDocument.bindTocDialogEditor` / `openTableOfContentsDialog` 与 Workbench 的真实 Ant 对话框相连。打开前复核正文目录、模型/输入/历史、真实布局及可用更新权限；失活、卸载、关闭、同 ID reopen 或新 owner 替换使旧能力失效。工具栏/右键/API 复用这一 operation，legacy 本地 ID 只作视图别名；操作本身不修改内容或历史。

malformed DTO、ACL/锁与原生保真失败使用本项目明确异常，比固定部分 false 路径严格；目标缺失、无 renderer 或等待取消返回 false。所有可替换控件继续使用 antdv-next。

追加未运行草稿：runtime 的全文/只页码、实际 field rangeId、默认目标、无输入会话、独立 story/typing marks、锁/权限、取消/陈旧 guard、无 renderer 更新拒绝、无 renderer 删除、三次总上界和单历史；真实 dialog operation 的权限/输入/布局、owner 替换、BeforeCommand 改选与同 ID 释放；浏览器的 classic/grid 实际更新对话框→全文重排→删除→undo/redo、公开 operation 打开 Ant Modal 与失活/重激活；原领域 Ribbon 草稿改为固定 operation ID。运行验证继续为 0；后续统一静态快照另记，不以此前 14 文件结果覆盖本续批。

手动更新/删除及公共对话框续批静态快照：**20 个生产/草稿文件、21 段脚本、2 个 Vue 模板、963 条本地 import/export 引用和 3 份证据的 67 处本地链接，错误 0**；4 份 mjs `node --check` 与限定 `git diff --check` 通过。草稿声明/字面量 each 展开分别为 runtime **27/38**、字段缓存 **6/6**、浏览器 **4/6**、既有领域 TOC **34/34**；领域文件只调整原入口 ID，没有新增 34 个功能。上述仅语法/路径/文档和计数检查，不包括类型、构建或任何运行验收。
