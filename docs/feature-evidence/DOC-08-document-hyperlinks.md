# DOC-08 文档链接

更新日期：2026-11-27。状态：已补领域命令、顶部/快捷键/Canvas 接线、当前文档公开 operation 定位与契约、浏览器草稿，未验收。测试、类型检查、构建、Rust/WASM、浏览器、Office 文件样本和性能基准均按当前要求留待统一验证阶段。

本记录对应 [完整功能清单](../UNIVER_FEATURE_PARITY.md) 的 DOC-08。书签的独立插入、修改、删除及全部入口仍属于 DOC-08 的后续闭环，不能用本批链接命令宣称整个功能组完成。

## 固定来源与入口

参考仓库：`E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

| 入口类别 | 挂载、顺序及条件 | 固定来源 |
| --- | --- | --- |
| 常驻插入入口 | INSERT → MEDIA，order 1，grid r1c5；新建要求非折叠文字选区，显式已有 linkInfo 可在光标位置编辑 | `packages/docs-hyper-link-ui/src/menu/schema.ts`、`menu/menu.ts`、`commands/operations/popup.operation.ts` |
| 浮动工具栏 | order 20，使用同一编辑弹层 | `packages/docs-hyper-link-ui/src/menu/schema.ts` |
| 正文右键 | MAIN_AREA → DATA，order 0 | 同上 |
| 段落插入下方 | PARAGRAPH → LAYOUT → INSERT_BELLOW，order 6 | 同上 |
| 空段落右键 | EMPTY_PARAGRAPH → LAYOUT，order 6 | 同上 |
| 快捷键 | Ctrl/Cmd+K 打开同一编辑弹层 | `packages/docs-hyper-link-ui/src/commands/operations/popup.operation.ts` |
| 已有链接信息弹层 | 打开、复制、修改、移除四个动作；编辑与复制分别检查权限 | `packages/docs-hyper-link-ui/src/views/DocLinkPopup.tsx` |

本功能静态分类为常驻根挂载 1、浮动挂载 1、右键上下文挂载 3、信息弹层动作 4。classic 与 grid 属于同一常驻入口的呈现变体；更多菜单沿共享 Ribbon 的溢出机制收纳。这些计数只描述本功能的声明，不替代全部 Docs/Pro 入口清点，也不证明运行态可见性或验收完成。

真实 ID：

- `docs.command.add-hyper-link`
- `docs.command.update-hyper-link`
- `docs.command.delete-hyper-link`
- `doc.operation.show-hyper-link-edit-popup`
- `doc.operation.toggle-hyper-link-info-popup`
- `doc.operation.click-hyper-link`

模型与行为来源：

- `packages/docs-hyper-link-ui/src/commands/commands/add-link.command.ts`
- `packages/docs-hyper-link-ui/src/commands/commands/update-link.command.ts`
- `packages/docs-hyper-link-ui/src/commands/commands/delete-link.command.ts`
- `packages/docs-hyper-link-ui/src/utils/hyper-link-target.ts`
- `packages/docs-hyper-link-ui/src/views/DocHyperLinkEdit.tsx`
- `packages/docs-hyper-link-ui/src/views/hyper-link-edit/use-doc-hyper-link-edit.ts`
- `packages/docs-hyper-link-ui/src/services/hyper-link-popup.service.ts`
- `packages/docs-hyper-link/src/index.ts`
- `packages/docs/src/utils/custom-range-factory.ts`
- `packages/core/src/docs/data-model/text-x/build-utils/custom-range.ts`
- `packages/engine-render/src/components/docs/layout/style/custom-range.ts`

## 本批实现与数据边界

| 闭环 | 本地实现 | 实际行为 |
| --- | --- | --- |
| 实体派生、capability 与三个编辑命令 | [document-hyperlinks.ts](../../packages/office-studio/src/features/document/services/document-hyperlinks.ts) | 从现有 runs、href 与 `univerBodyHyperlinkSource.rangeId` 派生完整逻辑范围；`documentHyperlinkForSelection` 只查询完整实体，state 与 Canvas 共用；新建要求真实文字；修改和移除按 linkId 检查完整实体权限、内容锁和生成引用字段限制，覆盖已有生成链接的 Add 同样拒绝 |
| 统一命令与辅助内容区 | [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) | 三个真实 ID 纳入 DocumentCommand；正文、表格文字和 story 复用现有命令、候选与宿主发布；内部目标分类保留只读完整正文上下文，不增加跨边界 DTO |
| 输入、控件绑定与历史 | [document-editor-commands.ts](../../packages/office-studio/src/api/document-editor-commands.ts)、[content-control-editing.ts](../../packages/office-studio/src/features/document/services/content-control-editing.ts) | 标签改变时先按 linkId 解析实际选区，再进入现有文字输入和绑定 finalizer；目标改变但标签不变保留原 runs、嵌套实体与作者样式；命令注册继续由 OfficeStudio 的既有注册表和内容历史处理 |
| 运行切片 | [document-run-range.ts](../../packages/office-studio/src/features/document/model/document-run-range.ts) | 复用既有 UTF-16 范围切片，只增加受类型约束的逐 run patch；普通格式 patch 的入口校验不变 |
| 导入与写出 | [document-body-ranges.ts](../../packages/office-studio/src/api/document-body-ranges.ts)、[document-body-write.ts](../../packages/office-studio/src/api/document-body-write.ts) | 导入 url、bookmarkId、headingId 和旧 anchor；写出保留稳定 rangeId，并让外部 url 与内部 bookmarkId/headingId 排他；无来源信息的连续同 href runs 沿首个 `hyperlink:${run.id}` 合并，空 run 不截断实体，普通文字和显式不同 rangeId 保持独立；移除同时清理 href、metadata.link、metadata.href 和旧链接来源，防止原始范围重放 |
| Canvas 外观 | [document-scene.ts](../../packages/office-studio/src/features/document/rendering/document-scene.ts) | 默认蓝色和下划线由链接实体派生；新命令不写作者 color/underline。`textStyleMode: text` 保留原外观，`textColorMode: text` 保留原文字颜色但仍显示链接下划线；移除恢复作者外观 |
| 内部目标定位 | [document-public-body.ts](../../packages/office-studio/src/api/document-public-body.ts) | `documentHyperlinkTargetSelection` 复用公开 body index：BOOKMARK=3 的 rangeId/bookmarkId 优先，再查 paragraphId/headingId；返回正文位置，不修改内容模型 |
| 公开 operation 参数适配 | [document-public-body.ts](../../packages/office-studio/src/api/document-public-body.ts)、[document-segments.ts](../../packages/office-studio/src/api/document-segments.ts) | `resolveDocumentHyperlinkOperation` 在参数入口校验来源 unit、linkId、segmentId 和页面提示；编辑严格匹配当前公开 HYPERLINK 范围的包含式 startIndex/endIndex，复用服务的完整实体 Edit/内容锁；点击按真实来源范围检查 View；返回完整来源、scoped model、现有链接实体和实际 host，不修改模型或历史 |

标签变化复用原有 `replaceSelectionText`，继承首个选中文字样式，并沿原控件、字段、批注与历史事务更新。标签不变时直接更新链接属性，不替换文字，也不压合原有混合 runs。跨 run、段落的稳定链接按相邻的实际文字片段合并；中间存在其他链接或普通文字时保持独立范围。

UI、弹层与操作已接入 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue)、[document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 和 [DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue)。顶部 `doc-link` 转发固定编辑 operation，旧 `doc.operation.open-link` 保留；Ctrl/Cmd+K 仅从实际正文/story 输入桥触发，composition/229 和不可编辑内容区不能打开。弹窗用 Ant Modal/Form/FormItem/Input/Button，已有链接显示整条文字与目标，并提供应用、移除和取消；提交共用三个真实编辑命令。

弹窗复用现有 `captureStructureTarget` / `structureTargetCurrent` / `closeStructureDialog`，文档、API、选区或 host 改变时关闭陈旧弹窗。成功后在 `nextTick` 捕获更新后的焦点目标。Canvas Ctrl 点击通过共享查询取得完整 linkId，并转发 `{ unitId, linkId, segmentId?, source }`；工作台拒绝陈旧 source 后调用公开 resolver，按实际完整来源范围检查 View。外部目标复用 `safeUrl`，内部书签/标题目标按正文索引与目标 View 同步选择、滚动和焦点，不进入内容历史。明暗主题、焦点恢复与输入法仍待统一浏览器验证。

当前文档的编辑 operation 已映射固定 `{ link: { unitId, linkId, segmentId?, segmentPage?, startIndex, endIndex } }`，显式链接在光标折叠或位于其他内容时仍解析完整实体，并由工作台切换实际 host 后捕获弹窗选区。无显式 link 的 `undefined` / `{}` 请求保留当前选区路径；错误 unit、未知或歧义 segment、缺失 ID 和陈旧公开偏移直接拒绝，不回退到当前链接。点击 operation 已接受固定 `{ unitId, linkId, segmentId? }`，来源范围检查 View，内部目标仍在同一完整来源文档的正文查找；Edit 拒绝和内容锁本身不阻止有 View 权限的点击。

公开坐标与本地文字选区分开：startIndex/endIndex 是当前 body.dataStream 的 UTF-16 下标，endIndex 包含末字符，可能包含范围结束哨兵 `\x1e`；跨段落内部 `\r`、其他范围标记和 emoji 代理对都占公开偏移。本地 writer 生成链接时把结束哨兵放在末段落 `\r` 前，操作参数必须使用当前公开 customRange，不能再把末段落标记加进 endIndex，也不能沿用未重新写出的导入偏移。resolver 返回服务已有完整 selection，不自行用公开偏移建立另一份链接实体。

segmentId 复用现有公开 header/footer ID、story alias 和 note 宿主映射；segmentPage 只保留为页面/弹层位置提示，同一 story 重复出现多页时不会据此选择另一个 host。跨文档参数解析要求调用方传入 unitId 对应的完整来源文档；当前工作台只处理已绑定的文档，公开 registry 转发和跨文档 UI 激活尚未接入，不将 resolver 的来源返回值计为跨文档导航闭环。

尚未实现固定链接信息浮层及打开/复制/修改/移除四动作的全部挂载；当前弹窗和 Ctrl 点击不替代这些入口。旧本地 `{ target, source? }` 调用仍保留既有行为，真实 Canvas 已使用固定 ID resolver，不再通过旧 target 路径。浮动工具栏、三个右键挂载、跨文档转发和页面提示的真实弹层定位继续保留待实现，不能将 DOC-08 标完整。

固定版 `DocHyperLinkPopupService.canEditLink` 禁止独立编辑 `referenceFieldLink: true` 以及被 TOC、INDEX、TOA、REF、PAGEREF、NOTEREF、CITATION、BIBLIOGRAPHY 字段完全包围的生成链接。本地已用同一断言接入 capability 和真实 Add/Update/Delete 命令：按完整链接检查当前 run 属性和链接来源，显式 false 覆盖旧生成标志；字段范围复用 [documentTextFieldRanges](../../packages/office-studio/src/features/document/model/document-text-fields.ts) 的当前模型坐标，支持跨段落，避免用不可变导入偏移判断当前归属。普通 AUTHOR、未知字段及仅部分重叠的 REF 不触发这项生成链接限制。新建覆盖已有生成链接同样拒绝，点击继续只检查 View。生成字段的目标更新仍由其所属字段功能负责；本批只补独立链接编辑边界，没有宣称这些字段的全部生成行为已完成。

地址在 Add/Update 的唯一命令入口归一：去除首尾空白；裸网址补 `https://`；裸邮箱补 `mailto:`；`#id` 保留；已有协议复用现有 `safeUrl` 判断。弹窗提交原始 payload，避免再维护一套 URL 规则。固定 UI 将裸邮箱写成 `mailto://`，本项目使用浏览器标准邮件 URI `mailto:`，属于已记录的实现差异。

编辑、选择与小范围链接变换继续位于 TypeScript 服务层，没有更换解析、公式或布局权威引擎，也未新增 Worker/WASM 协议。当前没有端到端性能测量，不宣称链接处理已优化或达到大文档指标。

## 草稿与未验收项

[document-hyperlinks.contract.test.mjs](../../packages/office-studio/tests/document-hyperlinks.contract.test.mjs) 的 23 个契约声明覆盖：唯一 URL 归一边界、半开选区与 UTF-16、新建/修改/移除的真实 OfficeStudio 注册与内容历史、混合 runs 和嵌套字段、内部目标排他写出、保存重开不回生、重叠链接与普通文字间隙、实体完整范围权限与内容锁、标签替换和绑定控件同行更新、story 到正文导航、公开段落/范围标记及书签优先、Canvas 两个 text mode 和删除后的作者外观。公开 operation 与原生 href 的 7 个声明覆盖光标分离、错误身份/未知或歧义 segment/陈旧范围拒绝、UTF-16 与末段落标记边界、公开 story ID/alias 与页面提示、完整实体 Edit/锁和点击 View 区别、原生 href 完整 ID/作者样式/公开 body 重开，以及空 run/普通文字间隙/显式不同 ID 的写出边界。另新增 4 个声明覆盖八类生成引用字段的真实命令拒绝、当前/来源生成标志与 false 覆盖、当前跨段落字段范围及 AUTHOR/部分重叠 REF 可编辑、共享查询与 Edit 能力分离；拒绝路径断言模型与内容历史不变。

[document-hyperlinks.browser.test.mjs](../../packages/office-studio/tests/document-hyperlinks.browser.test.mjs) 使用真实 OfficeStudio/FDocument，草稿覆盖 classic/grid 顶部插入、Ctrl+K 更新/整条重命名/移除、作者样式、历史与内部快照重开；空值/unsafe 地址/缺失内部目标、书签/标题 Canvas 跳转、真实外部浏览器页面、选区/文档替换/API 重绑/deactivate、保护链接、legacy header、主题与合成 composition。新增草稿覆盖显式编辑参数、错误身份与陈旧范围、按 ID 内部跳转、header segment 切换、Canvas 来源 View 拒绝，以及真实 HTML 样式拆分的多个 href runs 经公开 body 重开保持完整实体。草稿未运行，合成事件不代表真实 Windows IME；内部快照重开不代表 DOCX 验收。

本阶段只做代码、限定 diff 与来源路径的静态阅读，未运行测试或工程检查。上述草稿没有运行；不能据此声称命令、浏览器行为、Office 导出重开或性能验收通过。

仍需统一验证：

- 本批 UI 的插入、显示文字修改、移除、Ctrl/Cmd+K、Ctrl 点击和内部跳转；主题、焦点恢复与输入法。链接信息浮层和复制地址等剩余入口须先补实现再验收。
- 显式编辑/点击参数、陈旧偏移、story 页面提示、跨文档 registry 转发与 UI 激活；真实 Canvas 的完整来源 View、生成引用字段链接的编辑拒绝与所属字段更新，以及旧本地 target 调用的保留行为。
- 正文、跨段落、表格、页眉页脚、脚注尾注、嵌套内容控件、绑定目标拒绝与历史重入的运行态行为。
- 实际 DOCX 解析、写出和独立重开；链接、书签、标题、嵌套字段和零宽实体的来源保真。
- 删除/移动书签和标题后的引用变换，以及独立书签 CRUD 的剩余入口。
- 浏览器邮件客户端行为、外部地址边界及真实大文档性能。
