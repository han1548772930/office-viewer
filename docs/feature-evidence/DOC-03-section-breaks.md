# DOC-03 Section Break Evidence

Status: **Not accepted**. This evidence covers editable page, column, and section break markers. It does not establish full DOC-03 parity.

## 节布局接线后的当前实现记录（2026-11-22 / Asia/Shanghai）

本轮已将节级尺寸、连续节片段、多栏流动、下一栏、逻辑页码与奇偶补页接入正文布局，并实现六类页眉页脚 story 独立关联、活动 story 的字段刷新和补页导航过滤。续作增加有限数字字段开关、动态 mixed cache 迁移、真实初始 cache、DATE/TIME 显式 picture 的字面量/转义/语言格式、既有来源语言优先级与清除格式后的语言回退、自定义日期的公共历史与内部重开草稿、上下文权限/弹层范围保护，以及 story B/I/U 的共用格式命令；原生 DOCX 字段、首缓存 run 语言与 dormant first/even 保留也追加了写出/parser 重开草稿。模块、字段兼容、上游来源、公开规格缺口、测试草稿及真实未完成项见 [DOC-03 节布局与独立 story 关联](DOC-03-section-layout.md)。这些是静态实现记录；未运行测试、类型检查、构建、Rust/WASM、浏览器、格式/语法检查、链接检查或基准，DOC-03 仍为 **未验收**。本子任务只补所属日期/story 服务、测试与领域协议 helper，parser/writer 运行链及命令权限的共享实现由根任务记录；不能从这些草稿推断完整 Office 互操作已完成。

下面保留节布局接线前的阶段记录。该段中的“当前”“本轮”和剩余项描述当时的工作树；它已经被上面的实现记录及专项文件补充，不应继续据此认定现在没有节布局或只有全节关联。

## 节布局接线前阶段静态记录（2026-11-22 / Asia/Shanghai）

本节以恢复后的工作树为准，覆盖 DOC-03 的分隔符与页眉页脚入口，并关联 DOC-02 / COMMON-03 的字符格式、输入与历史接线。固定参考为 Univer v1.6.0 / `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；静态读取的本地 HEAD 与此一致，下面列出的菜单及控制器文件未显示在该参考仓库的修改列表中。本轮仅静态读取和更新证据；未运行测试、类型检查、构建、Rust/WASM、浏览器、格式/语法检查、链接检查或基准。下方英文记录保留为此前阶段的历史说明，其中的 modal-only、无 rich field rendering、固定提交未核实和通过数量不代表本轮工作树的现状或验证结果。DOC-03 仍为 **未验收**。

### 入口、顺序与统计口径

本页详细展开 **2 个常驻 Ribbon 根声明**（分隔符菜单、页眉页脚入口）及 **5 个页眉页脚上下文 Ribbon 根声明**（四组 control + Close）。扩展专用根、溢出专用根各 **0**；共享 More / 窄宽弹层复用既有入口，不再计根声明。固定上游分隔符菜单为分栏符 1 + 分节符 5 个可执行子项；本项目另加普通分页符，并把五种分节符放入一个“分节符”容器，故为 **7 个可执行叶项 + 1 个容器**。这不是全部 DOC-03/Pro 能力数，也不是运行态可见入口数或完成数。节设置另有正文右键 layout 入口，不计入 Ribbon 根声明。

| 入口 | 固定上游来源与布局 | 本项目静态接线、状态与差异 |
| --- | --- | --- |
| 分隔符 / 分节符 | [schema][section-schema] `RibbonInsertGroup.MEDIA / doc.menu.breaks`，classic `order:4`，grid `row:1,column:1,rowSpan:2,showLabel:true`；分栏符 `order:0`，`next-page / continuous / next-column / even-page / odd-page` 为 1–5。[factory][section-menu] 的 `disableMenuWhenNoDocRange` 合并无范围、相交代码块和文档 Edit 权限；非传统文档隐藏。[operation][section-operation] 要求传统模式、折叠正文范围与整数偏移。 | [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) `doc-breaks` 在插入 / 媒体组 `order:4`，grid 位置一致；modern 隐藏，正文不可编辑或处于 story 时禁用。叶项分派 `doc.command.insert-page-break / insert-column-break / insert-section-break`，五种分节参数是 `{ type }`。本项目会替换非折叠正文范围；分栏/分节的 `insertBreakBlock` 在 mutation 前拒绝表格范围，但根菜单未以表格范围单独禁用；普通分页符命令在表格中仅产生文本换行，不应记为完整表格分页符实现。 |
| 页眉和页脚常驻入口 | [schema][section-schema] `RibbonStartGroup.LAYOUT / OpenHeaderFooterPanelCommand`，classic `order:10`，grid `row:1,column:3,rowSpan:2,showLabel:true`；[factory][section-menu] 在传统 Docs 隐藏条件满足时显示，未单独注册 `disabled$`。 | `doc-header-footer` 在开始 / 段落与列表组，grid 位置一致；local 未显式填写 `order`，由稳定数组顺序排在列表入口之后，不能把上游 `order:10` 当作本地声明。`doc.operation.open-header-footer` 经 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) `openStory` 打开对应页的 header；Canvas 命中也走此 story。modern 或无匹配布局页时返回；正文选区独立保留。 |
| 页眉页脚上下文 / 插入 | [header/footer schema][section-header-footer-schema] `docs-ui.ribbon.header-footer` 为 contextual、`order:90`；insertion 组 `order:0`，control 根 `order:0`，grid `row:1,column:1,rowSpan:2`。 | `header-footer` 仅在传统模式且有活动 story 时生成；四组独立保留。插入组含页码菜单 PAGE / NUMPAGES / 更新字段，以及日期时间弹窗的 short / long / time 与自动更新。它们是 control 内部入口，不另计 Ribbon 根。[control](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonControl.vue) 以 section/story 权限禁用，关联前节时禁用内容编辑。 |
| 页眉页脚上下文 / 导航 | 同 schema navigation 组 `order:1`，control grid 位置同上；[upstream control][section-header-footer-control] 将当前区域对应的导航按钮禁用，第一节不能关联前节。 | 转到页眉 / 转到页脚走 `openStory`；链接开关走 `doc.command.set-section-header-footer`。当前区域的导航按钮与第一节链接开关禁用。本项目 `linkToPrevious` 是六类 story 共用的节级布尔值；上游 [SetSectionHeaderFooterLinkCommand][section-link-command] 按 `kind + variant` 独立关联，二者不等价。 |
| 页眉页脚上下文 / 选项 | options 组 `order:2`，control grid 位置同上。 | 首页不同 / 奇偶页不同修改 section header/footer 配置；布局按节首页与物理页奇偶选择默认、首页、偶数页 story。页码重启后的奇偶页语义仍未接布局。 |
| 页眉页脚上下文 / 位置 | position 组 `order:3`，control grid 位置同上；上游 InputNumber 为 0–200、precision 1。 | `marginHeader / marginFooter` 通过 `doc.command.set-section-setup` 写当前节，独立于内容关联；[layoutDocument](../../packages/office-studio/src/features/document/layout/document-layout.ts) 读取其位置，缺省 48px，零值保留。控件为 antdv-next Form / FormItem / InputNumber。 |
| 页眉页脚上下文 / 结束编辑 | Close 组 `order:5`，根 `order:0`，grid `row:1,column:1,rowSpan:2,showLabel:true`，命令 `doc.command.close-header-footer`。 | 本地 Close 组与根均声明 `order:5`，只有一个根项；`doc.operation.close-header-footer` 清除 story 输入状态、恢复此前 tab 与正文输入位置并请求焦点。该 operation 不产生内容历史。导航、关闭或观察者重入后的实际焦点效果仍待浏览器验证。 |

[DocumentHeaderFooterRibbonGroup.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonGroup.vue) 静态采用上游 control 的 toolbar 宽度阈值：position <1100px、options <900px、navigation <650px 时各自进入 Popover；insertion 不做此局部收纳，共享 overflow 状态下呈现完整 control。classic 使用行内 flex，grid 及收纳层使用 control 的分行布局。测量副本 inert，且不转发命令。尺寸、窄屏、整组 More、显隐/禁用、主题和焦点尚未实测。控件及日期弹窗均使用 antdv-next；本轮无新增 UI 或原生表单控件。

### 模型、命令、输入和持久化边界

- [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) `executeDocumentCommand / insertBreakBlock` 将分栏、分节保存为 first-class `columnBreak / sectionBreak`，五种分节类型单独保存。命令权限在服务入口再次检查；当前公共命令与快照历史接线见 [COMMON-03](COMMON-03-input-selection-axis-history.md)。历史、保存/重开草稿存在，未在本轮执行。
- [DocumentSectionSetup](../../packages/office-studio/src/domain/model.ts)、[document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts) 和 [DocumentSectionSettingsPanel.vue](../../packages/office-studio/src/components/document/DocumentSectionSettingsPanel.vue) 已有节尺寸/方向、页边距、栏数 1–12、栏间距、分隔线、节起始类型和 `pageNumberStart` 的可保存字段与命令。当前 `layoutDocument` 仍在开头只解析全篇 `model.page`，未使用这些字段实现节多栏、节尺寸、页码重启；不能再概括为“没有设置模型”，也不能把模型字段或 Drawer 控件记为排版实现。
- 当前 `pageBreak / columnBreak / nextPage / nextColumn` 均前进到下一页；`continuous` 留当前页；`evenPage / oddPage` 按物理页索引补页。连续节还会改写当前页的单一 `sectionId`，无法表达同页多节边界和各自内容范围。这些是待修复的布局差异。
- [section-header-footer.ts](../../packages/office-studio/src/features/document/services/section-header-footer.ts) 保存六类页眉页脚 story 字符串、`storyContent` 富运行/段落与 `pageFields`；[story-rendering.ts](../../packages/office-studio/src/features/document/services/story-rendering.ts) 与 Canvas 布局现已按实际运行样式绘制多行文字、字段 cache 和链接样式。输入缓冲里的 `\uFFFC` 只占一个逻辑对象范围，显示侧使用真实字段结果宽度，不能继续概括为“没有 rich field rendering”。这仍不证明复杂字段和完整 Office story 保真。
- PAGE / NUMPAGES 由当前布局计算；DATE / TIME 保留 instruction 与 cachedResult。`handleStoryAction('refresh-fields')` 仅按能识别的 locale 日期模板更改选中节的 cachedResult；混合 cache 运行刷新后的范围/样式迁移、任意字段开关、页码格式及节级继续/重启均未完成。本轮未改 Rust/parser/writer；DOCX 混合 cache 两次本项目 native writer/parser 草稿见 [DOC-02](DOC-02-clear-formatting.md)，未运行，不能代替 WASM 和独立 Office 重开。
- [DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 的正文与 story 输入均为 antdv-next `Input.TextArea`。公共会话捕获检查活动 facade、模型身份/快照、范围、revision 与 generation；clear / history / 内部重开后的真正输入可重建空 marks 会话，旧模型或旧 TextArea 的 composition / input / history / keyboard 事件被拒绝。展开字符格式、混合工具值和输入继承的实现/差异见 [DOC-02](DOC-02-clear-formatting.md)；这与 Toolbar/Canvas 的真实 Windows IME 验收分开。

### 本轮验证与仍未完成

静态回读涉及菜单、上游 operation/control/link command、本项目命令、设置服务、Canvas layout、story 输入以及现有草稿；没有执行任何验证命令。`document-header-footer-ribbon.contract.test.mjs`、`document-section-breaks.contract.test.mjs`、`document-section-header-footer.contract.test.mjs` 覆盖 schema/归一化、内容关联、首/偶页选择和位置/字段场景；`document-story-formatting.contract.test.mjs`、`document-story-api.contract.test.mjs`、`document-story-input.browser.test.mjs` 有富运行/cache、锁、clear、输入会话、history 与内部重开草稿，均不能当作本轮通过证明。story 浏览器 fixture 直接使用 FDocument，生产应用 history adapter 的接线/草稿另见 COMMON-03，不冒称 fixture 已验证生产宿主。

剩余目标：节多栏与下一栏、连续节边界/节几何、节页码重启/继续与格式；六类 story 独立关联和混合 cache 自动刷新；复杂/nested/未知字段、未知 story XML、正文中嵌入的分页/分栏符；脚注/尾注和线程评论的富内容/Canvas 编辑；classic/grid/More、权限、主题、焦点、真实 IME、独立 Office、Canvas2D/WebGPU 与大样本性能。上述仍为待开发/待修复或待验证，全部 Pro 目标保留。本轮只修改本证据，未新增公共接口或产品实现。

[section-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts
[section-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts
[section-operation]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/operations/insert-break.operation.ts
[section-header-footer-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/header-footer-ribbon.ts
[section-header-footer-control]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/header-footer/panel/DocHeaderFooterRibbonControl.tsx
[section-link-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/commands/commands/set-section-header-footer-link.command.ts

## Earlier reference and implementation records

### Earlier working-tree record, not rerun

The following supersedes the earlier modal-only observation below. The traditional Docs workbench now connects the Insert entry and Canvas story hit areas to a page/section/variant-specific antdv-next input bridge. The body selection is retained separately. The contextual tab has four inline control groups and Close. Responsive grouping uses the toolbar-width thresholds in `packages/docs-ui/src/views/header-footer/panel/DocHeaderFooterRibbonControl.tsx`. Insert/refresh date and time uses the locale-derived masks in upstream `utils/date-time-field.ts`. Positioned fields are represented independently of literal UTF-16 text, including repeated fields and field cached results. DOCX import reads the referenced story XML, and export places simple fields between literal runs. Contracts for field deletion, Unicode, distance rounding, writer/reader reopening and a browser editing path have been added, not executed. The WASM artifact has not been rebuilt for this change.

Still unverified/incomplete: classic and grid width/overflow, localized display, permissions and IME across focus transitions, rendering rich multiline stories, unresolved nested field instructions and preservation of unknown story XML. The text bridge denotes an editable field by an object replacement character; it is not rich field rendering. Manual refresh currently handles recognized date/time masks in the selected section. This work does not establish DOC-03 acceptance.

### Earlier implementation notes

2026-11-22 (implementation, unified tests pending): v1.6.0 `docs-ui/src/menu/header-footer-ribbon.ts` registers four distinct control groups plus Close; local contextual schema now separates them rather than grouping five items under Insert. The running workbench still opens the header/footer modal, not an in-Canvas story-editing session. `views/header-footer/panel/use-header-footer-options.ts` and its tests scope header/footer distances to the current section. Local InputNumber controls write `marginHeader`/`marginFooter` through `doc.command.set-section-setup`, independently of story linking. Canvas uses the stored distances for text placement. DOCX reader/writer map section `w:pgMar` header/footer twips to/from CSS pixels, retaining the existing writer default of 720 twips (48px) when unspecified. Text/field changes no longer request body focus on each keystroke; closing the modal requests it. No new tests or WASM builds have been run for this change; real file reopening, IME, field positioning and contextual activation remain unaccepted.

- The local Univer Docs UI registers column breaks and five section break types: `nextPage`, `continuous`, `nextColumn`, `evenPage`, and `oddPage`.
- The local command operation dispatches section break type as an argument to `doc.command.insert-section-break`. The section update command applies an undoable document mutation and constrains insertion to the document body.
- The fixed reference commit could not be independently checked out from the local Univer repository during this task. These observations come from the locally readable source tree.

## Implemented behavior

- The Docs Insert Ribbon exposes page, column, and all five section break commands through nested menus.
- Section and column commands split the current body text selection and add a first-class marker block. Table-cell insertion is rejected before mutation. The existing document snapshot history records the change for undo and redo.
- Layout starts following content on the next page for `nextPage` and `nextColumn`. It skips a page when required by `evenPage` and `oddPage`. `continuous` retains the current page. This editor has no multi-column section layout model, so `nextColumn` currently advances to the next page.
- DOCX export writes `w:br` page and column break markers. Section markers write paragraph-level `w:sectPr/w:type` for each supported break type.
- DOCX import maps section start metadata and standalone page/column break paragraphs back to marker payloads. A manual break inside a paragraph with visible text remains a line break and emits a diagnostic warning.
- Header/footer models now support structured `PAGE` and `NUMPAGES` fields. The editor recalculates their displayed values from the current page layout. DOCX export writes `w:fldSimple` instructions and import preserves the field names on the matching header/footer reference.

## Verification

The November 22 position changes are excluded from the historical test counts below. The Rust parser now extracts the current section's header/footer distances, skipping historical `sectPrChange` properties. The DTO and document normalizer retain zero and missing values distinctly. Deferred contracts cover XML extraction, section-local layout with inherited content, and native writer XML with fractional pixel rounding. These are not evidence of a completed Office round trip.

- `pnpm.cmd --filter @office-studio/office-studio check:types` passed.
- Document command, Ribbon, and Canvas layout contract tests passed, 29/29.
- The Docs browser test inserted an odd-page section break from the Ribbon and verified undo and redo, 1/1 focused test passed.
- `cargo test --manifest-path crates/office-parser/Cargo.toml` passed, 11/11.
- `cargo test --manifest-path crates/office-writer/Cargo.toml` passed, 42/42.
- The page-field extension adds a layout contract proving current page and total pages update after a page break is inserted. Writer/parser DOCX contract verifies both field instructions survive export and parse.

## Unverified and remaining scope

- Continuous sections do not yet carry section-specific page geometry or columns. Header/footer content, page fields, and link-to-previous are modeled per section, but numbering formats, section-specific restart/continue settings, and page-field switches are not implemented.
- Multi-column layout is not implemented. Consequently `nextColumn` cannot move to a remaining column.
- Breaks embedded among visible text in one paragraph are not split into separate editable blocks on import.
- Browser coverage verifies the odd-page Ribbon path. It does not exercise all break types or DOCX browser import/export round trips.
- A generated DOCX has not been reopened in desktop Office software. The page-field export leaves the cached result empty and requests Word field refresh on open; the editor independently renders values from its current pagination. DOC-03 remains unaccepted.
