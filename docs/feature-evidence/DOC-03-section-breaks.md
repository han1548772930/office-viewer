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
| 分隔符 / 分节符 | [schema][section-schema] `RibbonInsertGroup.MEDIA / doc.menu.breaks`，classic `order:4`，grid `row:1,column:1,rowSpan:2,showLabel:true`；分栏符 `order:0`，`next-page / continuous / next-column / even-page / odd-page` 为 1–5。[factory][section-menu] 的 `disableMenuWhenNoDocRange` 合并无范围、相交代码块和文档 Edit 权限；非传统文档隐藏。[operation][section-operation] 要求传统模式、折叠正文范围与整数偏移。 | [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) `doc-breaks` 在插入 / 媒体组 `order:4`，grid 位置一致；modern 隐藏。2026-11-27 续修将已传入的 editingHeaderFooter / editingAuxiliaryContent 和活动结构 Code 纳入根禁用态；表格单独禁用分栏/分节，不统一禁用普通分页符，详见[范围续核](#doc-view-break-scope)。叶项仍分派 `doc.command.insert-page-break / insert-column-break / insert-section-break`，五种分节参数是 `{ type }`。本项目会替换非折叠正文范围；分栏/分节的 `insertBreakBlock` 在 mutation 前拒绝表格范围，普通分页符 API 在表格中仍只产生文本换行，完整表格分页符尚未实现。 |
| 页眉和页脚常驻入口 | [schema][section-schema] `RibbonStartGroup.LAYOUT / OpenHeaderFooterPanelCommand`，classic `order:10`，grid `row:1,column:3,rowSpan:2,showLabel:true`；[factory][section-menu] 在传统 Docs 隐藏条件满足时显示，未单独注册 `disabled$`。 | `doc-header-footer` 在开始 / 段落与列表组，grid 位置一致；local 未显式填写 `order`，由稳定数组顺序排在列表入口之后，不能把上游 `order:10` 当作本地声明。`doc.operation.open-header-footer` 经 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) `openStory` 打开对应页的 header；Canvas 命中也走此 story。modern 或无匹配布局页时返回；正文选区独立保留。 |
| 页眉页脚上下文 / 插入 | [header/footer schema][section-header-footer-schema] `docs-ui.ribbon.header-footer` 为 contextual、`order:90`；insertion 组 `order:0`，control 根 `order:0`，grid `row:1,column:1,rowSpan:2`。 | `header-footer` 仅在传统模式且有活动 story 时生成；四组独立保留。插入组含页码菜单 PAGE / NUMPAGES / 更新字段，以及日期时间弹窗的 short / long / time 与自动更新。它们是 control 内部入口，不另计 Ribbon 根。[control](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonControl.vue) 以 section/story 权限禁用，关联前节时禁用内容编辑。 |
| 页眉页脚上下文 / 导航 | 同 schema navigation 组 `order:1`，control grid 位置同上；[upstream control][section-header-footer-control] 将当前区域对应的导航按钮禁用，第一节不能关联前节。 | 转到页眉 / 转到页脚走 `openStory`；链接开关走 `doc.command.set-section-header-footer`。当前区域的导航按钮与第一节链接开关禁用。本项目 `linkToPrevious` 是六类 story 共用的节级布尔值；上游 [SetSectionHeaderFooterLinkCommand][section-link-command] 按 `kind + variant` 独立关联，二者不等价。 |
| 页眉页脚上下文 / 选项 | options 组 `order:2`，control grid 位置同上。 | 首页不同 / 奇偶页不同修改 section header/footer 配置；布局按节首页与物理页奇偶选择默认、首页、偶数页 story。页码重启后的奇偶页语义仍未接布局。 |
| 页眉页脚上下文 / 位置 | position 组 `order:3`，control grid 位置同上；上游 InputNumber 为 0–200、precision 1。 | `marginHeader / marginFooter` 通过 `doc.command.set-section-setup` 写当前节，独立于内容关联；[layoutDocument](../../packages/office-studio/src/features/document/layout/document-layout.ts) 读取其位置，缺省 48px，零值保留。控件为 antdv-next Form / FormItem / InputNumber。 |
| 页眉页脚上下文 / 结束编辑 | Close 组 `order:5`，根 `order:0`，grid `row:1,column:1,rowSpan:2,showLabel:true`，命令 `doc.command.close-header-footer`。 | 本地 Close 组与根均声明 `order:5`，只有一个根项；`doc.operation.close-header-footer` 清除 story 输入状态、恢复此前 tab 与正文输入位置并请求焦点。该 operation 不产生内容历史。导航、关闭或观察者重入后的实际焦点效果仍待浏览器验证。 |

[DocumentHeaderFooterRibbonGroup.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonGroup.vue) 静态采用上游 control 的 toolbar 宽度阈值：position <1100px、options <900px、navigation <650px 时各自进入 Popover；insertion 不做此局部收纳，共享 overflow 状态下呈现完整 control。classic 使用行内 flex，grid 及收纳层使用 control 的分行布局。测量副本 inert，且不转发命令。尺寸、窄屏、整组 More、显隐/禁用、主题和焦点尚未实测。控件及日期弹窗均使用 antdv-next；本轮无新增 UI 或原生表单控件。

### 模型、命令、输入和持久化边界

- [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) `executeDocumentCommand / insertBreakBlock` 将分栏、分节保存为 first-class `columnBreak / sectionBreak`，五种分节类型单独保存。命令权限在服务入口再次检查；当前公共命令与快照历史接线见 [COMMON-03](COMMON-03-input-selection-axis-history.md)。历史、保存/重开草稿存在，未在本轮执行。
- [DocumentSectionSetup](../../packages/office-studio/src/domain/model.ts)、[document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts) 和 [DocumentSectionSettingsPanel.vue](../../packages/office-studio/src/components/document/DocumentSectionSettingsPanel.vue) 已有节尺寸/方向、页边距、栏数 1–12、栏间距、分隔线、节起始类型和 `pageNumberStart` 的可保存字段与命令。2026-11-27 本轮复查确认，当前 `layoutDocument` 已转到 [layoutDocumentWithNotes](../../packages/office-studio/src/features/document/layout/document-note-layout.ts)，由 [DocumentSectionFlow](../../packages/office-studio/src/features/document/layout/document-section-flow.ts) 读取 `documentEffectiveSectionSetup` 并生成实际页、节片段和栏，包含逻辑页码重启。此前“只读取全篇 page、未实现节多栏”的描述已经过时；本轮仅更正静态调用链，完整节布局与文件互操作仍未验收。
- 当前 [appendDocumentBlocks](../../packages/office-studio/src/features/document/layout/document-block-layout.ts) 将 `pageBreak` 分派到 `flow.pageBreak()`，将 `columnBreak` 分派到 `flow.nextColumn(true)`，分节标记进入 `flow.startSection`。`nextColumn` 在兼容页面且存在下一栏时保留物理页；`continuous` 在页面兼容、栏宽有效且有剩余空间时创建同页节片段；`evenPage / oddPage` 按逻辑页码补空页。页内节片段保存在 `page.sections`，不再只有单一节边界。这些是当前真实实现路径，所有模式、兼容条件、页码与 Office 样本仍需运行验证。
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

<a id="doc-view-break-scope"></a>

## VIEW 与分隔符菜单范围续核（2026-11-27）

本批按固定 v1.6.0 [schema][section-schema] 从左到右核对 VIEW 的三个常驻根：页面设置 `order:0 / r1c1 / rowSpan:2`、缩放 `order:1 / r1c2`、全屏 `order:2 / r2c2`；扩展、溢出专用和上下文根各零。另核对 INSERT 的一个分隔符根及其六个上游可执行叶项。本地对应四根均已声明，分隔符另保留一个普通分页符叶项；本批根及叶项增量均为零，计数不表示验收完成。

| VIEW 来源与实际接线 | 当前结论 |
| --- | --- |
| [display.menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/display.menu.ts)、[doc-zoom.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/services/doc-zoom.ts) | 本地三根 grid 位置一致；八个缩放预设为 50/80/100/130/150/170/200/400，状态栏和 Ribbon 同值。全屏沿 `doc.operation.toggle-fullscreen` 与真实 `fullscreenchange` 更新活动态。2026-11-28 已接每文档 `zoomRatio`、公开命令/事件与快照读写，详见[当前持久化记录](#doc-view-zoom-persistence)；工作台 ref 是该设置的显示投影，实际保存、切换与焦点仍未验收 |
| [PageSettingMenuItemFactory][section-menu]、[open-page-setting.operation.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/operations/open-page-setting.operation.ts) | 上游页面设置按文档 Edit 权限禁用，确认进入 DocPageSetupCommand 并仅在成功时关闭。此处保留 VIEW 续核时的阶段结论：本地 `doc.operation.open-page-settings` 打开现有 Ant 对话框，`applyPageSetup` 校验后写 page/documentFlavor 并调用 publish，缺少文档级 disabled，父层拒绝后子组件仍关闭。后续[页面设置 UI 续修](#doc-page-settings-ui-transaction)补权限禁用、父层成功关闭和文档切换清理；公共命令/历史、完整节尺寸覆盖及实际保存重开仍未完成或未验收 |

分隔符根原先只取 `editingDisabled`。Workbench 已传入活动页眉页脚/辅助宿主标志，而 [documentCommandSupportsBlockHost](../../packages/office-studio/src/features/document/commands/document-commands.ts) 明确排除三个 break 命令；Workbench 只给支持辅助宿主的命令补 host，旧菜单因而可能把辅助选区继续提交到根正文模型。此次在根菜单合并辅助/页眉页脚及活动结构 Code 的能力条件；Code 依据为固定 `menu.ts:331–369` 的 `disableMenuWhenNoDocRange`，明确排除与 CODE blockRange 相交的范围。主线已在 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 execute 开头合入三个 break 的辅助/legacy story 守卫，本批已静态重读；正文仍沿既有命令、权限与历史链路执行，modern 继续隐藏。

表格语义单独处理。固定 [update-document-section.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/commands/commands/update-document-section.command.ts) 的 `isValidTopLevelInsertionOffset`（288–294）排除 table、columnGroup 和 blockRange 内部；分栏命令（223）及分节命令（174）都调用它。因此本地用真实 `state.table` 禁用分栏叶项、分节容器和五个分节叶项，保留根菜单及普通分页符叶项。固定 schema 没有普通分页符入口；本地现有 page-break 表格分支只执行文本换行，没有建立页内 PAGE_BREAK 标记。该行为不等同于完整分页符，不从上游菜单缺项推断 Office/原始数据不支持表格分页符；其公开行为与文件模型保真仍待实现/验收。此次没有删掉这条既有路径。

[document-toolbar-state.contract.test.mjs](../../packages/office-studio/tests/document-toolbar-state.contract.test.mjs) 新增一个计划用例、七个范围场景，直接调用真实 Ribbon 和命令/工具栏状态；表格场景使用真实 create-table 返回的首单元格选区，区分仍可打开的根/普通分页符与禁用的分栏/分节，并核对隐藏、原分隔符参数和仍可用的缩放/全屏。[document-toolbar.browser.test.mjs](../../packages/office-studio/tests/document-toolbar.browser.test.mjs) 纠正已有两个视图草稿：大纲/专注沿实际“开始”扩展入口及必要的 More，全屏沿“视图”；页面设置后重开对话框检查宽度/上边距，移除已不存在的 doc-paper-size 断言。该浏览器夹具只使用外部 snapshot 历史，不覆盖 DocumentApi 页面设置事务。

本批改动限于 Ribbon 能力与既有草稿/证据；`document-toolbar-state.ts` 未确认新错态。没有新增组件、原生输入桥、命令、模型字段、Worker/WASM 或计算引擎。静态检查通过四段脚本（含一段夹具）、42 处本地导入/模块引用、本文 21 条本地链接与一处锚点，以及四个文件尾空白/冲突标记和限定 diff。测试、类型、构建、dev、浏览器、Rust/WASM、服务和基准均未运行；真实菜单禁用、Code 相交范围、跨表格的正文范围、辅助宿主保护、主题/箭头/hover、焦点和 IME、撤销重做与文件重开仍未验收；正文中嵌入 break 的模型保真和完整表格分页符继续保留缺口。

<a id="doc-page-settings-ui-transaction"></a>

## 页面设置权限、草稿与提交续核（2026-11-27）

本节前半保留弹窗与权限小修的阶段记录。2026-11-28 已实施下面的[统一页面设置事务](#doc-page-settings-command-transaction)；早期“没有新增命令”和待实现描述不表示当前仍直接发布页面模型。

状态：**未验收**。本批核对 VIEW 的一个既有页面设置根；常驻根增量零，扩展、溢出专用、上下文根及可执行叶项增量均为零。根保留 `doc.operation.open-page-settings`、528px Modal 和 grid `row:1 / column:1 / rowSpan:2 / showLabel:true`；没有新增页面设置命令或模型字段。

固定参考仍为 Univer v1.6.0 / `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。已读 [PageSettingMenuItemFactory][section-menu] 的文档 Edit 禁用、[打开 operation](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/operations/open-page-setting.operation.ts) 的 applied 成功关闭，以及 [DocPageSetupCommand](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/doc-page-setup.command.ts#L136) 的 `RichTextEditingMutation / UpdatePageLayout`。该上游命令还在传统模式更新全部顶层节的显式尺寸与边距，并对原本填满页面的表格按新内容宽度等比调整列宽（同文件 88、114、170、181 行）。这些行为在该阶段尚未接入，2026-11-28 的领域与事务实现见下节。

| 本批生产文件 | 实际接线与边界 |
| --- | --- |
| [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) | `DocumentRibbonContext` 增加可选 `pageSettingsDisabled`，既有页面设置根使用该值。未传字段的旧调用方保持原行为；选中段落的 Edit 禁用独立于文档级页面设置权限。没有改变缩放、全屏或原根的命令与位置 |
| [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) | 文档 Edit、失活和正在合成输入生成页面设置禁用态，共用于 Ribbon、Modal 和既有打开/纸张/方向 operation。父层保留原权限及几何校验，完成原 `publish` 后关闭弹窗，焦点改用现有 `scheduleInputFocus`。模型 ID 改变或 `documentApi` 引用切换时关闭旧弹窗；同文档权限变化不会丢掉草稿 |
| [DocumentPageSettings.vue](../../packages/office-studio/src/components/document/DocumentPageSettings.vue) | 增加缺省 false 的 `disabled` prop，使用 Ant Form.disabled 与 Modal.okButtonProps 禁用输入和提交，取消仍可用。`apply` 不再自行关闭，父层拒绝后草稿保持。错误反馈使用 Ant Alert；文本、分隔线和宽度说明使用当前 `theme.useToken()`，纸张、现代宽度、边距、尺寸与翻译行为继续复用原实现 |

所有涉及的可替换控件均使用 `antdv-next`；沿用宿主 ConfigProvider，不增加组件库、原生表单输入、Worker/WASM、计算域或跨边界协议。组件 API 根据本地 [docs/llm.md](../llm.md) 的 Form.disabled、Modal.okButtonProps、Alert.title 和 Design Token 章节核对，并用 context7 技能查阅同名官方仓库的 Form / Modal 文档。实际主题、禁用、取消、焦点恢复和 Windows IME 效果仍待浏览器验证。

本批给既有 [document-toolbar-state 契约草稿](../../packages/office-studio/tests/document-toolbar-state.contract.test.mjs)增加五个计划案例，直接使用真实权限服务与 Ribbon：文档拒绝、显式允许覆盖默认拒绝、默认拒绝、默认允许、段落单独锁定；同时检查旧调用方、根命令/位置和缩放/全屏保持。该文件当前 **20 处 test 声明 / 26 个计划案例**。[document-toolbar 浏览器草稿](../../packages/office-studio/tests/document-toolbar.browser.test.mjs)增加四个计划案例：两种初始主题下的权限撤销/恢复、失败保留、主题切换、成功一次历史、undo/redo、内部 JSON 重开和允许取消；无内容区的左右边距校验；切换模型 ID 后的草稿关闭、重开新尺寸及新文档历史。统一 AST 按全部声明及显式外层 for 展开后，该文件当前为 **20 处 test 声明 / 24 个计划案例**：17普通、2处 each 展开5例、1处外层 for 展开2例；先前19/22的文本计数漏了外层 for。夹具使用真实 Workbench 和 ConfigProvider，历史由外部 snapshot 适配提供；一次直接调用真实父层提交函数核对权限拒绝。它没有覆盖 FDocument 命令事务或 Office 文件往返，声明和计划案例数不代表执行结果。

<a id="doc-page-settings-command-transaction"></a>

### 统一页面设置事务与输入宿主（2026-11-28，已实施 / 未验收）

原七文件预估方案已被实际实现取代。[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 Modal、纸张和方向操作现在共用 `docs.command.page-setup`；有 Facade 时进入 `FDocument.executeCommand → applyDocumentEditorCommand → OfficeStudio` 原子内容事务，独立 Workbench 也调用同一领域实现后交给宿主历史发布。命令改变整份文档的页面设置；输入坐标及 typing marks 仍属于原正文、表格或辅助宿主。页面候选通过校验后一次发布，命令拒绝保持弹窗草稿，成功后才关闭。没有执行运行验收。

| 实际模块 | 当前行为与边界 |
| --- | --- |
| [document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts) | `DocumentPageSetupRequest / assertDocumentPageSetupRequest / prepareDocumentPageSetup` 分离 DTO、校验与候选。正数尺寸、非负边距、模式/方向及所有节的有效内容宽度必须成立；传统模式覆盖第一节及全部顶层 `sectionBreak.setup` 的显式页几何，现代模式保留节来源几何并使用现代默认边距。原栏配置、页眉页脚、资源与未知来源随候选保留 |
| 同领域服务的表格调整 | 旧页面正文宽与新宽差超过 2px 时，识别有显式列宽或表宽、列宽总和在旧正文宽 2px 容差内的表格，按原比例写新列宽和表宽；普通固定宽表格保持。递归处理正文、嵌套块、单元格/续格及已存辅助宿主。它是对固定 `resizePageFillTables` 的本地模型适配；不同节/嵌套宽度、不可拆行和真实 Office 行为仍需样本验收 |
| [document-editor-commands.ts](../../packages/office-studio/src/api/document-editor-commands.ts) / [document-page-setup-commands.ts](../../packages/office-studio/src/api/document-page-setup-commands.ts) | 编辑命令 DTO 与固定公开 `docs.command.page-setup` 共用类型化页设置和实际事务。公开入口支持平铺参数及显式编辑请求；文档 ID、选区、revision、history generation、input generation 都经原入口检查。注册函数注入现有提交函数，保留事务的私有边界 |
| [OfficeStudio](../../packages/office-studio/src/api/office-studio.ts) / [document-input-store.ts](../../packages/office-studio/src/features/document/stores/document-input-store.ts) | 页面设置走全局命令，不将辅助文档投影替换成正文。共享请求可选 `selectionHost`：省略时沿命令/有效会话，`null` 显式正文，对象显式辅助宿主；优先级高于自动宿主和 legacy story 继承。输入宿主已空时保留原非空坐标并失效输入、返回 `selectionCurrent:false`，不得把 `blockId:''` 发布为有效光标；整篇页设置仍可完成 |
| 平铺 API 的坐标契约 | `host?: DocumentAuxiliaryBlockHost` 指定选区所属空间。显式 selection 未带 host 时生成 `selectionHost:null`，不会继承当前脚注/页眉；省略 selection 可复用有效 body-kind 会话及其辅助 host，显式 host 可在该空间寻找首个合法目标。有效 legacy story 的默认调用省略 selectionHost 以保持原 story 输入；显式正文或 host 会覆盖它。所有传入选区与 host 都先分离复制 |
| Workbench / [DocumentPageSettings.vue](../../packages/office-studio/src/components/document/DocumentPageSettings.vue) | 页面提交保留模型、会话及历史代数 guards；观察者重入后的旧选区/焦点结果不再归还。Ant Modal/Form/InputNumber/Select/Button/Alert 沿统一主题，弹窗绑定文档 ID 与历史 generation，包含同 ID 重开失效检查，取消仍可用；失活、权限撤销及组字期间拒绝作者提交 |

[document-section-layout-api.contract.test.mjs](../../packages/office-studio/tests/document-section-layout-api.contract.test.mjs) 已补真实页设置候选、显式多节、现代模式、页面填满/固定宽表格、只读、字段和资源保留、历史/内部快照、BeforeCommandExecute 重入，以及平铺参数的正文/脚注/显式 host/legacy story/空辅助块合同草稿。前述外部 snapshot Workbench 草稿继续保留；它与 Facade 合同以及新增字符选择器浏览器草稿的覆盖空间分开，均未运行。

本批未新增常驻、扩展、上下文或溢出根；页设置 DTO、公开注册及共享 selectionHost 属实际接口变化。计算位于主线程可测试的领域服务，未新建 Worker/WASM 或第二套内容历史，不宣称性能已优化。静态检查结果在本批收尾单独登记；测试、类型、构建、dev、浏览器、Rust/WASM、Office 重开和基准仍后置。多节/辅助布局、空宿主的真实焦点、同 ID 重开、命令观察者重入、原生 IME、独立 Office 表格/节几何与大样本性能继续 **未验收**。

<a id="doc-view-fullscreen-focus"></a>

## VIEW 全屏切换与缩放焦点续核（2026-11-27）

本节为原全屏/焦点批的历史记录；文档级缩放与公开命令随后已接入，当前结论见[2026-11-28 持久化记录](#doc-view-zoom-persistence)。

状态：**未验收**。本批继续向右核对 VIEW 既有缩放与全屏两个常驻根；常驻、扩展、溢出专用、上下文根及叶项增量均为零。缩放仍为八个 50/80/100/130/150/170/200/400 预设、grid `r1c2`；全屏仍为 `r2c2`，沿原 `doc.operation.toggle-fullscreen`，命令、模型字段及公共 API 均无增量。

| 固定 v1.6.0 来源 | 已读契约与本地处理 |
| --- | --- |
| [display.menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/display.menu.ts) | 全屏状态先读取当前值，再订阅 `root.ownerDocument` 的 fullscreenchange；只有当前全屏元素等于该根时显示 Shrink/activated。本地增加挂载时的 `syncFullscreen`，状态取工作台 ownerDocument 并继续由真实事件更新 |
| [ToggleFullscreenOperation](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/commands/operations/toggle-fullscreen.operation.ts) | 宿主 document 已有任意全屏元素时先退出；没有全屏才请求根进入。本地原先只在工作台自身全屏时退出，其他根全屏则再次请求工作台。本批改为同一退出/进入顺序，使用工作台 ownerDocument，保留现有 Promise 失败反馈 |
| [getDocEffectiveZoomRatio / getDocZoomRatio$](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/services/doc-zoom.ts)、[SetDocZoomRatioOperation](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/operations/set-doc-zoom-ratio.operation.ts) | 缩放由 `getSettings().zoomRatio` 取值，operation 按 unitId 更新。已读 [DocumentDataModel](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/docs/data-model/document-data-model.ts) 的 setZoomRatio / getSettings，实际持久位置是 `snapshot.settings.zoomRatio`。该阶段本地使用工作台 ref 和宿主显示偏好，仅把缩放后的 nextTick 焦点请求改为 `scheduleInputFocus`；2026-11-28 已改读写文档设置，当前路径见下节 |

生产仅修改 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue)，约 17 行 diff。原页面设置、纸张/方向、输入会话、Ant 主题、Ribbon 声明和 footer 行为保留。代码仍属于主线程工作台视图状态和浏览器全屏 API；没有引入新的计算内核、Worker/WASM 或跨边界协议。

既有 [document-toolbar 浏览器草稿](../../packages/office-studio/tests/document-toolbar.browser.test.mjs)增加两处声明、三个计划案例。两个案例在普通/只读文档中，用 ConfigProvider 内的 Ant Button 触发真实 HTML 根全屏，再从实际 Ribbon 退出该根、进入工作台，并用真实 exitFullscreen 事件恢复菜单状态；核对模型、dirty 和外部历史保持。一个案例执行真实缩放 operation 后在同一轮替换文档，核对旧焦点请求取消、外部 Ant Button 继续持有焦点、新内容和历史保持。夹具按钮默认不挂载，不改变既有布局；没有模拟 Fullscreen API 或用消息代替交互。该草稿采用外部 snapshot 历史，未覆盖公共 FDocument 命令或 Office 文件。

按 AST 读取所有声明并展开显式外层循环，该批结束时浏览器文件为 **22 处 test 声明 / 27 个计划案例**：18 个普通案例、三处 each 展开七例、一处外层 for 展开两例。本批新增的三例与此前页面设置批次 **20 声明 / 24 计划** 分开记录，不恢复旧漏计数。契约文件本批未改，继续为 20 声明 / 26 计划。此次 AST 只作计数；测试、类型检查、构建、dev、浏览器、Rust/WASM、服务和基准执行结果均为 **0**，统一生产语法、引用、链接和 diff 收尾仍由主线完成。

该阶段文档级缩放状态、按 unitId 的公开 operation、[SetDocZoomRatioCommand](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/set-doc-zoom-ratio.command.ts) 的注册及保存重开尚未接入。2026-11-28 已实施文档设置与命令；内容历史/视图设置的本地选择、具体公开接口及仍未验收项见下节，不能继续用该历史缺口判断当前实现。

挂载初始全屏、异常/取消、iframe 或多 ownerDocument 的事件订阅、真实主题/RTL、浏览器焦点与 Windows IME、Office 样本和大规模缩放性能继续 **未验收**。该批结束时 fullscreenchange 的监听生命周期仍使用全局 document，后续[宿主生命周期续修](#doc-fullscreen-owner-document)补齐接线并增加草稿，跨 document 浏览器验收仍未执行。

<a id="doc-view-zoom-persistence"></a>

## VIEW 文档缩放设置、公开 API 与保存（2026-11-28）

状态：**已实施 / 未验收**。根任务已将每文档缩放接入当前模型、Facade、公开命令和保存链路，本节按实际源文件更新原“仅工作台 ref”的结论。固定来源为上述 `doc-zoom.ts`、`set-doc-zoom-ratio.operation.ts / command.ts`、`DocumentDataModel.getSettings / setZoomRatio`；本地沿用八个 Ribbon 预设及原 grid 位置，常驻、扩展、上下文、溢出根增量均为零。

[document-zoom.ts](../../packages/office-studio/src/features/document/services/document-zoom.ts) 以 `metadata.documentZoomRatio` 保存比例，缺省 1，显式输入要求 0.1–4 的有限数字。公开快照的 `settings.zoomRatio` 由 [document-input.ts](../../packages/office-studio/src/api/document-input.ts) 读取、[document-output.ts](../../packages/office-studio/src/api/document-output.ts) 写出，其他来源 settings 继续保留；内部快照保留模型设置。[document-zoom-commands.ts](../../packages/office-studio/src/api/document-zoom-commands.ts) 注册 `doc.operation.set-zoom-ratio`（`unitId / zoomRatio`）和 `doc.command.set-zoom-ratio`（`documentId? / zoomRatio?`）；后者沿固定缺省比例语义，省略文档 ID 使用活动文档。它们复用同一视图发布函数，不建立第二份缩放状态。

[FDocument](../../packages/office-studio/src/api/facade.ts) 暴露 `getZoomRatio / setZoomRatio / onZoomRatioChanged`；[DocumentZoomRatioChanged](../../packages/office-studio/src/api/events.ts) 带文档 ID、旧比例和新比例，订阅按文档过滤。Workbench 读取模型/事件后投影为百分比，Ribbon、Canvas、输入覆盖层和 Footer 使用同一值，文档切换重读该文档设置；Footer 真实缩放输入同步禁用态。焦点仍经过当前文档/会话守卫。来源快照和作者设置的实际重开行为已补草稿，未运行。

本地将缩放作为可随文档保存的视图设置：不推进内容 revision、dirty 或 typing marks，不加入内容撤销；内容 undo/redo 保留用户当前缩放。固定上游另有公共 Command 与撤销注册，本地保留命令入口并选择视图语义，差异明确登记，不能据此宣称内部实现相同。权限、BeforeCommandExecute 重入及宿主生命周期继续走根任务共享入口。[document-zoom.contract.test.mjs](../../packages/office-studio/tests/document-zoom.contract.test.mjs) 和 [document-zoom.browser.test.mjs](../../packages/office-studio/tests/document-zoom.browser.test.mjs) 有公共/内部快照、文档隔离、命令/事件、取消重入、内容历史和实际 Canvas/输入缩放草稿；它们属于根任务验证范围，执行结果仍为零。

这次只更新 Docs 证据，共享实现和根导出由根任务负责。实际跨 document/iframe、主题/RTL、放大后的坐标命中、全屏、原生 IME、保存重开及普通/大样本缩放性能均 **未验收**，未宣称 WASM 或端到端性能优化。完整 DOC-03 / DOC-23 与全部 Pro 目标不变。

<a id="doc-fullscreen-owner-document"></a>

## 全屏 ownerDocument 生命周期续修（2026-11-27）

状态：**未验收**。继续核对 VIEW 的既有全屏根，常驻、扩展、溢出专用、上下文根与菜单叶项增量均为零。固定参考仍为 Univer v1.6.0 / `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；[DocFullscreenMenuItemFactory](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/display.menu.ts#L52) 订阅 `root.ownerDocument` 并先发初始状态，[ToggleFullscreenOperation](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/commands/operations/toggle-fullscreen.operation.ts#L29) 从同一根获取宿主后执行真实全屏 API。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 增加内部 `fullscreenDocument` 资源，挂载在实际 ownerDocument 注册；卸载从原注册 Document 移除同一监听，即使根已被移到其他 Document。重新激活时比较当前宿主，变化后解绑旧宿主、绑定新宿主并同步全屏状态。失活只在工作台确为当前宿主全屏元素时退出，保留 Promise 失败反馈。生产改动沿现有 Vue 生命周期；上批缩放焦点路径保留，公共 API、命令、模型、内容历史、保存和 Worker/WASM 协议没有变化。

既有 [document-toolbar 浏览器草稿](../../packages/office-studio/tests/document-toolbar.browser.test.mjs) 增加两个普通计划案例。父页面创建 Vue 应用后挂到同源 iframe，构成模块全局 document 与根 ownerDocument 不同的真实条件。第一例使用 iframe 中的实际 Ribbon 进入全屏、调用原生退出，随后将根移到父 Document 再卸载，核对事件状态与原宿主监听清理。第二例在根仍挂载于 iframe 时直接调用真实 Vue 失活回调，核对原生退出发生在 iframe；随后通过真实 KeepAlive 失活、移动宿主容器和重新激活，核对旧订阅移除、新宿主订阅、全屏状态和最终清理。直接调用回调这一段用于隔离宿主退出契约，不替代真实缓存失活或输入法验收。

夹具仅包装 addEventListener、removeEventListener 和 exitFullscreen 来记录宿主及监听集合，全部转发原生调用；没有伪造 fullscreenElement 或替换 Fullscreen API 的结果。两例继续核对文档快照和外部历史保持。已有及新增工作台和主题宿主使用真实 `antdv-next` / ConfigProvider；本批新增 iframe 和挂载容器为测试宿主，没有新增产品 UI 控件。

按 AST 读取全部声明并展开显式外层循环，该阶段浏览器文件为 **24 处 test 声明 / 29 个计划案例**：20 个普通案例、三处 each 展开七例、一处外层 for 展开两例；toolbar-state 契约当时为 **20 声明 / 26 计划**。本批 AST 只作计数；测试、类型检查、构建、dev、浏览器、Rust/WASM、服务和基准执行结果均为 **0**，统一静态语法、引用、链接和 diff 收尾由主线完成。挂载初始全屏、拒绝/取消、跨宿主实际运行、主题/RTL、焦点恢复、Windows IME、Office 样本及性能继续 **未验收**；页面设置和缩放的后续实现见上述 2026-11-28 记录，不改变完整范围和验收要求。

<a id="doc-table-page-break-review"></a>

## 表格分页符链路与页眉页脚控件续核（2026-11-27）

本轮继续 DOC-03，并核对 DOC-06 / DOC-22 的既有表格布局和 DOCX 读写边界。固定仓库 HEAD 与上述 v1.6.0 提交一致，所读分隔符和页眉页脚来源没有在途修改。常驻、扩展、上下文、溢出根及菜单叶项均无增量；没有修改命令、领域模型、历史、Worker/WASM 或保存协议。

| 当前真实调用链 | 结论与可复用边界 |
| --- | --- |
| [insert-page-break](../../packages/office-studio/src/features/document/commands/document-commands.ts) `replaceSelectionText → blockContainer` | 先拆分选区文本，真实 cell/嵌套容器已能定位；`split.focus.row == null` 仍阻止向表格插入 `pageBreak`。表格分支只换行，不能记作分页功能 |
| [块测量](../../packages/office-studio/src/features/document/layout/document-block-measurement.ts) → [表格测量](../../packages/office-studio/src/features/document/layout/document-table-measurement.ts) → [表格分页](../../packages/office-studio/src/features/document/layout/document-table-layout.ts) | 块测量跳过三种 break，两个 measurement DTO 均没有强制分页位置。`tableCut` 只消费行、文字和图片的不可分割区间；跨页/栏沿 `nextColumn`。既有 `insertDocumentTablePageGap` 能平移派生几何，但尚无手工分页输入。仅放开 cell 标记仍不会实际分页 |
| [docx_table.rs](../../crates/office-writer/src/docx_table.rs) `cell_blocks / continuation_blocks → context.blocks` → [docx.rs](../../crates/office-writer/src/docx.rs) `block` | 普通单元格和续格都复用同一块写出器，已有 `pageBreak → w:p/w:r/w:br type=page`，嵌套表格也递归使用此入口。本轮未改 writer，未执行真实文件重开 |
| [parser.rs](../../crates/office-parser/src/parser.rs) `insert_paragraph_breaks / body_block_position` | 仅遍历根 `DocModel.nodes` 的 Paragraph；当前没有遍历 TableCellContentNode 提取 cell 分页符。根正文中的可见文字混合分页符也保留为换行并发出警告 |
| [空段落抽取](../../crates/office-parser/src/document_block_extraction.rs)、[表格来源](../../crates/office-parser/src/document_table_sidecar.rs)、[cellBlocks](../../packages/office-studio/src/parser/documentTables.ts) | 空 cell 段落可保留身份，表格来源可恢复属性/拓扑，cell children 共用正常块归一化；这些通路未恢复 `w:br` 的 page 类型。已读本地 `duke-pages 0.6.0/src/parse/paragraph.rs:156`，其文本 token 将 `w:br / w:cr` 转为 `\n`，`lastRenderedPageBreak` 单独忽略。不能据此宣称 cell 分页符往返完整 |

固定 [insert-break.operation.ts][section-operation] 和 [update-document-section.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/commands/commands/update-document-section.command.ts) 的分栏/分节仍要求传统模式和顶层插入位置，排除表格、栏组和结构块内部。本地普通分页符另有入口，继续保留。固定 [doc-skeleton.spec.ts:6150](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/__tests__/doc-skeleton.spec.ts#L6150) 的 fitting-table 用例将 PAGE_BREAK 登记在 `renderedPageBreaks` 中并保持一页；这是缓存分页场景，不能作为手工分页符的实现规格或验收证据。标准 Office 中 cell 内手工分页、不可拆行、嵌套表格和重复表头的行为仍需文件样本核实。

表格分页符属于较大改动，尚待确认；不在本轮修改生产链路。可复用现有 `pageBreak` 模型、真实 cell 定位、内容历史、节流布局和块写出器，预计范围如下：

| 文件 | 预计改动 | 具体内容 |
| --- | --- | --- |
| `features/document/commands/document-commands.ts` | 15–25 行 | 将标记插入真实 cell 容器，保留拆分后的地址、光标和原内容事务 |
| `features/document/layout/document-block-measurement.ts` | 35–60 行 | 输出带逻辑地址的强制分页位置，递归保持结构块和嵌套表格偏移 |
| `features/document/layout/document-table-measurement.ts` | 35–65 行 | 将 cell/续格分页位置投影到表格坐标，随现有派生 gap 平移，明确 padding、垂直对齐和 exact row/clip 行为 |
| `features/document/layout/document-table-layout.ts` | 65–100 行 | 在自动切片前消费手工分页，使用已有物理页切换；处理多栏、连续/首尾标记、重复表头和跨行对象，保证分页持续推进 |
| `crates/office-parser/src/parser.rs` | 100–160 行 | 复用 story/source 地址遍历 cell 和嵌套段落，恢复 standalone/混合分页符并保留运行、字段、对象和来源；不能把混合段落整体替换为空 marker |

上表前端路径均相对于 `packages/office-studio/src`；生产预计 250–410 行，writer 和领域枚举直接复用。验证草稿预计另需 260–405 行，追加到既有 [section-breaks 契约](../../packages/office-studio/tests/document-section-breaks.contract.test.mjs)、[table-layout 契约](../../packages/office-studio/tests/document-table-layout.contract.test.mjs)、[table-workbench 浏览器夹具](../../packages/office-studio/tests/document-table-workbench.browser.test.mjs) 和 [Rust 表格往返草稿](../../crates/office-parser/src/document_table_roundtrip_tests.rs)。覆盖真实 cell/续格/嵌套选择、普通与不可拆行、多栏中的物理分页、重复表头、首尾/连续分页、撤销重做、内部保存及 DOCX 导出重开；主线测试未在本轮修改或执行。

本轮实际小修仅为 [DocumentHeaderFooterRibbonControl.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonControl.vue) 的两处既有模板行：页码按钮以真实 `pageOpen` 声明 `aria-expanded`；日期和时间 Modal 显式使用 360px 宽度，来源为固定 [DocHeaderFooterFields.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/header-footer/panel/DocHeaderFooterFields.tsx) 的 `Dialog width={360}`。页码与紧凑组仍共用 caption，SVG 不收缩；text 按钮保持零边框、实际 hover 边界和主题 token，按下仍保留输入焦点。没有新增 UI，既有 Modal/Form/Select/Checkbox/Button 继续使用 `antdv-next`。

日期对话框的 locale 实际预览、命令拒绝后保留草稿与失败反馈，以及真正关闭弹层后的焦点恢复，仍与上游存在待核实/待实现差异。该阶段没有更改这些事务；页眉页脚 Modal 的独立打开路径仍未接入。完整字符选择器与页面设置命令随后已实施，见[字符目录/交互记录](DOC-03-header-footer-ribbon-sizing.md#doc-character-picker-plan)及[页面设置事务](#doc-page-settings-command-transaction)。该阶段静态检查通过脚本、模板、样式各一份，三处本地导入与三处包引用，本文当时 39 条本地链接、一处链接锚点、13 条固定源码路径与 11 处引用式链接；限定 diff、尾部空白和冲突标记检查通过。这些历史结果不覆盖后续改动。测试、类型、构建、Rust/WASM、dev、浏览器、服务和基准均后置，实际尺寸、展开状态、主题/RTL、hover、焦点、IME、分页、历史与文件闭环继续未验收。

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
