# DOC-03 节布局与独立 story 关联

状态：**未验收，开发中／待验证**。记录批次：2026-11-22 / Asia/Shanghai。固定参考为 Univer v1.6.0 / `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

本轮把已存在的节设置接入正文排版，增加六类页眉页脚的独立继承、有限数字字段格式、显式 DATE/TIME picture、来源语言优先级、多格式字段结果迁移和 UI 请求范围保护。这里只记录真实代码与静态对照；测试、类型检查、构建、Rust/WASM、浏览器、格式/语法检查、链接检查和性能基准均未执行。测试文件是待执行草稿，不能作为通过证明，也不代表 DOC-03 或 Pro 全量完成。

## 1. 上游入口与可观察契约

本轮沿用 [分节符入口证据](DOC-03-section-breaks.md) 中“节布局接线前阶段静态记录”的清点：2 个常驻 Ribbon 根声明、5 个页眉页脚上下文根声明，扩展专用根与溢出专用根各 0。根声明、菜单叶项、上下文 control 内部项分别计数；同一入口的 classic/grid/More 呈现不另计功能。该口径只覆盖此切片，不能推导全部 Docs/Pro 的运行态入口数量。

| 固定上游来源 | 已核对的契约及本地处理 |
| --- | --- |
| [docs-ui/menu/schema.ts][schema]、[menu/menu.ts][menu] | 分隔符菜单为插入／媒体组 `order:4`，grid `row:1,column:1,rowSpan:2`；分栏符与 nextPage/continuous/nextColumn/evenPage/oddPage 分节类型独立。现代模式隐藏；本地沿用已有 Ribbon/命令注册，不新增空入口。 |
| [docs-ui/menu/header-footer-ribbon.ts][hf-schema]、[DocHeaderFooterRibbonControl.tsx][hf-control] | 四组 control（插入、导航、选项、位置）与 Close，共 5 个上下文根；navigation 的关联开关作用于当前 header/footer variant，第一节拒绝关联。classic/grid/窄宽收纳语义沿用先前接线，尚未实测。 |
| [SetSectionHeaderFooterLinkCommand][hf-link] | `doc.command.set-section-header-footer-link` 参数包含 sectionId、kind、variant、linkedToPrevious。关联删除选中引用；取消关联克隆前节对应完整资源或建立空 story；其他五类引用保留。传统 Docs、第一节及资源引用边界在上游命令中判断。 |
| [DocumentSkeleton][skeleton] | `getSectionPageNumberStart`、`balanceFinalContinuousColumnSection`、`_addNewSectionByContinuous`、`_addNewSectionByNextColumn` 区分物理页兼容、同页节片段、剩余栏和逻辑页码。连续节与下一栏不能统一换页。 |
| [layout/model/page.ts][page] | `getNextSkeletonPageNumber` 区分共享旧物理页与新节后续页标签；首页与偶数页 story 按节起始及逻辑页码选择；长 story 可反向增加正文预留空间，本地尚未接入这一部分。 |
| [utils/sections.ts][sections]、[utils/section-columns.ts][columns]、[facade/f-document-section.ts][section-api]、[use-section-setting.ts][section-settings] | 顶层节枚举、节设置与栏参数是模型/公开操作契约。本地仍以顶层 sectionBreak 和可序列化设置为权威；嵌套节和完整 Facade 方法目录不由此切片宣称完成。 |
| [header-footer-field.ts][field-render]、[header-footer-field.spec.ts][field-tests]、[note-numbering.ts][field-numbering] | PAGE 用逻辑页码，NUMPAGES 使用完成后的物理页数、未完成时保留缓存；数字字段被锁定时仍动态显示，不改源缓存。有序 ARABIC、ROMAN/roman、ALPHABETIC/alphabetic、Upper/Lower 与引号格式可覆盖前一结果；字母编号 28 为 BB/bb。本地复用自有 numberingOrdinal，并对未知 switch、picture、额外参数保留整个原缓存。固定 renderer 把完整数字放在第一个 cache glyph、清空后续 glyph，本地按原格式跨度迁移，内部绘制方式不同。 |
| [InsertHeaderFooterFieldCommand][field-insert]、[RefreshHeaderFooterFieldsCommand][field-refresh]、[date-time-field.ts][field-date-utils]、[refresh-header-footer-fields.command.spec.ts][field-refresh-tests] | 上游数字插入写真实 instruction 和当前 skeleton 初始 cache；DATE/TIME 创建固定 Gregorian/latn 的三类 locale 模板，resolver 只精确匹配它们。刷新命令遍历正文及全部 header/footer，跳过 locked/权限不足/嵌套字段，含 selectionOnly/onOpen；按覆盖字段结果首字符 `start + 1` 的 textRun.ts.lang 优先于 UI locale。测试覆盖缓存、范围、历史、F9 与只读。本地数字插入读合法活动页；日期已扩展显式 picture 和既有解析来源语言，但仍只刷新选中未关联 story，未接全范围、onOpen、逐字段跳过及完整字段语言语义。 |
| [toolbar.shortcut.ts][toolbar-shortcuts]、[inline-format.command.ts][inline-format] | B/I/U 绑定 CTRL_COMMAND，要求 Docs 编辑器焦点；格式命令经 segmentId 获取正文或 header/footer body。本地 story TextArea 已接入相同字符格式命令链，并携带来源模型与明确 story 范围，不另造仅视觉的快捷键状态。 |

已静态阅读 [layout/__tests__/doc-skeleton.spec.ts][skeleton-tests] 中的连续/下一栏、几何兼容、节页码及平衡场景，包括 `keeps NEXT_COLUMN below the top of a preceding continuous section`、`keeps a continuous section on the current page when only its margins change`、`balances the final page of a multi-column section before a continuous section` 和显式 `pageNumberStart` 用例。文件路径与用例名称作为复查依据；本地契约按自有模型和 Canvas 坐标编写，没有运行上游测试。

数字字段对照还静态读取 `resolves numeric fields without changing cached glyphs and preserves pending values` 与 `keeps locked numeric page fields live and preserves other locked fields` 两个固定上游用例。数字显示、格式跨度迁移及 Office 重开效果均未执行验证；MERGEFORMAT/CHARFORMAT 只识别其不改变数值的部分，不能据此认定指令首字符样式语义已完整实现。

DATE/TIME 的新增 picture 实现按 Word 字段语义展开本地算法草稿。已静态读取 [Microsoft 的字段插入、编辑与显示说明][word-fields]，可确认字段名称、指令、专用与通用开关以及日期/时间格式属于字段结果契约；该说明中的 `format-field-results` 旧相对链接返回 404，Context7 的 Open XML 文档查询没有匹配到 picture 细节。这份公开页面不能单独证明全部 marker、转义、locale 或大小写语义；相关算法/fixture 仍须补 Office 规格与文件样本核对，不能冒称已完整复刻。

## 2. 正文排版的实际接线

入口仍为 [layoutDocument](../../packages/office-studio/src/features/document/layout/document-layout.ts)。解析 DTO、领域模型、布局与 DrawCommand 保持分离；没有用 antdv-next Table 或整页 contenteditable 替代正文。

| 模块 | 本轮真实行为 |
| --- | --- |
| [document-layout-types.ts](../../packages/office-studio/src/features/document/layout/document-layout-types.ts) | 布局页保存物理 index、逻辑 pageNumber/pageNumberStart、sectionPageIndex、方向、节片段、栏与 isParityBlank。正文行、单元格、对象和 marker 带 sectionId/sectionFragment/column。原 `document-layout.ts` 类型导出路径保留。 |
| [document-section-flow.ts](../../packages/office-studio/src/features/document/layout/document-section-flow.ts) | 按有效节设置建立真实纸张与栏。文本/表格/对象 consume 当前栏，耗尽后进入剩余栏，所有栏耗尽再换页。页 y 按前页真实高度累计。显式分页进入物理下一页，显式分栏优先剩余栏。 |
| [document-paragraph-layout.ts](../../packages/office-studio/src/features/document/layout/document-paragraph-layout.ts)、[document-text-layout.ts](../../packages/office-studio/src/features/document/layout/document-text-layout.ts) | 段落使用当前栏宽、缩进与行距；跨到边距不同的新页时只重排未放置后缀，保留绝对 UTF-16/grapheme 索引，首行缩进与列表 marker 只出现一次。保留既有 CJK 字体槽、LaTeX、换词及 checklist 路径。 |
| [document-table-layout.ts](../../packages/office-studio/src/features/document/layout/document-table-layout.ts) | 表格行按当前栏测量；溢出后栏宽不同则重测该行。当前整行不可拆，超高行在完整栏放置一次，避免无限翻页；没有完成跨页拆行、合并单元格或复杂嵌套表格增强。 |
| [document-column-balance.ts](../../packages/office-studio/src/features/document/layout/document-column-balance.ts) | 在需要平衡的连续节转接前，仅重排完整、等宽、纯文本片段，保留顺序、索引、glyph、baseline、marker 与末行之后的 spaceAfter。显式分栏、表格、图片、divider、满末栏等场景跳过；不是完整通用分栏平衡器。 |
| [document-page-arrangement.ts](../../packages/office-studio/src/features/document/layout/document-page-arrangement.ts) | 按最大页宽居中不同尺寸的物理页，同步平移正文、栏、选区相关坐标、表格、图片和分隔线；索引各页内容与栏分隔线。总高度来自最后一页的真实几何。 |
| [document-page-stories.ts](../../packages/office-studio/src/features/document/layout/document-page-stories.ts) | 第一物理正文页的 story 归属节保留；同页连续节不覆盖旧页所属节。PAGE 使用逻辑页码，NUMPAGES 使用物理总页数，首页按 sectionPageIndex，偶数页按逻辑标签。补页没有正文或 story 行。 |

连续节在尺寸和方向兼容时可共页，并追加自己的节片段。边距变化时当前物理页保持既有内容区域，后续新页使用新节边距；栏数/宽度/间距变化走相应片段及平衡判断。下一栏要求尺寸、方向与边距兼容，并保留先前片段顶部、跳过已使用栏；不兼容或没有剩余栏时建立新页。

显式页码重启与未设置重启分开：未设置延续之前的逻辑标签；同页连续节的重启用于后续页，不改写旧物理页标签。奇偶分节按逻辑页码补页；下一页重启在启用奇偶 story 的场景也处理物理/逻辑奇偶匹配。补页标记为 `isParityBlank`，没有节正文片段。显式分页已产生的空页可按边界规则复用，连续两次分页仍保留两次换页。这些边界仅有草稿，尚未实际验证。

[document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts) 在应用前以 prospective 模型校验有效尺寸、边距、栏间距与可用内容宽度，非法组合不会先写入局部设置。首节更新 legacy page 时处理存量几何覆盖，并校验其他依赖默认纸张的节。`documentEffectiveSectionSetup()` 保留显式 pageNumberStart 与 UI 默认显示 1 的区别。

## 3. 六类独立关联与旧快照兼容

模型增加 [DocumentSectionHeaderFooter.storyLinkToPrevious](../../packages/office-studio/src/domain/model.ts)：

```ts
storyLinkToPrevious?: Partial<Record<DocumentPageField['story'], boolean>>
```

六类 key 为 header、footer、firstPageHeader、firstPageFooter、evenPageHeader、evenPageFooter。缺项回退旧 `linkToPrevious`。读取或按 map 修改时先补全六项，再聚合 legacy 布尔值：六项全部 true 才为 true，混合态为 false。补全必须先于聚合，否则缺项会错误地跟随混合态 false。

- 旧 `patch.linkToPrevious` 继续表示整组六项切换；同次 `storyLinkToPrevious` 按具体 key 覆盖。Ribbon 将其单个开关意图转换成当前 story 的 map patch。
- [section-header-footer.ts](../../packages/office-studio/src/features/document/services/section-header-footer.ts) `resolveSectionHeaderFooter()` 按节序解析每类 story 的前节链，并返回深克隆。后续节没有配置时保持旧全关联语义；明确配置的首页/奇偶选项保持本节状态。
- 取消某类关联将当前有效的字面文本、runs、paragraphs、locks、字段及混合缓存保存为独立副本。已经独立的其他类不被覆盖；仍关联的其他类继续追踪前节更新。
- 重新关联只移除选中 story 的本地文本/富内容/字段 override，读取时动态解析前节。存储的关联 story 字符串为空不代表其显示内容为空；展示、输入、布局和文件转换必须使用有效继承状态。
- [document-section-stories.ts](../../packages/office-studio/src/features/document/model/document-section-stories.ts) 提供 `sectionStoryLinkedToPrevious()`、完整 map、单 story 替换和语义比较。未知 key、非 boolean map、非法对象和第一节 true 关联在写入前拒绝。关联状态为 true 时直接修改该类内容/字段被拒绝；未改变的其他继承 story 可以随现有编辑 patch 回传。
- 所有 rich 内容、字段缓存和锁校验以及 `assertDocumentStoryLocksPreserved()` 都在最终发布前完成。重新关联不能绕过当前内容锁；多类 patch 失败不发布部分状态。旧布尔值与部分 map 的读写迁移、第一节、深复制及锁相关草稿尚未执行。

本地复用 `doc.command.set-section-header-footer` 的可序列化 patch 和内容历史，没有新增上游独立命令 ID 或采用其 segment 资源表；这是内部架构/API 差异。上游用资源引用 ID、引用计数与资源复制，本地用内嵌 story 与关联 map。完整上游 Facade/command alias、Office story 关系与资源去重仍须单独核对。

根任务负责 `set-section-header-footer` 的实际 story/entity 权限判定，以及 Rust/parser/writer/DTO 的 map 映射与 Office 草稿；本文件不将这些共享在途改动记作本轮已验证能力。相关字符格式和来源证据见 [DOC-02](DOC-02-clear-formatting.md)，公共输入/历史见 [COMMON-03](COMMON-03-input-selection-axis-history.md)。

## 4. Canvas、工具栏、面板与字段刷新

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 Ribbon checkbox、内容禁用、输入只读、格式与 clear 状态均读取活动 story 的关联状态。[document-story-formatting.ts](../../packages/office-studio/src/features/document/services/document-story-formatting.ts) 与 [document-clear-formatting.ts](../../packages/office-studio/src/features/document/services/document-clear-formatting.ts) 在命令服务内也按请求 story 判断，不能通过混合态 legacy false 编辑仍关联的另一类内容。

[document-header-footer-capabilities.ts](../../packages/office-studio/src/components/document/document-header-footer-capabilities.ts) 被 Ribbon 根声明、compact trigger、control 和 Workbench 共用：导航只读取 story；位置使用节编辑权限；内容插入/刷新使用 story 编辑权限且拒绝关联类；关联与首页/奇偶选项使用 story 编辑权限。只读页眉可以转到页脚，单个 story 被拒绝编辑时仍可按节权限调整页眉/页脚位置；旧 state 缺少 sectionDisabled 时保留 disabled 回退。compound options/position 请求先判断两类权限，UI 实际控件分别发送各自 patch。

[DocumentHeaderFooterRibbonControl.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterRibbonControl.vue) 的 Button/Menu 项、日期 Modal OK、Select/Checkbox 与位置 InputNumber 显式传递对应 antdv-next disabled。HeaderFooterRibbonState/HeaderFooterAction 增加可选 requestScope，仅用于 UI 请求；Workbench 总是提供当前版本。模型、API、story 范围、输入 marks 和生命周期变化会失效该版本；页码菜单/日期弹窗保存打开时版本，范围失效、关闭或权限改变后不能回退成新目标请求。Workbench 解析动作时拒绝非法结构/参数，并在公开选择和输入会话回调后重查来源模型、API、节/variant、物理页与范围；即使新的 DOM 选区尚未发出 select，同步新范围后也拒绝旧请求，不重新捕获新版本授权插入。没有把版本 token 写入文档或更改公共命令 ID。

旧 [DocumentHeaderFooterPanel.vue](../../packages/office-studio/src/components/document/DocumentHeaderFooterPanel.vue) 保留为设置控件，改为默认、首页、偶数页各自 header/footer 的独立 Switch、TextArea 和 PAGE/NUMPAGES 按钮；可见变体由本节选项决定。开关与字段按钮按本节和对应 story/entity 权限判断；清除只作用于允许编辑且未关联的 story，保留其他类字段。字段事件参数扩展为完整 `DocumentStoryKey`。面板移除原 500 字符限制，使用 Form/FormItem、Input、Input.TextArea、Switch、InputNumber、Button、Modal，没有新增原生表单控件。

[document-page-field-result.ts](../../packages/office-studio/src/features/document/model/document-page-field-result.ts) 供 story 字符串与 Canvas 显示共用：PAGE 读取逻辑页码，NUMPAGES 读取物理总页数，按 instruction 顺序处理有限数字格式开关。不支持/未完成的 instruction 与未知页数保留原 cachedResult。`createDocumentNumericPageField()` 被上下文和旧面板插入使用，初始 instruction/cache 来自活动 story 的合法布局页，不使用虚构的固定页码；补页不能作为插入目标。节级 pageNumberFormat 尚未建模接线。

[document-story.ts](../../packages/office-studio/src/features/document/model/document-story.ts) 的 `documentStoryFieldDisplayRuns()` 先应用 modeled style/baseline，再通过 `remapDocumentStoryFieldRuns()` 将缓存运行迁移到动态显示结果。保留不变前后范围，改变范围按旧格式跨度分配，新增长度放到最后文本跨度；保留链接、来源 metadata 和空实体，并避免拆开 UTF-16 代理对。结果未变也深复制返回值，渲染不修改源 cache 或生成内容历史。每个字段显示可扩成多个 glyph，但输入仍为一个 U+FFFC 对象。

[document-story-fields.ts](../../packages/office-studio/src/features/document/services/document-story-fields.ts) 的 DATE/TIME 手动更新共用多格式迁移，重新建立与 cachedResult 一致的 cache，并保留 instruction、offset、来源 metadata、样式基线、链接与原子输入对象。它调用 [document-date-time-field.ts](../../packages/office-studio/src/features/document/services/document-date-time-field.ts) 的 `resolveDocumentDateTimeFieldResult()`，不再要求 instruction 与当前语言生成的三模板完全相等。`storyDateField` 保留原导入路径，改为同模块 `createDocumentDateTimeField` 的别名，仍返回 DATE/TIME、必需 instruction 和 cachedResult；没有改共享 DocumentPageField、run DTO 或 Rust/WASM 协议。

显式 picture 草稿区分 d/dd/ddd/dddd、M/MM/MMM/MMMM、yy/yyyy、h/hh/H/HH、m/mm、s/ss 和 AM/PM/am/pm；月份 M 与分钟 m 分开，12/24 小时独立。支持单引号字面量、双单引号、反斜线转义、外层字段参数中的转义双引号，以及引号/无空格的非引号 picture。星期和月份名称经 Intl 的 Gregorian/latn 按所选 language 格式化；有日数字时使用日期上下文的月份词形。按顺序处理 Upper/Lower，MERGEFORMAT 共用现有缓存格式迁移。非法日期/locale、未知 ASCII mask、未知开关、重复 picture、无 picture 的平台默认以及 CHARFORMAT 保留原缓存，不删除字段；插入非法参数仍明确拒绝。

日期刷新先读取第一条非空缓存 run 的 `metadata.source.style.language`，缺少时读字段的 `metadata.source.textStyle.language`，再回退调用方 locale；不从后续缓存 run 挑另一个语言。空 run 不决定语言；全空缓存扩展出文字后，仍指向空来源的 provenance 及其清除 marker 也不成为下次刷新语言权威，避免重复刷新改变语言。字段或首个有效结果 run 带 `officeStoryTextFormatCleared` 时，忽略旧来源语言，使用调用方 locale；非法/未支持的来源语言保留原缓存，不静默改用有效 UI 语言。上述只消费已有 provenance，没有修改共享 DocumentPageField、run DTO 或 Rust/WASM 协议，也没有删除源语言数据。

解析语言的静态来源为 `duke-pages@0.6.0/src/model.rs::TextStyle.language`、`src/parse/style.rs` 中的 `w:lang/w:val` 读取及 `src/parse/util.rs` 的语言继承；本地 [documentNormalizer.ts](../../packages/office-studio/src/parser/documentNormalizer.ts) `normalizeRun()` 将原始 run 保存到 metadata.source，因此语言路径在 source.style 内。字段 style 的回退与空缓存处理是本地缓存重建策略，不冒称上游使用相同 metadata 结构。语言尚未成为可编辑的共享样式字段，新插入字段的语言写出、East Asian/Bidi 语言槽和指令首字符样式仍未接线。AM/PM 的平台文本/大小写、完整 Office masks、默认日期/时间格式、Hijri/Saka/Um-al-Qura 等日历开关、CHARFORMAT、复杂/nested 字段与未知 story XML 继续待实现或待核实。锁与权限仍由既有最终 patch 保护，不能把本地整个 story 的事务拒绝认定为上游逐字段跳过已完成。多运行迁移是本地策略，尚未通过独立 Office 样本确认最终交互效果。

[document-story-navigation.ts](../../packages/office-studio/src/features/document/layout/document-story-navigation.ts) 按物理优先页、sectionId 和 variant 寻找合法 story 页，排除布局页及 story entry 的 isParityBlank。Canvas 命中、story 样式、reveal、Workbench 打开入口、公开输入状态恢复、重排定位与 [use-document-story-input.ts](../../packages/office-studio/src/components/document/use-document-story-input.ts) 的输入 scope/捕获/预览均拒绝补页。节选择退出 story 编辑后定位正文。

首页/奇偶选项改变当前可见 variant 时，Workbench 在新布局发布后重定位，并以文档 ID、模型身份、API 实例、epoch 与 revision 防止旧异步范围恢复覆盖新输入；公开输入 scope 同步到新的可见 story。重排使旧物理索引落到补页或另一节时，按先前 section/story 找回合法页。真实事件顺序、重入、快捷键、焦点、主题与 IME 仍待浏览器验证。

[DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 的正文和 story 输入继续使用 antdv-next Input.TextArea。story 的 Ctrl/Command+B/I/U 接到现有 toggle-format→doc.command.format-text；先同步真实 DOM 范围，再由 Workbench 核对来源模型、明确 story、会话与范围。只读、关联、锁与 IME composition 仍由同一能力/命令链判断；toggle-format 内部事件增加可选来源模型/story 参数，正文既有单参数调用保留。正文、页眉页脚与栏分隔线由 [document-scene.ts](../../packages/office-studio/src/features/document/rendering/document-scene.ts) 生成共享 DrawCommand，Canvas2D/WebGPU 使用同一布局；补页跳过 story 绘制。此处没有新增 GPU 管线或性能优化验收。

## 5. 待执行验证草稿

| 文件 | 预期覆盖及边界 |
| --- | --- |
| [document-section-layout.contract.test.mjs](../../packages/office-studio/tests/document-section-layout.contract.test.mjs) | 多栏溢出、显式分页/分栏、continuous/nextColumn、边距后缀重排、平衡及 trailing gap、不同纸张坐标、表格整行、逻辑页码、story 选择、空白补页、分隔线和非法设置。真实字体、图片/复杂表格/全部奇偶分节、Worker/性能未全面覆盖。 |
| [document-section-layout-api.contract.test.mjs](../../packages/office-studio/tests/document-section-layout-api.contract.test.mjs) | 公共节设置命令→布局→撤销/重做→内部保存/重开，非法几何、权限/取消与文档隔离。fixture 的 OfficeStudio 历史不替代生产应用 adapter。 |
| [document-section-story-links.contract.test.mjs](../../packages/office-studio/tests/document-section-story-links.contract.test.mjs) | 六类逐项关联/取消、前节传播、部分 map 与 legacy 迁移、完整富内容/锁/cache 深复制、关联编辑拒绝、原子锁保护、格式/clear、字段缓存刷新和补页导航。新增字段对象范围受 content lock 保护时拒绝缓存变更及整份设置 patch 的草稿；结果未变时允许独立选项变更。 |
| [document-section-story-links-api.contract.test.mjs](../../packages/office-studio/tests/document-section-story-links-api.contract.test.mjs) | 六类公共命令的内容历史和内部重开；story/entity 权限、继承内容原样回传、非法 map、第一节、取消与非活动文档隔离。包含 format-text→动态 PAGE 布局→undo/redo→内部保存/重开的 mixed cache/link 草稿，以及 DATE/TIME 自定义 picture→选中独立 story 更新→无变化不追加历史→undo/redo→内部重开/Canvas 字符串，保留源 story、footer、未知字段与混合缓存链接。另有法语来源优先→保存重开后再刷新→公共 clearFormatting→忽略旧来源语言→undo/redo→第二次内部重开与 Canvas 字符串草稿，保留来源 metadata、链接和原子输入。命令权限依赖根任务共享实现；尚未运行。 |
| [document-story-field-results.contract.test.mjs](../../packages/office-studio/tests/document-story-field-results.contract.test.mjs) | 19→20 混合颜色/强调/链接，Unicode 与空实体/深复制，动态显示与缓存刷新一致；PAGE/NUMPAGES 多 glyph 的原子输入范围、数字开关顺序/未知回退/初始 cache、锁定数字仍动态显示、Canvas 字符串一致及源模型不变。未覆盖任意掩码、CHARFORMAT 样式或 Office 重开。 |
| [document-date-time-field.contract.test.mjs](../../packages/office-studio/tests/document-date-time-field.contract.test.mjs) | 日期/时间 marker、M/m、午夜/正午、字面量/引号/转义/Unicode、月份词形和语言大小写、三预设生成/解析一致；非法 locale/date/instruction 与未知开关/cache 回退；选中未关联 story 的混合缓存迁移、其他 story/未知字段/输入对象不变。包含解析 DTO→normalizeDocument→来源语言优先、空 run/全空 cache 与重复刷新、字段来源/UI 回退、非法来源语言保持 cache，以及 field/首 run clear marker 的语言回退草稿。平台默认、可编辑语言/新字段写出/原生锁、完整日历与 CHARFORMAT 未覆盖，所有用例未执行。 |
| [document_story_roundtrip_tests.rs](../../crates/office-parser/src/document_story_roundtrip_tests.rs)、[docx.rs](../../crates/office-writer/src/docx.rs)、[parser.rs](../../crates/office-parser/src/parser.rs) | 新增 DATE/TIME 显式指令、转义双引号、两轮修改后缓存/颜色/强调/链接/未知 DATE 开关及 Tail 内容锁；DATE 首缓存 run 的 native w:lang/fr-FR 在 parser→领域 provenance→两次 DOCX 写出/重读后保留。dormant 六类 story 在关闭→首页启用→奇偶启用后真实 ZIP 写出/parser 重读，保留富内容和 DATE/TIME cache。旧 writer part 预期同步到 first/even 内容保留后的 header9/header11，parser 读取关闭状态的 first/even 引用。helper 明确组装领域协议，日期值是模型边界 fixture，未实现第二个 Rust 日期 formatter，未调用 TS normalizer/WASM/独立 Office，也未执行。 |
| [document-header-footer-capabilities.contract.test.mjs](../../packages/office-studio/tests/document-header-footer-capabilities.contract.test.mjs)、[document-header-footer-ribbon.contract.test.mjs](../../packages/office-studio/tests/document-header-footer-ribbon.contract.test.mjs) | readonly/linked/section 权限矩阵、legacy 回退、compound patch、捕获/失效请求范围及非法动作拒绝；同一根声明向 classic/grid/overflow 传递一致能力。没有实际挂载全部 UI 呈现。 |
| [document-section-layout.browser.test.mjs](../../packages/office-studio/tests/document-section-layout.browser.test.mjs) | Vue/ConfigProvider/Workbench/FDocument 下的第二栏输入、键盘历史、混合纸张点击、theme/grid、补页过滤、独立关联、首页 variant 重定位及内部重开。追加 mixed PAGE glyph→B/I/U→undo/redo/重开、只读/关联与合成 composition key 的格式拒绝、真实初始 cache、readonly header→位置/可编辑 footer、窄屏 classic More、date/page 旧弹层及同步选择观察者后的旧日期/快捷键不追加编辑的草稿。Edge headless 未启动；合成事件不是 Windows 原生 IME、生产 history adapter 或 Office 互操作证据。 |

所有草稿保留“未执行”。没有测试通过数量、浏览器结果、WASM 产物或实测性能数据。本轮仅使用静态读取、编辑、rg 和 git status/diff；未执行任何验证命令。

## 6. 尚未完成与未验收边界

- 长页眉页脚尚未反向预留正文空间，内容可能与正文相挤；跨页拆表格行、合并/嵌套表格、复杂对象环绕和完整通用分栏平衡仍未完成。
- 正文中嵌入的分页/分栏符、复杂字段/未知 story XML、节级 pageNumberFormat、有限开关之外的页码格式/picture、CHARFORMAT 指令首字符样式，以及 DATE/TIME 平台默认、日历开关、完整 masks、可编辑字段语言/新字段语言写出/East Asian 与 Bidi 槽和原生字段锁仍待实现或核实。动态 PAGE 和日期缓存的多运行迁移、既有来源语言选择已有实现与草稿，Office 格式效果尚未核对；DATE/TIME 全 segment/onOpen 与上游逐字段跳过仍未接线。
- DOCX 原始 story 关系、每类独立关联的 writer/parser 往返、WASM 与独立 Office 重开尚未验收。新增原生 DATE/TIME 指令/缓存及 dormant 六类 story 草稿仍未执行，内部 JSON 保存重开不能替代 Office 导出。原生 Word 的 evenAndOddHeaders 是文档级 settings；本地节级 differentOddEven 的跨节组合与原生激活行为尚未等价，不能从保存 six references 推断所有开关组合已保真。
- classic/grid/More、窄屏面板、全部权限与重入、焦点恢复、真实 Windows IME、真实字体度量、Canvas2D/WebGPU 一致性与设备丢失均未实际验证。
- 节布局仍在主线程 TypeScript 路径；Worker/Rust/WASM 热点迁移、真实长文档/大样本时间、帧、长任务、传输与内存没有基准。模块拆分或复用 DrawCommand 不构成性能优化证明。
- 本文件只描述 DOC-03 本轮切片；Docs、Sheets 与全部 Pro 目标不缩减。总表继续为未验收，剩余项按待开发、待修复或待验证保留。

[schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts
[menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts
[hf-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/header-footer-ribbon.ts
[hf-control]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/header-footer/panel/DocHeaderFooterRibbonControl.tsx
[hf-link]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/commands/commands/set-section-header-footer-link.command.ts
[skeleton]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/doc-skeleton.ts
[skeleton-tests]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/__tests__/doc-skeleton.spec.ts
[page]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/model/page.ts
[sections]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/utils/sections.ts
[columns]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/utils/section-columns.ts
[section-api]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/facade/f-document-section.ts
[section-settings]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/section-setting/use-section-setting.ts
[field-render]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/header-footer-field.ts
[field-tests]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/__tests__/header-footer-field.spec.ts
[field-numbering]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/note-numbering.ts
[field-insert]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/insert-header-footer-field.command.ts
[field-refresh]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/refresh-header-footer-fields.command.ts
[toolbar-shortcuts]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/shortcuts/toolbar.shortcut.ts
[inline-format]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/inline-format.command.ts
[field-date-utils]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/utils/date-time-field.ts
[field-refresh-tests]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/__tests__/refresh-header-footer-fields.command.spec.ts
[word-fields]: https://support.microsoft.com/en-us/office/insert-edit-and-view-fields-in-word-c429bbb0-8669-48a7-bd24-bab6ba6b06bb
