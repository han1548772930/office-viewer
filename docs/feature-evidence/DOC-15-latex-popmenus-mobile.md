# DOC-15 段落、浮动入口与移动公式编辑

记录日期：2026-11-26。状态：**实现与测试草稿已落盘，未验收**。续接 [公式编辑批](DOC-15-latex-editing.md) 与 [71 模板 / 上下文 Ribbon 批](DOC-15-latex-templates-ribbon.md)。本批先按固定来源清点剩余挂载，再在已有公式模型、命令、内容历史和保存协议上接线；下面的行为是实现说明，尚无运行验收证据。

按用户要求，本阶段不运行测试、类型检查、构建、dev、浏览器、Cargo/Rust/WASM、基准或 `git diff --check`。源码读取和文档记录不构成功能验收。

## 1. 实现前的固定入口清点

OSS 参考为 `E:\rust_test\univer` 的 v1.6.0，固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 来源是项目缓存的 [docs-latex-ui v1.6.0 ES](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/es/index.js)，本批只读取文本，没有执行商业包。

`W` 是普通公式入口 factory，`xt` 在 `W` 基础上增加结构 block range 显隐。它们均调用 `B.id = docs-latex.operation.open-formula-popup`，关闭 operation 为 `pt.id = docs-latex.operation.close-formula-popup`。标题 / tooltip 为 `docs-latex-ui.menu.formula`，icon 为 `FunctionIcon`。

| 自上而下的宿主 / 分组 | 上游入口 / order | factory 与状态 | 本批目标 |
| --- | --- | --- | --- |
| `ContextMenuPosition.PARAGRAPH → doc.menu.empty-paragraph → ContextMenuGroup.LAYOUT` | open popup / 8 | `W`；非 Docs hidden；header/footer editing disabled | 空段落 Layout 公式入口 |
| `doc.menu.paragraph-t.insert → quickBottom` | open popup / 4 | `W`；同上，quickBottom 是当前插入段的快捷操作 | 空段落插入菜单的快捷公式入口 |
| `doc.menu.paragraph-t.insert → insert` | open popup / 3 | `W`；同上，使用菜单捕获的文本范围 | 插入子菜单公式入口 |
| `doc.menu.paragraph-t.edit → quickBottom` | open popup / 5 | `W`；同上 | 非空段落编辑菜单的快捷公式入口 |
| `doc.menu.float-toolbar → doc.menu.float-text-style` | open popup / 13 | `xt`；另隐藏落在任一结构 block range 的文字选区 | 文字选择后的浮动公式入口 |

以上是 **5 个有效挂载点**，不是 5 个独立公式功能。浮动文字样式同时出现在 `Dt[float-text-style]` 与 `Et[float-toolbar][float-text-style]`；控制器 `Z` 的 `mergeMenu(Dt)` / `appendRootMenu(Et)` 是两条注册路径，不能重复计为两个可见浮动入口。`Tt` 的常规 Insert / Media order 8、grid row 2 / column 7 已在前两批接线，本批常规 Ribbon 根增量为 0；此前 11 个上下文 selector 和 71 个模板不重复计算。

OSS 名称与作用域来源：`packages/docs-ui/src/menu/paragraph-menu.ts` 的 `EMPTY_PARAGRAPH_MENU_ID`、`DOC_PARAGRAPH_T_INSERT_MENU_ID`、`DOC_PARAGRAPH_T_EDIT_MENU_ID`；`packages/docs-ui/src/menu/menu.ts` 的 `FLOAT_TOOLBAR_MENU_POSITION`、`FLOAT_TEXT_STYLE_MENU_ID`、`disableMenuWhenHeaderFooterEditing`、`hideMenuWhenSelectionInBlockRange`；`packages/docs-ui/src/menu/schema.ts` 的 empty/Layout 与 insert/edit quickBottom；`packages/docs-ui/src/views/ParagraphMenu.tsx` 的菜单捕获范围与执行后恢复。

上游 popup 服务 `z` 从真实 selection 读取 source，formula range 使用保存的原始 LaTeX，普通文本选区作为新公式初稿。popup open 不先写入文档。类型依据为 [formula-popup.operation.d.ts](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/types/commands/operations/formula-popup.operation.d.ts)，其显式 `startIndex/endIndex` 为 inclusive 范围，`selection.startOffset/endOffset` 为 exclusive 范围。

移动来源为 ES `Vt._captureFormulaPointer`：`DocMobileElementMenuService.capture` 提供 **编辑、删除 2 个动作**，编辑走同一 open popup，删除走 `docs-latex.command.remove-formula`；`Ht._createFormulaPopup` 使用 `docs-latex-ui.mobile-formula-dialog`，maskClosable=false。移动版与桌面是相同模型、命令及模板目录的呈现变体，不能重复增加公式类型或常规 Ribbon 根。

本地继续保留已支持的表格 / header / note host 编辑：相对上游 `W` 的 header/footer 禁用，本项目按真实 host、字段 / 控件锁和权限判断。上游 `isTextRangeInAnyBlockRange` 见固定 OSS `packages/docs-ui/src/menu/menu.ts:430`，判断的是 `body.blockRanges` 相交，不把所有原生表格一概当作 block range；本地原生表格内文字仍可使用公式浮动入口。

本批五处都不是常规 Ribbon 根；classic/grid 不新增 row/column。空段 Layout、quickBottom 和 insert 子菜单使用各自菜单位置与 order，浮动文字样式是单个扩展按钮；移动端两个动作使用同一模型的单行按钮组。动态显隐、最终菜单位置与两种 Ribbon 变体仍需运行验证。

| 本批分类 | 有效增量 | 口径 |
| --- | ---: | --- |
| 常规 Ribbon 根 / 上下文 Ribbon selector | 0 / 0 | 前两批的常规 Insert 根、1 分类 / 11 selector 不重计 |
| 段落右键 Layout | 1 | open popup / order 8 |
| 空段插入 quickBottom / insert 子菜单 | 1 / 1 | 同一 open popup，order 4 / 3 |
| 非空段编辑 quickBottom | 1 | 同一 open popup，order 5 |
| 浮动文字样式扩展 | 1 | order 13；两条上游注册路径合计一个有效入口 |
| 有效挂载合计 | **5** | 菜单与浮动挂载；不是五个独立数学功能 |
| 移动元素操作 | **2** | 编辑、删除，单列，不并入五个挂载或 Ribbon 根 |
| 模板目录 / 新增 More 根 | 0 / 0 | 复用既有 71 模板；不新增目录或溢出根 |

## 2. 实现与接口

本批属于 DOC-15 的 controller / API / view 层。公式模型、数学 parser / layout、命令及内容历史沿用前两批；没有改 Rust crate、Worker 协议或另建数学权威引擎。新增 popup target 是主线程运行时对象，包含源模型身份、unit、选区、host、初稿与版本，不进入快照或跨 Worker 传输。

| 责任 | 文件与落盘行为 |
| --- | --- |
| operation DTO | [document-latex-popup.ts](../../packages/office-studio/src/features/document/controllers/document-latex-popup.ts)：public/local open / close 常量、request / target / text-range 类型及参数校验 |
| source 范围与权限 | [API popup resolver](../../packages/office-studio/src/api/document-latex-popup.ts)、[latex-formula.ts](../../packages/office-studio/src/features/document/latex/latex-formula.ts)、[feature commands](../../packages/office-studio/src/features/document/commands/document-latex-commands.ts)：解析 public offsets / range ID / segment，复用原子边界、字段 / 控件锁和权限，读取真实 source 初稿 |
| runtime owner | [document-latex-popup-editors.ts](../../packages/office-studio/src/api/document-latex-popup-editors.ts)：按 studio / unit 隔离，替换 force-close、旧 handle 不移除新 owner，unit / studio 销毁清理 |
| API 接线 | [API commands](../../packages/office-studio/src/api/document-latex-commands.ts)、[document-latex.ts](../../packages/office-studio/src/api/document-latex.ts)、[facade.ts](../../packages/office-studio/src/api/facade.ts)、[office-studio.ts](../../packages/office-studio/src/api/office-studio.ts)、[index.ts](../../packages/office-studio/src/index.ts)：注册 operation、导出原子范围检查与 runtime 类型、Facade bind / open / close、销毁释放；共享文件只改 Docs 区段 |
| 入口与视图 | [document-latex-menu.ts](../../packages/office-studio/src/components/document/document-latex-menu.ts)、[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue)、[DocumentStructuredPopMenus.vue](../../packages/office-studio/src/components/document/DocumentStructuredPopMenus.vue)、[DocumentStructuredOverlay.vue](../../packages/office-studio/src/components/document/DocumentStructuredOverlay.vue)：五挂载、捕获目标、移动双动作、共同命令与 stale guards |
| 编辑与主题 | [DocumentLatexPanel.vue](../../packages/office-studio/src/components/document/DocumentLatexPanel.vue)、[DocumentLatexTemplateGallery.vue](../../packages/office-studio/src/components/document/DocumentLatexTemplateGallery.vue)、[Docs locale](../../packages/office-studio/src/i18n/docs/locales/document-latex.ts)：initialSource / editorKey / compact、source 重开重置、窄屏布局与英 / 简繁标签 |

注册 `docs-latex.operation.open-formula-popup`、`docs-latex.operation.close-formula-popup`，保留本地别名 `doc.operation.open-latex`、`doc.operation.close-latex`。FDocument 新增 `bindLatexPopupEditor`、`openLatexPopup`、`closeLatexPopup`；方法名是本地 UI 扩展，上游对应的是已核实的 operation ID。新增公共导出为 operation 常量、`DocumentLatexPopupOperationRequest`、`DocumentLatexPopupTextRange`、`DocumentLatexPopupTarget`、editor / handle 类型。

参数保留 inclusive `startIndex/endIndex` 和 exclusive `selection.startOffset/endOffset`，二者同时提供时检查一致性。固定 OSS `packages/core/src/sheets/typedef.ts:766` 的 ITextRange.direction 支持 none / forward / backward；也接受 ITextRangeParam 的 segmentId、segmentPage、isActive、rangeType 元数据。嵌套 / 顶层 segment 与 page 必须一致，backward 保留 anchor / focus 方向；RECT 在文字 popup resolver 中明确拒绝。省略 collapsed 是本地宽容扩展，提供时必须与范围一致。

普通选区使用实际文字，包括空格、完整 UTF-16 字符和组合字符；完整公式 atom 展开为原 LaTeX。读取先拒绝原始越界 / 已删除点，避免通用 `orderedSelection` 把错误范围夹回其他位置；拒绝跨段、表格结构选区和切开的 surrogate pair。public 范围仍先通过 `assertDocumentLatexAtomicRange`，不接受公式 source 中间的半个范围。初稿包括展开后的 source 使用 65,536 code units 预算。

打开 / 关闭只改变编辑器和输入位置。真正写入仍是既有 insert / update / remove / replace-with-text 内容命令，复用权限、单一历史和保存链路；打开不先创建段落或公式。quickBottom 使用捕获的当前文本范围，本批没有增加 after-block 内容 mutation。无 owner 时 open 返回 false，close 可返回 true；编辑器可拒绝普通 close，例如 composition 期间。

Workbench 的渲染模型经过 `normalizeDocument` 克隆，不能拿它与 API 权威模型作对象相等判断。本批用 `api.getModel() === target.source` 守护 API 身份，另捕获当前 UI 快照；nextTick 前后同时核对 unit、revision、history generation、input generation、host、选区、owner、active 与 IME 状态。选区比较复用 `sameDocumentInputLocation`，不靠对象属性排列。停用 / 卸载 force-close 并递增 opening generation，防止短暂重新激活让旧 queued open 复活；关闭命令异常保留 draft 并显示反馈。

## 3. UI、输入和移动边界

全部可替换控件使用 `antdv-next`，复用统一 ConfigProvider / token。Canvas 继续负责内容、命中和选择；source 输入沿用 Ant Input.TextArea 及最小原生 selection 桥。

菜单采用受控 Dropdown / Menu，移动菜单采用 Drawer，浮动动作使用 Button / Tooltip；既有代码控件继续使用 Popover / Ant Form。此前已通过 Context7 `/antdv-next/antdv-next`，请求版本 1.8.2 核对 controlled open、popup container 和 destroyOnHidden 接口，本批没有安装 UI 依赖或修改锁文件。

移动公式动作、结构 toolbar 和段落按钮放进同一 flex 停靠布局，按实际行高排列，避免三处都固定在 bottom 8px。公式 Modal 使用窄屏宽度、可滚动 body、大尺寸页脚按钮和 maskClosable=false；与前批一致保留非阻塞 Modal，让顶部 11 类模板仍可点击。小屏模板排两列，目录保持同一 71 项。

source 使用 Ant Form / Form.Item / Select / Input.TextArea，preview 仍为 Canvas。原生操作仅用于 TextAreaRef 暴露元素的 selectionStart / selectionEnd / setSelectionRange：组件没有公共 source-range setter，此桥不另建 textarea 或持久化内容。composition / draft generation、readonly 与 focus 恢复沿用同一编辑会话。已静态查看本批 UI，未引入原生表单控件或其他 UI 库；主题、焦点、真实 IME 与移动尺寸没有运行验收。

本地浮动公式在非空单段文字选区显示，完整 atom 使用公式编辑上下文；已识别结构节点 / 保留 block-range source 时隐藏。普通原生表格文字继续可用。与上游只按 body.blockRanges 判断相比，本地列组采用结构上下文隐藏；range 型 SDT / 未知结构组合、辅助 segment 的最终显隐仍待核实。浮动文字样式本批仅补 LaTeX 扩展，不冒称已补齐完整浮动字体格式工具栏。无 API 的 standalone 继续支持当前选区 / 公式 ID 初稿，显式 public offsets / segment 由 FDocument 路径处理。

## 4. 测试草稿与验收

| 本批新增草稿 | 静态 test 声明 | 计划展开用例 | 覆盖意图 |
| --- | ---: | ---: | --- |
| [document-latex-popup.contract.test.mjs](../../packages/office-studio/tests/document-latex-popup.contract.test.mjs) | 12 | 42 | 五挂载 / 同命令、真实 source 与 atom、Unicode / 越界 / 跨段 / 预算、完整范围参数及错误 |
| [document-latex-popup-facade.contract.test.mjs](../../packages/office-studio/tests/document-latex-popup-facade.contract.test.mjs) | 16 | 20 | public / local / Facade、方向与 source offsets、owner 隔离 / 替换 / 销毁、revision / history / input guards、锁 / 取消、header / note、真实 apply / undo / redo / save |
| [document-latex-popmenus.browser.test.mjs](../../packages/office-studio/tests/document-latex-popmenus.browser.test.mjs) | 14 | 18 | 五实际入口与保存闭环、同屏重开、关闭失败保留 draft、queued open 停用、observer / IME、移动双动作 / 短屏尺寸 / 主题、host 与内容锁 |
| 合计 | **42** | **80** | 全部未运行；计划展开不是收集 / 通过结果 |

source / menu 的参数声明分别展开 5 / 3 / 2 / 5 / 4 / 17 用例，加 6 个单项，共 42；Facade 两处三项参数声明使 16 条展开为 20；浏览器五入口参数声明使 14 条展开为 18。单个 test 内的场景循环不额外算用例。

复用 [document-latex.html](../../packages/office-studio/tests/fixtures/document-latex.html)，新增 `?entrances` 空段 / 普通文字 / quote / table 文字样本，以及 fixture-only open / close / range、queued bounce、observer 导航与 close 取消 hooks；用 KeepAlive 提供真实停用 / 激活路径。默认公式、invalid / locked / auxiliary / standalone / grid 场景保留，但整个夹具尚未重新运行。

更新旧 [document-latex-editing.contract.test.mjs](../../packages/office-studio/tests/document-latex-editing.contract.test.mjs) 的 context edit 预期为已核实 public open operation；这是旧声明的假设更新，不计入本批 42 条。前批 43 条、模板批 47 条 / 117 计划用例与本批分开统计，原有 7 条也不合并成本批新增。

本批只做源码 / 实现文本阅读、声明计数、已知来源与路径整理；没有运行 Vue / TypeScript 解析器、test runner、构建器、dev、浏览器、Rust/WASM 或基准，也没有执行 `git diff --check`。草稿中的合成 composition 仅准备 guard 回归，不证明 Windows 原生 IME。实际菜单装配、Office 文件重开、明暗主题 / 焦点 / 输入法和真实性能都保持未验收。

## 5. 剩余目标

完整 TeX / shaping / 数学字体、宏环境、原生 DOCX / OMML、数学 Worker / Rust / WASM 和取消 / lease / 复制预算仍在完整范围内。真实主题、焦点、Windows IME、浏览器交互、Office 样本及大样本性能尚未验收。本批不改中央清单、后端或 package / lock。

跨 source 多样式、字段 / 评论 / SDT / 复杂嵌套、note decoration、共享 header aliases 与重开资源冲突尚需统一验证；本批五入口接线不替代这些验收。当前 source seed / 数学布局仍走主线程 TypeScript，不声称已优化、Worker 零拷贝或大文档达标。

前端 popup 不要求新增服务端。OMML 解析 / 写出、协作版本中的 source / named resource 同步等后置需求继续由主任务保留在 [后端需求](../BACKEND_REQUIREMENTS.md)，沿用本项目解析 / 写出链路，不依赖 Univer Server / Exchange。

## 6. 向主任务交付

中央 Docs / 工具栏 / 顶部清单可关联本记录，登记常规根增量 0、五个有效菜单 / 浮动挂载和移动 2 动作，保留此前 1 分类 / 11 selector / 71 模板，以及本批 42 声明 / 80 计划用例的独立口径。COMMON-03 / COMMON-05 的输入与 runtime owner 边界发生接口扩展，不能据此验收其他功能组。

本批实现与草稿到此收尾，不自动开下一批；状态为未验收，保留完整 TeX / Office / Worker / UI / 性能目标。
