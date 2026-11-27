# SHEET-08 剪贴板 UI、入口能力、宿主保护和最近粘贴浮层

证据批次：2026-11-22。状态：实现与测试草稿已落盘，**未验收**。

本记录覆盖 `ExcelWorkbench`、共享内容别名与入口能力、Ribbon/右键菜单、Canvas 原生剪贴板事件、系统剪贴板适配、异步宿主保护及最近粘贴选项。领域 planner、命令/API、内容历史、资源移动和 Office 读写仍按 [SHEET-08 功能清单](../UNIVER_FEATURE_PARITY.md)验收。完整 Sheets / Docs / Pro 范围继续保留。

行为验证按统一安排后置，本文的测试草稿均未运行。源码静态核对、声明清点及主线统一语法/引用检查不作为测试、类型检查、构建、浏览器、Rust/WASM 或性能通过证据。下文的草稿数量不代表通过数量。

## 1. 固定来源和入口清点

Univer 固定版本为 v1.6.0、提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地来源为 `E:\rust_test\univer`。全工作台的分类、Pro 来源和未核实项沿用 [顶部入口清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)及 [工具栏对照](../UNIVER_TOOLBAR_PARITY.md)。本表只清点剪贴板子集，按 schema 声明区分根、容器和子项，未运行 factory 的显隐或装配。

| 分类 | 固定源码声明和顺序 | 分组或位置 | 命令、状态及来源 |
| --- | --- | --- | --- |
| 常驻 Ribbon | 1 个粘贴根，5 个子项：粘贴、值、格式、列宽、边框以外内容 | START / HISTORY，classic `order: 4`；grid `row: 1, column: 3, rowSpan: 2, showLabel: true` | `sheet.menu.paste` → paste、paste-value、paste-format、paste-col-width、paste-besides-border；根受 Sheets 实例、exclusive range 和 shape text editor 上下文控制。子项检查剪贴板与对应写值/样式/列宽权限。[schema][sheet-schema]、[factory][sheet-menu] |
| 主区、行头、列头 QUICK | 每处 3 个根，合计 9 个声明：复制、剪切、粘贴，`order: 0/1/2` | `quickLayout: tile` | 复制检查 copy/view 权限，剪切另检查编辑权限；粘贴检查剪贴板与范围、工作表、工作簿编辑权限。[schema][sheet-schema]、[factory][sheet-menu] |
| 主区、行头、列头 FORMAT | 每处 2 个容器，合计 6 个根声明：选择性复制、选择性粘贴，`order: 0/1` | `sheet.menu.copy-special`、`sheet.menu.paste-special` | 基础粘贴每处 4 个子项，`order: 0..3`；容器受 Sheets/exclusive range 显隐控制。[schema][sheet-schema]、[factory][sheet-menu] |
| 公式插件扩展 | 每处为上述容器各增加 1 个子项；3 个上下文合计 6 个扩展子项 | 仅复制公式 `order: 0`、仅粘贴公式 `order: 4`；该 schema 未向 `sheet.menu.paste` 注入公式子项 | `sheet.command.copy-formula-only`、`sheet.command.paste-formula`；copy/view 或 value/edit 权限。[公式 schema][formula-schema]、[factory][formula-menu] |
| 最近粘贴浮层 | 1 个浮层触发入口、6 个选项，顺序为全部、值、格式、列宽、边框以外内容、公式 | `pasteOptionsCache.target.pastedRange` 末端及对应 viewport | `sheet.command.optional-paste` 接 `{ type }`；`showMenu$` 和缓存控制显示，新旧模式都复查权限。[popup][clipboard-popup]、[命令][clipboard-command] |
| More 溢出 | 无新增独立剪贴板根或命令 | 同一 Ribbon 根的布局呈现 | 全量溢出清点见 [工具栏对照](../UNIVER_TOOLBAR_PARITY.md)。不重复计算为功能 |

固定上下文清点合计为 15 个根声明、18 个子项声明，其中 6 个子项由公式插件扩展；不把这些重复位置换算为独立功能数。浮层与 Ribbon 分开计数。该子集不能证明全部 Pro 或运行态入口已经完成清点。

本地现有 Ribbon 是 1 个 split/menu 根、6 个子项，在基础粘贴之后加入公式模式；主区、行头、列头和 corner 复用每处 5 个根/容器、6 个子项，合计 20 个根/容器声明、24 个子项声明。这些是本地声明数。默认粘贴禁用、但特殊模式许可时，根使用纯 menu，保留许可模式的访问；默认可用时使用 split。公式进入 Ribbon、corner 复用及根类型切换属于本地实现差异，不冒称固定版具有相同位置或内部实现。

## 2. 入口到模型的接线

`ExcelWorkbench.vue` 的 Ribbon 和右键首先通过共享 `resolveSheetClipboardCommand`，然后调用 `useSheetClipboard`。Ctrl/Cmd+C/X/V 保留浏览器默认动作，由 Canvas copy/cut/paste 事件进入相同 controller；Shift+V 标记下一次原生粘贴为 values。输入框、TextArea、contenteditable、composition 和 IME keyCode 229 继续由文本输入桥处理。画布与公式栏的运行保真输入依据另见 [SHEET-05 输入证据](SHEET-05-native-rich-text-input-2026-11-22.md)。

取数范围使用 `activeSelectionIndex` 对应范围，缺省取最后范围；不使用多选区包围矩形。固定来源的 copy / capturePasteTarget 使用 `getCurrentLastSelection()`；本地保留共享选区模型的显式 active range，差异由浏览器草稿追踪。[service][clipboard-service]

共享 resolver 支持本地 canonical 命令及固定版单数别名。粘贴参数由 `sheetClipboardPasteMode(mode?, value?, commandMode?, type?)` 校验：普通 paste 的 hook `{ value }`、optional 的 enum `{ type }`、本地 `{ mode }` 及专用命令必须一致。UI 保留已校验的 `type` 传给 API；standalone 采用相同映射。无效字符串、非字符串或冲突参数在系统读取和内容发布前拒绝。

`sheet-clipboard-input.ts` 导出的 `normalizeSheetClipboardContentAliases(input)` 由 UI 与主任务 API 参数解析共用，返回 `{ text?: string; html?: string }`。支持 `text/textContent`、`html/htmlContent`；同义字段相等可同时提供，不同值或非字符串拒绝。空字符串和空白保留，`undefined` 表示省略；显式空内容不触发系统读取或借用旧内部缓存。该函数不处理模式、目标、文件或媒体参数。

生产宿主执行 `OfficeStudio` 的共享命令，controller 不另建内容 undo 栈。optional 使用领域保存的原始请求 `target` 重规划，浮层使用 `recent.ranges` 的实际末端；`target` 与结果范围分离，避免筛选压缩/扩展后再次平铺到错误范围。生产内容历史通过同一 entry 的 amendment 保留最初 undo，最终 redo 由所选模式替换；其 API/历史契约与 UI 草稿均未执行。

standalone 只创建瞬态 Pinia clipboard store，模型解析、范围资源和公式仍走共享 planner，使用既有 view 适配和宿主发布。optional 恢复最初受影响的 sheets、metadata 和 typed workbook resources，再从原 target 重规划。受影响 sheets 跨模式取并集，metadataChanged 累计；比较最终候选与当前模型，允许“新模式对原始模型无差异，但需撤回旧模式”的发布。外部 consumer 每次收到真实内容 change；没有外部历史 amendment callback，不能宣称其外部 undo 也只增加一条。

共享 `assertSheetClipboardWorkbookChangesEditable(current, candidate, userId)` 已由 planner/API 转导出，standalone optional 在 restore+apply 之后、view 准备和发布之前复查当前模型到新候选的实际变化范围。检查包括真实单元格、规则范围差异、合并、列宽和 typed workbook resources；不能只按新模式的目标范围授权。资源原始快照复用 `cloneSheetClipboardWorkbookResources`。table/pivot/chart/slicer 的 capture/planner 返回也已接到既有 `workbookResources` / `originalResources` 分支，但资源保真和权限草稿仍未执行。

### 2.1 共享入口能力与权限边界

新增 `sheet-clipboard-capabilities.ts`，通过 `sheet-clipboard.ts` 转导出下列模块接口；本专项未修改包根 `src/index.ts` 的公开导出：

| 接口 | 输入与结果 |
| --- | --- |
| `sheetClipboardCapabilities(input)` | 输入 workbook、sheetId、完整 selection、userId、可选 snapshot、实际 reader 可用性、active/editing/exclusiveRange/shapeTextEditing、可选 permission 与 previousPasteMode；返回 copy、cut、copyFormulas、六模式 paste、copySpecial/pasteSpecial/ribbonPaste 和 source |
| `sheetClipboardCommandCapability(capabilities, commandId, mode?)` | canonical 与固定版别名复用 resolver；返回 `{ hidden, disabled, reason?, effectiveMode? }`，非剪贴板命令返回 undefined |
| `SheetClipboardPermissionEvaluator` | 同步接收 point、phase、workbook/sheet、unitId/sheetId、ranges、userId、mode；返回 boolean 或 `{ allowed, reason? }` |

workbook/sheet 使用不含 cells 的结构投影；生产提供领域模型，standalone 从 view 提取保护、合并和过滤等几何。入口预判不解析或克隆整本工作簿。snapshot 的 `hasValue`、`hasFormula` 仅描述缓存内容；空白源可清值或应用默认格式，无公式也不凭按钮名称禁用公式模式。`sourcePresent` 表示缓存携带 source 定位；cut 另检查源表仍存在并属于同一工作簿，普通 copy 不要求源仍存在。内部或 HTML snapshot 的列宽不依赖 source 定位；文本 snapshot 不提供源格式或列宽，无效列宽禁用宽度模式。

| 操作 | 预判权限 |
| --- | --- |
| copy / copyFormulas | 源范围 view + copy；cut 另要求 edit |
| values / formulas | 目标 view + edit + setCellValue |
| formats | 目标 view + edit + setCellStyle |
| columnWidths | 实际目标列的整列 view + edit + setColumnWidth；rowCount 缺省使用 1,048,576。单格源粘到完整合并区时按当前 planner 的实际锚点列判断宽度 |
| all / exceptBorders | 已知内部/HTML 源要求目标 value + style；纯文本保留目标样式，只需 value；formulaOnly 的 all 有效模式为 formulas |
| cut 的各粘贴模式 | 同工作簿、活源、源 view/copy/edit；只有实际 all 移动另检查源 setCellValue/setCellStyle，特殊模式保留源内容 |

worksheet/range 的 view/edit 保护先检查，注入 evaluator 不能绕过这些模型规则。`ExcelWorkbench.clipboardPermission` 是额外 UI 操作点判断，不等于持久权限服务或 API 的最终候选授权；未提供该 prop 时，命令仍由既有领域保护检查负责。

已知 snapshot 按区间预测目标扩展、重复、筛选跳过和完整合并范围，不逐行枚举整列。系统 reader 内容未知时，分别评估缓存与可能的新系统内容，任一可执行候选即可启用入口；旧 cut、已移除源或文本缓存不能否定新的外部 HTML。未知普通 all 可能是纯文本，因此可沿 value 权限启用；读取到实际 HTML 后 controller 再复查 style。可调用 reader 不证明系统读取权限已获准，也不保证当前内容可粘贴。

选区优先显式 activeRangeIndex，缺省最后范围；不把多选当成 exclusive range。后者来自固定版插件 interest group 的概念，由 `clipboardContext.exclusiveRange` 显式提供；特殊容器和 Ribbon paste 隐藏，普通 quick copy/cut/paste 单独判断。shapeTextEditing 禁用表格内容操作，并隐藏 Ribbon paste。[menu-util][clipboard-menu-util]、[factory][sheet-menu]。这些 UI hidden/disabled 条件不能直接用来拒绝合法的非活动/隐藏表 API 操作。

`useSheetClipboard` 返回新增 `capabilities` 和 `pasteOptionsCapabilities`，Ribbon、右键和浮层叶项共用对应结果；菜单分发时再守卫。最近粘贴将 `previousPasteMode` 传入同一 helper，切换模式同时检查旧模式恢复权限；撤回列宽模式后再切为值也必须具备整列授权。预判以原 sheet 的合并/过滤几何保留当前保护规则，最终恢复后的资源差异、指纹、历史可重规划状态和全候选权限仍由领域命令负责，不能由该 helper 证明全部资源权限闭环。

## 3. 系统剪贴板与异步保护

`sheet-system-clipboard.ts` 将浏览器 I/O 与模型分开，提供可注入 `read()` / `write()` 及可选 `canRead()`。`sheetSystemClipboardAvailable(adapter?)` 必须确认 read 实际可调用且可选 probe 未否定；浏览器适配仅在 native read/readText 为函数时视为有 reader。明确不可用时不调用 read，继续内部缓存路径。写入 `ClipboardItem` 的 plain 和 HTML；rich API 不存在或失败时可回退 readText/writeText。实际读取为空或仅媒体时返回明确空内容，避免用旧缓存替代；undefined 或最终读取错误才允许 controller 使用读取前捕获的内部 ID。显式 text/html、同义参数或 clipboardId 不读取系统剪贴板；只指定 mode/value/type 的普通粘贴仍读取系统内容。

只有实际 HTML token 与内部缓存匹配，且没有冲突的显式 clipboardId，才能恢复完整内部模型。外部 HTML/文本走共享 parser。token 缺失、过期或与显式 ID 冲突时，不能借用旧缓存里的 rich runs、hyperlink 或扩展 metadata。optional 不接受不同的显式 clipboardId。

`sheet-clipboard-ui-host.ts` 在读取前捕获 studio、UI workbook/sheet 和领域 workbook/sheet 引用、unit/sheet/range、content revision、history generation；standalone 另捕获 view JSON。新请求号、Escape、切文档/表、同 ID 重开、模型替换、历史重置、编辑与 deactivation 会使旧读取不能发布。

读取期间改选区不会改写捕获的粘贴目标。生产 API 只在当前活动范围仍等于请求范围时更新粘贴后选区；UI 发布前后复查 revision、generation、模型和预期 view JSON，在 apply 或 change listener 更换 UI view 后停止旧通知及焦点恢复。

延迟 focus 还复查请求号、宿主、view 内容、选区与读取前的 DOM 焦点。用户保持选区但转去外部 Input 时，内容仍可粘贴到原目标，焦点不会被 controller 拉回。只允许同一原焦点，或原控件已销毁且焦点退到 body 时恢复；该行为仅有未运行浏览器草稿。

取消复制先使读取任务失效。生产 cancel-copy 被 BeforeCommandExecute 拦截时返回 false，并只向原 view 报错；标记和剪切缓存按真实命令结果保留。原生 Escape 表示按键已处理，不据此虚报命令取消成功。普通 copy 的内容缓存与 copy mark 分离，成功取消标记不删除普通 copy 内容；cut 的标记由共享 store 消费或取消。

## 4. Canvas 浮层与组件库

copy 虚线复用 `sheetRangeDashedOutlineCommands`，仍由现有 Canvas/WebGPU 或 Canvas2D renderer 绘制。本专项不引入 Table 或整页 contenteditable。

`sheet-clipboard-overlay.ts` 按实际结果末端、对应 frozen/scrolling pane 与 zoom 计算 CSS 像素坐标。隐藏、零尺寸、越界、不可见末端、过小视口和非有限数值不显示按钮。按钮为 44×24px，间距 4px，限制在视口内容区；上游使用 padding 2px 及不同的边缘余量，这是明确的实现差异。[popup][clipboard-popup]

新增浮层使用 `antdv-next` 的 Dropdown、Button、Tooltip。菜单挂到 body，沿统一 ConfigProvider/theme token；独立 `useId()` 标识每个 Canvas 的按钮和 popup，避免焦点误认其他实例。涉及的 Canvas/公式输入继续使用 `antdv-next` Input/TextArea/AutoComplete 和前述最小原生事件桥。本专项没有新增原生表单控件。

安装版 `antdv-next` 为 1.8.2，Vue 为 3.5.45，`@v-c/dropdown` 为 1.1.2。安装版 Dropdown autoFocus 的 RAF 不承载本项目宿主/会话取消保护，因此关闭其 auto-focus，改用有请求号、宿主、revision、options、选区和当前焦点保护的最多三帧 menu focus。关闭、deactivate/unmount 取消待执行 focus；关闭后的 nextTick 另检查重新打开状态和请求号。ArrowDown 打开并聚焦菜单，popup capture Escape 关闭；外部控件获得焦点会关闭菜单，避免全局 Escape 夺回焦点。

上游 `onCloseAutoFocus` 委托 layoutService.focus。本地增加用户新焦点、宿主和重新打开保护；真实明暗主题、禁用/悬浮/焦点状态、输入法和多实例交互尚未验证。前序 Context7 文档对应 1.8.0，安装版类型与实现是本批次静态依据，不能以文档替代浏览器验收。

`ExcelWorkbench.hidePasteOptions?: boolean` 已接入 controller，只改变浮层呈现，不删除 recent、选区、内容或历史；配置撤回后可再次显示当前最近模式，隐藏期间 optional 命令仍可调用。右键 quick Button、禁用的右键/浮层 Menu label 使用 `antdv-next` Tooltip 呈现理由；禁用 Button 外保留接收 hover 的最小 span，右键提示层高于现有菜单/子菜单层级。保留原可访问名称并补 aria-description，动态 disabled 与点击守卫已接线。安装版 MenuItem 的字符串 title 在 vertical 模式仅传给底层原生 title，因此本轮不依赖它作为主题化工具提示。Context7 本轮请求安装版本 1.8.2，其解析目录列出的版本仅为 1.7.1/1.8.0，结果链接为 main；采用已读取的本地安装源码核对 Tooltip/Menu/Dropdown 接口，未据此声称 hover 或键盘行为验证通过。

### 4.1 公开入口、状态和媒体契约对照

按 Ribbon → 上下文 → 最近粘贴 → Facade → controller/service 顺序复查。以下记录本批已接入的 UI 契约和仍由主任务推进的 API/媒体缺口；没有运行行为验证。

| 契约 | 固定来源与可观察行为 | 本地静态差异和实现建议 |
| --- | --- | --- |
| 根 Facade 粘贴 | [f-univer.ts][clipboard-facade] 的 `IFUniverSheetsUIMixin.pasteIntoSheet(htmlContent?, textContent?, files?): Promise<boolean>` 与实现直接派发 `SheetPasteShortKeyCommand` | 对照时 `FOffice` 未提供该方法。复用共享命令、目标捕获、权限与内容历史，保留 boolean 异步结果和显式文件输入。无参固定实现走 `legacyPaste` 的 unrecognized hook，并非独立读取系统；如本地采用其他无参行为须记录差异 |
| 范围 HTML | [f-range.ts][clipboard-range-facade] 的 `generateHTML()` → `generateCopyContent(...).html ?? ''` | 对照时 `FRange` 未提供方法。复用共享 capture/content，不另维护 DOM 内容，也不执行 cut 或增加内容历史；生成 HTML 本身不设置系统剪贴板或 copy mark。默认样式、富文本、合并、行列属性和扩展保真分别待验证 |
| 四个公开事件 | [f-event.ts][clipboard-event-facade] 及 [f-univer.ts][clipboard-facade] 的 `BeforeClipboardChange`、`ClipboardChanged`、`BeforeClipboardPaste`、`ClipboardPasted`；copy 参数含 workbook/worksheet/text/html/fromSheet/fromRange，paste 参数含 workbook/worksheet/text/html；Before 的 cancel 可阻止操作 | 对照时 `ApiEventMap` / `FEventName` 只有通用命令事件。主任务在共享命令层接入可取消的 Before 和隔离监听器错误的 After，并在事件回调后复查目标、引用和版本，避免事件重入覆盖新内容。生产 UI 已走共享 `executeCommandSync`，无需额外发布一组事件；standalone 没有 FOffice 事件宿主 |
| Ribbon / context 状态 | [menu.ts][sheet-menu] 的 `menuClipboardDisabledObservable` 判断系统 reader 或 lastCopyId；copy/cut 分开检查 copy/view/edit；值、公式、格式和列宽使用各自权限；special 容器受 exclusive range 显隐控制，Ribbon paste 另排除 shape text editor | 已加入共享 capability，Ribbon/context 的容器、叶项和分发守卫共用状态。默认 paste 禁用但特殊模式许可时根改为 menu。UI evaluator、模型保护和最终 API 候选权限分开；上游 factory 与专用命令的权限集合有差异，不能把 UI 显隐用作全部 API 规则 |
| 粘贴浮层配置 | [config.ts][clipboard-config] 的 `clipboardConfig.hidePasteOptions` 以及 [clipboard-ui.controller.ts][clipboard-ui-controller] 控制注册；clipboard controller 同样检查该配置 | 已接 `ExcelWorkbench.hidePasteOptions`，只影响浮层呈现。保留 recent、内容、选区和历史；隐藏期间 optional 仍可调用，撤回配置后显示当前模式；草稿未运行 |
| UI 与 API 内容别名 | [clipboard.command.ts][clipboard-command] 的 short-key 参数为 htmlContent/textContent/files/formulaClipboardPayload/target | UI 与主任务 API 已改为共用 normalizeSheetClipboardContentAliases，接受一致同义值、拒绝冲突/非字符串，空字符串不触发系统读取。此 helper 只负责文本/HTML；文件、媒体和公式 MIME 仍为缺口，显式目标继续由各宿主/命令检查 |
| 媒体与聚焦对象 | [service][clipboard-service] 的 copy 先执行 `onBeforeCopyFocusedObject`，chart sheet 用对象 carrier；`legacyPaste` 存在文件且 HTML 含 img 时走 files，否则 HTML 优先。`type.ts` 声明 `onPasteFiles` 和 `onPasteUnrecognized`，controller 收集 DataTransfer 文件及公式专用 MIME | 本地系统/UI 已有 text/HTML/图片 File 和准备后媒体粘贴接线，完整聚焦对象、任意文件、chart sheet、公式专用 MIME 及资源互操作仍未验收。后续在共享 typed resource/媒体服务和命令内完成对象 ID、权限、撤销和保存闭环；data-source-only chart range 移动草稿不能替代绘图对象复制。上游 hooks 是公开 service 扩展契约，不能冒称都已暴露为 FOffice 方法 |

上述来源均为同一固定提交；本地 `E:\rust_test\univer` 的只读 `git status --short` 未显示在途改动。没有运行菜单 factory、Facade 或媒体行为。

## 5. 文件与测试草稿

本专项文件位于 `packages/office-studio/`：

- `src/components/ExcelWorkbench.vue`、`src/components/spreadsheet/SpreadsheetCanvas.vue`。
- `src/components/spreadsheet/use-sheet-clipboard.ts`、`sheet-clipboard-ui-host.ts`、`sheet-system-clipboard.ts`、`sheet-clipboard-overlay.ts`。
- `src/components/spreadsheet/sheet-ribbon.ts`、`sheet-context-menu.ts`、`SheetContextMenu.vue`。
- `src/features/sheet-workbench/sheet-clipboard-input.ts`、`sheet-clipboard-capabilities.ts`；`sheet-clipboard.ts` 仅追加本专项两个模块的转导出，planner/API/Facade/events 的并行实现由主任务负责。
- 下表测试文件及 `tests/fixtures/sheet-clipboard.html`；此前的 `SheetFormulaBar.vue` 原生输入修改归入 SHEET-05。

静态 `rg -c '^test\('` 声明数如下，**全部未运行**：

| `tests/` 文件 | 声明数 | 草稿覆盖意图 |
| --- | ---: | --- |
| `sheet-clipboard-input.contract.test.mjs` | 4 | 同义参数、冲突/非字符串、空字符串/空白、输入不变 |
| `sheet-clipboard-capabilities.contract.test.mjs` | 13 | reader/缓存候选、空白/formulaOnly、独立 copy/view/edit/value/style/width、目标扩展/过滤/合并、active/exclusive/shape、活源/cut、整列与合并锚点、无效列宽、旧模式恢复权限、别名与不修改输入 |
| `sheet-clipboard-menu.contract.test.mjs` | 4 | Ribbon/context 共用叶项状态、根 split/menu 切换、禁用分发、exclusive 与无缓存 |
| `sheet-clipboard-ui.contract.test.mjs` | 57 | 原生事件、rich token、相对公式、模式/type/内容别名冲突、显式空与回退、乱序/宿主/取消、生产 amendment、实际结果范围、发布回调、standalone no-op/资源恢复；hide 切换与隐藏期间 optional、独立操作点、新旧模式权限和缓存保留；formats 的 model/view 字体、run-only 历史/值验证隔离、recent 真实值恢复与稀疏内容事件；外部 forced-string、values 源局部字体与公式目标正文、cell image/publicBody、权限/stop 验证和内部重开；formulas 的普通值/空格、formulaOnly/cut 引用平移、隐藏/零宽列、宽度验证隔离和 recent 拒绝；单段均匀 run 的普通解析、目标字体、forced-string、局部/未知/图片/body 排除与 recent |
| `sheet-system-clipboard.contract.test.mjs` | 8 | plain/HTML 双格式、实际空内容、缺 reader、读取错误、Blob 写出、纯文本回退、实际 callable/probe |
| `sheet-clipboard-overlay.contract.test.mjs` | 5 | zoom、冻结、滚动/裁切、边缘余量、隐藏/零尺寸/无效 viewport、既有 whole-sheet 虚线命令 |
| `sheet-clipboard.browser.test.mjs` | 27 | 真实库控件与事件、active range、六模式/历史、HTML/多行输入、异步宿主/焦点、取消/权限、主题/冻结/缩放、menu 重开、外部 Input/standalone；旧列宽恢复禁用、跨表资源/undo/redo、hide 切换、无 reader、独立 copy/value/style 权限和 exclusive/shape；明暗主题的 formats run 字体、内容事件/重复粘贴/历史及内部保存重开；forced-string HTML 的 values 快捷键/最近模式、历史和内部重开；formulas/widths/exceptBorders 的目标恢复、局部字体/边框、事件顺序和单条历史；均匀 HTML span 的数值解析、整段字体归源 cell 样式、目标字体恢复与内部重开 |
| 合计 | 118 | 声明数量，非通过数量 |

浏览器草稿通过 Vite / Vue plugin 和 Playwright Edge 使用真实组件，注入系统剪贴板并强制 Canvas2D。没有启动浏览器；注入 reader、合成剪贴板事件及 composition 不能证明系统权限、跨应用剪贴板或真实 IME 行为。资源恢复与旧列宽权限草稿已跟随共享领域接口接线，仍需统一验证。

生产浏览器草稿先粘贴到 E5:F5，再切为列宽；独立撤销 E11 的编辑权限后，切为值会尝试恢复整个 E 列的旧宽度。菜单预判禁用该选项，合成 click 也不派发命令；草稿检查模型、权限、revision、undo/redo、最近模式和通知不变。UI 契约草稿另直接调用 optional 验证命令拒绝和错误反馈。夹具直接改变规则 owner，模拟内容 revision/history generation 不变的独立权限变化，不作为持久权限服务已经实现的证据。

另一个场景通过真实 `antdv-next` 工作表标签将 Source 的 A1:B2 table 剪切到 Report 的 E5:F6；图表数据源随区域移动到 Report，但绘图仍属于 Source，position/size 保留。最近粘贴改为值须恢复原 table/chart 资源并保留源内容，最终选择形成同一生产历史 entry 的 redo；草稿继续检查 undo/redo。系统剪贴板、保存和独立重开均未执行。

夹具新增 `?no-reader`、浮层配置、exclusive/shape 和独立操作点控制；默认不注入 permission。浏览器草稿分别检查只读 reader 缺失、许可格式但禁止值时根改为 menu、copy/cut 分开、未知 reader 纯文本可执行但实际 HTML 仍需 style。列宽断言使用领域 `sheetColumnWidth()`，兼容 span 与逐列存储，不再直接假设 columnWidths 字段存在。

## 6. 保留缺口

完整 SHEET-08 仍需核实系统剪贴板权限、跨应用 Microsoft Office HTML/RTF、特殊格式、公式与资源的全部场景、保存/XLSX 独立重开、中文 IME、真实明暗主题和焦点、多实例、WebGPU/Canvas2D 实机及大样本。跨工作簿 cut 当前拒绝并保留源内容，其真实移动闭环仍待接入。

本专项 evaluator 只增加同步 UI 操作点，持久权限、服务端执行和 API 候选授权仍需主任务接入并验收。入口 helper 不验证指纹/历史可重规划状态或完整 typed resource 候选；最近粘贴预判的旧模式权限不能代替恢复后的实际资源变化检查。包根 Facade、generateHTML 和四个剪贴板事件的公开契约继续由主任务完成，当前表中的缺口未据按钮或测试声明改为完成。

固定 [service][clipboard-service] 和 [controller][clipboard-controller] 还有图片、文件、聚焦绘图对象和 chart sheet 的剪贴板路径。本地系统适配已传 text/HTML/图片 File，`use-sheet-clipboard.ts` 已接准备后的媒体粘贴和图片对象路径；这些源码接线没有替代跨应用媒体、聚焦 chart / 嵌入对象和 chart sheet 的实际验收。typed table/chart 引用移动草稿也不能证明聚焦对象复制。全部 Pro 入口、权限和资源保真目标继续保留。

没有同一数据集的设备、浏览器、渲染后端、首屏/滚动/编辑帧时间、长任务、Worker/WASM 传输或内存实测。本次同步 clone/JSON 身份保护未据基准确定大任务阈值；不声称性能已优化或达标。

<a id="sheets-start-history-five-roots-review"></a>

## 2026-11-27 START/HISTORY 前五根续核

按固定 v1.6.0 / `eaf4cfa04fabdbcfc922c692f3450459f8a29546` 的 [ui schema][ui-history-schema] 与 [Sheets schema][sheet-schema] 复查前五根。常驻根 5；粘贴基础子项 5；格式刷 continuous `subId` 1，不计第二个根；More 复用原入口，本批新增溢出根 0。公式插件的上下文扩展仍按第 1 节单独清点，本地 Ribbon 额外公式子项 1 是既有差异，不能据子项数量宣称功能完成。

| 根 | classic order / grid | 固定命令与本地链路 |
| --- | --- | --- |
| 撤销 | 0 / r1c1 | UndoCommand；本地 core / univer / sheet 别名共用内容历史注册 |
| 重做 | 1 / r2c1 | RedoCommand；与撤销共用按文档隔离的 historyStore |
| 格式刷 | 2 / r1c2 | `ui.operation.activate-format-painter`，continuous 为 subId；本地会话、模型事务与选择历史见[专项续核](SHEET-04-format-painter.md#sheets-format-painter-history-run-style-review) |
| 清除格式 | 3 / r2c2 | `ui.command.clear-formatting` → Sheets host 的 clear-selection-format；本地有效 sheet 路径与共享公开派发缺口见[清除续核](SHEET-02-merge-clear.md#sheets-clear-formatting-dispatch-review) |
| 粘贴 | 4 / r1c3，跨两行 | `sheet.menu.paste`；全部、值、格式、列宽、边框以外内容五个基础子项 |

[固定历史 menu][ui-history-menu] 在无内容历史、cell editor 激活或公式栏编辑时禁用，并随 RTL 交换图标。[LocalUndoRedoService][undoredo-service] 按 unitID 隔离栈，mutation 全部执行成功后才移栈，dispose 清理对应文档。本地 [content-history-commands.ts](../../packages/office-studio/src/api/content-history-commands.ts) 与 `OfficeStudio.replayContentHistory` 共用目标解析和事件；[historyStore.ts](../../packages/office-studio/src/stores/historyStore.ts) 在 undo / redo 回调成功后移栈。工作台编辑态另读取 recorder / inline draft，走 `runCellDraftHistory`，组合输入时禁用；这是已有实际草稿历史行为，不能只按内容栈禁用而删除它。

粘贴调用链已重新读到 [use-sheet-clipboard.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-clipboard.ts) 的异步宿主捕获、系统读取、命令派发及选择 / 焦点保护，[sheet-clipboard-commands.ts](../../packages/office-studio/src/api/sheet-clipboard-commands.ts) 的目标 / 历史 generation、完整 before / after selections、候选发布与最近粘贴 amendment，以及 [planner](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard.ts) 的实际离散目标、验证、权限、源清除、引用 / 资源移动。最近选项恢复原数据后形成候选，复用原 undo 与新 redo；固定 service 则执行 top undo 后重放粘贴，内部实现不同，真实对外行为仍待运行验收。

固定 [clipboard utils][clipboard-utils] 的值粘贴也可携带源 richText；格式粘贴经 [SetRangeValuesMutation][set-range-values] / [handleStyle][sheet-cell-style] 将源字体应用于目标 run。本地 values 保留目标 cell style，源 richText 与旧公式目标的处理见[值粘贴续核](#sheets-values-forced-string-review)；原 formats 保留目标 run 字体覆盖的差异已按[本批局部修复](#sheets-format-paste-run-style-review)落盘，行为仍未验收。列宽写入现有 axis span，边框以外内容保留目标边框；公式子项及跨工作簿 cut 仍按既有差异记录。

保存已追到 `FWorkbook.save` 深克隆领域 WorkbookModel、markSaved 的文档 savepoint，重开经 `createWorkbook` / `workbookFromInput`。本批没有运行内部重开草稿、XLSX 写出 / 独立读取或完整资源互操作，也没有主题、焦点、真实 IME 和性能数据。源码记录、格式刷和 formats 小修已落盘，五根及完整 Sheets / Docs / Pro 仍保持 **未验收**；统一静态收尾由主线汇总。

<a id="sheets-format-paste-run-style-review"></a>

## 2026-11-27 formats 粘贴的局部字体与内容事件

固定 [clipboard command][clipboard-command] 的 `sheet.command.paste-format` 选择 `SPECIAL_PASTE_FORMAT`；[clipboard controller][clipboard-controller] 的 formats hook 调 `getSetCellStyleMutations` 并替换目标 merge。[utils][clipboard-utils] 只构造源样式 `s`，以 `isOverrideStyle: true` 交给 [SetRangeValuesMutation][set-range-values]；[handleStyle / mergeRichTextStyle][sheet-cell-style] 将 cell 字体写入目标富文本。该链路保留目标正文，不复制源 run，也不把 cell 背景当作 run 背景。

本地 [sheet-clipboard-cell.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-cell.ts) 的 formats 分支复用既有 [clearSheetCellRichTextFonts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-rich-text-style.ts)，清除目标 run 的 family、size、bold、italic、underline、strike、color 及其公开顶层别名，让目标正文继承刚粘贴的 cell 字体。正文、链接、run ID、未知 run 字段及嵌套 font 的未知字段保留；源 run 不复制。本地以继承表达字体，上游以合并 run 字体表达，存储结构不同。共享 rich-text helper 没有修改。

原 [planner](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard.ts) 的 formats 目标不进入值验证坐标；只有内容或公式路径才加入相应检查。run-only 字体清理仍形成真实 cell 差异，由原粘贴事务写入内容历史，重复的相同 formats 不增加历史。recent 仍恢复原 sheets / metadata / resources 后重规划，并保留最初 undo、替换最终 redo；本批没有改变 savepoint、选区、DTO 或 history 协议。

原 API 将权限受影响范围直接用于 `SheetValueChanged`，其中包含样式、merge、规则和列宽。[sheet-clipboard-changes.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-changes.ts) 新增 `sheetClipboardChangedValueRanges(before, after)`，按实际 scalar、formula、有效值/缓存类型、清除建模字体后的 richText、链接、单元格图片和 sparkline 比较内容，复用既有稀疏范围合并。已建模 run 字体、cell 样式、merge、规则和列宽本身不构成值变化；只有样式的空白单格也不虚报值事件。未知扩展仍参与保留后的内容比较，没有推断其语义。

[sheet-clipboard-commands.ts](../../packages/office-studio/src/api/sheet-clipboard-commands.ts) 保留 `assertSheetClipboardWorkbookChangesEditable(current, candidate)` 的完整权限检查，仅将内容范围用于值事件，列宽继续发维度事件，并保留原值事件先于维度事件的顺序。内容比较使用当前模型到最终候选，因此从 all 切为 formats 恢复真实值/公式时仍发事件，不按模式一概跳过。

[sheet-auto-fill.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-auto-fill.ts) 的 `ONLY_FORMAT` 也调用同一个 cell factory，会同步采用清除目标 run 建模字体的行为。本批没有修改 fill 命令、fill 内容事件、其他粘贴模式或单独扩展填充验收。

本批在既有 [UI 合同草稿](../../packages/office-studio/tests/sheet-clipboard-ui.contract.test.mjs) 新增 4 条声明：领域 planner 与 view 的字体继承/内容保留；生产及 standalone 的 run-only 格式、stop 值验证隔离和重复 no-op；recent 恢复真实值/公式并排除同区域纯格式单格的事件；稀疏内容变化与空白样式/几何隔离。生产历史草稿覆盖 undo、redo、`FWorkbook.save` 和内部 `createWorkbook` 重开。既有 [浏览器草稿](../../packages/office-studio/tests/sheet-clipboard.browser.test.mjs) 新增 1 条，配合 [fixture](../../packages/office-studio/tests/fixtures/sheet-clipboard.html) 经 `antdv-next` Ribbon / 最近粘贴菜单检查明暗主题、焦点、Canvas 字体场景、内容事件、重复模式和内部保存重开。

本 formats 子项仅完成源码核对、修改和 `rg` 声明清点；该批次落盘时合同为 49 条、浏览器为 24 条，运行均为 0。生产 UI 未新增可替换控件，继续沿用统一 ConfigProvider 与 `antdv-next`。浏览器草稿强制 Canvas2D，字体场景检查和合成事件不能证明实际像素、系统剪贴板权限、真实 IME、WebGPU、XLSX 写出/独立读取或性能。本修复及完整 SHEET-08 均保持 **未验收**。

<a id="sheets-values-forced-string-review"></a>

## 2026-11-27 values、公式目标与 forced-string 续核

固定 [clipboard controller][clipboard-controller] 的 `SPECIAL_PASTE_VALUE` 调 [getSetCellValueMutations][clipboard-utils]。普通值构造 `{ p: null, v, t }`；真正富文本构造 `{ p, v }`，源正文、局部字体和正文内链接可以随 p 复制。该 mutation 不写 s，目标 cell 样式保留。[isRichText][sheet-rich-text-predicate] 根据局部样式、特殊 `va`、bullet、多段落、customRanges 或 customBlocks 分类；仅有普通字体且覆盖整段正文的单一 run 不满足该分类。

该批次的本地 [sheet-clipboard-cell.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-cell.ts) 在非公式目标上仍复制源文字 runs，保留其字体、正文内链接、ID 和未知字段，目标 cell 字体、数值格式及独立 cell hyperlink 继续保留。当时本地任何 source.richText 数组都可复制，上游按完整 document body 分类；后续仅对信息完整的普通单段 run 实施[本次窄修](#sheets-uniform-rich-text-paste-review)，完整 body 分类仍未齐，没有统一清除所有源 run 字体。

[公式 clipboard hook][formula-clipboard-controller] 的 `getSpecialPasteValueValueMatrix` 对源公式/数组公式写计算值并清 f、si、p；旧目标是公式时也清 p，若源有正文则调用 `getCellRichText` 取 `dataStream` 并去掉末尾文档结束符。本地 values 已在既有 cell 工厂按目标原 formula 增加小分支：源文字 runs 拼成正文，清 richText，图片/publicBody 不复制到旧公式目标；沿用原公式/缓存清除、目标样式/独立链接、metadata 失效和显示格式。源 runs 为空正文时保留已有源值。非公式目标的源富文本保留行为不变。普通文字 runs 场景已写入草稿；原始 publicBody 的非文本块、图片占位符、仅正文元数据而无文字 runs、数组公式绑定和完整 Office 样本仍未核实，不能据本分支宣称全部 p 语义已齐。

[独立 hyperlink hook][hyperlink-clipboard-controller] 将 VALUE 列入 specialPastes 并跳过独立 link 资源复制；[drawing hook][drawing-clipboard-controller] 对 VALUE 跳过浮动绘图复制。单元格图片由 p / drawingsOrder 表达，[cell-type][sheet-cell-type] 的 `isCellImage` 与明确 FORCE_STRING 判断另有实际模型规则。本地非公式目标的 values 复制 cell image 的资源并分配新 ID，保留 source publicBody 和 prior drawing ID，公开写出复用 [sheetCellImageToPublicBody](../../packages/office-studio/src/features/sheet-images/sheet-image-public.ts) 重写 drawing 引用；后续 plain values 清旧 image/body。该路径与浮动对象复制分别记录，cut 特殊模式的图片身份及完整绘图互操作仍未验收。

原 externalCell 使用 `normalizeSheetCellValueType(source.dataType) === 'force_string'`；normalize 把 forced string 归为 text，条件不会成立。本批改为复用已有 [isForcedSheetCellValueType](../../packages/office-studio/src/domain/sheet-cell-value.ts)，使显式 forced-string 的 `'00123'` 和 `'=1+1'` 保持字面值。该共用入口同时影响外部 all、values、formulas：all/values 不重新解释字面文本，formulas 不把 forced 字面表达式当源公式；普通 text 继续走既有解析。未改类型协议、HTML parser 或值引擎。

权限和 stop 值验证继续由既有 planner 与 API 检查；本批没有另建权限、验证、事件或历史路径。values 在无 setCellStyle 权限时仍可保留目标样式并写值，缺 setCellValue 或违反 stop 验证时原子拒绝。recent 仍恢复最初模型后重规划，因此公式目标的清 runs 行为能够参与原 history entry 的 amendment；内容事件复用前一 formats 批次的实际内容范围比较。

既有 [UI 合同草稿](../../packages/office-studio/tests/sheet-clipboard-ui.contract.test.mjs) 新增 4 条声明：text/HTML 与三种 forced 类型别名及 all/values/formulas；生产/standalone 的局部字体、旧公式目标正文、目标样式/链接、recent no-op、内容事件及生产 undo/redo/内部重开；真实 PNG 的 cell image/publicBody 复制、公开 drawing ID 重写与 plain/公式目标清除；权限和 stop 验证拒绝后无内容/事件/历史/recent 变化。[浏览器草稿](../../packages/office-studio/tests/sheet-clipboard.browser.test.mjs) 新增 1 条，经名称框选 E5，合成 Ctrl+Shift+V 与无富文本的 typed HTML paste 检查两列 forced 字符串、最近模式、undo/redo、焦点和既有内部保存重开；fixture 未为该条新增接口。

该批次落盘时已用 `rg -c '^test\('` 核对 UI 合同 53 条、浏览器 25 条、本表专项合计 112 条；**运行均为 0**。本批修改一个生产 cell 工厂、两份既有草稿和本证据文件；前序 formats 文件保留。没有新增 UI 控件，现有输入、Ribbon 和最近粘贴菜单继续使用统一 ConfigProvider / `antdv-next`。合成事件和内部重开草稿不证明真实系统剪贴板、IME、XLSX 独立读取、WebGPU 或性能；本批及完整 SHEET-08 继续保持 **未验收**。

<a id="sheets-special-paste-modes-review"></a>

## 2026-11-27 formulas、columnWidths、exceptBorders 与最近粘贴续核

固定 [formula hook 的 `getSpecialPasteFormulaValueMatrix`][formula-special-paste] 对 f/si 来源写平移后的公式，并清 v/p；普通来源无条件写 v、清 f/si/p，有 p 时取正文。该 special hook 不因 cut 跳过引用平移。[公式专用复制][formula-empty-copy] 将非公式位置填为空格，[普通复制片段][clipboard-copy-fragment] 也将源空格显式写入 matrix，不能将本地稀疏快照的缺失位置解释为保留目标内容。外部来源继续由 `getValueMatrixOfPasteFromIsNull` 单独处理：公式字符串写 f，普通来源只清旧公式目标。

本批修改 [sheet-clipboard-cell.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-cell.ts)：内部 formulas 的 cut 与 copy 都复用既有 `translateFormulaReferenceAreas` 平移相对引用；非公式来源写普通值，真实文字 runs 转正文并清目标 runs/image/publicBody；稀疏或 formulaOnly 空格清旧目标内容，空目标保持稀疏。目标 cell 样式、独立 hyperlink 和无关 metadata 保留，公式缓存沿原路径失效。外部非公式来源在非公式目标上保留既有不写行为；values 的前批源局部字体行为保留。特殊 cut 模式仍保留源区域，只有 all 进入既有移动路径。

固定 [列宽 hook][column-width-paste] 不写 cells，只循环源列属性并写 truthy 且不等于配置默认宽的值；配置来自 `DEFAULT_WORKSHEET_COLUMN_WIDTH_KEY`，缺省 [88][column-width-default]。[SetWorksheetColWidthMutation][column-width-mutation] 跳过不可见目标列，缺少 colWidth 键时保留原宽，undo factory 捕获旧宽。本批修改 [sheet-clipboard.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard.ts)，复用既有 `hiddenSheetColumns` / `isSheetColumnVisible`，跳过源零宽及目标点隐藏、隐藏 span、折叠 outline 和零宽列，其他宽度继续写现有 axis span。**默认宽过滤仍未齐**：本地 snapshot 仅存 number[]，没有冻结的配置默认宽语义；本地无默认配置时首列 152 / 其余 112，合同与浏览器样本分别为 96 / 110，不能硬编码 88 或读取已变化的源配置代替协议。该 DTO 方案留待主线确认，本批没有新增协议。

固定 [exceptBorders hook][except-border-paste] 深克隆源值/样式并将 s.bd 置 null，经普通样式合并保留目标 border，再清/设 merge。本地既有 cell factory 复制源正文、真实局部字体及其他样式，将目标有效 border 写回 authoredStyle；继承/稀疏样式按本地模型表达，结构不同。目标有 border 的空格保留边框实体，无边框的空目标保持稀疏。独立 hyperlink 与 floating drawing hook 未将该模式列入跳过复制项，本地继续沿既有范围/图片资源链路；本批没有用空操作替代资源行为，也没有扩改资源模块。

菜单、工具栏和 UI 命令仍共用 resolver / capability。formulas 要求目标 view/edit/setCellValue，不要求 setCellStyle；本地 widths 按整列 view/edit/setColumnWidth 判断，固定 [`getSheetPastePermissions`][clipboard-paste-permissions] 的 widths 分支另要求 value/style/columnStyle，该权限粒度差异保留。planner 的宽度路径不加入值验证，formulas 的正文或清空继续经过 stop 验证。recent 先检查旧/新模式权限，再恢复最初 sheets/metadata/resources 重规划，API 复查当前到最终候选的真实变化；拒绝候选不改变内容、事件、历史或最近模式。内容事件复用前批实际内容比较，所有 `SheetValueChanged` 先于列宽 `SheetDimensionsChanged`；宽度模式从公式模式恢复原内容时仍发真实值事件。

固定 service 通过 top undo 后重放实现 recent；本地保留原 undo、替换最终 redo。已核对 `historyStore.ts` 的 generation 在文档重建/clear 时变化，execute/amend 不改变 generation；本批没有虚构 cut 的 generation 冲突或另建历史路径。保存继续由 `FWorkbook.save` 深克隆领域模型并 markSaved，内部重开经 `createWorkbook` / `workbookFromInput`，内容、列宽与隐藏状态沿既有模型持久化。

既有 [UI 合同草稿](../../packages/office-studio/tests/sheet-clipboard-ui.contract.test.mjs) 新增 2 条声明，分别覆盖 production/standalone 的普通 copy、formulaOnly、cut、普通值/稀疏空格、引用平移、目标样式/链接、权限、内容事件、重复 recent 与生产 undo/redo/内部重开；跨表 widths 覆盖三种隐藏原因、源/目标零宽、值权限及 stop 验证隔离、实际整列维度事件和被 stop 拒绝的 recent 候选。既有 [浏览器草稿](../../packages/office-studio/tests/sheet-clipboard.browser.test.mjs) 新增 1 条，配合 [fixture](../../packages/office-studio/tests/fixtures/sheet-clipboard.html) 新增的维度/事件顺序记录和样本接口，经真实 `antdv-next` 最近菜单从 all 切 formulas、widths、exceptBorders，检查目标正文/字体/边框和原宽恢复、源保留、重复 no-op、值事件先于维度事件、单条历史、焦点及内部保存重开。没有新增或迁移 UI 控件，仍使用统一 ConfigProvider；草稿切换明暗主题并强制 Canvas2D。

该批次落盘时已用 `rg -c '^test\('` 核对 UI 合同 **55** 条、浏览器 **26** 条、本表专项合计 **115** 条；运行均为 **0**。两份既有 mjs 与 fixture module 的静态语法、diff 空白、本页引用及固定源码路径检查通过；未运行产品、测试、类型检查、构建、浏览器或基准。本批生产改动仅上述两个剪贴板文件，前序 API/内容事件修复保留。共享/数组公式身份与 ft/fd 标志、完整 run/body 分类、非文本 p/body、默认宽过滤、完整范围/绘图资源、跨工作簿 cut、真实系统/Office 剪贴板、IME、XLSX 写出和独立读取、WebGPU 与性能仍未核实或未运行；本批及完整 SHEET-08 保持 **未验收**。

<a id="sheets-uniform-rich-text-paste-review"></a>

## 2026-11-27 普通单段均匀 run 的 values 与外部 HTML 续核

固定 [isRichText][sheet-rich-text-predicate] 去掉末尾 CRLF 后按正文覆盖范围分类：普通字体完整覆盖整段本身不构成 rich；有样式的局部 run、特殊 va、bullet、多段落、customRanges/customBlocks 才构成 rich。固定 [`getCellStyleBySnapshot`][sheet-whole-run-style] 另把覆盖整段的普通字体写为 cell 样式。[getSetCellValueMutations][clipboard-utils] 仅在 rich 分支携带 p，普通分支清 p，并在外部无数字格式时复用 numfmt 解析。`sheetCellRichTextMatches` 仅判断文字一致，现有字体/编辑 helper 没有相同的完整 body 分类方法。

本批仅修改 [sheet-clipboard-cell.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-cell.ts)：新增供 externalCell / values 两处复用的窄判定，要求无 image/publicBody、仅一个与 cell.value 完全一致的文字 run、无换行，且复用 `clearSheetCellRichTextFonts` 清已建模字体后只余 text/id。链接、va、未知 font/run 扩展、多个 runs 和原始正文均不进入这次归一；这是已读信息的能力判断，没有新增输入 schema 校验或推断丢失的段落数据。共享字体 helper、HTML parser、DTO、权限和历史文件未改。

eligible external HTML 用 `sheetCellRichTextRunView` 将整段字体应用于既有源 view，再交原 `editSheetCellText` 处理正文；复用 `sheetViewCellStylePatch` 将新字体/数值格式合入源 authoredStyle，避免 `applyExcelCellView` 对新 cell 重新 compose 时恢复旧字体。例如一个普通 Arial 24 span 的 `00123` 在 General 目标按既有路径变为数值 123，all 将 Arial 24 写为源 cell 样式；values 继续保留目标字体和独立 hyperlink。文本格式目标仍保留 `00123`。内部 values 对该单段 run 写标量并清 runs，内部 all 沿原源复制行为，真正局部字体与未知扩展继续随既有 values 路径保留。

显式 forced-string 仍由 `isForcedSheetCellValueType` 在外部解析前保护，`00123` / `=1+1` 保持字面值。该本地契约与固定源码并非完全相同：固定 [外部公式 hook][formula-external-paste] 只按 value.v 是否为公式字符串分支，没有检查来源 FORCE_STRING；`getExternalPasteNumberFormatResult` 只接收 v、目标样式和 parseOptions，也不读取来源 t。本地按明确类型与 Office 可观察语义保留既有保护；本批没有运行相应固定版/Office 样本，不能将两条上游路径标为完全一致。

旧公式目标的正文提取、公式/缓存失效和 runs/image/publicBody 清除保留前批路径，图片来源继续复用原 ID 复制与公开 drawing 引用重写。有 publicBody 的单段 run 不参与本次归一。非图片原始 publicBody 当前仍在 values 的 metadata 失效阶段清除；本地尚无完整正文类型、布局与复制分类，这个真实保真缺口保留，不能据普通 run 的小修宣称全部 p 支持完成。

既有 [UI 合同草稿](../../packages/office-studio/tests/sheet-clipboard-ui.contract.test.mjs) 新增 2 条声明：嵌套/顶层字体别名、内部 values、外部 General/文本格式、forced-string、局部/va/未知/链接/换行及图片/publicBody 排除；production/standalone 的目标字体、旧公式清除、独立权限、recent 切换/重复 no-op、内容事件、源保留及生产 undo/redo/内部重开。既有 [浏览器草稿](../../packages/office-studio/tests/sheet-clipboard.browser.test.mjs) 新增 1 条，复用现有 fixture 的 `specialPasteModes`，以合成 HTML paste / Ctrl+Shift+V 和 `antdv-next` 最近菜单检查普通 span 的解析、all 字体写入 authoredStyle、values 目标字体恢复、明转暗、Canvas 字体场景、重复事件、单条历史、焦点和内部重开；fixture 未新增接口或控件。

本批实际声明数为 UI 合同 **57**、浏览器 **27**，本表专项合计 **118**；运行均为 **0**。本批仅修改上述 cell 工厂、两份既有草稿和本证据，统一静态收尾由主线执行；不运行产品、测试、类型检查、构建、Rust/WASM、dev、浏览器或基准。现有控件继续使用 ConfigProvider / `antdv-next`，合成剪贴板与 Canvas2D 场景不证明真实系统权限、Office p/RTF、IME、XLSX 独立重开、WebGPU 或性能。完整 run/body 分类与前述 SHEET-08 缺口保持 **未验收**。

[sheet-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/schema.ts
[sheet-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/menu.ts
[formula-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-formula-ui/src/menu/schema.ts
[formula-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-formula-ui/src/menu/menu.ts
[clipboard-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/clipboard.command.ts
[clipboard-service]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/services/clipboard/clipboard.service.ts
[clipboard-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/clipboard/clipboard.controller.ts
[clipboard-ui-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/clipboard/clipboard-ui.controller.ts
[clipboard-config]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/config/config.ts
[clipboard-facade]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/facade/f-univer.ts
[clipboard-range-facade]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/facade/f-range.ts
[clipboard-event-facade]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/facade/f-event.ts
[clipboard-popup]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/views/clipboard/ClipboardPopupMenu.tsx
[clipboard-menu-util]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/menu-util.ts
[ui-history-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/schema.ts
[ui-history-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/history.menu.ts
[undoredo-service]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/services/undoredo/undoredo.service.ts
[clipboard-utils]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/clipboard/utils.ts
[sheet-cell-style]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/basics/cell-style.ts
[set-range-values]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/mutations/set-range-values.mutation.ts
[sheet-rich-text-predicate]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/editor/editing.render-controller.ts#L1176
[formula-clipboard-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-formula-ui/src/controllers/formula-clipboard.controller.ts#L396
[hyperlink-clipboard-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-hyper-link-ui/src/controllers/copy-paste.controller.ts#L157
[drawing-clipboard-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-drawing-ui/src/controllers/sheet-drawing-copy-paste.controller.ts#L115
[sheet-cell-type]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/basics/cell-type.ts
[formula-special-paste]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-formula-ui/src/controllers/formula-clipboard.controller.ts#L453
[formula-empty-copy]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-formula-ui/src/controllers/formula-clipboard.controller.ts#L83
[clipboard-copy-fragment]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/services/clipboard/clipboard.service.ts#L610
[column-width-paste]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/clipboard/clipboard.controller.ts#L715
[column-width-default]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/sheets/sheet-snapshot-utils.ts#L33
[column-width-mutation]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/mutations/set-worksheet-col-width.mutation.ts#L68
[except-border-paste]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/clipboard/clipboard.controller.ts#L793
[clipboard-paste-permissions]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/clipboard.command.ts#L215
[sheet-whole-run-style]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/editor/editing.render-controller.ts#L1215
[formula-external-paste]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-formula-ui/src/controllers/formula-clipboard.controller.ts#L359
