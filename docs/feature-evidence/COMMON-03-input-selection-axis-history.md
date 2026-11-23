# COMMON-03 输入、选区、行列命令与内容历史

记录日期：2026-11-22（Asia/Shanghai）。涉及 COMMON-02/03/05/06/11、SHEET-02/03/05/06、DOC-01/02/03。状态：**开发中，全部未验收**。

按用户安排，本批仅修改代码、补测试草稿和静态回读。未执行测试、类型检查、构建、Rust/WASM、浏览器、语法检查或基准；本记录不将任何草稿写成通过。

## 1. 固定依据与入口

Univer v1.6.0 固定提交为 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，只读检出 `E:\rust_test\univer`。完整 Ribbon 的 Tab/Group/order、classic/grid、常驻/扩展/More/上下文入口清点沿用 [Sheets 源码入口表](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)、[Docs 与 Sheets 完整入口表](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)及[工具栏清单](../UNIVER_TOOLBAR_PARITY.md)。本批增加共享行为和调用链，未重新计算运行时入口数量；历史按钮、别名、测试声明均不当作功能完成数。

| 固定源码，路径相对于上游根 | 可核实契约 | 本地实现与差异 |
| --- | --- | --- |
| `packages/sheets-ui/src/commands/commands/set-selection.command.ts`：`MoveSelectionEnterAndTabCommand`；`commands/commands/utils/selection-utils.ts` | Enter/Tab 的方向、按键、合并格、多个选区及 Tab 起点；选区通过 operation 发布 | `sheet-selection-navigation.ts` 和公共选区命令共享规划；用稀疏轴索引跳过手动隐藏、筛选、折叠与零尺寸。没有完整复刻上游 selection/render service，其他方向键与公式引用选区仍需继续核验 |
| `packages/sheets-ui/src/controllers/editor/editing.render-controller.ts` | 原值/公式与编辑文本分离；确认/取消；拒绝提交不移动 | `sheet-cell-edit-session.ts` 记录目标、源内容/类型/数字格式、日期系统、地区及年份。冲突协议是本地实现，不冒称上游采用相同版本/字段比较 |
| `packages/sheets/src/commands/commands/{insert-row-col,remove-row-col,move-rows-cols}.command.ts` | 插删、移动与引用资源/历史编排 | 本地结构 planner 用一次 detached workbook 计划更新单元格、尺寸、合并及可识别资源；连续资源被拆开时明确拒绝，复杂/未建模资源保真未验收 |
| `packages/sheets/src/commands/commands/{set-row-visible,set-col-visible}.command.ts`；`packages/sheets-ui/src/commands/commands/hide-row-col-confirm.command.ts` | 手动隐藏与零尺寸分开；全部可见行列不能隐藏；确认入口调用真实隐藏命令 | 本地紧凑 span 表示手动隐藏；保留过滤/大纲原因。隐藏全部返回稳定 `SHEET_AXIS_HIDE_ALL` 错误，UI 显示对应反馈，不复制上游仅警告的 confirm 宿主 |
| `packages/sheets/src/commands/commands/{set-worksheet-row-height,set-worksheet-col-width}.command.ts`；`packages/sheets-ui/src/commands/commands/set-worksheet-auto-col-width.command.ts` | 手动/自动尺寸、拖拽增量、对应可撤销模型修改 | 本地 dimension feature、Facade 与行列头共用注册命令；对话框捕获原目标/版本，拒绝过期确认。字体度量和 Office 单位换算仍待真实样本验证 |
| `packages/docs-ui/src/commands/commands/inline-format.command.ts`：`SetInlineFormatCommand` | 折叠选区只缓存下一次输入格式；空 mutation 不清瞬态格式；中西文字体槽 | `document-input-store.ts` 管理文档会话；非空范围由字符格式服务修改 runs。清色意图保留 own property 的 `undefined`，不通过 JSON 克隆丢掉 |
| `packages/docs-ui/src/services/doc-ime-input-manager.service.ts` | composition 捕获源范围，最终文本与历史组成一次输入 | 本地 body/story 输入桥校验公共 revision/generation，避免旧 composition 或提交结果复活旧范围。合成事件草稿不证明系统中文 IME 已通过 |
| `packages/core/src/services/undoredo/undoredo.service.ts` | unit 隔离、成功后移动栈项、销毁清理 | Pinia `content-history` 是本地内容历史源；保留 callback 事务和 savepoint。协作 mutation/逆 mutation、持久版本服务与该本地栈分别验收 |
| `packages/ui/src/menu/schema.ts`、`menu/history.menu.ts`、`controllers/shared-shortcut.controller.ts`；`packages/core/src/facade/f-univer.ts` 与 `__tests__/f-univer.spec.ts` | 开始区历史顺序 0/1、grid 首/次行第 1 列；菜单、快捷键与全局 API 调同一撤销/重做命令；Facade 将历史事件与一般 command 事件分开，全局异步取消返回 false | 本地 `content-history-commands.ts` 注册 core/univer/sheet/doc 别名，全部进入同一 replay；生产 Ribbon 与全局快捷键经 workspace adapter 执行。同步 unit API 保留既有取消异常契约，全局 FOffice 只把明确取消转为 false，普通监听器异常仍拒绝 Promise |

上述源码路径可按固定提交复查；详细 Docs 字符、字段缓存和来源样式差异见 [DOC-02 清除格式](DOC-02-clear-formatting.md)。

## 2. Sheets 的共享模型与事务

`packages/office-studio/src/features/sheet-workbench/sheet-selection.ts` 维护多个独立/重叠矩形、`activeRangeIndex` 和实际 `primary`；不是一个外包矩形。`sheet-selection-store.ts` 按 workbook/sheet 保存完整状态和 Tab 起点。公共 getter 与观察事件返回分离的选区对象，旧 `SelectionChanged` 是活动范围投影，完整状态来自 `SelectionsChanged`。

`sheet.operation.set-selections`、`sheet.command.select-range` 和 `sheet.command.move-selection-enter-tab` 在 `src/api/sheet-selection-commands.ts` 与 OfficeStudio 注册入口中解析统一请求。Facade 的 A1/范围列表、行列头、Canvas、名称框和 Enter/Tab 通过同一选区状态；navigation 不进入内容历史。结构修改按 rangeType 保留整行/整列意图，撤销恢复完整选区，随后切换的活动工作表可以保留。

`sheet.command.commit-cell-input` 的 `sheet-cell-input-command.ts` planner 先处理合并格 owner、输入设置/源会话、查看/编辑权限、类型/日期/数字格式及数据验证，再返回候选单元格。`src/api/sheet-cell-input-commands.ts` 经 `commitWorkbook` 提交一次内容历史并发布真实 `SheetValueChanged`。Facade `FRange.setInputText` 和 `ExcelWorkbench.vue` 的单元格/公式输入使用同一命令；多单元格赋值继续使用范围值 API。stop 拒绝、warning/information 反馈、无变化、冲突与错误各自保留结果。

F2 的既有富文本编辑保留未受影响的运行、链接、空标记和扩展字段；直接覆盖输入清掉旧运行。当前 `sheet-cell-rich-text-edit.ts` 通过最终字符串的前后公共段推断替换区间，没有逐次捕获 native beforeinput 源范围。重复文字中间的样式边界、多次编辑/原生删除、富文本局部格式与真实 Office 重开仍是待修复/待验证项，不能由单个运行保留草稿推断 SHEET-05 完整支持。

## 3. 行列命令与 UI

结构命令在 `src/api/sheet-axis-structure-commands.ts` 注册插入前/后、删除、移动/复制及公开别名，调用 `planWorkbookAxisStructureCommand`。规划时处理源/目标边界、合并格、筛选行排除、范围权限及 Office 上限；由 `sheet-axis-structure-resources.ts` 更新可识别的公式、条件格式/验证范围与公式起点、自动筛选、表格、保护、定义名称、超链接、迷你图、图表/透视绑定和冻结信息。资源类型全部组合及 vendor 公式/未知资源仍未验收。

显隐命令在 `src/api/sheet-axis-visibility-commands.ts` 注册 `set-rows-hidden` / `set-col-hidden`、`set-specific-rows-visible` / `set-col-visible-on-cols`、selected-visible、confirm 与本地别名。手动 hidden points 归并为 spans，不展开百万行；显示不会清掉过滤和折叠。零尺寸与手动 hidden 同时存在时按上游分阶段语义恢复。隐藏计划及公式/筛选刷新后的准备阶段均检查至少一个有效可见位置。

尺寸命令在 `src/api/sheet-axis-dimension-commands.ts` 统一注册手动、自动及拖拽增量。三个 `components/spreadsheet/use-sheet-axis-*` controller 在 `finishEdit` 成功后捕获宿主，通过 `sheet-axis-command-host.ts` 校验 studio、文档/工作表、activeKind、发布后的 view 身份与内容 revision。命令观察者切换宿主、改变选区或再次编辑后，不向旧宿主发布旧 view/选区/焦点。API 已发布的完整选区由 UI 投影，不另造轴选区。

`commitWorkbook` 在权威模型与历史变得可观察前完成 detached 候选、预算/权限、公式/筛选、自动行高、选区快照和 `prepareView`。prepareView 得到另一份候选克隆；修改它不会改变待提交权威模型。准备期间真实内容改变则拒绝旧计划，导航改变保留。standalone host 复用领域 planner，并检查 view 身份及数据签名；这条边界不替代宿主自身的历史、协作或大文件性能验收。

## 4. Docs 的输入会话

`document-input-session.ts` 的 `DocumentInputSession` 包含 body/story location、marks、modelRevision、generation、valid。Pinia `document-input-store.ts` 按文档隔离，输入偏好不写入快照/dirty/history。选择和格式、内容失效、成功输入、清理各有明确转换；public getter 克隆 location/metadata 并保留颜色/高亮清除的属性意图。

`doc.operation.set-input-selection` / `clear-input-marks` / `advance-input-marks` 与 editor command 共用 `document-input-commands.ts` 的位置、UTF-16 边界、权限/控件锁和格式校验。`document-indexed-input.ts` 把 Facade 的正文 offset 映射为段落/表格实际位置，保留既有换行语义，拒绝跨独立表格单元格或未寻址非文字块的 replacement。

body 插入接受成功后，只有原内容 revision、generation 和宿主仍对应，才一次发布新的 caret/marks。story 的 `DocumentStoryInputCommit` 声明源 range、结果 caret、modelRevision、generation；先验证真实拼接（含重复文字和 Unicode 删除），再授权一次 advance。无标记的普通 header/footer 修改、后续编辑、导航和旧 generation 都不能恢复旧草稿。调用者收到 `selectionCurrent: false` 时不应用旧返回选区。

DocumentWorkbench 与 Canvas/input 的闭环、CJK/字段缓存局部格式、关闭/重绑和 canceled undo 的详细修改由 [DOC-02 证据](DOC-02-clear-formatting.md)记录。正文和 story 使用 `antdv-next` Input/TextArea 的最小输入桥；内容仍由 Canvas 渲染。

## 5. 历史、观察事件与取消

`src/stores/historyStore.ts` 在 execute/undo/redo 回调及栈/revision 修改期间保持 `isApplying`。同 unit 重入 execute/undo/redo/clear 在修改前拒绝；其他 unit 可独立事务。退出先释放 unit lock，再降低全局 depth，使已提交后的观察者可以开始下一次编辑。失败回调释放 scope 并保留原栈项以便重试；领域命令自身仍负责修改前准备与原子性。

OfficeStudio 立即让过期 Docs 会话失效，但在 history 应用期间推迟 `DocumentInputStateChanged`；Sheets 选区同样等历史结束再发布。输入/选区事件可检查当前 generation/revision/宿主，重入后停止旧事件的后续投影。观察事件通过 `emitSettled` 给订阅者独立 payload，错误记录到 `EventListenerError`，不把已提交编辑改成失败；可取消 Before 事件继续使用同步 `emit`，取消和监听器异常对调用方可见。

`replayContentHistory` 在 BeforeUndo/BeforeRedo 之前捕获实际目标、历史对象/顶部 entry、savepoint revision/generation 与模型 revision。回调后若产品/文档、内容或历史改变，返回 false，不撤销另一个文档或新插入的编辑。成功 Docs replay 只在预期会话仍对应时清旧输入偏好；取消/拒绝保留待输入格式。FWorkbook 的既有 undo/redo 返回 facade，FDocument 与 OfficeStudio 返回实际 boolean，未改变该接口区别。

`core.command.undo/redo`、上游 `univer.command.undo/redo` 与本地 `sheet.command.undo/redo`、`doc.command.undo/redo` 均已注册到同一历史实现。公共请求校验 kind/unitId，产品别名拒绝相反 kind；同 ID 对应两种模型时必须明确类型。历史命令沿用固定版 Facade 的专属 Before/完成事件，不额外发一般 BeforeCommandExecute/CommandExecuted。`FOffice.undo/redo` 新增 Promise<boolean>；`ContentHistoryCanceledError` 只表示显式取消，异步全局 API 将其转为 false，普通 Before 监听器异常仍向调用方报告。同步 API 保留取消抛出的既有行为，尚不能由此宣称完整 Facade 兼容。

生产 `examples/workbench/src/App.vue` 的两种 Ribbon 与全局 Ctrl/Cmd+Z、Ctrl/Cmd+Shift+Z、Ctrl/Cmd+Y 已改由 `workspace-history-operations.ts` / `useWorkspaceController.ts` 执行上述 API 路径。适配器先捕获活动 kind/unitId，拒绝另一产品或已替换宿主发来的事件；Docs emit 使用可选 unitId 字符串，Sheets emit 使用可选 `{ workbookId }`，生产壳把两者转为同一目标参数。取消在宿主返回 false 并通过库控件反馈，继续由同一历史源决定 dirty/savepoint；不预先清除内容、选区或 marks。UI 的真实焦点恢复和缓存工作台切换仍需统一浏览器验证。

本批没有新增历史按钮。两产品各沿用开始区两个常驻根，宽度不足时由现有 More 呈现；新增插件根、上下文根均为 0，运行态溢出数量没有重测。8 个命令 ID 是兼容别名，不能当作 8 项功能或入口完成数。

这些保护不是服务端并发控制、完整 public Facade 兼容或持久版本服务的完成证明。内容历史仍保存本地快照/callback，100 项上限与大文件峰值内存需要继续验收。

## 6. 测试草稿与待统一执行

路径以下相对于 `packages/office-studio/tests/`；只记录覆盖意图，不记录通过数量。

| 草稿文件 | 主要场景 |
| --- | --- |
| `sheet-selection.contract.test.mjs` | 范围/重叠/primary、稀疏可见轴、四方向 oracle、百万行合并、多选区与 Tab 起点 |
| `sheet-selection-api.contract.test.mjs` | A1 与 unit/sheet 身份、多选区 getter/events、取消、跨表历史、观察者重入与发布顺序 |
| `sheet-cell-input-command.contract.test.mjs` | 运行时边界、日期/类型/格式、源冲突、合并 owner、验证与富文本保留/覆盖 |
| `sheet-cell-input-api.contract.test.mjs` | 共享注册命令、dirty/savepoint、无变化、取消/冲突、prepareView 失败、BeforeUndo/Redo 切换及新编辑、观察者隔离 |
| `sheet-axis-edit.contract.test.mjs` | 紧凑隐藏、零尺寸/过滤/折叠、结构引用与样式、边界/权限、API 和三个 controller 的宿主/发布保护 |
| `history-publication.contract.test.mjs` | 栈/revision 结束后观察、失败重试、同 unit 重入拒绝、跨 unit 独立、settled listener 的下一次编辑 |
| `content-history-command.contract.test.mjs` | 注册别名、全局异步返回/取消、同步 unit 兼容、非法/歧义目标、专属事件、生产适配器的取消与过期宿主拒绝 |
| `document-input-session.contract.test.mjs` | 临时 marks 与实际 runs、权限/锁/UTF-16、索引桥、取消与 undo/redo、Before/after 重入、story 拼接及一次授权 |
| `sheet-selection-input.browser.test.mjs` 与 `fixtures/sheet-selection-input.html` | 真实 ExcelWorkbench/Facade、Canvas/公式栏、多选区、确认/取消、合成 composition、隐藏合并、宿主切换、主题/布局与内部重开 |
| `document-story-input.browser.test.mjs` | 真实 Docs Canvas 输入/工具栏、body/story 会话、混合缓存格式、合成 IME、导航/重绑/关闭与 canceled undo；详见 Docs 专项 |
| 根 `tests/browser/office-workbench.spec.ts` 的本批追加草稿 | 生产 App 的 Sheets/Docs × undo/redo × Ribbon/全局键盘：取消后模型/savepoint/输入状态不变；Before 导航后旧请求停止；恢复原宿主后一次成功重放 |

既有 `api-session-history`、`history-savepoint`、轴尺寸和 Office parser/writer 草稿仍须一起执行。还需真实系统 IME、暗/亮弹层/焦点、WebGPU/Canvas2D/设备丢失、导出独立重开、跨模块与权限组合。百万行合成对象草稿只核对模型算法，未执行，更不构成浏览器大文件测量。

## 7. UI 与性能边界

本批涉及的 Sheets 公式/单元格输入、轴尺寸弹窗和 Docs input/覆盖层继续使用 `antdv-next`，由统一 ConfigProvider/token 提供主题。Canvas 绘制/命中/选择与最小输入桥分层。已补主题/焦点/composition 草稿，但本批未在浏览器检查，未迁移区域和真实 IME 效果不能由静态 import 推断。

生产壳的重命名迁入 Form/Form.Item；解析加载、错误和消息通过新 `WorkspaceFeedback.vue` 使用 Spin/Alert/Button；诊断详情由 `WorkspaceDiagnostics.vue` 使用 Collapse，移除原生 details/summary。两个子组件在应用统一 ConfigProvider 内读取 theme token，错误/警告/背景/边框随主题切换；原有文件操作与诊断 DTO 不归 UI 所有。移除对应硬编码反馈颜色和自绘 spinner，仅保留布局 CSS；RTL 消息定位、Form Enter 提交、Collapse 键盘/焦点及真实明暗主题均未运行验证。

2026-11-24 补应用壳文件选择间隔：`examples/workbench/src/App.vue` 将隐藏导入触发器设置为 `Upload :styles="{ root: { display: 'none' } }"`，隐藏最外层 wrapper，避免内层按钮隐藏后 wrapper 仍作为纵向 flex 项占位。静态依据是当前安装的 `antdv-next@1.8.2` 的 `dist/upload/Upload.js`：普通 `attrs.style` 合入 semantic `trigger`，最外层 span 使用 `mergedStyles.root`。文件菜单和 Ctrl/Cmd+O 仍执行库 Button 的 `$el.click()` 并通过 `beforeUpload` 调用文件服务；Upload 内部创建的原生 file input 属于库实现，未添加新的应用原生表单控件。整体仍在统一 ConfigProvider 内。本项没有新入口或业务协议；真实文件选择、同文件再次选择、取消、焦点恢复、明暗主题与无布局占位尚未浏览器复核，保持未验收。

共享候选/view 克隆、来源 JSON 物化、snapshot 历史和签名比较仍可能产生较大的临时分配。尚无同数据前后耗时、帧时间、主线程长任务、Worker/WASM 传输和内存实测；没有性能达标或优化完成结论。大任务的 Worker/Rust/WASM、取消、复杂 shaping/布局、跨用户协作/持久权限/历史及全部 Sheets/Docs Pro 目标继续保留，均未验收。
