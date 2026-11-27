# SHEET-04 格式刷（含 SHEET-14 条件格式插件）

记录日期：2026-11-22。验收状态：**未验收**。本批按续作约定只读取源码、实现及编写测试草稿，不执行测试、类型、构建、Rust/WASM、浏览器、基准或 `git diff --check`。下文的“已实现”仅指代码已落盘。

执行依据：[项目标准](../../PROJECT_EXECUTION_STANDARD.md)、[源码规则](../UNIVER_REFERENCE_RULES.md)、[功能清单](../UNIVER_FEATURE_PARITY.md)、[渲染与计算标准](../RENDERING_AND_COMPUTE_STANDARD.md)。完整 Sheets / Docs 及全部 Pro 目标保持原范围；本页只记录格式刷切片。

后续条件格式 Office 来源、dxf/theme、base/x14 合并、原生身份恢复、平均值与自动数据条阈值编辑，以及流式 CF/相关资源提取、UTF-8／整本模型预算、共享主题校验和失败接线见 [SHEET-14 Office 往返专项](SHEET-14-conditional-format-office-roundtrip.md)。该页区分历史新增 36 与本轮新增 54（流式/来源预算 46、整本接线 8），相关十一份文件当前共 112 个声明；仅为静态数量，均未运行。本页测试表保留格式刷与基准投影阶段的声明数量，不代表后续批次总数，所有新内容仍未验收。

## 1. 固定来源与入口清点

Univer **v1.6.0**，固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地只读根为 `E:\rust_test\univer`。下列路径均相对于该根，链接固定到同一提交。

| 来源 | 复查点 |
| --- | --- |
| [ui/menu/schema.ts][U1] | START / HISTORY 顺序、activate 入口，grid `row: 1, column: 2` |
| [ui/menu/format-painter.menu.ts][U2] | activate 根、continuous `subId`；OFF 且不能采样时禁用；激活后可取消 |
| [ui/controllers/format-painter.controller.ts][U3] | 共享交互与格式刷状态 |
| [sheets-ui/services/format-painter/format-painter.service.ts][U4] | 表格样本、状态、插件 hook 与格式应用编排 |
| [sheets-conditional-formatting-ui/controllers/cf.painter.controller.ts][U5] | `ConditionalFormattingPainterController`，`repeatByRange`、目标裁剪、同表扩展及跨表复制 |
| [sheets-conditional-formatting/services/conditional-formatting-range-transform.service.ts][U6] | 范围差量应用 |
| [sheets-conditional-formatting/models/conditional-formatting-rule-model.ts][U7] | 规则 ID、顺序，`addRule` 向首部插入 |
| [sheets-conditional-formatting-ui/menu/schema.ts][U8] | DATA / RULES 管理根，order 1，grid `row: 2, column: 1`；INSERT / CELL 的数据条和图标集快插根另计 |
| [sheets-conditional-formatting/commands/commands/set-cf.command.ts][U9] | `sheet.command.set-conditional-rule`、正向/逆 mutation 与工作簿历史 |
| [sheets-conditional-formatting-ui/views/panel/RuleEdit.tsx][U10] | `AddCfCommand` / `SetCfCommand`、范围与样式提交 |
| [sheets-conditional-formatting-ui/views/panel/RuleList.tsx][U11] | 当前选区/整表列表、`DeleteCfCommand` / `MoveCfCommand`、权限过滤 |
| [sheets-conditional-formatting-ui/menu/manage-rule.ts][U12] | 10 个管理 selector 项，清理旧规则的全部范围权限、当前范围禁用状态 |

入口分类：常驻根 **1**（格式刷），continuous `subId` **1**（双击连续模式，非第二个根）；条件格式插件新增格式刷根 **0**；More 溢出复用同一入口；本切片无新增上下文根。根计数仅说明清点口径，不能证明功能完成。Docs 的独立入口与全部 Pro 清点仍由[总入口表](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)记录。

本地 `sheet-ribbon.ts`：开始 / 剪贴板（组顺序 0），classic 第 3 项，grid `row: 1, column: 2`，图标 Brush，隐藏标签；激活态由格式刷会话提供，未激活时依有效工作表及源采样权限禁用。More 沿用相同 command / `subId`，不创建第二套状态。

条件格式管理是另一个既有根入口：上游此插件声明 DATA 管理根 **1**、INSERT 快插根 **2**，管理 selector 子项 **10**，均不计入格式刷根数。本轮没有新增 Ribbon 根。本地管理按钮位于数据 / 规则（组顺序 20），classic 第 2 项，grid `row: 2, column: 1`，标签“条件格式”，命令 `sheet.operation.open-conditional-formatting`。无工作表时禁用，打开时检查查看权限，提交时检查工作表及受影响范围的编辑权限。上游 selector 的快捷分类、新建、管理、清除，以及 INSERT 快插，不等于本地一个打开面板按钮；完整入口对齐仍由总清单记录，不能把本轮面板闭环视为这些入口全部完成。

## 2. 本地命令、模型与交互

- `sheet.operation.format-painter`、`sheet.operation.continuous-format-painter`、`sheet.operation.cancel-format-painter` 与应用命令共用 `SheetFormatPainterSession` 的 `off / once / continuous` 状态机。具体 ID 以 `sheet-format-painter-session.ts` 中注册表为准。
- `captureSheetFormatPainterSource` / `captureSheetViewFormatPainterSource` 保存稀疏、深冻结的源样式、合并及条件格式规则，不复制值、公式缓存或富文本内容。
- `createSheetFormatPainterPlan` 扩大跨合并采样；单格目标扩为源块；其他目标按源块铺设；多选重叠由后一个目标的相位控制。源交叉合并整块保留，不能只采样锚点。
- `applySheetFormatPainterCommand` / `applySheetViewFormatPainterCommand` 复用共享合并规划、权限和确认。样式、合并及条件规则在同一内容历史事务内写入；失败不提交部分内容。
- API 与 standalone 工作台都校验 workbook / sheet / session / 版本。失败保留 ONCE；成功 no-op 消耗 ONCE，且不新增内容历史；连续模式保留源样本。
- `OfficeStudio.sheetFormatPainter` 与 `SheetFormatPainterStateChanged`、`SheetFormatPainterConfirmationRequested` 已接入；本切片未新增专用条件格式事件。CF-only 不触发值验证或 `SheetValueChanged`。
- 工作簿替换、切换 Docs、源表删除、失去查看权限、dispose 清理会话；跨工作簿和重入拒绝。合并确认前后重新校验版本与权限，拒绝旧确认。
- Canvas 仅在真实左键选区手势完成后触发应用；pointercancel / lost capture 不应用。格式刷期间关闭填充柄、超链接跳转及双击编辑；源虚线是会话视图数据，不进入快照。
- 已有 Ribbon、单元格及公式栏 Input / TextArea、合并确认 Modal / Alert 使用 `antdv-next`；本切片未新增原生交互控件。明暗主题、焦点恢复、真实输入法尚未执行验收。
- 条件格式 Drawer、Form / FormItem、Input、Select、InputNumber、Checkbox、ColorPicker、Button、Tooltip 与 Alert 均使用 `antdv-next`，继承工作台统一 ConfigProvider。Drawer 关闭动画完成后由工作台恢复 Canvas 焦点；浏览器草稿覆盖提交、取消、主题切换与焦点，尚未运行。

## 3. 条件格式插件行为与差异

`sheet-format-painter-conditional.ts` 保存完整原始规则 ranges 以保留公式基准，先验证无循环 JSON、规则 ID、坐标、配置与阈值，再克隆。规则优先级、stopIfTrue、差异样式、数据条/色阶/图标及未知 metadata 保留。源无规则时仍清除目标内现有规则，范围外保留。

目标覆盖先裁剪规则，再铺设源相交部分，包含最后不足一个源块的部分区域。表达式、数值比较公式及 visual 的 min / max / thresholds / icons 中 formula 随目标移动；文本规则的 `"A1"` 等字符串保持文本。

同表只有采样规则整栈定义、优先级及公式语义一致时才全部复用源 ID；混用共享规则与新副本会改变目标叠加次序，因此整栈复用或整栈独立。独立目标连续重涂也尝试整栈复用，完整替换后保留原数组位置。范围覆盖等价时保留原表示，避免无意义历史。

### 3.1 部分有效引用、求值与标准 XLSX 投影

`SheetConditionalFormatRule.formulaAnchor?: { row, col }` 记录公式基准。未显式提供时按最早范围起点推导。`reanchorConditionalFormatRule` 比较相对引用语义；如果整体移动会将部分引用变成永久 `#REF!`，保留原公式和显式基准，在每个实际单元格求值时再移动引用。同基准的编辑保留原公式拼写，不为一次未改变的编辑改写大小写。

例如源 A2:A5、`=A1>5`，数值 A1:A5 为 `[7, 2, 9, 6, 10]`，刷到 A1:A4 后，这四格的条件结果应为 `[false, true, false, true]`；首格越界不应污染后续有效引用。同表可保留原规则范围外的 A5，仍是一条规则；跨表生成独立规则，并保留 `{ row: 1, col: 0 }`。面板改为 A1:A4、`=A1>8` 后应为 `[false, false, false, true]`。以上为测试草稿的预期，尚非运行结果。

数值比较的 formula1 / formula2 与视觉阈值中的 formula 已接入统一公式求值回调。`evaluateFormula` 接收条件单元格坐标，供 `ROW()` / `COLUMN()` 及引用返回使用；动态 `INDEX` 的引用端点、范围、错误传播与稀疏整表范围共用同一引擎。Canvas 场景通过 `conditionalFormatFormulaAt` 使用规则基准，不永久写入单元格底色。场景当前只构造当前工作表的公式工作簿，完整跨表、外部、结构化引用和函数覆盖仍是缺口。

`crates/office-writer/src/xlsx_conditional_formula.rs` 将基准投影到标准 SpreadsheetML：存在安全基准时调整同一条规则的 `sqref` 分片顺序并平移普通 A1 公式；不存在安全基准时使用引用返回的 `INDEX` / `ROW` / `COLUMN` 表达式，让每个操作数独立产生有效引用或错误。越界下标显式产生 `#REF!`，避免 `INDEX` 的 0 下标被解释成整行/整列。规则不拆开，统计使用同一单元格并集；没有把自定义 `formulaAnchor` 属性写到标准 worksheet XML。

Rust parser 按 `sqref` 第一项恢复非默认基准，并读取视觉阈值和显示标志。共享投影样本为 `tests/fixtures/conditional-formula-projection.json`；写出、解析重开及二次写出已有原生二进制合同草稿，均未执行。没有本项目扩展的标准 Office 来源从原生 UID/GUID 或稳定来源摘要生成本地身份；本地 native 写出则通过带来源指纹的私有投影恢复原 ID、公式表示、稀疏/相同优先级及未知 JSON。来源字段合并与未核实边界见 [Office 往返专项](SHEET-14-conditional-format-office-roundtrip.md)；完整未知扩展映射与实际文件验收仍未完成，内部 JSON 保真不能代替此项。

### 3.2 条件格式面板的真实提交闭环

`sheet.command.set-conditional-formats` 使用 `SetSheetConditionalFormatsInput`（unitId、sheetId / subUnitId、完整 rules、可选 expectedRevision）。规划先验证规则、范围、基准、颜色、公式和阈值，再确定真实受影响的旧/新范围。不同 priority 按数值排序，相同 priority 保留稳定数组顺序；单纯新增/删除造成未编辑规则数组下标变化，不把它们当作范围编辑。同 priority 的相对顺序变化仍检查受影响范围。

API 在提交前检查版本、查看权限、工作表编辑权限和旧/新范围权限，通过一次 `commitWorkbook(..., { recalculate: false })` 写入。no-op 返回 false，保留版本和历史；成功的规则编辑只有一个工作簿内容事务，不重算值缓存、不触发值验证或 `SheetValueChanged`。standalone 复用同一规划、验证和权限，开面板时捕获工作簿视图快照，拒绝过期提交。

面板的 `submit` 回调返回 `SheetConditionalFormatSubmitResult`：接受时提供 changed；拒绝时提供 error。失败保持完整本地规则草稿并显示 Alert，解除命令取消后可重新提交。接受后才关闭；Drawer 使用安装版 `antdv-next 1.8.2` 的 `afterOpenChange` 回调 **prop**，不是同名 emit。关闭动画结束触发 `closed`，工作台在文档、编辑会话及其他弹窗状态允许时归还 Canvas 焦点。工作簿/工作表切换清理旧上下文。

`conditional-format-rule-editor.ts` 深克隆未知根、style、config、threshold、icon 和 metadata 扩展；面板保留未编辑公式的原始文本与同语义 type / subType，支持 `highlightCell / formula` 和 `formula1=''`、公式保存在 value 的已有模型。未修改的导入公式可以保留当前 parser 尚不支持的语法，不能据此宣称该语法已可求值。新建或改写的公式必须通过共享 validator。归一化拒绝非法/部分非法规则 ranges；视图省略规则字段时，adapter 回退的模型规则也深克隆，避免输入与输出共享嵌套数组。

面板列表草稿已移除全量 priority 重编号：新增规则使用当前最小 priority 并放在数组首部，空列表用 1；删除仅移除目标；相邻移动只交换两条规则的位置与 priority，相同 priority 时保留新的稳定数组顺序。三个列表 helper 均返回隔离深克隆，未编辑规则保留原 priority 和未知扩展。这样对可编辑规则的新增、删除及互相移动不会因为无关受保护规则被重编号而误拒绝；跨越受保护规则仍检查双方范围并原子拒绝。该修正只有静态审阅和草稿，尚未运行验收。

上游 `_painterConfig` 保存源范围/表 ID，应用时读取 **live source rules**；本地选择**激活时不可变快照**，避免连续模式源后改影响已捕获样式与规则。这是实现选择，与上游行为不同，不能标为逐细节等价。上游同表扩展原规则范围，跨表反序添加新 ID 以配合 `addRule` 的 `unshift`；本地用显式 priority 与稳定数组顺序保持样本栈的可观察顺序。

上游面板新增、修改、删除和拖动重排分别派发逐规则命令；本地先编辑规则列表，再以一个完整列表提交形成一次历史事务。上游规则列表提供选区/整表过滤及按规则编辑权限过滤；本地当前显示整表规则，在最终提交时拒绝越权修改。这些界面、提交时机和权限展示差异仍需单独验收。

## 4. 当前真实边界

| 边界 | 开发进度与验收 |
| --- | --- |
| 相对公式复制到工作表边缘 | 保留公式基准与逐格有效/越界语义的实现已落盘；合同与浏览器草稿未运行，保持待验证/未验收 |
| visual 规则统计作用域 | 本地与标准 XLSX 投影保持一条规则的完整范围并集；统计、阈值及大样本行为待运行验收 |
| 默认样式继承 / 范围主题 | 样式采样复用现有模型适配；完整行/列/范围主题继承仍需独立核对，不据默认样式展开宣称覆盖所有继承 |
| 物理隐藏与筛除行 | 按真实坐标采样/铺设，仍需核对固定版及 Office 的隐藏/过滤选择行为 |
| 合并内容 donor | 复用现有 first-nonempty donor 合并契约；与 Office 内容保留/冲突提示差异待样本验收 |
| cellIs 的公式操作数 | 已接入统一公式回调与基准；函数目录、跨表/外部/结构化公式仍未覆盖，不能把普通样本当作全部规则完成 |
| 管理面板与历史/权限 | 完整列表原子提交、旧/新范围、局部优先级保留、取消/版本过期、no-op 与焦点草稿已补，均未运行；上游管理子项、选区过滤、权限展示及快插入口仍有差异 |
| 导入规则编辑保真 | 已保留未知嵌套扩展、公式基准、原公式文本、公式 subtype/value 存储；无效/未知规则类型和全部已知字段的编辑保真尚未覆盖 |
| Office 文件往返 | 标准基准投影、原生 writer→parser→writer 合同草稿已补；未运行，独立读取器、兼容程序、所有未知扩展与原 ID 映射均未完成验收 |

## 5. 测试草稿与剩余验收

2026-11-22 格式刷与基准投影阶段用 `rg -c '^test\('` / Rust `#[test]` 静态重数。这里保留该阶段测试声明，不把循环内的样本当成独立测试，不执行 runner；后续 Office 往返批次的当前声明总数单列在专项页。父任务维护的渲染测试另计。

| 测试文件 | 声明数 | 本次状态 |
| --- | ---: | --- |
| `sheet-format-painter.contract.test.mjs` | 56 | 未运行 |
| `sheet-format-painter-conditional.contract.test.mjs` | 44 | 未运行 |
| `conditional-format-formulas.contract.test.mjs` | 13 | 未运行 |
| `conditional-formatting.contract.test.mjs` | 25 | 未运行 |
| `sheet-conditional-formatting-commands.contract.test.mjs` | 21 | 未运行 |
| `sheet-format-painter.browser.test.mjs` | 51 | 未运行；本轮追加 16 个声明 |
| `crates/office-writer/src/xlsx_conditional_formula.rs` | 5 | 未运行 |
| `crates/office-parser/src/conditional_format_sidecar.rs` | 9 | 未运行 |

草稿覆盖模型/视图源捕获、冻结、样式与合并、相对/绝对公式、部分铺设、多选相位、随机小范围几何、整表规则裁剪、ID/顺序/no-op、非法输入与原子拒绝，以及 API 历史、回滚、内部重开、Canvas 条件结果动态变化和 standalone emit。新增浏览器草稿通过生产 `createSpreadsheetScene` 的完整单元格填充命令检查部分 REF 后的后续有效格，覆盖同表/跨表/standalone、面板编辑、旧保护范围、取消后重试、版本过期、未知配置/公式存储保留、明暗主题和关闭焦点。列表草稿合同补了不连续/相同 priority、深克隆、独立受保护规则不被重编号和跨受保护规则拒绝；`conditional-independent` 浏览器草稿用于验证 A1:A4 / D4:D5 的移动、删除及 G7 新建均保留 N11 规则与 priority 3，跨越 N11 则保留模型、版本及历史。既有未编辑面板 no-op 草稿新增实际 revision 断言。循环内样本仍只按所在 test 声明计数。新增内容未添加截图或像素比较；文件中的历史像素用例另列既有边界。

浏览器 fixture 强制 Canvas2D，不能用作 WebGPU、真实输入法或 Office 文件验收证据。既有 composition 草稿只派发合成事件；没有真实中文 IME 验收。本轮没有执行测试、类型、前端构建、Rust/WASM、浏览器、基准或 `git diff --check`，也不把 2026-11-21 的历史检查结果移作本次修改的通过证据。

同步单元格预算为 **100,000**（包括扩大目标及被替换的大合并）；条件范围数量预算 **100,000**，内部范围工作预算 **5,000,000**。范围归一化使用行事件及压缩列覆盖，不按规则范围逐格枚举；这些是实现与保护阈值，**没有真实性能数据**，不得称已经优化或达标。大型 Worker/WASM 格式刷、真实大样本性能/内存/取消、WebGPU、IME 与 Office 互操作均保持未验收。

## 6. 修改模块与接口

| 模块 | 本批改动 |
| --- | --- |
| `features/sheet-workbench/sheet-format-painter*.ts`、`api/sheet-format-painter-commands.ts`、`components/spreadsheet/use-sheet-format-painter.ts` | 会话、捕获、规划、规则插件、版本/权限/历史与工作台交互 |
| `domain/model.ts`、`parser/workbookNormalizer.ts`、`services/excelViewAdapter.ts` | 公式基准、规则归一化与双向深克隆 |
| `features/conditional-formatting/conditional-format-command.ts`、`api/sheet-conditional-formatting-commands.ts` | 完整规则列表规划、原子命令与 standalone 公用合同 |
| `features/conditional-formatting/conditional-format-formulas.ts`、`conditional-format-rule-editor.ts`、`conditional-formatting.ts` | 公式平移、可恢复边界、未编辑表达式/配置保留、列表局部优先级草稿、数值操作数求值 |
| `components/spreadsheet/ConditionalFormattingPanel.vue`、`components/ExcelWorkbench.vue` 的条件格式区域 | submit / closed 接口、拒绝保留草稿、上下文版本与关闭焦点 |
| `features/formula/formula-{types,parser,engine}.ts`、`rendering/spreadsheet/scene.ts` | 实际条件单元格坐标、引用返回的 INDEX 范围及 Canvas 求值接线；后续共享引用变换由主任务维护 |
| `crates/office-writer/src/xlsx.rs`、`xlsx_conditional_formula.rs`、`crates/office-parser/src/conditional_format_sidecar.rs` | 标准 XLSX 公式基准投影、第一 sqref 基准恢复、原生往返草稿 |
| 上表各测试文件、`tests/fixtures/conditional-formula-projection.json` | 模型、服务、命令、浏览器与 Rust 共享样本草稿 |

OfficeStudio 根注册、公共 index 导出及功能总清单由主任务维护。新增领域字段是可选 formulaAnchor；新增同步命令采用完整规则列表；面板由旧 apply emit 改为 submit 回调并新增 closed 事件。调用面板的宿主须返回接受/拒绝结果并处理关闭焦点，不能继续按旧的无返回 emit 假定提交成功。本轮列表草稿新增模块级 `prependConditionalFormatRuleDraft`、`removeConditionalFormatRuleDraft`、`moveConditionalFormatRuleDraft`，不改变完整列表命令的输入/返回契约。

<a id="sheets-format-painter-history-run-style-review"></a>

## 2026-11-27 START/HISTORY 格式刷续核

固定来源仍为 v1.6.0 / `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。[Sheets controller][U13] 的 `_getUndoRedoMutationInfo` 在单格扩大为源块时，将扩大范围的 redo selection 与原范围的 undo selection 一起放入内容 mutation。该 controller 只采样 cell style；随后 `SetRangeValuesMutation` 的 [handleStyle / mergeRichTextStyle][U14] 先随 `s: null` 清目标 run 样式，再把源字体应用于目标 run，不复制源正文。

[sheet-format-painter-commands.ts](../../packages/office-studio/src/api/sheet-format-painter-commands.ts) 原先先提交模型，再扩展选择，导致 redo 快照只记住原范围。当前复用 `WorkbookCommitOptions.selections`，before 保存实际完整范围、activeRangeIndex 与 primary，after 保存扩大范围及起点；模型与选择由同一内容历史项恢复。只在当前活动范围匹配实际目标时编排扩大，显式无关目标保留当前选择。成功但内容无变化的刷取仍扩大选择、消耗 ONCE，且不新增内容历史。

[sheet-format-painter.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-format-painter.ts) 的 model / view 路径复用既有 `clearSheetCellRichTextFonts`，清除目标 run 的七个已建模字体覆盖，让目标文字继承刷来的单元格字体。单元格样式相同但 run 字体不同也构成格式变化，可进入内容历史；原文、链接、run ID 与未知扩展保留。固定版直接重写 run 样式，本地以继承表达同一已建模字体效果；未知 run 样式约定及 Office 独立读取效果仍未核实。粘贴的 formats 分支保持独立，未随本次修改。

内部保存链路已读到 [FWorkbook.save](../../packages/office-studio/src/api/facade.ts) 深克隆当前 WorkbookModel 并标记 savepoint，重开经 `createWorkbook` / `workbookFromInput` 恢复领域数据。刷取会话和选择历史属于临时应用状态，不据此宣称 Office 文件往返或历史持久化已完成。

[合同草稿](../../packages/office-studio/tests/sheet-format-painter.contract.test.mjs) 该批新增 4 个声明，覆盖完整多选 undo / redo、冗余扩大不入历史、model / view 的 run-only 字体覆盖，以及 API 的历史、值验证隔离与内部重开；强化既有无关显式目标和单格扩展用例。[浏览器草稿](../../packages/office-studio/tests/sheet-format-painter.browser.test.mjs) 新增 1 个 Canvas 字体 / 历史用例，并强化既有名称框 `D4` ↔ `D4:E5`。该批完成时两文件分别为 60 / 56 个 `test` 声明，仅为静态数量，实际运行 0。

生产增量为 API `+11/-6`、model / view `+8/-3`；没有公共接口新增。语法、引用、链接与 diff 收尾由主线统一记录。测试、类型检查、构建、Rust/WASM、浏览器、Office 样本、真实 IME、主题 / 焦点和性能均未运行，状态继续 **未验收**。

<a id="sheets-format-painter-input-session-review"></a>

## 2026-11-27 格式刷输入会话与确认取消续核

本批继续原格式刷入口：START / HISTORY，classic 第 3 项、grid `row: 1, column: 2`；主按钮与 continuous `subId`、More 和共享 UI aliases 沿原命令表。常驻、扩展、上下文、溢出根和子项增量均为 0，没有修改 Ribbon 清点或新增按钮。

固定 [Sheets controller][U13] 的 `sheet-cells.canStart` 复用 [ClearSelectionFormatMenuItemFactory][U16] 的 disabled$。它调用 [getCurrentRangeDisable$][U15]，默认 `supportCellEdit=false`，可见单元格编辑器或 `FOCUSING_FX_BAR_EDITOR` 会返回禁用，并继续判断工作表/范围权限与绘图上下文。这是本批输入禁用的具体依据。固定 [共享 session][U17] 在 OFF 时检查 adapter.canStart；已有会话升级 continuous 不重新采样，普通 activate 取消已有会话。其 apply 检查 adapter、unit 与 applying，没有同样的 DOM 输入判断；固定低层 Sheets apply 也没有本项目的合并确认扩展，不能把新增 UI 判断描述为所有上游程序入口的共同限制。

[use-sheet-format-painter.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-format-painter.ts) 的本地 UI 上下文新增 `editing()`，由 [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 提供共享单元格/公式栏 `editingKey` 与 `editComposing`。OFF 时的新采样依该状态禁用；UI apply、pointer selection 完成后的 apply 与 confirm 在输入或组字时返回 false，保留已有源快照和确认请求。已激活后的 continuous 升级、取消刷子和正确归属的取消确认仍可执行，不删除正在输入时的原样本。

`executeRibbon` 先前在分派格式刷前调用 `finishEdit`，会提交输入草稿，使后续 UI 判断失去当前输入状态。本批仅让原格式刷命令表绕过这个自动提交，其他命令沿原路径；成功后的 Canvas 焦点恢复复用 `editingKey/editComposing`，输入期间不会强制切回网格。程序化 OfficeStudio 激活/应用保留原行为；原“输入中程序化激活、composition Esc 保留草稿，取消编辑后再 Esc 取消刷子”浏览器草稿保留。

[SheetMergeConfirmationDialog.vue](../../packages/office-studio/src/components/spreadsheet/SheetMergeConfirmationDialog.vue) 新增可选 UI 属性 `disabled?: boolean`，只传给 `Modal.okButtonProps.disabled`。全部调用方是 ExcelWorkbench 的普通 merge 与格式刷确认两个实例：仅格式刷传 `Boolean(editingKey) || editComposing`，普通 merge 省略属性，取消按钮沿原行为。`afterPainterClose` 在原 `nextTick` 中直接复用 `context.editing()`，取消关闭期间仍在输入或组字就不恢复 Canvas 焦点。本地已安装 `antdv-next` 1.8.2 的 `dist/modal/interface.d.ts:83` 确认 `okButtonProps?: ButtonProps`；[官方 Modal 文档][U18] 经 Context7 返回当前版，1.8.2 尚未索引，不能将当前文档视为固定版证据。

[sheet-format-painter-commands.ts](../../packages/office-studio/src/api/sheet-format-painter-commands.ts) 的 cancel-confirmation 复用 cancel / set-OFF 的显式 unitId 归属语义：指定其他工作簿时返回 false，不清确认、源或内容历史；省略 unitId 的原全局取消确认用法保留。standalone UI 同样判断取消确认的源工作簿。正确取消只释放当前确认请求，ONCE / continuous 样本仍可用于下一目标；版本、源/目标权限、合并值验证、候选事务、selection undo/redo 和 FWorkbook.save / 内部重开均复用既有实现，没有改写模型或历史协议。

本批生产增量为 UI hook `+9/-3`、ExcelWorkbench 接线 `+6/-3`、API `+4/-1`、共享确认组件 `+2/-2`，共 `+21/-9`；另修改两份原草稿和本页，共七文件。内部 UI context 新增 `editing()`，共享 UI 组件新增可选 `disabled`；`sheet-format-painter.ts`、共享 rich-text helper、先前四个 Clipboard 文件、公式文件、DTO / 公开 API / 公共命令 ID / 事件字段 / 保存 / Worker/WASM 协议均无本批改动。

[合同草稿](../../packages/office-studio/tests/sheet-format-painter.contract.test.mjs) 新增 1 个声明，直接调用真实 UI hook、Pinia 会话、领域应用与 OfficeStudio：production / standalone 分别检查编辑与组字时各激活 alias/status 的拒绝、pointer/apply/confirm 不改变候选/版本/历史、错误工作簿取消确认、连续升级、正确取消、输入结束后合并确认与生产 undo/redo/save/内部重开；standalone 检查原 view 发布与 JSON 保存读回，不伪造 OfficeStudio 历史。既有 API 取消确认用例补错 unitId、错 requestId 与内容/版本/历史不变。

[浏览器草稿](../../packages/office-studio/tests/sheet-format-painter.browser.test.mjs) 新增 1 个声明，循环 production / standalone × 单元格 / 公式栏，通过原 `antdv-next` Input / TextArea 和格式刷按钮检查 OFF 禁用、编辑草稿、合成 composition、明转暗、取消活动刷子不提交草稿/不新增值事件或历史、不强制 Canvas 焦点，取消输入后恢复网格焦点；四个循环样本仍只计 1 个声明。既有合并取消用例补错误工作簿调用后对话框、源、revision、内容与历史保留。fixture 沿原接口，没有新增产品 UI 控件或原生桥，统一 ConfigProvider / 主题 token 保留。

同一浏览器声明另循环 production / standalone：打开合并确认后，向未聚焦的真实公式 TextArea 派发合成 composition，检查“继续合并”禁用、取消仍可用、明转暗仍禁用、迟到的原生 `click()` 不改变内容/版本/源/历史；composition 结束后恢复按钮并确认，生产只新增一项内容历史，standalone 只发布一次 view。重新打开确认后在合成 composition 期间取消，等待两帧检查 `after-close` 未强制网格焦点、源和内容历史保留。现有 hook 合同继续覆盖 editing/composing 时直接确认被拦截；这里的禁用按钮 `click()` 不证明 hook 收到迟到的 `@ok`，也未覆盖真实输入焦点与已开 Modal 同时编辑或 Windows IME，不能据合成事件宣称这些场景已通过。

实际 `rg -c '^test\('` 为合同 **61**、浏览器 **57**，这两文件合计 **118** 个声明（含既有条件格式用例），运行均为 **0**。本批仅完成源码/diff 阅读与声明计数，统一静态语法、引用、链接/锚点和空白检查由主线收尾；没有运行产品、测试、类型检查、构建、Rust/WASM、dev、浏览器或基准。草稿不证明真实 Windows IME、系统焦点/剪贴板、两种渲染后端、Office 独立写出重开或性能，完整格式刷与 SHEET-04 / SHEET-14 仍保持 **未验收**。

[U1]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/schema.ts
[U2]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/format-painter.menu.ts
[U3]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/controllers/format-painter.controller.ts
[U4]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/services/format-painter/format-painter.service.ts
[U5]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/controllers/cf.painter.controller.ts
[U6]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting/src/services/conditional-formatting-range-transform.service.ts
[U7]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting/src/models/conditional-formatting-rule-model.ts
[U8]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/menu/schema.ts
[U9]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting/src/commands/commands/set-cf.command.ts
[U10]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/views/panel/RuleEdit.tsx
[U11]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/views/panel/RuleList.tsx
[U12]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/menu/manage-rule.ts
[U13]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/format-painter/format-painter.controller.ts
[U14]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/basics/cell-style.ts
[U15]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/menu-util.ts#L163
[U16]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/clear.menu.ts#L59
[U17]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/services/format-painter/format-painter-session.service.ts#L101
[U18]: https://github.com/antdv-next/antdv-next/blob/main/docs/src/pages/components/modal/index.en-US.md
