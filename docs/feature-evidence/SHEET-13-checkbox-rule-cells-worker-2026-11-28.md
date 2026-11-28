# SHEET-13 复选框规则与单元格联合事务、Worker/Rust 计划

续接记录日期：**2026-11-27（Asia/Shanghai 客户端日期）**；沿用前批文件名日期。固定 Univer **v1.6.0**，HEAD **eaf4cfa04fabdbcfc922c692f3450459f8a29546**。本阶段实现、草稿和静态检查已经写入，运行验收全部后置，状态 **未验收**。

入口、分组、grid/classic 位置、条件和命令的清点沿用 [复选框入口](SHEET-13-checkbox-entry-review.md)；画布单格切换、反馈和公开 dropdown operations 沿用 [前段证据](SHEET-13-checkbox-feedback-dropdown-2026-11-28.md)，本文件不重复清点入口或用按钮计数证明完成。

## 来源与语义

[getDataValidationDiffMutations](../../../univer/packages/sheets-data-validation/src/commands/commands/data-validation.command.ts) 在新增/扩展 checkbox 时填空白格的未勾选值；修改自定义操作数时迁移匹配的旧勾选/未勾选值，规则与内容一起进入 undo/redo。[CheckboxValidator](../../../univer/packages/sheets-data-validation/src/validators/checkbox-validator.ts) 使用缺省 **1/0** 及完整 TRUE/FALSE 比较别名；[CheckboxRender](../../../univer/packages/sheets-data-validation-ui/src/views/widgets/checkbox-widget.ts) 只切换实际主格。本地保留这些可观察契约，范围计算和内存所有权按本项目分层改写。

[validation-checkbox-rule-cells.ts](../../packages/office-studio/src/features/sheet-validation/validation-checkbox-rule-cells.ts) 共享准备和 scalar action 分类：新规则、范围变化或操作数变化填空 scalar；旧规则覆盖内只迁移匹配状态；不匹配文本、公式、富文本和图片保持原数据。改变提示字段或完全不变规则不额外物化空格；clear/remove 不改 scalar。公式操作数按保存 anchor 经共享 reader 求值，完整目标/合并区域及源 View 权限先于目标 scalar 读取；无法读取新操作数时拒绝整个事务。

本地差异：公式/富文本/图片不被规则 authoring 强制转换为 checkbox。范围先做并集并裁剪到真实 worksheet dimensions，合并格只物化规则范围内的实际主格；只选择合并成员不越界写入外部主格。需用独立 Office 样本和固定版行为验收这些保真选择，未宣称内部实现完全相同。

## 分层、协议与资源

模块按 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture) 和 [WASM 平台契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract) 归属：领域模型保存唯一内容；validation feature 判断 scalar/规则语义；compute Worker/Rust 规划 bulk 几何和写入动作；`office-wasm-bridge` 负责控制 DTO、独立 TypedArray copy 和宏任务让出；Vue/Ant 控件负责进度、取消和编辑 owner。

[checkbox-cell-protocol.ts](../../packages/office-studio/src/compute/checkbox-cell-protocol.ts) 的任务域为 `checkboxCells`，保留协议 2；[TS 平台枚举](../../packages/office-studio/src/platform/wasm/task-protocol.ts) 和 [Rust TaskKind](../../crates/office-wasm-bridge/src/control.rs) 同步增加该域。

| bulk/control 字段 | 契约 |
| --- | --- |
| `ranges` / `merges` | 独立 `Uint32Array`，四元闭区间 `startRow,endRow,startColumn,endColumn`；规则范围并集互不相交 |
| `positions` / `actions` | 独立位置 pairs 和 `Uint8Array`；`0` 保留、`1` unchecked、`2` checked；缺失位置隐式为 `1` |
| `params` | typed DTO：rowCount/columnCount/rangeCount/mergeCount/existingCount/coveredCells/maxOutputRuns |
| `runs` / `changedCells` | 独立 `Uint32Array` 五元压缩矩形加动作；变化数量匹配面积总和，未完成计划没有部分结果 |
| task control | requestId/unitId/baseRevision/kind/budget；权限和用户基线由 host 捕获并在读取、分块及发布前重查 |

[checkbox.rs](../../crates/office-compute/src/checkbox.rs) 分步建立稀疏动作索引、裁剪合并成员、扫描并压缩行区间。每次 `step` 限制工作项；range/merge 构造校验和必要 copy 在 Worker 中同步执行，此段不能接收普通 cancel 消息，仍由现有客户端超时/终止重建策略兜底。输出或碎片超过预算明确失败，不发布半成品。

[checkbox_wasm.rs](../../crates/office-compute/src/checkbox_wasm.rs) 实际调用共享 `from_control/to_control`（serde-wasm-bindgen）、`copy_indices_in/copy_bytes_in/copy_indices_out`（js-sys TypedArray 和受控 lease）。[compute.worker.ts](../../packages/office-studio/src/compute/compute.worker.ts) 实际调用 `compute_yield()`，沿 Rust Future/Promise + web-sys 宏任务适配让出，接收取消并 finally 释放 stepper/result。领域算法不依赖 DOM。

当前任务上限沿既有 compute 的 16,000,000 元素、worksheet 1,048,576 × 16,384 dimensions 和 4,096 个范围；本 feature 输出上限暂设 65,536 runs，默认 admission 256 MiB。4,096 cells 以下可沿共享同步 planner，较大区域沿 Worker/Rust；这是暂定有界策略，**未测量性能交叉点**。独立选区规模、格式/metadata 与 materialization budget 可能触发明确资源错误，不能把资源拒绝当作全量大表验收。

TS/Rust admission 估算包括 typed 输入、WASM 入站副本、稀疏索引、几何片段及独立输出。JS→WASM 每个非空 bulk 输入记一次真实 copy；WASM→Worker JS `runs` 记一次实际出站 copy，随后只 transfer 独立 buffer。4 个输入 transfer 含真实空数组，空数组不冒算入站 copy；control DTO 计量至少入/出各一次，无 JSON bulk 和 Rust buffer clone。几何转换、JS representation 准备、领域对象创建与内容历史另外记账；**不是端到端零拷贝**，不转移 linear-memory buffer。

[validation-checkbox-rule-cells-async.ts](../../packages/office-studio/src/features/sheet-validation/validation-checkbox-rule-cells-async.ts) 在 host 分块枚举/分类现有键，Worker 输出回来后每 256 个目标分块创建 detached cells。`maxMaterializationBytes` 是按序列化长度、对象及历史副本估算的 admission，包含保守固定余量；它不是浏览器实际峰值测量，也没有计入既有全工作簿常驻模型/视图。规则命令保留的完整 JSON 基线检查和最终 commit/视图准备仍须统一测量，当前不能声称整个主线程已经避免长任务。

## 命令、UI 与持久化

[validation-rule-command.ts](../../packages/office-studio/src/features/sheet-validation/validation-rule-command.ts) 增加 `SHEET_DATA_VALIDATION_ASYNC_COMMANDS.set/replace`，分别为 `sheet.command.set-data-validation-async` 和 `sheet.command.replace-data-validation-async`。保留同步命令；异步 payload 为原 `target/rule` 加顶层 `signal/onProgress/maxMaterializationBytes`，返回 `Promise<boolean>`。注册由 [sheet-data-validation-commands.ts](../../packages/office-studio/src/api/sheet-data-validation-commands.ts) 接收 Studio 私有 compute client 的依赖注入；共享 Studio/Facade/根导出由主任务协调接线。

`set/replace` 先规划规则、检查原/新完整 View/Edit，再规划 cells；准备、mutator 和 prepareView 重查内容、用户、history、规则与宿主。规则与 cells 同一次 `commitWorkbook`，一条文档隔离历史；仅有实际 scalar 变化时发布一次 `SheetValueChanged`，包含实际变化矩形。取消、失活、版本改变或资源不足保留原内容；clear/remove 不发布虚假的值事件。内部 save/reopen 草稿覆盖本地模型字段，Office 文件保真仍复用既有 parser/writer 链且未验收。

[use-sheet-data-validation-workbench.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-data-validation-workbench.ts) 的 `submitAsync/pending/progress/cancel` 保留完整 panel owner；[use-sheet-quick-validation-workbench.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-quick-validation-workbench.ts) 给无侧栏的快速规则建立独立异步 owner。用户离开工作台/表或组件销毁时撤销在途任务，旧完成回调不回焦或覆盖新上下文。[SheetDataValidationPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataValidationPanel.vue) 与 [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 使用 Ant Spin/Progress/Button，取消按钮在禁用 Form 外。真实阶段和已处理数量来自枚举、分类、WASM 与物化，不用 timer 伪造进度；现有 ConfigProvider/theme token 继续管理明暗主题。批量状态新增 [19 locales × 5 keys](../../packages/office-studio/src/i18n/sheets/locales/validation-input.ts)，validation 专项共 42 keys/locale。

## 草稿与实际静态检查

| 草稿 | 声明 / 计划场景 | 覆盖 |
| --- | --- | --- |
| [sheet-checkbox-rule-cells.contract.test.mjs](../../packages/office-studio/tests/sheet-checkbox-rule-cells.contract.test.mjs) | 18 / 29 | 初始化/值迁移/合并/公式源/权限/取消/单历史/重开/非活动工作簿 |
| [sheet-checkbox-bulk-protocol.contract.test.mjs](../../packages/office-studio/tests/sheet-checkbox-bulk-protocol.contract.test.mjs) | 6 / 24 | typed 形状/ownership/detached/budget/复制计量/actual client/取消与旧结果 |
| [sheet-checkbox-bulk-rule-cells.contract.test.mjs](../../packages/office-studio/tests/sheet-checkbox-bulk-rule-cells.contract.test.mjs) | 8 / 12 | 固定结果 transport、共享分类、分块物化、源 View、资源拒绝、真实 async command/history/重开 |
| [sheet-checkbox-feedback.browser.test.mjs](../../packages/office-studio/tests/sheet-checkbox-feedback.browser.test.mjs) | 19 / 42 | 已有 Canvas/Ant/feedback/dropdown 加真实 Worker/WASM 4,800 格初始化、computing/materializing 取消、classic/grid 历史与重开，以及 INDEX/numfmt 的单选/多选文字 |
| [sheet-checkbox-bulk-workbench.contract.test.mjs](../../packages/office-studio/tests/sheet-checkbox-bulk-workbench.contract.test.mjs) | 7 / 14 | 真实 Studio/Vue effectScope、panel/quick pending/进度/双提交/取消重试、owner 更换、失活及 scope disposal |
| [Rust 算法草稿](../../crates/office-compute/src/checkbox/tests.rs) | 8 `#[test]` | 独立小网格 oracle 的 81 个 action pattern、步长、合并/并集、预算、非法输入、部分 owner 释放 |
| [真实 Worker binding 草稿](../../crates/office-compute/tests/wasm_checkbox.rs) | 3 wasm-bindgen-test | typed DTO、实际 copy/lease、memory.grow/release、别名/预算/未知字段及 Future 宏任务让出 |

联合事务首次静态快照为 **67 文件、67 scripts、7 Vue templates、7 styles、577 本地 imports 和 22 次 Node 语法解析**，其中 **21 个 JS 草稿文件、256 声明、435 计划场景，运行 0**。后续 INDEX/numfmt 与 bulk workbench 草稿的实际静态快照为 **68 文件、68 scripts、7 templates、7 styles、592 imports、23 Node parses**，其中 **22 个 JS 草稿文件、271 声明、467 计划场景，运行 0**；后续细项见 [INDEX/numfmt 专项](SHEET-13-list-source-index-numfmt-2026-11-28.md)。只做 TS/SFC/PostCSS/import/Node parse；没有 TypeScript 类型检查、Rust/WASM 检查或运行测试。静态数量是相应覆盖集合，不是新增功能或通过场景数量。

## 仍需实施与验收

同步 Facade 大范围调用的调度/兼容策略、实际任务/物化/完整 JSON 基线/commit/视图成本必须在同数据集的 TS 与 Worker/Rust 端到端基准中核对。首屏、帧时、长任务、真实峰值内存和取消时延没有实测；协议中的计量字段和期望断言不代替真实报告。统一验证先按项目脚本生成匹配的新 WASM，再运行 native/binding/contract/browser、独立 Office 往返和基准。

SHEET-13 的 INDEX authoring 和候选 numfmt 已在 [后续专项](SHEET-13-list-source-index-numfmt-2026-11-28.md) 落盘且未验收；完整数组/spill、大列表源 Worker 和标签完整自动尺寸/旋转/缩进仍待推进。本批不替代全目标。真实 Windows IME、主题/焦点、两渲染后端、权限/协作/结构编辑和 Office 独立样本仍保持未验收。
