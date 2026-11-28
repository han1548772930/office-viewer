# SHEET-09 函数自动引用和批量插入

状态：生产链路已落盘，未验收。续接原线程 `01a220b5-b017-7723-b522-04979a516e34`；固定只读 Univer v1.6.0 为 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。原“较大改动待确认”方案由用户持续的全量修改授权覆盖，本记录替代该方案的未实施描述。

## 固定源码与可观察行为

- [InsertFunctionOperation](../../../univer/packages/sheets-formula-ui/src/commands/operations/insert-function.operation.ts)：公开 ID `formula-ui.operation.insert-function`，参数 `{ value }`。单格或单个无数值的二维选区进入未闭合公式编辑；引用优先上方、再左方。找到最近数值块后向前连续扫描，序列末端至目标之间的文本/空白仍属于引用终点，不能只截取数值单元格。
- [InsertFunctionCommand](../../../univer/packages/sheets-formula/src/commands/commands/insert-function.command.ts)：`formula.command.insert-function` 接收 `list / listOfRangeHasNumber`。无数值的一维/多选范围相对填充；有数值的行寻找结果列，多行范围按各列第一个不存在的单元格的最大行统一结果位置。存在空值但已物化的单元格仍占位。
- [快速求和快捷键](../../../univer/packages/sheets-formula-ui/src/controllers/shortcuts/quick-sum.shortcut.ts)：Windows Alt+=、Mac Cmd+Alt+=；沿固定编辑器焦点前提拒绝正在编辑/组字的会话。
- [引用序列化](../../../univer/packages/engine-formula/src/engine/utils/reference.ts)与[单元格类型](../../../univer/packages/core/src/sheets/typedef.ts)：单引用不用冒号；普通数字字符串不是数值格，富文本 body 中的数值文本按固定 `isRealNum` 语义识别。

## 本地实现与边界

[纯计划](../../packages/office-studio/src/features/formula/sheet-function-insertion.ts)生成编辑目标或两个批量列表。数值位置采用稀疏索引，复用原有引用平移器展开共享公式；没有增加第二个求值器。最多 1,000,000 个计划写入的显式任务预算会在提交前拒绝超量任务；尚未实现该计划的 Worker 大任务路径，也没有性能优化结论。

[命令与编辑器 owner](../../packages/office-studio/src/api/sheet-function-commands.ts)将固定公共 DTO 与本地 `CellRange` 归一到同一注册项。结果坐标、工作表查看/编辑、范围保护、完整候选的数据验证、模型/历史/用户版本和观察者重入均在发布前检查；所有公式经原 `commitWorkbook` 一次提交，保留样式并清理被替换的富文本/图片来源，发布真实内容范围事件。失败不提交部分单元格，不消耗原 redo 分支。

`registerSheetFunctionEditor` 显式绑定实际工作台。单格无宿主时返回 false；有宿主时恢复真实选择并启动已有输入桥，不写内容历史。多目标纯批量可直接调用公共命令，不需要伪造 UI。旧 owner 替换、工作簿关闭及 Studio 释放均调用 release。`BeforeCommandExecute` 改选区、换用户/工作簿、修改内容或重开历史时拒绝旧计划。

[ExcelWorkbench](../../packages/office-studio/src/components/ExcelWorkbench.vue)和[原 Ribbon schema](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts)将分类函数和快速求和接同一 operation；原完整参数 Form 保留为“插入函数”选择器。未闭合公式、组字、验证、确认/取消、合并锚点和焦点复用既有编辑路径。混合多选计划中的编辑草稿在其他单元格批量提交后继续保留。

固定 operation 在仅打开编辑器且没有批量写入时返回 false；本地返回实际 opened 状态 true，便于宿主判断成功。该返回值差异明确记录；编辑内容与提交/取消语义按固定行为实现。共享公式采用等价相对公式展开，不生成上游内部随机 shared formula ID；原生 Office shared-formula 编码与全部数组/spill 场景仍需单独验收。

## 数量、草稿与剩余目标

未新增函数算法或常规 Ribbon 根。唯一函数目录仍为 215 项、8 类，常用七项为同一目录的重复投影。函数菜单九根及选择器子项沿原清点口径；Sheets 72 / Docs 54 是常规根基线，不是完成数量。

- [模型契约草稿](../../packages/office-studio/tests/sheet-function-insertion-model.contract.test.mjs)：10 个声明，覆盖引用优先级、富文本身份、物化空格、结果行对齐、相对/绝对引用、多选中断及预算。
- [真实 API 草稿](../../packages/office-studio/tests/sheet-function-insertion-api.contract.test.mjs)：10 个声明，覆盖固定 DTO、内容事件、一次历史、撤销重做/内部保存重开、保护/整批验证拒绝、跨结果依赖、owner 替换、过期版本及观察者重入。
- [顶部浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)继续覆盖参数 Form、自动引用输入及批量单次撤销。

以上均未运行。测试、类型检查、构建、Rust/WASM、dev、浏览器、Office 文件与性能统一后置。完整函数目录、数组/spill、迭代计算、Pro、高级依赖与 Office 互操作继续保留全部目标；本批没有据此将 SHEET-09/10 或其他功能组标为已验收。
