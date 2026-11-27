# SHEET-15 查找替换入口与单次替换定位

状态：2026-11-27 已静态核对固定版入口与定位契约；前批补单次替换定位参数，本续批将查找面板迁移到 Ant Form / Form.Item 并接入主题 token。契约及浏览器草稿已落盘，全部未运行；功能未验收。固定来源为 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

## 固定入口和命令

[共享 find-replace schema](../../../univer/packages/find-replace/src/menu/schema.ts) 将“查找替换”放在 `RibbonDataGroup.OTHERS`，order 0，grid r1c1、跨两行两列；classic 由同一 order 排列。[菜单 factory](../../../univer/packages/find-replace/src/menu/find-replace.menu.ts) 只在 Sheets 上下文显示；编辑器激活、工作表未聚焦或 provider 不可用时禁用。本切片固定常驻根 1、子项 0；溢出复用该根，不据此推断完整 Data 或 Pro 入口数量。

[OpenFindDialogOperation](../../../univer/packages/find-replace/src/commands/operations/find-replace.operation.ts) 的真实 ID 是 `ui.operation.open-find-dialog`；首开调用 `start()`，已打开则聚焦查找输入。替换、关闭和前后导航同文件提供独立 operation；[ReplaceCurrentMatchCommand / ReplaceAllMatchesCommand](../../../univer/packages/find-replace/src/commands/commands/replace.command.ts) 分别转入统一查找服务，全部替换有确认与成功/失败反馈。

[SheetFindModel](../../../univer/packages/sheets-find-replace/src/controllers/sheet-find-replace.controller.ts) 的 `_findInRange` 按 query.findDirection 选择行/列迭代，跳过筛选隐藏行；`_findNextMatchByRange` 使用同一方向判断当前选区后的命中。`_findNextWorksheetThatHasAMatch` 使用 `workbook.getSheetOrders()`，未命中的工作表也保留在顺序中。`replace` 只替换已有、可替换的 currentMatch，通过 SetRangeValuesCommand 保持单次内容历史；`replaceAll` 按工作表生成内容矩阵并交给 [SheetReplaceCommand](../../../univer/packages/sheets-find-replace/src/commands/commands/sheet-replace.command.ts) 合并撤销记录。

[固定 FindReplaceDialog](../../../univer/packages/find-replace/src/views/dialog/FindReplaceDialog.tsx) 与 [FindBy / FindScope / FindDirection](../../../univer/packages/find-replace/src/services/find-replace.service.ts) 只有 value / formula 两种内容模式、当前工作表 / 工作簿两种范围、行 / 列两种方向。Sheets provider 开启区分大小写与匹配整个单元格，关闭 whole-word 能力；本地提供相同选项。格式化文本属于 value 的数据来源，固定 v1.6.0 没有独立第三个 FindBy 选项。固定 `hitCell` 从工作表 cellData 的富文本或值提取内容；本地普通格使用 rawValue / value，公式格使用 cachedValue / displayValue，货币、日期及区域格式样本的实际命中/替换一致性仍待核实。

固定 SUBUNIT 查询订阅选区、内容 mutation 和活动工作表变化，以 200ms 去抖重查。本地 `openSearch` 首开捕获多格选区并保留导航前的范围来源；工作簿范围忽略该选区限制，过滤不可见或没有查看权限的工作表。两者选区更新策略存在差异，本续批保留原行为。替换继续在原提交边界检查全部目标的编辑权限和数据验证；不以按钮有可替换命中就推断权限已通过。

<a id="single-replace-target-order-review"></a>
## 本批单次替换接线

[sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 的 Data 工具组已有 `sheet.operation.find`，grid 同为 r1c1、跨两行两列；本地标签为“查找”。常驻根、子项、扩展、上下文及溢出入口增量均为 0。

[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 已通过同一 `executeRibbon` 处理查找、前后导航、单次/全部替换和 Ctrl+F / Ctrl+H。`moveFindMatch` 原来已传入查找方向与所有工作表 ID；`applyFindReplace` 在选中单元格没有命中时遗漏这两个参数。定位批只补该行，沿用 [nextSheetMatchIndex / createSheetReplacePatches](../../packages/office-studio/src/features/sheet-find-replace/sheet-find-replace.ts)，不新增算法或入口。

- 按列查找，选中未命中的 B1，A2 / B2 都命中时，单次替换定位到 B2；此前默认按行会选择 A2。
- 工作表顺序为前表 / 空表 / 后表，当前空表没有命中，两侧 A1 都命中时，单次替换定位到后表；此前从命中项推导顺序会跳回前表。

本地保留既有“选中位置未命中时直接替换下一命中”的行为；固定上游 `replace` 只替换 currentMatch，不提供这个直接定位分支。本批借鉴固定方向及工作表顺序契约，没有把本地差异称为上游的相同实现。

提交仍走既有范围权限、数据验证、隔离候选、公式刷新和 `commitHistory`，再经 [excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) 与工作簿模型发布。未修改模型、公开 Facade、命令 ID、Worker/WASM 协议或内存生命周期；同步查找/替换计算保持原形态。

## 草稿与剩余缺口

[sheet-find-replace.contract.test.mjs](../../packages/office-studio/tests/sheet-find-replace.contract.test.mjs) 在原真实查找、定位及 patch 实现上追加上述两个未命中目标回归，当前 7 个声明、7 个计划用例。

[sheet-range-editor.browser.test.mjs](../../packages/office-studio/tests/sheet-range-editor.browser.test.mjs) 复用原真实 OfficeStudio / Workbench、seed、选区、主题、历史和 JSON 保存/重开能力。定位批追加两条 `test.each(['classic', 'grid'])`，仍为 2 个声明、4 个计划用例；本续批补超链接草稿后全文件为 19 个声明位置、23 个计划用例（15 个普通 test、1 个循环声明展开 2 例、3 个 test.each 展开 6 例）。查找每例经 Data 根和原 Ant 控件，检查非目标内容不变、一次历史、撤销/重做及内部重开；主题切换检查输入草稿与焦点。按列草稿改用替换 Input 的 Enter 和查找 Input 的 Escape，跨表草稿保留替换/关闭按钮；未新增 fixture 或产品测试接口，保留前批全部范围编辑及保护草稿。

<a id="find-form-review"></a>
## 查找 Form 与主题续修

本续批只将原两行查找面板的根改为 antdv-next Form，七个现有控件（2 个 Input、3 个 Select、2 个 Checkbox）各使用 FormItem noStyle；复用已安装 1.8.2 的 API，保持原 DOM 布局、ARIA 名称、keydown、Enter 导航/替换和 Escape 关闭。不增加字段校验、模型包装、按钮或业务接口；Form 的 submit.prevent 阻止原生提交。

面板背景、文本与分隔线复用已有 sheetThemeToken 的 colorBgContainer / colorTextSecondary / colorBorderSecondary，输入、选择、勾选和按钮继续通过原 ConfigProvider 管理主题及交互状态。原按列草稿增加实际 Ant Form 根和明暗背景差异断言；这些断言没有运行，不能称为主题或焦点已经验收。

本续批仅复读源码、diff 和文本计数；测试、类型、构建、dev、浏览器、Rust/WASM、文件样本与基准全部后置。完整结果交互、Worker 大样本与取消仍有缺口；value 模式中格式化文本的样本行为、本批浏览器交互、两渲染后端、真实 IME、Office 文件往返和性能均未验收。内部 JSON 重开草稿不代表 Office 互操作已通过。
