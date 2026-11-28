# SHEET-13 INDEX 列表创建与数字格式候选

续接日期：**2026-11-27（Asia/Shanghai 客户端日期）**；文件名沿用前批日期。固定参考为 Univer **v1.6.0**，HEAD **eaf4cfa04fabdbcfc922c692f3450459f8a29546**。已落盘的实现、未运行草稿和实际静态检查分别记录，全部行为仍为 **未验收**。

本续段沿用 [顶部入口和列表规则清点](SHEET-13-quick-dropdown-date.md)、[共享动态来源](SHEET-13-checkbox-feedback-dropdown-2026-11-28.md) 与 [规则事务](SHEET-13-checkbox-rule-cells-worker-2026-11-28.md)，没有新增 Ribbon 根或商业依赖。

## 固定来源与差异

[list-validator.ts](../../../univer/packages/sheets-data-validation/src/validators/list-validator.ts) 的 `supportedFormula/isValidListFormula` 在创建时只识别 IF、INDIRECT、CHOOSE、OFFSET 四个函数节点，或直接区域引用；[固定测试](../../../univer/packages/sheets-data-validation/src/validators/__tests__/list-validator.spec.ts) 覆盖 reference/IF/OFFSET 并拒绝 SUM。该白名单不能证明共享引用求值器只支持四函数。本地此前已经能读取 INDEX reference，本续段在 [validation-engine.ts](../../packages/office-studio/src/features/sheet-validation/validation-engine.ts) 的创建语法判断中加入 INDEX，允许 `=INDEX($D$1:$D$4,0,1)` 经面板和真实规则命令创建单选或多选规则。引用、错误、选择器、名称和权限仍复用共享 formula engine；没有另写 INDEX evaluator。**这是本地明确的创建扩展，与固定四函数白名单存在差异**。Office 样本及线上 INDEX 创建行为尚未独立验收，不把此改动冒称固定版 Univer 的已有 authoring 行为。

固定 [list-validator.ts](../../../univer/packages/sheets-data-validation/src/validators/list-validator.ts) 的 `getRuleFormulaResultSet` 和 [dv-list-cache.service.ts](../../../univer/packages/sheets-data-validation/src/services/dv-list-cache.service.ts) 的 `_getRuleFormulaResultSet` 对非 string 且带 number format 的值调用 `numfmt.format`。缓存保存最终 `list/listWithColor/set/colorMap`，颜色按去重后的候选位置映射。`ListValidator.isValidType` 将目标实际值转为文字后与候选做忽略大小写的比较；[dropdown-manager.service.ts](../../../univer/packages/sheets-data-validation-ui/src/services/dropdown-manager.service.ts) 将 `item.label` 同时作为 label/value，`handleSave` 写入所选文字。没有把候选文字反解析后修改来源单元格。

本地采用同样的可观察值关系。例如源值 `0.25`、源格式 `0%` 生成 `25%`；选择控件提交文字 `25%`，源数字及其缓存不变。格式化来源不额外创建 raw-value alias：目标原始数字 `0.25` 即使显示为 `25%`，也按固定验证路径的 `String(actual)` 比较。多选保留本项目已声明的 CSV 存储，包含逗号的 regional label 会正确加引号；固定版所用 JSON list 序列化仍是已记录的差异。

## 实现边界

[validation-engine.ts](../../packages/office-studio/src/features/sheet-validation/validation-engine.ts) 在同一 permission-aware reference 结果内按稀疏单元格行优先顺序读取当前 scalar。非字符串值使用 [sheet-style-inheritance.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-style-inheritance.ts) 的实际格式继承，以及 [sheet-number-format.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-number-format.ts) / [number-format-render.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-render.ts) 的唯一格式化路径。日期系统来自 workbook；地域通过可选 `region` 传入，不把 stale `display` 当成候选。源字符串保持原文字，不按数字格式强制数值化。

本地默认 General 保留既有 `String(scalar)` 路径；有意义的格式才交格式引擎。不可解释的导入格式返回原始 scalar 文字，并保留源格式/显示/值，不伪造当前格式化结果。这些与固定分支的具体错误格式/显式 General 处理仍需样本核对。源 View 和完整 merge 先于 value/style getter；隐藏 source 或嵌套公式依赖不输出候选。

[validation-list-input.ts](../../packages/office-studio/src/features/sheet-validation/validation-list-input.ts) 将 reference 颜色按最终候选索引映射到 Ant 标签，保持第一项的大小写去重行为；此前 reference source 的颜色被遗漏，现已接入。`sheetDataValidationOptions` / `validateSheetDataValidationChanges`、`sheetValidationListOptions` / `sheetValidationCellFeedback` 均增加末尾可选 `region`，旧调用缺省仍为 zh-CN。typed widget/no-op 校验、Canvas popup、反馈和规则预览使用同一地域与 source reader。widget target 的 `userId` 是内部验证上下文，未进入持久快照或 Worker DTO。

共享 [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 和 [facade.ts](../../packages/office-studio/src/api/facade.ts) 的候选/验证调用由主任务接入 `regionalSettings.region.value`；单格输入、自动填充、剪贴板和范围 editor validator 的 source policy/地域沿共享路径协调。 [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的批量准备和重填验证也传入地域/用户 reader，`useSheetImages` 上下文接同一 `officeRegion`，standalone 图片替换/删除的数据验证不再遗失当前地域。后续必须统一核对全部编辑路径，不能由一个控件草稿推断全局一致性。

Canvas 仍绘制内容，Vue/`antdv-next` 的 Form/Input/Checkbox/Tag/Popover 负责可替换控件；没有以 Table、整页 contenteditable 或额外 native UI 替代编辑器。主题沿 ConfigProvider/token，真实明暗、回焦和 IME 尚未运行验收。

## 未运行草稿

- [sheet-validation-source.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-source.contract.test.mjs)：纯 INDEX 面板创建/再读、单/多选 source、regional decimal/percent、字符串不数值化、当前公式/旧 cache、去重颜色、1900 特殊日/1904 epoch、继承格式以及源 value/style getter 权限顺序。
- [sheet-data-validation-commands.contract.test.mjs](../../packages/office-studio/tests/sheet-data-validation-commands.contract.test.mjs)：真实 INDEX set、候选/颜色、选择命令、源编辑、文档隔离历史、内部保存重开与 hidden-source no-op 拒绝。
- [sheet-validation-input.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-input.contract.test.mjs)：真实 regional typed 选择、一次值事件、no-op、undo/redo、源数字不变与独立内部重开。
- [sheet-checkbox-feedback.browser.test.mjs](../../packages/office-studio/tests/sheet-checkbox-feedback.browser.test.mjs) 和 [fixture](../../packages/office-studio/tests/fixtures/sheet-validation-cell.html)：classic/grid × single/multiple，Ant Checkbox/Tag、深色主题、含逗号数字文字、历史与内部重开；该草稿强制 Canvas2D，尚未运行，不覆盖 WebGPU 或真实 Windows IME。

新增 [sheet-checkbox-bulk-workbench.contract.test.mjs](../../packages/office-studio/tests/sheet-checkbox-bulk-workbench.contract.test.mjs) 复用真实 Studio 与 Vue effectScope，固定结果 transport 仅控制异步时机，覆盖 pending/进度/双提交/取消重试、panel/quick 失活、sheet/workbook/user/input 更换和 dispose。草稿编写中发现并修正旧取消任务晚到 progress 覆盖同 owner retry，以及 quick 离开再返回恢复资格的问题；取消绑定提交时的 panel owner，不能取消后出现的另一 owner。没有执行 runner 或接受草稿数量作为运行证据。

## 实际静态检查

本续段只读解析快照覆盖 **68 文件、68 scripts、7 Vue templates、7 styles、592 imports、23 Node parses**；**22 个 JS 草稿文件、271 声明、467 计划场景，运行 0**。其中 source 为 **14/33**、data-validation commands **14/24**、validation input **20/41**、checkbox-feedback browser **19/42**、bulk workbench **7/14**，分别为声明/计划而非通过数量。locale 集合为 **19 × 42 keys**。仅解析 TS/JS/SFC/PostCSS、核对本地引用、文档来源/链接及限定 diff；没有测试 runner、类型检查、Rust/WASM 检查、构建、dev、浏览器或性能验收。

## 后续缺口

完整数组/spill 和高级引用仍属于共享 formula engine。后续[标签与尺寸](SHEET-13-validation-tag-dimensions-2026-11-28.md)以及[内联缓存/Rust 查询/Ant 虚拟 UI](SHEET-13-validation-list-cache-query-2026-11-28.md)已落盘；大来源的 Worker/Rust 求值与完整动态依赖缓存仍未完成。此段 scalar 格式化沿现有 TS 引擎，没有测量大来源性能；不得据此声称主线程、大样本或端到端性能达标。上述 68 文件、271/467 和 19×42 keys 是来源批历史快照，不覆盖后续追加。真正的 Office 导出/独立重开、两渲染后端、结构编辑/协作权限、真实地域切换/输入法/焦点及可复现性能证据全部仍未验收。
