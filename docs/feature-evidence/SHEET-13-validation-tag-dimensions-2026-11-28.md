# SHEET-13 / SHEET-03 验证标签、候选颜色与自动尺寸

续接记录日期：2026-11-27。固定 Univer v1.6.0：eaf4cfa04fabdbcfc922c692f3450459f8a29546。状态：**开发中，未验收**。本记录承接[日期与列表输入](SHEET-13-validation-cell-input-2026-11-28.md)、[复选框与反馈](SHEET-13-checkbox-feedback-dropdown-2026-11-28.md)、[列表源 INDEX 与数字格式](SHEET-13-list-source-index-numfmt-2026-11-28.md)，不替代完整 Sheets / Docs / Pro 目标。

## 1. 固定源码与入口

| 固定来源 | 已核实的可观察契约 | 本次本地处理 |
| --- | --- | --- |
| [ListValidator][list-validator] skipDefaultFontRender/getExtraStyle/getOptionColor | 单选 default/custom 胶囊；arrow/text 保留普通文字并使用所选候选底色；OVERFLOW 转为 CLIP | 胶囊与普通文字分路；候选底色属于派生层，不改作者 fill；列表格不向邻格溢出 |
| [ListMultipleValidator][multiple-validator] | 多选 CSV 存储/读取，始终跳过默认 font，由独立 widget 绘制 | 保留引号、逗号、换行、无效现有值；本地 text 模式的普通文字是现有本地扩展 |
| [DropdownWidget][single-widget] calcCellAutoHeight/Width | 文字宽度扣除 icon/padding；自动高使用无限高度；text 模式交还普通文字测量 | 单选 arrow 普通文字绘制与 auto-fit 使用相同尾部槽；标签使用共享布局 |
| [DropdownMultipleWidget][multiple-widget] calcCellAutoHeight/Width | 两个自动尺寸入口共用 layoutDropdowns | 多选绘制和自动高/宽共用一份几何 |
| [shape/layout.ts][tag-layout] layoutDropdowns | item 按当前宽度换行；自动宽采用最大布局行宽度 | 保留 item 换行及最大行宽；补不同标签高度的同行排列 |
| [shape/dropdown.ts][tag-shape] | 标签底色、文字裁剪和装饰线 | polygon/text/line 使用同一 affine matrix，文字另有局部裁剪 |
| [DataValidationRenderController][dv-render] | AutoHeight/Width interceptor 与规则修改后的重算 | 添加/删除/范围和模式变化失效旧、新范围内的稀疏已填充行 |
| [DataValidationListCacheService][list-cache] | 按 unit/sheet/rule 缓存候选、颜色/索引；规则和公式结果变化失效 | 尺寸批时仅有单次操作缓存；后续[内联持续缓存与查询](SHEET-13-validation-list-cache-query-2026-11-28.md)已落盘，动态依赖 epoch 和来源求值仍未完成 |

本批没有顶部 Ribbon 修改。常驻根、扩展根、More 溢出根、上下文根的**入口增量分别为 0**。新增 helper、标签数和草稿数均不计为功能完成数。原验证管理、快捷下拉和日期的入口来源仍见[验证入口](SHEET-13-data-validation.md)、[快捷入口](SHEET-13-quick-dropdown-date.md)。原 sheet.command.set-row-is-auto-height、sheet.command.set-col-auto-width、手动行高列宽及别名继续走共享命令和内容历史。

## 2. 已落盘的本地实现

- [validation-tag-layout.ts](../../packages/office-studio/src/features/sheet-validation/validation-tag-layout.ts) 是唯一标签布局。CSV 标签使用候选的 canonical spelling/颜色；未知现有值仍可见。单选 arrow/text 交还普通文字，default/custom 使用标签，多选保留逐 item 换行。
- 长标签内部换行服从单元格 wrap，显式换行总是保留。竖排保持完整 grapheme。旋转、缩进、居中不缩进和 shrink-to-fit 使用共享字体/几何；自动高忽略固定高度，自动宽不缩小、也不在标签内部换行。
- [validation-cell-tags.ts](../../packages/office-studio/src/rendering/spreadsheet/validation-cell-tags.ts) 用现有 matrix/localClip 生成共享 Canvas2D/WebGPU 命令。标签底色、文字、下划线和删除线整体旋转/缩小；只为相交标签和文字 fragment 发出命令，固定行高通过裁剪保留完整自然布局。
- [validation-cell-options.ts](../../packages/office-studio/src/features/sheet-validation/validation-cell-options.ts) 只对请求主格读取格式化候选。合并范围及全部来源均检查 View；同一规则在一次读取中求值一次，输出只有 label/color，不携带 workbook、源 DTO 或规则。显式空候选集合不回退为引用源颜色。
- [use-sheet-validation-feedback.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-validation-feedback.ts) 的 read/readTagOptions 共用已有前后 owner、模型、revision、内容历史 generation、用户和区域检查。前者保持 popup blocker，后者允许画布在弹窗打开时保留候选颜色；读取中切换上下文会丢弃完整结果。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) → [SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) → [scene.ts](../../packages/office-studio/src/rendering/spreadsheet/scene.ts) 已传候选 reader。反馈与标签候选共用可见格/选中格坐标；scene 只消费候选结果，不新增列表公式求值。单选 arrow/text 颜色置于作者 fill 之后、现有条件格式 fill 之前；此顺序是本地选择，尚未在两后端验收。
- [sheet-auto-fit.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-auto-fit.ts) 标签和普通文字通过 sheetViewCellAt 测量有效继承样式。arrow 普通文字扣相同尾部槽；零宽度的列可重新自动适应，手动/大纲隐藏原因仍按共享可见性契约排除。
- [sheet-axis-dimensions.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-axis-dimensions.ts) 包含作者样式、目录引用、clear 属性、themeStyle、类型、缩进/对齐等尺寸签名。workbook 默认样式、目录、范围主题和表过滤造成的 zebra 样式变化会重算；引用列表源变化会失效相应目标范围。未改变的共享 worksheet 在构建完整 cell view 前跳过；确需修改其尺寸时先复制 worksheet，保护旧内容历史对象。
- [use-sheet-axis-dimensions.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-axis-dimensions.ts) 捕获用户、区域及同步上下文 generation。源 getter、候选测量或 prepareView 中的切换会在发布前丢弃候选；applyView 后的切换阻止旧通知，保存的 current 回调和 nextTick 焦点检查同一上下文。用户或区域切换会清除原尺寸弹窗；切回原值不会重新授权旧回调。共享 axis host 未被本批修改。

内容事务仍是唯一权威源。标签布局、反馈、候选 Map 和渲染矩阵不写快照。自动尺寸写入已有 rowAutoHeights/rowAutoHeightSpans/columnWidthSpans，由已有文档历史、内部保存及 Office 轴适配消费。**本批没有执行内部重开或 XLSX 往返，未确认文件保真**。

## 3. 内部接口与共享接线

新增均为可选内部接口，公共 workbook schema 和 Worker/WASM 协议没有本批增量：

| 接口 | 含义 |
| --- | --- |
| SheetValidationTagOptionsReader(positions) | 请求主格的只读候选 Map |
| SheetValidationTagOptionsAt(row, column) | 场景/尺寸消费已准备候选 |
| createSheetValidationTagOptionsResolver(workbook, sheetId, userId?, region?) | 在同步快照上下文按规则缓存；调用者负责 owner/revision 生命周期 |
| SheetAutoFitOptions.styleContext/validationTagOptions | 有效 workbook 样式和候选来源 |
| SheetAxisDimensionOptions.userId | 自动尺寸的来源 View 上下文 |
| refreshSheetViewAutoRowHeights(..., previousStyleContext?, invalidatedValidationRanges?) | 比较旧上下文及引用列表源失效 |

refreshWorkbookAutoRowHeights 已为每张表取 workbook 样式/候选。主动尺寸 API 的 [sheet-axis-dimension-commands.ts](../../packages/office-studio/src/api/sheet-axis-dimension-commands.ts) 已传入 candidate model 的 styleContext、currentUserId 和候选 resolver。standalone 分支已传 candidate view 的继承样式；仅自动尺寸/grow 通过既有 adapter 准备候选领域模型及 resolver，并用本地捕获的用户、区域和 generation 保护读取、准备、发布与后续回调。上述两个 caller 已落盘，运行验收仍后置。

共享 Studio/Facade 和 adapter 文件仍由主任务协调，本子任务未覆盖在途修改。Studio 对 refreshWorkbookAutoRowHeights 的调用尚需传入 currentUserId；detached sheetViewCellAt 默认样式上下文仍缺 tables，Canvas 空白格/继承范围主题的表过滤上下文尚待补齐。不得用本批服务函数或草稿记为这些共享缺口已完成。

## 4. 尺寸与行为差异

固定 multi shape 的 cell 横/纵 padding 为 6/6、tag 为 4/0、gap 为 4、icon slot 为 14；本地 cell 为 4/4、tag 为 6/3、gap 为 4、Ant 尾部槽为 22、圆角为 4。固定 glyph 高度来自 ascent/descent，本地标签使用共享 point→CSS-pixel fontSize 行盒。这些差异使本地标签有较大的纵向留白，本批保持绘制与测量一致，不声称像素相同。

固定 widgets 未显式处理标签旋转、缩进和缩小；本地将它们作为共享单元格样式扩展。长标签内部换行、不同高度的同行居中及文字 fragment 裁剪也按本地文本布局实现。字体、缩放、冻结合并、窄列、明暗主题、两绘制后端及真实交互仍待验收。

## 5. 静态结果与未运行草稿

本批只执行读取、语法、引用、计数检查：累计 **76 文件、76 段脚本、7 个 Vue 模板、7 段样式、664 处本地引用、25 次 Node 语法解析**通过，包含两个 caller 的最新接线及 standalone 上下文保护。没有执行产品功能、测试 runner、类型检查、前端构建、Rust/WASM 检查、dev、浏览器或基准。计数只覆盖本子任务静态文件集。

| 草稿 | 声明 / 计划 | 待执行行为 |
| --- | ---: | --- |
| [标签布局合同](../../packages/office-studio/tests/sheet-validation-tag-layout.contract.test.mjs) | 14 / 25 | CSV/颜色、源读取次数、权限、owner 重入、换行、不同高度、grapheme、旋转/缩进/隐藏箭头/缩小、独立尺寸几何、可见 fragment |
| [Canvas 验证格合同](../../packages/office-studio/tests/rendering/spreadsheet-validation-cells.contract.test.mjs) | 11 / 17 | 原标签/冻结合并；引用颜色、arrow/text 文字与裁剪、隐藏箭头、离屏 overflow、View 拒绝不调用候选 reader |
| [轴尺寸合同](../../packages/office-studio/tests/sheet-axis-dimensions.contract.test.mjs) | 29 / 38 | 原轴/权限/历史；规则范围、作者/引用/clear/theme 样式、表过滤 band、历史对象保护、源变化、真实 Studio 验证命令→同一尺寸历史→内部重开→手动尺寸；真实 Facade 继承字体/旋转自动列宽独立预期 68px；standalone 引用候选/无 ID，以及 getter/测量/准备/发布/延迟焦点/弹窗的用户和区域重入 |
| [复选框/反馈浏览器](../../packages/office-studio/tests/sheet-checkbox-feedback.browser.test.mjs) | 21 / 46 | 原复选框/反馈/源/Worker；新增 classic/grid 多选行高、列宽、Ant trigger、历史与内部重开 |

累计 24 份 JS 草稿为 **319 声明 / 541 计划例，运行 0**。standalone 新草稿使用 effectScope/shallowReactive，并在 finally 中释放 watcher 与可选 Studio。test.each 与外层 classic/grid 循环分开计算；声明数不是运行或功能覆盖证明。旧 74 文件、648 引用、311/527、68 文件、592 引用、271/467 和更早 67 文件结果仍为历史快照，不覆盖本批。

## 6. 继续实施与统一验收

1. 协调共享 Studio 的 currentUserId 与 adapter 的 table context；空白范围主题在 Canvas 上的表过滤继承仍属共享缺口。主动 API/standalone caller 已接线，运行验证留待统一阶段。
2. 与唯一公式引擎协调 array/spill、外部/结构化/动态引用的完整列表源。后续[列表缓存/查询批](SHEET-13-validation-list-cache-query-2026-11-28.md)已接内联跨 revision 缓存、Rust/Worker 查询、取消/预算与 Ant 虚拟 UI；大来源求值和完整动态依赖 epoch 仍未实施，尺寸算法保持原最大布局行宽契约。上述 76 文件等数字只属于尺寸批历史快照。
3. 统一阶段执行合同、真实 Studio/浏览器闭环和 UI 核对；验证 Ant 控件主题、弹层、焦点、IME、冻结合并及两后端。本批没有新增原生可替换控件。
4. 用真实 XLSX 样本验收验证规则、颜色、行列尺寸及手动/自动模式的导出/独立重开；内部 JSON 草稿不替代 Office 验收。
5. 同一大样本记录源读取、标签布局、首屏/滚动/编辑/缩放、长任务、Worker/WASM 传输/复制和 JS/WASM/GPU 内存；未实测，不宣称优化或达标。

[list-validator]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation/src/validators/list-validator.ts
[multiple-validator]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation/src/validators/list-multiple-validator.ts
[single-widget]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/views/widgets/dropdown-widget.ts
[multiple-widget]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/views/widgets/dropdown-multiple-widget.ts
[tag-layout]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/views/widgets/shape/layout.ts
[tag-shape]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/views/widgets/shape/dropdown.ts
[dv-render]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/controllers/dv-render.controller.ts
[list-cache]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation/src/services/dv-list-cache.service.ts
