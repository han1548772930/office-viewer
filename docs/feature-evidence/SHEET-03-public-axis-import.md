# SHEET-03 公开行列尺寸、自动行高与隐藏导入

记录日期：2026-11-25。固定参考根 `E:\rust_test\univer`，v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。状态：**前端已补代码，未验收**。接续[轴命令与内容历史](COMMON-03-input-selection-axis-history.md#3-行列命令与-ui)、[默认样式](SHEET-04-default-style-inheritance.md)和[公开样式读回](SHEET-04-public-style-readback.md)。按用户安排，测试、类型、构建、浏览器/dev、Rust/WASM、服务与性能检查均未执行。

本批没有增加常规 Ribbon、插件、溢出或上下文根入口；已有行列命令和输入桥消费新导入的真实状态。API 字段、用例数量与入口数量分开，不作为已完成功能总数。

## 固定来源与字段契约

下列路径相对固定参考根。

| 来源 | 已核实契约 |
| --- | --- |
| `packages/core/src/sheets/typedef.ts`：`IWorksheetData/IRowData/IColumnData` | `rowData` 的 `h/ah/ia/hd/s`、`columnData` 的 `w/hd/s`；尺寸是像素，状态使用 `BooleanNumber`，custom 等来源字段需保留。 |
| `packages/core/src/sheets/row-manager.ts`：`getRowHeight/isRowVisible` | ia 未指定或为 true 且 ah 是有效正尺寸时取 ah，否则取 h/默认行高；隐藏状态独立于保存的高度。 |
| `packages/core/src/sheets/column-manager.ts`：`getColumnWidth/isColumnVisible` | 列宽取 w/默认宽度；hd 明确隐藏，原宽度仍保留。 |
| `packages/core/src/sheets/sheet-dimensions.ts` | 有效尺寸、零值与默认尺寸校验的参考。 |
| `packages/sheets/src/facade/f-worksheet.ts`：行列尺寸/显隐/结构方法 | 公开尺寸 getter 与共享行列命令；结构修改连同关联模型进入内容历史。 |

本地严格拒绝非法字段；固定版部分结构输入会删除非法尺寸后继续。本地允许固定 0/1 和已有布尔状态，默认尺寸必须为正数。零尺寸仍可保存，并由现有可见性/几何路径跳过，与手动隐藏原因分开。默认行列尺寸保留本项目既有回退，与固定版缺省值的完整对照尚未验收。

## 实际接线与权威数据

- [sheet-public-axis.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-public-axis.ts) 的 `normalizeSheetPublicAxisData()` 将 h/w 映入 `rowHeights/columnWidths`，ah 映入 `rowAutoHeights`，ia 映入 `rowAutoHeightSpans`，hd 映入对应隐藏 spans，s 进入行列默认样式来源。
- ia=0 使用 h；ia=1 使用有效正 ah，没有时保存 h fallback；ia 缺省且存在有效正 ah 时启用自动，否则只有 h 的行沿本地手动尺寸语义。ia 缺省时后续自动重新计算与固定版全部组合仍需核验。
- `applySheetPublicAxisInput()` 先准备全部字段、默认尺寸和本地点/spans 覆盖，再一次应用。非法输入不会发布部分尺寸；helper 本身也在赋值前完成校验。显式本地尺寸/状态优先于公开输入；不为轴数据生成空白单元格。
- [model-utils.ts](../../packages/office-studio/src/api/model-utils.ts) 的公开工作表导入已接入 helper，修复原先仅保存 s/raw、遗漏 h/w/hd 和部分本地尺寸字段的路径。
- 原公开对象脱离调用方后保存在 `metadata.publicAxisData.rowData/columnData`，保留 custom、显式 0 与未知扩展。它是来源存档；后续真实尺寸、隐藏和样式编辑以领域字段为权威。
- [sheet-axis-structure-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-axis-structure-command.ts) 的结构 planner 调用 `transformSheetPublicAxisSources()`，插入/删除/移动/复制同步来源坐标，复制深克隆扩展。现有结构事务负责实际尺寸、模式、样式和资源变换。
- 本地 kind=workbook 快照重开使用已保存的当前领域尺寸，避免把旧 raw h/w 重放到已编辑模型。View adapter、Facade、Canvas 尺寸索引、导航和输入桥沿同一模型读取，内容撤销重做仍按文档隔离。

这是 Sheets 导入/领域适配，没有增加第二套轴引擎或 Office 解析入口。小型字段归一化当前在 TypeScript；大范围结构、自动度量和布局按[Worker/WASM 契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)续作。本批不改控制 DTO、transferable、WASM lease 或内存生命周期，也没有性能达标结论。

## 草稿与 UI 边界

| 文件 | 静态数量 | 计划覆盖 |
| --- | --- | --- |
| [sheet-public-axis.contract.test.mjs](../../packages/office-studio/tests/sheet-public-axis.contract.test.mjs) | 14（11 纯契约 + 3 Facade） | 手动/自动/fallback/默认、隐藏与零尺寸、来源隔离、span 合并、本地覆盖、非法输入原子拒绝、结构来源坐标、view/快照重开、命令历史和保护。 |
| [sheet-public-axis.browser.test.mjs](../../packages/office-studio/tests/sheet-public-axis.browser.test.mjs) | 3 | 实际 Canvas 输入边界/字体、导航跳过隐藏/零宽、尺寸历史/重开/明暗主题/输入提交。 |

合计 **17 个用例草稿，全部未运行**。复用[默认样式 fixture](../../packages/office-studio/tests/fixtures/sheet-default-style.html)；没有启动 Vite 或浏览器。Facade undo 草稿比较真实 sheets 内容，保存状态标记与轴内容分别处理。

本批没有新增可替换 UI。已有单元格、公式栏和 Ribbon/尺寸控件继续使用 `antdv-next`、Form 与统一 ConfigProvider；主题、焦点恢复、真实输入法和两种 Canvas 后端均待统一验证，不能以静态接线证明交互验收。

## 剩余范围

- 全部自动行高计算及 ia 缺省/显式 true/手动 fallback 的后续编辑组合、字体度量和复杂脚本仍未验收。
- 全套固定工作簿 snapshot DTO、轴来源当前值读回和目录编辑 API 仍需补齐；raw 存档保真不等于所有未知属性已生效。
- XLSX 原生单位与行列样式/主题的真实解析写出、Office 独立重开仍待开发或验证；内部 JSON 不替代文件往返。
- 全资源/复杂公式结构变换、权限组合、超大范围 Worker/Rust/WASM、取消、首屏/滚动/编辑/缩放、传输/峰值内存及可复现前后基准仍未验收。

持久内容、作者状态和来源存档需求登记在[后端需求](../BACKEND_REQUIREMENTS.md)；后端按用户顺序暂缓。全部 Sheets、Docs 及 Pro 范围保持。
