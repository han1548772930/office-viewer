# SHEET-04 默认样式、样式引用与稀疏空白格

记录日期：2026-11-25。固定参考根 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。状态：**前端已补代码，未验收**。按用户安排，本批只阅读源码、修改前端及编写用例草稿；测试、类型、构建、浏览器、dev、Rust/WASM、服务及性能检查均未执行。

本批接续[字体与字号](SHEET-04-font-family-size.md)、[富文本字体与 Canvas 编辑](SHEET-04-05-rich-text-font-editing.md)、[颜色](SHEET-04-color-reset.md)、[边框](SHEET-04-borders.md)和[格式刷](SHEET-04-format-painter.md)。已有字体组仍为 11 个常驻根声明，行内字体适用其中 9 项；新增常规 Ribbon、插件扩展、独立溢出及上下文根入口均为 **0**。默认样式五个本地命令、目录引用和下拉值不计作新 Ribbon 根或已完成功能数量。

## 固定来源与可观察契约

以下路径均相对固定参考根。

| 来源 | 本批核实的行为 |
| --- | --- |
| `packages/core/src/sheets/worksheet.ts`：`getDefaultCellStyle/Internal`、`getRowStyle`、`getColumnStyle`、`getCellStyle` | 工作表、行、列与单元格样式可存对象或目录 ID；raw 和 resolved 读取分开。 |
| 同文件：`getComposedCellStyle/ByCellData`、`_isRowStylePrecedeColumnStyle` | 按属性组合工作表默认、行列、themeStyle、直接样式；默认行优先，可按调用反转。 |
| `packages/core/src/sheets/styles.ts`、`types/interfaces/i-style-data.ts` | `cell.s` 目录/对象解析及短字段 DTO；字体开关、颜色、布局、数字格式和边框不能只按按钮名称推断。 |
| `packages/sheets/src/facade/f-range.ts`：`getCellStyleData/getFontFamily/getFontSize` | 默认 `row` 为行优先组合；`col` 为列优先组合；`cell` 只读直接样式。字体 setter 的 null 清除作者字段。 |
| `packages/sheets/src/commands/commands/set-style.command.ts`、`basics/cell-style.ts` | 稀疏样式 mutation、筛除行、富运行覆盖及内容撤销重做。 |
| `packages/sheets/src/controllers/auto-fill.controller.ts`：`_prepareFillData` | NO_FORMAT 使用序列计算而保留目标格式；ONLY_FORMAT 保留目标内容；格式、合并及范围资源与模式有关。 |

本地 workbook 默认层、稀疏行列 spans、按属性 reset mask 和五个默认样式命令属于本项目的协议/接口实现；没有将它们称为固定版相同 command ID 或五个新增上游入口。

## 权威模型与组合顺序

[domain/model.ts](../../packages/office-studio/src/domain/model.ts) 增加 `SheetCellStylePatch`、`SheetStyleSource` 和 `SheetCellStyleProperty`。`SpreadsheetCell.authoredStyle` 是稀疏直接格式，`authoredStyleReference` 保存公开 `cell.s` ID，`clearedStyleProperties` 仅屏蔽引用中明确重置的属性，`themeStyle` 为直接样式前的主题层。完整 `style` 是物化结果；持久 `styleMaterialized: true` 标记防止重开后将旧缓存再次捕获为作者格式。

工作簿保存 `defaultStyle/styles`；工作表保存 `defaultStyle`、行列点样式、行列 spans 和 `rowStylePriority`。当前顺序为：本地 Aptos/11 → workbook 默认 → worksheet 默认 → 列 → 行 → 范围主题 → cell theme → 扣除 reset mask 的引用直接样式 → 稀疏直接 patch。范围主题在[后续切片](SHEET-04-range-themes.md)加入，含 32 内置/九区段、稀疏规则和 46 个未运行草稿。行列优先级可反转；`wrap/wrapStrategy` 每层同步，直接 `wrap:false` 能覆盖继承换行策略。

旧快照没有 authored patch/reference 时，完整旧 `style` 继续视为直接样式，保留数据。新构造格在首次进入模型前的非默认 imperative 修改仍可采用；已经物化的快照不再这样推断。局部作者修改保留其他引用字段的实时解析；family/size/mark null 删除直接/run 字段并记录引用 reset，随后恢复继承来源。文字色/填充的 null 仍表示显式无色，保留既有颜色重置契约。

[sheet-style-inheritance.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-style-inheritance.ts) 集中解析、合并、清除、捕获和物化；[sheet-style-selection.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-style-selection.ts) 以来源边界/交点采样选区。稀疏空白格的渲染、编辑、字号步进和混合状态读取有效样式，不为读取而填满模型。

## API、输入、结构与历史

| 本地命令 | 作者内容 |
| --- | --- |
| `sheet.command.set-workbook-default-style` | workbook 默认来源 |
| `sheet.command.set-worksheet-default-style` | worksheet 默认来源 |
| `sheet.command.set-row-default-style` | 一个行的直接默认来源，null 清相应点/span |
| `sheet.command.set-column-default-style` | 一个列的直接默认来源，null 清相应点/span |
| `sheet.command.set-row-style-priority` | 工作表行列组合优先级 |

[sheet-default-style-commands.ts](../../packages/office-studio/src/api/sheet-default-style-commands.ts) 注册上述命令，先校验样式/ID/坐标和受影响的完整范围权限，再进入同一内容提交，关闭公式重算。Facade 提供 workbook/worksheet 默认读写、行列 raw/resolved 读取和设置；`FRange.getCellStyleData('row'|'col'|'cell')`、字体读取和 setter 复用共同来源。

公开输入接受 nested、固定短字段、字符串 ID 和 null；`rowData.s/columnData.s` 接入行列来源，原轴对象完整保存在 `metadata.publicAxisData`。未知字体字段、`n` 及颜色/旋转/边框扩展进入 `metadata.sourceStyle`；固定 `bd` 四边和六个斜线键转换为可绘制本地键，同时保留原数据。[公开样式读回续作](SHEET-04-public-style-readback.md)已输出 ff/fs 等短字段并保留 nested 兼容字段，另写 13 个未运行草稿。[公开轴续作](SHEET-03-public-axis-import.md)已映射 h/ah/ia/w/hd 到真实尺寸、自动模式和隐藏状态，并变换 raw 来源坐标；17 个草稿未运行，完整自动行高组合、snapshot 与 Office 往返仍未验收。

[excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) 携带 sparse patch、reference/reset/theme；单格和完整工作簿映射共用。空格输入按照继承数字格式解析，物化数字格式变更使用当前地区。轴插入/删除/移动/复制变换点/span 及 cell 来源；大量相同插入行使用紧凑 span。输入和结构计划用 `preserveStyleInheritance` 提交，不从移动前的物化缓存推断新增作者覆盖。

## 空白格复制、格式刷、填充与边框

[sheet-style-region.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-style-region.ts) 保存相对坐标的已解析默认、行列点/spans 和 priority，不保存 live IDs。连续范围裁剪 spans，筛除行按物理坐标映射。剪贴板与格式刷只存已有格和区域来源；按需投影虚拟空白。内部 HTML 包含空白格有效样式，cut fingerprint 也包含继承来源。

带格式粘贴/刷格式转成明确直接格式，清源 theme/ref/reset；values-only 与外部文本保留目标格式来源，exceptBorders 保留目标有效边框。bare blank 到 bare blank 的 all paste/复制填充保持空稀疏图。清除格式还删除已建模 rich-run 字体字段及公开别名，保留文字、ID、链接和未知扩展。

[sheet-auto-fill.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-auto-fill.ts) 用同一区域样式快照采样，包括空白源格和继承数字格式。COPY/ONLY_FORMAT 固化采样格式；NO_FORMAT 恢复目标有效样式及全部 authored/ref/reset/theme 来源，既有空内容可清，新空格保持稀疏。refill 沿既有一次内容历史修订；[sheet-fill-commands.ts](../../packages/office-studio/src/api/sheet-fill-commands.ts) 提交保留来源。

边框 plan 读取有效边框，只 author 变更的 border group。无继承边框时 none 保持既有稀疏扫描；存在继承时按需物化选区/相邻真正变化边，权限覆盖邻格。继承边框同步投影预算为 **100,000 个位置**；填充同步预算也是 **100,000 格源/目标总工作量**。这些是临时执行预算，Worker 大范围路径尚未实现，不能作为性能指标或完整大范围能力。

## 控件、计算域与验证状态

本批不新增 UI 控件；涉及的 Ribbon、颜色/边框弹层及 Canvas 单元格/公式输入仍使用 `antdv-next` 和统一 ConfigProvider。Canvas 内核负责绘制、命中、文字选区和拖拽。明暗主题、弹层、焦点恢复和真实中文 IME 均须浏览器验证；没有将静态库组件接线标为交互已验收。

样式来源/命令属 Sheets feature 和 API，物化视图通过 adapter 供 Canvas/Office DTO。未改变 Worker 控制 DTO、transferable、WASM lease/内存或网络边界。未来重样式投影、布局与填充按[平台契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)落实 Worker/Rust/WASM，不维护冲突的第二套引擎。

| 草稿 | 静态数量 | 计划覆盖 |
| --- | --- | --- |
| [sheet-default-style.contract.test.mjs](../../packages/office-studio/tests/sheet-default-style.contract.test.mjs) | 53 | 组合优先级、wrap、空格、null/legacy、raw/reference/reset、short DTO/边框/扩展、输入、轴 spans、区域采样、复制/格式刷/边框、保护、缓存重开、填充各模式/筛除、Facade/取消/隔离历史与快照。 |
| [sheet-default-style.browser.test.mjs](../../packages/office-studio/tests/sheet-default-style.browser.test.mjs) | 5 | Ant 空格输入、编辑中默认变化、字体/运行重置、百分比输入/快照、主题和 Ribbon 宿主。 |

静态数为 **58 个用例草稿，全部未运行**。[浏览器 fixture](../../packages/office-studio/tests/fixtures/sheet-default-style.html) 仅为后续样本。自动填充本批增加领域/Facade 草稿，Canvas 填充手势与默认来源的组合仍缺专门浏览器场景。

## 仍未完成或未验收

- 默认字体 Aptos/11 与固定 Arial/11 不同；完整主题字体/CJK/复杂脚本及未建模 rich-run 字体仍需继续。[范围主题](SHEET-04-range-themes.md)已补目录/九区段/API/Ant 应用移除、区域来源与轴/表复制；[局部移位/剪切续作](SHEET-04-range-theme-structure.md)保留逻辑片段与作者来源，32 个草稿未运行。九区段自定义 UI、移位共享命令/确认、全套 snapshot/Office 及浏览器/性能仍未完成或未验收。
- Facade 短字段读回和基础公开轴尺寸/自动模式/隐藏已续作，nested 字段是本地兼容扩展；全部未建模属性、每属性来源保真、全套 snapshot、完整自动行高组合及目录编辑 API 仍需补齐或验收。
- default/catalog/axis spans 与直接来源的真实 XLSX 解析/写出映射、主题/tint 和 Office 独立重开未开发或未验收，内部 JSON 不能代替原生文件保真。
- 大范围继承边框清除、填充和非重置样式 Worker 路径、取消、传输/内存与大样本测量仍缺；没有本批首屏、帧时间、长任务、复制字节或同数据集前后基准。
- 完整工作簿只读/设置样式权限、所有宿主与复杂 shaping/RTL/真实 IME 仍需组合验收。全部 Sheets / Docs / Pro 目标保留。

持久协议和原生文件需求同步到[后端需求](../BACKEND_REQUIREMENTS.md)，后端仍按用户顺序暂缓。
