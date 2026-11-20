# 在途功能证据记录（2026-11-20）

本记录依据当前共享工作区检查；并行任务仍在修改，内容是本次检查时的快照。它记录部分实现及实际验证边界，不构成完整功能验收。SHEET-22/23/25、COMMON-11、DOC-11/12 仍按主清单标为“未验收”。参考基线为 `D:\rust_test\univer-reference` 的 `6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68`。

## COMMON-11 Canvas WebGPU Rust WASM

- **规格依据**：本项目 `docs/RENDERING_AND_COMPUTE_STANDARD.md` 定义 WebGPU 优先、Canvas2D 回退、Worker/WASM 计算、取消/版本/内存边界；固定 Univer 源码只用于可观察渲染和交互契约，不表示上游采用相同后端。
- **模型与实现**：画布场景与 Renderer 分离；`packages/office-studio/src/rendering/createRenderSurface.ts` 管理后端生命周期，`canvas2d.ts` 和 `webgpu.ts` 提供绘制后端。Docs 布局已增加按页分桶的 `pageContent`，`document-scene.ts` 对可见页生成 Canvas 命令。计算层有 `compute/client.ts`、Worker 协议及 Rust `crates/office-compute`；新建 `features/sheet-compute/numeric-index-service.ts` 为数值排序/筛选打包并委派到 Worker Client，要求结果 revision 仍匹配。
- **行为与回归证据**：`node --experimental-strip-types --test packages/office-studio/tests/compute-index-service.contract.test.mjs packages/office-studio/tests/rendering/document-scene.contract.test.mjs packages/office-studio/tests/rendering/render-surface.contract.test.mjs`，检查时 **10/10 通过**。覆盖数值任务的 TypedArray 编码、筛选/排序委派、取消与过期结果、可见页场景以及 Canvas DPR / WebGPU 丢失回退；计算服务测试使用 fake client，不证明已加载真实 WASM。
- **闭环、界面与性能**：未完成真实大表 UI 排序/筛选命令调用、真实 Worker+编译 WASM 的浏览器运行、全文档绘制/编辑闭环和设备实测。本次无前后性能数据集、P50/P95、长任务或内存报告；不得声称已达到性能目标。WebGPU 运行设备、真实大文档滚动及浏览器无 GPU 情况仍待验收。
- **当前结论**：存在多个真实分层实现与契约测试；`NumericIndexService` 尚无已确认的业务调用点，不能称排序/筛选已在产品路径采用 Rust/WASM。

## SHEET-22 Editable Charts

- **规格依据**：`docs/UNIVER_FEATURE_PARITY.md` 中的 `[图表规格][sheet-charts]` 及固定提交参考中的 `packages/sheets-drawing-ui/src/services/sheet-chart-render-object.config.ts`。具体完整图型目录、轴/标签/主题和 Office 往返仍需逐项展开核实。
- **模型与实现**：`domain/model.ts` 有 `ChartModel`；`features/charts/chart-engine.ts` 读取源范围、生成 bar/line 布局及 Canvas DrawCommands。`FWorksheet`/`FChart` 与 `sheet.command.add-chart`、`update-chart`、`remove-chart` 接入模型历史；`ChartOverlay.vue` 和 `ExcelWorkbench.vue` 新增插入对话框、图层绘制以及选择/移动/缩放操作。
- **行为与回归证据**：`chart-facade.contract.test.mjs` 的 3 个契约用例已检查绑定源数据、数据变化后重读、命令撤销/重做、失败更新原子性和 log 轴跳过非正数。测试只覆盖 bar/line 引擎/API，没有覆盖工作台浏览器交互。
- **闭环、界面与性能**：尚无本记录可核验的图表 UI 浏览器测试、XLSX 导出后独立重开结果或 Office 程序验证。`office-writer/src/xlsx.rs` 当前正在改动；完成后需编译和重开核验。图表编辑仍未覆盖完整类型、轴标签/主题编辑面板、打印和透视图联动。没有相同数据集的前后性能基准。
- **待核实**：单元格源范围变更与图层刷新、跨 Sheet/多区域引用、图表数据缓存、键盘可访问命中、所有公开图型、坐标在冻结/缩放/滚动下是否一致，以及 XLSX 的 series/axis/style 的独立读取保真。

## SHEET-23 Pivot Tables

- **规格依据**：`[透视表规格][sheet-pivot]`、固定提交的 Sheets/Pivot 实现及测试。规格目标仍包括 cache、字段、公开聚合、排序筛选、小计/总计、分组、下钻和刷新。
- **模型与实现**：`PivotTableModel` 持久化 source、anchor、fields、options、revision 和下钻映射；`features/pivot/services/pivot-engine.ts` 实现字段设置、聚合、过滤/排序、数字/日期/手动分组、刷新及下钻；`OfficeStudio` 将相关操作注册为命令并同步生成输出区域。
- **行为与回归证据**：`pivot-facade.contract.test.mjs` 3 例覆盖基础字段/总计、筛选排序/设置/刷新/下钻、冲突；`pivot-grouping.contract.test.mjs` 4 例覆盖数字/日期/手动分组与错误原子性；`pivot-atomicity.contract.test.mjs` 2 例覆盖锚点冲突和失败操作不污染模型/历史。当前联合运行命令为 `node --experimental-strip-types --test packages/office-studio/tests/chart-facade.contract.test.mjs packages/office-studio/tests/pivot-facade.contract.test.mjs packages/office-studio/tests/pivot-grouping.contract.test.mjs packages/office-studio/tests/pivot-atomicity.contract.test.mjs`，检查时 12/12 通过。
- **闭环、界面与性能**：未完成透视编辑面板浏览器验收、XLSX PivotCache/Parts 写出与独立重开、完整 Excel 日期系统/日期派生类型和更广输入文件样本测试。当前分组部分以字符串 ISO 日期和数值实现，1904 日期系统、季度/年内边界和 Excel 日期序号兼容仍待核实。没有大型 cache 的 WASM 迁移或性能实测。
- **待核实**：所有聚合类型与空值/错误语义、复杂多层小计、多个 value fields、结构编辑和引用变换、刷新时锚点冲突、切片器和透视图联动、文件写出/重开。

## SHEET-25 Slicers and Timelines

- **规格依据**：`[切片器规格][sheet-slicer]` 及固定提交文档/API；目标包含表格和透视目标、交叉筛选、样式/布局与持久化。
- **模型与实现**：`SlicerModel` 记录 workbook、host sheet、pivot、source field、控制类型、选择/时间范围、布局与 revision。`features/slicer/slicer-engine.ts` 支持透视标签字段上的多选、清除、按 day/month/quarter/year 选取范围并写回透视字段过滤；OfficeStudio 已有 add/set/clear/remove 命令和 `FSheetSlicer` facade。
- **行为与回归证据**：检查时 `packages/office-studio/tests` 没有 slicer 契约测试；无界面或浏览器测试。故当前只有源码路径证据，关键输入/边界行为尚未由测试证明。
- **闭环、界面与性能**：没有切片器 UI、命令闭环测试、表格对象目标绑定、XLSX 导出重开、服务端协作/权限或性能实测。日期解析、不同日期系统、时区及重复值范围语义尚待验证。
- **待核实**：切片器是否能以正确目标绑定及保留格式/尺寸；多个透视目标共享一个切片器；同一透视字段交叉筛选；时间线粒度 UI、范围拖动和空范围；删除/撤销和工作簿重新载入；标准 XLSX 样本兼容。

## DOC-11 Content Controls

- **规格依据**：`[内容控件规格][doc-controls]`；完整公开类型目标包含文本、复选、日期、列表、图片、分组和重复节，及占位/属性/锁。
- **模型与实现**：`DocumentContentControl` 存储 blockId、UTF-16 段落范围、类型、占位、锁、值和列表项。`features/document/services/content-controls.ts` 实现创建/读取/更新/删除、文本/复选/列表值、范围替换和嵌套/部分重叠校验；`OfficeStudio` 注册命令，`FDocument` 暴露 API，并进入文档历史。
- **行为与回归证据**：本次检查时已有实现文件和命令，但 `packages/office-studio/tests` 中没有 content-control 专用契约测试；按钮、画布呈现和用户输入桥也未接入/验证。不能据 API 存在声称 UI 功能完成。
- **闭环、界面与性能**：未验证所有公开 control 类型、锁状态对普通文本编辑的拦截、修改前后偏移变换、撤销/重做、DOCX SDT 写出与独立重开；暂无视觉或性能数据。
- **待核实**：图片/分组/重复节节点模型，RichText 交互，内容锁与控件锁权限区别，控制范围跨段落行为、粘贴/结构编辑的位置映射、DOCX 结构映射和解析重开。

## DOC-12 Footnotes and Endnotes

- **规格依据**：`[脚注尾注规格][doc-footnotes]` 和固定提交 `packages/engine-render/src/components/docs/layout/footnote-layout.ts` 与对应测试。该规格说明需区分分页模式支持范围；目前本项目还没有此布局能力。
- **模型与实现**：`crates/office-writer/src/docx.rs` 已识别段落 run 上的 footnote/endnote 引用和顶层、metadata 或 notes collection 中的注释正文；构造 DOCX notes XML、relationships、content types，并为缺少正文的引用给出警告。当前 parser/DocumentModel 没有对应的规范化脚注/尾注模型或导入适配，DOCX writer 的 metadata 输入约定不等同于编辑器 API。
- **行为与回归证据**：`docx.rs` 有 Rust 单元测试检查 DOCX ZIP 中引用标记、脚注/尾注 parts 和 relationships，以及缺失正文时仍写入有效空 note 并报告 warning。检查时 `cargo test` 未通过编译，因为并行修改的 `xlsx.rs` 存在 format 参数数目与 `usize/u32` 类型错误；所以 DOCX 测试本轮未执行成功。
- **闭环、界面与性能**：没有文档模型命令、引用编号/重排、导航/跳转、分页注释区布局、撤销闭环、DOCX 导入解析或独立 Office 重开验证。没有性能基准。
- **待核实**：DOCX 新建与重开兼容、note body 的富文本/表格语义、删除/重排引用 ID 稳定性、同一事务的正文/注释内容历史、分页模式边界、DOCX importer 对 notes parts 的保留。

## 本记录使用的核验结果

- `node --experimental-strip-types --test packages/office-studio/tests/chart-facade.contract.test.mjs packages/office-studio/tests/pivot-facade.contract.test.mjs packages/office-studio/tests/pivot-grouping.contract.test.mjs packages/office-studio/tests/pivot-atomicity.contract.test.mjs`：12/12 通过。
- `node --experimental-strip-types --test packages/office-studio/tests/compute-index-service.contract.test.mjs packages/office-studio/tests/rendering/document-scene.contract.test.mjs packages/office-studio/tests/rendering/render-surface.contract.test.mjs`：10/10 通过。
- `cargo test`（`crates/office-compute`）：通过，但该 crate 当前为 0 Rust tests；TypeScript compute contract 才覆盖 adapter 分派行为。
- `cargo test`（`crates/office-writer`）：失败。当前并行 XLSX 代码编译错误阻断 suite；writer 回归项未验收。
- browser 实际界面闭环、可访问 Univer 演示截图对照、真实 WASM worker、Office reader/Office 程序重开和性能基准：本记录没有此类结果，全部保持待验收。
