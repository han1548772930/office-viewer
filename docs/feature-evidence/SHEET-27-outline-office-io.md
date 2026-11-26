# SHEET-27：大纲 XLSX 原生投影与保存重开

日期：2026-11-26。状态：本批代码和合同草稿已落盘，**未验收**。沿用用户“先对齐、测试统一后置”的安排；没有执行测试、类型检查、构建、Cargo/WASM、浏览器、服务、基准或 `git diff --check`。

## 对照与归属

固定 Univer v1.6.0：`E:\rust_test\univer`，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。本批只读确认 HEAD 和公开类型，没有安装或执行商业 SDK。

- [固定公开 group/resource 类型](../../.research.local/univer-1.6.0/packages/sheets-outline/lib/types/common/type.d.ts)：`id/axis/start/end/collapsed`，0-based inclusive，public owner 为 `unitId/subUnitId`；`outlineList` 为原插件资源契约。
- [固定公开算法类型](../../.research.local/univer-1.6.0/packages/sheets-outline/lib/types/common/algorithm.d.ts) 与 [上一批证据](SHEET-27-outline.md)：范围、相等层、最多八组、相邻身份、结构操作和六个 Facade API；汇总方向为本地 Office 增强。
- 原生来源为 worksheet `sheetPr/outlinePr`、`sheetFormatPr`、`cols/col`、`sheetData/row`。`outlineLevel` 是细节层级，`collapsed` 是相邻汇总维度的标记，`hidden` 是多种原因合成位；不能把标记直接解释为所有相邻层都折叠。
- 可后续复查的公开参考：[Row.OutlineLevel](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row.outlinelevel)、[Row.Collapsed](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row.collapsed)、[OutlineProperties.SummaryBelow](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.outlineproperties.summarybelow)、[OutlineProperties.SummaryRight](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.outlineproperties.summaryright)。这些链接本批未联网复核；Excel 对七个原生分组层级/八级显示的真实行为仍需 Office 样本确认，不能据此宣称完成 Office 验收。

独占模块属于 Office IO 归一化/写出层和 outline feature 的 DTO 适配层。UI 不参与解析或 XML 写出；现有 Worker/WASM 入口、任务控制和内存协议复用项目平台层。Root 维护 shared parser/model/writer/normalizer、API 和 Canvas，本子任务没有修改这些共享入口，也没有重做 Sparkline/Slicer。

## 本批文件和实际行为

本批新增 **10 个文件**，其中核心模块 4、专属合同草稿 2、浏览器草稿/夹具/配置 3、本证据 1。Rust common/writer 模块内另有单元合同草稿。

| 文件 | 已落盘行为 |
| --- | --- |
| [sheet_outline_source.rs](../../crates/office-io-common/src/sheet_outline_source.rs) | 共享 group/state/manual/source DTO；未知 state/group/source JSON 保留；Streaming namespace/XML 校验；读取 row/col 层级、标记、hidden 和汇总方向；压缩 source runs；独立模型/可见性指纹；完整项目扩展校验；未知 outlinePr 子树、混合文本、属性和 QName namespace bindings 保存 |
| [sheet_outline_sidecar.rs](../../crates/office-parser/src/sheet_outline_sidecar.rs) | 根据实际 OPC worksheet parts 读取，在 dimensions/filter restore 后恢复；独立 guards 保留模型与手动重叠；外部 outline 修改采用当前原生分组；空细节和汇总维度扩充必要 extent；保存 native/authored/projectSnapshot/诊断；XML、预算和无法构建完整模型通过 `failure` 交给主 parser |
| [xlsx_outline.rs](../../crates/office-writer/src/xlsx_outline.rs) | `OutlineInput` 保留 ID/summary/unknown；校验八层、不交叉范围；事件区间生成原生层级、折叠 hidden 和 summary marker；单一 sheetPr 合并 filterMode；sheetFormatPr 最高层；拆分 cols 几何且保留尺寸/手动隐藏；有序 row iterator；写出完整 guarded extension；未知 source 属性/namespace 与 JSON 来源保留 |
| [sheet-outline-import.ts](../../packages/office-studio/src/features/sheet-outline/sheet-outline-import.ts) | 严格 DTO -> domain 候选；保留 unknown/ID，剥离旧 public owner；必要 extent 不截断；独立 metadata.outlineOfficeSource；原生不精确原因与项目精确手动重叠区分；Rust 已归一化 flag 防止二次剔除；筛选推断同时处理数组和压缩 spans |
| [sheet_outline_roundtrip_tests.rs](../../crates/office-parser/src/sheet_outline_roundtrip_tests.rs) | 真实 ZIP/OOXML parser↔writer 合同草稿，含 source 再写出、编辑 staleness、未知 XML、八层/边界、filter 共存、损坏失败、实际 OPC part 移动和独立 Strict sidecar |
| [sheet-outline-office-import.contract.test.mjs](../../packages/office-studio/tests/sheet-outline-office-import.contract.test.mjs) | ID/unknown 深拷贝、边界 extent、八层、无损候选拒绝、手动 overlap、过滤 spans、Rust flag 与 metadata 合同草稿 |
| [sheet-outline-office-io.html](../../packages/office-studio/tests/fixtures/sheet-outline-office-io.html) | 真实 Facade → export Worker/WASM → XLSX → parse Worker/WASM/normalizer → Facade 夹具；无产品控件、无伪 writer/parser；再次展开、Undo/Redo 和第二次文件保存 |
| [sheet-outline-office-io.spec.ts](../../tests/browser/sheet-outline-office-io.spec.ts) | 实际二进制/Facade 往返、独立 Node ZIP/DEFLATE 读取、手动重叠、插件身份、未知 state/group 和八层边界诊断草稿 |
| [sheet-outline-office-io.playwright.config.ts](../../tests/browser/sheet-outline-office-io.playwright.config.ts) | 后置验证配置；须先重建本项目 parser/writer WASM；本批未启动其中的 Vite 或浏览器 |

静态回读修正了独立 `cols` 片段合入 main namespace 时可能生成 `col xmlns=""` 的问题。原生列节点显式绑定 worksheet main namespace；补入合同草稿，尚未运行证明。

## 跨模块协议

Root 已协调接入以下共享接口；该接线的实际运行仍未验证：

```rust
// common/lib.rs
pub mod sheet_outline_source;
// parser payload, camelCase -> outline/sourceOutline
outline: Option<office_io_common::sheet_outline_source::SheetOutlineState>
source_outline: serde_json::Value
// parser sidecar
SheetOutlineSidecar::read(bytes: &[u8]) -> Self
restore(&mut self, sheet: &mut SheetPayload)
warnings: Vec<String>
failure: Option<String>
// writer
outline::prepare(sheet: &SheetInput) -> Result<WorksheetOutlinePlan, ExportError>
plan.sheet_properties_xml(filter_mode: bool) -> Result<String, ExportError>
plan.format_xml(&axis.format_xml()) -> Result<String, ExportError>
plan.columns_xml(&axis.columns_xml()) -> Result<String, ExportError>
plan.row_indices() -> impl Iterator<Item = u32>
plan.row_hidden(row: u32) -> bool
plan.row_attributes(row: u32) -> String
plan.extension_xml(sheet, worksheet_prefix) -> Result<String, ExportError>
plan.warnings: Vec<String>
```

XML 插入顺序：worksheet 开头的原 sheetPr 用合并结果替换；sheetViews 后合并 sheetFormatPr/cols；在原有 cell/geometry/manual-hidden 三路 row iterator 加入 outline 第四路；row hidden OR 折叠原因，并添加 outline attrs；worksheet extLst additions 合并 outline extension，保持单一 extLst。

```ts
const outlineImport = applySheetOutlineImport(normalized, raw)
reconcileSheetOutlineImportedVisibility(normalized, outlineImport, filteredRows)
```

项目扩展 URI：`{94802E78-595D-40AD-BA94-29DDF46ED673}`；namespace：`urn:office-studio:sheet-outline:1`。JSON `version/state/manualVisibility/source` 保存完整模型，`outlineFingerprint` 和 `visibilityFingerprint` 分别校验原生分组投影及合成 hidden 位。

`sourceOutline` 保存本次 `native`、扩展中的 `authored`、完整 `projectSnapshot`、`manualVisibilityNormalized`、`projectOutlineExtension`。`restored/projectionMatched` 表示模型 guard；`visibilityProjectionMatched/hiddenReasonsExact` 表示隐藏原因 guard。未知筛选语义导致 manual 集合合入标准隐藏时不能沿用 axis restored 标记宣称精确。

## 原生限制与失败策略

- 项目八层模型全部保存。输出原生 `outlineLevel` 最多七；第八层产生 `OUTLINE_DEPTH_SNAPSHOT`，保留八组及全部 ID。读到原生 level 八时保留完整八组和 `NATIVE_LEVEL_OUTSIDE_OFFICE`；超过模型八层失败，避免只恢复七/八组后返回成功。
- 缺少相邻 summary dimension 的边界组不删除、不扩大已声明的全轴 extent。折叠存在完整扩展并记录 `OUTLINE_BOUNDARY_SNAPSHOT`；原生相邻 marker 不可用。没有扩展的原生边界 collapse 与手动隐藏不能无损反推。
- 同一原生 summary marker 不能表达每个同边界组的独立折叠状态；完整扩展保留不同层状态，必要时记录 `OUTLINE_SHARED_MARKER_SNAPSHOT`。原生推断仅恢复 marker 所指相邻层，且细节必须实际隐藏；不凭一个 marker 折叠全部层。
- `hidden` 只有合成位。原生文件没有独立项目原因时，用已观察到的 marker/细节和已知 filter 推断，保留完整 native runs 与 `NATIVE_HIDDEN_REASONS_AMBIGUOUS`；手动/outline/filter/zero-size 重叠不宣称精确保真。
- 已识别的本项目扩展损坏、嵌套 JSON、超预算、重复、空/缺失、异 namespace、不支持版本或 source/manual 校验失败直接失败。主 parser 必须在 read 后和 restore 后传播 `failure`，不能警告后返回缺组的成功模型。
- 合法外部 outline 编辑导致模型 fingerprint stale 时采用当前原生状态，同时保留完整旧 projectSnapshot/source；普通内容、样式和几何修改不使大纲模型 guard 失效。hidden 变化仅影响对应可见性恢复；指纹不提供防篡改认证。

## 检查与未验收项

本批仅做源码/工作树/参考类型和引用路径静态读取、编辑及草稿声明计数。新增测试声明草稿 **51 条**：common 16、writer 7、parser ZIP 11、TS import 14、真实 Worker/WASM/Facade 浏览器 3；**全部未运行**。

没有新增产品 UI 控件；夹具只暴露编程接口。新增模块不涉及 Ant 控件迁移；上一批 Ant 弹窗与 Root Canvas 输入的主题、焦点、IME 验证仍未执行，不能据此新增“UI 已验收”结论。

Streaming source/row iterator 的布局不等于实测优化。worksheet XML 上限 512 MiB，source/extension 各 32 MiB，source run 各轴最多 262,144，未知 XML 子树 16 MiB/128 depth/262,144 nodes。源 DTO、JSON 和 ZIP 写出仍有真实分配/复制；没有转移 WASM linear-memory buffer，也没有把 serde/transferable 宣称端到端零拷贝。

剩余验证包括：所有新增类型/Rust/WASM 编译；真实 Windows Edge + Canvas 两后端与 Office Excel 样本；Strict 全文件 parser；大量空 detail 行、密集 hidden runs、内存峰值、任务取消/超时与异步让出；外部 Office 去除项目扩展后的第八层/边界/重叠差异；原生 unknown source 在结构插删/移动后的可观察效果。没有测量首屏、滚动、编辑、缩放、Worker/WASM 传输或内存，性能继续“未验收”。
