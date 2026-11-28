# SHEET-11 / SHEET-18 / SHEET-29 颜色筛选的 Office 边界

状态：已落盘，未验收。固定 Univer v1.6.0：`eaf4cfa04fabdbcfc922c692f3450459f8a29546`。前端入口与固定模型依据见[筛选表头及颜色条件记录](SHEET-11-filter-header-color-2026-11-28.md)；本记录说明共享模型、派生数据和原生 Office 读写。

## 数据和读写

[领域模型](../../packages/office-studio/src/domain/model.ts)增加 `SheetColorFilter`：`cellFillColors?: (string | null)[]` 和 `cellTextColors?: string[]`。填充字段存在时优先，null 表示无填充；空数组保留“无颜色匹配”的语义。前端[边界归一化](../../packages/office-studio/src/parser/workbookAutoFilter.ts)检查有界数组、文本长度并去重；Rust 类型和快照按相同字段保存。

Microsoft 的[标准颜色筛选规格](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colorfilter)规定 `colorFilter` 通过 `dxfId` 指向 styles part 的差异样式，`cellColor` 区分填充/文字。核对日期 2026-11-28。标准单个条件表达一个颜色，项目中的多色数组没有等价的标准原生多选表达。

[writer 颜色适配](../../crates/office-writer/src/xlsx_filter_color.rs)实际向现有 StyleRegistry 注册 dxf，生成填充、无填充或字体的 `colorFilter`，不以值筛选冒充颜色筛选。空数组使用已有互斥条件表达无匹配。多色选择写首个原生条件，同时保留完整项目快照与当前隐藏行，并返回 `AUTOFILTER_MULTIPLE_COLORS` 诊断：Office 重新应用单个原生颜色时可能不同。该差异不记为原生多色互操作已完成。

工作表沿已有 axis/filter extension 保存完整 typed 状态；[结构化表 writer](../../crates/office-writer/src/xlsx_table.rs)在 table extLst 中增加独立的 `urn:office-studio:table-filter:1` 快照，保留原有未知扩展、列身份、按钮状态和排序来源。表格的 `filteredRows` 也参加 worksheet 隐藏行输出，与手动隐藏原因分别存储。

[共享颜色解码](../../crates/office-io-common/src/filter_color.rs)读取实际 dxf 中的填充/字体、主题+tint、自定义 indexed palette；复用已有颜色实现。[解析适配](../../crates/office-parser/src/sheet_filter_color.rs)按引用的索引流式提取 dxf，复用现有 OPC 主题/调色板读取。工作表及表格通过相同解码器恢复可编辑条件。

[过滤指纹](../../crates/office-io-common/src/filter_projection.rs)加入解析后的实际颜色依赖。Office 只改 dxf/theme/palette、保留原 dxfId 时也会使旧快照失效，转读当前原生规则。缺失资源、不可解释图案/渐变或无效来源不会伪造颜色结果；完整 native XML 和当时隐藏行保留，并输出诊断。更深层的图案/渐变匹配、富文本各运行颜色、未知样式资源的再次导出仍须继续实现和验收。

## 合并单元格与跨模块派生

固定 `Worksheet.iterateByColumn(false,false)` 返回完整 merge rowSpan，`FilterColumn.calc` 对整个组处理。因此 AutoFilter 的 ref 末行落在合并区内部时，`cachedFilteredOut` 可合法包含 ref 末行之后的行；入口不自动扩展 ref。源码路径和整组前端实现见上述筛选记录。

TypeScript 与 Rust 的 typed cache 保留这些行，仍拒绝标题行、重复及 Office 绝对越界；writer 另检查显式工作表 rowCount。criteria 列和范围本身的校验未放宽。标准 XLSX hidden 位不能独立区分手动/过滤原因，因此外部文件初始缓存仍保守读取 ref 内隐藏行，前端依照真实合并模型重新计算；项目精确快照保存完整整组结果。

共享 Studio、Facade、图表可见性/放置、切片器成员/原生筛选和 view/normalizer 已将整本工作簿传给颜色派生。颜色读取复用有效样式与既有条件格式/公式求值，保留跨表、日期系统、名称和继承；不会将条件格式结果写为单元格底色。切片器对空白单元格也传真实行列，避免错误读取 A1 的条件颜色。

## 草稿与检查边界

- [共享颜色/签名单元草稿](../../crates/office-io-common/src/filter_color.rs)：填充/无填充/字体、主题/调色板、默认布尔标志、缺资源及实际颜色变化。
- [解析写出往返草稿](../../crates/office-parser/src/sheet_filter_roundtrip_tests.rs)：工作表五种颜色状态、外部 dxf 编辑、缺失资源、结构化表多色快照/隐藏原因和跨 ref 合并缓存。
- 前端模型、API、面板与浏览器草稿见[筛选证据](SHEET-11-filter-header-color-2026-11-28.md)。

全部草稿未运行；没有执行类型、前端构建、Rust/WASM 检查、Office 文件重开、浏览器或性能测量。此处仅记录实现和只读语法/引用检查，不宣称已通过产品验收或优化达标。常规 Ribbon 根增量 0，全部 65 个顶层功能组继续未验收。
