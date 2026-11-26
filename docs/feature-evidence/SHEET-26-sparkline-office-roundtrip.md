# SHEET-26 迷你图 Office 读写与保存重开

状态：实现续作，所有合同草稿、浏览器、Office 独立重开与性能均未验收。本批按“先实现，后统一测试”约定，只静态读取、编辑及编写草稿，没有运行测试、类型检查、构建、Rust/WASM、浏览器或基准。

## 来源与契约

- Univer v1.6.0 OSS 基线为 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；Pro 公开包配置：`.research.local/univer-1.6.0/packages/sheets-sparkline/lib/types/common/type.d.ts`。类型是 Line / Column / Winloss / Pie；独立极值、六类标记、轴、隐藏源、空值/非数字处理、线宽、主题、渐变及圆角均在本项目范围内。菜单和完整配置追踪另见 [设置与分组证据](SHEET-26-sparkline-settings-and-groups.md)。
- 原生格式依据：[Microsoft Open XML SDK v3.3.0 x14 schema](https://github.com/dotnet/Open-XML-SDK/blob/v3.3.0/data/schemas/schemas_microsoft_com_office_spreadsheetml_2009_9_main.json)，2026-11-26 静态读取了 `SparklineGroup`、`Sparkline`、`SparklineTypeValues`、`SparklineAxisMinMaxValues` 和 `DisplayBlanksAsValues`。这是公开格式依据，未安装 SDK，也不是 Microsoft Excel 打开样本的证据。
- 原生仅有 `line`、`column`、`stacked`，分别对应本项目 line / column / winloss；原生 `span` 对应 connect。`minAxisType`、`maxAxisType` 各有 individual / group / custom，不能合并成一个旧 axis 开关。
- [Microsoft SparklineGroup 格式说明](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.sparklinegroup?view=openxml-3.0.1) 的公开规范文本，2026-11-26 静态读取：`lineWeight` 以 points 计，范围 0..1584，缺省 0.75；`displayEmptyCellsAs` 缺省 **zero**。本项目 UI/Canvas 的 lineWidth 为 CSS px，按共用 4/3 单位转换；纯模型允许 0..2112，常用 UI 仍提供 1..4 并保留导入宽度。原生缺省不套用本项目新建配置的 gap。x14 颜色不得含 auto 属性；无法按有效格式保真的输入显式失败。
- `colorSeries / colorNegative / colorAxis / colorMarkers / colorFirst / colorLast / colorHigh / colorLow` 必须按 schema 顺序写入；颜色值与 markers / first / last / high / low / negative 的布尔标志独立。仅保留颜色不能证明标记行为已保真。

## 实现边界

读写共用 `office-io-common::sparkline_io` 的可序列化配置、引用校验、原生投影与来源校验。parser 的 Sidecar 只解析 OPC/工作表片段；writer 只将统一单元格模型投影为 x14 与本地扩展。领域模型的 `officeSource` 保存不含成员的原生分组模板、颜色基线及必要的主题/索引调色板；可编辑内容仍以单元格 SparklineSpec 为准。

本地扩展命名空间为 `urn:office-studio:sparklines:1`。它保存 Pie、主题名、渐变、圆角、非数字策略、inactive 极值、groupId 与完整可编辑配置；不会以跳过 Pie 冒充成功。原生投影签名只用于意外陈旧检测，不是安全签名。外部编辑/删除原生迷你图后，陈旧本地记录不得覆盖新原生内容或复活已删除成员。

## 未运行合同草稿

本轮静态补齐了两个已由 `cfg(test)` 模块引用的文件。手工逐函数登记，Rust 共用层 12 条、writer 8 条、parser 12 条、Vue 11 条，共 43 条测试声明，全部未运行，不将声明数当作能力验收。

- `crates/office-io-common/src/sparkline_io_tests.rs`：引用与默认值、独立标记/极值、主题及索引/alpha、来源预算与字段、投影签名、导入线宽与 Pie。
- `crates/office-writer/src/xlsx_sparkline_tests.rs`：分组及颜色顺序、精度/引号、原生与本地配置、来源冲突拆组、主题/调色板冲突、0/非整数/大线宽、晚失败不返回部分 XML。
- `crates/office-parser/src/sparkline_sidecar_tests.rs`：由 ZIP helper 构造 relocated OPC 与 Strict 片段；本项目真实 writer → parser 的合同草稿；外部原生修改/删除/移动、主题改色/无关改色刷新、重复来源及日期轴元数据保留。
- `packages/office-studio/tests/sparkline-office-source.contract.test.mjs`：UTF8 预算、DTO → normalizer、逐成员模板、整组/合并/拆组/改源与移动、插件资源保存重开、非法晚成员原子拒绝。

合成片段位于 `crates/office-parser/tests/fixtures/sparklines/`。theme/palette 是字段合同片段，不代表完整主题 schema 或 Office 打开已通过。Strict 草稿仅核验命名空间与 Sidecar，完整 Strict 的第三方 cell parser 支持仍未验收。旧 writer XML 精确断言、旧原生 parser 的源 qualifier 断言需要主任务同步迁移，未运行本批检查。

## 未核实与后续验收

- 合成 OOXML 与本项目 writer/parser 合同只能覆盖本项目链路；Microsoft Excel / LibreOffice 独立打开、编辑、保存以及第三方工具是否保留未知扩展均未执行。
- 日期轴及原生未知扩展保留在 Office source 中；本轮没有声明日期轴 Canvas 语义已完成。
- 对外部工作簿、三维引用和无法解析的命名/公式源，不伪造 A1 数据源；需继续补模型与引用求值支持。
- 主题、焦点、IME、Ant 控件以及 WebGPU / Canvas2D 下的真实保存重开仍留待统一验收；本轮无新增 UI 控件。
- 大分组、主题复制、Worker/WASM 传输和内存仍未实测，不声称已优化或全链路零拷贝。
