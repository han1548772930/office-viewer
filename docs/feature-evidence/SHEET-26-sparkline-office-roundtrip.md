# SHEET-26 迷你图 Office 读写与保存重开

状态：实现续作，所有合同草稿、浏览器、Office 独立重开与性能均未验收。本批按“先实现，后统一测试”约定，只静态读取、编辑及编写草稿，没有运行测试、类型检查、构建、Rust/WASM、浏览器或基准。

## 来源与契约

- Univer v1.6.0 OSS 基线为 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；Pro 公开包配置：`.research.local/univer-1.6.0/packages/sheets-sparkline/lib/types/common/type.d.ts`。类型是 Line / Column / Winloss / Pie；独立极值、六类标记、轴、隐藏源、空值/非数字处理、线宽、主题、渐变及圆角均在本项目范围内。菜单和完整配置追踪另见 [设置与分组证据](SHEET-26-sparkline-settings-and-groups.md)。
- 原生格式依据：[Microsoft Open XML SDK v3.3.0 x14 schema](https://github.com/dotnet/Open-XML-SDK/blob/v3.3.0/data/schemas/schemas_microsoft_com_office_spreadsheetml_2009_9_main.json)，2026-11-26 静态读取了 `SparklineGroup`、`Sparkline`、`SparklineTypeValues`、`SparklineAxisMinMaxValues` 和 `DisplayBlanksAsValues`。这是公开格式依据，未安装 SDK，也不是 Microsoft Excel 打开样本的证据。
- 原生仅有 `line`、`column`、`stacked`，分别对应本项目 line / column / winloss；原生 `span` 对应 connect。`minAxisType`、`maxAxisType` 各有 individual / group / custom，不能合并成一个旧 axis 开关。
- `colorSeries / colorNegative / colorAxis / colorMarkers / colorFirst / colorLast / colorHigh / colorLow` 必须按 schema 顺序写入；颜色值与 markers / first / last / high / low / negative 的布尔标志独立。仅保留颜色不能证明标记行为已保真。

## 实现边界

读写共用 `office-io-common::sparkline_io` 的可序列化配置、引用校验、原生投影与来源校验。parser 的 Sidecar 只解析 OPC/工作表片段；writer 只将统一单元格模型投影为 x14 与本地扩展。领域模型的 `officeSource` 保存不含成员的原生分组模板、颜色基线及必要的主题/索引调色板；可编辑内容仍以单元格 SparklineSpec 为准。

本地扩展命名空间为 `urn:office-studio:sparklines:1`。它保存 Pie、主题名、渐变、圆角、非数字策略、inactive 极值、groupId 与完整可编辑配置；不会以跳过 Pie 冒充成功。原生投影签名只用于意外陈旧检测，不是安全签名。外部编辑/删除原生迷你图后，陈旧本地记录不得覆盖新原生内容或复活已删除成员。

## 未核实与后续验收

- 合成 OOXML 与本项目 writer/parser 合同只能覆盖本项目链路；Microsoft Excel / LibreOffice 独立打开、编辑、保存以及第三方工具是否保留未知扩展均未执行。
- 日期轴及原生未知扩展保留在 Office source 中；本轮没有声明日期轴 Canvas 语义已完成。
- 对外部工作簿、三维引用和无法解析的命名/公式源，不伪造 A1 数据源；需继续补模型与引用求值支持。
- 主题、焦点、IME、Ant 控件以及 WebGPU / Canvas2D 下的真实保存重开仍留待统一验收；本轮无新增 UI 控件。
- 大分组、主题复制、Worker/WASM 传输和内存仍未实测，不声称已优化或全链路零拷贝。
