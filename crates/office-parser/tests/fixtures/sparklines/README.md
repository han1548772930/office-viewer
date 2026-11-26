# SHEET-26 原生格式夹具草稿

这些 XML 是本项目合成的字段/OPC 合同输入，不是从用户文件提取的素材，也不是 Microsoft Office 打开过的独立互操作样本。没有运行生成 ZIP、测试、浏览器或任何验收。

- `native-settings-sheet.xml`：三个原生图型、两成员分组、带 apostrophe/空格/感叹号的跨表绝对引用、独立 min/max 模式、inactive manualMin、六类标记的颜色与显隐分离、日期轴、非整数 point 线宽、ARGB alpha、主题 tint 和自定义 indexed 颜色。
- `theme.xml`、`indexed-colors.xml`：小型颜色依赖合同片段；主题仅列需要的颜色结构，不能宣称通过完整主题 schema 验证。
- `sparkline_sidecar_tests.rs` 中的未运行 helper 将这些片段写入 relocated OPC ZIP；读写重开草稿另用真实 office-writer 构建主链路文件。

格式依据与差异：[Office 读写证据](../../../../../docs/feature-evidence/SHEET-26-sparkline-office-roundtrip.md)。固定 Univer 配置基线为 v1.6.0；OSS commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。
