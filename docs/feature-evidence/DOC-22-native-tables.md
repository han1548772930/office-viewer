# DOC-06 / DOC-22 原生表格读写与来源保留

更新：2026-11-24。开发中，**未验收**。本批按用户要求先实现、后统一验证；只读取源码、修改实现和编写测试草稿，未执行测试、类型检查、构建、Rust/WASM 检查、浏览器、dev 或基准。

## 来源与责任边界

- 固定 Univer：`E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Docs 表格与 Pro 入口、分组、命令及公开 Facade 依据见 [DOC-06 入口清单](DOC-06-enhanced-tables.md)。所有 Docs / Sheets Pro 能力继续在范围内。
- 公开对照：[Docs 表格](https://docs.univer.ai/zh-CN/guides/docs/features/table)、[Office 转换规则](../UNIVER_REFERENCE_RULES.md#12-exchange-office-互操作对照入口)。转换由本项目 `duke-pages` → parser DTO → 领域模型 → `office-writer` 实现，不依赖 Univer Server、Exchange 或商业 SDK。
- 当前解析依赖 `duke-pages 0.6.0`；本地 `src/parse/table.rs` 解析 `gridSpan` / `vMerge`、解析行格样式，`src/structured.rs` 提取逻辑列与正文。其 `indexed_cell_property_tokens` 只保留 `hMerge`，没有折叠旧式横向合并，本项目补独立适配。
- 计算/格式算法位于 Rust 读写层和独立领域模块，Vue / `antdv-next` 负责可替换 UI。界面、IME、Canvas、权限、历史与富文本粘贴验收由 DOC-06 对应证据记录；此文不将 IO 代码等同于整组产品完成。

## 已写入的实现

- `crates/office-io-common/src/document_table_source.rs`、`chart_xml.rs`：有预算的 namespace-aware XML、原生表格/拓扑 signature、作者 source entry 校验。
- `crates/office-parser/src/document_block_extraction.rs`：在真实 Duke 提取路径临时锚定空段、不可见字符和仅对象段，返回前移除内部 bookmark，保留稳定正文位置。
- `crates/office-parser/src/document_table_sidecar.rs`、`document_table_merges.rs`：实际物理行、嵌套表、表/行/格原生属性与 namespace 上下文；作者来源身份核对；旧式 `hMerge` 的逻辑格与正文归一化。
- `packages/office-studio/src/parser/documentTables.ts`、`documentNormalizer.ts`：compact anchor grid、真实列宽/行高、表头、列类型与下拉选项、原生属性 baseline 和来源 metadata；非法跨度、重叠、空洞及越界显式拒绝。
- `crates/office-writer/src/docx_table{,_styles}.rs`、`docx_source_tables.rs`：一个网格驱动原生表格、合并、物理续行与嵌套表；CSS px → twips，真实列宽及父格可用宽度；表/行/格样式、条件表格样式、theme 和 style ID 冲突重映射。
- 编号、正文字段、notes 和 controls 递归消费表格真实正文容器；`docx_field_topology.rs` 维护跨容器字段的安全边界。文档输入 session 已接入模型替换后的结构范围与稳定段落 ID 重新定位。

## 协议、来源与重新保存

- `columnIndex` 表示逻辑列；`TextTarget.col` 表示该行 `cells` 的存储索引。完全被跨行合并覆盖的物理行保留 `cells: []`，不创建假格、不夹取跨度。
- 作者扩展位于 `customXml/office-tables.xml`，namespace 为 `urn:office-studio:document-tables:1`；`ov:key`、tableId、native signature 与 topology signature 共同核对身份。文本变化保留作者设置，几何/样式变化使用当前 native，拓扑变化只保留显式 stale provenance。
- 属性与 signature 使用 namespace + local name，完整 expanded attributes 参与签名；同名扩展属性不能冒充 Word 属性。`hMerge` 已纳入 topology signature。
- 非空跨行续格使用 `style.nativeContinuationContent[offset] = { version: 1, ownerCellId, rowId, colSpan, blocks, source }`。以稳定 `rowId` 定位，offset 只作映射位置；当前 native 正文重建来源，旧 sidecar 不会复活被外部删除的文字。writer 的 `content_groups()` 按物理行/逻辑列顺序供正文、字段、编号、notes 和 controls 共用。
- 旧式同一行 `hMerge restart/continue` 折叠为一个逻辑 `columnSpan`，正文按成员顺序并入 owner，保留每段 ID 与 source refs。按折叠后的完整列区间重新核对 `vMerge`，一次物理行只增加一次 rowspan；孤立续格、非连续成员与不匹配的纵向区间拒绝。
- 原始横向成员的属性、resolved style、cell ID 与 block ID 保留为 `nativeProperties.sourceHorizontalMerge`，领域层将其保存在 `style.nativeHorizontalMergeSource` 的物理行来源记录中。它只作 provenance，不决定当前网格或恢复旧文字，也不进入正 offset 的纵向续格协议。
- 横向归一化后的 `tcW` / `tcBorders` 通过 `rewriteProperties` 重写，`gridSpan` 和 `vMerge` 由实际领域网格写出；不继续使用首成员的窄宽度。标准现代 `gridSpan` 表达替代 legacy `hMerge`，原始成员属性仍保留。

## 预算与测试草稿

- writer：10,000 物理行、512 逻辑列、整篇 1,000,000 grid slots、32 层、10,000 表；作者 metadata 共 16 MiB。归一化与 legacy merge adapter 也核对各自逻辑网格预算。预算存在不代表峰值内存或大样本性能已达标。
- 测试草稿：`crates/office-parser/src/document_table_roundtrip_tests.rs`、`document_table_merges.rs`；`crates/office-writer/src/docx_table_tests.rs`、`docx_source_tables_tests.rs`；`packages/office-studio/tests/document-table-import.contract.test.mjs`。
- 草稿覆盖真实 writer → Duke parser、覆盖行/嵌套表、来源身份与外部改动、namespace 冒名、续格正文/字段/列表、legacy 横纵合并、原始段落顺序、成员 provenance 和再次写出宽度。全部未运行；独立 Office 重开、浏览器编辑/撤销重做和 IME 尚未验证。
- XML/metadata 树转换和正文拼接会产生必要的对象/文本分配；没有整条链路零拷贝或性能优化达标的声明。复制字节、Worker/WASM 峰值内存及同数据集前后对照留待统一基准。

## 剩余差异与未核实项

- 原生 table/paragraph style 的 `numPr` 仍需导入原 numbering 资源并重映射 ID；当前明确拒绝 `DOCX_TABLE_STYLE_NUMBERING_UNMAPPED`，不能把拒绝路径标为支持完成。
- `gridBefore/gridAfter` 不规则行、零/塌缩列、零 exact height、tracked row/cell/property revisions 与样式编号仍须补齐。
- legacy 横向成员不一致的纵向状态、分段边框和复杂原生样本的 Office 可观察效果尚未核实；当前矩形区间规则及现代 span 输出需要实际文件验证。
- story 的完整 table blocks、跨 story 复杂 SDT/字段/批注/编号相交与富文本粘贴等编辑闭环继续由 Docs 模块实现；此批不能代替这些能力验收。
- 无真实大样本、浏览器或 Office 样本验证，DOC-06、DOC-22 和整个工作台保持未验收。
