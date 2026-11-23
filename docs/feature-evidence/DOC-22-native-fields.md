# DOC-22 正文原生字段与段落来源保留

更新：2026-11-24。状态：实现中，**未验收**。本批只静态阅读与修改代码、编写测试草稿；没有执行测试、类型检查、构建、Rust/WASM、浏览器或文件重开。用户要求先完成实现，再统一验证。

固定参考为 `E:/rust_test/univer` 的 v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。`packages/core/src/types/interfaces/i-document-data.ts` 中的 `IFieldRangeProperties` 区分指令、结果、锁定、指令文本样式和 `instructionParts`；`IFieldCustomRange` 允许结果内嵌套字段。`packages/docs-toc/src/commands/table-of-contents.command.ts` 的生成/更新路径保存 TOC 外层和 PAGEREF 内层，并按范围更新。本项目使用段落 ID 与半开 UTF-16 偏移表达可见结果，使用原生 DOCX 字段边界写出，不复用 Univer 内部数据流哨兵。

| 文件 | 本批实际变化 |
| --- | --- |
| [document_block_extraction.rs](../../crates/office-parser/src/document_block_extraction.rs) | 在真实 duke-pages 提取前为可能被过滤的正文/表格空段落加临时书签，提取后从源模型和输出移除。扫描正文、对象内容及其他 story 的真实书签名称，避免移除原有同名前缀书签。 |
| [parser.rs](../../crates/office-parser/src/parser.rs) | 正文解析调用该适配，再走已有 sidecar 和原生字段恢复；仍使用真实 duke-pages 解析模型。 |
| [document_text_fields.rs](../../crates/office-parser/src/document_text_fields.rs) | 保留字段身份、指令 XML、锁定、结果范围和 `parentFieldId`，新增原 XML 开始顺序 `beginSequence`。当前 runs 是导出结果的权威来源，导入 `cachedResult` 仅是来源缓存。 |
| [docx_field_topology.rs](../../crates/office-writer/src/docx_field_topology.rs) | 按显式父身份构建字段树。`null` 是显式根，缺少父来源才推断几何嵌套；检查悬空、自引用、循环、包含范围和兄弟交叉。父子可以为空或拥有相同范围。嵌套上限为 512，迭代遍历避免递归栈耗尽。 |
| [docx_text_fields.rs](../../crates/office-writer/src/docx_text_fields.rs) | 以字段树的 DFS 顺序生成同优先级开/关事件，保留父开→空子→父关、首尾空子与等范围父子。相邻根先关后开。结果语言使用实际树深度。 |
| [document_text_fields_roundtrip_tests.rs](../../crates/office-parser/src/document_text_fields_roundtrip_tests.rs) | 新草稿覆盖无书签空字段、跨空段结果、空/等范围父子、首尾空子、相邻空根和无效父图的实际 DOCX 导出/重开。 |

duke-pages 的本机版本为 0.6.0；具体依据是 `C:/Users/15487/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/duke-pages-0.6.0/src/structured.rs::without_empty_blocks`、`has_paragraph_content` 与 `model.rs::ParagraphBookmark`。依赖会过滤不可见正文并扁平化仅对象段落，因此本项目在提取边界保留可寻址段落，而不是在导入后凭字段偏移补造文本。

Docs TOC 的结构重建同时修复旧 PAGEREF 的显式父身份，生成页码归属当前 TOC 外层；详见 [DOC-04](DOC-04-table-of-contents.md)。这些事务与 Rust 写出仍待统一集成验证。

剩余缺口：指令区内的嵌套字段及被依赖投影进正文的隐藏代码文本当前仍保留来源并明确拒绝不安全导出，尚未完成与 `instructionParts` 等价的编辑/投影/写出闭环；其他 story 的空段适配、字段与复杂 SDT/批注/编号/对象相交、原生 Office 独立重开及大样本性能仍未验收。本文不表示 DOC-22 或完整工作台功能已经完成。
