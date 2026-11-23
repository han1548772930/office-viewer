# DOC-05 / DOC-22 原生编号资源映射

更新：2026-11-24。状态：开发中，未验收。此记录描述源码已落盘的资源适配，不把测试草稿、依赖声明或接口当作文件保真验收。

固定参考：Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。`E:\rust_test\univer\packages\docs-ui\src\commands\commands\list.command.ts` 的 `ListOperationCommand` / `ChangeListTypeCommand` 使用选区 `segmentId`、`getSelfOrHeaderFooterModel()` 和 `getRichTextEditPath()`，将列表操作写入对应正文或页眉页脚 segment 的统一 rich-text mutation。参考命令是 `doc.command.list-operation` / `doc.command.change-list-type`，不是全篇字符串替换。Pro 增强列表仍按 [DOC-05 清单列表记录](DOC-05-checklist.md) 和固定公开规格继续展开。

本项目列表与序列求值仍使用既有文档模型和 `docx_numbering.rs`；此次新增的是 Office 原生来源的资源依赖映射。原生表格条件样式、页眉页脚段落及样式继承中的 `numPr` 不再因编号依赖直接拒绝写出。

## 来源与边界

`officeSource.numbering` 保存实际 `partName/xml/relationshipsXml/resources`；媒体包含规范包路径、内容类型与标准 Base64。DOCX 读取编号所在 part 及其真实媒体字节，外部图片链接保留关系，读取过程不获取远程资源。DOC 复用 Duke 为同一模型生成的 `OoxmlPackage`；其 Word 编号资源属于合成适配，不能宣称是原始 DOC XML。ODT 原始来源不允许冒充原生 Word 编号。

共用适配位于 [document_numbering_source.rs](../../crates/office-io-common/src/document_numbering_source.rs)，解析位于 [parser 编号来源](../../crates/office-parser/src/document_numbering_source.rs)，写出位于 [writer 编号 registry](../../crates/office-writer/src/docx_source_numbering.rs)、[依赖图](../../crates/office-writer/src/docx_source_numbering_dependencies.rs) 和 [媒体映射](../../crates/office-writer/src/docx_source_numbering_media.rs)。前端验证在 [documentOfficeSource.ts](../../packages/office-studio/src/parser/documentOfficeSource.ts)。新增 Base64 依赖的锁文件由最终依赖解析正常生成，未手改。

写出先预留当前列表生成的 num、abstractNum、numPicBullet 与 VML ID，再按使用者可达依赖导入源资源；样式、实例、抽象、图片项目符号、pStyle/rStyle、styleLink/numStyleLink、级别覆盖和主题引用使用同一个 registry。编号 0 是抑制标记，不要求源编号。styleLink/pStyle/样式 numPr 的合法关联回链与真正的 numStyleLink 委托循环分开处理；basedOn 与委托循环拒绝，来源对象保留。

媒体使用独立输出路径；重复引用同一源媒体只写一次，关系 ID 和 VML shape/type/spid 避让。styles、numbering、theme 的最终打包推迟到所有当前消费者准备完成后，避免后准备的表格样式落在已经打包的资源之外。

单 XML 和单媒体各限 16 MiB，编号 XML、关系与媒体合计限 32 MiB，媒体项限 1024。路径、重复媒体、未知字段、Base64 填充及字节预算在 Rust 和 TypeScript 边界拒绝；读取失败保留可读取的 XML/资源并记录 resourceErrors。Base64 编码/解码与 OPC 输出存在必要复制，此链路不称为零拷贝。

## 尚未执行的验证

- common 草稿覆盖关系路径、规范 Base64、重复媒体和未知字段。
- writer 草稿覆盖 ID 冲突、条件表样式和 story 共用实例、级别覆盖、合法回链、继承的 numStyleLink 委托、图片关系/直接 src、默认 VML 图片类型、缺失媒体、循环和编号抑制。
- parser ZIP 草稿覆盖非默认 owner 路径、内部真实字节、外部链接以及缺失资源诊断；[DOCX 草稿](../../crates/office-parser/src/document_numbering_roundtrip_tests.rs) 覆盖表格条件样式与页眉段落共用编号、打包媒体及 Duke 重开。
- [TypeScript 草稿](../../packages/office-studio/tests/document-numbering-source.contract.test.mjs) 覆盖来源保留、编辑/历史/内部重开、预算与非法输入。

按用户要求以上草稿未运行；测试、类型、前端和 Rust/WASM 构建、浏览器、Office 独立打开及真实性能测量都留待统一验证。资源图不是完整 OOXML schema 验证器，重复必需子节点、复杂 DrawingML/VML/扩展及原生 Office 可观察效果仍待核实。页眉页脚和注释真实块树、表格/控件/字段跨 host 写出继续接线，不能由此次编号适配推断所有列表或所有 story 已完成。
