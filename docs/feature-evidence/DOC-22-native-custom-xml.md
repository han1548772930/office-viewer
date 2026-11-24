# DOC-11 / DOC-22：原生内容控件与自定义 XML 数据源

核对日期：2026-11-25。基线为 Univer v1.6.0，固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。状态：开发中、未验收；用户要求先实现后统一测试，本批未执行测试、类型检查、构建、Rust/WASM、浏览器或基准。

## 固定来源与可观察契约

- `E:/rust_test/univer/packages/core/src/docs/sdt-binding.ts` 的 `getSdtBindingKey`：只有非空 `storeItemID` 与 `xpath` 同时存在才激活数据绑定。
- `E:/rust_test/univer/packages/core/src/docs/__tests__/sdt-binding.spec.ts`：单独 store、单独 XPath、空配置是休眠绑定，不应按激活绑定要求其数据源存在。
- `E:/rust_test/univer/packages/core/src/types/interfaces/i-document-data.ts`：SDT 属性、XML 数据源和绑定配置与文档内容属于同一持久模型。十类控件入口、父子控件、锁与 UI 来源见 [DOC-11 清单](DOC-11-content-controls.md)。
- 本地固定 `duke-pages 0.6.0/src/parse/paragraph.rs`、`src/model.rs::FormFieldRunNode` 只给出简化 SDT 投影；原始 `sourceXml` 不等于本项目已接入全部控件属性和结构范围。
- WordprocessingML 的 `w:sdtPr/w:dataBinding`、`w:sdtContent`、`w15:repeatingSection`、`w15:repeatingSectionItem` 及 OPC customXml 关系是本项目原生适配目标；Pro 内部实现未公开，不能将本地算法写成上游实现事实。

## 已落盘的数据链

- 领域 `DocumentCustomXmlPart` 保存 Unicode XML、真实 part/store 身份、properties、关系、相关二进制 part、UTF-8/UTF-16 编码和 BOM。共享 Rust `document_custom_xml` 校验 OPC 名称、内容类型、关系图、缺资源、URI 别名歧义和有界资源。
- Parser 从真实 DOCX ZIP 读取数据源及依赖；writer 按项目 OPC 资源图写回。共享关系、循环相关资源、非规范 ZIP 名与 URI 关系名分别处理，不将业务 XML 误认为本项目 sidecar。
- 构造、create、XML source setter、binding setter 与普通候选事务接入资源/绑定校验。激活绑定刷新同一 store/XPath 的控件；候选失败不发布局部文本、资源、历史或字数更新。
- XML 编辑和内容控件命令复用权限、父级内容锁、输入会话守卫与文档内容历史。source getter 返回隔离副本；快照只保留一份权威 `customXmlParts`。
- 休眠配置按已提供字段校验类型和长度；只有激活绑定才要求明确数据源匹配，并写原生 `w:dataBinding`。这与“已配置即激活”不同，不能把空 XPath 当作错误绑定。
- 共享 namespace-aware XML reader 的保留树 admission budget 为 64 MiB，另有 16 MiB 输入、节点与深度上限。记录的是估算收费，不是实测 heap、性能优化或全链路零拷贝证据。

## 本批测试草稿

- `packages/office-studio/tests/document-custom-xml.contract.test.mjs`
- `packages/office-studio/tests/document-custom-xml-bindings.browser.test.mjs`
- `packages/office-studio/tests/document-custom-xml-candidate.browser.test.mjs`
- `crates/office-io-common/src/document_custom_xml_tests.rs`
- `crates/office-parser/src/document_custom_xml_roundtrip_tests.rs`
- `crates/office-writer/src/docx_custom_xml_tests.rs`

草稿覆盖 namespace/XPath、休眠配置、候选原子性、权限/锁、getter 隔离、历史/快照、真实 ZIP 资源图、编码/BOM 与 writer→parser reopen。全部未执行，不能作为浏览器或文件互操作通过的证据。

## 尚未闭合的边界

- 原生 SDT 完整 DTO、UTF-16 文本宿主映射、跨段/单元格/表格控件和真实 row SDT 正在接线；仅 XML 数据源重开不证明绑定控件重开。
- 文档打开时目前验证资源和绑定，尚未自动刷新绑定文本。部分 setter/common commit 路径可能重复解析，未测端到端成本。
- 休眠绑定现在保留于项目 JSON 快照；未输出原生 `w:dataBinding`，尚无专用 DOCX 保真 sidecar，不宣称 Office 重开保留。
- external custom XML store 获取策略、非 OPC 相对 XML/schemaLocation 引用迁移、XML 1.1 与浏览器一致性、复杂富文本/图片绑定细项仍需补齐。
- 原生图形/表格/控件属性与 namespace 的完整 Office 行为、独立 Office 样本重开、大样本内存/取消/传输和 IME/主题/焦点均未验收。

后续统一验证需分别核对解析、编辑、撤销重做、JSON 保存重开、真实 DOCX 导出及独立重导入；测试通过与全部功能目录覆盖是两项独立要求。
