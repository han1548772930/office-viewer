# COMMON-11 / SHEET-29 / DOC-22 WASM 依赖与编译适配

记录日期：2026-11-28。状态：开发中，**未验收**。本批按用户此前明确要求修复 WASM 编译错误并更新依赖；仅必要的 wasm32 库编译检查提前执行，功能测试及统一验收顺序保持不变。

## 实际依赖与改动

版本以各 crate 的 Cargo.toml / Cargo.lock 为准。当前解析链直接采用 `quick-xml 0.42.0`、`zip 9.8.2`、`duke-pages 0.6.5` 和 `rxls 0.1.4`；上游传递依赖仍可带其他 quick-xml / zip 版本，不将其误报为本项目直接绑定。WASM 平台锁定 `wasm-bindgen 0.2.130`、`js-sys/web-sys 0.3.107`、`wasm-bindgen-futures 0.4.80`、`serde-wasm-bindgen 0.6.5`；浏览器绑定测试使用 `wasm-bindgen-test 0.3.80`。

已修改以下真实读写与共享路径：

- `crates/office-io-common/src`：字符串化 QName / Namespace、属性值规范化、XML 声明与 processing instruction 的新版 API，以及共享错误适配。
- `crates/office-parser/src`：XML 名称/属性与 ZIP 名称的新版 Result/Cow API，批注模块的可见性与导出，以及 sparkline 所有权移动；不替换 duke-pages / rxls 的读取链路。
- `crates/office-parser/src/table_sidecar.rs`：最后一处弃用的 `unescape_value()` 改为 `normalized_value(quick_xml::XmlVersion::Implicit1_0)`，保留属性实体与换行规范化语义。
- `crates/office-writer/src`：DOCX 源 XML、自定义 XML 与编号资源的新版 XML/ZIP API；writer 的 Cargo.lock 由 Cargo 正常生成。
- parser 内的条件格式、自定义 XML、文档格式/legacy/story/table、工作表尺寸/筛选/大纲/验证与 sparkline 往返草稿：ZIP `.name()` 显式解包 Result 并取得 owned 字符串；writer 自定义 XML 草稿对 `file_names()` 的每项先解包再过滤。这些是未运行的草稿兼容改动。

现有 document-format、dimension、pivot 与共享 XML 树读取路径已有 GeneralRef 分支；不因升级盲目删除或再添加第二条实体读取路径。此静态事实仍须由真实文档与属性/文本实体样本验证。

## 已执行的检查

恢复的原线程记录与本次复查均使用项目内 crate，未生成 public WASM 产物：

| 命令 | 结果 | 证据边界 |
| --- | --- | --- |
| `cargo check --manifest-path crates/office-parser/Cargo.toml --target wasm32-unknown-unknown --lib --message-format short` | 成功；2026-11-28 在最后 table 属性改动后再次执行，exit 0 | 检查生产库编译，不包含 cfg(test) 草稿或浏览器绑定执行 |
| `cargo check --manifest-path crates/office-compute/Cargo.toml --target wasm32-unknown-unknown --lib --message-format short` | 原线程同批成功 | 仅生产库编译；当前数值行为、Worker 与内存绑定仍未验收 |
| `cargo check --manifest-path crates/office-wasm-bridge/Cargo.toml --target wasm32-unknown-unknown --lib --message-format short` | 原线程同批成功 | Promise / DOM / TypedArray 的真实浏览器契约尚未执行 |
| `cargo check --manifest-path crates/office-writer/Cargo.toml --target wasm32-unknown-unknown --lib --message-format short` | 原线程同批成功 | 不证明 XLSX / DOCX 输出与独立重开保真 |

仍有既有 unused / dead-code 等警告，本批不扩大到无关清理。没有运行 Rust 测试、wasm-bindgen-test、前端测试、类型检查、构建、dev、浏览器、Office 往返或性能基准。

## 待统一验证

统一验证阶段先生成协议 2 的 parser / compute / writer / bridge 产物；`public/` 现有生成物不由上述 cargo check 更新。随后编译并执行原生与 wasm32 测试草稿，验证 XML 实体/名称/属性、ZIP 名称、未知源 part 保留、DOCX/XLSX 往返，以及 Promise/取消、TypedArray lease、memory.grow 与 DOM/输入生命周期。旧数值基准不能验收这次依赖和绑定改写；性能仍使用相同数据集重新记录端到端耗时、传输复制及内存。

本记录只证明已落盘的兼容修复和以上编译结果；65 个功能组继续保持未验收。

关联：[功能清单](../UNIVER_FEATURE_PARITY.md)、[WASM 平台契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)、[既有基准及协议 2 草稿](COMMON-11-compute-benchmark.md)。
