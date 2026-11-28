# COMMON-11：WASM 打包与 Vite 重复导入修复

用户在运行 `pnpm dev` 时先报告找不到 wasm-pack，随后报告 WASM 编译错误与警告，并贴出 Vite 的 `Identifier HistorySavepoint has already been declared` 日志。本记录只覆盖这些构建问题；完整工作台功能、Office 往返、浏览器与性能仍未验收。

## 实际修改

- 根 [package.json](../../package.json) 将 `dev` 接到 `build:wasm:all`，按 parser → compute → writer → bridge 顺序构建，任一失败即中止。wasm-pack 固定为项目开发依赖 0.15.1，[pnpm-workspace.yaml](../../pnpm-workspace.yaml) 允许它的必要安装脚本，锁文件由 pnpm 正常生成，无全局安装。
- [formula_calendar_reference_scan.rs](../../crates/office-parser/src/formula_calendar_reference_scan.rs)：适配 ZIP 9 的可失败文件名迭代，先完整拥有名称再读取条目；名称无法解码时保留计算 helper 并给出诊断，不能把不完整扫描当成无引用的证据。
- [formula_calendar_sidecar.rs](../../crates/office-parser/src/formula_calendar_sidecar.rs)：对解析 DTO 的 u16 列坐标使用 `u32::from` 无损转换，与引用索引和指纹类型一致。
- [xlsx_formula_calendar.rs](../../crates/office-writer/src/xlsx_formula_calendar.rs)：返回名称可能借用输入或投影，显式统一二者与返回值的生命周期。
- 逐调用核对后清理 shared/parser/compute/writer 的未使用导入、参数、不可达分支和过时私有包装；正文抽取、文本输出与透视断言专用辅助项用 `cfg(test)` 保留。StyleInput 继续接收 `metadata`，以显式 serde 名称保留该字段；没有全局禁用 Rust lint。
- [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts)：删除 `HistorySavepoint` 的一处重复 type import，保留唯一 `history-types.ts` 导入。根因是有/无扩展名的两条 import 在同一模块绑定同名标识符。

## 已执行与未执行

针对用户明确报告的构建失败，提前执行 WASM release 构建；这属于统一验证后置规则的本次明确例外。四模块已分别通过，随后聚合入口也完整通过，包含 wasm-bindgen 与 wasm-opt。清理警告后的聚合输出记录为退出码 **0**、Rust error **0**、Rust warning **0**。

| 模块 | 真实生成目录 |
| --- | --- |
| office-parser | `public/wasm` |
| office-compute | `public/compute` |
| office-writer | `public/wasm-writer` |
| office-wasm-bridge | `public/wasm-bridge` |

构建输出中关于 Cargo `description/repository/license` 的缺省字段说明是 wasm-pack 的可选元数据 INFO；没有为消除这条信息杜撰许可证或仓库地址。构建不会发布包。

Vite 报错修复后，用项目已安装 Vite 的 `transformWithOxc` 转换实际 `office-studio.ts`，转换成功且 warnings 为空。只读扫描 **1,142** 份库与应用 TS/TSX/Vue 源文件的 import 绑定，重复绑定 **0**。用户现有 5174 开发服务器的该模块及筛选条件组件请求均返回 **HTTP 200**，没有启动第二个 dev，也没有运行浏览器交互。

单元/契约草稿、类型检查、前端完整构建、浏览器、Office 文件往返与性能基准仍后置；WASM 编译成功和 HTTP 200 不能替代上述功能验收。
