# Office Viewer / Office Studio

使用 Vue 3、TypeScript、Pinia 与 Rust/WASM 实现 Sheets / Docs 办公工作台。界面组件使用 antdv-next，正文与网格使用 Canvas；渲染后端优先 WebGPU，不可用时回退 Canvas2D。完整目标包括 Univer Sheets / Docs 的 Pro 能力，当前各功能仍按清单保持未验收，不将已有 API、按钮或构建成功等同于完整对齐。

## 开始开发

先按 [AGENTS.md](AGENTS.md) 顺序阅读项目标准、参考规则、功能清单与渲染计算标准。JavaScript 包管理只使用 pnpm。

准备支持 `--experimental-strip-types` 的 Node.js、pnpm、Rust 工具链和 wasm-pack。首次运行需生成项目内的三个 WASM 模块：

```powershell
pnpm install --frozen-lockfile
rustup target add wasm32-unknown-unknown
pnpm build:wasm
pnpm build:compute
pnpm build:writer
pnpm dev
```

PowerShell 若禁止执行 `pnpm.ps1`，使用 `pnpm.cmd` 执行相同命令。开发工作台位于 `examples/workbench`，默认端口 5174。WASM 输出分别放在 `public/wasm`、`public/compute` 和 `public/wasm-writer`；缺少这些构建产物时，真实文件导入、计算或导出不能完成。

## 代码位置

- `packages/office-studio/src/api`：公共 Facade、统一命令和应用装配。
- `packages/office-studio/src/domain`、`parser`：领域模型、序列化及解析 DTO 归一化。
- `packages/office-studio/src/features`：公式、筛选排序、图表、透视、切片器和文档业务逻辑。
- `packages/office-studio/src/rendering`、`components`：绘制后端、视口、输入桥和工作台组件。
- `crates/office-parser`、`office-compute`、`office-writer`：项目内的解析、计算与真实 Office 写出模块。
- `examples/workbench`：可运行的工作台集成示例。

## 验证

```powershell
pnpm check:types
pnpm test:unit
pnpm test:components
pnpm test:browser:features
pnpm test:browser:toolbars
pnpm test:browser
cargo test --locked --manifest-path crates/office-parser/Cargo.toml
cargo test --locked --manifest-path crates/office-compute/Cargo.toml
cargo test --locked --manifest-path crates/office-writer/Cargo.toml
pnpm build
pnpm test:browser:production
```

工作台浏览器测试使用本机 Microsoft Edge 的 headless 模式；计算专项配置使用 Chrome。实际 Office 保存/重导入测试还需要上面的 WASM 构建产物。`test:browser:features` 覆盖切片器、时间线与内容控件；`test:browser:toolbars` 联合覆盖 Sheets 样式/格式/函数/筛选排序、Docs 剪贴板/符号/表格上下文和 Pro 控件。内部快照测试不代替 XLSX/DOCX 往返保真。

完整范围、逐项证据及剩余缺口见 [功能对齐清单](docs/UNIVER_FEATURE_PARITY.md) 和 [本轮续作记录](docs/feature-evidence/continuation-2026-11-20.md)。性能与完整 Office 互操作分别验收。
