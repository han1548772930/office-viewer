# SHEET-23 / SHEET-24：ΣValues 的 position 与 index

记录日期：2026-11-26。状态：**核心实现和合同草稿已落盘，未验收**。

本批补齐值字段在行/列轴中的具体插入位置，普通报告与 owned 图表共用同一个透视计算器。按用户要求只做静态来源读取、编辑和草稿清点；未运行测试、类型检查、构建、Cargo/Rust/WASM、dev、服务、浏览器、性能基准或 `git diff --check`。不得将草稿数写成通过数。

## 固定来源

固定 OSS 为 `E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 发布文本位于项目 `.research.local/univer-1.6.0/packages/`，只查阅公开类型及 ES，没有安装或执行商业 SDK。

| 具体固定来源 | 可观察契约与本地差异 |
| --- | --- |
| `engine-pivot/lib/types/types/json-types.d.ts` 第 54–63 行 | snapshot 将 `valueIndex` 与 `valuePosition` 独立持久化，ΣValues 不成为原始字段。本地把它们放在 `PivotTableOptions` 中，实际值字段及 dimension local positions 保持原模型 |
| `engine-pivot/lib/types/types/enum.d.ts` 的 `PivotTableValuePositionEnum` | None=-1、Column=1、Row=2。本地领域使用 `none/column/row` 字符串，公开 API enum 适配由主任务负责 |
| `engine-pivot/lib/types/pivot/pivot-table.d.ts` 第 52–69 行 | `updateValuePosition(position,index)`、`getValueIndex()`、`getValuePosition()`；index=-1 表示特殊字段不存在/自动位置。本地 raw getter 保留该 sentinel，布局用当前轴末尾 |
| `sheets-pivot/lib/types/facade/f-pivot-table.d.ts` 第 593–613 行 | 公共更新接受位置与零基索引，示例以多 value 为前提。旧内部单参数调用仍以 index=-1 兼容 |
| `sheets-pivot-chart/lib/types/types.d.ts` 第 117–119 行、`facade/f-pivot-chart.d.ts` 第 219–226 行；对应 ES `type:'updateValuePosition',position:t.position,index:t.index` | owned/linked 编辑都传 position + index。图表 edit DTO、完整 fields config 与普通报告调用相同核心 |
| `engine-pivot/lib/es/index.js` 的 `updateValuePosition`、`prepareLayoutCtx`、`shouldMultipleNode`、`shouldMultipleTotal` 与行/列 header gap helpers | layout 中 index=-1 取末尾；多个 measures 且 None 时 fallback column；超长 index clamp 到轴 depth。本地保留 raw index，只在派生 layout clamp，不删 measures |
| `sheets-pivot/lib/es/plugin-VCBFwx6j.js` 的 `movePivotField`、`updateValuePosition` 分支 | 字段移动同步特殊 gap。本地保留 dimension 的纯 local position，插入 at-gap 放在 ΣValues 前，移除只扣其前面的字段；没有照搬上游 mutation 或内部位置实现 |

本批不新增顶部按钮，完整入口仍见[工作区来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)和[透视图 UI 记录](SHEET-24-pivot-chart-ui.md)。index helper、字段数或测试声明不能证明入口功能完成。

## 模块、协议与实际行为

按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，领域 DTO 在 [model.ts](../../packages/office-studio/src/domain/model.ts)，新纯服务 [pivot-values-position.ts](../../packages/office-studio/src/features/pivot/services/pivot-values-position.ts) 负责 raw/effective position、gap 保持及 measure 轴投影，不访问 UI 或 DOM。

领域契约为 `PivotValuePosition = 'none' | 'row' | 'column'`，`PivotTableOptions.valueIndex?: number`。省略或 -1 是自动当前末尾，非负安全整数是明确索引；无效类型、非整数、无限值或小于 -1 拒绝。明确的超长索引可持久化，派生布局 clamp。

[pivot-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-engine.ts) 提供：

- `getPivotValuesPosition(report): { position, index }`：返回防御性配置 snapshot，省略 index 返回 -1。
- `setPivotValuesPosition(workbook, report, position, index = -1, mode = 'worksheet')`：普通/owned 调用相同 options setter；自动末尾规范化为省略 index，重复状态 no-op。
- 先投影真实 dimension scopes，再按指定 gap 将 measures 插在层级之间。矩阵行/列排序、Values header 的物理位置、dimension header 偏移及 repeat-label 判断同步变化。
- subtotal、collapsed 和 grand-total 的 source scope、measure local ID、Show Values As、数字格式与 drill-down map 继续来自原聚合器；未建立第二个权威计算器。
- 真实字段 add/move/remove 保留特殊 gap。自动末尾在轴变化后自然保持尾部，不冻结为旧数值。

[pivot-chart-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-engine.ts) 的 `updateValuePosition` edit 新增可选 `index`，完整 `PivotChartFieldsConfig` 可含 top-level `valuePosition/valueIndex`，options 也允许相同字段；两处值不一致时拒绝。缺少这些字段的完整配置保持旧 position/index。linked 在 detached 候选完成输出冲突判断后才发布；owned 继续不虚构 worksheet generated cells。

[pivot-source-lifecycle.ts](../../packages/office-studio/src/features/pivot/services/pivot-source-lifecycle.ts) 在来源列退役及 shared cache 清理两条路径保留 gap。报告 clone/copy 保留 raw position/index；源字段被移除时 local identity 和剩余顺序继续权威，不让 ΣValues 冒充 canonical source。

[workbookNormalizer.ts](../../packages/office-studio/src/parser/workbookNormalizer.ts) 的普通报告 options normalization 保留 `none` 与非负 `valueIndex`。这仅补当前解析 DTO 边界，未增加 Rust writer 或原生 XLSX `dataPosition/dataOnRows` 写出。

API、Facade、Builder、命令、UI drag/menu 与 root barrel 接线由主任务和 UI 任务完成；本记录只证明上述核心代码已落盘，不将这些 exports 标为完整浏览器功能。没有新增 Worker DTO；后续 typed job 必须携带相同 position/index 和轴输出契约。

## 未运行草稿

新 [pivot-values-position.contract.test.mjs](../../packages/office-studio/tests/pivot-values-position.contract.test.mjs) 静态清点 **15 条 test 声明、83 个计划用例，全部未运行**。复用 [pivot-axis-state-fixture.mjs](../../packages/office-studio/tests/fixtures/pivot-axis-state-fixture.mjs) 的 32 条真实 source records：行与列各 3 层、SUM/AVG 两种 local measure、普通 linked 与 owned、总金额 1000。预期合计直接由原始 records 的 scope 计算，未以渲染快照代替业务断言。

计划覆盖两种所有权 × 两条轴 × 四个位置的真实顺序/header/聚合，自动末尾/越界 clamp、None fallback、单 measure legacy、字段 gap 移动/增删、来源字段退役、完整配置/冲突、非法参数原子性、subtotal/grand-total/collapse/drill-down、linked output 扩展冲突、JSON 重开及 no-op。为新增合法 index，既有 [pivot-chart-engine.contract.test.mjs](../../packages/office-studio/tests/pivot-chart-engine.contract.test.mjs) 的非法参数草稿从 index=1 改为 index=-2；未运行该文件。

## 差异、缺口与未验收

1. 现有本地单 measure row renderer 仍显示 Values header，以保留旧内部调用兼容；固定上游只有多 measures 形成特殊 ΣValue 维度。UI 可按多值显示特殊字段，但单值最终交互仍需统一迁移及浏览器样本验收。本批明确保留差异，不宣称已完全对齐。
2. None 保留模型中的 absent 配置，渲染时按固定 engine fallback columns；这不是删除 measures。原生 Office 的相同状态和索引边界仍缺独立 XLSX 样本。
3. XLSX 原生 cache、report/owned parts、值位置、保存重开与 Office 续算写出未完成；JSON 草稿不等同于 Office 保真。
4. 当前轴计算仍在主线程 TypeScript，Worker/Rust/WASM、取消、TypedArray lease、复制次数/字节与大样本内存/帧时间没有测量。[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)仍适用，不声称零拷贝或性能达标。
5. 本批核心无新增可替换 UI 控件；UI 的 `antdv-next`、ConfigProvider、主题切换、焦点和 IME 与全部浏览器闭环尚未验收。

完整 Sheets / Docs / Pro 目标保留，上述缺口不是最终范围缩减。
