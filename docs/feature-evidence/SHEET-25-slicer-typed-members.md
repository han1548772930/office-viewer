# SHEET-25：共享切片器的 typed 成员与 Table 原生限制

记录日期：2026-11-26。状态：**模型、引擎及合同草稿已落盘，未验收**。

本批连接 Pivot raw filter 与共享 Slicer cache，保留同 caption 不同 scalar 类型的 partial 状态，并接入既有 Table 原生 filter。按用户要求只做静态读取与编辑，未运行测试、类型检查、构建、Cargo/Rust/WASM、dev、服务、浏览器、性能基准或 `git diff --check`。

## 固定依据

OSS 为 `E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。公开 Pro 发布文本保存在项目 `.research.local/univer-1.6.0/packages/`，没有安装或执行商业 SDK。

| 固定包内来源 | 观察与本地差异 |
| --- | --- |
| `engine-pivot/lib/types/record-filter.d.ts`、`record-value.d.ts` | 原始值筛选区分类型，include/exclude/all 与显示文本分开。本地当前采用 `CellScalar`，日期/error 独立 identity 尚未完成 |
| `sheets-pivot/lib/types/record-filter-bridge.d.ts` | checked/partial 状态、manual filter 到 raw members 的桥接，`getSlicerSelectionFromPivotManualFilter(... current, isAll)`。本地用真实 records 投影，组 caption 不作为 raw identity |
| `sheets-slicer/lib/types/facade/f-slicer.d.ts` 第 405–460 行 | typed getSelection/setSelection、Table native restriction marker、已删除/未出现的显式成员保留及 bounded raw member window。当前实现底层 typed authority 和原生限制派生；正式 async bounded API 由后续主任务接入，不能以同步内部 getter 声称完整公开 API 已完成 |

UI 入口仍归[工作区来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)和[切片器与时间线记录](SHEET-25-slicer-timeline-affine.md)维护。本批没有按按钮数认定完成。

## authority、同步与公开核心接口

[model.ts](../../packages/office-studio/src/domain/model.ts) 新增 `SlicerCacheModel.recordSelection?: PivotRecordSelection`，`selectedItems` 仅为 legacy/display mirror。Pivot cache 的 raw state 是 authority；Table 原生 filter 是 authority。实现按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)置于以下模块：

- [slicer-record-selection.ts](../../packages/office-studio/src/features/slicer/slicer-record-selection.ts)：typed normalization/equality、caption 投影、checked/partial 派生及 partial edit planning。类型与值共同构成身份；重复和排序不产生内容变化。
- [slicer-connections.ts](../../packages/office-studio/src/features/slicer/slicer-connections.ts)：raw source values、原生 column filter 防御性 snapshot、来源校验与 legacy 规范化。
- [slicer-report-filter.ts](../../packages/office-studio/src/features/slicer/slicer-report-filter.ts)：所有连接报告先 staging，再统一发布；保留输出冲突原子性，普通、linked owner、owned report 使用同一透视引擎。
- [slicer-engine.ts](../../packages/office-studio/src/features/slicer/slicer-engine.ts)：共享 cache/view 更新、外部报告筛选同步、断开/重连、Table typed filter、condition UI 派生与保存重开数据。

内部可复用 exports 为 `readSlicerRecordSelection`、`readSlicerItemState`、`setSlicerRecordSelection`、`setSlicerItemStates`；连接模块提供 `getSlicerSourceValues` 与 `getSlicerTableNativeFilter`。主任务在原命令、权限、可取消事件、文档内容历史及 Facade 接线，未改 Worker/WASM 或 Office writer DTO。

实际语义：

- typed numeric `1` 不选择 text `"1"`，boolean `true` 不选择 text `"true"`，null 不选择 empty string。字符串 mirror 可相同，typed 冲突仍拒绝。
- 显示 caption 完整勾选选中当前所有 scalar 类型；partial 保留原 raw membership，切换其他按钮不扩大 partial。Canvas checked-all 且没有 partial 时清 member filter，空显示集合也如此；direct include-all 保持显式 include。
- `include` 空列表过滤全部；`all` 或 `exclude` 空列表清筛选。absent/deleted 的显式 raw 值保留，后续来源成员出现仍使用原 include/exclude 语义。
- 普通报告和 owned report 的 local field IDs 可不同，按 canonical source/cache fan-out。disconnect 保留报告旧筛选，reconnect 使用当前 raw cache，零连接状态可持久化。
- 同 raw 选择也处理来源 category 互斥，合法 grouped report 可保留自己的显示规则。所有输出规划先完成，拒绝不发布局部报告/cache。
- 旧字符串值有当前真实来源时映射 raw 类型；未观测到原类型的旧成员保留，typed getter 明确拒绝猜测。

## Table 原生 filter 与 condition UI

既有 [sort-filter-engine.ts](../../packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts) 的 `SheetValueFilter` 增加可选 `selectionMode: 'include' | 'exclude'`；省略仍为 include。exclude 在原 `matchesCriteria` 中执行，不增加第二个原生 filter 引擎。[workbookAutoFilter.ts](../../packages/office-studio/src/parser/workbookAutoFilter.ts) 保留该 mode 并拒绝非法值。

typed Table getter 从原生 values 读取，忽略 stale caption/cache；setter 只替换自身列，保留其他列 native criteria 和 absent/deleted values。同步遇到非 values native filter 保留原配置，不把 restriction 误写成 all。

`compileSheetCellFilterCriteria(criteria)` 复用原生 cell/error 读取和 condition matcher。`readSlicerItemState` 对 native condition 用真实 own-column cells 派生 checked/partial；`readSlicerSelection` 返回该显示状态的 mirror，`getInfo` 与 Canvas 投影不会因 typed getter 的 restriction 拒绝而丢掉控件。partial Canvas 操作从真实 native predicate 投影既有 raw membership，再作为用户显式编辑替换为 native values；读操作不转换 native condition。

当前 `readSlicerRecordSelection` 对 native condition 明确拒绝无法表达的 raw selection；正式公开 restriction marker 由主任务接入。既有引擎尚不支持真正的 color/manual criterion；未知 kind 保留，UI matcher 不支持时仍拒绝，不能标为 color/manual 完成。

## 未运行的合同草稿

[pivot-filter-fixture.mjs](../../packages/office-studio/tests/fixtures/pivot-filter-fixture.mjs) 包含总金额 280 的 7 条真实稀疏 records、两份普通报告、一个 linked、一个 owned、两个共享 views；Table fixture 使用同一原始数据和 native autoFilter。numeric `1`、boolean `true`、null 合计 90，另选 beta 合计 160；caption `"1"` 全选为 30，caption `"true"` 全选为 70。

| 草稿 | test 声明 | 展开计划用例 | 状态 |
| --- | ---: | ---: | --- |
| [slicer-record-selection.contract.test.mjs](../../packages/office-studio/tests/slicer-record-selection.contract.test.mjs) | 20 | 28 | 未运行 |
| [slicer-table-record-selection.contract.test.mjs](../../packages/office-studio/tests/slicer-table-record-selection.contract.test.mjs) | 11 | 13 | 未运行 |
| [slicer-record-selection-facade.contract.test.mjs](../../packages/office-studio/tests/slicer-record-selection-facade.contract.test.mjs) | 6 | 13 | 未运行 |
| 本记录合计 | 37 | 54 | 全部未运行 |

模型草稿含 collision identities、partial/exclude、checked-all、连接、来源 prune/new members、legacy、no-op、conflicting raw edits、worksheet output 冲突原子性与 JSON 重开。Table 草稿含 native authority、own-column replacement、condition partial UI、其他列保留、exclude parser fidelity。Facade 草稿经现有命令验证计划包括单次 history、undo/redo、cancel、linked/owned chart 与 shared view 通知、零连接 save/reopen；没有执行这些流程。

另见[日期/多类别合同记录](SHEET-23-24-pivot-filter-categories-and-typed-members.md)：19 条声明、86 个计划用例。两份记录合并为 **56 条声明、140 个计划用例，全部未运行**，不与前一批 source lifecycle 重复计数。

## 缺口与验收边界

1. 正式 async bounded typed getItems/getSelection、RecordValue 日期/error authority 和公开 native restriction marker 尚未完整接入；此记录不把内部同步 API 视为正式公开 API 完成。
2. Table color/manual、隐藏表头、无数据 range、完整 Office typed cache identity 尚未验收。condition 派生只基于既有原生引擎支持的 criteria。
3. XLSX cache/filter/slicer parts、Office 导出/保存/重开、独立 Office 样本、固定上游运行行为仍缺。JSON 合同不能代替 Office 文件保真。
4. 当前路径在主线程 TypeScript；Worker/Rust/WASM、取消、TypedArray lease、复制次数/字节与真实性能/内存基准尚未完成。[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)仍适用，不声称优化、零拷贝或达标。
5. 本批核心未新增可替换 UI 控件；Canvas 输入/覆盖层的 `antdv-next`、ConfigProvider 主题、焦点与 IME 需要浏览器验收。所有测试和浏览器行为保持“未验收”。

以上均保留在完整 Sheets / Docs / Pro 交付范围内。
