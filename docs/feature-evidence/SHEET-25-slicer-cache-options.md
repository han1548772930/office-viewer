# SHEET-25：共享切片器 cache options

记录日期：2026-11-26。状态：**核心、命令合同草稿和证据已落盘，未验收**。

本批补 `sortOrder/customListSort/crossFilter/hideItemsWithNoData`，选项由共享 cache 持久化，成员排序与可用性显示复用真实 raw source/query 契约。按用户要求只做静态读取与编辑；未运行测试、类型检查、构建、Cargo/Rust/WASM、dev、服务、浏览器、性能基准或 `git diff --check`。下文声明和计划用例均未运行，没有通过数。

## 固定公开来源

固定 OSS：`E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。公开发布包均为 1.6.0 的静态研究资料，没有安装或执行商业 SDK。

| 精确研究文件 | 可观察契约及实现差异 |
| --- | --- |
| `.research.local/univer-1.6.0/slicer/lib/types/definition.d.ts` 第 23–30、46–75 行 | ascending/descending；crossFilter 为 none、showItemsWithDataAtTop、showItemsWithNoData；cache 可选 customListSort，hideItemsWithNoData 属共享 cache。本地统一放入 `SlicerCacheModel.options`，不复刻 data/tableSlicerCache 内部对象层级 |
| `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/types.d.ts` 的 `SheetSlicerCacheOptions`、PatchCacheOptions intent | 四项通过共享 patch 更新，locked source、selection、connections 保留。本地采用同一 presence-aware 命令参数 |
| `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/facade/f-slicer.d.ts` 第 453–477 行 | `getCacheOptions()` 返回防御性作者状态；`setCacheOptions(patch)` 在执行时 patch，own undefined 清可选值，影响所有 cache views |
| `.research.local/univer-1.6.0/sheets-slicer/lib/es/plugin-Bf1aigaL.js` 的 PatchCacheOptions 分支 | 只允许四个 keys，按 presence patch；不覆盖来源、筛选或连接。本地核心只修改作者 options 和 cache/shared view revisions |
| 同 ES 的 `_comparator`、`_calendar` | 排序类型次序 blank/number/date/string/boolean/error，数值按数值；仅 customListSort===true 生成 locale 长/短月份和星期列表。月份排在星期前，未知项随后。本地普通文本使用 Intl.Collator numeric/base；固定 `createPivotSortCompare` 的全部 locale flags 没有单独运行核实，列为差异/未核实 |
| 同 ES 的 `_getAvailability`、`_availableKeys` | none 不应用 availability，hasData 为 true；缺省 crossFilter 仍计算 availability，不置顶；AtTop 置顶有数据项；hide 按可用项过滤。Table 忽略自身列 filter，Pivot 对所有连接报告的可用 records 取 union；零连接为 all |
| `.research.local/univer-1.6.0/slicer/lib/types/member-index.d.ts`、对应 ES `getWindow` | 先整体排序、availability hide/稳定分区，再 bounded slice；total 为投影后的真实总数。hide 优先于 dataAtTop，不改作者 selection |

本批没有新增顶部入口；分类仍见[顶部来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)、[切片器连接 UI 记录](SHEET-25-slicer-connections-ui.md)与[时间线记录](SHEET-25-slicer-timeline-affine.md)。四个选项或合同数不作为功能完成证明。

## 模块与实际核心

按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，[model.ts](../../packages/office-studio/src/domain/model.ts) 新增 `SlicerCacheSortOrder`、`SlicerCacheCrossFilter`、`SheetSlicerCacheOptions` 与 `SlicerCacheModel.options`。作者值保持可选，不把派生默认值写回作者状态。

新 [slicer-cache-options.ts](../../packages/office-studio/src/features/slicer/slicer-cache-options.ts) 提供：

- `normalizeSlicerCacheOptionsPatch`：只允许四项，拒绝非法 enum、null、非 boolean 和未知/locked keys。own undefined 保留在 patch 中，执行时删除可选作者字段。
- `readSlicerCacheOptions`：返回防御性 authored snapshot；Timeline 明确拒绝此 Slicer 排序契约。
- `patchSlicerCacheOptions`：true 表示改变，false 表示内容 no-op。只修改 cache options 及该 cache 的每个 view revision，保留 source、raw selection、native filter、连接、报告、图表和位置。最后一项清除后移除可选 options 对象。
- `effectiveSlicerCacheOptions`：升序、customListSort=false、crossFilter=showItemsWithNoData、hide=false 为派生默认；显式 false 仍是作者状态，可撤销和持久化。
- `createSlicerMemberComparator`：typed kind 与数值排序，显式 customListSort 使用 locale 月份/星期 long/short。等 rank 使用稳定 source 顺序。
- `applySlicerCacheItemOptions`：对完整 `{ key, value: RecordValue, hasData }` 列表生成防御性投影，整体排序，然后 hide 或 AtTop 稳定分区。none 将显示 hasData 置 true，忽略 hide；投影不改原列表或 retained selection。

主任务在现有 [slicer-member-query.ts](../../packages/office-studio/src/features/slicer/slicer-member-query.ts) 接入完整投影后才 total/slice；availability 继续使用原 [sort-filter-engine.ts](../../packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts) own-column exclusion 和 [pivot-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-engine.ts) source row predicates。没有新增第二套 filter authority。

Facade、command、normalizer、Canvas legacy/display items 与 root exports 接线由主任务负责。本记录仅证明核心与草稿已落盘；公开 API、history、取消及 Canvas 最终行为仍待统一验收。更新 cache options 的事务必须在原文档内容历史中完成，Before 取消不能发布部分 cache/view revisions。

## 未运行草稿与来源样本

[slicer-cache-options.contract.test.mjs](../../packages/office-studio/tests/slicer-cache-options.contract.test.mjs) 静态清点 **14 条 test 声明、48 个计划用例**。复用[真实筛选样本](../../packages/office-studio/tests/fixtures/pivot-filter-fixture.mjs)，typed raw 成员含 number/text、boolean/text、blank/empty，普通/owned 多报告及 shared views，Table 原生 filter 使用同一数据。

模型/查询草稿覆盖共享 revision、防御性作者 snapshot、explicit default/own undefined/no-op、非法 patch 原子性、typed kind 排序、calendar aliases、6 种 crossFilter/hide 组合、12 种真实 Pivot/Table availability 与 bounded slicing、own selection 排除、JSON 重开、零连接、不同 canonical cache 隔离、Timeline 拒绝与投影防御性。North 上下文样本的三个有数据 raw 值为 numeric `1`、empty string 和 text `"1"`；typed member 自身选择不使其余可用成员消失。

[slicer-cache-options-facade.contract.test.mjs](../../packages/office-studio/tests/slicer-cache-options-facade.contract.test.mjs) 静态清点 **5 条声明、14 个计划用例**：普通 Pivot/Table 的 author/clear/mixed 取消全部 cache/view revisions 与通知；接受后单次 history、所有 shared views 通知、报告/native filter 保留、no-op、undo/redo/save/reopen；own undefined 参数 presence；取消 options edit 不失效在途 typed window；Before listener 抛错不发布内容。

本批合计 **19 条声明、62 个计划用例，全部未运行**。不与前批 typed selection 或 Values position 合同重复计数。测试尚未执行，不能以这些计划证明功能通过或浏览器可用。

后续[显示设置与 cache options UI 批](SHEET-25-slicer-settings-ui.md)已接真实四项 getter/setter、Ant Select、共享成员投影与实际 owner 提交边界，新增草稿单列 **40 条声明 / 107 个计划场景，全未运行**。不重计本核心批 19/62，也不据此改为已验收。

## 未核实与剩余缺口

1. 固定上游 `createPivotSortCompare` 的全部 locale 行为、locale 切换与中文/其他语言日历名单需要运行核实；当前文本排序 flags 属本地明确实现选择。没有宣称全部 Office 自定义列表管理已经交付。
2. 原生 Table color/manual criteria、无数据/隐藏表头、独立 XLSX、Office 原生 cache/slicer parts 与 options 保存重开尚未验收。现有 native predicate 不支持的 kind 不能伪装为数据全部可用。
3. 选项菜单已在后续 UI 批接真实 setter；Canvas 的 sorted/hidden/hasData 最终投影、`antdv-next`、ConfigProvider 主题、焦点和 IME 尚需浏览器验证；本核心批没有新增可替换控件。
4. 本记录最初批次的 source/member list 与排序位于主线程 TypeScript。后续[成员 Worker/Rust/WASM 批](SHEET-25-slicer-member-worker.md)已落盘正式 typed window、本地 Canvas caption fold/tri-state 和 typed Date metadata 的真实计算路径，取消、leases、copy/transfer 与基准草稿单列；全部未运行、未验收。availability/selection 谓词准备仍有主线程边界，Canvas controller/renderer 接线由 Root 负责，完成状态以 Root 证据为准。[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)仍适用，没有真实性能/峰值内存结果，不声称优化、零拷贝或达标。
5. 公开 API 普通业务拒绝/权限、在途查询 source invalidation、observer 重入等其他统一合同仍由主任务维护；本批取消草稿不表示这些场景已全部验收。

所有剩余事项继续属于完整 Sheets / Docs / Pro 范围。
