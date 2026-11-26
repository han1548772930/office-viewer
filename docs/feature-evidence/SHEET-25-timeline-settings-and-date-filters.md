# SHEET-25：正式 Timeline 设置、半开日期范围与固定条件

记录日期：2026-11-26。状态：**领域与服务、合同草稿已落盘，未验收**。

本批按固定公开 Timeline 契约拆分独立 view settings 和共享 original-date-field filter。范围保留作者 Excel serial `[startDate, endDate)`，固定条件保留八种日期谓词，不转换为已枚举成员列表。按当前“先实现，后统一测试”要求，只静态读取与编辑；没有运行测试、类型检查、构建、Cargo/Rust/WASM、dev、服务、浏览器、基准或 `git diff --check`。没有通过数、Office 重开结果或性能达标结论。

## 固定公开来源与差异

固定 OSS 为 `E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。以下 1.6.0 发布包仅作静态研究材料，没有安装或执行商业 SDK。

| 精确来源 | 已核对契约与本地差异 |
| --- | --- |
| `.research.local/univer-1.6.0/slicer/lib/types/definition.d.ts`，`TimelineLevel`、`ITimelineRange`、`TimelineDateFilterOperator`、`ITimelineDateCondition`、`ITimelineCacheDefinition`、`ITimelineViewConfig` | owner levels 为 Year=0、Quarter=1、Month=2、Day=3；range 有 startDate/endDate/useWholeDay；condition 为 between/notBetween/equal/notEqual/newerThan/newerThanOrEqual/olderThan/olderThanOrEqual；cache selection 与 condition 互斥；level、selectionLevel、scrollPosition 属 view。模型继续保存 year/quarter/month/day 字符串，通过正式 DTO 显式转换为 numeric levels |
| `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/types.d.ts`，`SheetTimelineSettings`、SetDateRange、SetDateCondition、PatchView intents | settings 独立于 shared filter/source/connections；内部允许 SetDateCondition intent。本地没有凭空声称固定 Facade 已公开直接 setDateCondition，内部服务供 Pivot 命令同步使用 |
| `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/facade/f-timeline.d.ts`，`getDateRange/getDateCondition/setDateRange`（约 404–443 行） | 两个 getter 返回 retained shared defensive values，source loss 后仍可读；missing hosted view/cache 为 undefined。setter finite start<end，undefined 清筛选，不改独立 levels |
| 同 Facade，`setLevel/setSelectionLevel/setScrollPosition`（约 462–501 行）与 `updateSettings` | 显示/选择粒度互相独立；scrollPosition 是有限 Workbook Excel serial，允许负数和 fractional，不能解释为 px/item index。presence-aware patch 的 own undefined 清 optional overrides |
| 同 Facade，`getDateBounds`（约 444–460 行） | inclusive original-source minDate/maxDate 与 host date system，不能用 end-exclusive selection upper bound 或 viewport 代替。继续由独立真实源 async query 接线，本批没有声称 bounds 浏览器查询已经验收 |
| `.research.local/univer-1.6.0/slicer/lib/es/index.js`，`gt`（公开 compileTimelineRange） | 只匹配 typed Date，start inclusive/end exclusive；whole-day 派生 floor(start)/ceil(end)，不改作者边界。standalone compiler 允许相等原边界并明确匹配空集合，包含 fractional whole-day 相等输入；正式 setter/持久化 selection 仍要求严格非空 |
| 同 ES，`_t/vt`（公开 compileTimelineCondition/operands） | between/notBetween 两端 inclusive；whole-day 对 operands 与 actual 都 floor；其余为准确的等于/不等于/前后比较；只匹配 typed Date |
| 同 ES，`fn/vn/yn/pn`（cache/view 资源解析） | range 与 condition 同时存在拒绝，finite operands、typed numeric levels、finite scrollPosition、optional boolean 校验。本地服务额外拒绝 DTO 未知 keys；原始解析扩展字段保留属于 normalizer/IO 责任，不据此删弃源资源 |

本批没有新增顶部入口。常驻、扩展、溢出、上下文入口继续见[完整来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)、[切片器连接 UI](SHEET-25-slicer-connections-ui.md)与[既有 Timeline/仿射几何](SHEET-25-slicer-timeline-affine.md)。控件数与合同数不代表功能完成。

## 模块、权威数据与原子边界

按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，本批位于 `domain` 与 `features/slicer`；没有 UI、DOM、Vue 或 WASM 实例依赖。重计算仍由既有 Pivot 引擎完成，没有新增第二套聚合或源成员筛选权威。

- [timeline-filters.ts](../../packages/office-studio/src/domain/timeline-filters.ts)：正式 numeric levels、range、condition、互斥 filter state DTO。
- [timeline-date-filter.ts](../../packages/office-studio/src/features/slicer/timeline-date-filter.ts)：normalize/compile、effective whole-day bounds、numeric/model level 边界、Pivot operator 转换。只引用域类型与 typed RecordValue helper，**不 runtime import pivot-engine**；Pivot 同时复用 numeric condition compiler，避免依赖循环与冲突比较算法。
- [timeline-settings.ts](../../packages/office-studio/src/features/slicer/timeline-settings.ts)：只允许十个 view keys；required name/level/selectionLevel 不能清除，七个 optional own undefined 清 override；名称跨 Slicer/Timeline 唯一，caption≤32767，style≤255 且存在于 Timeline catalog，scroll finite。一次有效 patch 只增加当前 view revision；no-op 返回 false，不更改 shared cache、filters、source、reports、其他 views 或几何。
- [timeline-filter.ts](../../packages/office-studio/src/features/slicer/timeline-filter.ts)：retained getters 校验 hosted identity/cache ownership，不验证当前 source；live setter 校验 source 与真实 typed Date。范围和 fixed condition canonical 互斥；clear 遍历同源重复 label dimensions 的 date/range/condition，保留不相关 caption/value 条件。共享修改同步 Timeline views 的 canonical mirrors/revisions，不改变各自 level、selectionLevel、scrollPosition。

主任务在 [model.ts](../../packages/office-studio/src/domain/model.ts) 添加 cache/view 的 `timelineDateRange/timelineDateCondition`；PivotField 也保留 canonical range/condition。fixed condition 的 `dateFilter` 为正式 Pivot API 镜像，canonical marker 采用 typed Date source predicate；传统 Pivot date setter 清 marker，从而保留既有传统 ISO 值行为。主任务在 [pivot-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-engine.ts) / [pivot-filter-predicate.ts](../../packages/office-studio/src/features/pivot/services/pivot-filter-predicate.ts) 接入 condition category、清除/移入 Values、typed source serial 与 `setPivotTimelineDateRange/setPivotTimelineDateCondition`，并保留 optional `useWholeDay` 的 omitted/false 差别。

`applyTimelineReportFilter` 为每个候选报告建立 detached clone，以 deferred setters 先改字段；只为实际改变的 worksheet reports 复制 output sheet cell maps。先清所有旧 owned footprints，再渲染每个新 worksheet output、计算所有 owned charts。任何 output conflict 或计算错误都发生在提交之前；成功后才提交报告/输出，再写 cache/view state。多步 legacy retirement 与新日期规则合为一次 report revision，shared cache/view revision 也只增加一次。没有按现存成员计算 epsilon 上界，也没有把 source snapshot 当永久日期权威。

从旧 Timeline `recordSelection/selectedItems/timelineRange` 切换 canonical state 时，服务在报告 clone 中先清旧 member category，避免 allowMultipleFilters=true 把新范围继续截断为旧 include 快照。非迁移路径中，普通手工成员是否与日期条件并存继续服从 report multiple-filter policy。旧无 cache model 应由既有 normalization/create 路径 materialize；正式 retained getters 不重建 missing cache。

Facade/Studio/命令、source/date bounds query、generic Slicer synchronization、重连、保存 normalizer 和 exports 由主任务接线。它们必须保留 canonical range/condition，不能被旧 member-only synchronization 清掉。纯 helper 的 true/false 表示变化/no-op；正式公开 setter 的 boolean 是 command acceptance，普通业务拒绝/权限、Before 取消与异常传播继续由原命令宿主处理。

## 未运行合同与合成样本

[timeline-filter-fixture.mjs](../../packages/office-studio/tests/fixtures/timeline-filter-fixture.mjs) 复用固定真实 Pivot 模型服务；含两个 worksheet reports、一个 owned report、linked chart、两个 shared Timeline views 与不同源字段的普通 Slicer。十一条合成源记录包含 formatted typed Date、cachedDataType=date、同 serial 普通 number、ISO text string、cached Error、invalid date text；两个 date systems 和 grouped original date fields 共用样本。没有执行夹具。

该样本精确 `[first+.25,first+.75)` 的源汇总计划为 90：含 start Date 金额 10 和 cached Date 金额 80，排除 end Date 金额 20；whole-day 或 inclusive fixed between 为 110。普通同值 number 的 1000、ISO text 的 2000 与 Error 的 3000 均不属于 typed Date。未筛选汇总为 6360。以上是静态可复查的计划期望，不是运行结果。

| 草稿 | 静态 test 声明 | 按固定循环展开的计划用例 | 尚未运行的主要场景 |
| --- | ---: | ---: | --- |
| [timeline-date-filter.contract.test.mjs](../../packages/office-studio/tests/timeline-date-filter.contract.test.mjs) | 13 | 102 | numeric levels、非法 DTO、precise/whole-day half-open boundaries、negative fractions、equal empty standalone、typed Date-only、八个 inclusive conditions、三个 optional modes、defensive captured operands、Pivot 映射、Excel compatibility day 60 |
| [timeline-settings.contract.test.mjs](../../packages/office-studio/tests/timeline-settings.contract.test.mjs) | 12 | 62 | 16 个独立 grains 组合、serial navigation、四个 switches、七个 own undefined、required/refusal 原子性、名称/样式、no-op/revisions、source-loss view settings、防御读与 JSON 重开 |
| [timeline-filter.contract.test.mjs](../../packages/office-studio/tests/timeline-filter.contract.test.mjs) | 19 | 91 | 两个 date systems 与 grouping、48 个 fixed condition/mode 组合、typed cached source、新源日期刷新、范围/条件互斥、source loss/missing identities、no-op、零连接/重连、generic/external Pivot synchronization、legacy/manual member policy、source sibling clear、user-cell/new-output/owned-calculation 原子失败、JSON 重开 |
| 合计 | **44** | **255** | **全部未运行，没有通过数** |

计数仅含本批新增三个合同文件，不重复计入前批 Slicer cache options、typed members、style resources、Values position 或主任务后续 Facade/Studio 草稿。

## 仍未验收与剩余边界

1. Facade、命令 acceptance/no-op、权限/Before 取消、事件、撤销重做、queued settings patch、跨 view history 与保存重开必须在统一阶段运行。当前三份草稿以纯领域和 source/report integration 为主，不代替公开宿主和浏览器证明。
2. `getDateBounds` inclusive min/max、异步取消、source/identity/disposal late results、视图 display period windows、selectionLevel 到范围的实际拖动/键盘行为尚未运行核实。固定 predicates 与 serial navigation 数据契约不能证明完整 Timeline 绘制交互。
3. 本批没有新增或替换控件；正式 settings/Canvas overlay 的 `antdv-next`、ConfigProvider 明暗 token、focus recovery 与 IME 由主任务集成后进行真实浏览器验证。没有新增 native form controls。
4. XLSX 原生 timeline/cache/style parts、原生 Office serial/date-system 对照、独立 Office 文件重开与 parser/writer 映射仍未验收；JSON 计划不等于 Office 互操作完成。旧 legacy range 的迁移需保留原作者来源和记录差异，不能伪造精确半开边界。
5. 当前源 typed validation、报告 staging/渲染仍走主线程 TypeScript；缓存/报告/output-map 复制、Worker/WASM protocol、TypedArray lease、取消和真实大样本内存/性能未测量。[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)继续适用，不声称已优化、零拷贝或达标。

剩余事项继续属于完整 Sheets / Docs / Pro 目标；本批只交付上述真实领域/服务改动与未运行草稿。
