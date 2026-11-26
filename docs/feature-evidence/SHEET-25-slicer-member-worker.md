# SHEET-25：成员窗口、Canvas caption 与 Date metadata 的 Worker/Rust/WASM 路径

记录日期：2026-11-26。状态：**实现、合同/浏览器/binding 草稿和基准入口已落盘，未验收**。

本轮仅静态读取与编辑；未执行测试、类型检查、构建、Cargo/Rust/WASM、dev、服务、浏览器、性能基准或 `git diff --check`。没有通过数、实测耗时、内存峰值或优化结论。此记录补充[共享缓存选项](SHEET-25-slicer-cache-options.md)，不将计划用例数作为功能完成证明。

## 固定来源与可观察契约

固定 OSS：`E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。以下发布类型/ES 仅作静态研究，没有安装或执行商业 SDK。

| 来源 | 本地保留的契约、差异与未核实项 |
| --- | --- |
| [FSheetSlicer.getItems](../../.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/facade/f-slicer.d.ts)，约 438–455 行 | 防御性 raw typed window，默认 start=0/count=100，业务 cap 默认 200，真实 total；取消、source invalidation、unit/plugin disposal 拒绝；读取不改 selection/startItem。正式 getItems 没有加入 label/selected/partial 字段 |
| [SlicerMemberIndex.build/getWindow](../../.research.local/univer-1.6.0/slicer/lib/types/member-index.d.ts) | 建完整成员索引后按 availability/hide/dataAtTop 返回 bounded window；本地 typed 去重、最终排序与投影由 Rust 实现 |
| [cache enums/options](../../.research.local/univer-1.6.0/slicer/lib/types/definition.d.ts) | 三种 crossFilter 和四项共享 options；沿用已有 cache author/command/history/save 协议，没有另建过滤引擎 |
| `.research.local/univer-1.6.0/sheets-slicer/lib/es/plugin-Bf1aigaL.js` 的 `_comparator/_calendar/_getAvailability/_availableKeys`；具体研究结论见[前批记录](SHEET-25-slicer-cache-options.md) | typed 次序 blank/number/date/string/boolean/error；显式 customListSort 生成长/短月份与星期，hide 优先于 AtTop，none 全部 hasData=true。普通文字继续本地 Intl.Collator numeric/base；上游全部 locale flags 尚未运行核实 |
| [原 Canvas layout](../../packages/office-studio/src/features/slicer/slicer-renderer.ts) 和 [caption item state](../../packages/office-studio/src/features/slicer/slicer-record-selection.ts) | 原工作台按 raw caption 合并同名 typed 项，hasData OR；checked/partial 来自所有源行。Canvas 模式保留这些本地 UI 契约，包括隐藏行参与 partial 和 Date ISO 原 caption；不改变正式 raw API |

本批没有新增顶部常驻、扩展、溢出或上下文入口；仍参考[顶部来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)。UI 和 Pro 能力继续按完整项目范围交付。

## 模块与三种查询

按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)与[WASM 平台契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)，计算属于 `office-compute`，七个 TypedArray 经既有协议 2 计算队列转移到真实 Worker；模型读取和小窗口投影属于 feature adapter，Vue/Pinia/Ant 不承担算法。

- `SlicerMemberQueryService.getItems`：真实 Worker 中产生 typed total/window，仅返回 `{ key, value: RecordValue, hasData }`。保留 defaults/cap/defensive value；输出不会改作者 selection、options 或 startItem。
- `getDisplayMemberWindow`：本地 Canvas API，返回 `{ total, items: { key, value, label, hasData, selected, partial }[] }`。Rust 先按 typed identity 去重、排序及 hide/AtTop，再在完整投影上折叠同 caption，最后取窗口。total 为 caption 数；同 caption 不会跨窗口重复。selected 表示全选，partial 表示混选，两者互斥。
- `getSourceMetadata(workbook, slicer, context, signal?)`：只扫描 typed Date，返回真实分组/来源信息和 fractional serial bounds；cached Date 有效，普通 number、ISO text、numeric text、cached error 不推测为日期。第四参数复用同一准备/等待/发布取消链。

正式 items 和本地 Canvas 共用 `queryMemberWindow`。返回的 first-occurrence source offsets 仅在内部使用；host 对最多 200 个代表行读取原 cell，恢复 typed value 与 `String(cell?.value ?? '')` caption，没有为 label/状态再扫全源。Date 的 ISO 原值因此不会被 serial caption 替换。

Canvas 全选/半选在完整源上按原 caption 聚合，独立于 typed 代表行和 hide 投影。同 typed Date 的 ISO/serial 原 caption 可以不同；部分行的 native predicate 结果也可以不同。先聚合源 caption 的 selected/unselected，才能保持原行为。显示 hasData 则由投影后的 typed 成员合并。

Root 正在接 Canvas 异步 view cache/controller、Facade 本地入口和 renderer 消费；完成状态需以 Root 的后续证据和正式对齐清单为准。本记录不把 controller、Ant 覆盖层、焦点、IME 或实际画布行为标为已验收。既有同步 `getSlicerDisplayMembers` 仍保留，亦用于明确的 JS 基准。

## 协议、字典与权威算法

`kind='slicerMembers'`，`params.mode='items'|'canvas'|'metadata'`。`params` 明确 rowCount、dictionaryCount、textByteLength、start/count、locale 和四项 cache options。

| buffer | items/metadata | Canvas |
| --- | --- | --- |
| `values: Float64Array` | 每源行数值、Date serial、boolean 0/1；其余为 0 | 同一 typed 语义 |
| `states: Uint8Array` | blank0/number1/date2/string3/boolean4/error5 | 同一 typed 语义 |
| `textIds: Uint32Array` | rowCount，String/Error 的 raw dictionary ID；其余 UINT32_MAX | 2×rowCount：前段 raw IDs，后段每行 source-caption IDs；raw error 和显示 caption 可以不同 |
| `availability: Uint8Array` | 每源行 0/1 | bit0=hasData，bit1=selected；不存在两个冲突的来源状态引擎 |
| `textOffsets/textBytes/textRanks` | exact UTF-16LE 字典、Worker Intl ranks | raw text/captions 共用同一 interned 字典 |

UTF-16LE 保留 empty string、嵌入 null、emoji 和不同 lone surrogate，不通过 TextEncoder 替换造成 identity 合并。Rust 验证 dictionary exact interning、偏移、rank 和 typed row 数据；locale/calendar 相等 rank 仍保留不同 raw identities，稳定保留 source tie 顺序，降序不倒置相等项。

Worker 只承担 locale 边界：共享 Intl.Collator 与本地 calendar comparator，分块解码、稳定 merge 生成 rank。Rust 是 typed dedup、重复 availability OR、最终稳定排序、hide/AtTop、caption fold/tri-state、真实 total/window 和 typed Date bounds 的单一权威。none 忽略 hide 且 hasData=true；hide 优先于 AtTop。

两个结果 buffer 是 requested `indices:Uint32Array` 和 `hasData:Uint8Array`。items 输出 hasData 仍为 0/1；Canvas 输出 bit0=hasData、bit1=fully selected、bit2=partial，允许 0–5，拒绝 contradictory full+partial。metadata 返回真实 Date count/bounds、空数组、total=0。缺 Date 的 min/max 序列化为 null，桥接实际使用 `serialize_missing_as_null(true)`。

加载器要求真实 `create_slicer_members`，缺接口的旧产物明确拒绝；没有生产 JS fallback。需要后续统一重新构建 `public/compute`，本轮未构建或改写产物。

## 所有权、预算、取消与平台调用

主线程只准备源 typed columns、literal text/caption 字典与既有 availability/selection 谓词。每查询默认 256 MiB、服务最多 32 在途查询；准备期间累计所有 owner 的预留峰值，超预算拒绝。执行队列继续使用已有共享 `OfficeComputeClient`，无需为 task 单独创建新 worker pool。

两端预算公式相同，`r=rows,d=dictionaryCount,t=textByteLength`：

```text
I = (canvas ? 18 : 14)*r + 8*d + 4 + t
O = metadata ? 0 : 5*min(count,r)
estimatedPeak = 2*I + 128*r + 160*d + 2*t + 2*O + 8 MiB
```

这是 admission estimate，排除原模型 heap，不能报告为实测峰值。JS→Worker 转移七个独占 buffer；Rust copy-in 每个非空 buffer，最多七次、字节数为 I；输入 Vec 移入 owner，不 clone bulk。只复制有效 requested 输出，最多两次、实际字节数为 windowLength×5。空输出依然转移两个独立、可读的零长度 buffer；host 使用 caught zero-length constructor probe 拒绝 detached empty view。

Rust 复用 [bridge DTO](../../crates/office-wasm-bridge/src/dto.rs)、[buffer/lease](../../crates/office-wasm-bridge/src/buffers.rs) 和 [scheduler](../../crates/office-wasm-bridge/src/scheduler.rs)：`serde-wasm-bindgen` 实际转换 control DTO；`js-sys` 实际 copy/view/TypedArray；`compute_yield()` 实际进入共享 `wasm-bindgen-futures` Future/Promise 和 `web-sys` Worker/browser 宏任务让出。输出 helper 在 JS 分配后取得同步不可重入 lease，视图不跨 await；每次按当前 memory 重建，禁止转移整个 linear memory。

查询默认总 deadline 30 秒，包括源准备/排队/初始化/等待；准备和输出发布复查 unit、source revision、permission owner。等待期间有 owner-bound 50ms 复查，失效 abort 同一 signal。interval、timeout、abort listener、memory reservation 均在 finally 释放。Worker rank 和 Rust step 使用真实让出；candidate/result 均 finally free。较大的构造/copy/字典 hash 尚是单次调用，依赖现有 cancellation grace 后 kill/recreate 保证可取消，不宣称每个分配都可分块中断。

注入 Studio 的 computeClient 时 service dispose 只取消自己持有的 jobs，不 dispose 共享 client；自建 client 随 service 释放。返回结果校验 request/unit/revision/kind、维度、copy/transfer 统计和 owned window，过期或晚到结果不能发布。

## 草稿清点与真实验证边界

| 新增文件 | 静态声明 | 计划用例 | 覆盖范围；全部未运行 |
| --- | ---: | ---: | --- |
| [compute contract](../../packages/office-studio/tests/slicer-member-compute.contract.test.mjs) | 16 | 18 | packing、exact UTF-16、locale ranks/yield、shape/budget、transfer/detached/alias、owned output、context/copy stats、old artifact 拒绝、client cancel/stale/kill/reuse、Canvas raw/caption 分离、clone/短列/mode 错配与 packed tri-state |
| [native Rust](../../crates/office-compute/src/slicer/tests.rs) | 21 | 21 | 固定 typed representative expectations、OR、stable sort/ties、六种 options、total/tail/Date、malformed/failure/no partial output、steps/progress/consumption、Canvas 全局 fold/隐藏 partial/ISO Date/独立 error caption |
| [wasm-bindgen-test](../../crates/office-compute/tests/wasm_slicer.rs) | 8 | 8 | 专用 Worker 的真实 wasm32 binding、source copy、防 alias、memory.grow、独立输出 transfer/detach、control/budget、Future yield/drop/reuse、metadata/empty/null、Canvas packed output |
| [real Worker browser](../../packages/office-studio/tests/slicer-member-worker.browser.test.mjs) | 12 | 40 | 真实 OfficeComputeClient/Worker/WASM、fixed window、六种 options、3 locale×2 order、Table/Pivot predicates、cached Date metadata、empty 三模式、items/Canvas/metadata cancel、revision/owner/close/dispose、reuse、output release、全局 caption window 及六种 options 的 caption 顺序 |

本批新增独立文件共 **57 条声明 / 87 个计划用例，全部未运行**。浏览器 text case 内还有 customListSort true/false 子场景，Canvas case 内还有六种 options 子场景，均未另外重计。client contract 使用 port double 验证消息调度；source/Facade 原草稿用 [test-only reference client](../../packages/office-studio/tests/fixtures/slicer-member-compute-fixture.mjs)，明确没有执行 WASM。浏览器与 binding 另走真实产物，不能用 double 通过代替。

浏览器 late-result case 是将之前实际 Worker 返回的包故意重放给当前 handler，验证晚包 delivery guard；当前计算仍来自真实 WASM。source/permission 场景通过真实在途查询的 owner guard 失效，不证明整套应用权限 UI 已验收。

既有 [query contract](../../packages/office-studio/tests/slicer-member-query.contract.test.mjs) 当前 13 条声明 / 17 个计划用例，新增 bounded raw-caption lookup、全局 Canvas tri-state、metadata signal；旧 source/metadata/lifecycle 草稿亦迁移显式 test client/baseRevision。本文件总数单列，不与前批叠加。另改旧 cache-options source 草稿和 protocol-v2 mock module 的必需导出，未把已有声明计入上述 57/87。

## 可复现基准草稿

新增 [browser benchmark](../../tests/benchmarks/slicer-member-benchmark.mjs)、[HTML](../../tests/benchmarks/slicer-member-benchmark.html)、[CLI](../../scripts/benchmarks/slicer-member-benchmark.mjs)，共用 [deterministic dataset](../../packages/office-studio/tests/fixtures/slicer-member-worker-data.mjs)。版本 `typed-members-v1-lcg-1664525-1013904223`，seed `0x511ce2`；输出记录 row count、seed、版本与 exact UTF-16 checksum。输入包含六种 raw 类型、重复/空串/calendar aliases/lone surrogate/原生 Region availability。

后续统一验证时可执行（本轮未执行）：

```powershell
pnpm exec node scripts/benchmarks/slicer-member-benchmark.mjs --sizes 4096,16384,65536 --iterations 3 --locale en-US --channel msedge
```

items、Canvas、metadata 各用相同不变 Table，比较 main-thread JS baseline、每次新 Worker、warmup 后复用 Worker。记录 wall/query/source preparation、queue/compute、exact packed/copy/transfer bytes、frame intervals、主线程长任务、可用时 main JS heap/UA-specific memory、取消 rejection/真实 acknowledgement/recovery。默认 81 个 ordinary sample，warmup/cancellation 单列；没有执行或生成结果 JSON。

fresh Worker 不是保证磁盘冷启动。JS heap 和 admission estimate 不是 Worker resident/WASM peak；不支持的内存指标明确为 null。此夹具未挂载 Office Canvas，首屏、滚动、编辑、缩放、真实大文件/持续内存均保持“未验收”。

## 文件归属与剩余缺口

本批共 **36 个文件**（含本证据和旧 cache-options 证据补录），未修改 Root 的模型/Facade/Studio 装配/renderer 或 UI 子任务文件：

- 计算新增：`compute/slicer-member-{protocol,tasks,results}.ts`、`compute/slicer-member.worker.ts`；修改 `compute/{protocol,tasks,guards,compute.worker,wasm.worker,index}.ts`。
- feature 新增：`features/slicer/slicer-member-{order,input,transport}.ts`；修改 `slicer-member-query.ts`、`slicer-cache-options.ts`。
- 平台/Rust：`platform/wasm/task-protocol.ts`、`office-wasm-bridge/src/control.rs`、`office-compute/src/lib.rs`，新增 `src/slicer.rs`、`src/slicer_wasm.rs`、`src/slicer/tests.rs`、`tests/wasm_slicer.rs`。
- 前端新增：`tests/slicer-member-compute.contract.test.mjs`、`tests/slicer-member-worker.browser.test.mjs`、`tests/fixtures/slicer-member-{compute-fixture,worker-data,worker-runtime}.mjs`、`tests/fixtures/slicer-member-worker.html`；修改 `slicer-member-query.contract.test.mjs`、`slicer-cache-options.contract.test.mjs`、`wasm-protocol-v2.contract.test.mjs`。
- 基准新增：`tests/benchmarks/slicer-member-benchmark.{html,mjs}`、`scripts/benchmarks/slicer-member-benchmark.mjs`；证据新增本文件，并补录 `SHEET-25-slicer-cache-options.md`。

仍需统一类型/前端/Rust/binding/浏览器验证、重建产物，以及同一数据集真实大样本测量；在此之前保持未验收。Table/Pivot availability、native/retained selection 编译与源模型读取仍在主线程，Pivot 的既有 availability 可能同步重扫多个报告；没有将这些工作宣称为 WASM 化。旧 synchronous member helper 的其他消费者也需由 Root 完成核对。

本批没有新增生产 UI 可替换控件；Canvas Ant 输入/覆盖层、ConfigProvider 明暗 token、IME/焦点、Office 原生 slicer/cache parts、文件保真往返、协作权限/history 闭环及全部 Pro 最终功能仍按原完整范围验收，不因当前未验证而缩减。
