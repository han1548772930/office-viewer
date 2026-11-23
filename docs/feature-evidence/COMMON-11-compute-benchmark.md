# COMMON-11 compute benchmark evidence

This record documents a reproducible measurement harness for the numeric sort/filter path. It is evidence for the Worker/WASM boundary and result parity; it does not claim that a performance target has been met.

Run it from the repository root with:

```powershell
node scripts/benchmarks/compute-benchmark.mjs --sizes 4096,16384 --iterations 2 --output docs/feature-evidence/benchmarks
```

The runner starts a Vite page, loads the built `public/compute/office_compute.js` module through the production `OfficeComputeClient`, runs the browser harness in Chrome, and writes a JSON record plus a Markdown table. The smoke contract is:

```powershell
pnpm test:benchmark:smoke
```

The smoke test uses 128 and 512 rows and checks that TypeScript and Worker/WASM results have identical indices, that the route remains TypeScript below the configured threshold, and that JSON/Markdown output includes the limitations.

The recorded run on 2026-11-20 used Windows 10 x64, headless Google Chrome 157, 16 logical processors, reported 32 GiB device memory, a 1440×900 viewport, and DPR 1. It used two numeric key columns generated deterministically from row index, with nulls at fixed intervals. Sort used ascending then descending keys with nulls last. Filter used `column 0 between -1000 and 1000`, `column 1 >= 0`, and `match=all`. Every measured sample reported `parity: true`.

The run produced these route observations:

| Rows | Default sort route | Default filter route | TypeScript sort P50 wall | New Worker sort P50 wall | Reused Worker sort P50 wall | New Worker filter P50 wall | Reused Worker filter P50 wall |
| ---: | --- | --- | ---: | ---: | ---: | ---: | ---: |
| 4096 | Worker/WASM | Worker/WASM | 0.5 ms | 22.0 ms | 0.8 ms | 22.2 ms | 0.5 ms |
| 16384 | Worker/WASM | Worker/WASM | 1.9 ms | 32.9 ms | 11.7 ms | 22.3 ms | 1.0 ms |
| 262144 | Worker/WASM | Worker/WASM | 48.4 ms | 2160.8 ms | 1263.4 ms | 677.0 ms | 326.2 ms |

The same harness also records a 64-row run where the default route remains TypeScript. The 262144-row record is useful for exposing packing and transfer costs: the sort sample had a 2.16 s wall time while the Worker client total was about 1.57 s, so kernel time alone is not an end-to-end optimization claim. These are single Chrome/device samples and are not enough to calibrate the threshold. The configured `minRows: 4096` therefore remains a provisional routing value and must stay exposed as a calibration parameter until the same workloads are measured across supported browsers, devices, data densities, and realistic workbook operations.

Each sample includes:

- wall-clock operation time;
- Worker kernel time, client queue/startup time, and client total time;
- exact TypedArray input/result and boundary byte counts;
- the client’s schema-based `estimatedPeakBytes` value;
- Chrome `usedJSHeapSize` before/after when available;
- `PerformanceObserver` long-task count, total, and maximum duration;
- a parity flag against an independent TypeScript stable numeric implementation.

Memory fields have explicit limits: `estimatedPeakBytes` is not resident WASM memory, JS heap deltas include garbage collection and browser scheduling noise, and JavaScript array storage is not inferred from the TypedArray byte count. The generated JSON files under `docs/feature-evidence/benchmarks/` are the authoritative raw records for each run.

## Protocol 2 and chart statistics implementation, 2026-11-24

Status: implementation in progress, **not accepted**. The numeric measurements above predate the rewritten bindings and do not validate the current artifacts. The user requested implementation first and unified testing later. This batch has not run tests, type checks, builds, Rust/WASM checks, browsers or benchmarks, and has not restarted the dev server. The files under `public/` still require regeneration before unified validation; Workers intentionally reject old WASM interfaces instead of falling back.

`create_chart_statistics` and the `chartStatistics` Worker domain now use the shared protocol 2 context, typed numeric/state/category buffers, authored job descriptors, explicit result ownership and copy/transfer metrics. Histogram, boxplot, Pareto, waterfall, trendline and error-bar calculations have native implementations and matching TypeScript planning/decoding. The loader requires the chart constructor as part of the current interface. Result validation checks field/slot/job identity, contiguous output ranges, dimensions, independent complete output buffers and consistent byte/count metrics; it does not claim numerical parity from descriptor validation alone.

`ChartStatisticsService` serializes a chart's complete preparation, Worker wait and decoding lifecycle, limits pending requests/cache estimates, removes queued source references on cancellation and prioritizes queued interactive work. An interactive view can promote an already queued background request. Each consumer has its own deadline covering preparation and queueing; the original consumer's timeout does not cancel another view that is still waiting. The shared Worker task is aborted when the last consumer leaves, with the compute client's cooperative cancellation and terminate/recreate fallback. Version/configuration invalidation and result-delivery observer reentry reject stale snapshots. Other calculation domains continue to share the bounded compute client. Cache/resident memory remains an admission estimate, and callers may retain an evicted snapshot outside the service's cache accounting.

The shared WASM input adapter distinguishes a legitimate empty buffer from a detached buffer through a caught zero-length Uint8Array constructor, before allocating or copying data. It rejects the current WASM linear-memory buffer. Independent output helpers copy only the requested slice before the owner is released; their temporary WASM view does not escape the synchronous adapter. These bindings and memory claims still require compilation and actual browser/Worker execution.

Pending validation drafts:

The chart descriptors additionally require `seriesPosition`, the index in the prepared series array. A source Y field may appear in multiple groups with different X/sample columns. Planning now expands every matching group for each authored trendline/error-bar slot; decoding keeps `{ configIndex, seriesPosition }` instead of overwriting a slot with the first group. Series positions must be unique, and job identity is `(kind, seriesIndex, slot)`. Expanded columns and jobs have hard admission limits before allocating bulk buffers. The TypeScript/Rust contracts have changed together and deliberately accept no older descriptor shape.

Progress delivery stops if an observer invalidates the entry and skips consumers detached by an earlier observer. The new group and observer contracts remain unexecuted, and the limits are admission estimates rather than measured resident memory.

| File | Intended coverage |
| --- | --- |
| `crates/office-compute/src/chart/tests.rs` | Native statistical arithmetic, edge values and model/parameter contracts |
| `crates/office-compute/tests/wasm_chart.rs` | Real structured controls, typed input ownership, authored slots, result consumption and memory.grow independence |
| `crates/office-wasm-bridge/tests/worker_platform.rs` | Subviews, precision, copy budgets, empty/detached buffers and linear-memory aliases |
| `packages/office-studio/tests/wasm-protocol-v2.contract.test.mjs` | Required interfaces, nested contexts, chart input/output transfers, malformed descriptors/metrics, cancellation and stale results |
| `packages/office-studio/tests/chart-statistics-compute.contract.test.mjs` | Shared preparation/planning/decoding and TypeScript/native-compatible data contracts |
| `packages/office-studio/tests/chart-statistics-service.contract.test.mjs` | Shared consumers, independent deadlines, queue cancellation/priority, permissions, cache and observer reentry |
| `tests/browser/compute-runtime.ts` and `compute-runtime.spec.ts` | Production client/Worker/WASM chart fixture, actual workbook source preparation/cache/recompute and 800,000-row cancellation/revision fixtures |

The deterministic three-sample binding fixture `[4, 7, 9]` declares 39 input bytes, three inbound copies/transfers, and 88 output bytes for histogram plus boxplot, with two outbound copies/transfers. These numbers are expected contracts in unexecuted drafts, **not new measured evidence**. Empty outputs still own separate transferable buffers and record zero payload copies. The large fixture is a cancellation/validity scenario, not a performance acceptance run.

Remaining measurement work includes realistic chart projection/aggregation and pivot/visibility costs on the main thread, source packing/decoding, end-to-end queue time, Canvas/WebGPU uploads and scrolling/editing/zoom frame times, long tasks, and actual JS/WASM/GPU memory. Histogram/boxplot job stepping still performs a whole statistical job synchronously; terminate/recreate remains necessary for a job that exceeds the cooperative cancellation grace period. The same datasets must be compared before/after on recorded devices/browsers/backends before any optimization or large-file acceptance claim.
