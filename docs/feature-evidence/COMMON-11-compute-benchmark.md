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
