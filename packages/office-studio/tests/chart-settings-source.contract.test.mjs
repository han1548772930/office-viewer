import assert from 'node:assert/strict'
import { test } from 'vitest'
import { effectScope, nextTick, ref, shallowRef } from 'vue'
import { createChartSettingsDraft } from '../src/features/charts/chart-settings-draft.ts'
import { CHART_SETTINGS_PREVIEW_MAX_CELLS, chartSettingsBindingKey, chartSettingsDateAxisError, chartSettingsPreparedData, chartSettingsPreparedStatistics, projectChartSettings } from '../src/features/charts/chart-settings-data.ts'
import { chartParetoSettingsCategories } from '../src/features/charts/chart-pareto.ts'
import { CHART_CATEGORY_POINT_CHOICES_LIMIT, chartCategorySettingPoint, chartCategorySettingPoints, chartSeriesSettingPoint, chartSeriesSettingPoints, chartSeriesSettingSample, chartSeriesSettingSampleCount } from '../src/features/charts/chart-series-settings.ts'
import { useChartSettingsSource } from '../src/features/charts/views/use-chart-settings-source.ts'

// P24 drafts only. No test runner, server, browser, type checker or benchmark has run.
// 40,000 samples × two source fields = 80,002 positions, genuinely over the UI preview cap.
const SAMPLE_COUNT = 40_000
function config(patch = {}) {
  return { type: 'boxplot', source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: SAMPLE_COUNT, endCol: 1 } },
    categoryField: 0, valueFields: [1], position: { x: 80, y: 70 }, size: { width: 560, height: 400 }, ...patch }
}
function prepared(field = 1, name = 'Measurements', last = 100) {
  const categories = Array.from({ length: SAMPLE_COUNT }, (_value, index) => `Sample ${index + 1}`)
  const values = Array.from({ length: SAMPLE_COUNT }, (_value, index) => index === 1 ? null : index === SAMPLE_COUNT - 1 ? last : 1)
  const fields = field === 1 ? ['Sample', name] : ['Sample', 'Earlier measurements', name]
  const records = categories.map((label, index) => field === 1 ? [label, values[index]] : [label, 1, values[index]])
  return { fields, records, sourceNumericFields: field === 1 ? [1] : [1, 2], date1904: false,
    categories, series: [{ field, name, values }] }
}
function deferred() {
  let resolve, reject
  const promise = new Promise((yes, no) => { resolve = yes; reject = no })
  return { promise, resolve, reject }
}
async function flush() { await Promise.resolve(); await Promise.resolve(); await nextTick() }
function fixture(options = {}) {
  const active = ref(true), draft = ref(createChartSettingsDraft(config(options.config))), revision = ref('unit-a:0')
  const provided = shallowRef(options.provided), requests = [], statisticsRequests = []
  let cellReads = 0
  const sheet = { id: 'host', name: 'Large source', rowCount: 100_000, columnCount: 16_384,
    get cells() { cellReads++; throw new Error('The settings UI must not synchronously read source cells') } }
  const sheets = shallowRef([sheet]), hostSheetId = ref('host')
  const loader = shallowRef((value, readOptions) => {
    const completion = deferred()
    requests.push({ config: value, signal: readOptions.signal, ...completion })
    return completion.promise
  })
  const statisticsLoader = shallowRef((value, readOptions) => {
    const completion = deferred()
    statisticsRequests.push({ config: value, signal: readOptions.signal, ...completion })
    return completion.promise
  })
  const scope = effectScope()
  const source = scope.run(() => useChartSettingsSource({ active: () => active.value, draft: () => draft.value,
    sheets: () => sheets.value, hostSheetId: () => hostSheetId.value, provided: () => provided.value,
    date1904: () => false, sourceRevision: () => revision.value, load: () => loader.value, loadStatistics: () => statisticsLoader.value }))
  return { active, draft, revision, provided, sheet, sheets, hostSheetId, loader, statisticsLoader, requests, statisticsRequests, source,
    cellReads: () => cellReads, dispose: () => scope.stop() }
}

test('matching prepared source bypasses an over-budget worksheet without reading cells or replacing raw sample arrays', () => {
  const data = prepared(), draft = createChartSettingsDraft(config()), identity = { hostSheetId: 'host', sourceRevision: 'unit-a:0' }
  assert.ok(data.fields.length * (data.records.length + 1) > CHART_SETTINGS_PREVIEW_MAX_CELLS)
  const bound = chartSettingsPreparedData(draft, data, identity)
  const sheet = { id: 'host', name: 'Large', get cells() { throw new Error('No synchronous cell read allowed') } }
  const projection = projectChartSettings(draft, sheet, bound, false, identity)
  assert.equal(projection.error, undefined); assert.equal(projection.prepared, true)
  assert.equal(projection.data, data); assert.equal(projection.source.records, data.records)
  assert.equal(projection.data.series[0].values, data.series[0].values)
  assert.equal(chartSeriesSettingSampleCount(1, projection.data), SAMPLE_COUNT)
  assert.equal(chartSeriesSettingSample(1, SAMPLE_COUNT - 1, projection.data).label, '40000 · Sample 40000 · 100')
  assert.equal(chartSeriesSettingSample(1, 1, projection.data), undefined)
})

test('derived five-number data never supplies raw Boxplot samples even with an otherwise matching identity', () => {
  const draft = createChartSettingsDraft(config()), identity = { hostSheetId: 'host', sourceRevision: 'unit-a:0' }
  const derived = { categories: ['Measurements'], series: [{ field: 1, name: 'Measurements', values: [1, 1, 1, 1, 100] }], date1904: false }
  const input = { ...chartSettingsPreparedData(draft, derived, identity), kind: 'derived' }
  const projection = projectChartSettings(draft, undefined, input, false, identity)
  assert.match(projection.error, /统计结果不能作为原始数据源/)
  assert.equal(projection.data, undefined); assert.equal(projection.needsPreparation, true)
  assert.equal(chartSeriesSettingSampleCount(1, projection.data), 0)
})

test('unbound large records and series-only data request source preparation instead of silently using rendered rows', () => {
  const draft = createChartSettingsDraft(config()), data = prepared()
  const records = projectChartSettings(draft, undefined, data, false)
  assert.equal(records.data, undefined); assert.equal(records.needsPreparation, true)
  assert.match(records.error, /读取完整数据源/)
  const series = projectChartSettings(draft, undefined, { categories: ['Measurements'], series: [{ field: 1, name: 'Measurements', values: [1] }] })
  assert.equal(series.data, undefined); assert.equal(series.needsPreparation, true)
  assert.match(series.error, /当前字段绑定一致/)
})

test('prepared data is rejected after changing source range, mapping, source revision or host identity', () => {
  const draft = createChartSettingsDraft(config()), data = prepared(), identity = { hostSheetId: 'host', sourceRevision: 'unit-a:0' }
  const bound = chartSettingsPreparedData(draft, data, identity)
  for (const changed of [
    { draft: { ...draft, sourceText: 'A1:C40001' }, identity },
    { draft: { ...draft, config: { ...draft.config, valueFields: [0] } }, identity },
    { draft, identity: { ...identity, sourceRevision: 'unit-a:1' } },
    { draft, identity: { ...identity, hostSheetId: 'different-host' } },
  ]) {
    const projection = projectChartSettings(changed.draft, undefined, bound, false, changed.identity)
    assert.equal(projection.data, undefined); assert.equal(projection.needsPreparation, true)
    assert.match(projection.error, /绑定已更新/)
  }
})

test('source identity includes both the host and a separately selected source worksheet', () => {
  const draft = createChartSettingsDraft(config({ source: { ...config().source, sheetId: 'source' } }))
  assert.notEqual(chartSettingsBindingKey(draft, 'host-a', false), chartSettingsBindingKey(draft, 'host-b', false))
  assert.notEqual(chartSettingsBindingKey(draft, 'host-a', false), chartSettingsBindingKey({ ...draft,
    config: { ...draft.config, source: { ...draft.config.source, sheetId: 'source-b' } } }, 'host-a', false))
})

test('large-source loading yields current raw samples without synchronously projecting the worksheet', async () => {
  const f = fixture(), data = prepared()
  try {
    assert.equal(f.source.loading.value, true); assert.equal(f.source.projection.value.data, undefined)
    assert.throws(() => f.source.assertReady(), /仍在读取/)
    await flush(); assert.equal(f.requests.length, 1); assert.equal(f.cellReads(), 0)
    f.requests[0].resolve(data); await flush()
    assert.equal(f.source.loading.value, false); assert.equal(f.source.projection.value.data, data)
    assert.equal(f.source.assertReady().data.series[0].values.length, SAMPLE_COUNT)
    assert.equal(f.cellReads(), 0)
  } finally { f.dispose() }
})

test('paint, per-sample overrides and Boxplot summary controls reuse the prepared source and raw identities', async () => {
  const f = fixture(), data = prepared()
  try {
    await flush(); f.requests[0].resolve(data); await flush()
    const initial = f.source.projection.value
    f.draft.value = { ...f.draft.value, config: { ...f.draft.value.config, title: 'New title',
      options: { boxplot: { quartileMethod: 'exclusive', whiskerMethod: 'iqr', showOutliers: false, whiskerMultiplier: 2 },
        series: { 1: { color: 'red', dataPoints: { 39999: { size: 18, shape: 'diamond', fillOpacity: 0.35 } } } },
        namedSeries: [{ name: 'Measurements', style: { color: 'navy' } }] } } }
    await flush()
    assert.equal(f.requests.length, 1); assert.equal(f.source.loading.value, false)
    assert.equal(f.source.projection.value, initial); assert.equal(f.source.projection.value.data, data)
    assert.equal(chartSeriesSettingSample(1, 39999, f.source.projection.value.data).value, 39999)
    assert.equal(f.cellReads(), 0)
  } finally { f.dispose() }
})

test('binding edits immediately abort the old request and only the final source mapping is started', async () => {
  const f = fixture()
  try {
    await flush(); const old = f.requests[0]
    f.draft.value = { ...f.draft.value, sourceText: 'A1:C40001' }
    f.draft.value = { ...f.draft.value, config: { ...f.draft.value.config, valueFields: [2] } }
    assert.equal(old.signal.aborted, true); assert.equal(f.source.projection.value.data, undefined)
    await flush(); assert.equal(f.requests.length, 2)
    assert.equal(f.requests[1].config.source.range.endCol, 2); assert.deepEqual(f.requests[1].config.valueFields, [2])
    old.resolve(prepared(1, 'Old result')); await flush()
    assert.equal(f.source.projection.value.data, undefined); assert.equal(f.source.loading.value, true)
    const latest = prepared(2, 'New result', 300)
    f.requests[1].resolve(latest); await flush()
    assert.equal(f.source.projection.value.data, latest)
    assert.equal(chartSeriesSettingSample(2, 39999, latest).label, '40000 · Sample 40000 · 300')
  } finally { f.dispose() }
})

test('a source revision invalidates published samples before late old-source success or rejection can appear', async () => {
  const f = fixture()
  try {
    await flush(); f.requests[0].resolve(prepared()); await flush()
    f.revision.value = 'unit-a:1'
    assert.equal(f.source.projection.value.data, undefined); assert.equal(f.source.loading.value, true)
    await flush(); const obsolete = f.requests[1]
    f.revision.value = 'unit-a:2'; assert.equal(obsolete.signal.aborted, true)
    await flush(); obsolete.reject(Object.assign(new Error('obsolete source cancelled'), { code: 'COMPUTE_CANCELLED' })); await flush()
    assert.equal(f.source.projection.value.error, undefined); assert.equal(f.source.loading.value, true)
    const current = prepared(1, 'Latest measurements', 250)
    f.requests[2].resolve(current); await flush()
    assert.equal(f.source.projection.value.data, current); assert.equal(f.source.loading.value, false)
  } finally { f.dispose() }
})

test('close aborts preparation, clears samples and ignores late completion; reopen gets a fresh source', async () => {
  const f = fixture()
  try {
    await flush(); const closed = f.requests[0]
    f.active.value = false
    assert.equal(closed.signal.aborted, true); assert.equal(f.source.loading.value, false)
    assert.equal(f.source.projection.value.data, undefined)
    closed.resolve(prepared(1, 'Closed panel result')); await flush()
    assert.equal(f.source.projection.value.data, undefined)
    f.active.value = true; await flush(); assert.equal(f.requests.length, 2)
    const fresh = prepared(1, 'Reopened source')
    f.requests[1].resolve(fresh); await flush(); assert.equal(f.source.projection.value.data, fresh)
  } finally { f.dispose() }
})

test('changing the loader context isolates a different workbook even when range, host ID and revision match', async () => {
  const f = fixture(), next = deferred()
  try {
    await flush(); const old = f.requests[0]
    let nextSignal
    f.loader.value = (_config, options) => { nextSignal = options.signal; return next.promise }
    assert.equal(old.signal.aborted, true); await flush(); assert.equal(nextSignal.aborted, false)
    old.resolve(prepared(1, 'Previous workbook')); await flush()
    assert.equal(f.source.projection.value.data, undefined)
    const current = prepared(1, 'Current workbook')
    next.resolve(current); await flush(); assert.equal(f.source.projection.value.data, current)
  } finally { f.dispose() }
})

test('current memory/permission/timeout errors remain reviewable, and retry preserves the author draft', async () => {
  for (const [code, message] of [['COMPUTE_MEMORY_LIMIT', '源数据超过读取内存预算'], ['COMPUTE_INVALID_INPUT', '源工作表不可查看'], ['COMPUTE_TIMEOUT', '图表源读取超时']]) {
    const f = fixture()
    try {
      await flush(); const draft = JSON.stringify(f.draft.value)
      f.requests[0].reject(Object.assign(new Error(message), { code })); await flush()
      assert.equal(f.source.loading.value, false); assert.equal(f.source.projection.value.error, message)
      assert.throws(() => f.source.assertReady(), new RegExp(message))
      f.source.reload(); await flush(); assert.equal(f.requests.length, 2)
      const data = prepared(); f.requests[1].resolve(data); await flush()
      assert.equal(f.source.projection.value.error, undefined); assert.equal(f.source.assertReady().data, data)
      assert.equal(JSON.stringify(f.draft.value), draft); assert.equal(f.cellReads(), 0)
    } finally { f.dispose() }
  }
})

test('an explicitly bound prepared source is reused without asking the loader or applying the small-preview cap', async () => {
  const draft = createChartSettingsDraft(config()), data = prepared()
  const input = chartSettingsPreparedData(draft, data, { hostSheetId: 'host', sourceRevision: 'unit-a:0' })
  const f = fixture({ provided: input })
  try {
    await flush(); assert.equal(f.requests.length, 0); assert.equal(f.cellReads(), 0)
    assert.equal(f.source.loading.value, false); assert.equal(f.source.projection.value.data, data)
  } finally { f.dispose() }
})

test('derived or stale provided data is replaced by the current prepared source rather than treated as raw samples', async () => {
  const draft = createChartSettingsDraft(config()), data = prepared()
  for (const input of [{ ...chartSettingsPreparedData(draft, data, { hostSheetId: 'host', sourceRevision: 'unit-a:0' }), kind: 'derived' },
    chartSettingsPreparedData(draft, data, { hostSheetId: 'host', sourceRevision: 'unit-a:older' })]) {
    const f = fixture({ provided: input })
    try {
      await flush(); assert.equal(f.requests.length, 1); assert.equal(f.source.projection.value.data, undefined)
      const current = prepared(1, 'Current raw samples')
      f.requests[0].resolve(current); await flush(); assert.equal(f.source.projection.value.data, current)
    } finally { f.dispose() }
  }
})

test('invalid range text aborts pending work and reports the input error without starting another source read', async () => {
  const f = fixture()
  try {
    await flush(); f.draft.value = { ...f.draft.value, sourceText: 'invalid range' }
    assert.equal(f.requests[0].signal.aborted, true); assert.equal(f.source.loading.value, false)
    assert.match(f.source.projection.value.error, /有效的数据范围/)
    await flush(); assert.equal(f.requests.length, 1)
    f.requests[0].resolve(prepared()); await flush(); assert.equal(f.source.projection.value.data, undefined)
  } finally { f.dispose() }
})

test('prepared date preflight validates author bounds and calendar options without enumerating a large source', () => {
  const data = { date1904: false, get categories() { throw new Error('Do not enumerate dates') }, get series() { throw new Error('Do not enumerate samples') } }
  const valid = config({ options: { mapping: { useDateAxis: true } }, xAxis: { scaleType: 'linear', majorTimeUnit: 'months', majorUnit: 1 } })
  assert.equal(chartSettingsDateAxisError(valid, data, false, 'options'), undefined)
  assert.match(chartSettingsDateAxisError({ ...valid, xAxis: { ...valid.xAxis, min: -1 } }, data, false, 'options'), /边界/)
  assert.match(chartSettingsDateAxisError({ ...valid, xAxis: { ...valid.xAxis, majorUnit: 0.5 } }, data, false, 'options'), /安全整数/)
})

test('scope disposal aborts source preparation and drops late success without retaining visible samples', async () => {
  const f = fixture()
  await flush(); const request = f.requests[0]
  f.dispose(); assert.equal(request.signal.aborted, true); assert.equal(f.source.loading.value, false)
  request.resolve(prepared()); await flush(); assert.equal(f.source.projection.value.data, undefined)
})

// This fixture's known sort has one 100, 39,998 ones and one null. The service
// materializes it; settings never rebuild the preparation or the sort.
function largeParetoStatistics(data = prepared()) {
  const categories = [data.categories.at(-1), data.categories[0], ...data.categories.slice(2, -1)]
  const total = 100 + categories.length - 1
  return { preparedData: data, pareto: { categories, values: [100, ...Array(categories.length - 1).fill(1)],
    cumulativePercent: categories.map((_category, index) => (100 + index) / total * 100) } }
}

test('Pareto uses one statistical read for an over-budget source and retains its prepared arrays and materialized ordering', async () => {
  const f = fixture({ config: { type: 'pareto' } }), statistics = largeParetoStatistics()
  try {
    await flush(); assert.equal(f.statisticsRequests.length, 1); assert.equal(f.requests.length, 0); assert.equal(f.cellReads(), 0)
    f.statisticsRequests[0].resolve(statistics); await flush()
    const projection = f.source.assertReady()
    assert.equal(projection.data, statistics.preparedData); assert.equal(projection.statistics, statistics)
    assert.equal(projection.source.records, statistics.preparedData.records)
    assert.equal(chartParetoSettingsCategories(projection.statistics), statistics.pareto.categories)
    assert.equal(chartParetoSettingsCategories(projection.statistics)[0], 'Sample 40000')
    assert.equal(f.cellReads(), 0); assert.equal(f.requests.length, 0)
  } finally { f.dispose() }
})

test('settings preserve a prepared Top N / others projection and normalized null instead of applying them twice', async () => {
  const f = fixture({ config: { type: 'pareto', options: { invalidValueStrategy: 'zero',
    aggregation: { aggregate: true, topN: { count: 1, includeOthers: true, othersLabel: 'Other causes' } }, pareto: { includeZeroValues: true } } } })
  const data = { fields: ['Cause', 'Count'], records: [['A', 5], ['Other causes', 6]], sourceNumericFields: [1], date1904: false,
    categories: ['A', 'Other causes'], series: [{ field: 1, name: 'Count', values: [5, 6] }] }
  const statistics = { preparedData: data, pareto: { categories: ['Other causes', 'A'], values: [6, 5], cumulativePercent: [600 / 11, 100] } }
  try {
    await flush(); f.statisticsRequests[0].resolve(statistics); await flush()
    assert.equal(f.source.projection.value.data, data)
    assert.equal(chartParetoSettingsCategories(f.source.projection.value.statistics), statistics.pareto.categories)
    assert.deepEqual(f.source.projection.value.data.categories, ['A', 'Other causes'])
    assert.deepEqual(f.source.projection.value.statistics.pareto.values, [6, 5]); assert.equal(f.requests.length, 0)
    f.draft.value = { ...f.draft.value, config: { ...f.draft.value.config,
      options: { invalidValueStrategy: 'zero', pareto: { includeZeroValues: true } } } }
    await flush(); assert.equal(f.statisticsRequests.length, 2)
    const zeroData = { fields: ['Cause', 'Count'], records: [['missing', null], ['A', 3]], date1904: false,
      categories: ['missing', 'A'], series: [{ field: 1, name: 'Count', values: [0, 3] }] }
    const zeroStatistics = { preparedData: zeroData, pareto: { categories: ['A', 'missing'], values: [3, 0], cumulativePercent: [100, 100] } }
    f.statisticsRequests[1].resolve(zeroStatistics); await flush()
    assert.equal(f.source.assertReady().data, zeroData)
    assert.equal(f.source.projection.value.source.records[0][1], null)
    assert.equal(f.source.projection.value.data.series[0].values[0], 0)
    assert.equal(chartParetoSettingsCategories(f.source.projection.value.statistics), zeroStatistics.pareto.categories)
  } finally { f.dispose() }
})

test('materialized Pareto classification does not inspect or normalize the prepared source', () => {
  const categories = ['Others', 'Top cause'], statistics = { pareto: { categories, values: [6, 5], cumulativePercent: [600 / 11, 100] },
    get preparedData() { throw new Error('No settings-side source scan or re-preparation') } }
  assert.equal(chartParetoSettingsCategories(statistics), categories)
  assert.deepEqual(chartParetoSettingsCategories(), [])
  assert.throws(() => chartSettingsPreparedStatistics(createChartSettingsDraft(config({ type: 'pareto' })), {}), /未返回准备后的原始数据源/)
  assert.throws(() => chartSettingsPreparedStatistics(createChartSettingsDraft(config({ type: 'pareto' })), { preparedData: prepared() }), /未返回当前帕累托分类/)
})

test('Pareto zero inclusion is part of read identity, while paint and mark overrides reuse the statistical result', async () => {
  const f = fixture({ config: { type: 'pareto' } }), statistics = largeParetoStatistics()
  try {
    await flush(); f.statisticsRequests[0].resolve(statistics); await flush()
    const initial = f.source.projection.value, draft = f.draft.value, key = chartSettingsBindingKey(draft, 'host', false)
    f.draft.value = { ...draft, config: { ...draft.config, title: 'New Pareto title', options: { pareto: { includeZeroValues: false,
      barStyle: { color: 'navy', dataPoints: { 0: { fillOpacity: 0.2 } } }, cumulativeLineStyle: { width: 5, dataPoints: { 1: { size: 18 } } } } } } }
    await flush(); assert.equal(f.statisticsRequests.length, 1); assert.equal(f.source.projection.value, initial)
    assert.equal(chartSettingsBindingKey(f.draft.value, 'host', false), key)
    f.draft.value = { ...f.draft.value, config: { ...f.draft.value.config,
      options: { ...f.draft.value.config.options, pareto: { ...f.draft.value.config.options.pareto, includeZeroValues: true } } } }
    assert.notEqual(chartSettingsBindingKey(f.draft.value, 'host', false), key)
    assert.equal(f.source.projection.value.statistics, undefined); await flush(); assert.equal(f.statisticsRequests.length, 2)
    const next = { preparedData: statistics.preparedData, pareto: { categories: ['Sample 40000', 'zero included'], values: [100, 0], cumulativePercent: [100, 100] } }
    f.statisticsRequests[1].resolve(next); await flush(); assert.equal(f.source.projection.value.statistics, next)
    assert.equal(f.requests.length, 0)
  } finally { f.dispose() }
})

test('source-key and revision changes isolate old Pareto success and failure before a later statistical result arrives', async () => {
  const f = fixture({ config: { type: 'pareto' } })
  try {
    await flush(); const first = f.statisticsRequests[0]
    f.draft.value = { ...f.draft.value, sourceText: 'A1:C40001', config: { ...f.draft.value.config, valueFields: [2] } }
    assert.equal(first.signal.aborted, true); assert.equal(f.source.projection.value.statistics, undefined)
    await flush(); assert.equal(f.statisticsRequests.length, 2)
    f.revision.value = 'unit-a:1'; const second = f.statisticsRequests[1]
    assert.equal(second.signal.aborted, true); await flush(); assert.equal(f.statisticsRequests.length, 3)
    first.resolve(largeParetoStatistics()); second.reject(new Error('old source error')); await flush()
    assert.equal(f.source.projection.value.error, undefined); assert.equal(f.source.projection.value.statistics, undefined)
    const latest = largeParetoStatistics(prepared(2, 'Current source'))
    f.statisticsRequests[2].resolve(latest); await flush(); assert.equal(f.source.projection.value.statistics, latest)
    assert.equal(f.source.loading.value, false); assert.equal(f.requests.length, 0)
  } finally { f.dispose() }
})

test('switching chart family cancels its Pareto read and starts only the source-only path for Boxplot', async () => {
  const f = fixture({ config: { type: 'pareto' } })
  try {
    await flush(); const old = f.statisticsRequests[0]
    f.draft.value = { ...f.draft.value, config: { ...f.draft.value.config, type: 'boxplot' } }
    assert.equal(old.signal.aborted, true); await flush(); assert.equal(f.requests.length, 1)
    old.resolve(largeParetoStatistics()); await flush(); assert.equal(f.source.projection.value.statistics, undefined)
    const data = prepared(); f.requests[0].resolve(data); await flush()
    assert.equal(f.source.assertReady().data, data); assert.equal(f.source.projection.value.statistics, undefined)
  } finally { f.dispose() }
})

test('closing and replacing a Pareto loader context cancel old work without overwriting the reopened panel', async () => {
  const f = fixture({ config: { type: 'pareto' } }), replacement = deferred()
  try {
    await flush(); const old = f.statisticsRequests[0]
    f.active.value = false; assert.equal(old.signal.aborted, true); assert.equal(f.source.projection.value.statistics, undefined)
    f.active.value = true; await flush(); const reopened = f.statisticsRequests[1]
    let replacementSignal
    f.statisticsLoader.value = (_config, options) => { replacementSignal = options.signal; return replacement.promise }
    assert.equal(reopened.signal.aborted, true); await flush(); assert.equal(replacementSignal.aborted, false)
    old.resolve(largeParetoStatistics()); reopened.reject(new Error('closed-context failure')); await flush()
    assert.equal(f.source.projection.value.error, undefined); assert.equal(f.source.projection.value.statistics, undefined)
    const current = largeParetoStatistics(prepared(1, 'New workbook'))
    replacement.resolve(current); await flush(); assert.equal(f.source.assertReady().statistics, current)
  } finally { f.dispose() }
})

test('matching provided statistics bypass both readers; source-only prepared data still requests Pareto statistics', async () => {
  const draft = createChartSettingsDraft(config({ type: 'pareto' })), statistics = largeParetoStatistics()
  const identity = { hostSheetId: 'host', sourceRevision: 'unit-a:0' }
  for (const withStatistics of [true, false]) {
    const provided = withStatistics ? chartSettingsPreparedStatistics(draft, statistics, identity)
      : chartSettingsPreparedData(draft, statistics.preparedData, identity)
    const f = fixture({ config: { type: 'pareto' }, provided })
    try {
      await flush(); assert.equal(f.requests.length, 0); assert.equal(f.statisticsRequests.length, withStatistics ? 0 : 1)
      if (!withStatistics) { f.statisticsRequests[0].resolve(statistics); await flush() }
      assert.equal(f.source.assertReady().statistics, statistics); assert.equal(f.cellReads(), 0)
    } finally { f.dispose() }
  }
})

test('current Pareto failure is shown and retry retains authored mark styles without falling back to a synchronous source read', async () => {
  const f = fixture({ config: { type: 'pareto', options: { pareto: { cumulativeLineStyle: { width: 9 } } } } })
  try {
    await flush(); f.statisticsRequests[0].reject(Object.assign(new Error('Pareto statistics memory limit'), { code: 'COMPUTE_MEMORY_LIMIT' })); await flush()
    assert.match(f.source.projection.value.error, /statistics memory limit/); assert.equal(f.source.loading.value, false)
    assert.throws(() => f.source.assertReady(), /statistics memory limit/)
    f.source.reload(); await flush(); assert.equal(f.statisticsRequests.length, 2); assert.equal(f.requests.length, 0)
    const current = largeParetoStatistics(); f.statisticsRequests[1].resolve(current); await flush()
    assert.equal(f.source.assertReady().statistics, current); assert.equal(f.draft.value.config.options.pareto.cumulativeLineStyle.width, 9)
  } finally { f.dispose() }
})

test('materialized category choices stay bounded and direct lookup exposes a far mark without enumerating intervening categories', () => {
  const reads = [], categories = new Proxy(Array.from({ length: SAMPLE_COUNT }, (_value, index) => `Category ${index + 1}`), {
    get(target, key, receiver) {
      if (key === 'map') throw new Error('No unbounded Select-option map')
      if (typeof key === 'string' && /^\d+$/.test(key)) {
        const index = Number(key); reads.push(index)
        if (index >= CHART_CATEGORY_POINT_CHOICES_LIMIT && index !== SAMPLE_COUNT - 1) throw new Error('Only the selected far category may be read')
      }
      return Reflect.get(target, key, receiver)
    },
  })
  const items = chartCategorySettingPoints(categories, SAMPLE_COUNT - 1)
  assert.equal(items.length, CHART_CATEGORY_POINT_CHOICES_LIMIT + 1)
  assert.deepEqual(items.at(-1), { value: 39999, label: '40000 · Category 40000' })
  assert.equal(reads.length, CHART_CATEGORY_POINT_CHOICES_LIMIT + 1)
  for (const index of [-1, 0.5, SAMPLE_COUNT, NaN, Infinity]) assert.equal(chartCategorySettingPoint(categories, index), undefined)
  assert.deepEqual(chartSeriesSettingPoint(config({ type: 'pareto' }), 1, categories, undefined, 39999), items.at(-1))
})

test('bounded Bubble choices and direct number lookup use the same named-group membership', () => {
  const categories = Array.from({ length: SAMPLE_COUNT }, (_value, index) => `Point ${index + 1}`)
  const data = { categories, records: Array.from({ length: SAMPLE_COUNT }, (_value, index) => [index % 2 ? 'B' : 'A']),
    series: [{ field: 1, name: 'A', values: [] }, { field: 1, name: 'B', values: [] }] }
  const bubble = config({ type: 'bubble', options: { mapping: { bubble: { seriesIndex: 0 } } } })
  const items = chartSeriesSettingPoints(bubble, { name: 'B' }, categories, [], data, SAMPLE_COUNT - 1)
  assert.equal(items.length, CHART_CATEGORY_POINT_CHOICES_LIMIT / 2 + 1)
  assert.deepEqual(items.at(-1), { value: 39999, label: '40000 · Point 40000' })
  assert.deepEqual(chartSeriesSettingPoint(bubble, { name: 'B' }, categories, data, 39999), items.at(-1))
  assert.equal(chartSeriesSettingPoint(bubble, { name: 'B' }, categories, data, 39998), undefined)
  assert.deepEqual(chartSeriesSettingPoints(bubble, { name: 'missing' }, categories, [], data), [])
  assert.equal(chartSeriesSettingPoint(bubble, { name: 'missing' }, categories, data, 39999), undefined)
})
