import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Draft for the unified validation phase. No test/build/browser run is implied.
let server, statistics, math, geometry, catalog, derived
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[statistics, math, geometry, catalog, derived] = await Promise.all([
    'chart-statistics', 'chart-math', 'chart-geometry', 'chart-catalog', 'chart-derived',
  ].map(name => server.ssrLoadModule(`/packages/office-studio/src/features/charts/${name}.ts`)))
})
afterAll(async () => { await server?.close() })

function chart(options = {}, type = 'line') {
  return { id: 'chart', sheetId: 'sheet', type, source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 480, height: 300 }, categoryField: 0, valueFields: [1],
    legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 0, options }
}
function close(actual, expected, tolerance = 1e-12) { assert.ok(Math.abs(actual - expected) <= tolerance * Math.max(1, Math.abs(expected)), `${actual} ≈ ${expected}`) }

test('finite products apply small factors before deciding final overflow or underflow', () => {
  close(math.chartFiniteProduct([Number.MAX_VALUE, 2, 0.25]) / Number.MAX_VALUE, 0.5)
  assert.ok(math.chartFiniteProduct([Number.MIN_VALUE, Number.MAX_VALUE, 0.01]) > 0)
  assert.equal(math.chartFiniteProduct([0, Number.MAX_VALUE, 2]), 0)
  assert.throws(() => math.chartFiniteProduct([Number.MAX_VALUE, 2]), RangeError)
  assert.throws(() => math.chartFiniteProduct([0, Infinity]), RangeError)
})

test('normalized population moments and moving averages preserve extreme and subnormal samples', () => {
  const extremes = math.chartScaledMoments([-Number.MAX_VALUE, Number.MAX_VALUE])
  assert.equal(extremes.mean, 0)
  assert.equal(extremes.scaledVariance, 1)
  assert.equal(Object.hasOwn(extremes, 'deviation'), false)
  assert.equal(math.chartScaledMoments([Number.MAX_VALUE, Number.MAX_VALUE]).mean, Number.MAX_VALUE)
  assert.equal(math.chartScaledMoments([Number.MIN_VALUE, Number.MIN_VALUE]).mean, Number.MIN_VALUE)
  assert.deepEqual(statistics.chartTrendline({ field: 1, name: 'Tiny', values: [Number.MIN_VALUE, Number.MIN_VALUE] }, { seriesIndex: 1, type: 'movingAvg', period: 2 }), { points: [[2, Number.MIN_VALUE]] })
})

test('constant count histograms retain one unpadded point interval, including finite extremes', () => {
  const ordinary = statistics.chartHistogram([5, 5, 5], { binning: { mode: 'count', count: 4 } })
  assert.deepEqual(ordinary.bins, [{ kind: 'regular', lowerBound: 5, upperBound: 5, count: 3 }])
  const extreme = statistics.chartHistogram([Number.MAX_VALUE, Number.MAX_VALUE], { binning: { mode: 'count', count: 4 } })
  assert.deepEqual(extreme.bins, [{ kind: 'regular', lowerBound: Number.MAX_VALUE, upperBound: Number.MAX_VALUE, count: 2 }])
})

test('histogram thresholds and shared boundaries assign each finite sample once', () => {
  const result = statistics.chartHistogram([-10, -5, 0, 5, 10, 15, NaN, Infinity], { underflowThreshold: 0, overflowThreshold: 10, binning: { mode: 'count', count: 4 } })
  assert.equal(result.validSampleCount, 6)
  assert.deepEqual(result.bins.map(bin => [bin.kind, bin.count]), [['underflow', 3], ['regular', 1], ['regular', 1], ['overflow', 1]])
  const width = statistics.chartHistogram([1, 4, 7], { underflowThreshold: 0, overflowThreshold: 7, binning: { mode: 'width', width: 5 } })
  assert.deepEqual(width.bins.filter(bin => bin.kind === 'regular').map(bin => [bin.lowerBound, bin.upperBound, bin.count]), [[0, 5, 2], [5, 7, 1]])
})

test('boxplot reference keeps stable source identities for equal outliers instead of compacting empty slots', () => {
  const values = [100, null, 0, 0, 100, 0, 0, 0, 0, 0, 0, 0], before = [...values]
  const result = statistics.chartBoxplot(values, { quartileMethod: 'inclusive', whiskerMethod: 'iqr' })
  assert.deepEqual(result, { min: 0, q1: 0, median: 0, q3: 0, max: 0, outliers: [100, 100], outlierSourceIndices: [0, 4] })
  assert.deepEqual(values, before)
  const data = { categories: values.map((_value, index) => String(index)), series: [{ field: 1, name: 'Samples', values }] }
  assert.deepEqual(derived.deriveChartData(data, chart({ boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr' } }, 'boxplot')).boxes[0].statistics.outlierSourceIndices, [0, 4])
})

test('fixed boxplot defaults use each half median and every finite sample as the whisker range', () => {
  assert.deepEqual(statistics.chartBoxplot([4, 2, 1, 3]), { min: 1, q1: 1.5, median: 2.5, q3: 3.5, max: 4, outliers: [], outlierSourceIndices: [] })
  assert.deepEqual(statistics.chartBoxplot([100, null, 3, 2, 4, 1]), { min: 1, q1: 1.5, median: 3, q3: 52, max: 100, outliers: [], outlierSourceIndices: [] })
  assert.deepEqual(statistics.chartBoxplot([100, 5, 4, 3, 2, 1]), { min: 1, q1: 2, median: 3.5, q3: 5, max: 100, outliers: [], outlierSourceIndices: [] })
  assert.equal(statistics.chartBoxplot([null, 7]), undefined)
  assert.equal(statistics.chartBoxplot([]), undefined)
  assert.deepEqual(statistics.chartBoxplot([7, 7]), { min: 7, q1: 7, median: 7, q3: 7, max: 7, outliers: [], outlierSourceIndices: [] })
})

test('legacy interpolated boxplot options retain IQR while an explicit full range overrides them', () => {
  const legacy = statistics.chartBoxplot([1, 2, 3, 4, 100], { quartileMethod: 'inclusive' })
  assert.deepEqual(legacy, { min: 1, q1: 2, median: 3, q3: 4, max: 4, outliers: [100], outlierSourceIndices: [4] })
  const full = statistics.chartBoxplot([1, 2, 3, 4, 100], { quartileMethod: 'inclusive', whiskerMethod: 'range' })
  assert.equal(full.max, 100); assert.deepEqual(full.outliers, [])
  assert.equal(statistics.chartBoxplot([7], { quartileMethod: 'inclusive' }).median, 7)
  assert.throws(() => statistics.chartBoxplot([1, 2], { whiskerMethod: 'other' }), RangeError)
})

test('linear fits and R² remain finite when subtracting an original X mean would overflow', () => {
  const maximum = Number.MAX_VALUE
  const result = statistics.chartTrendline({ field: 1, name: 'Huge', xValues: [-maximum, maximum, maximum], values: [-maximum, maximum, maximum] }, { seriesIndex: 1, type: 'linear' })
  assert.ok(result.points.length >= 20)
  assert.ok(result.points.flat().every(Number.isFinite))
  close(result.r2, 1)
  assert.doesNotMatch(result.formula, /Infinity|NaN/)
})

test('polynomial fits use normalized Y instead of overflowing the normal-equation RHS', () => {
  const maximum = Number.MAX_VALUE
  const result = statistics.chartTrendline({ field: 1, name: 'Square', xValues: [-2, -1, 0, 1, 2], values: [0.8 * maximum, 0.2 * maximum, 0, 0.2 * maximum, 0.8 * maximum] }, { seriesIndex: 1, type: 'poly', order: 2 })
  assert.equal(result.points.length, 33)
  assert.ok(result.points.flat().every(Number.isFinite))
  close(result.r2, 1)
  assert.doesNotMatch(result.formula, /Infinity|NaN/)
})

test('Rust fit metadata uses the same finite equation formatter and rejects invalid normalized means', () => {
  const fit = { coefficients: [0, 1], xScale: Number.MAX_VALUE, xMean: 0, xSpread: 2, yScale: Number.MAX_VALUE, yMean: 0 }
  assert.doesNotMatch(statistics.formatChartTrendlineEquation('linear', fit), /Infinity|NaN/)
  assert.equal(statistics.formatChartTrendlineEquation('movingAvg', fit), undefined)
  assert.throws(() => statistics.formatChartTrendlineEquation('linear', { ...fit, yMean: 2 }), RangeError)
})

test('fixed errors do not precompute an unrepresentable sample deviation', () => {
  const series = { field: 1, name: 'Huge', values: [-Number.MAX_VALUE, Number.MAX_VALUE] }
  const result = statistics.chartErrorIntervals({ categories: ['A', 'B'], series: [series] }, series, { seriesIndex: 1, direction: 'y', type: 'both', valueType: 'fixedVal', value: 0 })
  assert.equal(result.length, 2)
  assert.ok(result.every(interval => interval.plus === 0 && interval.minus === 0))
})

test('error calculations validate only the selected direction and only complete X/Y samples', () => {
  const series = { field: 1, name: 'Huge', values: [Number.MAX_VALUE] }
  assert.deepEqual(statistics.chartErrorIntervals({ categories: ['A'], series: [series] }, series, { seriesIndex: 1, direction: 'y', type: 'minus', valueType: 'fixedVal', value: Number.MAX_VALUE }).map(interval => [interval.minus, interval.plus]), [[Number.MAX_VALUE, 0]])
  const sparse = { field: 1, name: 'Pairs', values: [10, 20, 1000], xValues: [1, 2, null] }
  const errors = statistics.chartErrorIntervals({ categories: ['A', 'B', 'C'], series: [sparse] }, sparse, { seriesIndex: 1, direction: 'y', type: 'minus', valueType: 'stdErr' })
  assert.equal(errors.length, 2)
  close(errors[0].minus, 5)
})

test('custom missing cells omit an interval, unused negative cells are ignored, and references must align', () => {
  const series = { field: 1, name: 'Sales', values: [10, 20] }
  const data = { categories: ['A', 'B'], series: [series], fields: ['Category', 'Sales', 'Plus', 'Unused'], records: [['A', 10, null, -3], ['B', 20, 2, -4]] }
  const config = { seriesIndex: 1, direction: 'y', type: 'plus', valueType: 'cust', plusIndex: 2, minusIndex: 3 }
  assert.deepEqual(statistics.chartErrorIntervals(data, series, config).map(interval => [interval.categoryIndex, interval.plus, interval.minus]), [[1, 2, 0]])
  assert.throws(() => statistics.chartErrorIntervals({ ...data, records: undefined }, series, config), /逐项对应/)
  assert.throws(() => statistics.chartErrorIntervals(data, series, { ...config, plusIndex: 4 }), /超出/)
  assert.throws(() => statistics.chartErrorIntervals(data, series, { ...config, plusIndex: 3 }), /负数/)
})

test('Top N reorders custom error records with points instead of silently substituting zeros', () => {
  const data = { categories: ['A', 'B', 'C'], fields: ['Category', 'Sales', 'Plus'], records: [['A', 1, 1], ['B', 30, 3], ['C', 20, 2]], series: [{ field: 1, name: 'Sales', values: [1, 30, 20] }] }
  const model = chart({ aggregation: { topN: { count: 2, includeOthers: false } }, errorBars: [{ seriesIndex: 1, direction: 'y', type: 'plus', valueType: 'cust', plusIndex: 2 }] })
  const result = derived.deriveChartData(data, model)
  assert.deepEqual(result.data.categories, ['B', 'C'])
  assert.deepEqual(result.data.records, [data.records[1], data.records[2]])
  assert.deepEqual(result.errors[0].intervals.map(interval => interval.plus), [3, 2])
  assert.throws(() => derived.deriveChartData({ ...data, categories: ['A', 'A', 'C'] }, chart({ ...model.options, aggregation: { aggregate: true } })), /逐项对应/)
})

test('catalog capabilities and derived statistics agree for every stacked/3D variant and mixed combination series', () => {
  for (const entry of catalog.CHART_TYPE_CATALOG) if (entry.stacked || entry.threeD) assert.equal(catalog.supportsChartStatistics(chart({}, entry.value)), false, entry.id)
  const mixed = chart({ series: { 1: { type: 'columnStacked' }, 2: { type: 'line' } } }, 'combination')
  assert.equal(catalog.supportsChartStatistics(mixed, 1, 0), false)
  assert.equal(catalog.supportsChartStatistics(mixed, 2, 1), true)
  assert.equal(catalog.supportsChartErrorDirection(chart({}, 'scatter'), 'x'), true)
  assert.equal(catalog.supportsChartErrorDirection(chart({}, 'bar'), 'x'), false)
  const data = { categories: ['A', 'B'], series: [{ field: 1, name: 'Sales', values: [1, 2] }] }
  const result = derived.deriveChartData(data, chart({ trendlines: [{ seriesIndex: 1, type: 'linear' }], errorBars: [{ seriesIndex: 1, direction: 'y', type: 'both', valueType: 'fixedVal' }] }, 'line3D'))
  assert.deepEqual(result.trends, [])
  assert.deepEqual(result.errors, [])
})

test('automatic and single-bound axes keep finite ordered extents and include zero for negative linear data', () => {
  assert.deepEqual(geometry.chartAxisExtent([-20, -10], { scaleType: 'linear' }), { min: -20, max: 0 })
  const above = geometry.chartAxisExtent([1, 2], { scaleType: 'linear', min: 10 })
  const below = geometry.chartAxisExtent([1, 2], { scaleType: 'linear', max: -10 })
  assert.ok(above.min === 10 && above.max > 10)
  assert.ok(below.max === -10 && below.min < -10)
  const extreme = geometry.chartAxisExtent([Number.MAX_VALUE], { scaleType: 'log' })
  assert.ok(extreme.min > 0 && extreme.min < extreme.max && Number.isFinite(extreme.max))
  assert.ok(geometry.chartAxisTicks(extreme.min, extreme.max, { scaleType: 'log' }).every(Number.isFinite))
  assert.throws(() => geometry.chartAxisExtent([1], { scaleType: 'linear', min: Number.MAX_VALUE }), /有限/)
})

test('nearby large logarithmic bounds use a finite ratio even when their rounded logs coincide', () => {
  const first = 1e308, last = first * (1 + Number.EPSILON)
  assert.ok(last > first)
  assert.equal(geometry.chartAxisRatio(first, first, last, { scaleType: 'log' }), 0)
  assert.equal(geometry.chartAxisRatio(last, first, last, { scaleType: 'log' }), 1)
  assert.doesNotMatch(geometry.formatChartNumber(Number.MAX_VALUE), /Infinity|NaN/)
})
