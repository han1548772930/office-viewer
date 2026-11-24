import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { init as initNativeChart, use as useNativeChart } from 'echarts/core'
import { GaugeChart } from 'echarts/charts'
import { SVGRenderer } from 'echarts/renderers'

// Draft for unified validation. This file has not been executed in this phase.
let server, catalog, dataModel, projection, derived, layout, draw, pie, special, threeD, echarts, color, options, settings, conversion, engine, gauge, gaugeApi, gaugeControls, controls, drafts, typeControls
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[catalog, dataModel, projection, derived, layout, draw, pie, special, threeD, echarts, color, options, settings, conversion, engine, gauge, gaugeApi, gaugeControls, controls, drafts, typeControls] = await Promise.all([
    'chart-catalog', 'chart-data', 'chart-record-projection', 'chart-derived', 'chart-layout', 'chart-draw', 'chart-pie-layout', 'chart-special-layout', 'chart-3d',
    'echarts-option', 'chart-color', 'chart-options', 'chart-settings-data', 'chart-type-conversion', 'chart-engine',
    'chart-gauge', 'chart-gauge-api', 'chart-gauge-settings-controls', 'chart-settings-controls', 'chart-settings-draft', 'chart-type-settings-controls',
  ].map(name => server.ssrLoadModule(`/packages/office-studio/src/features/charts/${name}.ts`)))
})
afterAll(async () => { await server?.close() })

function chart(type = 'line', overrides = {}) {
  return { id: 'chart', sheetId: 'sheet', type, source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 3 } },
    position: { x: 0, y: 0 }, size: { width: 480, height: 300 }, categoryField: 0, valueFields: [1],
    legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 0, ...overrides }
}
function prepared(source, model, computed = { preparedData: source }) { return derived.deriveChartData(source, model, computed) }
function commands(source, model, computed) { return draw.drawDerivedChartData(prepared(source, model, computed), model) }
function option(source, model, computed = { preparedData: source }) { return echarts.chartEChartsOption({ sheets: [] }, model, {}, computed) }
function near(actual, expected) { assert.ok(Math.abs(actual - expected) < 1e-8, `${actual} ≈ ${expected}`) }
function nearNative(actual, expected) { assert.ok(Math.abs(actual - expected) < 1e-4, `${actual} ≈ ${expected}`) }
function finiteCommands(scene) {
  for (const command of scene) {
    const coordinates = command.kind === 'polygon' ? command.points.flatMap(point => [point.x, point.y])
      : command.kind === 'rect' ? [command.x, command.y, command.width, command.height]
        : command.kind === 'line' ? [command.x1, command.y1, command.x2, command.y2] : [command.x, command.y]
    assert.ok(coordinates.every(Number.isFinite), command.kind)
  }
}
const ordinary = () => ({ categories: ['A', 'B', 'C'], series: [{ field: 1, name: 'Value', values: [10, 20, 30] }] })

test('all 45 catalog entries retain a finite empty-data scene and an explicit empty-state label', () => {
  const source = { fields: ['Category', 'Value'], records: [], visibleFields: [], categories: [], series: [] }
  for (const entry of catalog.CHART_TYPE_CATALOG) {
    const model = chart(entry.value, { title: entry.label }), result = prepared(source, model), geometry = layout.layoutDerivedChartData(result, model)
    assert.deepEqual(geometry.series, [], entry.id)
    const scene = draw.drawDerivedChartData(result, model)
    finiteCommands(scene)
    assert.ok(scene.some(command => command.kind === 'text' && command.text === '暂无可绘制数据'), entry.id)
    assert.ok(option(source, model).graphic.elements.some(element => element.type === 'text' && element.style.text === '暂无可绘制数据'), entry.id)
  }
})

test('source hierarchy DFS indices survive treemap area ordering and retain full paths', () => {
  const nodes = [{ name: 'Small', value: 1, children: [{ name: 'Leaf', value: 1 }] }, { name: 'Large', value: 9, children: [{ name: 'Leaf', value: 9 }] }]
  const items = dataModel.chartHierarchyItems(nodes)
  assert.deepEqual(items.map(item => [item.pointIndex, item.rootIndex, item.path]), [[0, 0, ['Small']], [1, 0, ['Small', 'Leaf']], [2, 1, ['Large']], [3, 1, ['Large', 'Leaf']]])
  const rectangles = special.chartTreemapLayout(nodes, { x: 10, y: 20, width: 300, height: 200 })
  assert.equal(rectangles[0].node.name, 'Large')
  for (const rectangle of rectangles) assert.equal(rectangle.pointIndex, items.find(item => item.node === rectangle.node).pointIndex)
  const model = chart('treemap'), source = { categories: nodes.map(node => node.name), hierarchy: nodes, series: [{ field: 1, name: 'Weight', values: [1, 9] }] }
  const geometry = layout.layoutDerivedChartData(prepared(source, model), model)
  assert.deepEqual(geometry.categories, ['Small', 'Small / Leaf', 'Large', 'Large / Leaf'])
  assert.deepEqual(new Set(geometry.series[0].points.map(point => point.categoryIndex)), new Set([0, 1, 2, 3]))
})

test('every hierarchy node resolves the same point color, opacity and border in both adapters', () => {
  const nodes = [{ name: 'Small', value: 1 }, { name: 'Large', value: 9, children: [{ name: 'Leaf', value: 9 }] }]
  for (const type of ['treemap', 'sunburst']) {
    const model = chart(type, { options: { series: { 1: { dataPoints: { 2: { color: 'magenta', fillOpacity: 0.4, border: { color: 'navy', width: 3 } } } } } } })
    const source = { categories: ['Small', 'Large'], hierarchy: nodes, series: [{ field: 1, name: 'Weight', values: [1, 9] }] }
    const leaf = option(source, model).series[0].data[1].children[0]
    assert.equal(color.normalizeChartColor(leaf.itemStyle.color), 'rgba(255, 0, 255, 1)')
    assert.equal(leaf.itemStyle.opacity, 0.4)
    assert.equal(leaf.itemStyle.borderWidth, 3)
    assert.ok(commands(source, model).some(command => (command.kind === 'rect' || command.kind === 'polygon') && command.fill === 'rgba(255, 0, 255, 0.4)'))
  }
})

test('single-ended gauge, chord and word-cloud parameters validate against effective defaults', () => {
  for (const invalid of [{ gauge: { min: 100 } }, { gauge: { max: 0 } }, { gauge: { ranges: [{ from: -1, to: 20, color: 'red' }] } },
    { chord: { innerRadiusRatio: 0.9 } }, { chord: { outerRadiusRatio: 0.6 } }, { wordCloud: { minFontSize: 60 } }, { wordCloud: { maxFontSize: 5 } }]) assert.throws(() => options.cloneChartOptions(invalid), RangeError)
  assert.doesNotThrow(() => options.cloneChartOptions({ gauge: { min: -10 }, chord: { innerRadiusRatio: 0.5 }, wordCloud: { minFontSize: 8 } }))
})

test('the categorical depth axis is independently preserved and rejects numeric-axis fields', () => {
  const current = { threeD: { view: { rotY: 44 } }, depthAxis: { title: 'Region', reverse: true, line: { color: 'blue' } } }
  assert.deepEqual(options.mergeChartOptions(current, { threeD: undefined }).depthAxis, current.depthAxis)
  assert.deepEqual(options.mergeChartOptions(current, { depthAxis: undefined }).threeD, current.threeD)
  assert.throws(() => options.cloneChartOptions({ depthAxis: { min: 0 } }), /分类深度轴/)
  assert.throws(() => options.cloneChartOptions({ depthAxis: { scaleType: 'linear' } }), /分类深度轴/)
})

test('pivot numeric roles reject numeric-looking categories while graph endpoints remain nominal', () => {
  const records = { fields: ['Source ID', 'Target ID', 'Weight', 'Size'], records: [[100, 200, 5, 8], [200, 300, 4, 7]], sourceCategoryFields: [0], sourceValueFields: [2, 3] }
  assert.throws(() => projection.projectChartRecords(records, chart('line', { valueFields: [1] })), /数值字段/)
  assert.throws(() => projection.projectChartRecords(records, chart('bubble', { valueFields: [2, 3], options: { mapping: { bubble: { xIndex: 0, yIndex: 2, sizeIndex: 3 } } } })), /数值字段/)
  const graph = projection.projectChartRecords(records, chart('sankey', { valueFields: [1, 2], options: { mapping: { edge: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 } } } }))
  assert.deepEqual(graph.graph.links.map(link => [link.source, link.target, link.value]), [['100', '200', 5], ['200', '300', 4]])
})

test('hidden invalid scatter Y and active custom-error references still fail before the empty-data branch', () => {
  const records = { fields: ['X', 'Y', 'Error'], records: [[1, 2, 0.5]], visibleFields: [] }
  assert.throws(() => projection.projectChartRecords(records, chart('scatter', { valueFields: [1], seriesToCategory: { 5: [0] } })), /Y 字段/)
  assert.throws(() => projection.projectChartRecords(records, chart('line', { options: { errorBars: [{ seriesIndex: 1, direction: 'y', type: 'plus', valueType: 'cust', plusIndex: 3 }] } })), /误差字段/)
  assert.doesNotThrow(() => projection.projectChartRecords(records, chart('line', { options: { errorBars: [{ seriesIndex: 1, direction: 'y', type: 'plus', valueType: 'cust', plusIndex: 2, minusIndex: 99 }] } })))
})

test('settings preview preserves explicit raw null and unrounded values', () => {
  const draft = { config: chart('line'), sourceText: 'A1:B3' }, sheet = { id: 'sheet', name: 'Raw', cells: {
    A1: { value: 'Category' }, B1: { value: 'Value' }, A2: { rawValue: 'A', value: 'A' }, B2: { rawValue: null, value: '100' },
    A3: { rawValue: 'B', value: 'B' }, B3: { rawValue: 1.234567, value: '1.23' },
  } }
  const result = settings.projectChartSettings(draft, sheet)
  assert.equal(result.error, undefined)
  assert.deepEqual(result.source.records, [['A', null], ['B', 1.234567]])
  assert.deepEqual(result.data.series[0].values, [null, 1.234567])
})

test('settings preview and the visibility switch share the historical visible-only default in both source orientations', () => {
  const fields = ['Category', 'Value', 'Other'], records = [['A', 10, 100], ['B', 20, 200], ['C', 30, 300]]
  const field = controls.chartCommonControls().find(control => control.key === 'plotVisibleOnly')
  for (const orientation of ['columns', 'rows']) {
    const matrix = orientation === 'columns' ? [fields, ...records] : fields.map((name, index) => [name, ...records.map(record => record[index])])
    const cells = Object.fromEntries(matrix.flatMap((row, rowIndex) => row.map((value, columnIndex) => [`${drafts.chartColumnName(columnIndex)}${rowIndex + 1}`, { rawValue: value }])))
    const sheet = { id: 'sheet', name: 'Visibility', rowCount: matrix.length, columnCount: matrix[0].length, cells, hiddenRows: [2], hiddenColumns: [2] }
    const range = { startRow: 0, startCol: 0, endRow: matrix.length - 1, endCol: matrix[0].length - 1 }
    const config = chart('line', { source: { orientation, range } }), draft = { config, sourceText: orientation === 'columns' ? 'A1:C4' : 'A1:D3' }
    for (const authored of [undefined, {}, { plotVisibleOnly: undefined }, { plotVisibleOnly: true }]) {
      config.options = authored
      assert.equal(field.read(config), true, orientation)
      const preview = settings.projectChartSettings(draft, sheet)
      assert.equal(preview.error, undefined, orientation)
      assert.deepEqual(preview.source.fields, fields)
      assert.deepEqual(preview.source.visibleFields, [0, 1])
      assert.deepEqual(preview.source.records, [records[0], records[2]])
      assert.deepEqual(preview.data.categories, ['A', 'C'])
      assert.deepEqual(preview.data.series[0].values, [10, 30])
    }
    field.write(config, false)
    assert.equal(field.read(config), false)
    const full = settings.projectChartSettings(draft, sheet)
    assert.equal(full.error, undefined)
    assert.equal(full.source.visibleFields, undefined)
    assert.deepEqual(full.source.records, records)
    assert.deepEqual(full.data.categories, ['A', 'B', 'C'])
    field.write(config, true)
    assert.deepEqual(settings.projectChartSettings(draft, sheet).data.categories, ['A', 'C'])
  }
})

test('type conversion infers actual numeric roles and preserves independent depth and authored appearance', () => {
  const records = { fields: ['Category', 'X', 'Y', 'Size'], records: [['A', 1, 10, 4], ['B', 2, 20, 8]] }
  const model = chart('column', { valueFields: [2], options: { appearance: { background: 'navy' }, depthAxis: { title: 'Region' } } })
  const bubble = conversion.convertChartType(model, 'bubble', records)
  assert.deepEqual(bubble.options.mapping.bubble, { xIndex: 1, yIndex: 2, sizeIndex: 3, labelIndex: 0 })
  assert.deepEqual(bubble.options.depthAxis, model.options.depthAxis)
  assert.equal(bubble.options.appearance.background, 'navy')
  const scatter = conversion.convertChartType(bubble, 'scatter', records)
  assert.equal(scatter.categoryField, 1)
  assert.ok(scatter.valueFields.includes(2) && !scatter.valueFields.includes(1))
})

test('create and update consume captured records and explicit role patches override conversion defaults atomically', () => {
  const source = { fields: ['Category', 'X', 'Y', 'Size'], records: [['A', 1, 10, 4], ['B', 2, 20, 8]] }
  const sheet = { id: 'sheet', name: 'Host', rowCount: 5, columnCount: 4,
    get cells() { throw new Error('Captured records must not be reread from host cells') } }
  const workbook = { id: 'workbook', sheets: [sheet] }
  const created = engine.createChart(workbook, 'sheet', chart('column', { valueFields: [2], options: { appearance: { background: 'navy' } } }), 'captured', source)
  const mapping = { xIndex: 2, yIndex: 3, sizeIndex: 1, labelIndex: 0 }
  engine.updateChart(workbook, created, { type: 'bubble', categoryField: 1, valueFields: [3], options: { mapping: { bubble: mapping } } }, source)
  assert.equal(created.categoryField, 1)
  assert.deepEqual(created.valueFields, [3])
  assert.deepEqual(created.options.mapping.bubble, mapping)
  assert.equal(created.options.appearance.background, 'navy')
  const projected = engine.readChartData(workbook, created, source)
  assert.deepEqual(projected.series[0].xValues, [10, 20])
  assert.deepEqual(projected.series[0].values, [4, 8])
  assert.deepEqual(projected.series[0].sizeValues, [1, 2])
  const before = structuredClone(created)
  assert.throws(() => engine.updateChart(workbook, created, { options: { mapping: { bubble: { ...mapping, xIndex: 99 } } } }, source))
  assert.deepEqual(created, before)
})

test('line markers allow a point shape to override none and resolve marker colors and opacity in both adapters', () => {
  const source = ordinary(), model = chart('line', { options: { allSeriesStyle: { point: { shape: 'none', color: 'green' } },
    series: { 1: { dataPoints: { 1: { shape: 'diamond', size: 12, color: 'magenta', fillOpacity: 0.4 } } } } } })
  const native = option(source, model).series[0]
  assert.equal(native.showSymbol, true)
  assert.equal(native.symbol, 'none')
  assert.equal(native.data[1].symbol, 'diamond')
  assert.equal(native.data[1].symbolSize, 12)
  assert.equal(color.normalizeChartColor(native.data[1].itemStyle.color), 'rgba(255, 0, 255, 1)')
  assert.equal(native.data[1].itemStyle.opacity, 0.4)
  assert.equal(commands(source, model).filter(command => command.kind === 'polygon' && command.fill === 'rgba(255, 0, 255, 0.4)').length, 2)
})

test('bubble point shapes, sizes and colors share the same semantics while zero weights remain hidden', () => {
  const source = { categories: ['A', 'B', 'C'], series: [{ field: 2, name: 'Y', values: [10, 20, 30], xValues: [1, 2, 3], sizeValues: [4, 8, 0] }] }
  const model = chart('bubble', { categoryField: 1, valueFields: [2], options: { series: { 2: { point: { shape: 'diamond', color: 'green' },
    dataPoints: { 0: { shape: 'rect', size: 13, color: 'lime', fillOpacity: 0.5 }, 1: { shape: 'none' }, 2: { shape: 'rect', size: 20 } } } } } })
  const native = option(source, model).series[0]
  assert.equal(native.symbol, 'diamond')
  assert.equal(native.data.length, 2)
  assert.equal(native.data[0].symbol, 'rect')
  assert.equal(native.data[0].symbolSize, 13)
  assert.equal(color.normalizeChartColor(native.data[0].itemStyle.color), 'rgba(0, 255, 0, 1)')
  assert.equal(native.data[1].symbol, 'none')
  const markers = commands(source, model).filter(command => command.kind === 'rect' && command.clip && command.fill === 'rgba(0, 255, 0, 0.5)')
  assert.equal(markers.length, 1)
  assert.equal(markers[0].width, 13)
  assert.equal(derived.chartBubblePointSize({ point: { size: 20 } }, 0, 8, model.options, 200), 0)
  assert.equal(derived.chartBubblePointSize({}, 4, 8, undefined, -10), 0)
  assert.equal(derived.chartBubblePointSize({ point: { size: 20 } }, 4, 8, undefined, -10), 0)
})

test('radar point overrides and positioned value labels use the shared scene in both adapters', () => {
  const source = ordinary(), model = chart('radar', { options: { series: { 1: {
    point: { shape: 'none', color: 'green' }, label: { visible: true, position: 'right', color: 'navy' },
    dataPoints: { 1: { shape: 'triangle', size: 12, color: 'magenta', fillOpacity: 0.4, border: { color: 'lime', width: 3 } } },
  } } } })
  const geometry = layout.layoutDerivedChartData(prepared(source, model), model), scene = commands(source, model)
  const markers = scene.filter(command => command.kind === 'polygon' && command.fill === 'rgba(255, 0, 255, 0.4)')
  assert.equal(markers.length, 1)
  assert.ok(scene.some(command => command.kind === 'line' && command.color === 'rgba(0, 255, 0, 1)' && command.width === 3))
  const label = scene.find(command => command.kind === 'text' && command.text === '20')
  assert.equal(label.align, 'left')
  near(label.x, geometry.series[0].points[1].x + 9)
  near(label.y, geometry.series[0].points[1].y)
  const native = option(source, model)
  const elements = native.graphic.elements.map(element => element.type === 'group' ? element.children[0] : element)
  assert.ok(elements.some(element => element.type === 'polygon' && element.style.fill === 'rgba(255, 0, 255, 0.4)'))
  assert.ok(elements.some(element => element.type === 'text' && element.style.text === '20'))
})

test('large boxplot outlier and waterfall result collections retain their complete axis extents', () => {
  const count = 140_000, boxModel = chart('boxplot')
  const box = { data: { categories: ['Samples'], series: [{ field: 1, name: 'Samples', values: [0] }] }, trends: [], errors: [],
    boxes: [{ field: 1, seriesPosition: 0, name: 'Samples', statistics: { min: 0, q1: 0, median: 0, q3: 0, max: 0, outliers: Array(count).fill(10), outlierSourceIndices: Array.from({ length: count }, (_value, index) => index) } }] }
  assert.equal(layout.layoutDerivedChartData(box, boxModel).max, 10)
  assert.equal(layout.layoutDerivedChartData(box, { ...boxModel, options: { boxplot: { showOutliers: false } } }).max, 1)
  const categories = Array.from({ length: count }, (_value, index) => String(index))
  // Fixed zero-change prefix and one final 140000 change. No test-side
  // accumulation is used to generate the authoritative result fixture.
  const values = Array(count).fill(0); values[count - 1] = 140_000
  const points = categories.map((_category, categoryIndex) => ({ seriesPosition: 0, dataPointIndex: categoryIndex,
    categoryIndex, start: 0, end: 0, value: 0, role: 'change', synthetic: false }))
  Object.assign(points[count - 1], { end: 140_000, value: 140_000 })
  const waterfall = { data: { categories, series: [{ field: 1, name: 'Changes', values }] }, trends: [], errors: [],
    waterfallLayout: { stackType: 'arrangement', useSubtotal: false, sourceCategoryCount: count, seriesFields: [1], cumulative: 140_000,
      points, connectors: categories.slice(1).map((_category, index) => ({ from: { categoryIndex: index, value: 0 }, to: { categoryIndex: index + 1, value: 0 } })) } }
  const geometry = layout.layoutDerivedChartData(waterfall, chart('waterfall', { options: { waterfall: { useSubtotal: false } } }))
  assert.equal(geometry.max, count)
  assert.equal(geometry.series[0].points.length, count)
  assert.equal(geometry.series[0].points.at(-1).baseValue, 0)
  assert.equal(geometry.series[0].points.at(-1).value, 140_000)
})

test('horizontal bar secondary series use a separate numeric X axis with matching Canvas labels and extents', () => {
  const source = { categories: ['A', 'B'], series: [{ field: 1, name: 'Small', values: [10, 20] }, { field: 2, name: 'Large', values: [100, 200] }] }
  const model = chart('bar-horizontal', { valueFields: [1, 2], yAxis: { scaleType: 'linear', min: 0, max: 20 },
    options: { rightYAxis: { scaleType: 'linear', min: 0, max: 200, title: 'Secondary', label: { color: 'magenta', rotate: 15 }, tick: { visible: true, color: 'navy', length: 7 } }, series: { 2: { rightAxis: true } } } })
  const geometry = layout.layoutDerivedChartData(prepared(source, model), model), native = option(source, model)
  assert.equal(native.xAxis.length, 2)
  assert.equal(native.xAxis[1].position, 'top')
  assert.equal(native.xAxis[1].max, 200)
  assert.equal(native.yAxis.type, 'category')
  assert.equal(native.yAxis.inverse, true)
  assert.equal(native.series[0].xAxisIndex, 0)
  assert.equal(native.series[1].xAxisIndex, 1)
  assert.equal(native.series[1].yAxisIndex, 0)
  near(geometry.series[0].points[0].x, geometry.series[1].points[0].x)
  assert.ok(geometry.plot.y >= 52)
  const scene = commands(source, model), title = scene.find(command => command.kind === 'text' && command.text === 'Secondary')
  assert.ok(title.y < geometry.plot.y)
  assert.equal(title.rotation, 0)
  const label = scene.find(command => command.kind === 'text' && command.text === '200')
  assert.ok(label.y < geometry.plot.y)
  assert.equal(label.rotation, 15)
  assert.ok(scene.some(command => command.kind === 'line' && command.color === 'rgba(0, 0, 128, 1)' && command.y2 < command.y1))
})

test('stacked combinations and secondary axes reject logarithmic stacking, and finite-range overflow fails explicitly', () => {
  const source = { fields: ['Category', 'A', 'B'], records: [['One', 10, 20], ['Two', 15, 25]] }, workbook = { id: 'workbook', sheets: [{ id: 'sheet', name: 'Sheet', cells: {} }] }
  assert.throws(() => engine.createChart(workbook, 'sheet', chart('combination', { valueFields: [1, 2], options: { rightYAxis: { scaleType: 'log' },
    series: { 2: { type: 'columnStacked', rightAxis: true } } } }), 'invalid', source), /堆积系列须使用线性数值轴/)
  const overflowing = { categories: ['One'], series: [{ field: 1, name: 'A', values: [Number.MAX_VALUE] }, { field: 2, name: 'B', values: [Number.MAX_VALUE] }] }
  assert.throws(() => layout.layoutDerivedChartData(prepared(overflowing, chart('stacked-column')), chart('stacked-column')), /堆积合计/)
})

test('grouped Bubble series sharing Y produce separate trends and custom error intervals for their actual X vectors', () => {
  const source = { fields: ['X', 'Y', 'Error'], records: [[1, 10, 1], [2, 20, 2], [3, 100, 3], [4, 200, 4]], categories: ['A1', 'A2', 'B1', 'B2'],
    series: [{ field: 1, name: 'A', values: [10, 20, null, null], xValues: [1, 2, null, null], sizeValues: [1, 1, null, null] },
      { field: 1, name: 'B', values: [null, null, 100, 200], xValues: [null, null, 3, 4], sizeValues: [null, null, 1, 1] }] }
  const model = chart('bubble', { options: { trendlines: [{ seriesIndex: 1, type: 'linear' }],
    errorBars: [{ seriesIndex: 1, direction: 'y', type: 'plus', valueType: 'cust', plusIndex: 2 }] } })
  const result = prepared(source, model)
  assert.deepEqual(result.trends.map(trend => [trend.field, trend.seriesPosition]), [[1, 0], [1, 1]])
  near(result.trends[0].statistics.points[0][1], 10)
  near(result.trends[0].statistics.points.at(-1)[1], 20)
  near(result.trends[1].statistics.points[0][1], 100)
  near(result.trends[1].statistics.points.at(-1)[1], 200)
  assert.deepEqual(result.errors.map(error => [error.seriesPosition, error.intervals.map(interval => interval.categoryIndex)]), [[0, [0, 1]], [1, [2, 3]]])
  assert.deepEqual(result.errors[1].intervals.map(interval => [interval.x, interval.plus]), [[3, 3], [4, 4]])
})

test('precomputed grouped statistics bind by config slot and source position in derived, layout and graphic drawing', () => {
  const source = { categories: ['A1', 'A2', 'B1', 'B2'], series: [
    { field: 1, name: 'A', values: [10, 20, null, null], xValues: [1, 2, null, null], sizeValues: [1, 1, null, null] },
    { field: 1, name: 'B', values: [null, null, 100, 200], xValues: [null, null, 3, 4], sizeValues: [null, null, 1, 1] },
  ] }
  const model = chart('bubble', { options: { trendlines: [{ seriesIndex: 1, type: 'linear', color: 'magenta' }],
    errorBars: [{ seriesIndex: 1, direction: 'y', type: 'plus', valueType: 'cust', plusIndex: 2, noEndCap: true, line: { color: 'lime' } }] } })
  const firstTrend = { points: [[1, 11], [2, 21]] }, secondTrend = { points: [[3, 101], [4, 201]] }
  const firstIntervals = [{ categoryIndex: 0, x: 1, y: 10, minus: 0, plus: 5 }, { categoryIndex: 1, x: 2, y: 20, minus: 0, plus: 5 }]
  const secondIntervals = [{ categoryIndex: 2, x: 3, y: 100, minus: 0, plus: 7 }, { categoryIndex: 3, x: 4, y: 200, minus: 0, plus: 7 }]
  const computed = { preparedData: source,
    trendlines: [{ configIndex: 0, seriesPosition: 1, statistics: secondTrend }, { configIndex: 0, seriesPosition: 0, statistics: firstTrend }],
    errorBars: [{ configIndex: 0, seriesPosition: 1, intervals: secondIntervals }, { configIndex: 0, seriesPosition: 0, intervals: firstIntervals }],
  }
  const result = derived.deriveChartData(source, model, computed)
  assert.equal(result.trends[0].statistics, firstTrend)
  assert.equal(result.trends[1].statistics, secondTrend)
  assert.equal(result.errors[0].intervals, firstIntervals)
  assert.equal(result.errors[1].intervals, secondIntervals)
  const geometry = layout.layoutDerivedChartData(result, model)
  assert.equal(geometry.max, 207)
  const scene = draw.drawDerivedChartData(result, model), errors = scene.filter(command => command.kind === 'line' && command.color === 'rgba(0, 255, 0, 1)')
  assert.equal(errors.length, 4)
  near(errors[2].x1, geometry.series[1].points[0].x)
  near(errors[2].y1, geometry.series[1].points[0].y)
  assert.ok(errors[2].y2 < errors[2].y1)
  const native = echarts.chartEChartsOption({ sheets: [] }, model, {}, computed)
  const elements = native.graphic.elements.map(element => element.type === 'group' ? element.children[0] : element)
  assert.equal(elements.filter(element => element.type === 'line' && element.style.stroke === 'rgba(0, 255, 0, 1)').length, 4)
  assert.throws(() => derived.deriveChartData(source, model, { ...computed, trendlines: computed.trendlines.slice(0, 1) }), /配置和系列位置/)
  assert.throws(() => derived.deriveChartData(source, model, { ...computed, errorBars: [] }), /配置和系列位置/)
})

test('boxplot output order and waterfall source identities keep repeated fields distinct', () => {
  const source = { categories: ['One'], series: [{ field: 1, name: 'A', values: [1] }, { field: 1, name: 'B', values: [10] }] }
  const first = { min: 1, q1: 1, median: 1, q3: 1, max: 1, outliers: [], outlierSourceIndices: [] }, second = { min: 10, q1: 10, median: 10, q3: 10, max: 10, outliers: [], outlierSourceIndices: [] }
  const box = derived.deriveChartData(source, chart('boxplot'), { preparedData: source, boxplots: [
    { field: 1, seriesPosition: 1, statistics: second }, { field: 1, seriesPosition: 0, statistics: first },
  ] })
  assert.deepEqual(box.boxes.map(item => item.seriesPosition), [0, 1])
  assert.equal(box.boxes[0].statistics, first)
  assert.equal(box.boxes[1].statistics, second)
  const firstPoint = { seriesPosition: 0, dataPointIndex: 0, categoryIndex: 0, start: 0, end: 1, value: 1, role: 'change', synthetic: false }
  const secondPoint = { seriesPosition: 1, dataPointIndex: 0, categoryIndex: 1, start: 1, end: 11, value: 10, role: 'change', synthetic: false }
  const waterfallLayout = { stackType: 'arrangement', useSubtotal: false, sourceCategoryCount: 1, seriesFields: [1, 1], cumulative: 11,
    points: [firstPoint, secondPoint], connectors: [{ from: { categoryIndex: 0, value: 1 }, to: { categoryIndex: 1, value: 1 } }] }
  const waterfall = derived.deriveChartData(source, chart('waterfall', { options: { waterfall: { useSubtotal: false } } }), { preparedData: source, waterfallLayout })
  assert.equal(waterfall.waterfallLayout, waterfallLayout)
  assert.equal(waterfall.waterfalls[0][0], firstPoint)
  assert.equal(waterfall.waterfalls[1][0], secondPoint)
  assert.equal(waterfall.waterfalls[1][0].start, 1)
})

test('public compound-pie anchors use actual sectors and secondary bar centers, including the aggregate item', () => {
  const source = ordinary(), model = chart('barOfPie', { options: { pie: { secondaryPlot: { split: { type: 'position', count: 2 } } } } })
  const geometry = layout.layoutDerivedChartData(prepared(source, model), model), partitions = pie.chartPieLayout(source, model, geometry.plot)
  for (const sector of partitions.sectors) {
    const point = geometry.series[0].points.find(point => point.categoryIndex === sector.part.index), radius = (sector.inner + sector.outer) / 2, angle = (sector.start + sector.end) / 2
    near(point.x, sector.center.x + Math.cos(angle) * radius); near(point.y, sector.center.y + Math.sin(angle) * radius)
  }
  for (const bar of partitions.bars) {
    const point = geometry.series[0].points.find(point => point.categoryIndex === bar.part.index)
    near(point.x, bar.rect.x + bar.rect.width / 2); near(point.y, bar.rect.y + bar.rect.height / 2)
  }
  assert.equal(geometry.categories[3], '其他')
  assert.equal(geometry.series[0].points.find(point => point.categoryIndex === 3).value, 50)
})

test('funnel sorting retains source point indices and ends at a finite zero-width tip', () => {
  const source = ordinary(), model = chart('funnel'), geometry = layout.layoutDerivedChartData(prepared(source, model), model), steps = pie.chartFunnelLayout(source, model, geometry.plot)
  assert.deepEqual(steps.map(step => step.categoryIndex), [2, 1, 0])
  assert.deepEqual(geometry.series[0].points.map(point => point.categoryIndex), [2, 1, 0])
  for (const [index, step] of steps.entries()) { near(geometry.series[0].points[index].x, step.center.x); near(geometry.series[0].points[index].y, step.center.y) }
  near(steps.at(-1).points[2].x, steps.at(-1).points[3].x)
})

test('gauge public anchors and native tile centers share the Canvas layout', () => {
  const source = ordinary(), model = chart('gauge', { options: { gauge: { min: 0, max: 100, mode: 'progress' } } })
  const geometry = layout.layoutDerivedChartData(prepared(source, model), model), tiles = pie.chartGaugeTiles(source, geometry.plot), native = option(source, model).series.filter(series => series.id.endsWith(':foreground'))
  for (const tile of tiles) {
    assert.deepEqual(native[tile.categoryIndex].center, [tile.center.x, tile.center.y])
    const angle = -(225 + (-45 - 225) * source.series[0].values[tile.categoryIndex] / 100) * Math.PI / 180, radius = tile.radius * 0.95
    const point = geometry.series[0].points[tile.categoryIndex]
    near(point.x, tile.center.x + Math.cos(angle) * radius); near(point.y, tile.center.y + Math.sin(angle) * radius)
  }
})

function gaugeData(value = 50, name = 'Observation') {
  return { fields: ['Name', 'Value'], records: [[name, value]], categories: [name], series: [{ field: 1, name: 'Value', values: [Number(value)] }] }
}
function gaugeParts(result, categoryIndex = 0) {
  const suffix = `:gauge:${categoryIndex}:`
  return Object.fromEntries((result.series ?? []).filter(series => series.id.includes(suffix)).map(series => [series.id.split(':').at(-1), series]))
}

/** Future validation oracle: native GaugeView displayables, using SVG SSR only in tests. */
function nativeGaugeDraw(source, model, inspect) {
  useNativeChart([GaugeChart, SVGRenderer])
  const instance = initNativeChart(null, undefined, { renderer: 'svg', ssr: true, width: model.size.width, height: model.size.height })
  try {
    instance.setOption({ animation: false, series: option(source, model).series }, { notMerge: true })
    assert.ok(instance.renderToSVGString().startsWith('<svg'))
    return inspect(instance, instance.getZr().storage.getDisplayList(true))
  } finally { instance.dispose() }
}

test('native GaugeView keeps independent annular widths, foreground order and the original value text', () => {
  const source = gaugeData('125.00'), model = chart('gauge', { options: { gauge: { mode: 'classic',
    axisLine: { width: 24, color: 'orange' }, progress: { width: 10, rounded: true, color: 'lime' },
    axisTick: { color: 'red' }, splitLine: { color: 'navy', length: 10 }, axisLabel: { distance: 4 },
    pointer: { color: 'blue', length: 0.6 }, anchor: { color: 'white' } } } })
  const geometry = layout.layoutDerivedChartData(prepared(source, model), model), tile = pie.chartGaugeTiles(source, geometry.plot)[0]
  nativeGaugeDraw(source, model, (instance, list) => {
    const findFill = fill => list.find(element => typeof element.style?.fill === 'string' && color.normalizeChartColor(element.style.fill) === fill)
    const axis = findFill('rgba(255, 165, 0, 1)'), progress = findFill('rgba(0, 255, 0, 1)'), anchor = findFill('rgba(255, 255, 255, 1)')
    assert.equal(axis.type, 'sector')
    assert.equal(progress.type, 'sausage')
    nearNative(axis.shape.r - axis.shape.r0, 24)
    nearNative(progress.shape.r - progress.shape.r0, 10)
    nearNative(progress.shape.r, tile.radius)
    const foreground = instance.getModel().getSeries().find(series => series.id.endsWith(':foreground'))
    assert.equal(foreground.getData().get('value', 0), 100)
    const pointer = foreground.getData().getItemGraphicEl(0), tip = pointer.transformCoordToGlobal(0, -tile.radius * 0.6)
    nearNative(tip[0], geometry.series[0].points[0].x); nearNative(tip[1], geometry.series[0].points[0].y)
    const ticks = list.filter(element => element.type === 'line' && color.normalizeChartColor(element.style.stroke) === 'rgba(255, 0, 0, 1)')
    assert.ok(ticks.length > 0)
    for (const element of [...ticks, anchor, pointer]) assert.ok(list.indexOf(progress) < list.indexOf(element))
    assert.ok(list.indexOf(axis) < list.indexOf(progress))
    const valueText = list.find(element => element.type === 'tspan' && element.style.text === '125.00')
    assert.ok(valueText)
    assert.ok(list.indexOf(pointer) < list.indexOf(valueText))
    const label = instance.getViewOfSeriesModel(foreground).group.children().find(element => element.type === 'text' && element.style.text === '0')
    const radius = tile.radius - 10 - 4, angle = -225 * Math.PI / 180
    nearNative(label.style.x, tile.center.x + radius * Math.cos(angle)); nearNative(label.style.y, tile.center.y + radius * Math.sin(angle))
    assert.deepEqual([label.style.align, label.style.verticalAlign], ['left', 'middle'])
  })
})

test('native Gauge pointer transforms match public tips for all profiles and authored angle directions', () => {
  for (const shape of ['needle', 'slim', 'arrow']) for (const [startAngle, endAngle] of [[225, -45], [-45, 225], [0, 360], [360, 0], [-360, 90], [-360, 360], [0, 0]]) {
    const source = gaugeData(25), model = chart('gauge', { options: { gauge: { startAngle, endAngle,
      axisLine: { visible: false }, axisTick: { visible: false }, splitLine: { visible: false }, axisLabel: { visible: false },
      progress: { visible: false }, title: { visible: false }, detail: { visible: false }, pointer: { shape, width: 6, length: 0.6 } } } })
    const geometry = layout.layoutDerivedChartData(prepared(source, model), model), tile = pie.chartGaugeTiles(source, geometry.plot)[0]
    nativeGaugeDraw(source, model, (instance, list) => {
      const foreground = instance.getModel().getSeries().find(series => series.id.endsWith(':foreground'))
      const pointer = foreground.getData().getItemGraphicEl(0), tip = pointer.transformCoordToGlobal(0, -tile.radius * 0.6), rect = pointer.getBoundingRect()
      assert.ok(list.includes(pointer))
      nearNative(rect.width, 6); nearNative(rect.height, tile.radius * 0.6)
      nearNative(tip[0], geometry.series[0].points[0].x); nearNative(tip[1], geometry.series[0].points[0].y)
    })
  }
})

test('modern, classic and historical Gauge modes preserve their distinct scale and indicator behavior', () => {
  const modern = gauge.resolveChartGaugeOptions(undefined, 100), classic = gauge.resolveChartGaugeOptions({ mode: 'classic' }, 100)
  assert.deepEqual([modern.min, modern.max, modern.startAngle, modern.endAngle, modern.splitNumber], [0, 100, 225, -45, 5])
  assert.deepEqual([modern.axisLine.width, modern.axisTick.visible, modern.splitLine.visible, modern.axisLabel.visible, modern.anchor.visible, modern.progress.rounded], [18, false, false, false, false, true])
  assert.deepEqual([classic.axisLine.width, classic.axisTick.visible, classic.axisTick.splitNumber, classic.splitLine.visible, classic.axisLabel.visible, classic.anchor.visible, classic.progress.rounded], [8, true, 5, true, true, true, false])
  for (const mode of ['modern', 'classic', 'pointer', 'progress']) {
    const model = chart('gauge', { options: { gauge: { mode } } }), source = gaugeData(), parts = gaugeParts(option(source, model))
    assert.equal(parts.foreground.pointer.show, mode !== 'progress')
    assert.equal(Boolean(parts.progress), mode !== 'pointer')
    finiteCommands(commands(source, model))
  }
  const pointer = gauge.resolveChartGaugeOptions({ mode: 'pointer' }, 100), progress = gauge.resolveChartGaugeOptions({ mode: 'progress' }, 100)
  assert.deepEqual([pointer.axisLine.width, pointer.pointer.length, pointer.anchor.visible, pointer.title.offsetY, pointer.detail.offsetY], [10, 0.7, true, 0.56, 0.3])
  assert.deepEqual([progress.pointer.visible, progress.progress.visible, progress.axisLabel.visible], [false, true, true])
})

test('all five public Gauge presets replace previous appearance and retain independent source and chart options', () => {
  assert.deepEqual(gauge.CHART_GAUGE_PRESETS.map(preset => preset.value), ['modern', 'classic', 'statusDial', 'progress', 'minimal'])
  const current = { mapping: { gauge: { valueIndex: 1, labelIndex: 0 } }, appearance: { background: 'navy' }, gauge: {
    min: -100, max: -10, splitNumber: 7, ranges: [{ from: -90, to: -30, color: 'red' }], label: { visible: false, format: '0.0' },
    axisLine: { color: 'red', width: 42 }, axisLabel: { color: 'red', family: 'Serif', format: '0.0000' },
    pointer: { shape: 'arrow', color: 'red', length: 0.9 }, progress: { color: 'red', width: 42 },
    title: { family: 'Serif', color: 'red', offsetY: 0.1 }, detail: { format: '0.0000', color: 'red', offsetY: 0.1 },
  } }
  for (const preset of gauge.CHART_GAUGE_PRESETS) {
    const result = options.mergeChartOptions(current, { gauge: gaugeApi.patchChartGaugeOptions(current.gauge, gauge.chartGaugePresetOptions(preset.value)) })
    assert.deepEqual(result.mapping, current.mapping)
    assert.deepEqual(result.appearance, current.appearance)
    assert.deepEqual([result.gauge.min, result.gauge.max, result.gauge.pointer.shape], [0, 100, 'needle'])
    assert.equal(result.gauge.ranges, undefined)
    assert.equal(result.gauge.label, undefined)
    for (const group of ['axisLine', 'axisLabel', 'pointer', 'progress', 'title', 'detail']) assert.equal(result.gauge[group].color, undefined)
    assert.equal(result.gauge.axisLabel.format, undefined)
    assert.equal(result.gauge.detail.format, undefined)
    const model = chart('gauge', { options: result }), parts = gaugeParts(option(gaugeData(), model))
    assert.equal(parts.foreground.pointer.show, preset.value !== 'progress')
    assert.equal(Boolean(parts.progress), preset.value !== 'minimal')
    finiteCommands(commands(gaugeData(), model))
  }
  const status = gauge.chartGaugePresetOptions('statusDial')
  assert.deepEqual([status.splitNumber, status.axisLine.width, status.splitLine.visible, status.axisLabel.visible, status.anchor.visible], [2, 10, true, true, true])
  assert.equal(gauge.chartGaugePresetOptions('progress').progress.width, 14)
  assert.throws(() => gauge.chartGaugePresetOptions('unknown'), TypeError)
  assert.equal(current.gauge.min, -100)
})

test('Gauge invalid observations compact tiles while preserving source indices, point colors and legend entries', () => {
  const records = { fields: ['Name', 'Value'], records: [['First', 10], ['Empty', null], ['Third', 20], ['Text', 'bad'], ['Infinite', Infinity]] }
  const model = chart('gauge', { legend: { visible: true, position: 'bottom' }, options: { invalidValueStrategy: 'zero', gauge: { axisLine: { visible: false }, progress: { visible: false } },
    series: { 1: { dataPoints: { 2: { color: 'magenta' } } } } } })
  const source = projection.projectChartRecords(records, model), normalized = derived.chartComputeInputs(source, model)
  assert.deepEqual(normalized.series[0].values, [10, null, 20, null, null])
  const geometry = layout.layoutDerivedChartData(prepared(normalized, model), model), tiles = pie.chartGaugeTiles(normalized, geometry.plot)
  assert.deepEqual(tiles.map(tile => tile.categoryIndex), [0, 2])
  assert.deepEqual(geometry.series[0].points.map(point => point.categoryIndex), [0, 2])
  const compact = pie.chartGaugeTiles({ categories: ['First', 'Third'], series: [{ field: 1, name: 'Value', values: [10, 20] }] }, geometry.plot)
  assert.deepEqual(tiles.map(tile => [tile.center, tile.radius]), compact.map(tile => [tile.center, tile.radius]))
  const native = option(normalized, model), parts = gaugeParts(native, 2)
  assert.deepEqual(native.legend.data, ['First', 'Third'])
  assert.equal(color.normalizeChartColor(parts.foreground.pointer.itemStyle.color), 'rgba(255, 0, 255, 1)')
  const scene = commands(normalized, model)
  assert.ok(scene.some(command => command.kind === 'polygon' && command.fill === 'rgba(255, 0, 255, 1)'))
  assert.ok(!scene.some(command => command.kind === 'text' && ['Empty', 'Text', 'Infinite'].includes(command.text)))
  assert.equal(source.records[1][1], null)
  assert.equal(records.records[4][1], Infinity)
})

test('Gauge with only invalid observations uses the shared empty state without fabricating zero indicators', () => {
  const source = { categories: ['Empty', 'Invalid'], series: [{ field: 1, name: 'Value', values: [null, NaN] }] }, model = chart('gauge')
  const geometry = layout.layoutDerivedChartData(prepared(source, model), model), scene = commands(source, model), native = option(source, model)
  assert.equal(geometry.series[0].points.length, 0)
  assert.ok(scene.some(command => command.kind === 'text' && command.text === '暂无可绘制数据'))
  assert.ok(native.graphic.elements.some(element => element.type === 'text' && element.style.text === '暂无可绘制数据'))
  assert.equal(native.series, undefined)
  finiteCommands(scene)
})

test('Gauge retains original numeric-string text while clamping only its indicator geometry and native value', () => {
  const records = { fields: ['Name', 'Value'], records: [['High', '125.00'], ['Low', '-25.00'], ['Exponent', '5.20e1']] }, model = chart('gauge')
  const source = projection.projectChartRecords(records, model), geometry = layout.layoutDerivedChartData(prepared(source, model), model), native = option(source, model)
  const tiles = pie.chartGaugeTiles(source, geometry.plot), ratios = [1, 0, 0.52]
  for (const [index, tile] of tiles.entries()) {
    const parts = gaugeParts(native, index), angle = (-225 + 270 * ratios[index]) * Math.PI / 180, point = geometry.series[0].points[index]
    assert.equal(parts.foreground.data[0].value, [100, 0, 52][index])
    assert.equal(parts.foreground.detail.formatter(), records.records[index][1])
    near(point.x, tile.center.x + tile.radius * 0.62 * Math.cos(angle))
    near(point.y, tile.center.y + tile.radius * 0.62 * Math.sin(angle))
    assert.equal(point.value, [125, -25, 52][index])
    assert.ok(commands(source, model).some(command => command.kind === 'text' && command.text === records.records[index][1]))
  }
  const formatted = chart('gauge', { options: { gauge: { detail: { format: '0.00%' } } } }), decimal = gaugeData(0.526)
  assert.equal(gaugeParts(option(decimal, formatted)).foreground.detail.formatter(), '52.60%')
  assert.ok(commands(decimal, formatted).some(command => command.kind === 'text' && command.text === '52.60%'))
})

test('Gauge clockwise, zero, full-circle and reverse-over-one-circle angles follow the native arc contract', () => {
  for (const [startAngle, endAngle, degrees] of [[225, -45, 270], [-45, 225, 90], [0, 360, 360], [360, 0, 360], [-360, 90, 270], [-360, 360, 360], [0, 0, 0]]) {
    const sweep = gauge.chartGaugeSweep({ startAngle, endAngle })
    near(sweep.end - sweep.start, degrees * Math.PI / 180)
    const source = gaugeData(), model = chart('gauge', { options: { gauge: { startAngle, endAngle, pointer: { length: 0.5 } } } })
    const geometry = layout.layoutDerivedChartData(prepared(source, model), model), tile = pie.chartGaugeTiles(source, geometry.plot)[0], point = geometry.series[0].points[0]
    const angle = (-startAngle + degrees / 2) * Math.PI / 180
    near(point.x, tile.center.x + tile.radius * 0.5 * Math.cos(angle)); near(point.y, tile.center.y + tile.radius * 0.5 * Math.sin(angle))
    const parts = gaugeParts(option(source, model))
    assert.deepEqual([parts.foreground.startAngle, parts.foreground.endAngle], [startAngle, endAngle])
    assert.equal(Boolean(parts.progress), degrees > 0)
    finiteCommands(commands(source, model))
  }
})

test('zero-width or zero-length Gauge indicators and anchors disappear from both render paths', () => {
  const base = { axisLine: { visible: false }, axisTick: { visible: false }, splitLine: { visible: false }, axisLabel: { visible: false },
    anchor: { visible: true, size: 0 }, title: { visible: false }, detail: { visible: false } }
  for (const [value, indicator] of [[50, { pointer: { width: 0 }, progress: { visible: false } }], [50, { pointer: { length: 0 }, progress: { visible: false } }],
    [50, { pointer: { visible: false }, progress: { width: 0 } }], [0, { pointer: { visible: false } }], [50, { pointer: { visible: false }, startAngle: 0, endAngle: 0 }]]) {
    const source = gaugeData(value), model = chart('gauge', { options: { gauge: { ...base, ...indicator } } }), geometry = layout.layoutDerivedChartData(prepared(source, model), model)
    const tile = pie.chartGaugeTiles(source, geometry.plot)[0], point = geometry.series[0].points[0], parts = gaugeParts(option(source, model))
    near(point.x, tile.center.x); near(point.y, tile.center.y)
    assert.equal(parts.foreground.pointer.show, false)
    assert.equal(parts.foreground.anchor.show, false)
    assert.equal(parts.progress, undefined)
    assert.ok(!commands(source, model).some(command => command.kind === 'polygon'))
  }
})

test('Gauge negative-domain partial patches, null resets and unsafe objects preserve atomic option updates', () => {
  const current = { min: -100, max: -10, pointer: { color: 'blue', width: 4 }, detail: { format: '0.000' } }
  const patch = gaugeApi.patchChartGaugeOptions(current, { max: -5, pointer: { width: 9 } })
  assert.deepEqual([patch.min, patch.max, patch.pointer.color, patch.pointer.width], [-100, -5, 'blue', 9])
  const reset = gaugeApi.patchChartGaugeOptions(patch, gaugeApi.chartGaugeNullableValues({ min: null, max: null, splitNumber: null }))
  assert.deepEqual([reset.min, reset.max, reset.splitNumber], [undefined, undefined, undefined])
  assert.deepEqual([gauge.resolveChartGaugeOptions(reset, 100).min, gauge.resolveChartGaugeOptions(reset, 100).max], [0, 100])
  assert.deepEqual(gaugeApi.chartGaugeTicksOptions({ majorDivisions: 4, minorDivisionsPerMajor: 3 }), { splitNumber: 4, axisTick: { splitNumber: 3 } })
  assert.deepEqual(gaugeApi.chartGaugeTicksOptions({ majorDivisions: null, minorDivisionsPerMajor: null }), { splitNumber: undefined, axisTick: { splitNumber: undefined } })
  class Unsafe { width = 8 }
  for (const value of [new Unsafe(), Object.create(null), JSON.parse('{"__proto__":{}}'), { constructor: 'unsafe' }, { prototype: 'unsafe' }]) {
    assert.throws(() => gaugeApi.patchChartGaugeOptions(current, value), TypeError)
    assert.throws(() => gaugeApi.patchChartGaugeOptions(current, { pointer: value }), TypeError)
    assert.throws(() => gaugeApi.chartGaugeNullableValues(value), TypeError)
  }
  assert.throws(() => gaugeApi.chartGaugeRangesOptions(Array(1)), TypeError)
  assert.throws(() => gaugeApi.patchChartGaugeOptions(current, { ranges: Array(1) }), TypeError)
  assert.throws(() => options.cloneChartOptions({ gauge: { ranges: Array(1) } }), TypeError)
  assert.throws(() => gaugeApi.patchChartGaugeOptions(current, { max: -200 }), RangeError)
  assert.deepEqual(current, { min: -100, max: -10, pointer: { color: 'blue', width: 4 }, detail: { format: '0.000' } })
})

test('Gauge per-field resets recover inherited fonts, visibility and number format in both adapters', () => {
  const model = chart('gauge', { options: { appearance: { font: { family: 'Georgia', fontSize: 19, color: 'navy', bold: true, italic: true } },
    allSeriesStyle: { label: { visible: false, format: '0.0%' } }, gauge: { label: { visible: undefined, format: undefined },
      axisLabel: { visible: true, family: undefined, bold: undefined }, title: { family: undefined, visible: undefined },
      detail: { visible: true, family: undefined, format: undefined, bold: undefined } } } })
  const source = gaugeData(0.52), parts = gaugeParts(option(source, model)), scene = commands(source, model)
  assert.equal(parts.foreground.title.show, false)
  assert.equal(parts.foreground.detail.formatter(), '52.0%')
  for (const group of ['axisLabel', 'detail']) {
    assert.equal(parts.foreground[group].fontFamily, 'Georgia')
    assert.equal(parts.foreground[group].fontSize, 19)
    assert.equal(parts.foreground[group].fontWeight, 'bold')
    assert.equal(parts.foreground[group].fontStyle, 'italic')
  }
  const detail = scene.find(command => command.kind === 'text' && command.text === '52.0%')
  assert.equal(detail.font, 'italic 600 19px Georgia')
  assert.equal(detail.color, 'rgba(0, 0, 128, 1)')
  assert.ok(!scene.some(command => command.kind === 'text' && command.text === 'Observation'))
})

test('Gauge inherited series labels use the authored series name in both render paths', () => {
  const model = chart('gauge', { options: { series: { 1: { name: 'Authored metric', label: { contentType: 'series' } } } } }), source = gaugeData()
  assert.equal(gaugeParts(option(source, model)).foreground.detail.formatter(), 'Authored metric')
  assert.ok(commands(source, model).some(command => command.kind === 'text' && command.text === 'Authored metric'))
})

test('Gauge controls drive axis, ticks, pointer, anchor, progress, text, layout and native layer settings', () => {
  const model = chart('gauge'), fields = new Map(gaugeControls.chartGaugeSettingsControls(model).map(control => [control.key, control]))
  const authored = { 'gauge.mode': 'classic', 'gauge.splitNumber': 4, 'gauge.axisLine.width': 24, 'gauge.axisLine.color': 'orange',
    'gauge.axisTick.splitNumber': 3, 'gauge.axisTick.length': 7, 'gauge.axisTick.width': 2, 'gauge.axisTick.color': 'red',
    'gauge.splitLine.length': 10, 'gauge.splitLine.width': 3, 'gauge.splitLine.color': 'navy',
    'gauge.axisLabel.distance': 4, 'gauge.axisLabel.format': '0.0', 'gauge.axisLabel.family': 'Georgia', 'gauge.axisLabel.fontSize': 16,
    'gauge.axisLabel.bold': true, 'gauge.axisLabel.italic': true, 'gauge.axisLabel.color': 'purple',
    'gauge.pointer.length': 0.5, 'gauge.pointer.width': 6, 'gauge.pointer.shape': 'arrow', 'gauge.pointer.color': 'blue',
    'gauge.anchor.visible': true, 'gauge.anchor.size': 12, 'gauge.anchor.color': 'white', 'gauge.anchor.borderColor': 'blue', 'gauge.anchor.borderWidth': 2,
    'gauge.progress.width': 10, 'gauge.progress.rounded': true, 'gauge.progress.color': 'lime',
    'gauge.title.offsetY': 0.8, 'gauge.title.color': 'green', 'gauge.detail.offsetY': 0.2, 'gauge.detail.format': '0.0%', 'gauge.detail.fontSize': 18 }
  for (const [key, value] of Object.entries(authored)) { assert.ok(fields.has(key), key); fields.get(key).write(model, value) }
  const source = gaugeData(0.52), scene = commands(source, model), geometry = layout.layoutDerivedChartData(prepared(source, model), model)
  const tile = pie.chartGaugeTiles(source, geometry.plot)[0], parts = gaugeParts(option(source, model)), foreground = parts.foreground
  assert.deepEqual([parts.axis.z, parts.progress.z, foreground.z], [1, 2, 3])
  assert.equal(parts.progress.progress.overlap, false)
  assert.equal(parts.axis.axisLine.lineStyle.width, 24)
  assert.equal(parts.progress.axisLine.lineStyle.width, 10)
  assert.equal(parts.progress.progress.width, 10)
  assert.equal(foreground.axisLine.lineStyle.width, 24)
  assert.equal(foreground.progress.show, false)
  assert.deepEqual([foreground.axisTick.show, foreground.axisTick.splitNumber, foreground.axisTick.length, foreground.axisTick.lineStyle.width], [true, 3, 7, 2])
  assert.deepEqual([foreground.splitLine.show, foreground.splitLine.length, foreground.splitLine.lineStyle.width], [true, 10, 3])
  assert.equal(scene.filter(command => command.kind === 'line' && command.color === 'rgba(255, 0, 0, 1)').length, 16)
  assert.equal(scene.filter(command => command.kind === 'line' && command.color === 'rgba(0, 0, 128, 1)').length, 5)
  const firstTick = scene.find(command => command.kind === 'text' && command.text === '0.0')
  const labelRadius = tile.radius - 10 - 4, angle = -225 * Math.PI / 180
  near(firstTick.x, tile.center.x + labelRadius * Math.cos(angle)); near(firstTick.y, tile.center.y + labelRadius * Math.sin(angle))
  assert.deepEqual([firstTick.align, firstTick.baseline, firstTick.font], ['left', 'middle', 'italic 600 16px Georgia'])
  assert.equal(foreground.axisLabel.formatter(25), '25.0')
  assert.equal(foreground.pointer.keepAspect, false)
  assert.equal(foreground.pointer.length, '50%')
  assert.equal(foreground.pointer.width, 6)
  assert.equal(foreground.anchor.size, 12)
  assert.equal(foreground.anchor.itemStyle.borderWidth, 2)
  assert.equal(parts.progress.progress.roundCap, true)
  const title = scene.find(command => command.kind === 'text' && command.text === 'Observation'), detail = scene.find(command => command.kind === 'text' && command.text === '52.0%')
  near(title.y, tile.center.y + tile.radius * 0.8); near(detail.y, tile.center.y + tile.radius * 0.2)
  assert.equal(title.color, 'rgba(0, 128, 0, 1)')
  assert.equal(detail.font, '600 18px Aptos, Arial, sans-serif')
  assert.equal(foreground.detail.formatter(), '52.0%')
  assert.equal(foreground.title.offsetCenter[1], '80%')
  assert.equal(foreground.detail.offsetCenter[1], '20%')
  finiteCommands(scene)
})

test('Gauge settings resolve implicit single-field roles and preserve explicit label omission when editing fields', () => {
  const source = { fields: ['Metric'], records: [['0.25']] }
  for (const orientation of ['columns', 'rows']) {
    const range = { startRow: 0, startCol: 0, endRow: orientation === 'columns' ? 1 : 0, endCol: orientation === 'columns' ? 0 : 1 }
    const model = chart('gauge', { source: { orientation, range }, valueFields: [], options: { series: { 0: { label: { format: '0.0%' } } } } })
    const fields = new Map(typeControls.chartDataMappingControls(model, 1).map(control => [control.key, control]))
    assert.equal(fields.get('mapping.gauge.valueIndex').read(model), 0)
    assert.equal(fields.get('mapping.gauge.labelIndex').read(model), undefined)
    assert.equal(gaugeControls.chartGaugeSettingsControls(model, 1).find(control => control.key === 'gauge.detail.format').read(model), '0.0%')
    fields.get('mapping.gauge.labelIndex').write(model, undefined)
    fields.get('mapping.gauge.valueIndex').write(model, undefined)
    assert.deepEqual(model.valueFields, [0])
    assert.deepEqual(projection.projectChartRecords(source, model).categories, ['Metric'])
    assert.equal(gaugeParts(option(projection.projectChartRecords(source, model), model)).foreground.detail.formatter(), '25.0%')
  }
  const oldRange = chart('gauge', { valueFields: [], options: { series: { 0: { label: { format: '0.00' } } } } })
  assert.equal(gaugeControls.chartGaugeSettingsControls(oldRange, 1).find(control => control.key === 'gauge.detail.format').read(oldRange), '0.00')
  const records = { fields: ['Name', 'First', 'Second'], records: [['A', 10, 0.5], ['B', 20, 0.75]] }
  const model = chart('gauge', { options: { mapping: { gauge: { valueIndex: 1 } }, series: { 2: { label: { format: '0.0%' } } } } })
  const fields = new Map(typeControls.chartDataMappingControls(model, 3).map(control => [control.key, control]))
  assert.equal(fields.get('mapping.gauge.labelIndex').read(model), undefined)
  fields.get('mapping.gauge.valueIndex').write(model, 2)
  assert.deepEqual(model.valueFields, [2])
  assert.equal(model.options.mapping.gauge.labelIndex, undefined)
  assert.deepEqual(projection.projectChartRecords(records, model).categories, ['Second', 'Second'])
  assert.equal(gaugeControls.chartGaugeSettingsControls(model, 3).find(control => control.key === 'gauge.detail.format').read(model), '0.0%')
  fields.get('mapping.gauge.labelIndex').write(model, 0)
  assert.deepEqual(projection.projectChartRecords(records, model).categories, ['A', 'B'])
  fields.get('mapping.gauge.labelIndex').write(model, undefined)
  assert.equal(fields.get('mapping.gauge.labelIndex').read(model), undefined)
  assert.deepEqual(projection.projectChartRecords(records, model).categories, ['Second', 'Second'])
})

test('Gauge draft preset and reset actions reach model updates while retaining mapping and other chart configuration', () => {
  const records = { fields: ['Name', 'Value'], records: [['A', 10], ['B', 20], ['C', 30]] }, workbook = { id: 'workbook', sheets: [{ id: 'sheet', name: 'Host', cells: {}, rowCount: 4, columnCount: 2 }] }
  const model = engine.createChart(workbook, 'sheet', chart('gauge', { options: { mapping: { gauge: { valueIndex: 1, labelIndex: 0 } }, appearance: { background: 'navy' }, gauge: { mode: 'classic' } } }), 'gauge-draft', records)
  const draft = drafts.createChartSettingsDraft(model)
  gaugeControls.applyChartGaugePreset(draft.config, 'progress')
  engine.updateChart(workbook, model, drafts.chartSettingsPatch(draft), records)
  assert.equal(gaugeParts(option(projection.projectChartRecords(records, model), model)).foreground.pointer.show, false)
  gaugeControls.resetChartGaugeSettings(draft.config)
  const cloned = drafts.cloneChartDraftConfig(draft.config), patch = drafts.chartSettingsPatch({ ...draft, config: cloned })
  assert.ok(Object.hasOwn(patch.options, 'gauge'))
  assert.equal(patch.options.gauge, undefined)
  engine.updateChart(workbook, model, patch, records)
  assert.equal(model.options.gauge, undefined)
  assert.deepEqual(model.options.mapping.gauge, { valueIndex: 1, labelIndex: 0 })
  assert.equal(model.options.appearance.background, 'navy')
  const parts = gaugeParts(option(projection.projectChartRecords(records, model), model))
  assert.equal(parts.foreground.pointer.show, true)
  assert.equal(parts.progress.progress.width, 18)
})

test('Gauge ranges leave uncovered scale portions in the axis color and pointer profiles retain their real tips', () => {
  const rangeOptions = { min: -10, max: 90, ranges: [{ from: 0, to: 30, color: 'red' }, { from: 40, to: 90, color: 'blue' }] }
  assert.deepEqual(gauge.chartGaugeColorStops(rangeOptions, 'gray'), [[0.1, 'gray'], [0.4, 'red'], [0.5, 'gray'], [1, 'blue']])
  for (const shape of ['needle', 'slim', 'arrow']) {
    const model = chart('gauge', { options: { gauge: { axisLine: { visible: false }, progress: { visible: false }, pointer: { shape, length: 0.6, width: 6, color: 'red' } } } })
    const source = gaugeData(25), geometry = layout.layoutDerivedChartData(prepared(source, model), model), tile = pie.chartGaugeTiles(source, geometry.plot)[0]
    const angle = (-225 + 270 / 4) * Math.PI / 180, tip = { x: tile.center.x + tile.radius * 0.6 * Math.cos(angle), y: tile.center.y + tile.radius * 0.6 * Math.sin(angle) }
    const vertices = commands(source, model).filter(command => command.kind === 'polygon' && command.fill === 'rgba(255, 0, 0, 1)').flatMap(command => command.points)
    assert.ok(vertices.some(point => Math.hypot(point.x - tip.x, point.y - tip.y) < 1e-8), shape)
    const pointer = gaugeParts(option(source, model)).foreground.pointer
    assert.ok(pointer.icon.startsWith('path://'))
    assert.deepEqual([pointer.keepAspect, pointer.length, pointer.width], [false, '60%', 6])
  }
})

test('horizontal series controls bind the top value axis and retain its authored axis title', () => {
  const model = chart('bar', { options: { rightYAxis: { scaleType: 'linear', title: 'Top values' } } })
  const field = controls.chartSeriesControls(undefined, false, true, true).find(control => control.key === 'allSeriesStyle.rightAxis')
  assert.equal(field.label, '使用顶部数值轴')
  field.write(model, true)
  const native = option(ordinary(), model)
  assert.equal(native.xAxis[1].name, undefined)
  assert.ok(native.graphic.elements.some(element => (element.type === 'text' ? element : element.children?.[0])?.style.text === 'Top values'))
  const geometry = layout.layoutDerivedChartData(prepared(ordinary(), model), model)
  assert.equal(geometry.series[0].rightAxis, true)
  assert.ok(commands(ordinary(), model).some(command => command.kind === 'text' && command.text === 'Top values'))
})

test('heatmap color ramps remain finite at extreme bounds and both adapters consume point overrides', () => {
  const source = { categories: ['Low', 'Mid', 'High'], series: [{ field: 1, name: 'Row', values: [-Number.MAX_VALUE, 0, Number.MAX_VALUE] }] }
  const model = chart('heatmap', { palette: ['black', 'white'], options: { series: { 1: { dataPoints: { 1: { color: 'magenta', fillOpacity: 0.4, border: { width: 3, color: 'navy' } } } } }, heatmap: { min: -Number.MAX_VALUE, max: Number.MAX_VALUE, label: { visible: true } } } })
  const extent = derived.chartHeatmapExtent(source, model)
  assert.equal(color.normalizeChartColor(derived.chartHeatmapPointStyle(model, 1, 0, 0, -Number.MAX_VALUE, extent, model.palette).color), 'rgba(0, 0, 0, 1)')
  assert.equal(option(source, model).series, undefined, 'extreme color extent uses the shared finite graphic scene')
  const native = option(source, { ...model, options: { ...model.options, heatmap: { ...model.options.heatmap, min: -10, max: 10 } } }).series[0].data[1]
  assert.equal(native.itemStyle.color, 'rgba(255, 0, 255, 0.4)')
  assert.equal(native.itemStyle.borderWidth, 3)
  const scene = commands(source, model)
  finiteCommands(scene)
  assert.ok(scene.some(command => command.kind === 'rect' && command.fill === native.itemStyle.color && command.strokeWidth === 3))
})

test('heatmap public points, field headers and category row labels follow both reversed categorical axes', () => {
  const source = { categories: ['A', 'B'], series: [{ field: 1, name: 'First', values: [1, 2] }, { field: 2, name: 'Second', values: [3, 4] }] }
  const model = chart('heatmap', { xAxis: { scaleType: 'linear', reverse: true }, yAxis: { scaleType: 'linear', reverse: true } }), geometry = layout.layoutDerivedChartData(prepared(source, model), model)
  near(geometry.series[0].points[0].x, geometry.plot.x + geometry.plot.width * 0.75)
  near(geometry.series[0].points[0].y, geometry.plot.y + geometry.plot.height * 0.25)
  const label = commands(source, model).find(command => command.kind === 'text' && command.text === 'First')
  near(label.x, geometry.series[0].points[0].x)
  assert.ok(label.y > geometry.plot.y + geometry.plot.height)
  const category = commands(source, model).find(command => command.kind === 'text' && command.text === 'A')
  near(category.y, geometry.series[0].points[0].y)
})

test('Pareto exposes and draws its cumulative series with a matching Canvas legend entry', () => {
  const source = ordinary(), model = chart('pareto', { legend: { visible: true, position: 'bottom' } }), geometry = layout.layoutDerivedChartData(prepared(source, model), model)
  assert.equal(geometry.series.length, 2)
  assert.equal(geometry.series[1].name, '累计百分比')
  assert.equal(geometry.series[1].rightAxis, true)
  assert.equal(geometry.series[1].points.at(-1).value, 100)
  assert.ok(commands(source, model).some(command => command.kind === 'text' && command.text === '累计百分比'))
})

test('waterfall right-axis values use their own extent and reversed connectors respect bar spacing', () => {
  const source = { categories: ['A', 'B', 'C'], series: [{ field: 1, name: 'Small', values: [10, -5, 20] }, { field: 2, name: 'Large', values: [100, -50, 200] }] }
  const model = chart('waterfall', { valueFields: [1, 2], xAxis: { scaleType: 'linear', reverse: true }, yAxis: { scaleType: 'linear', min: 0, max: 30 },
    options: { rightYAxis: { scaleType: 'linear', min: 0, max: 300 }, series: { 2: { rightAxis: true } }, bar: { gapWidth: 0 },
      waterfall: { useSubtotal: false, connector: { color: '#fedcba', dashType: 'solid' } } } })
  const computed = { preparedData: source, waterfallLayout: { stackType: 'arrangement', useSubtotal: false, sourceCategoryCount: 3, seriesFields: [1, 2], cumulative: 275,
    points: [
      { seriesPosition: 0, dataPointIndex: 0, categoryIndex: 0, start: 0, end: 10, value: 10, role: 'change', synthetic: false },
      { seriesPosition: 0, dataPointIndex: 1, categoryIndex: 1, start: 10, end: 5, value: -5, role: 'change', synthetic: false },
      { seriesPosition: 0, dataPointIndex: 2, categoryIndex: 2, start: 5, end: 25, value: 20, role: 'change', synthetic: false },
      { seriesPosition: 1, dataPointIndex: 0, categoryIndex: 3, start: 25, end: 125, value: 100, role: 'change', synthetic: false },
      { seriesPosition: 1, dataPointIndex: 1, categoryIndex: 4, start: 125, end: 75, value: -50, role: 'change', synthetic: false },
      { seriesPosition: 1, dataPointIndex: 2, categoryIndex: 5, start: 75, end: 275, value: 200, role: 'change', synthetic: false },
    ], connectors: [
      { from: { categoryIndex: 0, value: 10 }, to: { categoryIndex: 1, value: 10 } },
      { from: { categoryIndex: 1, value: 5 }, to: { categoryIndex: 2, value: 5 } },
      { from: { categoryIndex: 2, value: 25 }, to: { categoryIndex: 3, value: 25 } },
      { from: { categoryIndex: 3, value: 125 }, to: { categoryIndex: 4, value: 125 } },
      { from: { categoryIndex: 4, value: 75 }, to: { categoryIndex: 5, value: 75 } },
    ] } }
  const result = prepared(source, model, computed), geometry = layout.layoutDerivedChartData(result, model)
  assert.equal(geometry.series[1].points[0].baseValue, 25)
  near(geometry.series[0].points[2].y, geometry.plot.y + geometry.plot.height / 6)
  near(geometry.series[1].points[0].y, geometry.plot.y + geometry.plot.height * 7 / 12)
  assert.equal(geometry.rightMax, 300)
  const wide = draw.drawDerivedChartData(result, model), narrow = commands(source, { ...model, options: { ...model.options, bar: { gapWidth: 500 } } }, computed)
  const widths = scene => scene.filter(command => command.kind === 'rect' && command.clip).map(command => command.width)
  assert.ok(widths(wide)[0] > widths(narrow)[0])
  assert.ok(wide.some(command => command.kind === 'line' && command.color === 'rgba(254, 220, 186, 1)' && command.x2 < command.x1))
})

test('large independent weights produce finite pie, hierarchy and directed-link geometries', () => {
  const maximum = Number.MAX_VALUE, source = { categories: ['A', 'B'], series: [{ field: 1, name: 'Weight', values: [maximum, maximum] }] }
  for (const type of ['pie', 'pie3D']) finiteCommands(commands(source, chart(type)))
  const nodes = [{ name: 'A', value: maximum }, { name: 'B', value: maximum }], graph = { nodes, links: [{ source: 'A', target: 'B', value: maximum }], cyclic: false }
  assert.ok(special.chartTreemapLayout(nodes, { x: 0, y: 0, width: 300, height: 200 }).every(rect => [rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)))
  assert.ok(special.chartSunburstLayout(nodes, { x: 0, y: 0, width: 300, height: 200 }).every(sector => [sector.start, sector.end, sector.inner, sector.outer].every(Number.isFinite)))
  finiteCommands(commands({ ...source, graph }, chart('sankey')))
  finiteCommands(commands({ ...source, graph }, chart('chord')))
})

test('3D pie uses distinct category colors and preserves explicit fill opacity', () => {
  const source = { categories: ['A', 'B'], series: [{ field: 1, name: 'Weight', values: [10, 20] }] }, scene = commands(source, chart('pie3D', { options: { allSeriesStyle: { fillOpacity: 0.25 } } }))
  assert.ok(scene.some(command => command.kind === 'polygon' && command.fill === 'rgba(37, 99, 235, 0.25)'))
  assert.ok(scene.some(command => command.kind === 'polygon' && command.fill === 'rgba(220, 38, 38, 0.25)'))
})

test('3D depth reversal moves authored points and the scene consumes its independent title', () => {
  const source = { categories: ['A'], series: [{ field: 1, name: 'First', values: [10] }, { field: 2, name: 'Second', values: [10] }] }
  const model = chart('column3D', { valueFields: [1, 2], options: { depthAxis: { title: 'Region' } } }), reversed = { ...model, options: { depthAxis: { title: 'Region', reverse: true } } }
  const before = layout.layoutDerivedChartData(prepared(source, model), model), after = layout.layoutDerivedChartData(prepared(source, reversed), reversed)
  assert.ok(Math.abs(before.series[0].points[0].x - after.series[0].points[0].x) > 1e-6)
  assert.ok(commands(source, model).some(command => command.kind === 'text' && command.text === 'Region'))
})

test('an isolated 3D line sample draws its configured point marker', () => {
  const source = { categories: ['Only'], series: [{ field: 1, name: 'Value', values: [10] }] }
  const model = chart('line3D', { options: { series: { 1: { dataPoints: { 0: { color: 'lime', fillOpacity: 0.5, shape: 'diamond', size: 12 } } } } } })
  const geometry = layout.layoutDerivedChartData(prepared(source, model), model), scene = threeD.chart3DScene(source, model, geometry)
  assert.equal(scene.points[0].length, 1)
  assert.equal(scene.commands.filter(command => command.kind === 'polygon' && command.fill === 'rgba(0, 255, 0, 0.5)').length, 2)
})

test('3D horizontal bars place category labels left and numeric labels below their own axis', () => {
  const source = ordinary(), model = chart('bar3D', { options: { threeD: { view: { rotX: 0, rotY: 0, rightAngleAxes: true } } },
    xAxis: { scaleType: 'linear', title: 'Categories' }, yAxis: { scaleType: 'linear', title: 'Values' } })
  const geometry = layout.layoutDerivedChartData(prepared(source, model), model), scene = threeD.chart3DScene(source, model, geometry)
  const category = scene.commands.find(command => command.kind === 'text' && command.text === 'A')
  assert.equal(category.align, 'right')
  assert.ok(category.x < geometry.plot.x)
  const value = scene.commands.find(command => command.kind === 'text' && command.text === '0')
  assert.equal(value.align, 'center')
  assert.equal(value.baseline, 'top')
  assert.ok(value.y > geometry.plot.y + geometry.plot.height)
  const title = scene.commands.find(command => command.kind === 'text' && command.text === 'Values')
  assert.equal(title.rotation, 0)
})
