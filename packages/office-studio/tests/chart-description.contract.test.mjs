import assert from 'node:assert/strict'
import { test } from 'vitest'
import { CHART_TYPE_CATALOG } from '../src/features/charts/chart-catalog.ts'
import { describeChartConfiguration, chartDescriptionFieldIndexes } from '../src/features/charts/chart-description.ts'
import { chartLinearGradient } from '../src/features/charts/chart-color.ts'

// P16 drafts for the unified validation phase. This file has not been executed.
const fields = { 0: 'Year', 2: 'Revenue', 8: 'Cost', 20: 'Dormant' }
const configuration = (type = 'column', options = {}) => ({ type, categoryField: 0, valueFields: [2, 8], options })
const describe = (config, context = {}) => describeChartConfiguration(config, { id: 'chart-description', fields, ...context })

function freeze(value) {
  if (value && typeof value === 'object') { Object.values(value).forEach(freeze); Object.freeze(value) }
  return value
}

test('semantic descriptions retain authored gradients and inactive colors without derived paints or shared stop objects', () => {
  const color = chartLinearGradient({ start: { x: -0.5, y: 0 }, end: { x: 1.5, y: 1 }, stops: [
    { offset: 1, color: 'blue', opacity: 0.25 }, { offset: 0, color: 'red' },
  ] })
  const config = freeze(configuration('line', { autoGradientFill: false, allSeriesStyle: { color }, series: { 2: { dataPoints: { 0: { color } } } } })), before = structuredClone(config)
  const first = describe(config), second = describe(config)
  assert.deepEqual(first.series[2].color, color)
  assert.deepEqual(first.series[2].dataPoints[0].color, color)
  assert.equal(first.autoGradientFill, false)
  assert.deepEqual(first.series[2].color.stops.map(stop => stop.offset), [1, 0])
  first.series[2].color.start.x = 0
  first.series[2].color.stops[0].opacity = 1
  first.series[2].dataPoints[0].color.stops[0].color = 'purple'
  assert.deepEqual(config, before)
  assert.deepEqual(second.series[2].color, color)
  assert.equal(first.series[8].color.stops[0].opacity, 0.25)
  const serialized = JSON.stringify(second)
  for (const field of ['colorStops', 'authored', 'bounds', 'global']) assert.equal(serialized.includes(`"${field}"`), false)
})

test('descriptions detach authored plot, styles and axes while leaving an automatic plot unauthored', () => {
  const config = freeze({ ...configuration('area', {
    plotArea: { left: 0.2, top: 0.25, width: 0.6, height: 0.5 },
    rightYAxis: { scaleType: 'linear', title: 'Secondary', titleStyle: { color: 'blue' } },
    series: { 2: { color: 'red', dataPoints: { 0: { color: 'purple', border: { width: 2 } } } } },
  }), title: 'Sales', subtitle: 'Annual', palette: ['red', 'blue'], legend: { position: 'right', visible: true },
    xAxis: { scaleType: 'linear', titleStyle: { fontSize: 15 } }, yAxis: { scaleType: 'log', min: 1 } })
  const before = structuredClone(config), first = describe(config), second = describe(config)
  first.plotArea.left = 0.1
  first.palette[0] = 'green'
  first.legend.position = 'bottom'
  first.series[2].dataPoints[0].border.width = 8
  first.axes.x.titleStyle.fontSize = 20
  first.axes.rightY.titleStyle.color = 'black'
  assert.deepEqual(config, before)
  assert.equal(second.plotArea.left, 0.2)
  assert.equal(second.series[2].dataPoints[0].border.width, 2)
  assert.equal(first.rightYAxis.titleStyle.color, 'blue')
  assert.equal(second.axes.x.titleStyle.fontSize, 15)
  const automatic = describe(configuration('area'))
  assert.equal(automatic.autoGradientFill, true)
  assert.equal(Object.hasOwn(automatic, 'plotArea'), false)
})

test('all catalog descriptions use canonical names and effective gradient defaults without mutating config', () => {
  for (const entry of CHART_TYPE_CATALOG) {
    const config = configuration(entry.value), before = structuredClone(config), description = describe(config)
    assert.equal(description.type, entry.id)
    assert.equal(description.autoGradientFill, entry.family === 'area')
    const disabled = describe({ ...config, options: { autoGradientFill: false } })
    assert.equal(disabled.autoGradientFill, false)
    assert.deepEqual(config, before)
  }
  assert.equal(describe(configuration('bar')).type, 'column')
  assert.equal(describe(configuration('columnStacked')).type, 'columnStacked')
  assert.equal(describe(configuration('stacked-column')).type, 'columnStacked')
  assert.equal(describe(configuration('bar-horizontal')).type, 'bar')
  assert.equal(describe(configuration('doughnut')).type, 'donut')
})

test('source-indexed descriptions resolve shared, field and known-name inheritance and retain inactive styles', () => {
  const config = { ...configuration('column', {
    allSeriesStyle: { color: 'gray', fillOpacity: 0.4, border: { width: 3, color: 'red' }, label: { visible: true }, dataPoints: { 0: { color: 'red', fillOpacity: 0.25 } } },
    series: { 2: { name: 'Gross', border: { width: 1 }, rightAxis: true, dataPoints: { 0: { fillOpacity: 0.6 } } }, 20: { color: 'green' } },
    namedSeries: [{ name: 'Revenue', style: { name: 'Net', color: 'blue', dataPoints: { 0: { color: 'purple' } } } }],
  }), valueFields: [8, 2] }
  const before = structuredClone(config), description = describe(config)
  assert.deepEqual(Object.keys(description.series).map(Number), [2, 8, 20])
  assert.equal(description.series[2].name, 'Net')
  assert.equal(description.series[2].color, 'blue')
  assert.deepEqual(description.series[2].border, { width: 1, color: 'red' })
  assert.deepEqual(description.series[2].dataPoints[0], { color: 'purple', fillOpacity: 0.6, border: {} })
  assert.equal(description.series[8].name, 'Cost')
  assert.equal(description.series[20].name, 'Dormant')
  assert.equal(description.series[20].color, 'green')
  assert.equal(Array.isArray(description.series), false)
  assert.equal(description.series[0], undefined)
  assert.equal(description.series[1], undefined)
  assert.deepEqual(config, before)
  assert.deepEqual(describe({ ...config, valueFields: [2] }).series[2], description.series[2])
})

test('theme colors stay outside semantic descriptions while explicit palettes and colors survive', () => {
  const config = configuration(), description = describe(config)
  assert.deepEqual(description.series[2], { name: 'Revenue', type: 'column' })
  assert.equal(describe({ ...config, palette: [] }).series[2].color, undefined)
  const palette = describe({ ...config, valueFields: [8, 2], palette: ['red', 'blue'] })
  assert.equal(palette.series[8].color, 'red')
  assert.equal(palette.series[2].color, 'blue')
  assert.equal(describe(configuration('column', { series: { 2: { color: 'purple' } } })).series[2].color, 'purple')
})

test('description reads only requested header properties and never touches raw source, values or geometry', () => {
  const config = configuration('column', { series: { 20: { color: 'green' } } }), reads = []
  for (const key of ['source', 'records', 'values', 'cells', 'position', 'size']) Object.defineProperty(config, key, { enumerable: true, get() { throw new Error(`raw ${key} was read`) } })
  const headers = new Proxy(fields, {
    ownKeys() { throw new Error('whole header iteration') },
    get(target, key) { reads.push(key); if (!Object.hasOwn(target, key)) throw new Error(`unrequested field ${String(key)}`); return target[key] },
  })
  assert.deepEqual(chartDescriptionFieldIndexes(config), [0, 2, 8, 20])
  const description = describeChartConfiguration(config, { id: 'metadata-only', fields: headers })
  assert.deepEqual(new Set(reads), new Set(['0', '2', '8', '20']))
  assert.equal(description.category.name, 'Year')
  for (const key of ['source', 'records', 'values', 'cells', 'position', 'size', 'categories']) assert.equal(Object.hasOwn(description, key), false)
  for (const entry of Object.values(description.series)) for (const key of ['values', 'items', 'points', 'xValues', 'sizeValues']) assert.equal(Object.hasOwn(entry, key), false)
})

test('missing source identities keep named overrides detached instead of applying a fabricated default name', () => {
  const config = { ...configuration('column', { namedSeries: [{ name: 'Series 3', style: { color: 'purple' } }] }), valueFields: [2] }
  const unknown = describeChartConfiguration(config, { id: 'unknown-name' })
  assert.deepEqual(unknown.series[2], { name: 'Series 3', type: 'column' })
  assert.equal(unknown.category, undefined)
  assert.equal(unknown.namedSeries[0].style.color, 'purple')
  const knownEmptyHeader = describe(config, { fields: { 0: 'Year', 2: '' } })
  assert.equal(knownEmptyHeader.series[2].name, 'Series 3')
  assert.equal(knownEmptyHeader.series[2].color, 'purple')
  const emptyIdentity = describe(configuration('column', { namedSeries: [{ name: '', style: { color: 'green' } }] }), { seriesNames: { 2: '' } })
  assert.equal(emptyIdentity.series[2].name, '')
  assert.equal(emptyIdentity.series[2].color, 'green')
})

test('category metadata requires all configured headers or an explicit category name', () => {
  const config = configuration('column', { mapping: { categoryIndexes: [0, 5] } })
  assert.equal(describe(config).category, undefined)
  assert.deepEqual(describe(config, { fields: { ...fields, 5: 'Quarter' } }).category, { name: 'Year / Quarter' })
  assert.deepEqual(describe(config, { categoryName: '' }).category, { name: '' })
  assert.deepEqual(describe(config, { categoryName: 'Fiscal period' }).category, { name: 'Fiscal period' })
  assert.equal(describe(configuration('histogram')).category, undefined)
})

test('known source-role metadata can describe inferred roles without reading cells', () => {
  const config = { type: 'column', categoryField: 0, options: { series: { 20: { color: 'green' } } } }
  assert.deepEqual(Object.keys(describe(config).series).map(Number), [20])
  const resolved = describe(config, { seriesFields: [8, 2] })
  assert.deepEqual(Object.keys(resolved.series).map(Number), [2, 8, 20])
  assert.deepEqual(chartDescriptionFieldIndexes(config, [8, 2]), [0, 8, 2, 20])
  assert.equal(resolved.series[8].name, 'Cost')
  assert.equal(resolved.series[2].name, 'Revenue')
})

test('special mappings describe actual value roles instead of auxiliary fields as series', () => {
  const headers = { 0: 'X', 1: 'Y', 2: 'Size', 3: 'Label', 4: 'Source', 5: 'Target', 6: 'Weight', 7: 'Close', 8: 'High', 9: 'Low' }
  const bubble = describe({ type: 'bubble', options: { mapping: { bubble: { xIndex: 0, yIndex: 1, sizeIndex: 2, labelIndex: 3 } } } }, { fields: headers })
  assert.deepEqual(Object.keys(bubble.series).map(Number), [1])
  assert.deepEqual(bubble.category, { name: 'X' })
  const candle = describe({ type: 'candlestick', categoryField: 3, options: { mapping: { candlestick: { highIndex: 8, lowIndex: 9, closeIndex: 7 } } } }, { fields: headers })
  assert.deepEqual(Object.keys(candle.series).map(Number), [7])
  const edge = describe({ type: 'sankey', options: { mapping: { edge: { sourceIndex: 4, targetIndex: 5, valueIndex: 6 } } } }, { fields: headers })
  assert.deepEqual(Object.keys(edge.series).map(Number), [6])
  assert.deepEqual(edge.category, { name: 'Source / Target' })
  const gauge = describe({ type: 'gauge', options: { mapping: { gauge: { valueIndex: 6, labelIndex: 3 } } } }, { fields: headers })
  assert.deepEqual(Object.keys(gauge.series).map(Number), [6])
  assert.deepEqual(gauge.category, { name: 'Label' })
  assert.equal(gauge.axisBindings, undefined)
})

test('five public series types describe local combinations, aliases and stacking without derived series', () => {
  const expected = [['lineStacked', 'line'], ['columnPercentStacked', 'columnStacked'], ['barStacked', 'columnStacked'], ['areaPercentStacked3D', 'areaStacked'], ['waterfall', 'column']]
  for (const [type, seriesType] of expected) assert.equal(describe(configuration(type)).series[2].type, seriesType)
  const combination = describe({ ...configuration('combination', { series: { 8: { type: 'areaStacked' } } }), valueFields: [20, 2, 8] })
  assert.equal(combination.series[20].type, 'column')
  assert.equal(combination.series[2].type, 'line')
  assert.equal(combination.series[8].type, 'areaStacked')
  const pareto = describe(configuration('pareto'))
  assert.deepEqual(Object.keys(pareto.series).map(Number), [2])
  assert.equal(pareto.series[2].type, 'column')
  assert.equal(describe(configuration('pie')).series[2].type, undefined)
})

test('axis bindings follow active secondary styles, horizontal bars, heatmap rows and nonrectangular charts', () => {
  const options = { series: { 20: { rightAxis: true } }, rightYAxis: { scaleType: 'linear', title: 'Dormant' } }
  const inactive = describe(configuration('column', options))
  assert.equal(inactive.axisBindings.secondaryValue, undefined)
  assert.deepEqual(inactive.axisBindings.category, { dimension: 'horizontal', position: 'bottom' })
  assert.deepEqual(inactive.axisBindings.primaryValue, { dimension: 'vertical', position: 'left' })
  assert.equal(inactive.axes.rightY.title, 'Dormant')
  const active = describe(configuration('column', { allSeriesStyle: { rightAxis: true } }))
  assert.deepEqual(active.axisBindings.secondaryValue, { dimension: 'vertical', position: 'right' })
  const bar = describe(configuration('bar-horizontal', { series: { 2: { rightAxis: true } } }))
  assert.deepEqual(bar.axisBindings, { category: { dimension: 'vertical', position: 'left' }, primaryValue: { dimension: 'horizontal', position: 'bottom' }, secondaryValue: { dimension: 'horizontal', position: 'top' } })
  assert.deepEqual(describe(configuration('heatmap')).axisBindings, { category: { dimension: 'vertical', position: 'left' }, secondaryCategory: { dimension: 'horizontal', position: 'bottom' } })
  assert.deepEqual(describe(configuration('pareto')).axisBindings.secondaryValue, { dimension: 'vertical', position: 'right' })
  for (const type of ['radar', 'pie', 'doughnut', 'gauge', 'treemap', 'sankey']) assert.equal(describe(configuration(type)).axisBindings, undefined)
})

test('invalid roles and metadata fail explicitly without changing the authoring configuration', () => {
  const config = configuration(), before = structuredClone(config)
  for (const roles of [[-1], [1.5], [NaN], [Infinity], [1_048_576], ['1'], [{ name: 'one' }], Array(1), [1, 1]]) {
    assert.throws(() => describe(config, { seriesFields: roles }))
  }
  for (const id of ['', '  ', 0, null]) assert.throws(() => describe(config, { id }))
  assert.throws(() => describe(config, { fields: { 2: 42 } }), TypeError)
  assert.throws(() => describe(config, { fields: null }), TypeError)
  assert.throws(() => describe(config, { seriesNames: { 2: false } }), TypeError)
  assert.throws(() => describe(config, { categoryName: false }), TypeError)
  assert.throws(() => describe({ ...config, categoryField: '0' }), TypeError)
  assert.throws(() => describe({ ...config, type: 'pie', valueFields: Array(2) }), TypeError)
  assert.throws(() => describe({ ...config, type: 'pie', valueFields: [2, 2] }), RangeError)
  assert.throws(() => chartDescriptionFieldIndexes(config, [2, 2]), RangeError)
  assert.deepEqual(config, before)
})
