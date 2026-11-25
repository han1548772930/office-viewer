import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartCanvasElementAuthorSignature, chartCanvasElementCompatible, chartCanvasElementControls, chartCanvasElementDeleteDraft, chartCanvasElementHidden,
  chartCanvasElementPatch, resetChartCanvasElementDraft } from '../src/features/charts/chart-element-edit.ts'
import { chartCanvasParetoMark, chartCanvasSeriesBindings } from '../src/features/charts/chart-element-interaction.ts'
import { cloneChartDraftConfig } from '../src/features/charts/chart-settings-draft.ts'
import { chartPieAggregateStyle, chartPieConnector } from '../src/features/charts/chart-pie.ts'
import { chartParetoBarStyle, chartParetoCumulativeStyle } from '../src/features/charts/chart-pareto.ts'
import { chartWaterfallRoleStyle } from '../src/features/charts/chart-waterfall.ts'
import { chartLinearGradient } from '../src/features/charts/chart-color.ts'
import { cloneChartOptions, mergeChartOptions, validateChartOptions } from '../src/features/charts/chart-options.ts'

// P20 implementation drafts only. No runner, type/build check or browser was started.
const palette = ['#112233', '#223344', '#334455', '#445566', '#556677']
const json = value => JSON.parse(JSON.stringify(value))
const model = (type, options = {}) => ({ id: 'special-marks', sheetId: 'sheet', revision: 0, type,
  categoryField: 0, valueFields: [1, 2], source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 4, endCol: 2 } },
  position: { x: 0, y: 0 }, size: { width: 600, height: 420 }, palette, title: 'Chart',
  legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options })
const sources = () => ({ categories: ['A', 'B', 'C', 'D'], series: [
  { field: 1, name: 'Display Alpha', sourceName: ' Alpha ', values: [6, -2, 4, 1] },
  { field: 2, name: 'Display Beta', sourceName: 'Beta', values: [2, 3, -1, 4] },
] })
const mark = directEdit => ({ kind: 'mark', directEdit })
const bound = (chart, part = 'negative', position = 0, data = sources()) => mark({ kind: 'waterfallStyle', part,
  binding: chartCanvasSeriesBindings(data, chart.options)[position] })
const aggregate = () => mark({ kind: 'pieAggregateStyle', sourceField: 1, sourceName: ' Alpha ', aggregateIndex: 4 })
const connector = () => mark({ kind: 'pieConnectorStyle' })
const merged = (chart, draft, element) => mergeChartOptions(chart.options, chartCanvasElementPatch(chart, draft, element).options)

test('Pareto bar and cumulative-line edits use their own author groups and ignore unrelated draft fields and point maps', () => {
  const chart = model('pareto', { series: { 1: { color: '#111111' } }, pareto: {
    barStyle: { color: '#222222', dataPoints: { 0: { color: '#333333' } } },
    cumulativeLineStyle: { width: 2, dataPoints: { 1: { size: 12 } } },
  } }), before = json(chart), bar = chartCanvasParetoMark('bar', sources().series[0]), draft = cloneChartDraftConfig(chart)
  draft.options.pareto.barStyle.opacity = 0.4; draft.options.pareto.barStyle.dataPoints[0].color = '#ffffff'
  draft.options.pareto.cumulativeLineStyle.width = 9; draft.options.series[1].color = '#ffffff'; draft.title = 'Ignored'
  const patch = chartCanvasElementPatch(chart, draft, bar), options = mergeChartOptions(chart.options, patch.options)
  assert.deepEqual(Object.keys(patch), ['options']); assert.deepEqual(Object.keys(patch.options), ['pareto'])
  assert.equal(options.pareto.barStyle.opacity, 0.4); assert.deepEqual(options.pareto.barStyle.dataPoints, chart.options.pareto.barStyle.dataPoints)
  assert.deepEqual(options.pareto.cumulativeLineStyle, chart.options.pareto.cumulativeLineStyle); assert.deepEqual(options.series, chart.options.series)
  const line = chartCanvasParetoMark('line'), lineDraft = cloneChartDraftConfig(chart)
  lineDraft.options.pareto.cumulativeLineStyle.width = 4; lineDraft.options.pareto.barStyle.color = '#ffffff'
  const afterLine = merged(chart, lineDraft, line)
  assert.equal(afterLine.pareto.cumulativeLineStyle.width, 4); assert.deepEqual(afterLine.pareto.barStyle, chart.options.pareto.barStyle)
  assert.equal(chartParetoCumulativeStyle({ options: afterLine }, palette).border.width, 4); assert.deepEqual(json(chart), before)
})

test('Pareto format reset restores inherited appearance while retaining independent bar and line point overrides', () => {
  const chart = model('pareto', { series: { 1: { color: '#123456' } }, pareto: {
    barStyle: { color: '#654321', opacity: 0.6, border: { width: 3 }, label: { visible: true }, dataPoints: { 1: { color: '#abcdef' } } },
    cumulativeLineStyle: { width: 5, point: { size: 14 }, label: { visible: true }, dataPoints: { 2: { size: 16 } } },
  } })
  for (const [part, group] of [['bar', 'barStyle'], ['line', 'cumulativeLineStyle']]) {
    const element = chartCanvasParetoMark(part, sources().series[0]), next = merged(chart, resetChartCanvasElementDraft(chart, element), element)
    assert.deepEqual(next.pareto[group].dataPoints, chart.options.pareto[group].dataPoints)
    assert.equal(next.pareto[group].color, undefined); assert.equal(next.pareto[group].label, undefined)
    assert.deepEqual(next.series, chart.options.series)
    if (part === 'bar') assert.equal(chartParetoBarStyle({ options: next }, 1, 0, palette).color, '#123456')
    else assert.equal(chartParetoCumulativeStyle({ options: next }, palette).border.width, 2)
  }
})

test('a numeric Waterfall role patch preserves source fields, sibling roles, generic styles and independent role points', () => {
  const chart = model('waterfall', { waterfall: { pointRoles: [], connector: true }, series: {
    1: { name: 'Renamed', color: '#112233', waterfallStyles: { positive: { color: '#00aa00' },
      negative: { color: '#aa0000', dataPoints: { 1: { fillOpacity: 0.7 } } }, subtotal: { color: '#888888' } } }, 2: { color: '#223344' },
  } }), element = bound(chart), before = json(chart), draft = cloneChartDraftConfig(chart)
  draft.options.series[1].waterfallStyles.negative.fillOpacity = 0.5
  draft.options.series[1].waterfallStyles.negative.dataPoints[1].fillOpacity = 0.1
  draft.options.series[1].waterfallStyles.positive.color = '#ffffff'; draft.options.series[1].name = 'Ignored'
  draft.options.waterfall.pointRoles = [{ seriesIndex: 1, dataPointIndex: 1, role: 'total' }]
  const next = merged(chart, draft, element)
  assert.equal(next.series[1].waterfallStyles.negative.fillOpacity, 0.5)
  assert.deepEqual(next.series[1].waterfallStyles.negative.dataPoints, chart.options.series[1].waterfallStyles.negative.dataPoints)
  assert.deepEqual(next.series[1].waterfallStyles.positive, chart.options.series[1].waterfallStyles.positive)
  assert.deepEqual(next.series[1].waterfallStyles.subtotal, chart.options.series[1].waterfallStyles.subtotal)
  assert.deepEqual(next.series[2], chart.options.series[2]); assert.equal(next.series[1].name, 'Renamed')
  assert.deepEqual(next.waterfall, chart.options.waterfall); assert.deepEqual(json(chart), before)
  assert.equal(chartWaterfallRoleStyle({ options: next }, 1, 0, 'negative', palette, ' Alpha ').fillOpacity, 0.5)
})

test('Waterfall direct editing selects the exact-name author layer, and resets only the selected role appearance', () => {
  const chart = model('waterfall', { waterfall: { pointRoles: [] }, series: { 1: { waterfallStyles: { negative: { color: '#111111' } } } },
    namedSeries: [{ name: ' Alpha ', style: { waterfallStyles: { negative: { name: 'Loss', color: '#222222', dataPoints: { 1: { color: '#333333' } } }, positive: { color: '#444444' } } } },
      { name: 'Beta', style: { color: '#555555' } }] }), element = bound(chart), draft = cloneChartDraftConfig(chart)
  assert.deepEqual(element.directEdit.binding.selector, { name: ' Alpha ' })
  draft.options.namedSeries[0].style.waterfallStyles.negative.color = '#ee0000'
  const next = merged(chart, draft, element)
  assert.equal(next.namedSeries[0].style.waterfallStyles.negative.color, '#ee0000'); assert.deepEqual(next.series, chart.options.series)
  assert.deepEqual(next.namedSeries[1], chart.options.namedSeries[1])
  const reset = merged(chart, resetChartCanvasElementDraft(chart, element), element), negative = reset.namedSeries[0].style.waterfallStyles.negative
  assert.equal(negative.color, undefined); assert.equal(negative.name, 'Loss'); assert.deepEqual(negative.dataPoints, { 1: { color: '#333333' } })
  assert.deepEqual(reset.namedSeries[0].style.waterfallStyles.positive, chart.options.namedSeries[0].style.waterfallStyles.positive)
  assert.equal(chartWaterfallRoleStyle({ options: reset }, 1, 0, 'negative', palette, ' Alpha ').color, '#111111')
})

test('the synthetic stacked Waterfall subtotal writes global role style and cannot create an invented source selector', () => {
  const chart = model('waterfall', { waterfall: { stackType: 'stacked', useSubtotal: true, styles: { positive: { color: '#111111' }, subtotal: { color: '#222222', dataPoints: { 4: { color: '#333333' } } } } },
    series: { 1: { waterfallStyles: { subtotal: { color: '#444444' } } } } }), element = mark({ kind: 'waterfallStyle', part: 'subtotal' }), draft = cloneChartDraftConfig(chart)
  draft.options.waterfall.styles.subtotal.fillOpacity = 0.5; draft.options.series[1].waterfallStyles.subtotal.color = '#ffffff'
  const patch = chartCanvasElementPatch(chart, draft, element), next = mergeChartOptions(chart.options, patch.options)
  assert.deepEqual(Object.keys(patch.options), ['waterfall']); assert.equal(next.waterfall.styles.subtotal.fillOpacity, 0.5)
  assert.deepEqual(next.waterfall.styles.subtotal.dataPoints, chart.options.waterfall.styles.subtotal.dataPoints)
  assert.deepEqual(next.waterfall.styles.positive, chart.options.waterfall.styles.positive); assert.deepEqual(next.series, chart.options.series)
  assert.equal(chartCanvasElementDeleteDraft(chart, element), undefined)
  assert.throws(() => chartCanvasElementControls(element, model('waterfall', { waterfall: { stackType: 'arrangement' } })), /全局/)
  const disabled = model('waterfall', { waterfall: { stackType: 'stacked', useSubtotal: false } })
  assert.equal(chartCanvasElementCompatible(disabled, element), false)
  assert.throws(() => chartCanvasElementControls(element, disabled), /全局/)
  assert.throws(() => chartCanvasElementPatch(disabled, cloneChartDraftConfig(disabled), element), /全局/)
})

test('a compound Pie aggregate patch authors its independent style and label while preserving source points and split settings', () => {
  const chart = model('pieOfPie', { series: { 1: { dataPoints: { 1: { color: '#111111' }, 4: { color: '#222222' } } } },
    pie: { aggregateLabel: 'Others', aggregateStyle: { color: '#333333' }, secondaryPlot: { type: 'pie', size: 65, split: { type: 'position', count: 2 }, connector: { width: 2 } } } }),
    before = json(chart), element = aggregate(), draft = cloneChartDraftConfig(chart)
  draft.options.pie.aggregateStyle.color = '#ee0000'; draft.options.pie.aggregateStyle.explosion = 25; draft.options.pie.aggregateLabel = 'Remaining'
  draft.options.series[1].dataPoints[1].color = '#ffffff'; draft.options.pie.secondaryPlot.split.count = 1
  const patch = chartCanvasElementPatch(chart, draft, element), next = mergeChartOptions(chart.options, patch.options)
  assert.deepEqual(Object.keys(patch.options), ['pie']); assert.equal(next.pie.aggregateStyle.color, '#ee0000'); assert.equal(next.pie.aggregateStyle.explosion, 25)
  assert.equal(next.pie.aggregateLabel, 'Remaining'); assert.deepEqual(next.series, chart.options.series); assert.deepEqual(next.pie.secondaryPlot, chart.options.pie.secondaryPlot)
  assert.equal(next.pie.aggregateStyle.dataIndex, undefined); assert.equal(next.pie.aggregateStyle.aggregateIndex, undefined); assert.deepEqual(json(chart), before)
  const reset = merged(chart, resetChartCanvasElementDraft(chart, element), element)
  assert.equal(reset.pie.aggregateStyle, undefined); assert.equal(reset.pie.aggregateLabel, 'Others'); assert.deepEqual(reset.series, chart.options.series)
  assert.equal(chartPieAggregateStyle({ options: reset }, sources().series[0], 4, palette).color, '#222222')
  assert.equal(chartCanvasElementDeleteDraft(chart, element), undefined)
})

test('the Pie connector patch and Delete address both connection segments without replacing the split or aggregate authorship', () => {
  const chart = model('barOfPie', { pie: { aggregateLabel: 'Other', aggregateStyle: { fillOpacity: 0.6 }, secondaryPlot: {
    type: 'bar', size: 65, gapWidth: 70, split: { type: 'custom', pointIndexes: [2, 3] }, connector: { color: '#123456', width: 2 },
  } } }), element = connector(), draft = cloneChartDraftConfig(chart)
  draft.options.pie.secondaryPlot.connector.opacity = 0.4; draft.options.pie.secondaryPlot.connector.dashType = 'dotted'
  draft.options.pie.secondaryPlot.gapWidth = 10; draft.options.pie.aggregateStyle.fillOpacity = 0.1
  const next = merged(chart, draft, element)
  assert.deepEqual(next.pie.secondaryPlot.connector, { color: '#123456', width: 2, opacity: 0.4, dashType: 'dotted' })
  assert.equal(next.pie.secondaryPlot.gapWidth, 70); assert.deepEqual(next.pie.secondaryPlot.split, chart.options.pie.secondaryPlot.split)
  assert.deepEqual(next.pie.aggregateStyle, chart.options.pie.aggregateStyle)
  const hidden = chartCanvasElementDeleteDraft(chart, element); assert.ok(chartCanvasElementHidden(hidden, element))
  assert.equal(chartPieConnector(hidden, '#ffffff'), undefined); assert.deepEqual(hidden.source, chart.source)
  const reset = merged(chart, resetChartCanvasElementDraft(chart, element), element)
  assert.equal(reset.pie.secondaryPlot.connector, undefined); assert.equal(chartPieConnector({ options: reset }, '#abcdef').color, '#abcdef')
  assert.deepEqual(reset.pie.secondaryPlot.split, chart.options.pie.secondaryPlot.split)
})

test('special-mark signatures ignore outside scopes and no-op drafts, and stale chart types reject before patching', () => {
  const chart = model('pareto', { pareto: { barStyle: { dataPoints: { 0: { color: '#123456' } } } } }), bar = chartCanvasParetoMark('bar'), draft = cloneChartDraftConfig(chart)
  draft.options.pareto.barStyle.dataPoints[0].color = '#ffffff'; draft.title = 'Ignored'
  assert.equal(chartCanvasElementAuthorSignature(draft, bar), chartCanvasElementAuthorSignature(chart, bar))
  assert.deepEqual(chartCanvasElementPatch(chart, draft, bar), {})
  draft.type = 'line'; assert.throws(() => chartCanvasElementPatch(chart, draft, bar), /类型/)
  assert.equal(chartCanvasElementCompatible(draft, bar), false)
  assert.throws(() => chartCanvasElementControls(connector(), chart), /类型/)
  const empty = model('pieOfPie'), emptyDraft = cloneChartDraftConfig(empty); emptyDraft.options.pie = { aggregateStyle: {} }
  assert.equal(chartCanvasElementCompatible({ ...empty, type: 'pie' }, aggregate()), false)
  assert.deepEqual(chartCanvasElementPatch(empty, emptyDraft, aggregate()), {})
})

test('direct forms reuse effective full-panel controls, and their writes reach the intended materialized appearance', () => {
  const chart = model('pareto', { namedSeries: [{ name: ' Alpha ', style: { label: { visible: true } } }] }), bar = chartCanvasParetoMark('bar', sources().series[0])
  const controls = chartCanvasElementControls(bar, chart), draft = cloneChartDraftConfig(chart)
  assert.equal(controls.find(control => control.key === 'pareto.barStyle.label.visible').read(chart), true)
  controls.find(control => control.key === 'pareto.barStyle.opacity').write(draft, 0.4)
  assert.equal(chartParetoBarStyle({ options: merged(chart, draft, bar) }, 1, 0, palette, ' Alpha ').fillOpacity, 0.4)
  const pie = model('pieOfPie', { pie: { secondaryPlot: { split: { type: 'position', count: 2 } } } }), pieDraft = cloneChartDraftConfig(pie)
  const aggregateControls = chartCanvasElementControls(aggregate(), pie)
  aggregateControls.find(control => control.key === 'pie.aggregateStyle.color').write(pieDraft, '#abcdef')
  assert.equal(chartPieAggregateStyle({ options: merged(pie, pieDraft, aggregate()) }, sources().series[0], 4, palette).color, '#abcdef')
  assert.equal(pieDraft.options.pie.aggregateStyle.fillPattern, null)
  const connectorControls = chartCanvasElementControls(connector(), pie)
  assert.equal(connectorControls.length, 5); assert.equal(connectorControls[0].fallback, true)
  connectorControls.find(control => control.key === 'pie.secondaryPlot.connector.width').write(pieDraft, 4)
  assert.equal(chartPieConnector({ options: merged(pie, pieDraft, connector()) }, '#123456').width, 4)
})

test('new Pie author DTOs validate bounds and clone nested colors and lines without persisting runtime defaults', () => {
  const gradient = chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 1 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] })
  const options = { pie: { aggregateStyle: { color: gradient, explosion: 20, border: { width: 2 } }, secondaryPlot: { connector: { opacity: 0.4, dashType: 'dashed' } } } }
  const detached = cloneChartOptions(options), before = json(detached)
  gradient.stops[0].color = 'lime'; options.pie.aggregateStyle.border.width = 9; options.pie.secondaryPlot.connector.opacity = 1
  assert.deepEqual(json(detached), before)
  for (const pie of [{ aggregateStyle: { explosion: -1 } }, { aggregateStyle: { explosion: 101 } }, { aggregateStyle: { fillOpacity: 2 } },
    { aggregateStyle: { border: { width: -1 } } }, { secondaryPlot: { connector: { opacity: -1 } } }, { secondaryPlot: { connector: { visible: 'true' } } }])
    assert.throws(() => validateChartOptions({ pie }))
  const chart = model('pieOfPie'), snapshot = json(chart)
  assert.equal(chartPieConnector(chart, '#123456').width, 1); assert.equal(chartPieConnector(chart, '#abcdef').color, '#abcdef')
  chartPieAggregateStyle(chart, sources().series[0], 4, palette); assert.deepEqual(json(chart), snapshot)
})
