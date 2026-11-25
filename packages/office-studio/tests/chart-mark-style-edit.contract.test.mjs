import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartCanvasElementAuthorSignature, chartCanvasElementCompatible, chartCanvasElementControls, chartCanvasElementDeleteDraft,
  chartCanvasElementHidden, chartCanvasElementPatch, resetChartCanvasElementDraft } from '../src/features/charts/chart-element-edit.ts'
import { chartCanvasElementAsSeries, chartCanvasElementKey, chartCanvasElementRoute, chartCanvasPieSliceMark,
  chartCanvasSeriesBindings, chartCanvasSeriesMark, readChartCanvasElement } from '../src/features/charts/chart-element-interaction.ts'
import { cloneChartDraftConfig, createChartSettingsDraft, chartSettingsPatch } from '../src/features/charts/chart-settings-draft.ts'
import { cloneChartOptions, mergeChartOptions, chartSeriesStyle, validateChartOptions } from '../src/features/charts/chart-options.ts'
import { chartSeriesOverride } from '../src/features/charts/chart-series-selector.ts'
import { chartPointStyle } from '../src/features/charts/chart-derived.ts'
import { chartPieSliceAllowsExplosion, chartPieSliceStyle } from '../src/features/charts/chart-pie.ts'
import { chartPieSliceControls } from '../src/features/charts/chart-pie-settings-controls.ts'
import { chartTypeSettingsControls } from '../src/features/charts/chart-type-settings-controls.ts'
import { chartWaterfallStyleControls } from '../src/features/charts/chart-waterfall-settings-controls.ts'
import { chartLinearGradient } from '../src/features/charts/chart-color.ts'
import { validateChartGradientAuthoring } from '../src/features/charts/chart-gradient.ts'
import { CHART_EDIT_PANEL_SECTIONS, normalizeChartEditPanelRoute, planChartEditPanelRoute } from '../src/features/charts/chart-edit-panel-route.ts'

// P21 implementation drafts. No test, type/build check, browser or service was run.
const palette = ['#112233', '#223344', '#334455', '#445566']
const json = value => JSON.parse(JSON.stringify(value))
const source = () => ({ categories: ['A', 'B', 'C', 'D'], series: [
  { field: 1, name: 'Renamed Alpha', sourceName: ' Alpha ', values: [8, 5, 3, 2] },
  { field: 2, name: 'Renamed Beta', sourceName: 'Beta', values: [2, 3, 5, 7] },
] })
const model = (type = 'line', options = {}) => ({ id: 'mark-style', sheetId: 'sheet', revision: 0, type,
  categoryField: 0, valueFields: [1, 2], source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 4, endCol: 2 } },
  position: { x: 0, y: 0 }, size: { width: 600, height: 420 }, palette, title: 'Original chart',
  legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options })
const mark = directEdit => ({ kind: 'mark', directEdit })
const series = (chart, position = 0) => chartCanvasSeriesMark(chartCanvasSeriesBindings(source(), chart.options)[position])
const slice = (index = 2, plot = 'primary') => chartCanvasPieSliceMark(source().series[0], index, plot)
const merged = (chart, draft, element) => mergeChartOptions(chart.options, chartCanvasElementPatch(chart, draft, element).options)
function write(chart, element, key, value) {
  const control = chartCanvasElementControls(element, chart).find(control => control.key === key)
  assert.ok(control, key); assert.notEqual(control.visible?.(chart), false); control.write(chart, value)
}

test('a numeric seriesStyle edit owns only its appearance and retains source mapping, independent points and other series', () => {
  const chart = model('combination', { series: {
    1: { name: 'Authored name', type: 'line', rightAxis: true, color: '#112233', dataPoints: { 1: { size: 18 } },
      waterfallStyles: { negative: { color: '#445566' } } },
    2: { color: '#223344', label: { visible: true } },
  } }), before = json(chart), element = series(chart), draft = cloneChartDraftConfig(chart)
  write(draft, element, 'series.1.fillOpacity', 0.45); write(draft, element, 'series.1.border.width', 3)
  draft.options.series[1].name = 'Ignored'; draft.options.series[1].type = 'column'; draft.options.series[1].rightAxis = false
  draft.options.series[1].dataPoints[1].size = 7; draft.options.series[2].color = '#ffffff'; draft.source.range.endRow = 9
  const patch = chartCanvasElementPatch(chart, draft, element), next = mergeChartOptions(chart.options, patch.options)
  assert.deepEqual(Object.keys(patch), ['options']); assert.equal(next.series[1].fillOpacity, 0.45); assert.equal(next.series[1].border.width, 3)
  assert.equal(next.series[1].name, 'Authored name'); assert.equal(next.series[1].type, 'line'); assert.equal(next.series[1].rightAxis, true)
  assert.deepEqual(next.series[1].dataPoints, chart.options.series[1].dataPoints)
  assert.deepEqual(next.series[1].waterfallStyles, chart.options.series[1].waterfallStyles); assert.deepEqual(next.series[2], chart.options.series[2])
  assert.deepEqual(json(chart), before); assert.equal(chartCanvasElementKey(element), 'mark:seriesStyle:field:1')
  const keys = chartCanvasElementControls(element, chart).map(control => control.key)
  assert.ok(keys.includes('series.1.fillPattern.preset')); assert.ok(keys.includes('series.1.label.fontSize'))
  assert.ok(keys.every(key => !['.name', '.type', '.rightAxis'].some(suffix => key.endsWith(suffix))))
})

test('an exact-name seriesStyle reset restores lower appearance while preserving named point maps, sibling names and author identity', () => {
  const chart = model('line', { allSeriesStyle: { label: { visible: true } }, series: { 1: { color: '#112233', point: { size: 6 } } },
    namedSeries: [{ name: ' Alpha ', style: { name: 'Display name', color: '#445566', point: { size: 12 },
      label: { fontSize: 20 }, dataPoints: { 2: { color: '#abcdef', size: 17 } } } }, { name: 'Beta', style: { color: '#223344' } }] }),
    element = series(chart), draft = cloneChartDraftConfig(chart)
  assert.deepEqual(element.directEdit.binding.selector, { name: ' Alpha ' })
  write(draft, element, 'namedSeries." Alpha ".fillOpacity', 0.5)
  const edited = merged(chart, draft, element); assert.equal(edited.namedSeries[0].style.fillOpacity, 0.5)
  assert.deepEqual(edited.series, chart.options.series); assert.deepEqual(edited.namedSeries[1], chart.options.namedSeries[1])
  const next = merged(chart, resetChartCanvasElementDraft(chart, element), element), style = next.namedSeries[0].style
  assert.equal(style.color, undefined); assert.equal(style.point, undefined); assert.equal(style.label, undefined)
  assert.equal(style.name, 'Display name'); assert.deepEqual(style.dataPoints, chart.options.namedSeries[0].style.dataPoints)
  assert.equal(chartSeriesStyle({ options: next }, 1, 0, palette, ' Alpha ').color, '#112233')
  assert.equal(chartPointStyle(chartSeriesStyle({ options: next }, 1, 0, palette, ' Alpha '), 2, 'line').point.size, 17)
  assert.deepEqual(chartCanvasElementAsSeries(element), { kind: 'series', binding: element.directEdit.binding })
  assert.deepEqual(chartCanvasElementRoute(element), { tab: 'style', section: 'series', seriesId: 'name:" Alpha "' })
})

test('a real Boxplot outlier can author size and shape at its raw sample index while summary and sibling samples remain independent', () => {
  const chart = model('boxplot', { boxplot: { showOutliers: true }, series: { 1: { color: '#112233', point: { size: 4 },
    dataPoints: { 0: { color: '#abcdef' }, 13: { size: 9 } } } } }), binding = chartCanvasSeriesBindings(source(), chart.options)[0],
    element = { kind: 'dataPoint', binding, dataIndex: 13 }, draft = cloneChartDraftConfig(chart)
  write(draft, element, 'point.size', 16); write(draft, element, 'point.shape', 'diamond')
  const next = merged(chart, draft, element), style = chartPointStyle(chartSeriesStyle({ options: next }, 1, 0, palette, ' Alpha '), 13, 'boxplot')
  assert.equal(style.point.size, 16); assert.equal(style.point.shape, 'diamond')
  assert.deepEqual(next.series[1].dataPoints[0], chart.options.series[1].dataPoints[0]); assert.deepEqual(next.series[1].point, chart.options.series[1].point)
  const hidden = cloneChartDraftConfig(chart); hidden.options.boxplot.showOutliers = false
  for (const target of [element, { kind: 'series', binding }, chartCanvasSeriesMark(binding)])
    assert.ok(chartCanvasElementControls(target, hidden).filter(control => control.key.endsWith('.size') || control.key.endsWith('.shape'))
      .every(control => control.visible?.(hidden) === false))
})

test('pieStyle edits, resets and Delete address the global label layer without replacing slice, aggregate or legacy series authors', () => {
  const chart = model('pieOfPie', { series: { 1: { label: { visible: true, fontSize: 11 } } }, pie: {
    label: { visible: true, position: 'inside', fontSize: 14 }, sliceStyles: { 2: { color: '#abcdef' } },
    aggregateStyle: { color: '#445566' }, aggregateLabel: 'Others', secondaryPlot: { split: { type: 'position', count: 2 } },
  } }), element = mark({ kind: 'pieStyle' }), draft = cloneChartDraftConfig(chart)
  write(draft, element, 'pie.label.fontSize', 19); draft.options.pie.aggregateLabel = 'Ignored'; draft.options.pie.sliceStyles[2].color = '#ffffff'
  const next = merged(chart, draft, element)
  assert.equal(next.pie.label.fontSize, 19); assert.equal(next.pie.aggregateLabel, 'Others')
  assert.deepEqual(next.pie.sliceStyles, chart.options.pie.sliceStyles); assert.deepEqual(next.series, chart.options.series)
  const reset = merged(chart, resetChartCanvasElementDraft(chart, element), element)
  assert.equal(reset.pie.label, undefined); assert.deepEqual(reset.pie.aggregateStyle, chart.options.pie.aggregateStyle)
  const deleted = chartCanvasElementDeleteDraft(chart, element)
  assert.equal(chartCanvasElementHidden(deleted, element), true); assert.equal(merged(chart, deleted, element).pie.label.visible, false)
  assert.deepEqual(chartCanvasElementRoute(element), { tab: 'style', section: 'pieStyle', focus: { kind: 'pieLabel' } })
})

test('pieSliceStyle is the highest independent author layer and preserves legacy source points, sibling slices and the generated aggregate', () => {
  const chart = model('pieOfPie', { series: { 1: { dataPoints: { 2: { color: '#112233', border: { width: 2 } }, 4: { color: '#223344' } } } },
    pie: { sliceStyles: { 0: { color: '#445566' }, 2: { color: '#abcdef', border: { opacity: 0.6 } } },
      aggregateStyle: { color: '#334455' }, secondaryPlot: { split: { type: 'position', count: 2 } } } }),
    element = slice(2, 'secondary'), draft = cloneChartDraftConfig(chart)
  write(draft, element, 'pie.sliceStyles.2.fillOpacity', 0.35); write(draft, element, 'pie.sliceStyles.2.explosion', 22)
  draft.options.series[1].dataPoints[2].color = '#ffffff'; draft.options.pie.sliceStyles[0].color = '#ffffff'
  const next = merged(chart, draft, element), effective = chartPieSliceStyle({ options: next }, source().series[0], 2, palette)
  assert.equal(effective.color, '#abcdef'); assert.equal(effective.fillOpacity, 0.35)
  assert.deepEqual(effective.border, { width: 2, opacity: 0.6 }); assert.equal(next.pie.sliceStyles[2].explosion, 22)
  assert.deepEqual(next.series, chart.options.series); assert.deepEqual(next.pie.sliceStyles[0], chart.options.pie.sliceStyles[0])
  assert.deepEqual(next.pie.aggregateStyle, chart.options.pie.aggregateStyle)
  assert.equal(chartCanvasElementControls(element, chart).length, 11)
  const reset = merged(chart, resetChartCanvasElementDraft(chart, element), element)
  assert.equal(chartPieSliceStyle({ options: reset }, source().series[0], 2, palette).color, '#112233')
  assert.deepEqual(reset.series, chart.options.series); assert.deepEqual(reset.pie.aggregateStyle, chart.options.pie.aggregateStyle)
  assert.equal(chartCanvasElementDeleteDraft(chart, element), undefined)
})

test('a Bar-of-Pie secondary segment hides explosion, ignores it in scoped signatures and preserves its inactive authored value', () => {
  const chart = model('barOfPie', { pie: { sliceStyles: { 2: { color: '#112233', explosion: 90 } },
    secondaryPlot: { split: { type: 'position', count: 2 } } } }), element = slice(2, 'secondary'), draft = cloneChartDraftConfig(chart)
  assert.equal(chartPieSliceAllowsExplosion(chart, 2, source(), 'secondary'), false)
  const controls = chartCanvasElementControls(element, chart)
  assert.equal(controls.length, 10); assert.ok(controls.every(control => !control.key.endsWith('.explosion')))
  draft.options.pie.sliceStyles[2].explosion = 5
  assert.equal(chartCanvasElementAuthorSignature(chart, element), chartCanvasElementAuthorSignature(draft, element))
  assert.deepEqual(chartCanvasElementPatch(chart, draft, element), {})
  write(draft, element, 'pie.sliceStyles.2.fillOpacity', 0.5)
  const next = merged(chart, draft, element); assert.equal(next.pie.sliceStyles[2].explosion, 90); assert.equal(next.pie.sliceStyles[2].fillOpacity, 0.5)
  const reset = merged(chart, resetChartCanvasElementDraft(chart, element), element)
  assert.equal(reset.pie.sliceStyles[2].color, undefined); assert.equal(reset.pie.sliceStyles[2].explosion, 90)
})

test('the full slice form reset survives the real settings patch and deep option merge without reviving removed appearance', () => {
  const chart = model('barOfPie', { pie: { label: { visible: true }, sliceStyles: { 0: { color: '#223344' },
    2: { color: '#abcdef', fillOpacity: 0.4, border: { width: 3 }, explosion: 70 } }, aggregateStyle: { color: '#334455' },
    secondaryPlot: { split: { type: 'position', count: 2 } } } }), draft = createChartSettingsDraft(chart)
  draft.config = resetChartCanvasElementDraft(draft.config, slice(2, 'secondary'))
  const patch = chartSettingsPatch(draft), next = mergeChartOptions(chart.options, patch.options)
  assert.ok(Object.hasOwn(patch.options.pie.sliceStyles[2], 'color')); assert.equal(next.pie.sliceStyles[2].color, undefined)
  assert.equal(next.pie.sliceStyles[2].fillOpacity, undefined); assert.equal(next.pie.sliceStyles[2].border, undefined)
  assert.equal(next.pie.sliceStyles[2].explosion, 70); assert.deepEqual(next.pie.sliceStyles[0], chart.options.pie.sliceStyles[0])
  assert.deepEqual(next.pie.label, chart.options.pie.label); assert.deepEqual(next.pie.aggregateStyle, chart.options.pie.aggregateStyle)
})

test('independent slice DTOs validate bounds and retain deep-cloned gradient, texture and border author data', () => {
  const color = chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 1 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }),
    options = { pie: { sliceStyles: { 2: { color, explosion: 30, fillOpacity: 0.6,
      fillPattern: { preset: 'ltUpDiag', foregroundColor: 'black', backgroundColor: 'white' }, border: { width: 2 } } } } }
  validateChartOptions(options); validateChartGradientAuthoring('pie', options)
  const cloned = cloneChartOptions(options); cloned.pie.sliceStyles[2].color.stops[0].color = 'green'; cloned.pie.sliceStyles[2].border.width = 5
  assert.equal(options.pie.sliceStyles[2].color.stops[0].color, 'red'); assert.equal(options.pie.sliceStyles[2].border.width, 2)
  assert.throws(() => validateChartGradientAuthoring('radar', options), /纯色/)
  for (const sliceStyles of [{ '-1': {} }, { '01': {} }, { 2: { explosion: 101 } }, { 2: { fillOpacity: -0.1 } }, { 2: { border: { width: Infinity } } }, []])
    assert.throws(() => validateChartOptions({ pie: { sliceStyles } }))
})

test('new mark metadata reconstructs exact ownership, rejects malformed indices and cannot author a changed or incompatible chart type', () => {
  const chart = model(), element = series(chart), pie = mark({ kind: 'pieStyle' }), point = slice(2, 'secondary')
  for (const value of [element, pie, point]) { const read = readChartCanvasElement(value); assert.deepEqual(read, value); assert.notEqual(read, value) }
  assert.notEqual(readChartCanvasElement(element).directEdit.binding, element.directEdit.binding)
  const binding = element.directEdit.binding
  for (const directEdit of [{ kind: 'seriesStyle', binding: { ...binding, seriesId: 'native:0' } },
    { kind: 'seriesStyle', binding: { ...binding, selector: 2 } }, { ...point.directEdit, sourceField: -1 },
    { ...point.directEdit, sourcePointIndex: -1 }, { ...point.directEdit, sourcePointIndex: 1.5 },
    { ...point.directEdit, plot: 'unknown' }, { kind: 'pieSliceStyle', sourceField: 1, aggregateIndex: 4 }])
    assert.equal(readChartCanvasElement(mark(directEdit)), undefined)
  assert.equal(chartCanvasElementCompatible(model('line3D'), element), false); assert.equal(chartCanvasElementCompatible(model('pie3D'), pie), false)
  const draft = cloneChartDraftConfig(chart); draft.type = 'column'
  assert.throws(() => chartCanvasElementPatch(chart, draft, element), /类型已变更/)
  assert.deepEqual(chartCanvasElementPatch(chart, cloneChartDraftConfig(chart), element), {})
})

test('local exact focus preserves the fixed 27 sections and five original route fields while validating its optional sixth DTO', () => {
  assert.equal(CHART_EDIT_PANEL_SECTIONS.length, 27)
  const focuses = [{ kind: 'pieAggregate' }, { kind: 'pieConnector' }, { kind: 'pieLabel' }, { kind: 'pieSlice', sourcePointIndex: 2 }]
  for (const focus of focuses) {
    const route = { tab: 'style', section: 'pieStyle', focus }, normalized = normalizeChartEditPanelRoute(route)
    assert.deepEqual(normalized, route); assert.notEqual(normalized.focus, focus)
  }
  for (const focus of [null, [], new Date(), Object.create(null), { kind: 'pieConnector', selector: '#arbitrary' },
    { kind: 'waterfallRole', role: 'base' }, { kind: 'pieSlice', sourcePointIndex: -1 }, { kind: 'pieSlice', sourcePointIndex: 0.5 },
    { kind: 'pieLabel', [Symbol('private')]: true }]) assert.throws(() => normalizeChartEditPanelRoute({ tab: 'style', section: 'pieStyle', focus }), TypeError)
  for (const route of [{ tab: 'data', focus: { kind: 'pieLabel' } }, { tab: 'style', section: 'series', focus: { kind: 'pieLabel' } },
    { tab: 'style', section: 'pieStyle', focus: { kind: 'waterfallRole', role: 'negative' } },
    { tab: 'style', section: 'waterfallSeries', focus: { kind: 'pieSlice', sourcePointIndex: 2 } }])
    assert.throws(() => normalizeChartEditPanelRoute(route), TypeError)
})

test('Waterfall exact focus opens the selected numeric or named role color, and global stacked totals never resolve to a fake source', () => {
  const chart = model('waterfall', { waterfall: { pointRoles: [] }, namedSeries: [{ name: ' Alpha ', style: { color: '#112233' } }] })
  for (const [seriesId, selector, prefix] of [['field:1', 1, 'series.1'], ['name:" Alpha "', { name: ' Alpha ' }, 'namedSeries." Alpha "']]) {
    for (const role of ['positive', 'negative', 'subtotal']) {
      const plan = planChartEditPanelRoute(chart, { tab: 'style', section: 'waterfallSeries', seriesId, focus: { kind: 'waterfallRole', role } }, source())
      assert.equal(plan.seriesTarget, seriesId); assert.equal(plan.waterfallRole, role); assert.equal(plan.optionKey, `${prefix}.waterfallStyles.${role}.color`)
      assert.ok(chartWaterfallStyleControls(role, selector).some(control => control.key === plan.optionKey))
    }
  }
  const stacked = model('waterfall', { waterfall: { stackType: 'stacked', useSubtotal: true } }), route = {
    tab: 'style', section: 'waterfallSeries', focus: { kind: 'waterfallRole', role: 'subtotal' },
  }, global = planChartEditPanelRoute(stacked, route, source())
  assert.equal(global.optionKey, 'waterfall.styles.subtotal.color'); assert.equal(global.seriesTarget, undefined)
  assert.throws(() => planChartEditPanelRoute(stacked, { ...route, seriesId: 'field:1' }, source()), /来源合计/)
  assert.throws(() => planChartEditPanelRoute(model('waterfall', { waterfall: { stackType: 'stacked', useSubtotal: false } }), route, source()), /自动合计/)
  assert.throws(() => planChartEditPanelRoute(chart, { ...route, seriesId: 'field:9' }, source()), /不存在/)
})

test('Pie exact focus resolves actual aggregate, connector, label and source-slice fields and rejects unavailable source or 3D targets', () => {
  const chart = model('pieOfPie', { pie: { secondaryPlot: { split: { type: 'position', count: 2 } } } }), prepared = source(),
    controls = chartTypeSettingsControls(chart, 3, ' Alpha ', prepared)
  for (const [kind, key] of [['pieAggregate', 'pie.aggregateStyle.color'], ['pieConnector', 'pie.secondaryPlot.connector.visible'], ['pieLabel', 'pie.label.visible']]) {
    const plan = planChartEditPanelRoute(chart, { tab: 'style', section: 'pieStyle', focus: { kind } }, prepared)
    assert.equal(plan.optionKey, key); assert.deepEqual(plan.panels, ['type']); assert.ok(controls.some(control => control.key === key))
  }
  const route = { tab: 'style', section: 'pieStyle', focus: { kind: 'pieSlice', sourcePointIndex: 2 } }, plan = planChartEditPanelRoute(chart, route, prepared)
  assert.equal(plan.pieSourcePointIndex, 2); assert.equal(plan.optionKey, 'pie.sliceStyles.2.color')
  assert.ok(chartPieSliceControls(2, prepared.series[0]).some(control => control.key === plan.optionKey))
  assert.throws(() => planChartEditPanelRoute(chart, route), /尚未就绪/)
  assert.throws(() => planChartEditPanelRoute(chart, { ...route, focus: { kind: 'pieSlice', sourcePointIndex: 4 } }, prepared), /不存在/)
  assert.throws(() => planChartEditPanelRoute(model('pie'), { ...route, focus: { kind: 'pieAggregate' } }, prepared), /主副图/)
  assert.throws(() => planChartEditPanelRoute(model('pie3D'), route, prepared), /三维/)
  assert.deepEqual(chartCanvasElementRoute(slice(2)), route)
  assert.equal(chartSeriesOverride(chart.options, 1), undefined)
})
