import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { chartColorChannels } from '../src/features/charts/chart-color.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { drawChartData } from '../src/features/charts/chart-draw.ts'
import { chartColorOpacity, chartLabelText, chartSymbol, triangulateChartPolygon } from '../src/features/charts/chart-geometry.ts'
import { CHART_LABEL_CONTENT_CHOICES, chartLabelContentFromValues, chartLabelContentMask, chartLabelContentValues, chartPointLabelAnchor } from '../src/features/charts/chart-labels.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { CHART_POINT_SHAPES, chartMarkerPoints, chartMarkerRequiresSharedGeometry, chartMarkerShape, chartNativeMarkerSymbol, isChartPointShape } from '../src/features/charts/chart-markers.ts'
import { cloneChartOptions, mergeChartOptions, validateChartOptions } from '../src/features/charts/chart-options.ts'
import { chartParetoCumulativeMarkerStyle, chartParetoCumulativeStyle, chartParetoValuePercentages } from '../src/features/charts/chart-pareto.ts'
import { chartParetoBarControls, chartParetoDataPointControls, chartParetoLineControls } from '../src/features/charts/chart-pareto-settings-controls.ts'
import { chartLabelControls, chartSeriesControls } from '../src/features/charts/chart-settings-controls.ts'
import { cloneChartDraftConfig, createChartSettingsDraft } from '../src/features/charts/chart-settings-draft.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'

// P19 implementation/verification drafts only. No runner, browser or build has run.
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, String(actual) + ' ≈ ' + expected)
const palette = ['#112233', '#223344', '#334455', '#445566']
const source = (values = [6, 4], categories = ['A', 'B']) => ({ categories, fields: ['Cause', 'Count'],
  series: [{ field: 1, name: 'Count', sourceName: 'Count', values }] })
function model(type = 'pareto', options = {}) {
  return { id: 'marker-label-draft', sheetId: 'sheet', revision: 0, type, categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } }, position: { x: 0, y: 0 },
    size: { width: 440, height: 320 }, palette, legend: { visible: false, position: 'bottom' },
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options }
}
const optionFor = (data, chart) => chartEChartsOption(createBlankWorkbook(), chart, {}, { preparedData: data })
const graphics = elements => elements.flatMap(element => element.type === 'group' ? graphics(element.children ?? []) : [element])
function area(points) {
  return Math.abs(points.reduce((sum, point, index) => {
    const next = points[(index + 1) % points.length]
    return sum + point.x * next.y - next.x * point.y
  }, 0)) / 2
}

test('all fixed public marker shapes validate in series and materialized-point style contracts', () => {
  assert.equal(CHART_POINT_SHAPES.length, 19)
  assert.equal(new Set(CHART_POINT_SHAPES).size, 19)
  for (const shape of CHART_POINT_SHAPES) {
    assert.equal(isChartPointShape(shape), true)
    assert.doesNotThrow(() => validateChartOptions({ allSeriesStyle: { point: { shape }, dataPoints: { 0: { shape } } },
      pareto: { cumulativeLineStyle: { point: { shape }, dataPoints: { 0: { shape } } } } }))
  }
  for (const shape of ['star', 'emptyDot', 'EmptyCircle', {}, 1, null]) {
    assert.equal(isChartPointShape(shape), false)
    assert.throws(() => chartMarkerShape(shape), TypeError)
    assert.throws(() => validateChartOptions({ allSeriesStyle: { point: { shape } } }), TypeError)
  }
})

test('complex marker polygons stay bounded and their triangulated fill preserves actual shape area', () => {
  const center = { x: -30, y: 12 }, before = structuredClone(center), size = 24
  for (const shape of CHART_POINT_SHAPES) {
    const points = chartMarkerPoints(center, size, shape)
    if (shape === 'none') { assert.deepEqual(points, []); continue }
    assert.equal(points.length >= 3, true)
    points.forEach(point => {
      assert.equal(Number.isFinite(point.x) && Number.isFinite(point.y), true)
      assert.equal(Math.abs(point.x - center.x) <= size / 2 + 1e-8, true)
      assert.equal(Math.abs(point.y - center.y) <= size / 2 + 1e-8, true)
    })
    const triangles = triangulateChartPolygon(points)
    assert.equal(triangles.length > 0, true)
    near(triangles.reduce((sum, triangle) => sum + area(triangle), 0), area(points))
    assert.equal(area(points) > 0, true)
  }
  assert.deepEqual(center, before)
  assert.equal(area(chartMarkerPoints(center, size, 'dash')) < area(chartMarkerPoints(center, size, 'rect')), true)
  assert.equal(area(chartMarkerPoints(center, size, 'roundRect')) < area(chartMarkerPoints(center, size, 'rect')), true)
  near(area(chartMarkerPoints(center, size, 'dot')), size * size * 0.4 * 0.2)
  near(area(chartMarkerPoints(center, size, 'dash')), size * size * 0.8 * 0.2)
  assert.equal(area(chartMarkerPoints(center, size, 'dot')) < area(chartMarkerPoints(center, size, 'dash')), true)
  assert.deepEqual(chartMarkerPoints(center, size, 'square'), chartMarkerPoints(center, size, 'rect'))
  for (const invalid of [-1, NaN, Infinity]) assert.throws(() => chartMarkerPoints(center, invalid, 'circle'), RangeError)
  assert.throws(() => chartMarkerPoints({ x: Number.MAX_VALUE, y: 0 }, Number.MAX_VALUE, 'rect'), RangeError)
  for (const shape of CHART_POINT_SHAPES) assert.deepEqual(chartMarkerPoints(center, 0, shape), [])
})

test('hollow marks contain only clipped outlines and consume fallback source alpha once', () => {
  const clip = { x: 0, y: 0, width: 40, height: 40 }, color = 'rgba(255, 0, 0, 0.4)'
  for (const shape of CHART_POINT_SHAPES) {
    const commands = []
    chartSymbol(commands, { x: 20, y: 20 }, 16, color, shape, undefined, clip)
    if (shape === 'none') { assert.deepEqual(commands, []); continue }
    if (chartMarkerShape(shape).hollow) {
      assert.equal(commands.length > 0, true)
      commands.forEach(command => {
        assert.equal(command.kind, 'line')
        near(chartColorChannels(command.color)[3], 0.4)
        assert.deepEqual(command.clip, clip)
      })
    } else {
      assert.equal(commands.some(command => command.kind === 'polygon' || command.kind === 'rect'), true)
      assert.equal(commands.some(command => command.kind === 'line'), false)
    }
  }
})

test('explicit hollow border hiding and zero width are retained without a substitute filled center', () => {
  for (const shape of CHART_POINT_SHAPES.filter(shape => shape.startsWith('empty'))) {
    for (const border of [{ width: 0 }, { visible: false }, { visible: false, width: 9 }]) {
      const commands = []
      chartSymbol(commands, { x: 20, y: 20 }, 12, 'red', shape, border)
      assert.deepEqual(commands, [])
    }
    const commands = []
    chartSymbol(commands, { x: 20, y: 20 }, 12, 'rgba(255, 0, 0, 0.2)', shape, { color: 'rgba(0, 128, 0, 0.8)', width: 3, opacity: 0.75 })
    assert.equal(commands.length > 0, true)
    commands.forEach(command => { assert.equal(command.kind, 'line'); assert.equal(command.width, 3); near(chartColorChannels(command.color)[3], 0.6) })
  }
})

test('Pareto hollow fallback borders follow marker color and opacity instead of cumulative curve width', () => {
  const chart = model('pareto', { pareto: { cumulativeLineStyle: { color: 'rgba(255, 0, 0, 0.8)', opacity: 0.5, width: 9,
    point: { shape: 'emptyCircle', size: 12 }, dataPoints: { 0: { fillOpacity: 0.25 }, 1: { border: { width: 0 } } } } } })
  const style = chartParetoCumulativeStyle(chart, palette)
  for (const [index, alpha] of [[undefined, 0.4], [0, 0.2]]) {
    const marker = chartParetoCumulativeMarkerStyle(style, index), commands = []
    chartSymbol(commands, { x: 20, y: 20 }, marker.point.size, chartColorOpacity(marker.color, marker.fillOpacity), marker.point.shape, marker.border)
    assert.equal(commands.length > 0, true)
    commands.forEach(command => { assert.equal(command.kind, 'line'); assert.equal(command.width, 1); near(chartColorChannels(command.color)[3], alpha) })
  }
  const suppressed = chartParetoCumulativeMarkerStyle(style, 1), commands = []
  chartSymbol(commands, { x: 20, y: 20 }, 12, chartColorOpacity(suppressed.color, suppressed.fillOpacity), suppressed.point.shape, suppressed.border)
  assert.deepEqual(commands, [])
  assert.equal(style.border.width, 9)
})

test('native aliases preserve authoring while complex and hollow active markers share the generated scene', () => {
  for (const [shape, native] of [['square', 'rect'], ['circle', 'circle'], ['rect', 'rect']]) {
    const chart = model('line', { allSeriesStyle: { point: { shape } } }), before = structuredClone(chart), option = optionFor(source(), chart)
    assert.equal(chartNativeMarkerSymbol(shape), native)
    assert.equal(option.series[0].symbol, native)
    assert.deepEqual(chart, before)
  }
  assert.equal(chartNativeMarkerSymbol('dot'), 'path://M0 0L0 0M10 10L10 10M3 4H7V6H3Z')
  assert.equal(chartNativeMarkerSymbol('dash'), 'path://M0 0L0 0M10 10L10 10M1 4H9V6H1Z')
  for (const shape of ['dot', 'dash', 'roundRect', 'pin', 'arrow', 'emptyCircle', 'emptySquare', 'emptyArrow']) {
    assert.equal(chartMarkerRequiresSharedGeometry(shape), true)
    const chart = model('pareto', { pareto: { cumulativeLineStyle: { point: { shape, size: 13 } } } })
    const option = optionFor(source(), chart), commands = drawChartData(source(), chart), items = graphics(option.graphic.elements)
    assert.deepEqual(option.series ?? [], [])
    assert.deepEqual(items.filter(item => item.type === 'line').map(item => item.style.stroke),
      commands.filter(command => command.kind === 'line').map(command => command.color))
    assert.deepEqual(items.filter(item => item.type === 'polygon').map(item => item.style.fill),
      commands.filter(command => command.kind === 'polygon').map(command => command.fill))
  }
})

test('inactive or zero-size marker overrides do not change the active native render route', () => {
  const cases = [
    model('pareto', { pareto: { cumulativeLineStyle: { point: { shape: 'emptyCircle', size: 0 } } } }),
    model('pareto', { pareto: { cumulativeLineStyle: { dataPoints: { 99: { shape: 'emptyPin', size: 18 } } } } }),
    model('line', { allSeriesStyle: { point: { shape: 'circle' }, dataPoints: { 99: { shape: 'emptyPin' } } } }),
    model('column', { allSeriesStyle: { point: { shape: 'emptyArrow' } } }),
  ]
  for (const chart of cases) assert.equal(optionFor(source(), chart).series.length > 0, true)
  const blankLine = model('pareto', { rightYAxis: { scaleType: 'log' }, pareto: { includeZeroValues: true, cumulativeLineStyle: { point: { shape: 'emptyPin' } } } })
  assert.deepEqual(optionFor(source([0, 0]), blankLine).series.at(-1).data, [null, null])
})

test('content flags accept existing single strings, composed masks and explicit empty labels while rejecting unknown bits', () => {
  assert.equal(chartLabelContentMask(undefined), 8)
  for (const [content, _label, flag] of CHART_LABEL_CONTENT_CHOICES) {
    assert.equal(chartLabelContentMask(flag), flag)
    if (content !== 'bubbleSize') assert.equal(chartLabelContentMask(content), flag)
  }
  assert.deepEqual(chartLabelContentValues(30), ['series', 'category', 'value', 'percentage'])
  assert.equal(chartLabelContentFromValues(['percentage', 'category', 'value', 'series'], false), 30)
  assert.equal(chartLabelContentFromValues([], false), 0)
  assert.equal(chartLabelText({ contentType: 0 }, 12, 'A', 'Count', 60), '')
  for (const invalid of [-1, 1, 3, 31, 63, 64, 1.5, NaN, Infinity, 'bubbleSize', 'unknown', null, {}]) {
    assert.throws(() => chartLabelContentMask(invalid), TypeError)
  }
  assert.throws(() => chartLabelContentMask(32, false), TypeError)
  assert.throws(() => chartLabelContentFromValues(['value', 'value']), TypeError)
  assert.throws(() => chartLabelContentFromValues(['value', 'bogus']), TypeError)
})

test('combined label text uses fixed comma order and formats business percentage and bubble value independently', () => {
  assert.equal(chartLabelText({ contentType: 30, format: '0.00' }, 6, 'A', 'Count', 60), 'Count,A,6.00,60.00%')
  assert.equal(chartLabelText({ contentType: 18, format: '0.0%' }, 6, 'A', 'Count', 60), 'A,60.0%')
  assert.equal(chartLabelText({ contentType: 40, format: '0.00' }, 10, 'A', 'Count', undefined, 7), '10.00,7.00')
  assert.equal(chartLabelText({ contentType: 32 }, 10, 'A', 'Count', undefined, 7), '7')
  assert.equal(chartLabelText({ contentType: 32 }, 10, 'A', 'Count'), '')
  assert.equal(chartLabelText({ contentType: 'series' }, 10, 'A', 'Count'), 'Count')
})

test('Pareto validation rejects BubbleSize masks atomically and keeps numeric masks through deep cloning and patches', () => {
  const original = { pareto: { barStyle: { label: { visible: true, contentType: 30, position: 'insideBase' } },
    cumulativeLineStyle: { point: { shape: 'emptyRoundRect' }, label: { contentType: 0, position: 'auto' } } } }
  const before = structuredClone(original), cloned = cloneChartOptions(original), merged = mergeChartOptions(cloned, { pareto: { barStyle: { label: { position: 'insideEnd' } } } })
  assert.equal(merged.pareto.barStyle.label.contentType, 30)
  assert.equal(merged.pareto.barStyle.label.position, 'insideEnd')
  assert.equal(merged.pareto.cumulativeLineStyle.label.contentType, 0)
  for (const group of ['barStyle', 'cumulativeLineStyle']) {
    assert.throws(() => cloneChartOptions({ pareto: { [group]: { label: { contentType: 32 } } } }), TypeError)
    assert.throws(() => mergeChartOptions(original, { pareto: { [group]: { label: { contentType: 33 } } } }), TypeError)
  }
  assert.deepEqual(original, before)
})

test('multiselect controls publish whole validated masks into detached drafts and keep hidden label fields', () => {
  const chart = model('pareto', { pareto: { cumulativeLineStyle: { label: { visible: false, contentType: 'category', bold: true } } } })
  const before = structuredClone(chart), config = cloneChartDraftConfig(createChartSettingsDraft(chart).config), controls = chartParetoLineControls()
  const content = controls.find(control => control.key.endsWith('.label.contentType')), visible = controls.find(control => control.key.endsWith('.label.visible'))
  assert.equal(content.kind, 'multiselect')
  assert.deepEqual(content.read(config), ['category'])
  assert.equal(content.visible(config), false)
  visible.write(config, true)
  content.write(config, ['percentage', 'category', 'series'])
  assert.equal(config.options.pareto.cumulativeLineStyle.label.contentType, 22)
  assert.deepEqual(content.read(config), ['series', 'category', 'percentage'])
  const authored = structuredClone(config)
  for (const invalid of [['value', 'bad'], ['value', 'value'], ['bubbleSize'], [8], null]) {
    assert.throws(() => content.write(config, invalid), TypeError)
    assert.deepEqual(config, authored)
  }
  visible.write(config, false)
  assert.equal(config.options.pareto.cumulativeLineStyle.label.contentType, 22)
  assert.equal(config.options.pareto.cumulativeLineStyle.label.bold, true)
  content.write(config, [])
  assert.equal(config.options.pareto.cumulativeLineStyle.label.contentType, 0)
  assert.deepEqual(chart, before)
})

test('generic label editors retain inactive imported BubbleSize flags while bubble-specific controls can edit them', () => {
  const config = createChartSettingsDraft(model('column', { allSeriesStyle: { label: { contentType: 40 } } })).config
  const controls = chartSeriesControls(), content = controls.find(control => control.key.endsWith('.label.contentType'))
  assert.deepEqual(content.read(config), ['value'])
  content.write(config, ['category'])
  assert.equal(config.options.allSeriesStyle.label.contentType, 34)
  const bubble = chartSeriesControls(undefined, false, false, false, undefined, true).find(control => control.key.endsWith('.label.contentType'))
  assert.deepEqual(bubble.read(config), ['category', 'bubbleSize'])
  bubble.write(config, ['bubbleSize'])
  assert.equal(config.options.allSeriesStyle.label.contentType, 32)
  const generic = chartLabelControls('label', () => ({ contentType: 0 }), () => ({})).find(control => control.key.endsWith('.contentType'))
  assert.deepEqual(generic.read(config), [])
})

test('Pareto controls expose every real marker and all public label positions with cumulative outside excluded', () => {
  const line = chartParetoLineControls(), bar = chartParetoBarControls(), point = chartParetoDataPointControls('cumulativeLineStyle', 0)
  assert.deepEqual(line.find(control => control.key.endsWith('.point.shape')).choices.map(choice => choice.value), CHART_POINT_SHAPES)
  assert.deepEqual(point.find(control => control.key.endsWith('.shape')).choices.map(choice => choice.value), CHART_POINT_SHAPES)
  for (const position of ['auto', 'insideBase', 'insideEnd']) {
    assert.equal(line.find(control => control.key.endsWith('.label.position')).choices.some(choice => choice.value === position), true)
    assert.equal(bar.find(control => control.key.endsWith('.label.position')).choices.some(choice => choice.value === position), true)
  }
  assert.equal(line.find(control => control.key.endsWith('.label.position')).choices.some(choice => choice.value === 'outside'), false)
  assert.equal(line.find(control => control.key.endsWith('.label.contentType')).choices.some(choice => choice.value === 'bubbleSize'), false)
})

test('baseline and endpoint anchors follow projected sign, reversed axes and short marks instead of fixed top offsets', () => {
  for (const horizontal of [false, true]) {
    for (const direction of [-1, 1]) {
      const point = horizontal ? { x: 50 + direction * 100, y: 70, baseX: 50, value: 1 } : { x: 70, y: 50 + direction * 100, baseY: 50, value: 1 }
      const base = chartPointLabelAnchor('insideBase', point, horizontal), end = chartPointLabelAnchor('insideEnd', point, horizontal), outside = chartPointLabelAnchor('auto', point, horizontal)
      near(horizontal ? base.x : base.y, 50 + direction * 9)
      near(horizontal ? end.x : end.y, 50 + direction * 91)
      near(horizontal ? outside.x : outside.y, 50 + direction * 109)
      if (horizontal) { assert.equal(base.align, direction > 0 ? 'left' : 'right'); assert.equal(end.align, direction > 0 ? 'right' : 'left') }
    }
  }
  const tiny = { x: 12, y: 0, baseX: 10 }
  assert.equal(chartPointLabelAnchor('insideBase', tiny, true).x, 11)
  assert.equal(chartPointLabelAnchor('insideEnd', tiny, true).x, 11)
  assert.equal(chartPointLabelAnchor('insideEnd', { x: 10, y: 20 }).y, 20)
})

test('real column/bar render commands place internal labels near their own baseline/end under negative and reversed values', () => {
  const data = source([8, -6])
  for (const type of ['column', 'bar-horizontal']) for (const reverse of [false, true]) for (const position of ['insideBase', 'insideEnd']) {
    const chart = model(type, { allSeriesStyle: { label: { visible: true, contentType: 'category', position } } })
    chart.yAxis.reverse = reverse
    const layout = layoutDerivedChartData(deriveChartData(data, chart), chart), commands = drawChartData(data, chart), option = optionFor(data, chart)
    assert.deepEqual(option.series ?? [], [])
    for (const point of layout.series[0].points) {
      const label = commands.find(command => command.kind === 'text' && command.text === point.category
        && command.clip?.x === layout.plot.x && command.clip?.y === layout.plot.y && command.clip?.width === layout.plot.width)
      const coordinate = type === 'bar-horizontal' ? label.x : label.y, base = type === 'bar-horizontal' ? point.baseX : point.baseY, end = type === 'bar-horizontal' ? point.x : point.y
      assert.equal(coordinate >= Math.min(base, end) && coordinate <= Math.max(base, end), true)
      assert.equal(position === 'insideBase' ? Math.abs(coordinate - base) < Math.abs(coordinate - end) : Math.abs(coordinate - end) < Math.abs(coordinate - base), true)
    }
  }
})

test('equal business shares stay equal after cumulative truncation, with identical native and Canvas combination labels', () => {
  const data = source([4, 4, 4], ['A', 'B', 'C']), chart = model('pareto', { pareto: {
    barStyle: { label: { visible: true, contentType: 18, format: '0.00%' } },
    cumulativeLineStyle: { label: { visible: true, contentType: 18, format: '0.00%' } },
  } })
  const derived = deriveChartData(data, chart), option = optionFor(data, chart), labels = drawChartData(data, chart).filter(command => command.kind === 'text').map(command => command.text)
  assert.deepEqual(derived.pareto.cumulativePercent, [33.33, 66.66, 100])
  chartParetoValuePercentages(derived.pareto.values).forEach(value => near(value, 100 / 3))
  assert.deepEqual(option.series[0].data.map(point => point.label.formatter), ['A,33.33%', 'B,33.33%', 'C,33.33%'])
  assert.deepEqual(option.series[1].data.map(point => point.label.formatter), ['A,33.33%', 'B,66.66%', 'C,100.00%'])
  for (const label of ['A,33.33%', 'B,33.33%', 'C,33.33%', 'B,66.66%', 'C,100.00%']) assert.equal(labels.includes(label), true)
  assert.equal(labels.some(label => label.includes('33.34')), false)
  chartParetoValuePercentages([Number.MAX_VALUE, Number.MAX_VALUE, Number.MAX_VALUE]).forEach(value => near(value, 100 / 3))
  assert.deepEqual(chartParetoValuePercentages([0, 0]), [0, 0])
  assert.deepEqual(chartParetoValuePercentages([6, 4, 0]), [60, 40, 0])
  assert.throws(() => chartParetoValuePercentages([1, -1]), RangeError)
  assert.throws(() => chartParetoValuePercentages([Infinity]), RangeError)
})
