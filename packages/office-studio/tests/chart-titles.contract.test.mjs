import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createChartDefinition, updateChartDefinition } from '../src/features/charts/chart-engine.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { drawDerivedChartData } from '../src/features/charts/chart-draw.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { chartPlotFrame } from '../src/features/charts/chart-layout.ts'
import { chartTitleControls, chartTitleFontControls } from '../src/features/charts/chart-settings-controls.ts'
import { chartTitleCommands, chartTitleFont, chartTitleGeometry, chartTitlePatch } from '../src/features/charts/chart-titles.ts'
import { createChartHeatmapScale } from '../src/features/charts/chart-heatmap.ts'
import { createChartHeatmapLegend } from '../src/features/charts/chart-heatmap-draw.ts'

// Drafts for the later verification phase; no runner or browser has executed.
const workbook = { sheets: [{ id: 'sheet', cells: {} }] }
const source = { categories: ['North', 'South'], series: [{ field: 1, name: 'Revenue', values: [4, 7] }, { field: 2, name: 'Costs', values: [2, 3] }] }
function model(patch = {}) {
  return { id: 'title-chart', sheetId: 'sheet', type: 'line', revision: 0,
    source: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: 520, height: 360 }, categoryField: 0, valueFields: [1, 2],
    title: 'Main', subtitle: 'Sub', legend: { visible: true, position: 'bottom' },
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, ...patch }
}
const elements = option => (option.graphic?.elements ?? []).flatMap(function flatten(element) { return element.children ? element.children.flatMap(flatten) : [element] })

test('title specs preserve explicit clears, detach flat fonts and reject invalid fields before changing a candidate', () => {
  const value = { text: 'New', fontSize: 18, color: '#123456', family: 'Georgia', bold: false, italic: true, position: 'left', alignment: 'right' }
  const patch = chartTitlePatch(value)
  value.family = 'Arial'
  assert.equal(patch.options.titleStyle.family, 'Georgia')
  assert.equal(patch.title, 'New'); assert.equal(Object.hasOwn(patch.options.titleStyle, 'text'), false)
  assert.deepEqual(chartTitlePatch('Text'), { title: 'Text' })
  assert.deepEqual(chartTitlePatch({ text: undefined }), { title: undefined })
  assert.deepEqual(chartTitlePatch(undefined), { title: undefined, options: { titleStyle: undefined } })
  assert.deepEqual(chartTitlePatch(undefined, 'subtitle'), { subtitle: undefined, options: { subtitleStyle: undefined } })
  for (const bad of [null, [], { text: 1 }, { visible: 'yes' }, { position: 'middle' }, { alignment: 'start' }, { family: {} }, { color: '#ggg' }]) assert.throws(() => chartTitlePatch(bad))
  assert.throws(() => chartTitlePatch({ fontSize: 0 }), RangeError)
  assert.throws(() => chartTitlePatch({ position: 'left' }, 'subtitle'), TypeError)
  assert.throws(() => chartTitlePatch('x'.repeat(100_001)), TypeError)
})

test('legacy string snapshots support independent object patches and full title clears without erasing subtitle or appearance', () => {
  const legacy = createChartDefinition(workbook, 'sheet', model({ options: { appearance: { font: { family: 'Georgia' } }, titleStyle: { fontSize: 18, color: '#123456' } } }), 'title-chart')
  const before = structuredClone(legacy), moved = updateChartDefinition(workbook, legacy, chartTitlePatch({ position: 'bottom', alignment: 'right' }))
  assert.equal(moved.title, 'Main'); assert.equal(moved.subtitle, 'Sub'); assert.equal(moved.options.titleStyle.fontSize, 18)
  const replaced = updateChartDefinition(workbook, moved, chartTitlePatch('Replacement'))
  assert.equal(replaced.options.titleStyle.position, 'bottom')
  const cleared = updateChartDefinition(workbook, replaced, chartTitlePatch(undefined))
  assert.equal(cleared.title, undefined); assert.equal(cleared.options.titleStyle, undefined); assert.equal(cleared.subtitle, 'Sub')
  assert.equal(cleared.options.appearance.font.family, 'Georgia')
  assert.deepEqual(legacy, before)
})

test('all four title positions reserve the matching edge, place the subtitle inward and Hide removes both texts and reservation', () => {
  const bare = chartPlotFrame(model({ title: undefined, subtitle: undefined }), source.series).plot
  for (const position of ['top', 'bottom', 'left', 'right']) {
    const chart = model({ options: { titleStyle: { position, alignment: 'left' }, subtitleStyle: { alignment: 'right' } } })
    const geometry = chartTitleGeometry(chart), frame = chartPlotFrame(chart, source.series), commands = chartTitleCommands(chart, '#334155')
    assert.ok(geometry.bands[position] > 0)
    assert.ok(commands.length >= 2)
    if (position === 'top') { assert.ok(frame.plot.y > bare.y); assert.ok(commands[0].y < commands.at(-1).y) }
    if (position === 'bottom') { assert.ok(frame.plot.height < bare.height); assert.ok(commands[0].y > commands.at(-1).y) }
    if (position === 'left') { assert.ok(frame.plot.x > bare.x); assert.ok(commands[0].x < commands.at(-1).x) }
    if (position === 'right') { assert.ok(frame.plot.width < bare.width); assert.ok(commands[0].x > commands.at(-1).x) }
    assert.ok(commands.every(command => command.clip.width === chart.size.width && Number.isFinite(command.x) && Number.isFinite(command.y)))
  }
  for (const titleStyle of [{ position: 'hide' }, { visible: false, position: 'left' }]) {
    const hidden = model({ options: { titleStyle } })
    assert.deepEqual(chartTitleCommands(hidden, '#334155'), [])
    assert.deepEqual(chartPlotFrame(hidden, source.series).plot, bare)
  }
})

test('settings expose inherited fonts, independent resets and restoration from hidden position without writing defaults into authors', () => {
  const chart = model({ options: { appearance: { font: { family: 'Georgia', fontSize: 20, color: '#123456', italic: true } },
    titleStyle: { visible: false, position: 'bottom', alignment: 'right', fontSize: 30 }, subtitleStyle: { alignment: 'left' } } })
  const fontControls = chartTitleFontControls('title'), size = fontControls.find(control => control.key === 'titleStyle.fontSize')
  assert.equal(size.read(chart), 30); size.reset(chart); assert.equal(size.read(chart), 20)
  assert.equal(chart.options.titleStyle.alignment, 'right'); assert.equal(chart.options.titleStyle.visible, false)
  assert.equal(fontControls.find(control => control.key === 'titleStyle.family').read(chart), 'Georgia')
  const controls = chartTitleControls('title'), position = controls.find(control => control.key === 'titleStyle.position')
  assert.equal(position.read(chart), 'hide'); position.write(chart, 'left'); assert.equal(chart.options.titleStyle.visible, true)
  position.reset(chart); assert.equal(chart.options.titleStyle.position, undefined); assert.equal(chart.options.titleStyle.alignment, 'right')
  controls.find(control => control.key === 'title').reset(chart)
  assert.equal(chart.title, undefined); assert.equal(chart.options.titleStyle, undefined); assert.equal(chart.subtitle, 'Sub')
  assert.equal(chart.options.subtitleStyle.alignment, 'left'); assert.equal(chartTitleFont(chart, 'title').family, 'Georgia')
})

test('native and shared Canvas title graphics use the same positions and styles, retaining heatmap color scale graphics', () => {
  const chart = model({ options: { titleStyle: { position: 'bottom', alignment: 'left', fontSize: 17, color: '#123456', italic: true }, subtitleStyle: { alignment: 'right' } } })
  const derived = deriveChartData(source, chart, { preparedData: source }), common = drawDerivedChartData(derived, chart).filter(command => command.kind === 'text' && ['Main', 'Sub'].includes(command.text))
  const native = chartEChartsOption(workbook, chart, {}, { preparedData: source }), graphics = elements(native).filter(element => element.type === 'text' && ['Main', 'Sub'].includes(element.style.text))
  assert.equal(native.title, undefined); assert.equal(graphics.length, 2)
  assert.deepEqual(graphics.map(element => [element.style.text, element.x, element.y, element.style.font, element.style.align, element.style.fill]), common.map(command => [command.text, command.x, command.y, command.font, command.align, command.color]))
  const shared = chartEChartsOption(workbook, { ...chart, options: { ...chart.options, allSeriesStyle: { label: { visible: true, position: 'insideEnd' } } } }, {}, { preparedData: source })
  assert.deepEqual(elements(shared).filter(element => element.type === 'text' && ['Main', 'Sub'].includes(element.style.text)).map(element => [element.x, element.y]), graphics.map(element => [element.x, element.y]))
  const heatmap = chartEChartsOption(workbook, { ...chart, type: 'heatmap' }, {}, { preparedData: source })
  assert.equal(elements(heatmap).filter(element => element.type === 'text' && ['Main', 'Sub'].includes(element.style.text)).length, 2)
  assert.ok(elements(heatmap).some(element => element.type === 'rect'))
})

test('text layout fits bounded rows without dropping author text, preserves code points and leaves manual plot rectangles exact', () => {
  const chart = model({ title: 'AB', subtitle: undefined, size: { width: 80, height: 50 }, legend: { visible: false, position: 'bottom' }, options: { titleStyle: { position: 'left', fontSize: 10 } } })
  assert.deepEqual(chartTitleGeometry(chart).entries[0].lines, ['A', 'B'])
  const long = { ...chart, title: 'ABC' }
  assert.deepEqual(chartTitleGeometry(long).entries[0].lines, ['A', '…']); assert.equal(long.title, 'ABC')
  const emoji = { ...chart, title: '😀A' }
  assert.deepEqual(chartTitleGeometry(emoji).entries[0].lines, ['😀', 'A'])
  const wrapped = model({ title: 'Long text '.repeat(500), subtitle: undefined, size: { width: 90, height: 70 } })
  const geometry = chartTitleGeometry(wrapped)
  assert.ok(geometry.entries[0].lines.length <= Math.floor((70 - 24) / geometry.entries[0].lineHeight))
  assert.equal(wrapped.title.length, 5000)
  const manual = { ...model(), options: { titleStyle: { position: 'left' }, plotArea: { left: 0.2, top: 0.3, width: 0.5, height: 0.4 } } }
  assert.deepEqual(chartPlotFrame(manual, source.series).plot, { x: 104, y: 108, width: 260, height: 144 })
})

test('continuous and piecewise heatmap legends keep bars/hits outside the title band on each matching edge', () => {
  for (const type of ['continuous', 'piecewise']) for (const position of ['top', 'bottom', 'left', 'right']) {
    const chart = model({ type: 'heatmap', legend: { visible: true, position }, options: { titleStyle: { position }, heatmap: { visualMapType: type } } })
    const layout = chartPlotFrame(chart, source.series), bands = chartTitleGeometry(chart).bands
    const scale = createChartHeatmapScale({ min: 0, max: 8 }, ['#ff0000', '#0000ff'], type), legend = createChartHeatmapLegend(chart, layout, scale, '#334155')
    const rectangles = legend.bar ? [legend.bar] : legend.pieces.map(piece => piece.hit)
    assert.ok(rectangles.length > 0)
    for (const rect of rectangles) {
      if (position === 'top') assert.ok(rect.y >= bands.top)
      if (position === 'bottom') assert.ok(rect.y + rect.height <= chart.size.height - bands.bottom)
      if (position === 'left') assert.ok(rect.x >= bands.left)
      if (position === 'right') assert.ok(rect.x + rect.width <= chart.size.width - bands.right)
    }
    assert.equal(legend.frame, layout.frame)
  }
})
