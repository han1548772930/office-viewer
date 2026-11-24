import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartAxisTitleCommand, chartAxisTitleFont, chartRectangularAxisTitleCommands } from '../src/features/charts/chart-axis-title.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { drawDerivedChartData } from '../src/features/charts/chart-draw.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { normalizeChartColor } from '../src/features/charts/chart-color.ts'
import { chartAxisControls } from '../src/features/charts/chart-settings-controls.ts'
import { cloneChartDraftConfig } from '../src/features/charts/chart-settings-draft.ts'

// P12 implementation drafts only. Fixed wke physical-box behavior is static evidence.
// No runner, browser, build, pixel comparison or Office sample has been executed.
const source = () => ({ categories: ['A', 'B', 'C'], series: [{ field: 1, name: 'Units', values: [1, 3, 5] }, { field: 2, name: 'Rate', values: [2, 4, 6] }] })
function model(type = 'line', patch = {}) {
  return { id: 'axis-title-draft', sheetId: 'sheet', type, revision: 0, categoryField: 0, valueFields: [1, 2],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 520, height: 360 }, legend: { visible: false, position: 'bottom' },
    xAxis: { scaleType: 'linear', title: 'Category title', titleAlignment: 'start' },
    yAxis: { scaleType: 'linear', title: 'Value title', titleAlignment: 'end' }, ...patch }
}
const texts = (option) => (option.graphic?.elements ?? []).flatMap(element => element.type === 'text' ? [element] : element.children?.filter(child => child.type === 'text') ?? [])
const normalize = command => ({ text: command.text, x: command.x, y: command.y, rotation: -(command.rotation ?? 0) * Math.PI / 180,
  font: command.font, align: command.align, color: command.color, width: command.maxWidth })
const nativeText = element => ({ text: element.style.text, x: element.x, y: element.y, rotation: element.rotation,
  font: element.style.font, align: element.style.align, color: normalizeChartColor(element.style.fill), width: element.style.width })

test('physical horizontal start/end anchors ignore axis reversal and preserve either endpoint input order', () => {
  const chart = model(), first = { x: 30, y: 200 }, last = { x: 230, y: 200 }
  for (const [alignment, x, align] of [['start', 30, 'left'], ['center', 130, 'center'], ['end', 230, 'right']]) {
    const axis = { title: 'Physical', titleAlignment: alignment }, before = structuredClone(axis)
    const ordinary = chartAxisTitleCommand(chart, axis, first, last, 'navy')
    assert.deepEqual([ordinary.x, ordinary.y, ordinary.align, ordinary.maxWidth], [x, 200, align, 200])
    assert.deepEqual(chartAxisTitleCommand(chart, { ...axis, reverse: true }, first, last, 'navy'), ordinary)
    assert.deepEqual(chartAxisTitleCommand(chart, axis, last, first, 'navy'), ordinary)
    assert.deepEqual(axis, before)
  }
})

test('vertical start/end are top/bottom and rotated glyph alignment keeps text within the physical span', () => {
  const chart = model(), bottom = { x: 20, y: 200 }, top = { x: 20, y: 30 }
  for (const [rotation, startAlign, endAlign] of [[-90, 'right', 'left'], [90, 'left', 'right']]) {
    const start = chartAxisTitleCommand(chart, { title: 'Vertical', titleAlignment: 'start', reverse: true }, bottom, top, 'black', rotation)
    const end = chartAxisTitleCommand(chart, { title: 'Vertical', titleAlignment: 'end' }, bottom, top, 'black', rotation)
    assert.deepEqual([start.x, start.y, start.align, start.maxWidth], [20, 30, startAlign, 170])
    assert.deepEqual([end.x, end.y, end.align, end.maxWidth], [20, 200, endAlign, 170])
  }
})

test('rectangular titles bind horizontal Bar category/value/top-secondary to their real edges and omit inactive secondary axes', () => {
  const chart = model('bar-horizontal', { options: { rightYAxis: { scaleType: 'linear', title: 'Top rate', titleAlignment: 'end' } } })
  const layout = { frame: { x: 0, y: 0, width: 520, height: 360 }, plot: { x: 90, y: 90, width: 350, height: 160 }, horizontal: true, rightMin: 0, rightMax: 10 }
  const commands = chartRectangularAxisTitleCommands(chart, layout, '#64748b'), titles = Object.fromEntries(commands.map(command => [command.text, command]))
  assert.equal(titles['Category title'].y, layout.plot.y)
  assert.equal(titles['Category title'].rotation, -90)
  assert.equal(titles['Value title'].x, layout.plot.x + layout.plot.width)
  assert.equal(titles['Value title'].rotation, 0)
  assert.equal(titles['Top rate'].x, layout.plot.x + layout.plot.width)
  assert.ok(titles['Top rate'].y < layout.plot.y)
  assert.equal(titles['Top rate'].rotation, 0)
  assert.equal(chartRectangularAxisTitleCommands(chart, { ...layout, rightMin: undefined, rightMax: undefined }, 'black').length, 2)
})

test('title font inheritance, empty titles and default 12px preserve author fields and exact manual plot spans', () => {
  const chart = model('line', { options: { appearance: { font: { family: 'Aptos', fontSize: 18, color: 'purple', bold: true } } },
    xAxis: { scaleType: 'linear', title: 'Inherited', titleAlignment: 'end', titleStyle: { family: undefined, fontSize: 22, italic: true } } })
  const before = structuredClone(chart), layout = { frame: { x: 100, y: 70, width: 520, height: 360 }, plot: { x: 140, y: 130, width: 260, height: 180 } }
  assert.deepEqual(chartAxisTitleFont(chart, chart.xAxis), { family: 'Aptos', fontSize: 22, color: 'purple', bold: true, italic: true })
  const title = chartRectangularAxisTitleCommands(chart, layout, 'black').find(command => command.text === 'Inherited')
  assert.equal(title.x, 400); assert.equal(title.maxWidth, 260)
  assert.equal(title.font, 'italic 600 22px Aptos'); assert.equal(title.color, 'purple')
  assert.match(chartAxisTitleCommand(model(), { title: 'Default' }, { x: 0, y: 0 }, { x: 200, y: 0 }, 'black').font, /12px/)
  assert.equal(chartAxisTitleCommand(chart, { title: '' }, { x: 0, y: 0 }, { x: 200, y: 0 }, 'black'), undefined)
  assert.deepEqual(chart, before)
})

test('native Cartesian content and common Canvas share physical title geometry and never emit native axis names', () => {
  for (const type of ['line', 'bar-horizontal', 'scatter', 'bubble', 'boxplot', 'radar', 'heatmap']) {
    const values = source(), chart = model(type, { options: { appearance: { font: { family: 'Georgia', fontSize: 16 } } } })
    if (type === 'scatter' || type === 'bubble') for (const series of values.series) series.xValues = [1, 2, 3]
    if (type === 'bubble') for (const series of values.series) series.sizeValues = [1, 1, 1]
    const before = structuredClone(chart), derived = deriveChartData(values, chart, { preparedData: values })
    const commands = drawDerivedChartData(derived, chart), option = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: values })
    const titleText = texts(option).filter(element => element.style.text === 'Category title' || element.style.text === 'Value title')
    const canvasTitles = commands.filter(command => command.kind === 'text' && (command.text === 'Category title' || command.text === 'Value title'))
    assert.equal(titleText.length, 2, type)
    assert.deepEqual(titleText.map(nativeText), canvasTitles.map(normalize), type)
    for (const axis of [option.xAxis, option.yAxis].flat().filter(Boolean)) assert.equal(axis.name, undefined, type)
    assert.equal(titleText.every(element => element.style.overflow === 'truncate'), true)
    assert.deepEqual(chart, before)
  }
})

test('Heatmap legends survive title appending, and reversing data axes leaves native title anchors unchanged', () => {
  const values = source(), chart = model('heatmap', { legend: { visible: true, position: 'bottom' } })
  const option = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: values })
  assert.ok(option.graphic.elements.some(element => element.id.startsWith(chart.id + ':heatmap-legend')))
  assert.ok(option.graphic.elements.some(element => element.id.startsWith(chart.id + ':axis-titles')))
  const titleText = texts(option).filter(element => element.id.startsWith(chart.id + ':axis-titles')).map(nativeText)
  chart.xAxis.reverse = true; chart.yAxis.reverse = true
  const reverse = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: values })
  assert.deepEqual(texts(reverse).filter(element => element.id.startsWith(chart.id + ':axis-titles')).map(nativeText), titleText)
  const layout = layoutDerivedChartData(deriveChartData(values, chart, { preparedData: values }), chart)
  assert.equal(titleText.find(element => element.text === 'Category title').x, layout.plot.x)
  assert.equal(titleText.find(element => element.text === 'Value title').y, layout.plot.y + layout.plot.height)
})

test('title controls expose inherited font and seven real resets while preserving unrelated axis fields and draft isolation', () => {
  const original = model('line', { options: { appearance: { font: { family: 'Georgia', fontSize: 18, bold: true } } },
    yAxis: { scaleType: 'linear', min: 1, max: 20, reverse: true, title: 'Units', titleAlignment: 'start', titleStyle: { fontSize: 24, italic: true } } })
  const before = structuredClone(original), draft = cloneChartDraftConfig(original), controls = chartAxisControls('yAxis')
  const title = controls.find(control => control.key === 'yAxis.title'), size = controls.find(control => control.key === 'yAxis.titleStyle.fontSize')
  const fields = controls.filter(control => control.key === 'yAxis.title' || control.key === 'yAxis.titleAlignment' || control.key.startsWith('yAxis.titleStyle.'))
  assert.equal(fields.length, 7); assert.equal(fields.every(control => typeof control.reset === 'function'), true)
  assert.equal(size.read(draft), 24); size.reset(draft); assert.equal(size.read(draft), 18)
  assert.equal(size.canReset(draft), false)
  assert.equal(controls.find(control => control.key === 'yAxis.titleStyle.family').read(draft), 'Georgia')
  title.reset(draft)
  assert.deepEqual(draft.yAxis, { scaleType: 'linear', min: 1, max: 20, reverse: true, title: undefined, titleStyle: undefined, titleAlignment: undefined })
  assert.equal(title.canReset(draft), false)
  assert.equal(size.read(model()), 12)
  assert.deepEqual(original, before)
})
