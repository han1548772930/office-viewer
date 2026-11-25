import assert from 'node:assert/strict'
import { test } from 'vitest'
import { updateChartDefinition } from '../src/features/charts/chart-engine.ts'
import { chartEChartsOption, chartGraphics } from '../src/features/charts/echarts-option.ts'
import { chartTitleCommands, chartTitleGeometry, chartTitlePatch } from '../src/features/charts/chart-titles.ts'
import { chartTitleAuthorContextMatches, chartTitleDragLayout, chartTitleGraphicLayouts, chartTitleOutlineCommands,
  clipChartTitlePolygon, createChartTitleTargets, hitChartTitle } from '../src/features/charts/chart-title-interaction.ts'

// Drafts only. No runner, type checker, browser or benchmark has executed.
const workbook = { sheets: [{ id: 'sheet', cells: {} }] }
const source = { categories: ['A', 'B'], series: [{ field: 1, name: 'Main', values: [4, 7] }] }
function model(options = {}) {
  return { id: 'chart', sheetId: 'sheet', type: 'line', revision: 0, title: 'Main', subtitle: 'Sub', categoryField: 0, valueFields: [1],
    source: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: 400, height: 300 }, legend: { visible: true, position: 'bottom' },
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options }
}
const rect = { x: 0, y: 0, width: 100, height: 100 }
function layout(kind = 'title', patch = {}) { return { kind, lineIndex: 0, anchor: { x: 40, y: 20 }, clip: rect, manual: false, vertical: false, alignment: 'center', ...patch } }

test('manual anchors detach from public title specs and invalid pairs/rotation reject before changing the chart', () => {
  const spec = { manualLayout: { left: 0.2, top: 0.3 }, rotate: 30, text: 'Authored' }, patch = chartTitlePatch(spec)
  spec.manualLayout.left = 0.9
  assert.deepEqual(patch.options.titleStyle.manualLayout, { left: 0.2, top: 0.3 })
  const before = model()
  for (const value of [{ manualLayout: { left: 0.2 } }, { manualLayout: null }, { manualLayout: { left: -0.1, top: 0.2 } },
    { manualLayout: { left: 0.2, top: Infinity } }, { manualLayout: { left: NaN, top: 0.2 } }, { manualLayout: { left: 0.2, top: 1.1 } }, { rotate: 361 }, { rotate: -361 }, { rotate: '30' }]) {
    assert.throws(() => chartTitlePatch(value)); assert.deepEqual(before, model())
  }
  assert.equal(chartTitlePatch({ rotate: -360 }, 'subtitle').options.subtitleStyle.rotate, -360)
})

test('restoring automatic position clears only the selected manual anchor and preserves rotation, text, fonts and the other title', () => {
  const original = model({ titleStyle: { manualLayout: { left: 0.1, top: 0.2 }, rotate: 30, fontSize: 20, color: '#123456' },
    subtitleStyle: { manualLayout: { left: 0.3, top: 0.4 }, rotate: -15 } })
  const restored = updateChartDefinition(workbook, original, chartTitlePatch({ manualLayout: undefined }))
  assert.equal(restored.options.titleStyle.manualLayout, undefined)
  assert.equal(restored.options.titleStyle.rotate, 30); assert.equal(restored.options.titleStyle.fontSize, 20); assert.equal(restored.title, 'Main')
  assert.deepEqual(restored.options.subtitleStyle, original.options.subtitleStyle)
  assert.deepEqual(original.options.titleStyle.manualLayout, { left: 0.1, top: 0.2 })
})

test('manual titles stop reserving automatic bands, use normalized anchors and retain unwrapped author text behind the chart clip', () => {
  const chart = model({ titleStyle: { manualLayout: { left: 0.2, top: 0.3 }, alignment: 'right' }, subtitleStyle: { manualLayout: { left: 0.3, top: 0.4 } } })
  chart.title = 'Long horizontal title '.repeat(50)
  assert.deepEqual(chartTitleGeometry(chart).bands, { top: 0, bottom: 0, left: 0, right: 0 })
  const commands = chartTitleCommands(chart, '#334155'), title = commands.find(command => command.chartTitle.kind === 'title')
  assert.equal(title.text, chart.title); assert.equal(title.maxWidth, undefined); assert.equal(title.align, 'left')
  assert.deepEqual(title.chartTitle.anchor, { x: 80, y: 90 }); assert.deepEqual(title.clip, { x: 0, y: 0, width: 400, height: 300 })
  const oneAutomatic = model({ titleStyle: { manualLayout: { left: 0.2, top: 0.3 } } })
  assert.ok(chartTitleGeometry(oneAutomatic).bands.top > 0)
})

test('rotated multiline titles share the same group anchor, clockwise geometry and tagged stable graphics on native/shared paths', () => {
  const chart = model({ titleStyle: { manualLayout: { left: 0.2, top: 0.3 }, rotate: 90, fontSize: 16 } }); chart.title = 'One\nTwo'
  const commands = chartTitleCommands(chart, '#334155').filter(command => command.chartTitle.kind === 'title')
  assert.equal(commands.length, 2); assert.ok(commands[1].x < commands[0].x); assert.ok(Math.abs(commands[0].y - commands[1].y) < 1e-8)
  assert.deepEqual(commands[0].chartTitle.anchor, commands[1].chartTitle.anchor)
  const graphics = chartGraphics(commands, chart.id).flatMap(element => element.children ?? [element])
  assert.equal(new Set(graphics.map(element => element.id)).size, 2)
  assert.ok(graphics.every(element => element.info.chartElementHit.kind === 'title' && element.silent === false))
  assert.ok(graphics.every(element => Math.abs(element.rotation + Math.PI / 2) < 1e-8))
  const native = chartEChartsOption(workbook, chart, {}, { preparedData: source })
  const shared = chartEChartsOption(workbook, { ...chart, options: { ...chart.options, allSeriesStyle: { label: { visible: true, position: 'insideEnd' } } } }, {}, { preparedData: source })
  for (const option of [native, shared]) {
    const targets = [...chartTitleGraphicLayouts(option.graphic.elements).values()]
    assert.equal(targets.filter(target => target.kind === 'title').length, 2); assert.equal(targets.filter(target => target.kind === 'subtitle').length, 1)
  }
})

test('graphic metadata rejects malformed targets and never identifies series/axis text merely because it equals the title', () => {
  const valid = layout(), metadata = chartTitleGraphicLayouts([{ id: 'clip', children: [{ id: 'title-0', type: 'text', info: { chartTitle: valid } }] },
    { id: 'series-label', type: 'text', style: { text: 'Main' } }, { id: 'invalid', info: { chartTitle: { ...valid, anchor: { x: NaN, y: 0 } } } },
    { id: 'axis', info: { chartTitle: { ...valid, kind: 'axis-title' } } }])
  assert.deepEqual([...metadata.keys()], ['title-0'])
  valid.anchor.x = 99; assert.equal(metadata.get('title-0').anchor.x, 40)
})

test('clipped transformed polygons hit glyph slots, exclude bounding-box corners/outside charts and resolve subtitle overlap last', () => {
  const polygon = [{ x: 50, y: -10 }, { x: 110, y: 50 }, { x: 50, y: 110 }, { x: -10, y: 50 }]
  const clipped = clipChartTitlePolygon(polygon, rect)
  assert.ok(clipped.every(point => point.x >= 0 && point.x <= 100 && point.y >= 0 && point.y <= 100))
  const targets = createChartTitleTargets([{ layout: layout(), polygon, leading: { x: -10, y: 0 } }, { layout: layout('subtitle'), polygon }])
  assert.equal(hitChartTitle(targets, { x: 50, y: 50 }).kind, 'subtitle')
  assert.equal(hitChartTitle(targets.slice(0, 1), { x: 50, y: 0 }).kind, 'title')
  assert.equal(hitChartTitle(targets, { x: 0, y: 0 }), undefined); assert.equal(hitChartTitle(targets, { x: 101, y: 50 }), undefined)
  assert.equal(hitChartTitle(targets, { x: NaN, y: 0 }), undefined)
  assert.deepEqual(targets[0].anchor, { x: 30, y: 20 })
})

test('outline geometry and normalized dragging use measured anchors without changing source targets or creating a model mutation', () => {
  const target = createChartTitleTargets([{ layout: layout('title', { manual: true }), polygon: [{ x: 30, y: 20 }, { x: 50, y: 20 }, { x: 50, y: 40 }, { x: 30, y: 40 }] }])[0]
  const before = structuredClone(target), commands = chartTitleOutlineCommands(target, { x: 200, y: 300 }, '#1677ff', 1.5)
  assert.equal(commands.length, 4); assert.deepEqual(commands[0].clip, { x: 200, y: 300, width: 100, height: 100 })
  assert.equal(commands[0].x1, 230); assert.equal(commands[0].y1, 320)
  assert.deepEqual(chartTitleDragLayout(target, { width: 100, height: 100 }, { x: 20, y: 10 }), { left: 0.6, top: 0.3 })
  assert.deepEqual(chartTitleDragLayout(target, { width: 100, height: 100 }, { x: 200, y: -100 }), { left: 1, top: 0 })
  assert.throws(() => chartTitleDragLayout(target, { width: 0, height: 100 }, { x: 0, y: 0 }), RangeError)
  assert.deepEqual(target, before)
})

test('author edits are rejected after owner/workbook/sheet/chart/revision/snapshot changes, even when chart IDs are reused', () => {
  const owner = {}, context = { owner, workbookId: 'book', sheetId: 'sheet', chartId: 'chart', revision: 4, snapshot: JSON.stringify(model()) }
  assert.equal(chartTitleAuthorContextMatches(context, { ...context }), true)
  for (const patch of [{ owner: {} }, { workbookId: 'other' }, { sheetId: 'other' }, { chartId: 'other' }, { revision: 5 }, { snapshot: 'changed' }])
    assert.equal(chartTitleAuthorContextMatches(context, { ...context, ...patch }), false)
  assert.equal(chartTitleAuthorContextMatches(context, undefined), false)
})
