import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P18 deferred drafts. No test, server, browser or build ran in this batch.
let server, shapes, paths, options, controls, draft, settings, catalog
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[shapes, paths, options, controls, draft, settings, catalog] = await Promise.all([
    'chart-line-and-area', 'chart-line-path', 'chart-options', 'chart-type-settings-controls',
    'chart-settings-draft', 'chart-settings-data', 'chart-catalog',
  ].map(name => server.ssrLoadModule(`/packages/office-studio/src/features/charts/${name}.ts`)))
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const near = (a, b) => assert.ok(Math.abs(a - b) < 1e-8, `${a} ≈ ${b}`)
const point = (x, y, baseY, categoryIndex, sequenceIndex) => ({ x, y, baseY, categoryIndex, ...(sequenceIndex === undefined ? {} : { sequenceIndex }) })
function config(type = 'line', authored = {}) {
  return { type, position: { x: 20, y: 30 }, size: { width: 480, height: 320 },
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 } },
    categoryField: 0, valueFields: [1, 2], options: authored }
}
const shapeControl = value => controls.chartTypeSettingsControls(value).find(item => item.key === 'area.lineStyle')

test('published string enums and capabilities include all eleven line/area/combination variants and exclude Scatter shape', () => {
  assert.deepEqual(shapes.ChartAreaLineStyles, { Line: 'line', Smooth: 'smooth', Step: 'step' })
  assert.deepEqual(shapes.ChartInvalidValueTypes, { Zero: 'zero', Break: 'break', Link: 'link' })
  assert.ok(Object.isFrozen(shapes.ChartAreaLineStyles) && Object.isFrozen(shapes.ChartInvalidValueTypes))
  const supported = catalog.CHART_TYPE_CATALOG.filter(item => shapes.supportsChartAreaLineStyle(item.value))
  assert.equal(supported.length, 11)
  for (const entry of supported) assert.equal(shapes.supportsChartLineAndAreaSection(entry.value), true)
  assert.equal(shapes.supportsChartAreaLineStyle('scatter'), false)
  assert.equal(shapes.supportsChartLineAndAreaSection('scatter'), true)
  assert.equal(shapes.supportsChartLineAndAreaSection('column'), false)
  assert.equal(shapeControl(config('scatter')), undefined)
})

test('runtime enums and dedicated patches reject malformed values or unsupported chart types before mutation', () => {
  const before = config('area', { area: { lineStyle: 'step' } }), snapshot = json(before)
  assert.throws(() => shapes.chartAreaLineStyleValue(undefined), TypeError)
  for (const value of [0, null, {}, [], 'Smooth', 'bezier', 'middle', true]) {
    assert.throws(() => shapes.chartAreaLineStyleValue(value), TypeError)
    assert.throws(() => shapes.chartAreaLineStylePatch('area', value), TypeError)
  }
  for (const value of [0, null, {}, 'connect', 'Break', false]) assert.throws(() => shapes.chartInvalidValueStrategyValue(value), TypeError)
  for (const type of ['scatter', 'column', 'radar', 'pareto']) assert.throws(() => shapes.chartAreaLineStylePatch(type, 'step'), TypeError)
  assert.deepEqual(before, snapshot)
})

test('option validation preserves detached unknown area fields while rejecting invalid authored shape and legacy values', () => {
  const source = { area: { lineStyle: 'smooth', smooth: false, future: { axis: [1, 3], note: 'retained' } }, future: { enabled: true } }
  const copy = options.cloneChartOptions(source)
  assert.deepEqual(copy, source); copy.area.future.axis.push(8)
  assert.deepEqual(source.area.future.axis, [1, 3])
  for (const value of [null, 1, 'spline', {}]) assert.throws(() => options.cloneChartOptions({ area: { lineStyle: value } }))
  assert.throws(() => options.cloneChartOptions({ area: { smooth: 'true' } }))
})

test('explicit lineStyle wins over the old switch and legacy smooth remains confined to actual area series', () => {
  for (const type of ['line', 'area', 'line3D', 'area3D', 'combination']) {
    const value = config(type, { area: { lineStyle: 'step', smooth: true } })
    assert.equal(shapes.resolveChartAreaLineStyle(value, 'line'), 'step')
    assert.equal(shapes.resolveChartAreaLineStyle(value, 'area'), 'step')
  }
  assert.equal(shapes.resolveChartAreaLineStyle(config('line', { area: { smooth: true } })), 'line')
  assert.equal(shapes.resolveChartAreaLineStyle(config('area', { area: { smooth: true } })), 'smooth')
  const mixed = config('combination', { area: { smooth: true } })
  assert.equal(shapes.resolveChartAreaLineStyle(mixed, 'line'), 'line')
  assert.equal(shapes.resolveChartAreaLineStyle(mixed, 'area'), 'smooth')
  assert.equal(shapes.resolveChartAreaLineStyle(config('scatter', { area: { lineStyle: 'step' } })), 'line')
})

test('shape writes and explicit clears retain other options and cannot reactivate a legacy switch after JSON reopen', () => {
  const current = { area: { smooth: true, future: { retained: 8 } }, invalidValueStrategy: 'link', appearance: { background: 'navy' } }
  const snapshot = json(current), authored = options.mergeChartOptions(current, shapes.chartAreaLineStylePatch('area', 'step'))
  assert.equal(authored.area.lineStyle, 'step'); assert.equal(authored.area.smooth, undefined)
  assert.deepEqual(authored.area.future, { retained: 8 }); assert.equal(authored.invalidValueStrategy, 'link')
  const reset = options.mergeChartOptions(authored, shapes.chartAreaLineStylePatch('area', undefined))
  assert.equal(reset.area.lineStyle, undefined); assert.equal(reset.area.smooth, undefined)
  assert.equal(shapes.resolveChartAreaLineStyle(config('area', json(reset))), 'line')
  assert.deepEqual(reset.appearance, current.appearance); assert.deepEqual(current, snapshot)
})

test('a legacy switch-only patch intentionally leaves explicit shape mode without losing unrelated author fields', () => {
  const current = { area: { lineStyle: 'step', future: 'kept' }, invalidValueStrategy: 'zero' }
  const legacy = options.mergeChartOptions(current, { area: { smooth: true } })
  assert.equal(legacy.area.lineStyle, undefined); assert.equal(legacy.area.smooth, true)
  assert.equal(shapes.resolveChartAreaLineStyle(config('area', legacy)), 'smooth')
  assert.equal(shapes.resolveChartAreaLineStyle(config('line', legacy)), 'line')
  const unrelated = options.mergeChartOptions(current, { area: { future: 'updated' } })
  assert.equal(unrelated.area.lineStyle, 'step'); assert.equal(unrelated.area.future, 'updated')
  const cleared = options.mergeChartOptions(legacy, { area: { lineStyle: undefined } })
  assert.equal(cleared.area.smooth, undefined); assert.equal(cleared.area.future, 'kept')
})

test('the real settings Select writes only a draft, offers three shapes and resets both new and legacy authors', () => {
  const source = config('area3D', { area: { smooth: true, future: { keep: 4 } }, threeD: { view: { rotY: 47 } } }), snapshot = json(source)
  const working = draft.createChartSettingsDraft(source), control = shapeControl(working.config)
  assert.equal(control.kind, 'select'); assert.equal(control.label, '折线与面积线型')
  assert.deepEqual(control.choices, [{ value: 'line', label: '直线' }, { value: 'smooth', label: '平滑' }, { value: 'step', label: '阶梯' }])
  assert.equal(control.read(working.config), 'smooth'); assert.equal(control.canReset(working.config), true)
  control.write(working.config, 'step')
  const patch = draft.chartSettingsPatch(working)
  assert.equal(patch.options.area.lineStyle, 'step'); assert.equal(patch.options.area.smooth, undefined)
  assert.deepEqual(patch.options.threeD, source.options.threeD); assert.deepEqual(source, snapshot)
  control.reset(working.config)
  assert.equal(control.read(working.config), 'line'); assert.equal(control.canReset(working.config), false)
  assert.equal(Object.hasOwn(working.config.options.area, 'lineStyle'), true)
  assert.equal(Object.hasOwn(working.config.options.area, 'smooth'), true)
  assert.deepEqual(working.config.options.area.future, { keep: 4 })
})

test('legacy combination settings describe inherited mixed shapes without falsely selecting one global shape', () => {
  const value = config('combination', { area: { smooth: true }, series: { 1: { type: 'area' }, 2: { type: 'line' } } }), before = json(value)
  const control = shapeControl(value)
  assert.equal(control.read(value), undefined); assert.equal(control.fallback, undefined)
  assert.match(control.extra, /面积系列平滑/); assert.match(control.extra, /折线系列直线/)
  assert.deepEqual(value, before)
  control.write(value, 'line')
  assert.equal(shapeControl(value).read(value), 'line'); assert.equal(value.options.area.smooth, undefined)
})

test('shape drafts reuse the prepared data binding while invalid-value edits require a new binding', () => {
  const working = draft.createChartSettingsDraft(config()), source = { categories: ['A', 'B', 'C'],
    series: [{ field: 1, name: 'Value', values: [9, null, 8] }] }
  const snapshot = json(source), original = settings.chartSettingsBindingKey(working, 'source', false)
  const prepared = settings.chartSettingsPreparedData(working, source, { hostSheetId: 'source', sourceRevision: 7 }, false)
  shapeControl(working.config).write(working.config, 'smooth')
  assert.equal(settings.chartSettingsBindingKey(working, 'source', false), original)
  assert.equal(prepared.bindingKey, original); assert.deepEqual(prepared.data, snapshot)
  working.config.options.invalidValueStrategy = 'zero'
  assert.notEqual(settings.chartSettingsBindingKey(working, 'source', false), original)
  assert.deepEqual(source, snapshot); assert.deepEqual(working.config.source, config().source)
})

test('middle steps move the top and stacked base at the same x while keeping authored endpoints intact', () => {
  const values = [point(2, 5, 1, 0), point(10, 9, 4, 1)], snapshot = json(values)
  assert.deepEqual(paths.chartLineSegment(values, 1, 'step'), [
    { x: 2, y: 5, baseY: 1 }, { x: 6, y: 5, baseY: 1 }, { x: 6, y: 9, baseY: 4 }, { x: 10, y: 9, baseY: 4 },
  ])
  const reversed = [point(10, 9, 4, 0), point(2, 5, 1, 1)]
  assert.equal(paths.chartLineSegment(reversed, 1, 'step')[1].x, 6)
  assert.deepEqual(values, snapshot)
})

test('Smooth does not borrow neighbours across Break or Zero gaps, while Link may use the connected sequence', () => {
  const values = [point(-4, 1000, 12, 0), point(0, 1, 0, 2), point(2, 4, 3, 3), point(6, 20, 9, 5)]
  const isolated = paths.chartLineSegment(values.slice(1, 3), 1, 'smooth')
  assert.deepEqual(paths.chartLineSegment(values, 2, 'smooth', 'break'), isolated)
  assert.deepEqual(paths.chartLineSegment(values, 2, 'smooth', 'zero'), isolated)
  assert.notDeepEqual(paths.chartLineSegment(values, 2, 'smooth', 'link'), isolated)
  assert.equal(isolated.length, paths.CHART_LINE_CURVE_STEPS + 1)
})

test('Smooth handles irregular reversed dates by sequence identity and falls back safely for repeated x', () => {
  const values = [point(9, 3, 2, 2, 0), point(4, 10, 5, 0, 1), point(0, 2, 1, 1, 2)]
  const curve = paths.chartLineSegment(values, 1, 'smooth'), first = curve[0], last = curve.at(-1)
  assert.deepEqual(first, { x: 9, y: 3, baseY: 2 }); assert.deepEqual(last, { x: 4, y: 10, baseY: 5 })
  assert.ok(curve.some((value, index) => Math.abs(value.y - (3 + 7 * index / paths.CHART_LINE_CURVE_STEPS)) > 1e-3))
  for (const value of curve) {
    assert.ok(value.x >= 4 && value.x <= 9 && value.y >= 3 && value.y <= 10 && value.baseY >= 2 && value.baseY <= 5)
  }
  const repeated = [point(4, 3, 1, 7, 0), point(4, 10, 2, 2, 1)]
  assert.deepEqual(paths.chartLineSegment(repeated, 1, 'smooth'), paths.chartLineSegment(repeated, 1, 'line'))
  near(curve[6].x, 6.5)
})

test('gap connectivity preserves missing positions and invalid path indices or enum values fail without changing inputs', () => {
  const values = [point(0, 3, 0, 0), point(4, 7, 0, 2)], before = json(values)
  assert.equal(paths.chartLinePointsConnected(values[0], values[1]), false)
  assert.equal(paths.chartLinePointsConnected(values[0], values[1], 'link'), true)
  for (const style of ['line', 'smooth', 'step']) {
    assert.deepEqual(paths.chartLineSegment(values, 1, style, 'break'), [])
    assert.deepEqual(paths.chartLineSegment(values, 1, style, 'zero'), [])
    assert.ok(paths.chartLineSegment(values, 1, style, 'link').length >= 2)
  }
  for (const index of [-1, 0, 2, 1.5, NaN]) assert.throws(() => paths.chartLineSegment(values, index, 'line'), RangeError)
  assert.throws(() => paths.chartLineSegment(values, 1, 'spline'), TypeError)
  assert.throws(() => paths.chartLineSegment(values, 1, 'line', 'connect'), TypeError)
  assert.deepEqual(values, before)
})
