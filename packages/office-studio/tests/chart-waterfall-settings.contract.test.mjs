import assert from 'node:assert/strict'
import { test } from 'vitest'
import { cloneChartOptions, mergeChartOptions, validateChartOptions } from '../src/features/charts/chart-options.ts'
import { chartSettingsPatch, createChartSettingsDraft } from '../src/features/charts/chart-settings-draft.ts'
import { chartTypeSettingsControls } from '../src/features/charts/chart-type-settings-controls.ts'
import { initializeChartWaterfallOptions } from '../src/features/charts/chart-waterfall.ts'
import { CHART_WATERFALL_POINT_CHOICE_LIMIT, chartWaterfallLayoutControls, chartWaterfallPointControls,
  chartWaterfallPointRole, chartWaterfallPointRoleControls, chartWaterfallStyleControls,
  chartWaterfallStylePointIsActive, chartWaterfallStylePoints } from '../src/features/charts/chart-waterfall-settings-controls.ts'

// P18 authoring/UI contracts; no runner, rendered component or browser ran.
const palette = ['#112233', '#223344', '#334455', '#445566']
function model(options = {}) {
  return { id: 'waterfall-settings', sheetId: 'sheet', revision: 0, type: 'waterfall', categoryField: 0, valueFields: [1, 2],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 640, height: 400 }, palette,
    legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options }
}
const role = (seriesIndex, dataPointIndex, value) => ({ seriesIndex, dataPointIndex, role: value })

test('fixed creation defaults use the first finite source identity without changing runtime authoring', () => {
  const sources = [{ field: 7, values: [null, NaN, Infinity] }, { field: 3, values: [null, -5, 8] }]
  const config = model({ appearance: { background: 'pink' } }), before = structuredClone(config)
  const options = initializeChartWaterfallOptions(config, sources)
  assert.deepEqual(options.waterfall.pointRoles, [role(3, 1, 'total')])
  assert.equal(options.waterfall.connector, true)
  assert.equal(options.waterfall.stackType, undefined); assert.equal(options.waterfall.useSubtotal, undefined)
  options.appearance.background = 'blue'
  assert.deepEqual(config, before)
  const allInvalid = initializeChartWaterfallOptions(model(), [{ field: 9, values: [null, NaN] }])
  assert.equal(Object.hasOwn(allInvalid.waterfall, 'pointRoles'), false); assert.equal(allInvalid.waterfall.connector, true)
})

test('creation defaults distinguish missing properties from explicit empty/reset roles and connector false', () => {
  const sources = [{ field: 9, values: [4, -2] }]
  for (const pointRoles of [[], undefined, [role(9, 1, 'change')]]) {
    const config = model({ waterfall: { pointRoles, connector: false } }), before = structuredClone(config)
    const options = initializeChartWaterfallOptions(config, sources)
    assert.equal(Object.hasOwn(options.waterfall, 'pointRoles'), true)
    assert.deepEqual(options.waterfall.pointRoles, pointRoles)
    assert.equal(options.waterfall.connector, false)
    assert.deepEqual(config, before)
  }
  const explicitReset = initializeChartWaterfallOptions(model({ waterfall: { connector: undefined } }), sources)
  assert.equal(Object.hasOwn(explicitReset.waterfall, 'connector'), true); assert.equal(explicitReset.waterfall.connector, undefined)
  const stacked = initializeChartWaterfallOptions(model({ waterfall: { stackType: 'stacked' } }), sources)
  assert.equal(Object.hasOwn(stacked.waterfall, 'pointRoles'), false); assert.equal(stacked.waterfall.connector, true)
  const other = { type: 'column', options: { series: { 1: { color: 'red' } } } }, options = initializeChartWaterfallOptions(other, sources)
  assert.deepEqual(options, other.options); assert.notEqual(options, other.options); assert.equal(options.waterfall, undefined)
})

test('layout controls restore runtime defaults independently and preserve isolated source authoring', () => {
  const original = model({ waterfall: { stackType: 'arrangement', useSubtotal: false, connector: { color: 'red' },
    pointRoles: [role(1, 0, 'total')], styles: { negative: { color: 'blue' } } } }), before = structuredClone(original)
  const draft = createChartSettingsDraft(original), controls = chartWaterfallLayoutControls(), [stack, subtotal] = controls
  assert.deepEqual(controls.map(control => control.key), ['waterfall.stackType', 'waterfall.useSubtotal'])
  assert.equal(stack.fallback, 'arrangement'); assert.equal(subtotal.fallback, true)
  stack.write(draft.config, 'stacked')
  assert.equal(draft.config.options.waterfall.pointRoles, undefined)
  assert.equal(draft.config.options.waterfall.stackType, 'stacked')
  assert.equal(draft.config.options.waterfall.useSubtotal, false)
  assert.equal(draft.config.options.waterfall.connector.color, 'red')
  subtotal.write(draft.config, true); subtotal.reset(draft.config); stack.reset(draft.config)
  assert.equal(stack.canReset(draft.config), false); assert.equal(subtotal.canReset(draft.config), false)
  const patch = chartSettingsPatch(draft)
  assert.equal(patch.options.waterfall.stackType, undefined); assert.equal(patch.options.waterfall.useSubtotal, undefined)
  assert.equal(patch.options.waterfall.styles.negative.color, 'blue')
  assert.deepEqual(original, before)
})

test('partial options can switch to stacked while newly authored roles reject atomically', () => {
  const current = { waterfall: { pointRoles: [role(1, 0, 'total')], connector: true, useSubtotal: false } }, before = structuredClone(current)
  const stacked = mergeChartOptions(current, { waterfall: { stackType: 'stacked' } })
  assert.equal(stacked.waterfall.pointRoles, undefined)
  assert.equal(stacked.waterfall.connector, true); assert.equal(stacked.waterfall.useSubtotal, false)
  assert.throws(() => mergeChartOptions(stacked, { waterfall: { pointRoles: [role(1, 0, 'total')] } }), /堆积瀑布图/)
  assert.throws(() => mergeChartOptions(current, { waterfall: { stackType: 'stacked', pointRoles: [role(1, 0, 'change')] } }), /堆积瀑布图/)
  assert.throws(() => cloneChartOptions({ waterfall: { stackType: 'stacked', pointRoles: [role(1, 0, 'total')] } }), /堆积瀑布图/)
  assert.deepEqual(current, before)
  const retained = structuredClone(stacked)
  assert.throws(() => validateChartOptions({ waterfall: { stackType: 'grouped' } }), TypeError)
  assert.throws(() => validateChartOptions({ waterfall: { useSubtotal: 'true' } }), TypeError)
  assert.deepEqual(stacked, retained)
})

test('duplicate authored point identities use their final input while the saved array keeps its order', () => {
  const pointRoles = [role(1, 0, 'change'), role(2, 2, 'total'), role(1, 0, 'total'), role(2, 2, 'change')]
  const config = model({ waterfall: { pointRoles } }), cloned = cloneChartOptions(config.options)
  assert.deepEqual(cloned.waterfall.pointRoles, pointRoles)
  assert.notEqual(cloned.waterfall.pointRoles, pointRoles)
  assert.equal(chartWaterfallPointRole(config, 1, 0), 'total')
  assert.equal(chartWaterfallPointRole(config, 2, 2), 'change')
  assert.equal(chartWaterfallPointRole(config, 2, 0), 'change')
  assert.deepEqual(JSON.parse(JSON.stringify(cloned)).waterfall.pointRoles, pointRoles)
})

test('source-role controls edit/reset only the selected identity and reject stacked writes', () => {
  const original = model({ waterfall: { pointRoles: [role(1, 0, 'change'), role(2, 2, 'total'), role(1, 0, 'total'), role(2, 2, 'change')] } })
  const draft = createChartSettingsDraft(original).config, control = chartWaterfallPointRoleControls(1, 0)[0]
  assert.deepEqual(control.choices.map(choice => choice.value), ['change', 'total'])
  assert.equal(control.read(draft), 'total'); assert.equal(control.canReset(draft), true)
  control.write(draft, 'change')
  assert.deepEqual(draft.options.waterfall.pointRoles, [role(2, 2, 'total'), role(2, 2, 'change'), role(1, 0, 'change')])
  control.reset(draft)
  assert.deepEqual(draft.options.waterfall.pointRoles, [role(2, 2, 'total'), role(2, 2, 'change')])
  assert.equal(control.read(draft), 'change'); assert.equal(control.canReset(draft), false)
  assert.equal(original.options.waterfall.pointRoles.length, 4)
  const stacked = model({ waterfall: { stackType: 'stacked' } }), before = structuredClone(stacked)
  assert.throws(() => control.write(stacked, 'total'), /堆积瀑布图/)
  assert.throws(() => control.reset(stacked), /堆积瀑布图/)
  assert.deepEqual(stacked, before)
})

test('source-role index bounds accept the final source identity and reject invalid indices before writing', () => {
  const config = model(), control = chartWaterfallPointRoleControls(1_048_575, 1_048_575)[0]
  control.write(config, 'total'); validateChartOptions(config.options)
  assert.equal(chartWaterfallPointRole(config, 1_048_575, 1_048_575), 'total')
  for (const [field, index] of [[-1, 0], [1_048_576, 0], [0, -1], [0, 1_048_576], [0, 0.5], [NaN, 0], [0, Infinity]]) {
    assert.throws(() => chartWaterfallPointRoleControls(field, index), RangeError)
  }
})

test('style choices reflect source roles, invalid values as zero and the true synthetic source index', () => {
  const data = { categories: ['A', 'B', 'C', 'D', 'E'], series: [
    { field: 1, name: 'First', values: [6, -2, null, NaN, 0] }, { field: 2, name: 'Second', values: [-1, 2, -5, Infinity, 1] },
  ] }, config = model({ waterfall: { pointRoles: [role(1, 0, 'total')] } })
  const choices = (target, selector) => chartWaterfallStylePoints(config, target, selector, data).map(point => point.value)
  assert.deepEqual(choices('positive', 1), [2, 3, 4])
  assert.deepEqual(choices('negative', { name: 'Second' }), [0, 2])
  assert.deepEqual(choices('negative', undefined), [0, 1, 2])
  assert.deepEqual(choices('subtotal', 1), [0, 5])
  assert.deepEqual(choices('subtotal', 2), [5])
  config.options.waterfall.useSubtotal = false
  assert.deepEqual(choices('subtotal', 1), [0]); assert.deepEqual(choices('subtotal', 2), [])
  for (const index of [-1, 5.5, 6, Infinity]) assert.equal(chartWaterfallStylePointIsActive(config, 'positive', index, undefined, data), false)
})

test('stacked style choices expose only the shared global subtotal and suppress source total authoring', () => {
  const data = { categories: ['A', 'B'], series: [{ field: 1, name: 'First', values: [10, -3] }, { field: 2, name: 'Second', values: [4, -2] }] }
  const config = model({ waterfall: { stackType: 'stacked' } })
  assert.deepEqual(chartWaterfallStylePoints(config, 'subtotal', undefined, data), [{ value: 2, label: '自动全局合计' }])
  assert.deepEqual(chartWaterfallStylePoints(config, 'subtotal', 1, data), [])
  assert.equal(chartWaterfallStylePointIsActive(config, 'subtotal', 2, { name: 'First' }, data), false)
  assert.deepEqual(chartWaterfallStylePoints(config, 'positive', undefined, data).map(point => point.value), [0])
  config.options.waterfall.useSubtotal = false
  assert.deepEqual(chartWaterfallStylePoints(config, 'subtotal', undefined, data), [])
})

test('large source style choices stay bounded while direct source and automatic subtotal indices remain editable', () => {
  const count = CHART_WATERFALL_POINT_CHOICE_LIMIT + 5, values = Array(count).fill(1); values[count - 1] = -1
  const data = { categories: Array.from({ length: count }, (_unused, index) => String(index + 1)), series: [{ field: 1, name: 'Large', values }] }, config = model()
  assert.equal(chartWaterfallStylePoints(config, 'positive', 1, data).length, CHART_WATERFALL_POINT_CHOICE_LIMIT)
  assert.deepEqual(chartWaterfallStylePoints(config, 'negative', 1, data), [])
  assert.equal(chartWaterfallStylePointIsActive(config, 'negative', count - 1, 1, data), true)
  assert.equal(chartWaterfallStylePointIsActive(config, 'subtotal', count, 1, data), true)
  assert.deepEqual(chartWaterfallStylePoints(config, 'subtotal', 1, data).map(point => point.value), [count])
  const negative = chartWaterfallPointControls('negative', count - 1, 1), subtotal = chartWaterfallPointControls('subtotal', count, 1)
  negative.find(control => control.kind === 'chart-color').write(config, 'red')
  subtotal.find(control => control.kind === 'chart-color').write(config, 'blue')
  validateChartOptions(config.options)
  assert.equal(config.options.series[1].waterfallStyles.negative.dataPoints[count - 1].color, 'red')
  assert.equal(config.options.series[1].waterfallStyles.subtotal.dataPoints[count].color, 'blue')
})

test('global stacked controls inherit shared role and point styles without leaking selected field overrides', () => {
  const config = model({ waterfall: { stackType: 'stacked', styles: { subtotal: { color: 'orange', dataPoints: { 2: { color: 'purple' } } } } },
    allSeriesStyle: { waterfallStyles: { subtotal: { fillOpacity: 0.7 } } },
    series: { 1: { waterfallStyles: { subtotal: { color: 'red', fillOpacity: 0.2, dataPoints: { 2: { color: 'blue' } } } } } } })
  const source = { field: 1, position: 0, name: 'First' }, controls = chartWaterfallStyleControls('subtotal', undefined, source)
  assert.equal(controls.find(control => control.kind === 'chart-color').fallback(config), 'orange')
  assert.equal(controls.find(control => control.key.endsWith('.fillOpacity')).read(config), 0.7)
  const point = chartWaterfallPointControls('subtotal', 2, undefined, source)
  assert.equal(point.find(control => control.kind === 'chart-color').fallback(config), 'purple')
  point.find(control => control.kind === 'chart-color').write(config, 'pink')
  assert.equal(config.options.waterfall.styles.subtotal.dataPoints[2].color, 'pink')
  assert.equal(config.options.series[1].waterfallStyles.subtotal.dataPoints[2].color, 'blue')
})

test('waterfall keeps a working category gap control and omits the unrelated series overlap value', () => {
  const config = model({ bar: { overlap: 75 } }), controls = chartTypeSettingsControls(config)
  assert.equal(controls.some(control => control.key === 'bar.overlap'), false)
  const gap = controls.find(control => control.key === 'bar.gapWidth')
  assert.ok(gap); gap.write(config, 250)
  assert.equal(config.options.bar.gapWidth, 250); assert.equal(config.options.bar.overlap, 75)
})
