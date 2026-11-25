import assert from 'node:assert/strict'
import { test } from 'vitest'
import { sunburstFixture } from './fixtures/chart-sunburst-fixtures.mjs'
import { CHART_SUNBURST_RESIDUAL_PREFIX, chartSunburstCenterText, chartSunburstLabelAnchor, chartSunburstLabelFits,
  chartSunburstLabelText, chartSunburstLabelVisible, chartSunburstRenderNodes, chartSunburstSectorRect, chartSunburstShade,
  chartSunburstTotal, resolveChartSunburstCenterLabel, resolveChartSunburstLabel } from '../src/features/charts/chart-sunburst.ts'
import { chartSunburstLayout } from '../src/features/charts/chart-sunburst-layout.ts'
import { drawChartSunburst } from '../src/features/charts/chart-sunburst-draw.ts'
import { assertChartSunburstType, chartSunburstCenterLabelValues, chartSunburstHierarchyFields, chartSunburstLabelValues,
  chartSunburstValueField } from '../src/features/charts/chart-sunburst-api.ts'
import { chartSunburstBackgroundMode, chartSunburstCenterLabelControls, chartSunburstNodeLabelControls, clearChartSunburstSettings,
  setChartSunburstBackground, setChartSunburstBackgroundMode } from '../src/features/charts/chart-sunburst-settings-controls.ts'
import { chartSunburstCenterGraphic, chartSunburstEChartsData, chartSunburstTooltip, createChartSunburstEChartsSeries } from '../src/features/charts/echarts-sunburst-option.ts'
import { CHART_SUNBURST_BACKGROUND_KEY, CHART_SUNBURST_FIT_KEY } from '../src/features/charts/echarts-sunburst-layout.ts'
import { CHART_HIERARCHY_PATTERN_KEY } from '../src/features/charts/echarts-hierarchy-pattern.ts'
import { chartCommonControls, chartSeriesControls } from '../src/features/charts/chart-settings-controls.ts'
import { createChartSettingsDraft, chartSettingsPatch } from '../src/features/charts/chart-settings-draft.ts'
import { mergeChartOptions, validateChartOptions } from '../src/features/charts/chart-options.ts'
import { updateChart } from '../src/features/charts/chart-engine.ts'
import { estimateChartFinalLayoutMemory } from '../src/features/charts/chart-compute-memory.ts'

// P27 implementation drafts. No tests, compilers, Canvas, Worker or benchmarks ran.
const json = value => JSON.parse(JSON.stringify(value))
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-10, `${actual} ≈ ${expected}`)
const plot = { x: 10, y: 20, width: 600, height: 400 }
const context = { palette: ['#112233', '#ddeeff'], color: '#64748b', axisColor: '#64748b', splitLineColor: '#eeeeee',
  backgroundColor: '#ffffff', grid: { left: 12, right: 12, top: 12, bottom: 12, containLabel: false } }
const control = (controls, key) => { const found = controls.find(item => item.key === key); assert.ok(found, key); return found }
const gradient = () => ({ type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 1 },
  stops: [{ offset: 0, color: 'red', opacity: 0.4 }, { offset: 1, color: 'blue' }] })
const leaf = (name, value) => ({ id: name, name, value, selfValue: value, path: [name] })

test('signed parents retain semantic values, positive own facts become residuals and the center is 22', () => {
  const { data } = sunburstFixture(), before = json(data.hierarchy), roots = chartSunburstRenderNodes(data.hierarchy)
  assert.deepEqual(roots.map(node => [node.name, node.value, node.node.selfValue, node.rootIndex]), [['North', 13, -5, 1], ['South', 9, 3, 0]])
  assert.deepEqual(roots[0].children.map(node => [node.name, node.value]), [['Leaf B', 12], ['Leaf C', 6]])
  assert.equal(roots[0].children.some(node => node.isSunburstResidual), false)
  const residual = roots[1].children.find(node => node.isSunburstResidual)
  assert.equal(residual.id, CHART_SUNBURST_RESIDUAL_PREFIX + roots[1].id)
  assert.equal(residual.semanticParentId, roots[1].id); assert.equal(residual.name, ''); assert.equal(residual.value, 3)
  assert.equal(residual.node, roots[1].node); assert.deepEqual(residual.path, ['South'])
  assert.equal(chartSunburstTotal(data.hierarchy), 22); assert.equal(chartSunburstCenterText(data.hierarchy, {}), '22')
  assert.equal(chartSunburstCenterText(data.hierarchy, { format: '0.00' }), '22.00')
  assert.deepEqual(json(data.hierarchy), before)
})

test('global angle units preserve children wider than a negative-own parent and fixed 18/88 percent radii', () => {
  const { data } = sunburstFixture(), sectors = chartSunburstLayout(data.hierarchy, plot)
  const north = sectors.find(node => node.name === 'North'), south = sectors.find(node => node.name === 'South')
  near(north.start, -Math.PI / 2); near(north.end - north.start, 13 / 22 * Math.PI * 2)
  near(south.start, north.end); near(south.end - south.start, 9 / 22 * Math.PI * 2)
  const children = sectors.filter(node => node.depth === 1 && node.rootIndex === 1)
  near(children[0].start, north.start); near(children[1].end - children[0].start, 18 / 22 * Math.PI * 2)
  assert.ok(children[1].end > north.end)
  near(north.inner, 36); near(north.outer, 106); near(children[0].inner, 106); near(children[0].outer, 176)
  near(sectors.find(node => node.isSunburstResidual).end - sectors.find(node => node.isSunburstResidual).start, 3 / 22 * Math.PI * 2)
  assert.equal(Object.hasOwn(north.node, 'children'), false); assert.equal(Object.hasOwn(north.renderNode, 'children'), false)
  assert.deepEqual(chartSunburstLayout([], plot), []); assert.deepEqual(chartSunburstLayout(data.hierarchy, { ...plot, width: 0 }), [])
})

test('default labels and direct center fonts preserve legacy snapshot font inheritance', () => {
  const defaults = resolveChartSunburstLabel({}), center = resolveChartSunburstCenterLabel({})
  assert.equal(defaults.visible, true); assert.equal(defaults.contentType, 'category'); assert.equal(defaults.fontSize, 12); assert.equal(defaults.position, undefined)
  assert.equal(center.visible, false); assert.equal(center.fontSize, 16)
  const chart = { options: { sunburst: { centerLabel: { visible: true, font: { fontSize: 23, family: 'Legacy Font', bold: true }, fontSize: 19, bold: false } } } }
  const resolved = resolveChartSunburstCenterLabel(chart)
  assert.equal(resolved.fontSize, 19); assert.equal(resolved.family, 'Legacy Font'); assert.equal(resolved.bold, false)
  assert.equal(resolveChartSunburstLabel({}, { position: 'left' }).position, undefined)
  assert.equal(resolveChartSunburstLabel({ options: { sunburst: { label: { position: 'outside' } } } }).position, 'outside')
})

test('only category/value flags render complete strings, with residual, empty-mask and four-percent visibility rules', () => {
  const { data } = sunburstFixture(), roots = chartSunburstRenderNodes(data.hierarchy), north = roots[0]
  assert.equal(chartSunburstLabelText(north, { contentType: 10, format: '0.0' }), 'North\n13.0')
  assert.equal(chartSunburstLabelText(north, { contentType: 31 }), 'North\n13')
  assert.equal(chartSunburstLabelText(north, { contentType: 0 }), '')
  const residual = roots[1].children.find(node => node.isSunburstResidual)
  assert.equal(chartSunburstLabelText(residual, { contentType: 10 }), ''); assert.equal(chartSunburstLabelVisible(residual, 22, {}), false)
  assert.equal(chartSunburstLabelVisible({ ...north, value: 3.999 }, 100, {}), false)
  assert.equal(chartSunburstLabelVisible({ ...north, value: 4 }, 100, {}), true)
  assert.equal(chartSunburstLabelVisible(north, 22, { visible: false }), false)
})

test('depth tint and root tangential versus descendant radial labels preserve upright outside alignment', () => {
  assert.equal(chartSunburstShade('#000000', 1), '#2e2e2e'); assert.equal(chartSunburstShade('#000000', 3), '#8a8a8a')
  assert.equal(chartSunburstShade('#000000', 30), '#8a8a8a'); assert.equal(chartSunburstShade('navy', 3), 'navy')
  const center = { x: 50, y: 60 }, sector = { start: -Math.PI / 2, end: Math.PI / 2, inner: 100, outer: 200, depth: 0 }
  const root = chartSunburstLabelAnchor(center, sector), child = chartSunburstLabelAnchor(center, { ...sector, depth: 1 })
  near(root.x, 200); near(root.y, 60); near(root.rotation, -90); near(child.rotation, 0)
  const outside = chartSunburstLabelAnchor(center, { ...sector, depth: 1 }, 'outside')
  near(outside.x, 255); assert.equal(outside.align, 'left')
  const west = chartSunburstLabelAnchor(center, { ...sector, depth: 1, start: Math.PI / 2, end: Math.PI * 1.5 }, 'outside')
  near(west.x, -155); near(west.rotation, 0); assert.equal(west.align, 'right')
  for (const [bounds, expected] of [
    [chartSunburstSectorRect(center, { ...sector, start: -Math.PI / 2, end: 0 }), { x: 50, y: -140, width: 200, height: 200 }],
    [chartSunburstSectorRect(center, { ...sector, start: 0, end: Math.PI * 2 }), { x: -150, y: -140, width: 400, height: 400 }],
  ]) for (const key of ['x', 'y', 'width', 'height']) near(bounds[key], expected[key])
})

test('complete rotated bounds hide entire labels without truncating or squeezing their source content', () => {
  assert.equal(chartSunburstLabelFits({ width: 50, height: 20 }, { width: 50, height: 20 }), true)
  assert.equal(chartSunburstLabelFits({ width: 50, height: 20 }, { width: 50.001, height: 20 }), false)
  assert.equal(chartSunburstLabelFits({ width: 50, height: 20 }, { width: 40, height: 20.001 }), false)
  assert.equal(chartSunburstLabelFits(undefined, { width: 1, height: 1 }), false)
  assert.equal(chartSunburstLabelFits({ width: 50, height: 20 }, { width: Infinity, height: 1 }), false)
  const { data, chart } = sunburstFixture({ sunburst: { centerLabel: { visible: true } } })
  const name = '完整长标签'.repeat(200), hierarchy = [{ ...data.hierarchy[0], name, children: undefined }]
  const source = { ...data, hierarchy }, commands = []
  drawChartSunburst(commands, source, chart, { frame: plot, plot, sunburst: chartSunburstLayout(hierarchy, plot) }, '#64748b', context.palette)
  assert.equal(commands.some(command => command.kind === 'text' && command.text === name), false)
  assert.equal(commands.some(command => command.kind === 'text' && command.text === '13'), true)
  assert.equal(commands.filter(command => command.kind === 'text').some(command => command.maxWidth !== undefined), false)
})

test('native data uses flat semantic facts, explicit empty/residual state hides and actual label gradient/pattern markers', () => {
  const fill = { preset: 'cross', foregroundColor: 'black', backgroundColor: 'white' }
  const { data, chart } = sunburstFixture({ allSeriesStyle: { fillPattern: fill, fillOpacity: 0.4 }, sunburst: { label: { contentType: 10, backgroundColor: gradient() } } })
  const roots = chartSunburstEChartsData(data, chart, context), residual = roots[1].children.find(node => node.isSunburstResidual)
  assert.equal(roots[0].value, 13); assert.equal(roots[0].semanticSelfValue, -5); assert.equal(roots[0].semanticChildValue, 18)
  assert.equal(Object.hasOwn(roots[0], 'semanticNode'), false); assert.equal(roots[0].label.formatter(), 'North\n13')
  assert.equal(roots[0].label.rotate, 'tangential'); assert.equal(roots[0].children[0].label.rotate, 'radial')
  assert.equal(roots[0][CHART_SUNBURST_BACKGROUND_KEY].global, false)
  assert.equal(roots[0][CHART_HIERARCHY_PATTERN_KEY].foregroundColor, 'rgba(0, 0, 0, 0.4)')
  for (const state of [residual, ...chartSunburstEChartsData(data, { ...chart, options: { sunburst: { label: { contentType: 0 } } } }, context)]) {
    assert.equal(state.label.show, false); assert.equal(state.label.formatter(), '')
    for (const name of ['emphasis', 'select', 'blur']) assert.equal(state[name].label.show, false)
  }
  const series = createChartSunburstEChartsSeries(data, chart, context)
  assert.equal(series.type, 'sunburst'); assert.equal(series.animation, false); assert.equal(series.sort, null)
  assert.equal(series.nodeClick, false); assert.equal(series.stillShowZeroSum, false); assert.equal(series[CHART_SUNBURST_FIT_KEY], true)
  assert.deepEqual(series.center, [280, 200]); near(series.radius[0], 33.84); near(series.radius[1], 165.44)
})

test('native tooltip reports semantic signed parents, positive own residual values and center total independently', () => {
  const { data, chart } = sunburstFixture({ sunburst: { centerLabel: { visible: true, format: '0.00', family: 'Center Font', fontSize: 21, bold: true } } })
  const roots = chartSunburstEChartsData(data, chart, context), tooltip = chartSunburstTooltip(data, chart, context)
  assert.equal(tooltip.renderMode, 'richText')
  assert.equal(tooltip.showContent, false)
  assert.equal(tooltip.triggerOn, 'none')
  assert.equal(tooltip.formatter({ data: roots[0] }), 'North\nWeight: 13 (59.1%)\n自身数值: -5\n子节点合计: 18')
  assert.equal(tooltip.formatter({ data: roots[1].children.find(node => node.isSunburstResidual) }), 'South\nWeight: 3 (13.6%)\n自身数值: 3\n子节点合计: 6')
  assert.equal(tooltip.formatter({ data: { value: Infinity } }), '')
  const center = chartSunburstCenterGraphic(data, chart, context).elements[0]
  assert.equal(center.style.text, '22.00'); assert.equal(center.style.fontFamily, 'Center Font'); assert.equal(center.style.fontSize, 21)
  assert.equal(center.style.fontWeight, 'bold'); assert.equal(center.silent, true)
  assert.equal(chartSunburstCenterGraphic(data, { ...chart, options: undefined }, context), undefined)
})

test('center/node conditional controls and independent resets publish one isolated chart draft', () => {
  const { workbook, chart } = sunburstFixture({ sunburst: { label: { visible: true, contentType: 27, position: 'outside', format: '0.0' }, centerLabel: { visible: true, fontSize: 22 } } })
  const before = json(chart), draft = createChartSettingsDraft(chart), nodes = chartSunburstNodeLabelControls(), centers = chartSunburstCenterLabelControls()
  assert.deepEqual(control(nodes, 'sunburst.label.contentType').choices.map(choice => choice.value), ['category', 'value'])
  control(nodes, 'sunburst.label.contentType').write(draft.config, ['category']); assert.equal(draft.config.options.sunburst.label.contentType, 19)
  assert.equal(control(nodes, 'sunburst.label.format').visible(draft.config), false)
  control(nodes, 'sunburst.label.position').write(draft.config, 'auto'); assert.equal(draft.config.options.sunburst.label.position, undefined)
  control(centers, 'sunburst.centerLabel.fontSize').write(draft.config, 24)
  assert.equal(draft.config.options.sunburst.centerLabel.fontSize, 24); assert.equal(draft.config.options.sunburst.centerLabel.font, undefined)
  control(nodes, 'sunburst.label.visible').write(draft.config, false)
  assert.equal(control(nodes, 'sunburst.label.contentType').visible(draft.config), false)
  assert.equal(control(centers, 'sunburst.centerLabel.fontSize').visible(draft.config), true)
  clearChartSunburstSettings(draft.config, 'label'); assert.equal(draft.config.options.sunburst.label, undefined)
  assert.equal(draft.config.options.sunburst.centerLabel.fontSize, 24); assert.deepEqual(json(chart), before)
  updateChart(workbook, chart, chartSettingsPatch(draft)); assert.equal(chart.revision, before.revision + 1)
  assert.deepEqual(json(chart.options.sunburst), { centerLabel: { visible: true, fontSize: 24 } })
  const second = createChartSettingsDraft(chart); clearChartSunburstSettings(second.config, 'centerLabel'); updateChart(workbook, chart, chartSettingsPatch(second))
  assert.equal(resolveChartSunburstCenterLabel(chart).visible, false); assert.equal(resolveChartSunburstLabel(chart).visible, true)
  clearChartSunburstSettings(second.config, 'all'); assert.equal(second.config.options.sunburst, undefined)
  assert.equal(chartCommonControls(chart).some(item => item.key.startsWith('legend.')), false)
  for (const item of chartSeriesControls().filter(item => item.key.includes('.label.') || item.key.includes('.point.'))) assert.equal(item.visible(chart), false)
})

test('background union replacements detach callers and invalid backgrounds never create partial option groups', () => {
  const bare = { type: 'sunburst' }, before = json(bare)
  assert.throws(() => setChartSunburstBackground(bare, { type: 'unknown' })); assert.deepEqual(bare, before)
  const { chart } = sunburstFixture(), color = gradient(), pattern = { type: 'pattern', preset: 'pct25', foregroundColor: 'black', backgroundColor: 'white' }
  setChartSunburstBackground(chart, color); color.stops[0].color = 'lime'; assert.equal(chart.options.sunburst.label.backgroundColor.stops[0].color, 'red')
  setChartSunburstBackgroundMode(chart, 'pattern'); assert.equal(chartSunburstBackgroundMode(chart), 'pattern')
  const merged = mergeChartOptions({ sunburst: { label: { fontSize: 20, backgroundColor: gradient() }, centerLabel: { visible: true } } }, { sunburst: { label: { backgroundColor: pattern } } })
  assert.equal(merged.sunburst.label.fontSize, 20); assert.deepEqual(merged.sunburst.label.backgroundColor, pattern)
  assert.equal(Object.hasOwn(merged.sunburst.label.backgroundColor, 'stops'), false); assert.equal(merged.sunburst.centerLabel.visible, true)
  setChartSunburstBackgroundMode(chart, 'color'); assert.equal(chartSunburstBackgroundMode(chart), 'color')
  assert.doesNotThrow(() => validateChartOptions({ sunburst: { label: { backgroundColor: { ...pattern, preset: 'importedPattern' } } } }))
})

test('six public-operation inputs accept nullable direct fonts, detach values and reject invalid mappings/positions before commit', () => {
  assert.deepEqual(chartSunburstHierarchyFields([1, 0]), [1, 0]); assert.equal(chartSunburstValueField(2), 2)
  const color = gradient(), values = chartSunburstLabelValues({ backgroundColor: color, position: null, fontSize: 18 })
  color.stops[0].color = 'lime'; assert.equal(values.backgroundColor.stops[0].color, 'red'); assert.equal(values.position, undefined)
  assert.deepEqual(chartSunburstCenterLabelValues({ visible: true, family: 'Direct Font', fontSize: 20, bold: null }), { visible: true, family: 'Direct Font', fontSize: 20, bold: undefined })
  for (const invalid of [[], [0, 0], Array(2), [-1], [0.5]]) assert.throws(() => chartSunburstHierarchyFields(invalid))
  for (const value of [-1, 0.5, NaN, Infinity]) assert.throws(() => chartSunburstValueField(value))
  for (const label of [{ position: 'left' }, { fontSize: 0 }, { visible: 'yes' }, { contentType: -1 }, { unknown: true }]) assert.throws(() => chartSunburstLabelValues(label))
  assert.throws(() => chartSunburstCenterLabelValues({ font: { fontSize: 12 } })); assert.throws(() => assertChartSunburstType('treemap'))
})

test('Sunburst admits descendant/ring/gradient staging conservatively and rejects an unrepresentable semantic total', () => {
  const { data, chart } = sunburstFixture(), base = estimateChartFinalLayoutMemory(data, chart, {})
  const shallow = estimateChartFinalLayoutMemory({ ...data, hierarchy: undefined }, chart, {})
  assert.ok(base - shallow >= 6 * 2048)
  const painted = { ...chart, options: { ...chart.options, sunburst: { label: { backgroundColor: gradient() } } } }
  assert.ok(estimateChartFinalLayoutMemory(data, painted, {}) > base)
  assert.throws(() => chartSunburstTotal([leaf('A', Number.MAX_VALUE), leaf('B', Number.MAX_VALUE)]), /旭日图层级总量/)
  assert.ok(chartSunburstLayout([leaf('A', Number.MAX_VALUE / 4), leaf('B', Number.MAX_VALUE / 4)], plot).every(node => [node.start, node.end, node.inner, node.outer].every(Number.isFinite)))
})
