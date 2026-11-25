import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P17 drafts only. No runner, Vite process, browser, type check or build ran.
let server, routes, controls, typeControls, targets, echartsStyle
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[routes, controls, typeControls, targets, echartsStyle] = await Promise.all([
    'chart-edit-panel-route', 'chart-settings-controls', 'chart-type-settings-controls', 'chart-series-settings', 'echarts-style',
  ].map(name => server.ssrLoadModule(`/packages/office-studio/src/features/charts/${name}.ts`)))
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function chart(type = 'line', options = {}) {
  return { type, source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 } },
    categoryField: 0, valueFields: [1, 2], title: 'Original', subtitle: 'Sub', position: { x: 16, y: 24 }, size: { width: 480, height: 320 },
    legend: { visible: true, position: 'bottom' }, xAxis: { title: 'Category', scaleType: 'linear' }, yAxis: { title: 'Primary', scaleType: 'linear' }, options }
}
function data(series = [{ field: 1, name: 'Value', values: [7, 4] }, { field: 2, name: 'Renamed comparison', sourceName: 'Comparison', values: [3, 9] }]) {
  return { categories: ['North', 'South'], fields: ['Label', 'Value', 'Comparison'], records: [['North', 7, 3], ['South', 4, 9]], series }
}
function plan(model, route, source = data()) { return routes.planChartEditPanelRoute(model, route, source) }

test('route defaults infer the title, axis and series groups without mutating the caller DTO', () => {
  const normalize = routes.normalizeChartEditPanelRoute, input = { tab: 'style', titleKey: 'rightYAxisTitle' }, before = json(input)
  assert.deepEqual(normalize({ tab: 'style' }), { tab: 'style', section: 'chartStyle' })
  assert.deepEqual(normalize({ tab: 'data' }), { tab: 'data' })
  assert.deepEqual(normalize({ tab: 'style', axisId: 'yAxis' }), { tab: 'style', section: 'verticalAxis', axisId: 'yAxis' })
  assert.deepEqual(normalize({ tab: 'style', seriesId: 'field:2' }), { tab: 'style', section: 'series', seriesId: 'field:2' })
  assert.deepEqual(normalize(input), { tab: 'style', section: 'chartAndAxisTitles', titleKey: 'rightYAxisTitle', axisId: 'rightYAxis' })
  assert.deepEqual(input, before)
})

test('malformed or unknown route fields fail explicitly, including symbols and non-plain DTOs', () => {
  const invalid = [undefined, null, [], new Date(), Object.create(null), { tab: 'format' }, { tab: 'style', section: 'absent' },
    { tab: 'style', axisId: 'depthAxis' }, { tab: 'style', titleKey: 'caption' }, { tab: 'style', seriesId: 2 },
    { tab: 'style', seriesId: ' ' }, { tab: 'style', seriesId: 'x'.repeat(16_385) }, { tab: 'style', hidden: true },
    { tab: 'style', [Symbol('private')]: 'value' }]
  for (const input of invalid) assert.throws(() => routes.normalizeChartEditPanelRoute(input), TypeError)
})

test('contradictory tab, group, axis, title and series combinations cannot silently choose a different section', () => {
  for (const route of [
    { tab: 'data', section: 'chartStyle' }, { tab: 'data', seriesId: 'field:1' },
    { tab: 'style', section: 'legend', titleKey: 'title' }, { tab: 'style', section: 'legend', axisId: 'yAxis' },
    { tab: 'style', section: 'horizontalAxis', axisId: 'yAxis' }, { tab: 'style', section: 'chartAndAxisTitles', axisId: 'yAxis', titleKey: 'title' },
    { tab: 'style', section: 'chartAndAxisTitles', axisId: 'xAxis', titleKey: 'yAxisTitle' },
    { tab: 'style', section: 'pieStyle', seriesId: 'field:1' },
  ]) assert.throws(() => routes.normalizeChartEditPanelRoute(route), TypeError)
})

test('local series IDs keep exact field/source-name identity, case, spaces and empty names', () => {
  assert.equal(routes.chartEditPanelSeriesId(0), 'field:0')
  assert.equal(routes.chartEditPanelSeriesId(2), 'field:2')
  for (const name of ['', 'A', 'a', ' A ', '__proto__', 'A"B\\C']) {
    const id = routes.chartEditPanelSeriesId({ name })
    assert.deepEqual(routes.resolveChartEditPanelSeries(id, data([{ field: 1, name: 'Displayed', sourceName: name, values: [1, 2] }])), { name })
  }
  for (const value of [-1, 1.5, NaN, 1_048_576]) assert.throws(() => routes.chartEditPanelSeriesId(value), RangeError)
})

test('duplicate names require an explicit field, and unresolved, non-canonical or unready series never fall back by position', () => {
  const source = data([{ field: 1, name: 'A', values: [1, 2] }, { field: 1, name: 'A', values: [3, 4] }])
  assert.equal(routes.resolveChartEditPanelSeries('field:1', source), 1)
  assert.throws(() => routes.resolveChartEditPanelSeries('name:"A"', source), /名称重复/)
  assert.throws(() => routes.resolveChartEditPanelSeries('field:1'), /尚未就绪/)
  for (const id of ['field:01', 'field:2', 'field:-1', 'field:1048576', 'name: "A"', 'name:1', 'runtime-opaque-series-id']) {
    assert.throws(() => routes.resolveChartEditPanelSeries(id, source), /不存在|标识无效/)
  }
})

test('all five title routes select their own real title controls and leave authored content unchanged', () => {
  const model = chart('line', { series: { 2: { rightAxis: true } }, rightYAxis: { title: 'Secondary', scaleType: 'linear' } }), before = json(model)
  const cases = [
    ['title', ['common', 'fonts'], [], 'title'], ['subtitle', ['common', 'fonts'], [], 'subtitle'],
    ['xAxisTitle', ['axes'], ['x'], 'xAxis.title'], ['yAxisTitle', ['axes'], ['y'], 'yAxis.title'],
    ['rightYAxisTitle', ['axes'], ['right'], 'rightYAxis.title'],
  ]
  for (const [titleKey, panels, axes, optionKey] of cases) {
    const result = plan(model, { tab: 'style', titleKey })
    assert.deepEqual(result.panels, panels); assert.deepEqual(result.axes, axes); assert.equal(result.optionKey, optionKey)
    const actual = titleKey === 'title' || titleKey === 'subtitle' ? controls.chartCommonControls(model)
      : controls.chartAxisControls(result.route.axisId)
    assert.equal(actual.find(control => control.key === optionKey)?.kind, 'text')
  }
  assert.deepEqual(model, before)
})

test('three axis and gridline routes address the correct axis; absent secondary axes and non-axis chart types fail', () => {
  const model = chart('line', { series: { 2: { rightAxis: true } } })
  for (const [axisId, section, group] of [['xAxis', 'horizontalAxis', 'x'], ['yAxis', 'verticalAxis', 'y'], ['rightYAxis', 'rightVerticalAxis', 'right']]) {
    const axis = plan(model, { tab: 'style', section, axisId }), gridline = plan(model, { tab: 'style', section: 'gridlinesAndTicks', axisId })
    assert.deepEqual(axis.axes, [group]); assert.equal(axis.optionKey, `${axisId}.title`)
    assert.deepEqual(gridline.axes, [group]); assert.equal(gridline.optionKey, `${axisId}.gridLine`)
    assert.ok(controls.chartAxisControls(axisId).some(control => control.key.startsWith(`${gridline.optionKey}.`)))
  }
  assert.throws(() => plan(chart(), { tab: 'style', axisId: 'rightYAxis' }), /编辑分组|坐标轴/)
  assert.throws(() => plan(chart('pie'), { tab: 'style', titleKey: 'xAxisTitle' }), /坐标轴/)
  assert.throws(() => plan(chart('line3D', { series: { 2: { rightAxis: true } } }), { tab: 'style', axisId: 'rightYAxis' }), /编辑分组|坐标轴/)
})

test('line/area/combination and derived routes reach the real shape control while Scatter retains the null-value control', () => {
  for (const type of ['line', 'lineStacked', 'linePercentStacked', 'line3D', 'combination', 'area', 'areaStacked', 'areaPercentStacked', 'area3D', 'areaStacked3D', 'areaPercentStacked3D']) {
    const model = chart(type), route = plan(model, { tab: 'style', section: 'lineAndArea' })
    assert.deepEqual(route.panels, ['common', 'type', 'series']); assert.deepEqual(route.series, ['all'])
    assert.equal(route.optionKey, 'area.lineStyle')
    const shape = typeControls.chartTypeSettingsControls(model).find(control => control.key === route.optionKey)
    assert.equal(shape?.kind, 'select'); shape.write(model, 'step'); assert.equal(model.options.area.lineStyle, 'step')
    const nullControl = controls.chartCommonControls(model).find(control => control.key === 'invalidValueStrategy')
    assert.equal(nullControl?.kind, 'select'); nullControl.write(model, 'link'); assert.equal(model.options.invalidValueStrategy, 'link')
    const control = controls.chartSeriesControls(undefined, type === 'combination', true).find(control => control.key === 'allSeriesStyle.border.width')
    assert.equal(control?.kind, 'number'); control.write(model, 8)
    assert.equal(echartsStyle.chartEChartsLineStyle(model.options.allSeriesStyle).width, 8)
  }
  const scatter = plan(chart('scatter'), { tab: 'style', section: 'lineAndArea' })
  assert.deepEqual(scatter.panels, ['common']); assert.equal(scatter.optionKey, 'invalidValueStrategy')
  assert.equal(typeControls.chartTypeSettingsControls(chart('scatter')).some(control => control.key === 'area.lineStyle'), false)
})

test('the 27 declared sections have explicit applicable chart families and reject requests against incompatible families', () => {
  const cases = [
    ['chartStyle', 'line'], ['chartAndAxisTitles', 'line'], ['legend', 'line'], ['horizontalAxis', 'line'], ['verticalAxis', 'line'],
    ['rightVerticalAxis', 'line'], ['series', 'line'], ['pieStyle', 'pie'], ['paretoBarSeries', 'pareto'], ['paretoLineSeries', 'pareto'],
    ['waterfallSeries', 'waterfall'], ['relation', 'relation'], ['wordCloud', 'wordCloud'], ['radar', 'radar'], ['funnel', 'funnel'],
    ['heatmap', 'heatmap'], ['lineAndArea', 'line'], ['gridlinesAndTicks', 'line'], ['indicatorLine', 'line'], ['trendline', 'line'],
    ['errorBars', 'line'], ['candlestick', 'candlestick'], ['histogram', 'histogram'], ['treemap', 'treemap'], ['sunburst', 'sunburst'], ['gauge', 'gauge'], ['chord', 'chord'],
  ]
  assert.deepEqual(cases.map(([section]) => section), [...routes.CHART_EDIT_PANEL_SECTIONS])
  for (const [section, type] of cases) {
    const model = chart(type, section === 'rightVerticalAxis' ? { series: { 2: { rightAxis: true } } } : {}), before = json(model)
    assert.ok(plan(model, { tab: 'style', section }).panels.length > 0); assert.deepEqual(model, before)
  }
  for (const section of ['pieStyle', 'paretoBarSeries', 'paretoLineSeries', 'waterfallSeries', 'relation', 'wordCloud', 'radar', 'funnel', 'heatmap', 'candlestick', 'histogram', 'treemap', 'sunburst', 'gauge', 'chord']) {
    assert.throws(() => plan(chart(), { tab: 'style', section }), /当前图表类型/)
  }
})

test('statistics routes use existing trendline/error groups without adding authors, and stacked/3D statistics remain unsupported', () => {
  const model = chart(), before = json(model)
  assert.equal(plan(model, { tab: 'style', section: 'trendline' }).statisticsKind, 'trendlines')
  assert.equal(plan(model, { tab: 'style', section: 'errorBars' }).statisticsKind, 'errorBars')
  assert.deepEqual(model, before); assert.equal(model.options.trendlines, undefined); assert.equal(model.options.errorBars, undefined)
  for (const type of ['lineStacked', 'columnPercentStacked', 'area3D']) for (const section of ['trendline', 'errorBars']) {
    assert.throws(() => plan(chart(type), { tab: 'style', section }), /当前图表类型/)
  }
})

test('the 11 public context scopes preserve v1.6.0 axis/group/series/title route semantics', () => {
  const cases = [
    [{ scope: 'chartArea' }, { tab: 'style', section: 'chartStyle' }],
    [{ scope: 'title', titleKey: 'subtitle' }, { tab: 'style', section: 'chartAndAxisTitles', titleKey: 'subtitle' }],
    [{ scope: 'axisTitle', axisId: 'yAxis', titleKey: 'yAxisTitle' }, { tab: 'style', section: 'chartAndAxisTitles', axisId: 'yAxis', titleKey: 'yAxisTitle' }],
    [{ scope: 'axis', axisId: 'xAxis' }, { tab: 'style', section: 'horizontalAxis', axisId: 'xAxis' }],
    [{ scope: 'gridLine', axisId: 'yAxis' }, { tab: 'style', section: 'verticalAxis', axisId: 'yAxis' }],
    [{ scope: 'tick', axisId: 'rightYAxis' }, { tab: 'style', section: 'rightVerticalAxis', axisId: 'rightYAxis' }],
    [{ scope: 'legend' }, { tab: 'style', section: 'legend' }],
    [{ scope: 'mark', panel: 'data' }, { tab: 'data' }],
    [{ scope: 'series', seriesId: 'field:2' }, { tab: 'style', section: 'series', seriesId: 'field:2' }],
    [{ scope: 'dataPoint', seriesId: 'field:2' }, { tab: 'style', section: 'series', seriesId: 'field:2' }],
    [{ scope: 'dataLabel', seriesId: 'field:2' }, { tab: 'style', section: 'series', seriesId: 'field:2' }],
  ]
  for (const [target, route] of cases) assert.deepEqual(routes.resolveChartEditPanelRoute(target), route)
  assert.equal(routes.resolveChartEditPanelRoute({ scope: 'dataLabel' }), undefined)
})

test('mark direct-edit precedence, seven panel targets and missing-series fallbacks follow the fixed resolver', () => {
  const resolveRoute = routes.resolveChartEditPanelRoute
  for (const panel of ['pieStyle', 'paretoBarSeries', 'paretoLineSeries', 'waterfallSeries', 'relation', 'wordCloud', 'data']) {
    const expected = panel === 'data' ? { tab: 'data' } : { tab: 'style', section: panel, ...(panel === 'waterfallSeries' ? { seriesId: 'field:2' } : {}) }
    assert.deepEqual(resolveRoute({ scope: 'mark', panel, seriesId: 'field:2' }), expected)
  }
  assert.deepEqual(resolveRoute({ scope: 'mark', panel: 'data', directEdit: { kind: 'seriesStyle', seriesId: 'field:1' } }), { tab: 'style', section: 'series', seriesId: 'field:1' })
  assert.deepEqual(resolveRoute({ scope: 'mark', directEdit: { kind: 'pieStyle' } }), { tab: 'style', section: 'pieStyle' })
  for (const part of ['bar', 'line']) assert.deepEqual(resolveRoute({ scope: 'mark', directEdit: { kind: 'paretoStyle', part } }), { tab: 'style', section: part === 'bar' ? 'paretoBarSeries' : 'paretoLineSeries' })
  assert.deepEqual(resolveRoute({ scope: 'mark', directEdit: { kind: 'waterfallStyle', seriesId: 'field:2' } }), { tab: 'style', section: 'waterfallSeries', seriesId: 'field:2' })
  assert.equal(resolveRoute({ scope: 'mark', directEdit: { kind: 'pieConnectorStyle' } }), undefined)
  assert.deepEqual(resolveRoute({ scope: 'mark', directEdit: { kind: 'pieSliceStyle' }, panel: 'pieStyle' }), { tab: 'style', section: 'pieStyle' })
  assert.equal(resolveRoute({ scope: 'mark' }), undefined)
})

test('route-selected source names remain stable after series reordering and display renaming in the existing selector model', () => {
  const model = chart('bubble', { mapping: { bubble: { xIndex: 0, yIndex: 1, sizeIndex: 2, seriesIndex: 3 } } }), source = data([
    { field: 1, name: 'Display B', sourceName: 'B', values: [1, 2] }, { field: 1, name: 'Display A', sourceName: 'A', values: [3, 4] },
  ]), route = { tab: 'style', section: 'series', seriesId: 'name:"A"' }
  const first = plan(model, route, source), second = plan(model, route, { ...source, series: [...source.series].reverse() })
  assert.equal(first.seriesTarget, 'name:"A"'); assert.equal(second.seriesTarget, first.seriesTarget)
  const actual = targets.chartSeriesSettingTargets(model, [{ value: 1, label: 'Y' }], source).find(target => target.value === first.seriesTarget)
  assert.deepEqual(actual.selector, { name: 'A' }); assert.equal(actual.source.name, 'A'); assert.equal(actual.source.position, 1)
  assert.throws(() => plan(model, { ...route, seriesId: 'name:"Display A"' }, source), /不存在/)
})
