import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartCanvasElementAuthorSignature, chartCanvasElementControls, chartCanvasElementDeleteDraft, chartCanvasElementHidden,
  chartCanvasElementPatch, resetChartCanvasElementDraft } from '../src/features/charts/chart-element-edit.ts'
import { chartCanvasSeriesBindings } from '../src/features/charts/chart-element-interaction.ts'
import { cloneChartDraftConfig } from '../src/features/charts/chart-settings-draft.ts'
import { mergeChartOptions } from '../src/features/charts/chart-options.ts'

// P19 scoped-authoring drafts; no execution or acceptance is claimed.
const chart = () => ({ id: 'direct', sheetId: 'sheet', type: 'line', revision: 0, categoryField: 0, valueFields: [1, 2],
  source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 } }, position: { x: 8, y: 9 }, size: { width: 480, height: 320 },
  title: 'Chart', legend: { visible: true, position: 'bottom' }, xAxis: { title: 'Categories', scaleType: 'linear', line: { color: '#111111' } },
  yAxis: { title: 'Values', scaleType: 'linear' }, options: { series: { 1: { name: 'Displayed', color: '#123456', rightAxis: true,
    label: { visible: true, fontSize: 12 }, point: { size: 7 }, dataPoints: { 0: { color: '#555555' }, 1: { color: '#666666', border: { width: 3 } } } },
    2: { color: '#abcdef' } }, rightYAxis: { title: 'Secondary', scaleType: 'linear' } } })
const binding = () => chartCanvasSeriesBindings({ series: [{ field: 1, name: 'Displayed', sourceName: 'Original', values: [2, 4, 6] }] })[0]
const point = () => ({ kind: 'dataPoint', binding: binding(), dataIndex: 1 })
const json = value => JSON.parse(JSON.stringify(value))

test('axis/title/grid/tick controls expose only their real author fields and fixed label rotation choices', () => {
  const model = chart(), title = chartCanvasElementControls({ kind: 'axisTitle', axisId: 'rightYAxis' }, model)
  assert.ok(title.length > 0 && title.every(control => control.key.startsWith('rightYAxis.title')))
  const axis = chartCanvasElementControls({ kind: 'axis', axisId: 'xAxis', part: 'label' }, model)
  assert.ok(axis.every(control => !control.key.startsWith('xAxis.title') && !control.key.startsWith('xAxis.gridLine') && !control.key.startsWith('xAxis.tick')))
  assert.deepEqual(axis.find(control => control.key === 'xAxis.label.rotate').choices.map(choice => choice.value), ['0', '30', '45', '60', '90'])
  for (const kind of ['gridLine', 'tick']) assert.ok(chartCanvasElementControls({ kind, axisId: 'yAxis' }, model).every(control => control.key.startsWith(`yAxis.${kind}.`)))
})

test('axis title apply ignores unrelated draft edits and preserves the original line/source/legend values', () => {
  const model = chart(), before = json(model), draft = cloneChartDraftConfig(model), element = { kind: 'axisTitle', axisId: 'xAxis' }
  draft.xAxis.title = 'Authored axis'; draft.xAxis.titleAlignment = 'end'; draft.legend.position = 'left'; draft.options.series[2].color = '#ffffff'
  const patch = chartCanvasElementPatch(model, draft, element)
  assert.equal(patch.xAxis.title, 'Authored axis'); assert.equal(patch.xAxis.titleAlignment, 'end'); assert.deepEqual(patch.xAxis.line, model.xAxis.line)
  assert.equal(patch.legend, undefined); assert.equal(patch.source, undefined); assert.equal(patch.options, undefined); assert.deepEqual(json(model), before)
  const unrelated = cloneChartDraftConfig(model); unrelated.legend.position = 'right'
  assert.deepEqual(chartCanvasElementPatch(model, unrelated, element), {})
})

test('point edits and null resets retain sibling points, other series and unrelated selected-series authorship', () => {
  const model = chart(), draft = cloneChartDraftConfig(model), element = point()
  draft.options.series[1].dataPoints[1].color = '#ee0000'; draft.options.series[1].name = 'Ignored rename'
  const patch = chartCanvasElementPatch(model, draft, element), merged = mergeChartOptions(model.options, patch.options)
  assert.equal(merged.series[1].dataPoints[1].color, '#ee0000'); assert.equal(merged.series[1].dataPoints[1].border.width, 3)
  assert.deepEqual(merged.series[1].dataPoints[0], model.options.series[1].dataPoints[0]); assert.equal(merged.series[1].name, 'Displayed')
  assert.deepEqual(merged.series[2], model.options.series[2])
  const reset = resetChartCanvasElementDraft(model, element), resetOptions = mergeChartOptions(model.options, chartCanvasElementPatch(model, reset, element).options)
  assert.equal(resetOptions.series[1].dataPoints[1], undefined)
  assert.deepEqual(json(resetOptions).series[1].dataPoints, { 0: { color: '#555555' } })
  assert.equal(resetOptions.series[1].rightAxis, true)
})

test('duplicate-field named point editing keeps the shared field layer and other named series intact', () => {
  const model = chart(), sources = { series: [{ field: 1, name: 'Alpha', values: [1] }, { field: 1, name: 'Beta', values: [2] }] }
  model.options.namedSeries = [{ name: 'Alpha', style: { dataPoints: { 0: { color: '#112233' } } } }, { name: 'Beta', style: { color: '#334455' } }]
  const element = { kind: 'dataPoint', binding: chartCanvasSeriesBindings(sources)[0], dataIndex: 0 }, draft = cloneChartDraftConfig(model)
  draft.options.namedSeries[0].style.dataPoints[0].fillOpacity = 0.5
  const patch = chartCanvasElementPatch(model, draft, element), merged = mergeChartOptions(model.options, patch.options)
  assert.equal(patch.options.series, undefined); assert.deepEqual(merged.series, model.options.series)
  assert.equal(merged.namedSeries[0].style.dataPoints[0].fillOpacity, 0.5); assert.deepEqual(merged.namedSeries[1], model.options.namedSeries[1])
})

test('ordinary data labels author the selected series while pie labels author the global pie-label group', () => {
  const model = chart(), element = { ...point(), kind: 'dataLabel' }, draft = cloneChartDraftConfig(model)
  draft.options.series[1].label.fontSize = 18; draft.options.series[1].dataPoints[1].color = '#ffffff'
  const next = mergeChartOptions(model.options, chartCanvasElementPatch(model, draft, element).options)
  assert.equal(next.series[1].label.fontSize, 18); assert.deepEqual(next.series[1].dataPoints, model.options.series[1].dataPoints)
  const pie = { ...model, type: 'pie', options: { ...model.options, pie: { label: { visible: true, fontSize: 10 } } } }, pieDraft = cloneChartDraftConfig(pie)
  pieDraft.options.pie.label.fontSize = 20; pieDraft.options.series[1].label.fontSize = 99
  const piePatch = chartCanvasElementPatch(pie, pieDraft, element)
  assert.deepEqual(piePatch.options, { pie: { label: { visible: true, fontSize: 20 } } })
  assert.ok(chartCanvasElementControls(element, pie).every(control => control.key.startsWith('pie.label.')))
})

test('format reset preserves series name/type/axis assignment and point overrides, and point signatures ignore series-only drafts', () => {
  const model = chart(), element = { kind: 'series', binding: binding() }; model.options.series[1].type = 'line'
  const reset = resetChartCanvasElementDraft(model, element), next = mergeChartOptions(model.options, chartCanvasElementPatch(model, reset, element).options)
  for (const key of ['color', 'fillOpacity', 'fillPattern', 'border', 'point', 'label']) assert.equal(next.series[1][key], undefined)
  assert.equal(next.series[1].name, 'Displayed'); assert.equal(next.series[1].type, 'line'); assert.equal(next.series[1].rightAxis, true)
  assert.deepEqual(next.series[1].dataPoints, model.options.series[1].dataPoints)
  assert.equal(chartCanvasElementAuthorSignature(model, point()), chartCanvasElementAuthorSignature({ ...model, options: { ...model.options,
    series: { ...model.options.series, 1: { ...model.options.series[1], color: '#ffffff' } } } }, point()))
})

test('Delete hides supported chrome/labels without changing source data, while point and series selections consume no destructive author patch', () => {
  const model = chart(), before = json(model)
  for (const element of [{ kind: 'axisTitle', axisId: 'xAxis' }, { kind: 'axis', axisId: 'yAxis', part: 'label' },
    { kind: 'gridLine', axisId: 'yAxis' }, { kind: 'tick', axisId: 'xAxis' }, { kind: 'legend' }, { ...point(), kind: 'dataLabel' }]) {
    const draft = chartCanvasElementDeleteDraft(model, element)
    assert.ok(chartCanvasElementHidden(draft, element)); assert.deepEqual(draft.source, model.source); assert.deepEqual(draft.valueFields, model.valueFields)
  }
  assert.equal(chartCanvasElementDeleteDraft(model, point()), undefined)
  assert.equal(chartCanvasElementDeleteDraft(model, { kind: 'series', binding: binding() }), undefined)
  assert.deepEqual(json(model), before)
})
