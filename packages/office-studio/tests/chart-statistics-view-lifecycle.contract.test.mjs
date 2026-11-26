import assert from 'node:assert/strict'
import { afterEach, test, vi } from 'vitest'
import { effectScope, nextTick, shallowRef } from 'vue'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { OfficeComputeClient } from '../src/compute/client.ts'
import { createChartDefinition } from '../src/features/charts/chart-engine.ts'
import { ChartStatisticsService, chartStatisticsSignature } from '../src/features/charts/services/chart-statistics-service.ts'
import { useChartStatistics } from '../src/features/charts/views/use-chart-statistics.ts'

// Deferred contracts isolate Vue host lifecycle from Worker computation. The
// service's revision notifications are real; only its prepared cache is supplied.
const hooks = vi.hoisted(() => ({ mounted: [], activated: [], deactivated: [], unmounted: [] }))
vi.mock('vue', async importOriginal => ({
  ...await importOriginal(),
  onMounted: callback => hooks.mounted.push(callback),
  onActivated: callback => hooks.activated.push(callback),
  onDeactivated: callback => hooks.deactivated.push(callback),
  onUnmounted: callback => hooks.unmounted.push(callback),
}))
const cleanups = []
afterEach(() => {
  for (const callback of hooks.unmounted) callback()
  for (const cleanup of cleanups.splice(0)) cleanup()
  for (const callbacks of Object.values(hooks)) callbacks.length = 0
  vi.restoreAllMocks()
})
async function settled() { await nextTick(); await nextTick() }
function fixture() {
  const initial = createBlankWorkbook(); initial.id = 'chart-view-lifecycle'
  const sheet = initial.sheets[0]; sheet.id = 'host'; initial.activeSheetId = sheet.id
  sheet.cells = { '0:0': createCell(0, 0, 'Label'), '0:1': createCell(0, 1, 'Value'),
    '1:0': createCell(1, 0, 'North'), '1:1': createCell(1, 1, 7) }
  initial.sheets.push({ ...sheet, id: 'other', cells: { ...sheet.cells } })
  const workbook = shallowRef(initial), activeSheet = shallowRef('host'), scope = effectScope()
  const client = new OfficeComputeClient(), service = new ChartStatisticsService(client)
  let revision = 0
  service.setDocumentRevision(initial.id, revision)
  const subscribe = vi.spyOn(service, 'subscribe')
  const cache = vi.spyOn(service, 'getCached').mockImplementation((chart, context) => ({
    documentId: initial.id, modelRevision: context.modelRevision, chartId: chart.id, chartRevision: chart.revision,
    signature: chartStatisticsSignature(chart), preparationMs: 0, decodingMs: 0, estimatedCacheBytes: 64,
    statistics: { preparedData: { categories: ['North'], series: [{ name: 'Value', field: 1, values: [7] }] } },
  }))
  const studio = { chartStatisticsService: service, currentUserId: 'local-user', getWorkbookRevision: () => revision }
  const view = scope.run(() => useChartStatistics({ studio: () => studio, workbook: () => workbook.value,
    activeSheetId: () => activeSheet.value, externalStatistics: () => undefined }))
  cleanups.push(() => { scope.stop(); service.dispose(); client.dispose() })
  const chart = (id, sheetId = 'host', type = 'column') => createChartDefinition(workbook.value, sheetId, {
    type, source: { orientation: 'columns', range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 } },
    categoryField: 0, valueFields: [1], position: { x: 64, y: 48 }, size: { width: 420, height: 300 },
  }, id)
  function publish(charts) { workbook.value = { ...workbook.value, charts }; service.setDocumentRevision(initial.id, ++revision) }
  return { view, subscribe, cache, chart, publish }
}

test('the first chart on an initially empty workbook receives current prepared data without a host change', async () => {
  const state = fixture(); for (const callback of hooks.mounted) callback(); await settled()
  assert.equal(state.subscribe.mock.calls.length, 1); assert.deepEqual(state.view.statistics.value, {})
  const first = state.chart('first'); state.publish([first]); await settled()
  assert.deepEqual(state.view.statistics.value.first.preparedData.series[0].values, [7])
  assert.equal(state.view.snapshots.value.first.modelRevision, 1)
  state.publish([]); await settled(); assert.deepEqual(state.view.statistics.value, {})
  state.publish([state.chart('replacement', 'host', 'line')]); await settled()
  assert.ok(state.view.statistics.value.replacement); assert.equal(state.view.statistics.value.first, undefined)
  assert.equal(state.subscribe.mock.calls.length, 1)
})

test('an empty active sheet remains connected when the workbook has a chart on another sheet', async () => {
  const state = fixture(); state.publish([state.chart('other-chart', 'other')])
  for (const callback of hooks.mounted) callback(); await settled()
  assert.deepEqual(state.view.statistics.value, {}); assert.equal(state.cache.mock.calls.length, 0)
  state.publish([state.chart('other-chart', 'other'), state.chart('active-chart')]); await settled()
  assert.deepEqual(Object.keys(state.view.statistics.value), ['active-chart'])
  assert.equal(state.view.snapshots.value['active-chart'].modelRevision, 2)
})

test('deactivation releases an empty view subscription and activation sees charts inserted while hidden', async () => {
  const state = fixture(); for (const callback of hooks.mounted) callback(); await settled()
  for (const callback of hooks.deactivated) callback()
  state.publish([state.chart('hidden-insertion')]); await settled()
  assert.deepEqual(state.view.statistics.value, {}); assert.equal(state.cache.mock.calls.length, 0)
  for (const callback of hooks.activated) callback(); await settled()
  assert.ok(state.view.statistics.value['hidden-insertion']); assert.equal(state.subscribe.mock.calls.length, 2)
})
