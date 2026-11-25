import assert from 'node:assert/strict'
import { afterAll, beforeAll, test, vi } from 'vitest'
import { effectScope, nextTick, ref, shallowRef } from 'vue'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Draft for unified validation. No tests or server have been run in this phase.
let server, useSheetChartPanels, FOffice, EventHub
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  const modules = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/features/charts/views/use-sheet-chart-panels.ts'),
    server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'),
    server.ssrLoadModule('/packages/office-studio/src/api/events.ts'),
  ])
  ;[{ useSheetChartPanels }, { FOffice }, { EventHub }] = modules
})
afterAll(async () => { vi.restoreAllMocks(); await server?.close() })

function deferred() {
  let resolve, reject
  const promise = new Promise((yes, no) => { resolve = yes; reject = no })
  return { promise, resolve, reject }
}
function config() {
  return { type: 'column', source: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, orientation: 'columns' }, position: { x: 0, y: 0 }, size: { width: 480, height: 300 }, valueFields: [1] }
}
function model(sheetId = 'sheet-a', id = 'chart-a') {
  return { ...config(), id, sheetId, revision: 0, categoryField: 0, legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' } }
}
function fixture(handlers = {}) {
  const makeStudio = () => {
    const hub = new EventHub(), subscriptions = new Set()
    return { hub, subscriptions, events: { addEvent(name, listener) {
      const handle = hub.addEvent(name, listener)
      const disposable = { dispose() { handle.dispose(); subscriptions.delete(disposable) } }
      subscriptions.add(disposable)
      return disposable
    } } }
  }
  const firstStudio = makeStudio(), secondStudio = makeStudio()
  const studio = shallowRef(firstStudio), unitId = ref('unit-a'), sheetId = ref('sheet-a'), active = ref(true), canFocus = ref(true)
  const calls = [], sourceCalls = [], statisticsCalls = [], refreshes = [], notices = []
  let focusCount = 0, dataReads = 0
  const spy = vi.spyOn(FOffice, 'newAPI').mockImplementation(owner => ({ getWorkbook(unit) {
    return { getSheetBySheetId(sheet) {
      return {
        insertChart(value) { calls.push({ kind: 'insert', owner, unit, sheet, value }); return handlers.insert?.({ owner, unit, sheet, value }) ?? Promise.resolve() },
        getChartSettingsData(value, options) {
          sourceCalls.push({ owner, unit, sheet, value, signal: options.signal })
          if (handlers.readError) return Promise.reject(handlers.readError)
          return handlers.source?.({ owner, unit, sheet, value, options }) ?? Promise.resolve({ fields: ['Category', 'Value'], records: [['A', 10]],
            sourceNumericFields: [1], categories: ['A'], series: [{ field: 1, name: 'Value', values: [10] }] })
        },
        getChartSettingsStatistics(value, options) {
          statisticsCalls.push({ owner, unit, sheet, value, signal: options.signal })
          if (handlers.readError) return Promise.reject(handlers.readError)
          return handlers.statistics?.({ owner, unit, sheet, value, options }) ?? Promise.resolve({
            preparedData: { fields: ['Category', 'Value'], records: [['A', 10]], sourceNumericFields: [1],
              categories: ['A'], series: [{ field: 1, name: 'Value', values: [10] }] },
            pareto: { categories: ['A'], values: [10], cumulativePercent: [100] },
          })
        },
        getChartById(id) {
          return {
            update(patch) { calls.push({ kind: 'update', owner, unit, sheet, id, patch }); return handlers.update?.({ owner, unit, sheet, id, patch }) ?? Promise.resolve() },
            getInfo() { return model(sheet, id) },
            getData() { dataReads++; throw new Error('Settings panels must use the bounded draft-source service') },
          }
        },
      }
    } }
  } }))
  const scope = effectScope()
  const panels = scope.run(() => useSheetChartPanels({ studio: () => studio.value, unitId: () => unitId.value, sheetId: () => sheetId.value,
    sheets: () => [{ id: sheetId.value, name: sheetId.value, cells: {} }], active: () => active.value, canRestoreFocus: () => canFocus.value,
    refresh: (owner, unit) => refreshes.push({ owner, unit }), focus: () => { focusCount++ }, notify: message => notices.push(message) }))
  return { panels, firstStudio, secondStudio, studio, unitId, sheetId, active, canFocus, calls, sourceCalls, statisticsCalls, refreshes, notices,
    focusCount: () => focusCount, dataReads: () => dataReads, dispose() { scope.stop(); spy.mockRestore() } }
}

test('a pending insert retains its unit while late rejection cannot touch the next unit session', async () => {
  const first = deferred(), second = deferred(), f = fixture({ insert: ({ unit }) => unit === 'unit-a' ? first.promise : second.promise })
  try {
    f.panels.openInsertChart(config())
    const oldSession = f.panels.chartInsertContext.value, oldApply = oldSession.apply(config())
    assert.equal(f.panels.chartInsertBusy.value, true)
    f.unitId.value = 'unit-b'
    assert.equal(f.panels.chartInsertContext.value, undefined)
    assert.equal(f.panels.chartInsertBusy.value, false)
    f.panels.openInsertChart(config())
    const nextSession = f.panels.chartInsertContext.value, nextApply = nextSession.apply(config())
    assert.notEqual(nextSession.key, oldSession.key)
    first.reject(new Error('old-unit failure'))
    await oldApply
    oldSession.setOpen(false); oldSession.closed()
    assert.equal(f.panels.chartInsertContext.value, nextSession)
    assert.equal(f.panels.chartInsertBusy.value, true)
    assert.equal(f.panels.chartModalOpen.value, true)
    assert.equal(f.panels.chartError.value, '')
    assert.deepEqual(f.refreshes, [])
    assert.equal(f.focusCount(), 0)
    assert.deepEqual(f.calls.map(call => call.unit), ['unit-a', 'unit-b'])
    second.resolve()
    await nextApply
    assert.deepEqual(f.refreshes, [{ owner: f.firstStudio, unit: 'unit-b' }])
    nextSession.closed()
    assert.equal(f.focusCount(), 1)
  } finally { f.dispose() }
})

test('a pending editor update is isolated after switching sheets', async () => {
  const first = deferred(), second = deferred(), f = fixture({ update: ({ sheet }) => sheet === 'sheet-a' ? first.promise : second.promise })
  try {
    f.panels.openChartEditor(model())
    const oldSession = f.panels.chartEditorContext.value, oldApply = oldSession.apply({ title: 'First' })
    f.sheetId.value = 'sheet-b'
    f.panels.openChartEditor(model('sheet-b', 'chart-b'))
    const nextSession = f.panels.chartEditorContext.value, nextApply = nextSession.apply({ title: 'Second' })
    first.resolve()
    await oldApply
    assert.equal(f.panels.chartEditorContext.value, nextSession)
    assert.equal(f.panels.chartEditorBusy.value, true)
    assert.deepEqual(f.refreshes, [])
    assert.deepEqual(f.calls.map(call => [call.sheet, call.id]), [['sheet-a', 'chart-a'], ['sheet-b', 'chart-b']])
    second.resolve()
    await nextApply
    assert.equal(f.panels.chartEditorOpen.value, true)
    assert.deepEqual(f.refreshes, [{ owner: f.firstStudio, unit: 'unit-a' }])
  } finally { f.dispose() }
})

test('pivot callbacks cannot reopen, mark busy or refresh after deactivation', () => {
  const f = fixture()
  try {
    f.panels.openPivotChart()
    const session = f.panels.pivotChartContext.value
    session.setBusy(true)
    f.active.value = false
    assert.equal(f.panels.pivotChartContext.value, undefined)
    assert.equal(f.panels.pivotChartPanelOpen.value, false)
    session.setOpen(true); session.setBusy(true); session.applied(); session.closed()
    assert.equal(f.panels.pivotChartPanelOpen.value, false)
    assert.deepEqual(f.refreshes, [])
    f.active.value = true
    f.panels.openInsertChart(config())
    assert.equal(f.panels.chartModalOpen.value, true)
    assert.equal(f.focusCount(), 0)
  } finally { f.dispose() }
})

test('workbook disposal releases the current session and studio/scope changes dispose event listeners', async () => {
  const f = fixture()
  try {
    f.panels.openChartEditor(model())
    f.panels.chartEditorChart.value
    const before = f.panels.chartSourceRevision.value
    f.firstStudio.hub.emit('SheetValueChanged', { workbookId: 'another-unit' })
    assert.equal(f.panels.chartSourceRevision.value, before)
    f.firstStudio.hub.emit('SheetValueChanged', { workbookId: 'unit-a' })
    assert.equal(f.panels.chartSourceRevision.value, before + 1)
    assert.equal(f.dataReads(), 0)
    f.firstStudio.hub.emit('WorkbookDisposed', { workbookId: 'unit-a' })
    assert.equal(f.panels.chartEditorContext.value, undefined)
    assert.equal(f.panels.chartEditorOpen.value, false)
    assert.equal(f.focusCount(), 0)
    f.studio.value = f.secondStudio
    await nextTick()
    assert.equal(f.firstStudio.subscriptions.size, 0)
    assert.equal(f.secondStudio.subscriptions.size, 8)
  } finally { f.dispose() }
  assert.equal(f.secondStudio.subscriptions.size, 0)
  f.panels.openInsertChart(config())
  assert.equal(f.panels.chartModalOpen.value, false)
})

test('current-session failures remain reviewable and closing restores focus only when eligible', async () => {
  const f = fixture({ insert: () => Promise.reject(new Error('源范围无效')) })
  try {
    f.panels.openInsertChart(config())
    const session = f.panels.chartInsertContext.value
    await session.apply(config())
    assert.equal(f.panels.chartError.value, '源范围无效')
    assert.equal(f.panels.chartInsertBusy.value, false)
    assert.equal(f.panels.chartModalOpen.value, true)
    session.setOpen(false); session.closed(); session.closed()
    assert.equal(f.panels.chartError.value, '')
    assert.equal(f.focusCount(), 1)
    f.canFocus.value = false
    f.panels.openPivotChart()
    const pivot = f.panels.pivotChartContext.value
    pivot.setOpen(false); pivot.closed()
    assert.equal(f.focusCount(), 1)
  } finally { f.dispose() }
})

test('source loader failures retain the detached chart model and propagate to the shared panel source hook', async () => {
  const f = fixture({ readError: new Error('源工作表不可查看') })
  try {
    f.panels.openChartEditor(model())
    assert.equal(f.panels.chartEditorChart.value.id, 'chart-a')
    await assert.rejects(f.panels.chartEditorContext.value.loadSource(config(), { signal: new AbortController().signal }), /源工作表不可查看/)
    assert.equal(f.panels.chartEditorChart.value.id, 'chart-a')
    assert.equal(f.panels.chartEditorError.value, '')
    assert.equal(f.dataReads(), 0)
  } finally { f.dispose() }
})

test('editor and insert drafts read their current candidate binding through the captured worksheet and signal', async () => {
  const f = fixture()
  try {
    const controller = new AbortController(), candidate = { ...config(), valueFields: [0] }
    f.panels.openChartEditor(model())
    await f.panels.chartEditorContext.value.loadSource(candidate, { signal: controller.signal })
    f.panels.openInsertChart(config())
    await f.panels.chartInsertContext.value.loadSource(candidate, { signal: controller.signal })
    assert.deepEqual(f.sourceCalls.map(call => [call.unit, call.sheet, call.value.valueFields]), [['unit-a', 'sheet-a', [0]], ['unit-a', 'sheet-a', [0]]])
    assert.ok(f.sourceCalls.every(call => call.signal === controller.signal))
    assert.equal(f.dataReads(), 0); assert.deepEqual(f.calls, [])
  } finally { f.dispose() }
})

test('closed or replaced panel sessions cannot start reads in a newly active workbook', async () => {
  const f = fixture()
  try {
    f.panels.openChartEditor(model())
    const old = f.panels.chartEditorContext.value
    f.unitId.value = 'unit-b'; f.panels.openChartEditor(model())
    await assert.rejects(old.loadSource(config(), { signal: new AbortController().signal }), error => error.code === 'COMPUTE_STALE_RESULT')
    assert.deepEqual(f.sourceCalls, [])
    const current = f.panels.chartEditorContext.value
    current.setOpen(false)
    await assert.rejects(current.loadSource(config(), { signal: new AbortController().signal }), error => error.code === 'COMPUTE_STALE_RESULT')
    assert.deepEqual(f.sourceCalls, [])
  } finally { f.dispose() }
})

test('late source completion after switching worksheets is rejected without modifying another panel session', async () => {
  const pending = deferred(), f = fixture({ source: () => pending.promise })
  try {
    f.panels.openChartEditor(model())
    const old = f.panels.chartEditorContext.value
    const read = old.loadSource(config(), { signal: new AbortController().signal })
    f.sheetId.value = 'sheet-b'; f.panels.openChartEditor(model('sheet-b', 'chart-b'))
    const current = f.panels.chartEditorContext.value
    const rejected = assert.rejects(read, error => error.code === 'COMPUTE_STALE_RESULT')
    pending.resolve({ categories: ['Old'], series: [{ field: 1, name: 'Old', values: [1] }] })
    await rejected
    assert.equal(f.panels.chartEditorContext.value, current)
    assert.equal(f.panels.chartEditorError.value, ''); assert.equal(f.panels.chartEditorBusy.value, false)
    assert.deepEqual(f.sourceCalls.map(call => [call.unit, call.sheet]), [['unit-a', 'sheet-a']])
    assert.deepEqual(f.calls, []); assert.deepEqual(f.refreshes, [])
  } finally { f.dispose() }
})

test('aborted panel source reads are rejected before lookup or after a provider ignores cancellation', async () => {
  const pending = deferred(), f = fixture({ source: () => pending.promise })
  try {
    f.panels.openInsertChart(config())
    const session = f.panels.chartInsertContext.value, alreadyCancelled = new AbortController()
    alreadyCancelled.abort()
    await assert.rejects(session.loadSource(config(), { signal: alreadyCancelled.signal }), error => error.code === 'COMPUTE_CANCELLED')
    assert.deepEqual(f.sourceCalls, [])
    const controller = new AbortController(), read = session.loadSource(config(), { signal: controller.signal })
    const rejected = assert.rejects(read, error => error.code === 'COMPUTE_CANCELLED')
    controller.abort(); pending.resolve({ categories: [], series: [] }); await rejected
    assert.equal(f.sourceCalls[0].signal.aborted, true)
    assert.equal(f.panels.chartModalOpen.value, true); assert.equal(f.panels.chartError.value, '')
    assert.deepEqual(f.calls, []); assert.equal(f.dataReads(), 0)
  } finally { f.dispose() }
})

test('Pareto editor and insert statistical readers retain their captured host, candidate binding and AbortSignal', async () => {
  const f = fixture()
  try {
    const controller = new AbortController(), candidate = { ...config(), type: 'pareto', options: { pareto: { includeZeroValues: true } } }
    f.panels.openChartEditor({ ...model(), type: 'pareto' })
    const editorStatistics = await f.panels.chartEditorContext.value.loadStatistics(candidate, { signal: controller.signal })
    f.panels.openInsertChart(candidate)
    const insertStatistics = await f.panels.chartInsertContext.value.loadStatistics(candidate, { signal: controller.signal })
    assert.deepEqual(editorStatistics.pareto.categories, ['A']); assert.deepEqual(insertStatistics.pareto.categories, ['A'])
    assert.deepEqual(f.statisticsCalls.map(call => [call.unit, call.sheet, call.value.options.pareto.includeZeroValues]),
      [['unit-a', 'sheet-a', true], ['unit-a', 'sheet-a', true]])
    assert.ok(f.statisticsCalls.every(call => call.signal === controller.signal))
    assert.deepEqual(f.sourceCalls, []); assert.deepEqual(f.calls, []); assert.deepEqual(f.refreshes, []); assert.equal(f.dataReads(), 0)
  } finally { f.dispose() }
})

test('a held statistical result from a closed editor cannot enter a replacement session with the same host IDs', async () => {
  const pending = deferred(), f = fixture({ statistics: () => pending.promise })
  try {
    f.panels.openChartEditor({ ...model(), type: 'pareto' })
    const old = f.panels.chartEditorContext.value, read = old.loadStatistics({ ...config(), type: 'pareto' }, { signal: new AbortController().signal })
    const rejected = assert.rejects(read, error => error.code === 'COMPUTE_STALE_RESULT')
    old.setOpen(false); old.closed(); f.panels.openChartEditor({ ...model(), type: 'pareto' })
    const current = f.panels.chartEditorContext.value
    pending.resolve({ preparedData: { categories: ['Old'], series: [] }, pareto: { categories: ['Old'], values: [1], cumulativePercent: [100] } })
    await rejected
    assert.notEqual(current.key, old.key); assert.equal(f.panels.chartEditorContext.value, current)
    assert.equal(f.panels.chartEditorError.value, ''); assert.equal(f.panels.chartEditorBusy.value, false)
    await assert.rejects(old.loadStatistics({ ...config(), type: 'pareto' }, { signal: new AbortController().signal }), error => error.code === 'COMPUTE_STALE_RESULT')
    assert.equal(f.statisticsCalls.length, 1); assert.deepEqual(f.sourceCalls, []); assert.deepEqual(f.calls, [])
  } finally { f.dispose() }
})

test('statistical readers check cancellation before lookup and after a provider returns despite abort', async () => {
  const pending = deferred(), f = fixture({ statistics: () => pending.promise })
  try {
    f.panels.openInsertChart({ ...config(), type: 'pareto' })
    const session = f.panels.chartInsertContext.value, cancelled = new AbortController()
    cancelled.abort()
    await assert.rejects(session.loadStatistics({ ...config(), type: 'pareto' }, { signal: cancelled.signal }), error => error.code === 'COMPUTE_CANCELLED')
    assert.deepEqual(f.statisticsCalls, [])
    const controller = new AbortController(), read = session.loadStatistics({ ...config(), type: 'pareto' }, { signal: controller.signal })
    const rejected = assert.rejects(read, error => error.code === 'COMPUTE_CANCELLED')
    controller.abort(); pending.resolve({ preparedData: { categories: [], series: [] }, pareto: { categories: [], values: [], cumulativePercent: [] } })
    await rejected; assert.equal(f.statisticsCalls[0].signal.aborted, true)
    assert.equal(f.panels.chartModalOpen.value, true); assert.equal(f.panels.chartError.value, '')
    assert.deepEqual(f.sourceCalls, []); assert.deepEqual(f.calls, []); assert.equal(f.dataReads(), 0)
  } finally { f.dispose() }
})

test('P17 same-chart navigation retains one editor session and does not issue author or source reads', () => {
  const f = fixture()
  try {
    assert.equal(f.panels.openChartEditor(model(), { tab: 'style', titleKey: 'subtitle' }), true)
    const session = f.panels.chartEditorContext.value, routeKey = f.panels.chartEditorRouteKey.value
    assert.equal(f.panels.openChartEditor(model(), { tab: 'data' }), true)
    assert.equal(f.panels.chartEditorContext.value, session); assert.equal(f.panels.chartEditorContext.value.key, session.key)
    assert.equal(f.panels.chartEditorRouteKey.value, routeKey + 1); assert.deepEqual(f.panels.chartEditorRoute.value, { tab: 'data' })
    assert.equal(f.panels.chartEditorOpen.value, true); assert.deepEqual(f.calls, []); assert.deepEqual(f.sourceCalls, []); assert.deepEqual(f.refreshes, [])
  } finally { f.dispose() }
})

test('P17 invalid routes and a busy editor keep the previous session and requested location intact', async () => {
  const pending = deferred(), f = fixture({ update: () => pending.promise })
  try {
    f.panels.openChartEditor(model(), { tab: 'style', section: 'legend' })
    const session = f.panels.chartEditorContext.value, route = f.panels.chartEditorRoute.value, key = f.panels.chartEditorRouteKey.value
    assert.throws(() => f.panels.openChartEditor(model(), { tab: 'data', section: 'legend' }), /数据标签页/)
    assert.equal(f.panels.chartEditorContext.value, session); assert.equal(f.panels.chartEditorRoute.value, route); assert.equal(f.panels.chartEditorRouteKey.value, key)
    const applying = session.apply({ title: 'Pending' })
    assert.equal(f.panels.openChartEditor(model(), { tab: 'data' }), false)
    assert.equal(f.panels.chartEditorRoute.value, route); pending.resolve(); await applying
    assert.equal(f.panels.chartEditorContext.value, session); assert.equal(f.panels.chartEditorRouteKey.value, key)
  } finally { f.dispose() }
})

test('P17 navigationCurrent rejects IME/focus exclusions, replaced identities and closed sessions', () => {
  const f = fixture()
  try {
    f.panels.openChartEditor(model(), { tab: 'style', axisId: 'yAxis' })
    const old = f.panels.chartEditorContext.value
    assert.equal(old.navigationCurrent(), true); f.canFocus.value = false; assert.equal(old.navigationCurrent(), false)
    f.canFocus.value = true; assert.equal(old.navigationCurrent(), true)
    old.setOpen(false); assert.equal(old.navigationCurrent(), false); old.closed()
    f.panels.openChartEditor(model(), { tab: 'data' })
    const next = f.panels.chartEditorContext.value
    assert.equal(old.navigationCurrent(), false); assert.equal(next.navigationCurrent(), true)
    f.studio.value = f.secondStudio; assert.equal(next.navigationCurrent(), false)
    assert.equal(f.panels.chartEditorRoute.value, undefined); assert.equal(f.panels.chartEditorOpen.value, false)
    assert.equal(f.focusCount(), 1)
  } finally { f.dispose() }
})
