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
  const calls = [], refreshes = [], notices = []
  let focusCount = 0, dataReads = 0
  const spy = vi.spyOn(FOffice, 'newAPI').mockImplementation(owner => ({ getWorkbook(unit) {
    return { getSheetBySheetId(sheet) {
      return {
        insertChart(value) { calls.push({ kind: 'insert', owner, unit, sheet, value }); return handlers.insert?.({ owner, unit, sheet, value }) ?? Promise.resolve() },
        getChartById(id) {
          return {
            update(patch) { calls.push({ kind: 'update', owner, unit, sheet, id, patch }); return handlers.update?.({ owner, unit, sheet, id, patch }) ?? Promise.resolve() },
            getInfo() { return model(sheet, id) },
            getData() { dataReads++; if (handlers.readError) throw handlers.readError; return { fields: ['Category', 'Value'], records: [['A', 10]], categories: ['A'], series: [{ field: 1, name: 'Value', values: [10] }] } },
          }
        },
      }
    } }
  } }))
  const scope = effectScope()
  const panels = scope.run(() => useSheetChartPanels({ studio: () => studio.value, unitId: () => unitId.value, sheetId: () => sheetId.value,
    sheets: () => [{ id: sheetId.value, name: sheetId.value, cells: {} }], active: () => active.value, canRestoreFocus: () => canFocus.value,
    refresh: (owner, unit) => refreshes.push({ owner, unit }), focus: () => { focusCount++ }, notify: message => notices.push(message) }))
  return { panels, firstStudio, secondStudio, studio, unitId, sheetId, active, canFocus, calls, refreshes, notices,
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
    f.panels.chartEditorData.value
    const before = f.dataReads()
    f.firstStudio.hub.emit('SheetValueChanged', { workbookId: 'another-unit' })
    f.panels.chartEditorData.value
    assert.equal(f.dataReads(), before)
    f.firstStudio.hub.emit('SheetValueChanged', { workbookId: 'unit-a' })
    f.panels.chartEditorData.value
    assert.equal(f.dataReads(), before + 1)
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

test('preview source-read failures retain the chart model and become panel errors', () => {
  const f = fixture({ readError: new Error('源工作表不可查看') })
  try {
    f.panels.openChartEditor(model())
    assert.equal(f.panels.chartEditorChart.value.id, 'chart-a')
    assert.equal(f.panels.chartEditorData.value, undefined)
    assert.equal(f.panels.chartEditorError.value, '源工作表不可查看')
  } finally { f.dispose() }
})
