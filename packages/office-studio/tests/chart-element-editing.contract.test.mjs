import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { effectScope, nextTick, ref, shallowRef } from 'vue'

// P19 host/canonical-command drafts. No Vite server or test was run this batch.
let server, api, elements, edits, editing, commands
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, elements, edits, editing, commands] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-element-interaction.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-settings-draft.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/views/use-chart-element-editing.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-object-commands.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const undoCount = (studio, id) => studio.editor.history.units[id]?.undo.length ?? 0
const deferred = () => { let resolve; const promise = new Promise(done => { resolve = done }); return { promise, resolve } }
async function fixture() {
  const studio = new api.OfficeStudio(), facade = api.FOffice.newAPI(studio), workbook = facade.getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C4').setValues([['Label', 'Original', 'Comparison'], ['A', 2, 6], ['B', 4, 5], ['C', 6, 3]])
  const chart = await sheet.insertChart(sheet.newChart('line').setSource('A1:C4').setValueFields([1, 2]).setTitle('Chart')
    .setXAxis({ title: 'Categories' }).setYAxis({ title: 'Values' }).build())
  await chart.getComputedData()
  const slot = shallowRef(studio), active = ref(true), editable = ref(true), blocked = ref(false), subUnit = ref(sheet.getSheetId()), unit = ref(workbook.getId())
  const id = chart.getChartId(), opened = [], focused = [], events = [], scope = effectScope()
  const current = () => slot.value.getWorkbookModel(unit.value)?.charts?.find(chart => chart.id === id && chart.sheetId === subUnit.value)
  const capture = () => { const value = current(); return value && { unitId: unit.value, sheetId: subUnit.value, chartId: id,
    expectedRevision: slot.value.getWorkbookRevision(unit.value), snapshot: JSON.stringify(value), hostSessionId: 'element-host' } }
  const registration = commands.registerChartObjectUIHost(studio, { sessionId: 'element-host', target: capture,
    current: target => { const value = capture(); return !!value && target.unitId === value.unitId && target.sheetId === value.sheetId
      && target.chartId === value.chartId && target.expectedRevision === value.expectedRevision && target.snapshot === value.snapshot },
    currentSheet: (unitId, sheetId) => active.value && unitId === unit.value && sheetId === subUnit.value,
    openPanel: (target, route) => { opened.push({ target, route }); return true }, exportImage: async () => true, writeClipboard: async () => true })
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  const ui = scope.run(() => editing.useChartElementEditing({ studio: () => slot.value, workbookId: () => unit.value, sheetId: () => subUnit.value,
    active: () => active.value, chart: chartId => chartId === id ? current() : undefined, canEdit: () => editable.value, busy: () => blocked.value,
    target: (chartId, element) => chartId === id && current() ? { element, key: elements.chartCanvasElementKey(element), bounds: { x: 10, y: 20, width: 30, height: 40 } } : undefined,
    changed() {}, focus: chartId => focused.push(chartId), openPanel: async (_chartId, route) => studio.executeCommand(commands.CHART_OBJECT_COMMANDS.edit, { ...capture(), route }),
  }))
  const binding = () => elements.chartCanvasSeriesBindings({ series: [{ field: 1, name: 'Displayed', sourceName: 'Original', values: [2, 4, 6] }] })[0]
  const select = element => ui.select(id, { element, key: elements.chartCanvasElementKey(element), bounds: { x: 10, y: 20, width: 30, height: 40 } })
  return { studio, slot, active, editable, blocked, subUnit, unit, workbook, sheet, chart, id, current, ui, binding, select, opened, focused, events,
    dispose() { scope.stop(); registration.dispose(); studio.dispose() } }
}

test('one scoped FChart update owns history/save/reopen, while a no-op draft adds no author command', async () => {
  const f = await fixture(); let reopened
  try {
    const count = undoCount(f.studio, f.unit.value), element = { kind: 'axisTitle', axisId: 'xAxis' }, selected = f.select(element)
    const draft = edits.cloneChartDraftConfig(f.current()); draft.xAxis.title = 'Authored'; draft.legend.position = 'right'
    await f.ui.apply(selected, draft)
    assert.equal(f.chart.getInfo().xAxis.title, 'Authored'); assert.equal(f.chart.getInfo().legend.position, 'bottom')
    assert.equal(undoCount(f.studio, f.unit.value), count + 1); assert.equal(f.events.filter(id => id === 'sheet.command.update-chart').length, 1)
    await f.ui.apply(f.ui.selected.value, edits.cloneChartDraftConfig(f.current()))
    assert.equal(undoCount(f.studio, f.unit.value), count + 1)
    const authored = f.chart.getInfo(); f.workbook.undo(); await nextTick(); assert.equal(f.chart.getInfo().xAxis.title, 'Categories')
    assert.equal(f.ui.selected.value, null); f.workbook.redo(); assert.deepEqual(f.chart.getInfo(), authored)
    const saved = json(f.workbook.save()); assert.equal(saved.charts[0].element, undefined); assert.equal(saved.charts[0].binding, undefined)
    reopened = new api.OfficeStudio({ workbook: saved })
    assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(f.id).getInfo(), authored)
  } finally { reopened?.dispose(); f.dispose() }
})

test('an external model revision retires old menu callbacks before they can create another author command', async () => {
  const f = await fixture()
  try {
    const selected = f.select({ kind: 'axisTitle', axisId: 'xAxis' }), draft = edits.cloneChartDraftConfig(f.current()); draft.xAxis.title = 'Stale draft'
    await f.chart.setTitle('External title'); await nextTick()
    const before = json(f.workbook.save()), count = undoCount(f.studio, f.unit.value)
    assert.equal(f.ui.selected.value, null)
    await assert.rejects(() => f.ui.apply(selected, draft), /重新选择/)
    assert.equal(await f.ui.more(selected), false); assert.deepEqual(json(f.workbook.save()), before); assert.equal(undoCount(f.studio, f.unit.value), count)
  } finally { f.dispose() }
})

test('reusing workbook/sheet/chart IDs in a new owner does not revive the previous author session', async () => {
  const f = await fixture(); let replacement
  try {
    const selected = f.select({ kind: 'legend' }), draft = edits.cloneChartDraftConfig(f.current()); draft.legend.position = 'left'
    replacement = new api.OfficeStudio({ workbook: json(f.workbook.save()) }); f.slot.value = replacement; await nextTick()
    const before = json(api.FOffice.newAPI(replacement).getActiveWorkbook().save())
    assert.equal(f.ui.selected.value, null); await assert.rejects(() => f.ui.apply(selected, draft), /重新选择/)
    assert.deepEqual(json(api.FOffice.newAPI(replacement).getActiveWorkbook().save()), before)
  } finally { replacement?.dispose(); f.dispose() }
})

test('busy and read-only hosts disable author controls, and a permission change retires an editable selection', async () => {
  const f = await fixture()
  try {
    const selected = f.select({ kind: 'axis', axisId: 'yAxis', part: 'label' }), draft = edits.cloneChartDraftConfig(f.current()); draft.yAxis.reverse = true
    const before = json(f.workbook.save()), count = undoCount(f.studio, f.unit.value)
    f.blocked.value = true; assert.equal(f.ui.disabled.value, true); await assert.rejects(() => f.ui.apply(selected, draft), /重新选择/)
    f.blocked.value = false; assert.equal(f.ui.disabled.value, false)
    f.editable.value = false; await nextTick(); assert.equal(f.ui.selected.value, null)
    const readOnly = f.select({ kind: 'legend' }); assert.ok(readOnly); assert.equal(f.ui.disabled.value, true)
    await assert.rejects(() => f.ui.apply(readOnly, draft), /重新选择/)
    assert.deepEqual(json(f.workbook.save()), before); assert.equal(undoCount(f.studio, f.unit.value), count)
  } finally { f.dispose() }
})

test('Delete hides a selected title with one history entry, preserves point source data, and guards deferred focus restoration', async () => {
  const f = await fixture()
  try {
    const count = undoCount(f.studio, f.unit.value)
    f.select({ kind: 'axisTitle', axisId: 'yAxis' }); assert.equal(await f.ui.remove(), true); await nextTick()
    assert.equal(f.chart.getInfo().yAxis.title, undefined); assert.equal(f.ui.selected.value, null); assert.equal(undoCount(f.studio, f.unit.value), count + 1)
    const before = json(f.workbook.save()); f.select({ kind: 'dataPoint', binding: f.binding(), dataIndex: 1 })
    assert.equal(await f.ui.remove(), false); assert.deepEqual(json(f.workbook.save()), before)
    const focused = f.focused.length; f.ui.close(true); f.active.value = false; await nextTick(); assert.equal(f.focused.length, focused)
  } finally { f.dispose() }
})

test('promoting a real point selects its bound series and More uses the canonical focused-panel operation outside history', async () => {
  const f = await fixture()
  try {
    const selected = f.select({ kind: 'dataPoint', binding: f.binding(), dataIndex: 1 }), promoted = f.ui.promote(selected)
    assert.equal(promoted.key, 'mark:seriesStyle:field:1'); assert.equal(promoted.element.directEdit.binding.sourceName, 'Original')
    const before = json(f.workbook.save()), count = undoCount(f.studio, f.unit.value)
    assert.equal(await f.ui.more(promoted), true)
    assert.deepEqual(f.opened[0].route, { tab: 'style', section: 'series', seriesId: 'field:1' })
    assert.equal(f.ui.selected.value, null); assert.deepEqual(json(f.workbook.save()), before); assert.equal(undoCount(f.studio, f.unit.value), count)
  } finally { f.dispose() }
})

test('a late completion of a real committed command cannot reopen a retired menu or steal focus from an inactive host', async () => {
  const f = await fixture(), original = api.FChart.prototype.update, gate = deferred(), committed = deferred()
  try {
    api.FChart.prototype.update = async function(patch) { const result = await original.call(this, patch); committed.resolve(); await gate.promise; return result }
    const selected = f.select({ kind: 'axisTitle', axisId: 'xAxis' }), draft = edits.cloneChartDraftConfig(f.current()); draft.xAxis.title = 'Committed before leave'
    const count = undoCount(f.studio, f.unit.value), pending = f.ui.apply(selected, draft)
    await committed.promise; assert.equal(f.ui.busy.value, true); assert.equal(undoCount(f.studio, f.unit.value), count + 1)
    f.active.value = false; await nextTick(); assert.equal(f.ui.selected.value, null)
    gate.resolve(); await pending; await nextTick()
    assert.equal(f.ui.busy.value, false); assert.equal(f.ui.selected.value, null); assert.deepEqual(f.focused, [])
    assert.equal(f.chart.getInfo().xAxis.title, 'Committed before leave')
  } finally { gate.resolve(); api.FChart.prototype.update = original; f.dispose() }
})
