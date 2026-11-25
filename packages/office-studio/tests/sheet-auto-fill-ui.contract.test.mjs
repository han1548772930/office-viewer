import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { nextTick, shallowRef } from 'vue'
import { createServer } from 'vite'

// Draft only. This exercises the command/view/selection boundary without a DOM snapshot test.
let server, sdk, controller
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[sdk, controller] = await Promise.all(['/packages/office-studio/src/index.ts',
    '/packages/office-studio/src/components/spreadsheet/use-sheet-auto-fill.ts'].map(path => server.ssrLoadModule(path)))
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
const range = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })
const gridRange = value => ({ startRow: value.startRow, endRow: value.endRow, startColumn: value.startCol, endColumn: value.endCol })
function host(t, options = {}) {
  const model = sdk.createBlankWorkbook(); model.id = 'fill-ui-book'; model.calculationMode = 'manual'
  model.sheets[0].id = 'data'; model.activeSheetId = 'data'; model.sheets[0].rowCount = 30; model.sheets[0].columnCount = 12
  model.sheets[0].cells['0:0'] = sdk.createCell(0, 0, 1); model.sheets[0].cells['1:0'] = sdk.createCell(1, 0, 2)
  const studio = options.standalone ? undefined : new sdk.OfficeStudio({ workbook: model, sheetTextMeasure: text => [...text].length * 8 })
  if (studio) t.onTestFinished(() => studio.dispose())
  const owner = shallowRef(studio), view = shallowRef(sdk.toExcelWorkbookView(studio?.getWorkbookModel(model.id) ?? model))
  if (options.standalone) delete view.value.id
  const selected = shallowRef(range(0, 0, 1)), active = shallowRef(true), editing = shallowRef(false), calls = { applied: 0, changed: 0, focus: 0, notices: [] }
  if (studio) studio.setSelection(model.id, 'data', gridRange(selected.value))
  const context = {
    studio: () => owner.value, workbook: () => view.value, sheetId: () => view.value.activeSheetId,
    range: () => selected.value, userId: () => 'local-user', region: () => 'en-US', active: () => active.value, editing: () => editing.value,
    finishEdit: () => !editing.value, prepareView: next => options.prepare ? options.prepare(next) : next,
    applyView: next => {
      view.value = copy(next); calls.applied++
      const state = owner.value?.getSelectionState(model.id, 'data'), actual = state?.ranges[state.activeRangeIndex] ?? state?.ranges.at(-1)
      if (actual) selected.value = range(actual.startRow, actual.startColumn, actual.endRow, actual.endColumn)
      options.onApply?.(context)
    },
    selection: value => { selected.value = copy(value) },
    changed: request => { if (request.current()) { calls.changed++; options.onChanged?.(context) } },
    focus: () => calls.focus++, notify: value => calls.notices.push(value),
  }
  const ui = controller.useSheetAutoFill(context)
  return { studio, owner, view, selected, active, editing, calls, context, ui,
    current: () => studio?.getWorkbookModel(model.id) ?? sdk.applyExcelWorkbookView(model, view.value, 'en-US') }
}
function key(key, extra = {}) { return { key, ctrlKey: true, metaKey: false, altKey: false, shiftKey: false, isComposing: false,
  keyCode: 0, defaultPrevented: false, preventDefault() { this.defaultPrevented = true }, ...extra } }

test('pointer fill publishes the complete model once and shows options only while its exact selection/history remain current', async t => {
  const h = host(t)
  assert.equal(h.ui.fillSelection(gridRange(range(0, 0, 1)), { row: 5, column: 0 }), true)
  await nextTick()
  assert.deepEqual([2, 3, 4, 5].map(row => h.current().sheets[0].cells[`${row}:0`].value), [3, 4, 5, 6])
  assert.equal(h.calls.applied, 1); assert.equal(h.calls.changed, 1); assert.equal(h.calls.focus, 1)
  const option = h.ui.fillOptions.value; assert.equal(option.applyType, 'SERIES'); assert.equal(option.supportsSeries, true)
  assert.equal(h.ui.refill({ sessionId: option.id, applyType: 'COPY' }), true)
  assert.deepEqual([2, 3, 4, 5].map(row => h.current().sheets[0].cells[`${row}:0`].value), [1, 2, 1, 2])
  assert.equal(h.studio.editor.history.units['fill-ui-book'].undo.length, 1)
  h.selected.value = range(10, 3); assert.equal(h.ui.fillOptions.value, undefined)
  h.selected.value = range(0, 0, 5); assert.equal(h.ui.fillOptions.value.id, option.id)
  sdk.FOffice.newAPI(h.studio).getActiveWorkbook().getActiveSheet().getRange('D8').setValue('later')
  assert.equal(h.ui.fillOptions.value, undefined)
})

test('keyboard copy uses canonical commands while editing, composing and other modifiers retain native behavior', t => {
  const h = host(t), down = key('d')
  assert.equal(h.ui.handleFillKeydown(down), true); assert.equal(down.defaultPrevented, true)
  assert.equal(h.current().sheets[0].cells['1:0'].value, 1); assert.equal(h.ui.fillOptions.value, undefined)
  for (const input of [key('r', { isComposing: true }), key('d', { keyCode: 229 }), key('d', { altKey: true }), key('d', { shiftKey: true })]) {
    assert.equal(h.ui.handleFillKeydown(input), false); assert.equal(input.defaultPrevented, false)
  }
  h.editing.value = true; assert.equal(h.ui.handleFillKeydown(key('d')), false)
})

test('an equivalent refill updates the options and focus without publishing another model or content history', async t => {
  const h = host(t)
  assert.equal(h.ui.fillSelection(gridRange(range(0, 0, 1)), { row: 5, column: 0 }), true)
  await nextTick()
  const session = h.ui.fillOptions.value, revision = h.studio.getWorkbookRevision('fill-ui-book')
  assert.equal(h.ui.refill({ sessionId: session.id, applyType: 'NO_FORMAT' }), true)
  await nextTick()
  assert.equal(h.calls.applied, 1); assert.equal(h.calls.changed, 1); assert.equal(h.calls.focus, 2)
  assert.equal(h.ui.fillOptions.value.applyType, 'NO_FORMAT'); assert.equal(h.studio.getWorkbookRevision('fill-ui-book'), revision)
  assert.equal(h.studio.editor.history.units['fill-ui-book'].undo.length, 1)
})

test('controlled standalone fill uses the shared model planner and one existing change callback without a private undo stack', async t => {
  const h = host(t, { standalone: true })
  assert.equal(h.ui.fillSelection(gridRange(range(0, 0, 1)), { row: 4, column: 0 }), true)
  await nextTick()
  assert.deepEqual([2, 3, 4].map(row => h.current().sheets[0].cells[`${row}:0`].value), [3, 4, 5])
  assert.deepEqual(h.selected.value, range(0, 0, 4)); assert.equal(h.calls.changed, 1)
  assert.equal(h.view.value.id, undefined); assert.equal(h.ui.fillOptions.value, undefined)
})

test('a failed standalone preparation never publishes and a changed host suppresses later callbacks/focus', async t => {
  const failed = host(t, { standalone: true, prepare: () => { throw new Error('prepare failed') } }), before = copy(failed.view.value)
  assert.equal(failed.ui.fillSelection(gridRange(range(0, 0, 1)), { row: 4, column: 0 }), false)
  assert.deepEqual(failed.view.value, before); assert.equal(failed.calls.changed, 0); assert.equal(failed.calls.applied, 0)
  const switched = host(t, { onApply: context => { context.workbook().activeSheetId = 'another' } })
  assert.equal(switched.ui.fillSelection(gridRange(range(0, 0, 1)), { row: 4, column: 0 }), true)
  await nextTick(); assert.equal(switched.calls.changed, 0); assert.equal(switched.calls.focus, 0)
})
