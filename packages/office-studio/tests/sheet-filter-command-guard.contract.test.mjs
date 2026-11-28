import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

let server, api, panel, guards
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  panel = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-sort-filter/sheet-data-panel-state.ts')
  guards = await server.ssrLoadModule('/packages/office-studio/src/api/workbook-command-guards.ts')
})
afterAll(async () => { await server?.close() })

function setup() {
  const studio = new api.OfficeStudio(), workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:B3').setValues([['Region', 'Value'], ['East', 2], ['West', 1]])
  const request = { unitId: workbook.getId(), sheetId: sheet.getSheetId(), expectedRevision: studio.getWorkbookRevision(workbook.getId()),
    range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, columns: { 0: { kind: 'values', selectedItems: ['East'] } } }
  return { studio, workbook, sheet, request }
}

test('a panel invalidated by a before-command observer cannot configure a filter or consume redo', () => {
  const h = setup()
  try {
    const source = JSON.stringify(h.workbook.save()), history = JSON.stringify(h.studio.editor.history.units[h.workbook.getId()])
    let current = true
    guards.bindWorkbookCommandGuard(h.request, () => { if (!current) throw new Error('panel owner changed') })
    const listener = h.studio.events.addEvent('BeforeCommandExecute', event => {
      if (event.id === panel.CONFIGURE_SHEET_FILTER_COMMAND) current = false
    })
    try { assert.throws(() => h.studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, h.request), /panel owner changed/) }
    finally { listener.dispose() }
    assert.equal(JSON.stringify(h.workbook.save()), source)
    assert.equal(JSON.stringify(h.studio.editor.history.units[h.workbook.getId()]), history)
    assert.equal(h.sheet.getAutoFilter(), null)
  } finally { h.studio.dispose() }
})

test('a nested content edit from a command observer survives while the stale filter is rejected', () => {
  const h = setup()
  try {
    delete h.request.expectedRevision
    const listener = h.studio.events.addEvent('BeforeCommandExecute', event => {
      if (event.id === panel.CONFIGURE_SHEET_FILTER_COMMAND) h.sheet.getRange('D1').setValue('observer edit')
    })
    try { assert.throws(() => h.studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, h.request), /更改/) }
    finally { listener.dispose() }
    assert.equal(h.sheet.getRange('D1').getValue(), 'observer edit'); assert.equal(h.sheet.getAutoFilter(), null)
    h.workbook.undo(); assert.equal(h.sheet.getRange('D1').getValue(), null)
  } finally { h.studio.dispose() }
})

test('guarded filter commands retain a single content transaction and snapshot save/reopen behavior', () => {
  const h = setup()
  try {
    const history = h.studio.editor.history.units[h.workbook.getId()].undo.length
    guards.bindWorkbookCommandGuard(h.request, () => {})
    h.studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, h.request)
    assert.deepEqual(h.sheet.getFilteredRows(), [2])
    assert.equal(h.studio.editor.history.units[h.workbook.getId()].undo.length, history + 1)
    h.workbook.undo(); assert.equal(h.sheet.getAutoFilter(), null)
    h.workbook.redo(); assert.deepEqual(h.sheet.getFilteredRows(), [2])
    const reopened = new api.OfficeStudio({ workbook: JSON.parse(JSON.stringify(h.workbook.save())) })
    try { assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getFilteredRows(), [2]) }
    finally { reopened.dispose() }
  } finally { h.studio.dispose() }
})

test('the first guarded edit of an inactive workbook does not mistake history creation for a reopen', () => {
  const h = setup()
  try {
    const active = h.studio.activeWorkbookId
    const model = h.studio.createWorkbook({ id: 'inactive-grid', sheetOrder: ['data'], sheets: { data: { name: 'Data',
      cellData: { 0: { 0: { v: 'Region' } }, 1: { 0: { v: 'East' } }, 2: { 0: { v: 'West' } } } } } }, false)
    h.studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { unitId: model.id, sheetId: model.sheets[0].id,
      range: { startRow: 0, startCol: 0, endRow: 2, endCol: 0 }, columns: { 0: { kind: 'values', selectedItems: ['East'] } }, expectedRevision: 0 })
    assert.deepEqual(h.studio.getWorkbookModel(model.id).sheets[0].autoFilter.filteredRows, [2])
    assert.equal(h.studio.activeWorkbookId, active)
    assert.equal(h.studio.getWorkbookRevision(model.id), 1)
  } finally { h.studio.dispose() }
})
