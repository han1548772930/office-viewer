import assert from 'node:assert/strict'
import { beforeAll, afterAll, test } from 'vitest'
import { createServer } from 'vite'

// Implementation-phase drafts. No runtime verification has been performed.
let server, api, panel
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  panel = await server.ssrLoadModule('/packages/office-studio/src/api/sheet-filter-panel-commands.ts')
})
afterAll(async () => { await server?.close() })
function setup() {
  const studio = new api.OfficeStudio(), book = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = book.getActiveSheet()
  sheet.getRange('A1:B3').setValues([['Region', 'Value'], ['East', 2], ['West', 1]])
  sheet.setAutoFilter('A1:B3')
  return { studio, book, sheet, input: { unitId: book.getId(), subUnitId: sheet.getSheetId(), col: 1 } }
}

test('the fixed panel operation opens the requested column without touching model, selection or content history', () => {
  const { studio, book, input } = setup(), calls = []
  try {
    const editor = panel.registerSheetFilterPanelEditor(studio, input.unitId, { open: target => { calls.push(target); return true }, close: () => true, release: () => {} })
    const before = JSON.stringify(book.save()), history = JSON.stringify(studio.editor.history.units[input.unitId])
    const selection = structuredClone(studio.getSelectionState(input.unitId, input.subUnitId))
    assert.equal(editor.execute(panel.SHEET_FILTER_PANEL_COMMANDS.open, input), true)
    assert.equal(calls[0].column, 1); assert.equal(calls[0].sheetId, input.subUnitId)
    assert.equal(JSON.stringify(book.save()), before); assert.equal(JSON.stringify(studio.editor.history.units[input.unitId]), history)
    assert.deepEqual(studio.getSelectionState(input.unitId, input.subUnitId), selection)
    assert.equal(editor.execute(panel.SHEET_FILTER_PANEL_COMMANDS.close), true)
    editor.dispose(); assert.equal(editor.execute(panel.SHEET_FILTER_PANEL_COMMANDS.open, input), false)
  } finally { studio.dispose() }
})

test('replaced owners, workbook close, missing hosts and mismatched target aliases never open an obsolete panel', () => {
  const { studio, book, input } = setup(), releases = [], opens = []
  try {
    assert.equal(studio.executeCommandSync(panel.SHEET_FILTER_PANEL_COMMANDS.open, input), false)
    const first = panel.registerSheetFilterPanelEditor(studio, input.unitId, { open: () => { opens.push('first'); return true }, close: () => true, release: () => releases.push('first') })
    const second = panel.registerSheetFilterPanelEditor(studio, input.unitId, { open: () => { opens.push('second'); return true }, close: () => true, release: () => releases.push('second') })
    assert.deepEqual(releases, ['first']); assert.equal(first.execute(panel.SHEET_FILTER_PANEL_COMMANDS.open, input), false)
    assert.throws(() => second.execute(panel.SHEET_FILTER_PANEL_COMMANDS.open, { ...input, column: 0 }), /不一致/)
    assert.throws(() => second.execute(panel.SHEET_FILTER_PANEL_COMMANDS.open, { ...input, extra: true }), /未知/)
    assert.equal(second.execute(panel.SHEET_FILTER_PANEL_COMMANDS.open, { ...input, col: 8 }), false)
    studio.disposeUnit(book.getId())
    assert.deepEqual(releases, ['first', 'second']); assert.equal(second.execute(panel.SHEET_FILTER_PANEL_COMMANDS.open, input), false)
    assert.deepEqual(opens, [])
  } finally { studio.dispose() }
})

test('before-command observer invalidation rejects the captured owner instead of forwarding to a replacement host', () => {
  const { studio, input } = setup()
  let opened = false
  try {
    const editor = panel.registerSheetFilterPanelEditor(studio, input.unitId, { open: () => { opened = true; return true }, close: () => true, release: () => {} })
    const listener = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === panel.SHEET_FILTER_PANEL_COMMANDS.open) editor.dispose() })
    try { assert.throws(() => editor.execute(panel.SHEET_FILTER_PANEL_COMMANDS.open, input), /宿主/) }
    finally { listener.dispose() }
    assert.equal(opened, false)
  } finally { studio.dispose() }
})

test('a protected merged tail outside the filter ref prevents the public panel operation from opening', () => {
  const original = setup(), saved = JSON.parse(JSON.stringify(original.book.save()))
  original.studio.dispose()
  saved.sheets[0].mergedRanges = [{ startRow: 1, endRow: 4, startCol: 0, endCol: 0 }]
  saved.sheets[0].rangeProtectionRules = [{ id: 'tail', permissionId: 'tail-permission', ownerUserId: 'owner', editState: 'onlyMe',
    viewState: 'othersCanView', allowedUserIds: [], ranges: [{ startRow: 4, endRow: 4, startCol: 0, endCol: 0 }] }]
  const studio = new api.OfficeStudio({ workbook: saved })
  let opened = false
  try {
    const before = JSON.stringify(studio.getWorkbookModel(saved.id)), editor = panel.registerSheetFilterPanelEditor(studio, saved.id,
      { open: () => { opened = true; return true }, close: () => true, release: () => {} })
    assert.throws(() => editor.execute(panel.SHEET_FILTER_PANEL_COMMANDS.open, { unitId: saved.id, subUnitId: saved.sheets[0].id, col: 0 }), /保护/)
    assert.equal(opened, false); assert.equal(JSON.stringify(studio.getWorkbookModel(saved.id)), before)
  } finally { studio.dispose() }
})
