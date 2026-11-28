import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred command contracts: a protected merge tail is outside the native ref.
let server, sdk, context
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  sdk = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  context = await server.ssrLoadModule('/packages/office-studio/src/api/sheet-filter-command-context.ts')
})
afterAll(async () => { await server?.close() })
const range = { startRow: 0, endRow: 2, startCol: 0, endCol: 0 }
function fixture(t, existing = true, protectedRange = { startRow: 4, endRow: 4, startCol: 0, endCol: 0 }) {
  const model = sdk.createBlankWorkbook(); model.id = 'filter-permission-book'
  const sheet = model.sheets[0]; sheet.id = 'data'; model.activeSheetId = sheet.id
  sheet.cells['0:0'] = sdk.createCell(0, 0, 'Header'); sheet.cells['1:0'] = sdk.createCell(1, 0, 'West')
  sheet.mergedRanges = [{ startRow: 1, endRow: 4, startCol: 0, endCol: 0 }]
  if (existing) sheet.autoFilter = { range: { ...range }, columns: { 0: { kind: 'values', selectedItems: ['East'] } }, filteredRows: [1, 2, 3, 4] }
  if (protectedRange) sheet.rangeProtectionRules = [{ id: 'tail', permissionId: 'tail', ranges: [protectedRange],
    ownerUserId: 'another-user', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '' }]
  const studio = new sdk.OfficeStudio({ workbook: model, sheetTextMeasure: text => text.length * 8 })
  t.onTestFinished(() => studio.dispose())
  const book = sdk.FOffice.newAPI(studio).getActiveWorkbook()
  return { studio, book, request: { unitId: book.getId(), sheetId: 'data' } }
}

test('criteria, clearing, removal, toggle, configure and reapply all reject protection on an intersecting merge tail', t => {
  const h = fixture(t), before = JSON.stringify(h.book.save())
  const history = JSON.stringify(h.studio.editor.history.units[h.book.getId()])
  const commands = [
    ['sheet.command.set-filter-criteria', { column: 0, criteria: { kind: 'values', selectedItems: ['West'] } }],
    ['sheet.command.clear-filter', {}], ['sheet.command.set-auto-filter', { range: null }],
    ['sheet.command.toggle-auto-filter', { range }],
    ['sheet.command.configure-filter', { range, columns: { 0: { kind: 'values', selectedItems: ['West'] } } }],
    ['sheet.command.reapply-filter', {}],
  ]
  for (const [id, params] of commands) assert.throws(() => h.studio.executeCommandSync(id, { ...h.request, ...params }), /保护|权限/)
  assert.equal(JSON.stringify(h.book.save()), before)
  assert.equal(JSON.stringify(h.studio.editor.history.units[h.book.getId()]), history)
})

test('new filters and configuration check their new footprint before a protected merge group becomes filtered', t => {
  const h = fixture(t, false), before = JSON.stringify(h.book.save())
  assert.throws(() => h.studio.executeCommandSync('sheet.command.set-auto-filter', { ...h.request, range }), /保护|权限/)
  assert.throws(() => h.studio.executeCommandSync('sheet.command.configure-filter', { ...h.request, range, columns: {} }), /保护|权限/)
  assert.equal(JSON.stringify(h.book.save()), before)
})

test('an unrelated protected cell does not prevent filtering and reapply creates no content history', t => {
  const h = fixture(t, true, { startRow: 8, endRow: 8, startCol: 6, endCol: 6 })
  const depth = h.studio.editor.history.units[h.book.getId()].undo.length
  assert.equal(h.studio.executeCommandSync('sheet.command.set-filter-criteria', { ...h.request, column: 0, criteria: { kind: 'values', selectedItems: ['West'] } }), true)
  assert.deepEqual(h.book.getActiveSheet().getFilteredRows(), [])
  assert.equal(h.studio.editor.history.units[h.book.getId()].undo.length, depth + 1)
  h.studio.executeCommandSync('sheet.command.reapply-filter', h.request)
  assert.equal(h.studio.editor.history.units[h.book.getId()].undo.length, depth + 1)
})

test('filter removal ignores an unrelated passed selection, while retaining authorization over its original range', t => {
  const protectedRange = { startRow: 8, endRow: 9, startCol: 6, endCol: 6 }, h = fixture(t, true, protectedRange)
  assert.equal(h.studio.executeCommandSync('sheet.command.toggle-auto-filter', { ...h.request, range: protectedRange }), true)
  assert.equal(h.book.getActiveSheet().getAutoFilter(), null)
})

test('captured filter contexts reject changed requests, users and history before publication', t => {
  const h = fixture(t, false, null), request = { ...h.request, range: { ...range } }
  const guard = context.sheetFilterCommandContext(h.studio, request, request.range)
  request.range.endRow++
  assert.throws(() => guard.assertCurrent(), /参数|更改/)
  request.range.endRow--
  h.studio.editor.history.clear(h.book.getId())
  assert.throws(() => guard.assertCurrent(), /历史|更改/)
})
