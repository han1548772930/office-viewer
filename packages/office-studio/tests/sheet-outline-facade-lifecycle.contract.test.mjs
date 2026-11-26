import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Real Facade / Studio / history drafts. The SSR harness is only created when
// the deferred suite is explicitly run; nothing was started in this batch.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const history = f => f.studio.editor.history.units[f.book.getId()]?.undo.length ?? 0
function setup() {
  const workbook = api.createBlankWorkbook('Outline lifecycle'), sheet = workbook.sheets[0]
  sheet.rowCount = 30; sheet.columnCount = 20
  workbook.sheets.push({ ...json(sheet), id: 'outline-other-sheet', name: 'Other' })
  const studio = new api.OfficeStudio({ workbook }), office = api.FOffice.newAPI(studio), book = office.getActiveWorkbook()
  return { studio, office, book, sheet: book.getActiveSheet(), dispose: () => studio.dispose() }
}

test('six public outline APIs use quantities, stable owners and isolated content history', () => {
  const f = setup()
  try {
    const baseline = history(f)
    assert.equal(f.sheet.addRowOutline(1, 8).addColumnOutline(2, 3), f.sheet)
    assert.deepEqual(f.sheet.getDimensionOutlines().map(group => [group.axis, group.start, group.end]), [['row', 1, 8], ['column', 2, 4]])
    const rows = f.sheet.getDimensionOutlines(f.office.Enum.DimensionOutlineAxis.ROW), rowId = rows[0].id
    assert.equal(rows[0].unitId, f.book.getId()); assert.equal(rows[0].subUnitId, f.sheet.getSheetId())
    assert.equal(f.sheet.setDimensionOutlineCollapsed(rowId, true), f.sheet)
    assert.equal(f.sheet.getDimensionOutlines('row')[0].collapsed, true)
    f.book.undo(); assert.equal(f.sheet.getDimensionOutlines('row')[0].collapsed, false)
    f.book.redo(); assert.equal(f.sheet.getDimensionOutlines('row')[0].collapsed, true)
    assert.equal(f.sheet.removeDimensionOutline(rowId), f.sheet); assert.equal(f.sheet.getDimensionOutlines('column').length, 1)
    assert.equal(f.sheet.clearDimensionOutlines('column', 2, 4), f.sheet); assert.deepEqual(f.sheet.getDimensionOutlines(), [])
    assert.equal(history(f), baseline + 5)
    assert.deepEqual(f.book.getSheetBySheetId('outline-other-sheet').getDimensionOutlines(), [])
  } finally { f.dispose() }
})

test('invalid facade additions are ignored without creating revisions or history entries', () => {
  const f = setup()
  try {
    f.sheet.addRowOutline(2, 8)
    const before = json(f.book.save()), revision = f.studio.getWorkbookRevision(f.book.getId()), baseline = history(f)
    for (const [start, count] of [[-1, 2], [0, 0], [1, -1], [1.5, 2], [29, 2], [6, 10]]) assert.equal(f.sheet.addRowOutline(start, count), f.sheet)
    assert.deepEqual(json(f.book.save()), before); assert.equal(history(f), baseline)
    assert.equal(f.studio.getWorkbookRevision(f.book.getId()), revision)
  } finally { f.dispose() }
})

test('eighth equal-range layer is allowed while the ninth addition leaves the workbook intact', () => {
  const f = setup()
  try {
    for (let depth = 0; depth < 8; depth += 1) f.sheet.addRowOutline(2, 8)
    const before = json(f.book.save()), baseline = history(f)
    f.sheet.addRowOutline(3, 2)
    assert.equal(f.sheet.getDimensionOutlines('row').length, 8); assert.deepEqual(json(f.book.save()), before); assert.equal(history(f), baseline)
    f.sheet.addColumnOutline(2, 3); assert.equal(f.sheet.getDimensionOutlines('column').length, 1)
    f.sheet.setDimensionOutlineCollapsed(f.sheet.getDimensionOutlines('row')[7].id, true)
    assert.equal(f.sheet.getDimensionOutlines('row').every(group => group.collapsed), true)
  } finally { f.dispose() }
})

test('facade permission denial is surfaced and is not swallowed as an ignored invalid range', () => {
  const f = setup()
  try {
    const model = f.studio.getWorkbookModel(f.book.getId())
    model.sheets[0].rangeProtectionRules = [{ id: 'outline-guard', permissionId: 'guard', ownerUserId: 'other-user', editState: 'onlyOwnerCanEdit',
      allowedUserIds: [], ranges: [{ startRow: 3, endRow: 3, startCol: 0, endCol: 19 }] }]
    const before = json(model), baseline = history(f)
    assert.throws(() => f.sheet.addRowOutline(2, 3), /受保护/)
    assert.deepEqual(json(f.studio.getWorkbookModel(f.book.getId())), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('native and public snapshots reopen group IDs, directions and collapsed state defensively', () => {
  const f = setup(); let native, publicStudio
  try {
    f.sheet.addRowOutline(2, 3).addColumnOutline(1, 3)
    f.sheet.setDimensionOutlineCollapsed(f.sheet.getDimensionOutlines('row')[0].id, true)
    f.studio.executeCommandSync('sheet.command.set-outline-summary-settings', { unitId: f.book.getId(), subUnitId: f.sheet.getSheetId(), summaryAbove: true, summaryLeft: true })
    const expected = f.sheet.getDimensionOutlines(), snapshot = f.book.getSnapshot()
    assert.equal(snapshot.resources.filter(resource => resource.name === 'SHEET_OUTLINE_PLUGIN').length, 1)
    native = new api.OfficeStudio({ workbook: f.book.save() }); publicStudio = new api.OfficeStudio({ workbook: snapshot })
    for (const studio of [native, publicStudio]) {
      const reopened = api.FOffice.newAPI(studio).getActiveWorkbook().getActiveSheet()
      assert.deepEqual(reopened.getDimensionOutlines(), expected)
      assert.equal(studio.getWorkbookModel(studio.activeWorkbookId).sheets[0].outline.summaryAbove, true)
      assert.equal(studio.getWorkbookModel(studio.activeWorkbookId).sheets[0].outline.summaryLeft, true)
    }
    const detached = f.sheet.getDimensionOutlines(); detached[0].collapsed = false
    assert.equal(f.sheet.getDimensionOutlines('row')[0].collapsed, true)
  } finally { native?.dispose(); publicStudio?.dispose(); f.dispose() }
})

test('insert/delete/move update the same group ID and a split move rejects every workbook resource', () => {
  const f = setup()
  try {
    f.sheet.addRowOutline(2, 6)
    const id = f.sheet.getDimensionOutlines('row')[0].id
    f.sheet.insertRowsBefore(2, 2)
    assert.deepEqual(f.sheet.getDimensionOutlines('row').map(group => [group.id, group.start, group.end]), [[id, 4, 9]])
    f.sheet.deleteRows(4, 2)
    assert.deepEqual(f.sheet.getDimensionOutlines('row').map(group => [group.id, group.start, group.end]), [[id, 4, 7]])
    const before = json(f.book.save()), baseline = history(f)
    assert.throws(() => f.sheet.moveRows(f.sheet.getRange('A6:T6'), 20), /拆散|split/i)
    assert.deepEqual(json(f.book.save()), before); assert.equal(history(f), baseline)
    f.sheet.moveRows(f.sheet.getRange('A5:T8'), 20)
    assert.deepEqual(f.sheet.getDimensionOutlines('row').map(group => [group.id, group.start, group.end]), [[id, 16, 19]])
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test('captured UI owner fields reject changed content, reset history and another active worksheet', () => {
  const f = setup()
  try {
    const make = () => ({ unitId: f.book.getId(), sheetId: f.sheet.getSheetId(), action: { kind: 'group', axis: 'row', start: 2, end: 4, groupId: 'draft' },
      host: 'workbench', expectedRevision: f.studio.getWorkbookRevision(f.book.getId()),
      expectedHistoryGeneration: f.studio.editor.history.captureSavepoint(f.book.getId()).generation, expectedUserId: f.studio.currentUserId })
    const stale = make(); f.sheet.addColumnOutline(1, 3)
    assert.throws(() => f.studio.executeCommandSync('sheet.command.change-outline', stale), /过期|更改|版本/)
    const reset = make(); f.studio.editor.history.clear(f.book.getId())
    assert.throws(() => f.studio.executeCommandSync('sheet.command.change-outline', reset), /过期|更改|版本|历史/)
    const wrongHost = make(); f.book.setActiveSheet('outline-other-sheet')
    assert.throws(() => f.studio.executeCommandSync('sheet.command.change-outline', wrongHost), /过期|更改|目标|活动|工作表/)
  } finally { f.dispose() }
})

test('content reentry in BeforeCommandExecute rejects a stale draft without publishing its group', () => {
  const f = setup()
  try {
    const request = { unitId: f.book.getId(), sheetId: f.sheet.getSheetId(), action: { kind: 'group', axis: 'row', start: 2, end: 4, groupId: 'stale-draft' },
      host: 'workbench', expectedRevision: f.studio.getWorkbookRevision(f.book.getId()),
      expectedHistoryGeneration: f.studio.editor.history.captureSavepoint(f.book.getId()).generation, expectedUserId: f.studio.currentUserId }
    const subscription = f.studio.events.addEvent('BeforeCommandExecute', event => {
      if (event.id !== 'sheet.command.change-outline') return
      subscription.dispose(); f.sheet.addColumnOutline(1, 3)
    })
    assert.throws(() => f.studio.executeCommandSync('sheet.command.change-outline', request), /过期|更改|版本/)
    assert.deepEqual(f.sheet.getDimensionOutlines('row'), []); assert.equal(f.sheet.getDimensionOutlines('column').length, 1)
  } finally { f.dispose() }
})
