import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { createSheetDataValidationRule } from '../src/features/sheet-validation/validation-engine.ts'
import { planSheetCheckboxRuleCells } from '../src/features/sheet-validation/validation-checkbox-rule-cells.ts'
import { planSheetDataValidationCommand, sheetDataValidationCommandInput, SHEET_DATA_VALIDATION_COMMANDS } from '../src/features/sheet-validation/validation-rule-command.ts'

// Implementation-phase contracts. Models, actual command/history boundaries and internal reopen remain unrun.
// Small deterministic worksheet bounds here do not establish Worker, memory, cancellation or performance acceptance.
let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = (row, column = 0, endRow = row, endCol = column) => ({ startRow: row, endRow, startCol: column, endCol })
function modelFixture(configure = () => {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'checkbox-rule-cells'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'
  sheet.rowCount = 12; sheet.columnCount = 8; sheet.cells = {}; sheet.dataValidations = []
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other); configure(model, sheet)
  return { model, sheet }
}
function commandFixture(configure = () => {}) {
  const fixture = modelFixture(configure), studio = new OfficeStudio({ workbook: fixture.model, currentUserId: 'local-user' })
  const book = FOffice.newAPI(studio).getActiveWorkbook(), current = () => studio.getWorkbookModel(fixture.model.id)
  const target = () => ({ unitId: fixture.model.id, sheetId: 'data', expectedRevision: studio.getWorkbookRevision(fixture.model.id),
    expectedHistoryGeneration: studio.editor.history.captureSavepoint(fixture.model.id).generation,
    expectedUserId: studio.currentUserId, expectedRulesSignature: JSON.stringify(current().sheets[0].dataValidations) })
  const execute = (action, payload) => studio.executeCommandSync(SHEET_DATA_VALIDATION_COMMANDS[action], { ...target(), ...payload })
  return { studio, book, current, target, execute, count: () => studio.editor.history.units[fixture.model.id]?.undo.length ?? 0, dispose: () => studio.dispose() }
}
const rule = (ranges, options = {}) => createSheetDataValidationRule({ id: 'checkbox', type: 'checkbox', ...options }, ranges)
const plan = (model, sheet, next, previous) => planSheetCheckboxRuleCells(model, sheet, next, previous, 'local-user', 'zh-CN')

test('new checkbox initialization visits overlapping discontinuous ranges once, clips to real worksheet bounds and preserves unrelated cells', () => {
  const { model, sheet } = modelFixture((_model, sheet) => {
    sheet.rowCount = 3; sheet.columnCount = 3; sheet.cells['0:0'] = createCell(0, 0, 1); sheet.cells['0:1'] = createCell(0, 1, 'keep')
    sheet.cells['2:2'] = createCell(2, 2, true)
  })
  const before = cloneModel(model), next = rule([area(0, 0, 1, 1), area(1, 1, 20, 20)])
  const prepared = plan(model, sheet, next)
  assert.deepEqual(prepared.cells.map(cell => cell.id), ['1:0', '1:1', '1:2', '2:1'])
  assert.ok(prepared.cells.every(cell => cell.value === 0 && cell.dataType === 'number'))
  assert.deepEqual(model, before); assert.equal(new Set(prepared.cells.map(cell => cell.id)).size, prepared.cells.length)
})

test.each([
  [{ formula1: 'Yes', formula2: 'No' }, 'No', 'text'],
  [{ formula1: 'TRUE', formula2: 'FALSE' }, 'FALSE', 'text'],
  [{ formula1: '1', formula2: '0' }, '0', 'text'],
  [{ formula1: 'Done' }, 0, 'number'],
  [{ formula2: 'Waiting' }, 'Waiting', 'text'],
  [{ formula1: '=5', formula2: '=1+1' }, 2, 'number'],
])('unchecked initialization saves the original scalar for %j without parsing text as a formula or boolean', (options, expected, type) => {
  const { model, sheet } = modelFixture(), prepared = plan(model, sheet, rule([area(0)], options))
  assert.equal(prepared.cells.length, 1); assert.equal(prepared.cells[0].value, expected); assert.equal(prepared.cells[0].dataType, type)
  assert.equal(prepared.cells[0].formula, null)
})

test('styled blanks retain authored formatting, hyperlink and metadata while stale value caches are replaced', () => {
  const { model, sheet } = modelFixture((_model, sheet) => {
    const blank = createCell(0, 0, null); blank.style.numberFormat = '0.00'; blank.style.font.bold = true
    blank.authoredStyle = { font: { bold: true }, numberFormat: '0.00' }; blank.hyperlink = 'https://example.test/retained'
    blank.metadata = { retained: ['cell'] }; blank.cachedDataType = 'date'; blank.display = 'old display'; sheet.cells['0:0'] = blank
  })
  const before = cloneModel(sheet.cells['0:0']), changed = plan(model, sheet, rule([area(0)])).cells[0]
  assert.equal(changed.value, 0); assert.equal(changed.dataType, 'number'); assert.equal(changed.cachedDataType, null)
  assert.deepEqual(changed.authoredStyle, before.authoredStyle); assert.equal(changed.style.font.bold, true); assert.equal(changed.style.numberFormat, '0.00')
  assert.equal(changed.hyperlink, before.hyperlink); assert.deepEqual(changed.metadata, before.metadata); assert.deepEqual(sheet.cells['0:0'], before)
})

test('existing invalid text, rich text, images and formulas are preserved when a rule is added or its custom values are changed', () => {
  const { model, sheet } = modelFixture((_model, sheet) => {
    sheet.cells['0:0'] = createCell(0, 0, 'invalid'); sheet.cells['1:0'] = createCell(1, 0, 0); sheet.cells['1:0'].richText = [{ text: '0' }]
    sheet.cells['2:0'] = createCell(2, 0, null); sheet.cells['2:0'].image = { id: 'retained-image',
      resource: { id: 'retained-resource', name: 'retained.png', mimeType: 'image/png', source: 'data:image/png;base64,AA==', width: 1, height: 1, byteLength: 1 },
      altText: 'Retained image', title: 'Retained image' }
    sheet.cells['3:0'] = createCell(3, 0, 0); sheet.cells['3:0'].formula = '=1-1'
  })
  const before = cloneModel(model), previous = rule([area(0, 0, 3)]), next = rule(previous.ranges, { formula1: 'Yes', formula2: 'No' })
  assert.deepEqual(plan(model, sheet, previous).cells, []); assert.deepEqual(plan(model, sheet, next, previous).cells, []); assert.deepEqual(model, before)
})

test.each([0, 1, false, true, 'FALSE', 'TRUE'])('custom value migration carries state for old scalar %j and preserves invalid text', actual => {
  const { model, sheet } = modelFixture((_model, sheet) => { sheet.cells['0:0'] = createCell(0, 0, actual); sheet.cells['1:0'] = createCell(1, 0, 'invalid') })
  const previous = rule([area(0, 0, 1)]), next = rule(previous.ranges, { formula1: 'Yes', formula2: 'No' })
  const prepared = plan(model, sheet, next, previous), checked = actual === 1 || actual === true || actual === 'TRUE'
  assert.equal(prepared.cells.length, 1); assert.equal(prepared.cells[0].value, checked ? 'Yes' : 'No'); assert.equal(prepared.cells[0].dataType, 'text')
  assert.equal(sheet.cells['1:0'].value, 'invalid'); assert.equal(sheet.cells['0:0'].value, actual)
})

test('range expansion fills missing old and new blanks while keeping existing checked values and hidden rows', () => {
  const { model, sheet } = modelFixture((_model, sheet) => { sheet.cells['0:0'] = createCell(0, 0, 1); sheet.hiddenRows = [2]; sheet.filteredRows = [3] })
  const previous = rule([area(0, 0, 1)]), next = rule([area(0, 0, 3)])
  const prepared = plan(model, sheet, next, previous)
  assert.deepEqual(prepared.cells.map(cell => [cell.row, cell.value]), [[1, 0], [2, 0], [3, 0]])
  assert.deepEqual(prepared.ranges, [area(1, 0, 3)])
})

test('a range shrink or move leaves outside values intact and migrates only surviving old coverage', () => {
  const { model, sheet } = modelFixture((_model, sheet) => {
    sheet.cells['0:0'] = createCell(0, 0, 'Yes'); sheet.cells['1:0'] = createCell(1, 0, 'No'); sheet.cells['2:0'] = createCell(2, 0, 'Yes')
    sheet.cells['4:0'] = createCell(4, 0, 'No')
  })
  const previous = rule([area(0, 0, 2)], { formula1: 'Yes', formula2: 'No' }), next = rule([area(1), area(4, 0, 5)], { formula1: 'Done', formula2: 'Waiting' })
  const prepared = plan(model, sheet, next, previous)
  assert.deepEqual(prepared.cells.map(cell => [cell.row, cell.value]), [[1, 'Waiting'], [5, 'Waiting']])
  assert.equal(sheet.cells['0:0'].value, 'Yes'); assert.equal(sheet.cells['2:0'].value, 'Yes'); assert.equal(sheet.cells['4:0'].value, 'No')
})

test('message-only edits and unchanged imported definitions never materialize additional scalar cells', () => {
  const { model, sheet } = modelFixture(), previous = rule([area(0, 0, 3)]), next = { ...cloneModel(previous), prompt: 'Choose a value', showInputMessage: true }
  assert.deepEqual(plan(model, sheet, next, previous), { cells: [], ranges: [] }); assert.deepEqual(plan(model, sheet, previous, previous), { cells: [], ranges: [] })
  assert.deepEqual(sheet.cells, {})
})

test('a merged rule materializes only its main cell and permission planning includes every merged member', () => {
  const { model, sheet } = modelFixture((_model, sheet) => { sheet.mergedRanges = [area(0, 0, 1, 1)] })
  const next = rule([area(0)]), input = sheetDataValidationCommandInput('set', { unitId: model.id, sheetId: sheet.id, rule: next })
  assert.deepEqual(plan(model, sheet, next).cells.map(cell => cell.id), ['0:0'])
  assert.ok(planSheetDataValidationCommand(sheet, input).ranges.some(range => JSON.stringify(range) === JSON.stringify(sheet.mergedRanges[0])))
  assert.deepEqual(plan(model, sheet, rule([area(0, 0, 1, 1)])).cells.map(cell => cell.id), ['0:0'])
})

test('formula operands use the rule anchor and current scalar source, while missing or cyclic operands prevent initialization', () => {
  const { model, sheet } = modelFixture((_model, sheet) => { sheet.cells['2:3'] = createCell(2, 3, 9); sheet.cells['2:4'] = createCell(2, 4, 8) })
  const next = rule([area(0)], { formula1: '=D3', formula2: '=E3', formulaAnchor: { row: 2, col: 0 } })
  assert.equal(plan(model, sheet, next).cells[0].value, 8)
  sheet.cells['2:4'].value = 7; assert.equal(plan(model, sheet, next).cells[0].value, 7)
  sheet.cells['2:4'].formula = '=E3'; assert.throws(() => plan(model, sheet, next), /无法读取/); assert.equal(sheet.cells['0:0'], undefined)
})

test('View denial is checked before target scalar getters and formula source permissions prevent copying hidden defaults', () => {
  const { model, sheet } = modelFixture((_model, sheet) => {
    const hidden = createCell(0, 0, null); Object.defineProperty(hidden, 'value', { configurable: true, get: () => { throw new Error('hidden target was read') } })
    sheet.cells['0:0'] = hidden; sheet.rangeProtectionRules = [{ id: 'hidden-target', ranges: [area(0)], ownerUserId: 'peer',
      allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'onlyMe' }]
  })
  assert.throws(() => plan(model, sheet, rule([area(0)])), /查看/)
  const source = modelFixture((_model, sheet) => {
    sheet.cells['0:3'] = createCell(0, 3, 'hidden unchecked'); sheet.rangeProtectionRules = [{ id: 'source', ranges: [area(0, 3)], ownerUserId: 'peer',
      allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'onlyMe' }]
  })
  assert.throws(() => plan(source.model, source.sheet, rule([area(0)], { formula1: 'Yes', formula2: '=D1' })), /无法读取/)
  assert.equal(source.sheet.cells['0:0'], undefined)
})

test('the actual set command commits rules and initialized cells in one history entry and one coalesced value event', () => {
  const h = commandFixture(), events = [], subscription = h.studio.events.addEvent('SheetValueChanged', event => events.push(cloneModel(event)))
  try {
    const before = cloneModel(h.current()), next = rule([area(0, 0, 1), area(0, 2, 1, 2)])
    const frozen = Object.freeze({ ...h.target(), rule: Object.freeze(cloneModel(next)) }), original = JSON.stringify(frozen)
    assert.equal(h.studio.executeCommandSync(SHEET_DATA_VALIDATION_COMMANDS.set, frozen), true); assert.equal(h.count(), 1); assert.equal(JSON.stringify(frozen), original)
    const after = cloneModel(h.current()); assert.deepEqual(Object.values(after.sheets[0].cells).map(cell => cell.value), [0, 0, 0, 0])
    assert.deepEqual(after.sheets[1], before.sheets[1]); assert.equal(events.length, 1)
    assert.deepEqual(events[0].ranges.map(range => [range.startRow, range.endRow, range.startColumn, range.endColumn]), [[0, 1, 0, 0], [0, 1, 2, 2]])
    h.book.undo(); assert.deepEqual(h.current(), before); h.book.redo(); assert.deepEqual(h.current(), after)
    const reopened = new OfficeStudio({ workbook: h.book.save() })
    try { assert.deepEqual(reopened.getWorkbookModel(after.id).sheets[0], after.sheets[0]) } finally { reopened.dispose() }
  } finally { subscription.dispose(); h.dispose() }
})

test('custom setting replacement migrates both states and blank scalars atomically but preserves unmatched and formula cells', () => {
  const h = commandFixture((_model, sheet) => {
    sheet.dataValidations = [rule([area(0, 0, 4)])]
    sheet.cells['0:0'] = createCell(0, 0, 1); sheet.cells['1:0'] = createCell(1, 0, false); sheet.cells['2:0'] = createCell(2, 0, 'invalid')
    sheet.cells['4:0'] = createCell(4, 0, 0); sheet.cells['4:0'].formula = '=1-1'
  })
  try {
    const before = cloneModel(h.current()), next = { ...cloneModel(before.sheets[0].dataValidations[0]), formula1: 'Yes', formula2: 'No' }
    assert.equal(h.execute('replace', { rule: next }), true); assert.equal(h.count(), 1)
    const after = cloneModel(h.current()); assert.equal(after.sheets[0].cells['0:0'].value, 'Yes'); assert.equal(after.sheets[0].cells['1:0'].value, 'No')
    assert.equal(after.sheets[0].cells['3:0'].value, 'No'); assert.deepEqual(after.sheets[0].cells['2:0'], before.sheets[0].cells['2:0'])
    assert.deepEqual(after.sheets[0].cells['4:0'], before.sheets[0].cells['4:0'])
    h.book.undo(); assert.deepEqual(h.current(), before); h.book.redo(); assert.deepEqual(h.current(), after)
  } finally { h.dispose() }
})

test('a partial same-ID set leaves the residual old rule and its cells untouched', () => {
  const h = commandFixture((_model, sheet) => { sheet.dataValidations = [rule([area(0, 0, 2)])]; sheet.cells['1:0'] = createCell(1, 0, 1); sheet.cells['2:0'] = createCell(2, 0, 0) })
  try {
    assert.equal(h.execute('set', { rule: rule([area(1)], { formula1: 'Done', formula2: 'Waiting' }) }), true)
    const sheet = h.current().sheets[0]; assert.equal(sheet.cells['0:0'], undefined); assert.equal(sheet.cells['1:0'].value, 'Done'); assert.equal(sheet.cells['2:0'].value, 0)
    const residual = sheet.dataValidations.find(rule => rule.id !== 'checkbox'); assert.ok(residual)
    assert.equal(residual.formula1, undefined); assert.equal(residual.formula2, undefined); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test.each(['view', 'edit'])('a protected member of a merge rejects %s access before publishing any rule or initialized cell', access => {
  const h = commandFixture((_model, sheet) => {
    sheet.mergedRanges = [area(0, 0, 1, 1)]; sheet.rangeProtectionRules = [{ id: 'merge-member', ranges: [area(1, 1)], ownerUserId: 'peer', allowedUserIds: [],
      viewState: access === 'view' ? 'noOneElseCanView' : 'othersCanView', editState: 'onlyMe' }]
  })
  try {
    const before = cloneModel(h.current()); assert.throws(() => h.execute('set', { rule: rule([area(0)]) }), /权限|查看|保护/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('cancellation and preparation-time revocation retain the exact rule/cell snapshot and history', () => {
  const h = commandFixture(), before = cloneModel(h.current())
  try {
    const cancel = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === SHEET_DATA_VALIDATION_COMMANDS.set) event.cancel = true })
    assert.throws(() => h.execute('set', { rule: rule([area(0, 0, 2)]) }), /canceled/); cancel.dispose(); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    const original = h.studio.assertCanEditRange.bind(h.studio); let candidatePrepared = false
    const commit = h.studio.commitWorkbook.bind(h.studio)
    h.studio.commitWorkbook = (id, label, mutate, options) => commit(id, label, model => { mutate(model); candidatePrepared = true }, options)
    h.studio.assertCanEditRange = (...args) => { if (candidatePrepared) throw new Error('revoked after cell initialization'); return original(...args) }
    assert.throws(() => h.execute('set', { rule: rule([area(0, 0, 2)]) }), /revoked/); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('unchanged rules, clear and remove preserve scalar cells and do not publish spurious value events', () => {
  const h = commandFixture((_model, sheet) => { sheet.dataValidations = [rule([area(0, 0, 2)])]; sheet.cells['0:0'] = createCell(0, 0, 1) }), events = []
  const subscription = h.studio.events.addEvent('SheetValueChanged', event => events.push(event))
  try {
    const original = cloneModel(h.current().sheets[0]), existing = cloneModel(original.dataValidations[0])
    assert.equal(h.execute('replace', { rule: existing }), false); assert.equal(h.count(), 0); assert.equal(h.current().sheets[0].cells['1:0'], undefined)
    assert.equal(h.execute('clearTypes', { ranges: [area(0)], types: ['checkbox'] }), true); assert.deepEqual(h.current().sheets[0].cells, original.cells)
    assert.equal(h.execute('remove', { ruleId: 'checkbox' }), true); assert.deepEqual(h.current().sheets[0].cells, original.cells); assert.deepEqual(events, [])
  } finally { subscription.dispose(); h.dispose() }
})

test('an inactive API workbook keeps its rule/cell history isolated and a value observer may switch sheets without rolling back the committed scalar', () => {
  const h = commandFixture(), peer = FOffice.newAPI(h.studio).createWorkbook({ id: 'checkbox-peer' })
  try {
    const beforePeer = cloneModel(h.studio.getWorkbookModel(peer.getId()))
    assert.equal(h.execute('set', { rule: rule([area(0)]) }), true); assert.equal(h.count(), 1); assert.deepEqual(h.studio.getWorkbookModel(peer.getId()), beforePeer)
    h.studio.activateWorkbook(h.book.getId()); const subscription = h.studio.events.addEvent('SheetValueChanged', event => { if (event.workbookId === h.book.getId()) h.book.getSheetBySheetId('other').activate() })
    try {
      assert.equal(h.execute('replace', { rule: rule([area(0)], { formula1: 'Yes', formula2: 'No' }) }), true)
      assert.equal(h.current().sheets[0].cells['0:0'].value, 'No'); assert.equal(h.current().activeSheetId, 'other'); assert.equal(h.count(), 2)
    } finally { subscription.dispose() }
  } finally { h.dispose() }
})
