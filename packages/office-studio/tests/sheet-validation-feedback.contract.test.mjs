import assert from 'node:assert/strict'
import { test } from 'vitest'
import { shallowReactive } from 'vue'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { sheetValidationCellFeedback } from '../src/features/sheet-validation/validation-cell-feedback.ts'
import { sheetValidationSourceReader } from '../src/features/sheet-validation/validation-value-read.ts'
import { serializeValidationListItems, validateSheetDataValidationChanges } from '../src/features/sheet-validation/validation-engine.ts'
import { useSheetValidationFeedback } from '../src/components/spreadsheet/use-sheet-validation-feedback.ts'

// Implementation-phase drafts. The contracts and shared formula permission boundary have not been run.
const area = (row, column, endRow = row, endCol = column) => ({ startRow: row, startCol: column, endRow, endCol })
const point = (row, column) => ({ row, column })
const lock = ranges => ({ id: 'private', ranges, ownerUserId: 'peer', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'onlyMe' })
function workbook() {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'feedback'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'; sheet.rowCount = 80; sheet.columnCount = 20
  sheet.dataValidations = [{ id: 'whole', type: 'whole', operator: 'between', formula1: '1', formula2: '5', ranges: [area(0, 0)],
    showInputMessage: true, promptTitle: 'Quantity', prompt: 'Enter 1 through 5', errorTitle: 'Integer', error: 'Use a valid quantity' }]
  sheet.cells['0:0'] = createCell(0, 0, 1.5)
  return { model, sheet, read: positions => sheetValidationCellFeedback(model, sheet.id, positions, 'local-user') }
}

test('requested positions deduplicate to one merged main cell and preserve both prompt and invalid feedback', () => {
  const h = workbook(); h.sheet.mergedRanges = [area(0, 0, 1, 1)]
  const before = cloneModel(h.model), result = h.read([point(0, 0), point(1, 1), point(0, 1), point(0, 0), point(20, 12), point(-1, 0), point(80, 0)])
  assert.deepEqual([...result.keys()], ['A1'])
  assert.deepEqual(result.get('A1'), { key: 'A1', row: 0, column: 0, prompt: { ruleId: 'whole', title: 'Quantity', message: 'Enter 1 through 5' },
    issues: [{ ruleId: 'whole', type: 'whole', severity: 'stop', title: 'Integer', message: 'Use a valid quantity' }] })
  assert.deepEqual(h.model, before)
})

test.each([
  ['whole', 1.5, {}], ['decimal', 'not a number', {}], ['textLength', 'a', { operator: 'equal', formula1: '2' }],
  ['date', 'not a date', {}], ['time', -1, {}], ['checkbox', 2, {}],
  ['list', 'missing', { formula1: serializeValidationListItems(['allowed']) }],
  ['listMultiple', 'allowed,missing', { formula1: serializeValidationListItems(['allowed']) }],
  ['custom', 1, { formula1: '=FALSE()' }],
])('%s feedback comes from the real validator without carrying an original cell value', (type, value, settings) => {
  const h = workbook(); h.sheet.cells['0:0'] = createCell(0, 0, value)
  h.sheet.dataValidations = [{ id: type, type, ranges: [area(0, 0)], ...settings }]
  const feedback = h.read([point(0, 0)]).get('A1')
  assert.equal(feedback.issues.length, 1); assert.equal(feedback.issues[0].type, type)
  assert.equal(Object.hasOwn(feedback.issues[0], 'value'), false); assert.equal(Object.hasOwn(feedback, 'rule'), false)
  assert.equal(Object.hasOwn(feedback, 'ranges'), false)
})

test('any rules can carry an input prompt, and required blank cells receive an invalid notice', () => {
  const h = workbook(); delete h.sheet.cells['0:0']
  h.sheet.dataValidations[0].allowBlank = false
  assert.equal(h.read([point(0, 0)]).get('A1').issues.length, 1)
  h.sheet.dataValidations[0].type = 'any'
  const feedback = h.read([point(0, 0)]).get('A1')
  assert.equal(feedback.issues.length, 0); assert.equal(feedback.prompt.title, 'Quantity')
  h.sheet.dataValidations[0].showInputMessage = false
  assert.equal(h.read([point(0, 0)]).get('A1').prompt, undefined)
})

test('a hidden part of a merged target excludes the entire prompt and invalid notice', () => {
  const h = workbook(); h.sheet.mergedRanges = [area(0, 0, 1, 1)]; h.sheet.rangeProtectionRules = [lock([area(1, 1)])]
  assert.equal(h.read([point(0, 0), point(1, 1)]).size, 0)
  assert.equal(sheetValidationCellFeedback(h.model, 'absent', [point(0, 0)], 'local-user').size, 0)
})

test('a denied source range yields no option values or source coordinates in a target notice', () => {
  const h = workbook(), secret = createBlankSheet('Secret'); secret.id = 'secret'; secret.rowCount = 80; secret.columnCount = 20
  secret.cells['0:0'] = createCell(0, 0, 'private-token'); secret.rangeProtectionRules = [lock([area(0, 0)])]; h.model.sheets.push(secret)
  h.sheet.dataValidations = [{ id: 'list', type: 'list', formula1: '=Secret!$A$1', ranges: [area(0, 0)], metadata: { source: 'retained-only' } }]
  h.sheet.cells['0:0'] = createCell(0, 0, 'private-token')
  const feedback = h.read([point(0, 0)]).get('A1'), serialized = JSON.stringify(feedback)
  assert.equal(feedback.issues.length, 1); assert.equal(serialized.includes('private-token'), false)
  assert.equal(serialized.includes('Secret'), false); assert.equal(serialized.includes('retained-only'), false)
})

test('source permission checks expand a merged anchor before any reference value is read', () => {
  const h = workbook(); h.sheet.mergedRanges = [area(0, 2, 1, 3)]; h.sheet.cells['0:2'] = createCell(0, 2, 'private-token')
  h.sheet.rangeProtectionRules = [lock([area(1, 3)])]
  const canRead = sheetValidationSourceReader('local-user')
  assert.equal(canRead(h.sheet, area(0, 2)), false); assert.equal(canRead(h.sheet, area(20, 2)), true)
  h.sheet.dataValidations = [{ id: 'list', type: 'list', formula1: '=$C$1', ranges: [area(0, 0)] }]
  h.sheet.cells['0:0'] = createCell(0, 0, 'private-token')
  assert.equal(h.read([point(0, 0)]).get('A1').issues.length, 1)
})

test.each(['custom', 'checkbox', 'cell-formula'])('%s nested references respect the shared formula source reader', kind => {
  const h = workbook(), secret = createBlankSheet('Secret'); secret.id = 'secret'; secret.rowCount = 80; secret.columnCount = 20
  secret.cells['0:0'] = createCell(0, 0, 1); secret.rangeProtectionRules = [lock([area(0, 0)])]; h.model.sheets.push(secret)
  h.sheet.cells['0:2'] = createCell(0, 2, 1); h.sheet.cells['0:2'].formula = '=Secret!A1'
  if (kind === 'custom') h.sheet.dataValidations = [{ id: kind, type: 'custom', formula1: '=C1=1', ranges: [area(0, 0)] }]
  else if (kind === 'checkbox') h.sheet.dataValidations = [{ id: kind, type: 'checkbox', formula1: '=C1', formula2: '0', ranges: [area(0, 0)] }]
  else { h.sheet.cells['0:0'].formula = '=C1'; h.sheet.dataValidations[0].operator = 'equal'; h.sheet.dataValidations[0].formula1 = '1' }
  h.sheet.cells['0:0'].value = 1
  assert.equal(h.read([point(0, 0)]).get('A1').issues.length, 1)
  secret.rangeProtectionRules = []
  assert.equal(h.read([point(0, 0)]).get('A1').issues.length, 0)
})

test('cell formulas are evaluated separately from stale cached values and source changes refresh the invalid result', () => {
  const h = workbook(); h.sheet.cells['0:2'] = createCell(0, 2, 2); h.sheet.cells['0:0'].value = 'stale'; h.sheet.cells['0:0'].formula = '=C1'
  const before = cloneModel(h.model)
  assert.equal(h.read([point(0, 0)]).get('A1').issues.length, 0); assert.deepEqual(h.model, before)
  h.sheet.cells['0:2'].value = 8
  assert.equal(h.read([point(0, 0)]).get('A1').issues.length, 1)
  assert.equal(h.sheet.cells['0:0'].value, 'stale')
})

test('disabled custom error text is omitted while severity and prompt remain independent', () => {
  const h = workbook(); Object.assign(h.sheet.dataValidations[0], { showErrorMessage: false, errorStyle: 'warning' })
  const feedback = h.read([point(0, 0)]).get('A1')
  assert.deepEqual(feedback.issues, [{ ruleId: 'whole', type: 'whole', severity: 'warning' }]); assert.equal(feedback.prompt.title, 'Quantity')
})

test('one batch resolves a shared list candidate range once and validates every requested target', () => {
  const h = workbook(); h.sheet.cells['0:2'] = createCell(0, 2, 'allowed'); h.sheet.cells['1:2'] = createCell(1, 2, 'second')
  for (let row = 0; row < 30; row++) h.sheet.cells[`${row}:0`] = createCell(row, 0, row % 2 ? 'missing' : 'allowed')
  h.sheet.dataValidations = [{ id: 'list', type: 'list', formula1: '=$C$1:$C$2', ranges: [area(0, 0, 29, 0)] }]
  let firstReads = 0, secondReads = 0
  Object.defineProperty(h.sheet.cells['0:2'], 'value', { enumerable: true, configurable: true, get: () => { firstReads++; return 'allowed' } })
  Object.defineProperty(h.sheet.cells['1:2'], 'value', { enumerable: true, configurable: true, get: () => { secondReads++; return 'second' } })
  const result = validateSheetDataValidationChanges(h.model, h.sheet.id, Array.from({ length: 30 }, (_, row) => point(row, 0)),
    () => true)
  assert.equal(firstReads, 1); assert.equal(secondReads, 1); assert.equal(result.issues.length, 15)
})

test('standalone feedback is available with View-only permission and rejects inactive or temporarily blocked hosts', () => {
  const h = workbook(); h.sheet.rangeProtectionRules = [{ ...lock([area(0, 0)]), viewState: 'othersCanView' }]
  const state = shallowReactive({ model: toExcelWorkbookView(h.model), sheetId: h.sheet.id, userId: 'local-user', region: 'en-US', active: true, blocked: false })
  const feedback = useSheetValidationFeedback(() => state)
  assert.equal(feedback.read([point(0, 0)]).get('A1').issues.length, 1)
  state.blocked = true; assert.equal(feedback.read([point(0, 0)]).size, 0)
  state.blocked = false; state.active = false; assert.equal(feedback.read([point(0, 0)]).size, 0)
  state.active = true; state.sheetId = 'absent'; assert.equal(feedback.read([point(0, 0)]).size, 0)
})
