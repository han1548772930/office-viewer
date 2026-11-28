import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { createSheetDataValidationRule, sheetDataValidationOptions, validateSheetDataValidationChanges } from '../src/features/sheet-validation/validation-engine.ts'
import { sheetValidationSourceReader } from '../src/features/sheet-validation/validation-value-read.ts'
import { validationPanelSettings, validationRuleFromPanel } from '../src/features/sheet-validation/validation-panel-state.ts'
import { planSheetValidationListInput, sheetValidationListOptions } from '../src/features/sheet-validation/validation-list-input.ts'

// Unexecuted name/dynamic-source contracts for the shared scalar/reference engine. No second formula evaluator is used.
const area = (row, column, endRow = row, endCol = column) => ({ startRow: row, startCol: column, endRow, endCol })
function sample() {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'validation-source'; model.activeSheetId = sheet.id = 'data'; sheet.name = 'Data'; model.calculationMode = 'manual'
  sheet.rowCount = 100; sheet.columnCount = 20
  for (const [row, value] of [[0, 'Red'], [1, 'Green, dark'], [2, 'Blue'], [3, 'Red']]) sheet.cells[`${row}:3`] = createCell(row, 3, value)
  sheet.cells['0:4'] = createCell(0, 4, true); sheet.cells['1:4'] = createCell(1, 4, 2); sheet.cells['0:5'] = createCell(0, 5, 'Data!$D$1:$D$2')
  const other = createBlankSheet("Other's data"); other.id = 'other'; other.cells['0:0'] = createCell(0, 0, 'other'); model.sheets.push(other)
  model.metadata.definedNames = [{ name: 'Choices', refersTo: '=Data!$D$1:$D$2' }, { name: '选择项', refersTo: '=Data!$D$1:$D$2' }]
  const options = (formula, extra = {}) => sheetDataValidationOptions({ id: 'rule', type: 'list', ranges: [area(1, 1)], formula1: formula, ...extra },
    model, sheet.id, sheetValidationSourceReader('local-user'))
  return { model, sheet, other, options }
}

test.each([
  ['=$D$1:$D$2', ['Red', 'Green, dark']], ['=Choices', ['Red', 'Green, dark']], ['=选择项', ['Red', 'Green, dark']],
  ['=IF($E$1,$D$1:$D$2,$D$3:$D$4)', ['Red', 'Green, dark']], ['=CHOOSE($E$2,$D$1:$D$2,$D$3:$D$4)', ['Blue', 'Red']],
  ['=OFFSET($D$1,0,0,COUNTA($D$1:$D$4),1)', ['Red', 'Green, dark', 'Blue']],
  ['=INDIRECT($F$1)', ['Red', 'Green, dark']], ['=INDIRECT("R[-1]C[2]:R[0]C[2]",FALSE)', ['Red', 'Green, dark']],
  ['=INDEX($D$1:$D$4,0,1)', ['Red', 'Green, dark', 'Blue']], ['=Data!D:D', ['Red', 'Green, dark', 'Blue']],
])('%s resolves live row-major candidates from one shared reference result', (formula, expected) => {
  const h = sample(), before = cloneModel(h.model)
  assert.deepEqual(h.options(formula), expected); assert.deepEqual(h.model, before)
})

test('name scope, explicitly qualified names and quoted sheet names keep their canonical source', () => {
  const h = sample(); h.model.metadata.definedNames.push({ name: 'Choices', refersTo: '=Data!$D$3:$D$4', localSheetId: h.sheet.id },
    { name: 'Choices', refersTo: "='Other''s data'!$A$1", localSheetId: h.other.id })
  assert.deepEqual(h.options('=Choices'), ['Blue', 'Red'])
  assert.deepEqual(h.options("='Other''s data'!Choices"), ['other'])
  assert.deepEqual(h.options("='Other''s data'!$A$1"), ['other'])
})

test('a dynamic range follows source/control edits and keeps current candidate spelling and order', () => {
  const h = sample(), formula = '=IF($E$1,$D$1:$D$2,$D$3:$D$4)'
  assert.deepEqual(h.options(formula), ['Red', 'Green, dark'])
  h.sheet.cells['0:4'].value = false; assert.deepEqual(h.options(formula), ['Blue', 'Red'])
  h.sheet.cells['2:3'].value = 'Approved'; assert.deepEqual(h.options(formula), ['Approved', 'Red'])
  h.sheet.cells['0:5'].value = 'Data!$D$3:$D$4'; assert.deepEqual(h.options('=INDIRECT($F$1)'), ['Approved', 'Red'])
})

test('OFFSET can use a nested count formula without being rejected as an unsupported dropdown function', () => {
  const formula = '=OFFSET($D$1,0,0,COUNTA($D$1:$D$4),1)', rule = createSheetDataValidationRule({ id: 'offset', type: 'list', formula1: formula }, [area(1, 1)])
  assert.equal(rule.formula1, formula)
  assert.throws(() => createSheetDataValidationRule({ type: 'list', formula1: '=SUM(D1:D4)' }, [area(1, 1)]), /仅支持/)
})

test.each(['list', 'listMultiple'])('a pure INDEX %s source can be authored through panel settings and follows reference edits without changing scalar data', type => {
  const h = sample(), settings = validationPanelSettings(undefined, type), formula = '=INDEX($D$1:$D$4,0,1)', before = cloneModel(h.model)
  settings.listMode = 'reference'; settings.listReference = formula; settings.formula2 = '#ff0000,#008000,#0000ff'
  const rule = validationRuleFromPanel(settings, [area(1, 1)], { id: 'index-source', metadata: { retained: ['authoring'] } })
  assert.equal(rule.formula1, formula); assert.equal(rule.type, type); assert.equal(rule.formula2, settings.formula2)
  assert.deepEqual(sheetDataValidationOptions(rule, h.model, h.sheet.id, sheetValidationSourceReader('local-user')), ['Red', 'Green, dark', 'Blue'])
  assert.deepEqual(h.model, before); const reopened = validationPanelSettings(rule)
  assert.equal(reopened.listMode, 'reference'); assert.equal(reopened.listReference, formula)
  assert.deepEqual(validationRuleFromPanel(reopened, rule.ranges, {}, rule), rule)
  h.sheet.cells['2:3'].value = 'Approved'; assert.deepEqual(h.options(formula), ['Red', 'Green, dark', 'Approved'])
  h.sheet.rangeProtectionRules = [{ id: 'private-index', ranges: [area(2, 3)], ownerUserId: 'peer', allowedUserIds: [],
    viewState: 'noOneElseCanView', editState: 'onlyMe' }]
  assert.deepEqual(sheetDataValidationOptions(rule, h.model, h.sheet.id, sheetValidationSourceReader('local-user')), [])
})

test('relative R1C1 uses the saved formula anchor rather than the current selected cell', () => {
  const h = sample(), formula = '=INDIRECT("R[-1]C[2]:R[0]C[2]",FALSE)'
  assert.deepEqual(h.options(formula, { ranges: [area(50, 10)], formulaAnchor: { row: 1, col: 1 } }), ['Red', 'Green, dark'])
})

test('hidden unused branches remain unread while a hidden selector or resolved source yields no candidates', () => {
  const h = sample(), formula = '=IF($E$1,$D$1:$D$2,$D$3:$D$4)'
  h.sheet.rangeProtectionRules = [{ id: 'private', ranges: [area(2, 3, 3, 3)], ownerUserId: 'peer', allowedUserIds: [],
    viewState: 'noOneElseCanView', editState: 'onlyMe' }]
  assert.deepEqual(h.options(formula), ['Red', 'Green, dark'])
  h.sheet.cells['0:4'].value = false; assert.deepEqual(h.options(formula), [])
  h.sheet.cells['0:4'].value = true; h.sheet.rangeProtectionRules[0].ranges = [area(0, 4)]
  assert.deepEqual(h.options(formula), [])
})

test('whole-axis sources respect a hidden blank location instead of checking populated cells alone', () => {
  const h = sample(); h.sheet.rangeProtectionRules = [{ id: 'blank', ranges: [area(50, 3)], ownerUserId: 'peer', allowedUserIds: [],
    viewState: 'noOneElseCanView', editState: 'onlyMe' }]
  assert.deepEqual(h.options('=D:D'), []); assert.deepEqual(h.options('=D1:D4'), ['Red', 'Green, dark', 'Blue'])
})

test('nested source-cell formulas honor View and omit formula errors without leaking private values', () => {
  const h = sample(); h.other.cells['0:0'] = createCell(0, 0, 'private-token')
  h.other.rangeProtectionRules = [{ id: 'source', ranges: [area(0, 0)], ownerUserId: 'peer', allowedUserIds: [],
    viewState: 'noOneElseCanView', editState: 'onlyMe' }]
  h.sheet.cells['1:3'].formula = "='Other''s data'!A1"; h.sheet.cells['1:3'].value = 'private-token'
  assert.deepEqual(h.options('=D1:D2'), ['Red'])
  h.other.rangeProtectionRules = []; assert.deepEqual(h.options('=D1:D2'), ['Red', 'private-token'])
})

test.each(['=1+2', '=INDIRECT("1+2")', '=OFFSET(D1,-1,0)', '=OFFSET(D1,0,0,0,1)', '=MissingName', '=Circular'])('invalid reference %s yields an empty dropdown without mutating its rule', formula => {
  const h = sample(); h.model.metadata.definedNames.push({ name: 'Circular', refersTo: '=Circular' })
  const before = cloneModel(h.model); assert.deepEqual(h.options(formula), []); assert.deepEqual(h.model, before)
})

test.each([
  ['zh-CN', '1234.50'], ['de-DE', '1234,50'], ['fr-FR', '1234,50'],
])('referenced %s candidates format evaluated numeric scalars and retain original strings, formula caches and deduplicated colors', (region, decimal) => {
  const h = sample(); h.sheet.cells = {}
  for (const [row, value, pattern] of [[0, 0.25, '0%'], [1, 1234.5, '0.00'], [2, '0.25', '0%'], [3, 0.251, '0%']]) {
    const cell = createCell(row, 3, value); cell.style.numberFormat = pattern; h.sheet.cells[`${row}:3`] = cell
  }
  h.sheet.cells['0:4'] = createCell(0, 4, 1); h.sheet.cells['0:3'].formula = '=E1/4'; h.sheet.cells['0:3'].value = 0.9
  const rule = createSheetDataValidationRule({ id: 'formatted', type: 'list', formula1: '=D1:D4', formula2: '#ff0000,#008000,#0000ff', renderMode: 'custom' }, [area(1, 1)])
  const before = cloneModel(h.model)
  assert.deepEqual(sheetDataValidationOptions(rule, h.model, h.sheet.id, sheetValidationSourceReader('local-user'), region), ['25%', decimal, '0.25'])
  assert.deepEqual(sheetValidationListOptions(rule, h.model, h.sheet.id, 'local-user', region), [
    { label: '25%', value: '25%', color: '#ff0000' }, { label: decimal, value: decimal, color: '#008000' }, { label: '0.25', value: '0.25', color: '#0000ff' },
  ])
  assert.deepEqual(h.model, before); assert.equal(h.sheet.cells['0:3'].value, 0.9); assert.equal(h.sheet.cells['0:3'].formula, '=E1/4')
})

test.each([
  [false, 60, '1900-02-29'], [true, 0, '1904-01-01'],
])('referenced date1904=%s labels preserve the workbook epoch and exceptional day instead of the source display cache', (date1904, serial, expected) => {
  const h = sample(); h.model.date1904 = date1904; h.sheet.cells = { '0:3': createCell(0, 3, serial) }
  h.sheet.cells['0:3'].style.numberFormat = 'yyyy-mm-dd'; h.sheet.cells['0:3'].display = 'stale imported display'
  const rule = createSheetDataValidationRule({ type: 'list', formula1: '=D1' }, [area(1, 1)]), before = cloneModel(h.model)
  assert.deepEqual(sheetDataValidationOptions(rule, h.model, h.sheet.id), [expected]); assert.deepEqual(h.model, before)
})

test.each(['list', 'listMultiple'])('formatted %s choices validate their stored labels and typed widget plans preserve the numeric source', type => {
  const h = sample(); h.sheet.cells = { '0:3': createCell(0, 3, 0.25), '1:1': createCell(1, 1, 0.25) }
  h.sheet.cells['0:3'].style.numberFormat = '0%'; h.sheet.cells['1:1'].style.numberFormat = '0%'
  const rule = createSheetDataValidationRule({ id: 'formatted', type, formula1: '=D1', allowBlank: false }, [area(1, 1)])
  h.sheet.dataValidations = [rule]; const before = cloneModel(h.model), reader = sheetValidationSourceReader('local-user')
  assert.equal(validateSheetDataValidationChanges(h.model, h.sheet.id, [{ row: 1, column: 1 }], reader).blocked, true)
  const plan = planSheetValidationListInput(h.model, { unitId: h.model.id, sheetId: h.sheet.id, row: 1, column: 1, ruleId: rule.id, values: ['25%'] },
    { region: 'en-US', currentYear: 2026, userId: 'local-user' })
  assert.equal(plan.result.accepted, true); assert.equal(plan.cell.value, '25%'); assert.equal(plan.cell.dataType, 'text'); assert.deepEqual(h.model, before)
  h.sheet.cells['1:1'] = plan.cell
  assert.equal(validateSheetDataValidationChanges(h.model, h.sheet.id, [{ row: 1, column: 1 }], reader, 'en-US').valid, true)
  assert.equal(h.sheet.cells['0:3'].value, 0.25); assert.equal(h.sheet.cells['0:3'].dataType, 'number')
})

test('source format inheritance is resolved live, while a hidden source rejects both numeric and style getters before candidate extraction', () => {
  const h = sample(); h.sheet.cells = { '0:3': createCell(0, 3, 0.25) }; h.sheet.cells['0:3'].authoredStyle = {}
  h.sheet.columnStyles = { 3: { numberFormat: '0%' } }
  assert.deepEqual(h.options('=D1'), ['25%']); h.sheet.columnStyles[3].numberFormat = '0.00'; assert.deepEqual(h.options('=D1'), ['0.25'])
  Object.defineProperty(h.sheet.cells['0:3'], 'value', { get() { throw Error('private numeric read') } })
  Object.defineProperty(h.sheet.cells['0:3'], 'style', { get() { throw Error('private style read') } })
  h.sheet.rangeProtectionRules = [{ id: 'private-format', ranges: [area(0, 3)], ownerUserId: 'peer', allowedUserIds: [],
    viewState: 'noOneElseCanView', editState: 'onlyMe' }]
  assert.deepEqual(h.options('=D1'), [])
})
