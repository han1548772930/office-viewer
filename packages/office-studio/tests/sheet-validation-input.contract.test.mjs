import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createSheetCellEditSession } from '../src/features/sheet-workbench/sheet-cell-edit-session.ts'
import { DateSystem } from '../src/features/sheet-workbench/number-format-engine/date-system.ts'
import { excelDateTimePartsToSerial, excelSerialToDateTimeParts, getDateSystemMaxSerial } from '../src/features/sheet-workbench/number-format-engine/serial-date.ts'
import { COMMIT_SHEET_DATE_INPUT_COMMAND, planSheetDateInput, resolveSheetDateInputChange, sheetDateInputCommandInput,
  sheetDateInputState } from '../src/features/sheet-validation/validation-date-input.ts'
import { COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, sheetValidationListInput, sheetValidationListOptions,
  sheetValidationListSelected, sheetValidationListValue } from '../src/features/sheet-validation/validation-list-input.ts'
import { serializeValidationListItems, validateSheetDataValidationChanges } from '../src/features/sheet-validation/validation-engine.ts'

// Implementation-phase drafts. Tests, the SSR harness and Office samples have not been run.
let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = (row = 0, column = 0, endRow = row, endCol = column) => ({ startRow: row, startCol: column, endRow, endCol })
const clock = (hours, minutes = 0, seconds = 0, fractionalSecond = 0) => ({ kind: 'time', hours, minutes, seconds, fractionalSecond })
const date = (year, month, day) => ({ kind: 'date', year, month, day })
function fixture(configure = () => {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'validation-input'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'
  sheet.rowCount = 80; sheet.columnCount = 20; sheet.cells['0:0'] = createCell(0, 0, 60.125)
  sheet.dataValidations = [{ id: 'rule', type: 'date', ranges: [area()], metadata: { retained: ['data'] } }]
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other); configure(model, sheet)
  const studio = new OfficeStudio({ workbook: model, currentUserId: 'local-user' }), book = FOffice.newAPI(studio).getActiveWorkbook()
  const current = () => studio.getWorkbookModel('validation-input')
  const target = (withSession = true) => {
    const model = current(), sheet = model.sheets.find(item => item.id === 'data'), view = toExcelWorkbookView(model)
    const session = createSheetCellEditSession({ documentId: model.id, sheetId: sheet.id, key: 'A1', date1904: model.date1904,
      cell: view.sheets.find(item => item.id === sheet.id).cells.A1 }, { region: 'en-US', currentYear: 2026 })
    return { unitId: model.id, sheetId: sheet.id, row: 0, column: 0, ruleId: 'rule', ...(withSession ? { session } : {}),
      expectedRevision: studio.getWorkbookRevision(model.id), expectedHistoryGeneration: studio.editor.history.captureSavepoint(model.id).generation,
      expectedUserId: studio.currentUserId, expectedRuleSignature: JSON.stringify(sheet.dataValidations[0]) }
  }
  return { studio, book, current, target, cell: () => current().sheets.find(sheet => sheet.id === 'data').cells['0:0'],
    count: () => studio.editor.history.units['validation-input']?.undo.length ?? 0, dispose: () => studio.dispose() }
}

test.each([
  { serial: 0, year: 1900, month: 1, day: 0 }, { serial: 59, year: 1900, month: 2, day: 28 },
  { serial: 60, year: 1900, month: 2, day: 29 }, { serial: 61, year: 1900, month: 3, day: 1 },
])('calendar anchor $serial stays exact on an untouched widget and a clock-only edit', item => {
  const original = item.serial + 0.125, state = sheetDateInputState({ type: 'date', metadata: { showTime: true } }, original, 'yyyy-mm-dd hh:mm:ss', false)
  assert.deepEqual([state.parts.year, state.parts.month, state.parts.day], [item.year, item.month, item.day])
  assert.deepEqual(resolveSheetDateInputChange(state, { kind: 'unchanged' }), { changed: false, serial: original })
  assert.deepEqual(resolveSheetDateInputChange(state, clock(18)), { changed: true, serial: item.serial + 0.75 })
  const same = { kind: 'date-time', ...state.parts }
  assert.equal(resolveSheetDateInputChange(state, same).changed, false)
})

test.each([false, true])('workbook date1904=%s uses its epoch and rejects selections before that epoch', date1904 => {
  const h = fixture(model => { model.date1904 = date1904 })
  try {
    const epoch = date1904 ? date(1904, 1, 1) : date(1900, 1, 1), result = h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: epoch })
    assert.equal(result.accepted, true); assert.equal(h.cell().value, (date1904 ? 0 : 1) + 0.125); assert.equal(h.count(), 1)
    const before = cloneModel(h.current())
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: date(date1904 ? 1903 : 1899, 12, 31) }), /范围/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 1)
    const serial1900 = excelDateTimePartsToSerial({ year: 1904, month: 1, day: 1, hours: 0, minutes: 0, seconds: 0, fractionalSecond: 0 }, { dateSystem: DateSystem.Date1900 })
    assert.equal(serial1900, 1462)
  } finally { h.dispose() }
})

test('a pure date change keeps the exact hidden time and an existing date format in one real content transaction', () => {
  const hidden = 12345.123456789 / 86400, h = fixture((_model, sheet) => {
    sheet.cells['0:0'].value = 59 + hidden; sheet.cells['0:0'].style.numberFormat = 'dd/mm/yyyy'
  })
  try {
    const original = h.cell().value, first = h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: date(2024, 2, 29) })
    const expectedDay = excelDateTimePartsToSerial({ year: 2024, month: 2, day: 29, hours: 0, minutes: 0, seconds: 0, fractionalSecond: 0 }, { dateSystem: DateSystem.Date1900 })
    assert.equal(first.accepted, true); assert.equal(h.cell().value, expectedDay + (original - Math.floor(original)))
    assert.equal(h.cell().style.numberFormat, 'dd/mm/yyyy'); assert.equal(h.cell().dataType, 'number'); assert.equal(h.count(), 1)
    const saved = cloneModel(h.current()), again = h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: date(2024, 2, 29) })
    assert.equal(again.changed, false); assert.deepEqual(h.current(), saved); assert.equal(h.count(), 1)
    h.book.undo(); assert.equal(h.cell().value, original); h.book.redo(); assert.deepEqual(h.current().sheets[0].cells, saved.sheets[0].cells)
  } finally { h.dispose() }
})

test.each([60, '60', 0.123456789012345])('untouched value %s keeps its original scalar, type and format without history', original => {
  const h = fixture((_model, sheet) => { sheet.cells['0:0'].value = original; sheet.cells['0:0'].dataType = typeof original === 'string' ? 'text' : 'number'; sheet.cells['0:0'].style.numberFormat = 'General' })
  try {
    const before = cloneModel(h.current()), result = h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: { kind: 'unchanged' } })
    assert.deepEqual(result, { accepted: true, changed: false, issues: [] }); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('an out-of-calendar serial can be corrected numerically, while negative, nonfinite and overflowing candidates never publish', () => {
  const maximum = getDateSystemMaxSerial(DateSystem.Date1900), h = fixture((_model, sheet) => { sheet.cells['0:0'].value = maximum + 20 })
  try {
    assert.equal(sheetDateInputState(h.current().sheets[0].dataValidations[0], h.cell().value, 'General', false).parts, null)
    for (const serial of [-1, maximum + 1, Number.POSITIVE_INFINITY, Number.NaN]) {
      const before = cloneModel(h.current())
      assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: { kind: 'serial', serial } }))
      assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    }
    assert.equal(h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: { kind: 'serial', serial: 60.5 } }).accepted, true)
    assert.equal(h.cell().value, 60.5); assert.equal(h.cell().style.numberFormat, 'yyyy-MM-dd'); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test.each([0, 60, 45292])('a clock rule preserves day %s and an elapsed-time rule stores more than a day', day => {
  const h = fixture((_model, sheet) => { sheet.cells['0:0'].value = day + 0.125; sheet.cells['0:0'].style.numberFormat = 'hh:mm:ss'; sheet.dataValidations[0].type = 'time' })
  try {
    const changed = h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: clock(12, 34, 56, 0.25) })
    assert.equal(changed.accepted, true); assert.equal(h.cell().value, day + (12 * 3600 + 34 * 60 + 56.25) / 86400)
    assert.equal(h.cell().style.numberFormat, 'hh:mm:ss'); assert.equal(h.count(), 1)
    h.studio.commitWorkbook('validation-input', 'elapsed format', model => { model.sheets[0].cells['0:0'].style.numberFormat = '[h]:mm:ss'; model.sheets[0].cells['0:0'].authoredStyle = undefined }, { preserveStyleInheritance: true })
    assert.equal(h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: { kind: 'serial', serial: 2.5 } }).accepted, true)
    assert.equal(h.cell().value, 2.5); assert.equal(h.cell().style.numberFormat, '[h]:mm:ss')
  } finally { h.dispose() }
})

test('typed replacement clears a formula and rich body while retaining unrelated metadata, formatting, and per-workbook undo', () => {
  const h = fixture((_model, sheet) => {
    const cell = sheet.cells['0:0']; cell.formula = '=60'; cell.dataType = 'formula'; cell.cachedDataType = 'number'; cell.richText = [{ text: '60' }]
    cell.metadata = { retained: { stamp: 7 }, formulaError: '#REF!', publicBody: { dataStream: '60\r\n' } }; cell.style.font.bold = true
    sheet.dataValidations[0].metadata.showTime = true
  })
  try {
    const before = cloneModel(h.cell()), change = { kind: 'date-time', year: 2026, month: 11, day: 28, hours: 12, minutes: 3, seconds: 4, fractionalSecond: 0.125 }
    const result = h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change })
    assert.equal(result.accepted, true); assert.equal(h.cell().formula, null); assert.equal(h.cell().richText, undefined); assert.equal(h.cell().cachedDataType, null)
    assert.deepEqual(h.cell().metadata.retained, { stamp: 7 }); assert.equal(h.cell().metadata.publicBody, undefined); assert.equal(h.cell().metadata.formulaError, undefined)
    assert.equal(h.cell().style.font.bold, true); assert.equal(h.cell().style.numberFormat, 'yyyy-MM-dd hh:mm:ss'); assert.equal(h.count(), 1)
    const after = cloneModel(h.cell()), saved = h.book.save(), other = FOffice.newAPI(h.studio).createWorkbook({ id: 'independent-input' })
    other.getActiveSheet().getRange('A1').setValue('peer'); h.book.undo(); assert.deepEqual(h.cell(), before)
    assert.equal(other.getActiveSheet().getRange('A1').getValue(), 'peer'); h.book.redo(); assert.deepEqual(h.cell(), after)
    const reopened = new OfficeStudio({ workbook: saved })
    try { assert.equal(reopened.getWorkbookModel(saved.id).sheets[0].cells['0:0'].value, after.value); assert.equal(reopened.getWorkbookModel(saved.id).sheets[0].dataValidations[0].metadata.showTime, true) }
    finally { reopened.dispose() }
  } finally { h.dispose() }
})

test.each(['stop', 'warning', 'info'])('date validation severity %s produces a real rejection or issue-bearing accepted transaction', severity => {
  const h = fixture((_model, sheet) => { Object.assign(sheet.dataValidations[0], { operator: 'equal', formula1: '60', errorStyle: severity, errorTitle: 'Date rule', error: 'Keep day 60' }) })
  try {
    const before = cloneModel(h.current()), result = h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: date(1900, 3, 1) })
    assert.equal(result.issues.length, 1); assert.equal(result.issues[0].severity, severity)
    assert.equal(result.accepted, severity !== 'stop'); assert.equal(h.count(), severity === 'stop' ? 0 : 1)
    if (severity === 'stop') { assert.deepEqual(h.current(), before); assert.match(result.message, /Date rule.*Keep day 60/) }
  } finally { h.dispose() }
})

test.each(['date-system', 'number-format', 'content', 'rich-text'])('the shared edit session rejects a changed %s without publishing the widget draft', conflict => {
  const h = fixture()
  try {
    const target = h.target()
    h.studio.commitWorkbook('validation-input', 'external edit', model => {
      if (conflict === 'date-system') model.date1904 = !model.date1904
      else if (conflict === 'number-format') { model.sheets[0].cells['0:0'].style.numberFormat = '0.00'; model.sheets[0].cells['0:0'].authoredStyle = undefined }
      else if (conflict === 'content') model.sheets[0].cells['0:0'].value = 61
      else model.sheets[0].cells['0:0'].richText = [{ text: 'external run' }]
    }, { preserveStyleInheritance: true })
    delete target.expectedRevision; const before = cloneModel(h.current()), count = h.count()
    const result = h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...target, change: date(2026, 11, 28) })
    assert.equal(result.accepted, false); assert.equal(result.conflict, conflict); assert.deepEqual(h.current(), before); assert.equal(h.count(), count)
  } finally { h.dispose() }
})

test('merged widgets address the main cell and check view/edit rights over every cell in that merge, including no-op input', () => {
  const h = fixture((_model, sheet) => { sheet.mergedRanges = [area(0, 0, 1, 1)]; sheet.dataValidations[0].ranges = [area(0, 0, 1, 1)] })
  try {
    const target = { ...h.target(), row: 1, column: 1 }, result = h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...target, change: date(1900, 3, 1) })
    assert.equal(result.accepted, true); assert.equal(h.cell().value, 61.125); assert.equal(h.current().sheets[0].cells['1:1'], undefined)
    h.current().sheets[0].rangeProtectionRules = [{ id: 'hidden-tail', ranges: [area(1, 1)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'othersCanEdit' }]
    const before = cloneModel(h.current())
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(false), row: 1, column: 1, change: { kind: 'unchanged' } }), /查看/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test.each([false, true])('single/multiple=%s keeps numeric-looking and formula-looking labels as text, in candidate order', multiple => {
  const labels = ['1', '=1+1', 'Green, dark', 'It"s blue', '"Quoted"', 'Line\nbreak']
  const h = fixture((_model, sheet) => { sheet.cells['0:0'].value = null; Object.assign(sheet.dataValidations[0], { type: multiple ? 'listMultiple' : 'list', formula1: serializeValidationListItems(labels) }) })
  try {
    const requested = multiple ? [labels[4], labels[1], labels[0], '1'] : [labels[1]], expected = multiple ? [labels[0], labels[1], labels[4]] : [labels[1]]
    const result = h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { ...h.target(), values: requested })
    assert.equal(result.accepted, true); assert.equal(h.cell().value, sheetValidationListValue(expected, multiple)); assert.equal(h.cell().dataType, 'text'); assert.equal(h.cell().formula, null)
    assert.equal(h.count(), 1); assert.deepEqual(validateSheetDataValidationChanges(h.current(), 'data', [{ row: 0, column: 0 }]).issues, [])
    const options = sheetValidationListOptions(h.current().sheets[0].dataValidations[0], h.current(), 'data')
    assert.deepEqual(sheetValidationListSelected(options, h.cell().value, multiple), expected)
    const saved = h.book.save(), reopened = new OfficeStudio({ workbook: saved })
    try { assert.equal(reopened.getWorkbookModel(saved.id).sheets[0].cells['0:0'].value, h.cell().value) } finally { reopened.dispose() }
  } finally { h.dispose() }
})

test('multiple CSV selection round-trips separators, escaped quotes and line breaks without a second parser', () => {
  const labels = ['Green, dark', 'It"s blue', '"Quoted"', 'Line\nbreak'], options = labels.map(label => ({ value: label, label }))
  for (const selected of labels.map(label => [label]).concat([labels])) {
    const value = sheetValidationListValue(selected, true)
    assert.deepEqual(sheetValidationListSelected(options, value, true), selected)
  }
  assert.deepEqual(sheetValidationListSelected(options, '"green, DARK",it""s BLUE', true), labels.slice(0, 2))
})

test('candidate colors stay aligned after case-fold deduplication and reference edits provide new live candidates', () => {
  const h = fixture((_model, sheet) => {
    sheet.cells['0:0'].value = null; sheet.cells['0:3'] = createCell(0, 3, 'Red'); sheet.cells['1:3'] = createCell(1, 3, 'Green, dark')
    Object.assign(sheet.dataValidations[0], { type: 'listMultiple', formula1: '=$D$1:$D$2' })
  })
  try {
    assert.equal(h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { ...h.target(), values: ['green, DARK', 'red'] }).accepted, true)
    assert.equal(h.cell().value, 'Red,"Green, dark"')
    h.book.getActiveSheet().getRange('D2').setValue('Approved')
    const before = cloneModel(h.current()), count = h.count()
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { ...h.target(), values: ['Green, dark'] }), /候选/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), count)
    assert.equal(h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { ...h.target(), values: ['approved', 'RED'] }).accepted, true)
    assert.equal(h.cell().value, 'Red,Approved')
    const rule = { type: 'listMultiple', formula1: serializeValidationListItems(['Red', 'red', 'Green']), formula2: '#abc,#ff0000,#123456' }
    assert.deepEqual(sheetValidationListOptions(rule, h.current(), 'data'), [
      { value: 'Red', label: 'Red', color: '#aabbcc' }, { value: 'Green', label: 'Green', color: '#123456' },
    ])
  } finally { h.dispose() }
})

test('clearing a required list rejects both an existing selection and an already invalid blank without history', () => {
  const h = fixture((_model, sheet) => { sheet.cells['0:0'].value = 'Red'; Object.assign(sheet.dataValidations[0], { type: 'listMultiple', formula1: '"Red,Green"', allowBlank: false }) })
  try {
    const before = cloneModel(h.current()), first = h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { ...h.target(), values: [] })
    assert.equal(first.accepted, false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    h.cell().value = ''; const blank = cloneModel(h.current()), again = h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { ...h.target(), values: [] })
    assert.equal(again.accepted, false); assert.equal(again.reason, 'validation'); assert.deepEqual(h.current(), blank); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('a protected reference source cannot disclose options or be consumed by a typed list choice', () => {
  const h = fixture((_model, sheet) => {
    sheet.cells['0:0'].value = null; sheet.cells['0:3'] = createCell(0, 3, 'Secret')
    Object.assign(sheet.dataValidations[0], { type: 'list', formula1: '=$D$1:$D$1' })
    sheet.rangeProtectionRules = [{ id: 'private-source', ranges: [area(0, 3)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'onlyMe' }]
  })
  try {
    assert.deepEqual(sheetValidationListOptions(h.current().sheets[0].dataValidations[0], h.current(), 'data', 'local-user'), [])
    const before = cloneModel(h.current())
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { ...h.target(), values: ['Secret'] }), /候选/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('runtime readers copy frozen input and read each date-change field once before planning an immutable candidate', () => {
  const h = fixture(), reads = {}
  try {
    const fields = { kind: 'date-time', year: 2026, month: 11, day: 28, hours: 12, minutes: 0, seconds: 0, fractionalSecond: 0.25 }, change = {}
    for (const [key, value] of Object.entries(fields)) Object.defineProperty(change, key, { get() { reads[key] = (reads[key] ?? 0) + 1; return value }, enumerable: true })
    const input = sheetDateInputCommandInput(Object.freeze({ ...h.target(), change })), before = cloneModel(h.current())
    assert.deepEqual(reads, Object.fromEntries(Object.keys(fields).map(key => [key, 1])))
    assert.equal(planSheetDateInput(h.current(), input, { region: 'en-US', currentYear: 2026, userId: 'local-user' }).result.accepted, true)
    assert.deepEqual(h.current(), before)
    const values = Object.freeze(['Red']), parsed = sheetValidationListInput(Object.freeze({ ...h.target(), values }))
    parsed.values.push('Green'); assert.deepEqual(values, ['Red'])
    assert.throws(() => sheetDateInputCommandInput({ ...h.target(), change: { ...fields, hours: 25 } }), /时间/)
    assert.throws(() => sheetDateInputCommandInput({ ...h.target(), change: { ...fields, fractionalSecond: null } }), /字段/)
  } finally { h.dispose() }
})

test.each(['revision', 'history', 'user', 'rule', 'host'])('stale %s context refuses an API widget write, including a no-op', field => {
  const h = fixture()
  try {
    const input = { ...h.target(), change: { kind: 'unchanged' } }
    if (field === 'revision') input.expectedRevision += 1
    else if (field === 'history') input.expectedHistoryGeneration += 1
    else if (field === 'user') input.expectedUserId = 'other'
    else if (field === 'rule') input.expectedRuleSignature = '{}'
    else { input.host = 'workbench'; h.book.getSheetBySheetId('other').activate() }
    const before = cloneModel(h.current())
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, input), /更改|规则/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('an inactive API workbook remains independently editable and command cancellation keeps content/history untouched', () => {
  const h = fixture()
  try {
    const other = FOffice.newAPI(h.studio).createWorkbook({ id: 'active-peer' }), beforePeer = cloneModel(h.studio.getWorkbookModel(other.getId()))
    const cancel = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === COMMIT_SHEET_DATE_INPUT_COMMAND) event.cancel = true })
    const before = cloneModel(h.current())
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), change: date(1900, 3, 1) }), /canceled/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0); cancel.dispose()
    assert.equal(h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...h.target(), host: 'api', change: date(1900, 3, 1) }).accepted, true)
    assert.equal(h.count(), 1); assert.deepEqual(h.studio.getWorkbookModel(other.getId()), beforePeer)
  } finally { h.dispose() }
})

test.each(['list', 'listMultiple'])('formatted %s selection stores the regional label as text, emits one value change and reopens without changing numeric sources', type => {
  const h = fixture((_model, sheet) => {
    sheet.cells['0:0'] = createCell(0, 0, '25%'); sheet.cells['0:3'] = createCell(0, 3, 0.25); sheet.cells['0:3'].style.numberFormat = '0%'
    sheet.cells['1:3'] = createCell(1, 3, 1234.5); sheet.cells['1:3'].style.numberFormat = '0.00'
    sheet.dataValidations = [{ id: 'rule', type, ranges: [area()], formula1: '=INDEX($D$1:$D$2,0,1)', formula2: '#ff0000,#008000', renderMode: 'custom' }]
  })
  const events = [], subscription = h.studio.events.addEvent('SheetValueChanged', event => events.push(cloneModel(event)))
  try {
    h.studio.regionalSettings.setRegion('de-DE')
    const before = cloneModel(h.current()), options = sheetValidationListOptions(h.current().sheets[0].dataValidations[0], h.current(), 'data', 'local-user', 'de-DE')
    assert.deepEqual(options.map(option => option.value), ['25%', '1234,50']); assert.deepEqual(h.book.getActiveSheet().getCellValidationOptions(0, 0), ['25%', '1234,50'])
    const result = h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { ...h.target(false), region: 'de-DE', values: ['1234,50'] })
    assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(h.cell().value, type === 'listMultiple' ? '"1234,50"' : '1234,50')
    assert.equal(h.cell().dataType, 'text'); assert.equal(h.count(), 1); assert.equal(events.length, 1)
    assert.deepEqual(h.current().sheets[0].cells['0:3'], before.sheets[0].cells['0:3']); assert.deepEqual(h.current().sheets[0].cells['1:3'], before.sheets[0].cells['1:3'])
    const after = cloneModel(h.current()), unchanged = h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { ...h.target(false), region: 'de-DE', values: ['1234,50'] })
    assert.equal(unchanged.accepted, true); assert.equal(unchanged.changed, false); assert.equal(h.count(), 1); assert.equal(events.length, 1)
    h.book.undo(); assert.deepEqual(h.current(), before); h.book.redo(); assert.deepEqual(h.current(), after)
    const reopened = new OfficeStudio({ workbook: h.book.save(), region: 'de-DE' })
    try {
      const restored = FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet()
      assert.equal(restored.getRange('A1').getValue(), h.cell().value); assert.deepEqual(restored.getCellValidationOptions(0, 0), ['25%', '1234,50'])
      assert.equal(reopened.getWorkbookModel(after.id).sheets[0].cells['1:3'].value, 1234.5)
    } finally { reopened.dispose() }
  } finally { subscription.dispose(); h.dispose() }
})

test.each(['planning', 'mutator', 'prepareView'])('view permission revoked during %s prevents the planned cell from being published', stage => {
  const h = fixture()
  try {
    const target = h.target(), revoke = () => { h.current().sheets[0].rangeProtectionRules = [{ id: 'reentrant-lock', ranges: [area()], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'othersCanEdit' }] }
    if (stage === 'planning') {
      const original = h.studio.assertCanViewRange.bind(h.studio); let calls = 0
      h.studio.assertCanViewRange = (...args) => { original(...args); if (++calls === 1) revoke() }
    } else {
      const commit = h.studio.commitWorkbook.bind(h.studio)
      h.studio.commitWorkbook = (id, label, update, options) => commit(id, label, model => { if (stage === 'mutator') revoke(); update(model) },
        { ...options, prepareView: model => { if (stage === 'prepareView') revoke(); options.prepareView(model) } })
    }
    const originalValue = h.cell().value
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_DATE_INPUT_COMMAND, { ...target, change: date(1900, 3, 1) }), /权限|查看|更改/)
    assert.equal(h.cell().value, originalValue); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})
