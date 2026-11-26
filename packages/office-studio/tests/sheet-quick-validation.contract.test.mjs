import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let apiModule
let quickValidation
let validationEngine

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'packages/office-studio/vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  quickValidation = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-validation/quick-validation.ts')
  validationEngine = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-validation/validation-engine.ts')
})

afterAll(async () => {
  await server?.close()
})

test('quick checkbox, dropdown and date presets create rules over all selected ranges', () => {
  const ranges = [
    { startRow: 0, startCol: 0, endRow: 2, endCol: 0 },
    { startRow: 4, startCol: 1, endRow: 4, endCol: 2 },
  ]
  const checkbox = quickValidation.createQuickSheetDataValidationRule('checkbox', ranges)
  assert.equal(checkbox.type, 'checkbox')
  assert.deepEqual(checkbox.ranges, ranges)

  const dropdown = quickValidation.createQuickSheetDataValidationRule('dropdown', ranges, {
    dropdownValues: [' Red ', 'Green, dark', 'It"s blue', ''],
  })
  assert.equal(dropdown.type, 'list')
  assert.equal(dropdown.formula1, '"Red,"Green, dark","It""s blue""')
  assert.equal(dropdown.renderMode, 'custom')

  const date = quickValidation.createQuickSheetDataValidationRule('date', ranges)
  assert.equal(date.type, 'date')
  assert.equal(date.operator, undefined)
  assert.equal(date.formula1, undefined)
  const dateTime = quickValidation.createQuickSheetDataValidationRule('date-time', ranges)
  assert.deepEqual(dateTime.metadata, { showTime: true })
  assert.throws(() => quickValidation.createQuickSheetDataValidationRule('dropdown', ranges), /至少需要一个选项/)
})

test('quick dropdown items containing separators remain intact in real editing, undo and reopening', () => {
  const studio = new apiModule.OfficeStudio()
  try {
    const office = apiModule.FOffice.newAPI(studio)
    const book = office.createWorkbook({ id: 'quick-list-special-items', calculationMode: 'manual' })
    const sheet = book.getActiveSheet()
    const values = ['Red', 'Green, dark', 'It"s blue', 'Line\nbreak']
    const ranges = [
      { startRow: 1, endRow: 1, startCol: 1, endCol: 1 },
      { startRow: 3, endRow: 3, startCol: 1, endCol: 1 },
    ]
    const rule = quickValidation.createQuickSheetDataValidationRule('dropdown', ranges, { dropdownValues: values })
    office.executeCommandSync('sheet.command.set-data-validation', { unitId: book.getId(), sheetId: sheet.getSheetId(), rule })
    assert.equal(studio.editor.history.units[book.getId()].undo.length, 1)
    assert.deepEqual(sheet.getCellValidationOptions(1, 1), values)
    assert.deepEqual(sheet.getCellValidationOptions(3, 1), values)
    assert.deepEqual(sheet.getCellValidationOptions(2, 1), [])
    const beforeInvalid = JSON.stringify(book.save())
    assert.throws(() => sheet.getRange('B2').setValue('Green'))
    assert.equal(JSON.stringify(book.save()), beforeInvalid)
    sheet.getRange('B2').setValue('Green, dark')
    assert.equal(sheet.getRange('B2').getValue(), 'Green, dark')
    book.undo()
    assert.equal(sheet.getRange('B2').getValue(), null)
    assert.deepEqual(sheet.getCellValidationOptions(1, 1), values)
    book.redo()
    const reopened = new apiModule.OfficeStudio({ workbook: book.save() })
    try {
      const reopenedSheet = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet()
      assert.equal(reopenedSheet.getRange('B2').getValue(), 'Green, dark')
      assert.deepEqual(reopenedSheet.getCellValidationOptions(1, 1), values)
      assert.deepEqual(reopenedSheet.getCellValidationOptions(3, 1), values)
    } finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('quick actions target every selected rectangle, localize list contents and open only explicit edits', () => {
  const ranges = [{ startRow: 0, endRow: 1, startCol: 1, endCol: 2 }, { startRow: 6, endRow: 6, startCol: 0, endCol: 0 }]
  const label = (key, fallback) => ({ 'sheet.validation.preset.yes': 'Yes', 'sheet.validation.preset.no': 'No' })[key] ?? fallback
  const plan = action => quickValidation.planQuickValidation(action, ranges, label)
  const defaultRule = plan('add-default')
  assert.deepEqual({ type: defaultRule.rule.type, operator: defaultRule.rule.operator, formula1: defaultRule.rule.formula1 },
    { type: 'decimal', operator: 'equal', formula1: '100' })
  assert.deepEqual(defaultRule.rule.ranges, ranges)
  assert.equal(defaultRule.edit, true)
  assert.deepEqual(plan('clear-dropdown'), { kind: 'clear', ranges, types: ['list', 'listMultiple'] })
  assert.deepEqual(plan('dropdown-yes-no').rule.ranges, ranges)
  assert.equal(plan('dropdown-yes-no').rule.formula1, '"Yes,No"')
  assert.equal(plan('dropdown-status').edit, false)
  assert.equal(plan('edit-dropdown').edit, true)
  assert.equal(plan('edit-date').rule.operator, undefined)
  assert.equal(plan('clear-other'), undefined)
  assert.equal(quickValidation.planQuickValidation('clear-date', [], label), undefined)
})

for (const formulas of [{ formula1: 'TRUE', formula2: 'FALSE' }, { formula1: '1', formula2: '0' }, { formula1: 'DONE' }, { formula2: 'NO' }]) {
  test(`checkbox ${JSON.stringify(formulas)} accepts real boolean input and each fallback candidate`, () => {
    const studio = new apiModule.OfficeStudio()
    try {
      const office = apiModule.FOffice.newAPI(studio)
      const book = office.createWorkbook({ id: 'checkbox-input', calculationMode: 'manual' })
      const sheet = book.getActiveSheet()
      sheet.setDataValidation('A1', { type: 'checkbox', ...formulas })
      const options = [formulas.formula1 ?? 'TRUE', formulas.formula2 ?? 'FALSE']
      assert.deepEqual(sheet.getCellValidationOptions(0, 0), options)
      studio.editor.history.clear(book.getId())
      const target = { unitId: book.getId(), sheetId: sheet.getSheetId(), row: 0, column: 0 }
      const checkedInput = formulas.formula1 === 'DONE' ? 'DONE' : 'TRUE'
      const checked = office.executeCommandSync('sheet.command.commit-cell-input', { ...target, input: checkedInput })
      assert.equal(checked.accepted, true)
      assert.equal(sheet.getRange('A1').getValue(), checkedInput === 'DONE' ? 'DONE' : true)
      assert.equal(studio.editor.history.units[book.getId()].undo.length, 1)
      const beforeInvalid = JSON.stringify(book.save())
      const rejected = office.executeCommandSync('sheet.command.commit-cell-input', { ...target, input: formulas.formula1 === 'DONE' ? 'done' : 'unexpected' })
      assert.equal(rejected.accepted, false)
      assert.equal(JSON.stringify(book.save()), beforeInvalid)
      book.undo()
      assert.equal(sheet.getRange('A1').getValue(), null)
      assert.deepEqual(sheet.getCellValidationOptions(0, 0), options)
      book.redo()
      const uncheckedInput = formulas.formula2 === 'NO' ? 'NO' : 'FALSE'
      assert.equal(office.executeCommandSync('sheet.command.commit-cell-input', { ...target, input: uncheckedInput }).accepted, true)
      const reopened = new apiModule.OfficeStudio({ workbook: book.save() })
      try {
        const reopenedSheet = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet()
        assert.equal(reopenedSheet.getRange('A1').getValue(), uncheckedInput === 'NO' ? 'NO' : false)
        assert.deepEqual(reopenedSheet.getCellValidationOptions(0, 0), options)
      } finally { reopened.dispose() }
    } finally { studio.dispose() }
  })
}

for (const source of ['inline', 'reference']) {
  test(`multi-select ${source} keeps quoted labels intact through real input and candidate changes`, () => {
    const studio = new apiModule.OfficeStudio()
    try {
      const office = apiModule.FOffice.newAPI(studio)
      const book = office.createWorkbook({ id: 'multi-list-input', calculationMode: 'manual' })
      const sheet = book.getActiveSheet()
      const values = ['Red', 'Green, dark', 'It"s blue', '"Quoted"']
      sheet.getRange('A1:A4').setValues(values.map(value => [value]))
      sheet.setDataValidation('B1', { type: 'listMultiple', formula1: source === 'inline' ? validationEngine.serializeValidationListItems(values) : '=$A$1:$A$4' })
      assert.deepEqual(sheet.getCellValidationOptions(0, 1), values)
      const target = { unitId: book.getId(), sheetId: sheet.getSheetId(), row: 0, column: 1 }
      for (const selected of [[values[1]], [values[2]], [values[3]], values, ['red', 'green, DARK', 'it"s BLUE']]) {
        const csv = validationEngine.serializeValidationListItems(selected).slice(1, -1)
        assert.equal(office.executeCommandSync('sheet.command.commit-cell-input', { ...target, input: csv }).accepted, true)
        assert.equal(sheet.getRange('B1').getValue(), csv)
        assert.deepEqual(sheet.getDataValidationErrors('B1'), [])
      }
      studio.editor.history.clear(book.getId())
      const previousValue = sheet.getRange('B1').getValue()
      const nextValues = ['Red', 'It"s blue', '"Quoted"', 'Approved']
      sheet.setDataValidation('B1', { type: 'listMultiple', formula1: validationEngine.serializeValidationListItems(nextValues) })
      assert.deepEqual(sheet.getCellValidationOptions(0, 1), nextValues)
      assert.equal(sheet.getDataValidationErrors('B1').length, 1)
      const beforeInvalid = JSON.stringify(book.save())
      const rejected = office.executeCommandSync('sheet.command.commit-cell-input', {
        ...target, input: validationEngine.serializeValidationListItems(['Green, dark']).slice(1, -1),
      })
      assert.equal(rejected.accepted, false)
      assert.equal(JSON.stringify(book.save()), beforeInvalid)
      const nextValue = validationEngine.serializeValidationListItems(['It"s blue', 'Approved']).slice(1, -1)
      assert.equal(office.executeCommandSync('sheet.command.commit-cell-input', { ...target, input: nextValue }).accepted, true)
      assert.equal(studio.editor.history.units[book.getId()].undo.length, 2)
      book.undo()
      assert.equal(sheet.getRange('B1').getValue(), previousValue)
      book.undo()
      assert.deepEqual(sheet.getCellValidationOptions(0, 1), values)
      assert.deepEqual(sheet.getDataValidationErrors('B1'), [])
      book.redo()
      book.redo()
      const reopened = new apiModule.OfficeStudio({ workbook: book.save() })
      try {
        const reopenedSheet = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet()
        assert.equal(reopenedSheet.getRange('B1').getValue(), nextValue)
        assert.deepEqual(reopenedSheet.getCellValidationOptions(0, 1), nextValues)
        assert.deepEqual(reopenedSheet.getDataValidationErrors('B1'), [])
      } finally { reopened.dispose() }
    } finally { studio.dispose() }
  })
}

test('quick validation commands persist, undo and clear only matching rule types', () => {
  const studio = new apiModule.OfficeStudio()
  const office = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = office.getActiveWorkbook()
    const sheet = workbook.getActiveSheet()
    const checkboxRange = { startRow: 0, startCol: 0, endRow: 1, endCol: 0 }
    const dropdownRange = { startRow: 0, startCol: 1, endRow: 1, endCol: 1 }
    const dateRange = { startRow: 3, startCol: 2, endRow: 3, endCol: 2 }
    const checkbox = quickValidation.createQuickSheetDataValidationRule('checkbox', [checkboxRange])
    office.executeCommandSync('sheet.command.set-data-validation', { unitId: workbook.getId(), sheetId: sheet.getSheetId(), rule: checkbox })
    const dropdown = quickValidation.createQuickSheetDataValidationRule('dropdown', [dropdownRange], { dropdownValues: ['Open', 'Closed'] })
    office.executeCommandSync('sheet.command.set-data-validation', { unitId: workbook.getId(), sheetId: sheet.getSheetId(), rule: dropdown })
    const date = quickValidation.createQuickSheetDataValidationRule('date', [dateRange])
    office.executeCommandSync('sheet.command.set-data-validation', { unitId: workbook.getId(), sheetId: sheet.getSheetId(), rule: date })

    assert.deepEqual(sheet.getDataValidations().map(rule => rule.type), ['checkbox', 'list', 'date'])
    assert.equal(sheet.getDataValidation('A1').type, 'checkbox')
    assert.equal(sheet.getDataValidation('B1').type, 'list')
    assert.equal(sheet.getDataValidation('C4').type, 'date')
    sheet.getRange('C4').setValue('2026-11-21')
    assert.equal(sheet.getDataValidationErrors('C4').length, 0)
    assert.throws(() => sheet.getRange('C4').setValue('not a date'), /符合日期条件/)
    assert.equal(sheet.getDataValidationErrors('C4').length, 0)
    assert.equal(sheet.getRange('C4').getValue(), '2026-11-21')

    const beforeClear = sheet.getDataValidations()
    office.executeCommandSync('sheet.command.clear-quick-data-validation', {
      unitId: workbook.getId(), sheetId: sheet.getSheetId(), ranges: [{ ...checkboxRange, endRow: 0 }], types: ['checkbox'],
    })
    assert.equal(sheet.getDataValidation('A1'), null)
    assert.equal(sheet.getDataValidation('A2').type, 'checkbox')
    assert.equal(sheet.getDataValidation('B1').type, 'list')
    workbook.undo()
    assert.deepEqual(sheet.getDataValidations(), beforeClear)
    office.executeCommandSync('sheet.command.clear-quick-data-validation', {
      unitId: workbook.getId(), sheetId: sheet.getSheetId(), ranges: [checkboxRange, dateRange], types: ['checkbox', 'date'],
    })
    assert.deepEqual(sheet.getDataValidations().map(rule => rule.type), ['list'])
    workbook.undo()
    assert.deepEqual(sheet.getDataValidations(), beforeClear)
    workbook.redo()
    assert.deepEqual(sheet.getDataValidations().map(rule => rule.type), ['list'])
  } finally {
    studio.dispose()
  }
})

test('quick rule changes and type clears reject protected selected ranges atomically', () => {
  const studio = new apiModule.OfficeStudio({ currentUserId: 'guest' })
  const office = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = office.getActiveWorkbook()
    const sheet = workbook.getActiveSheet()
    const rules = [quickValidation.createQuickSheetDataValidationRule('checkbox', [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }])]
    sheet.getRange('A1').setDataValidation({ type: 'whole', operator: 'greaterThan', formula1: '0' })
    const model = studio.getWorkbookModel(workbook.getId())
    model.sheets.find(item => item.id === sheet.getSheetId()).rangeProtectionRules = [{
      id: 'protected-a1', permissionId: 'permission-a1', ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
      ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView',
    }]
    const before = sheet.getDataValidations()
    assert.throws(() => office.executeCommandSync('sheet.command.set-data-validation', {
      unitId: workbook.getId(), sheetId: sheet.getSheetId(), rule: rules[0],
    }), /受保护/)
    assert.throws(() => office.executeCommandSync('sheet.command.clear-quick-data-validation', {
      unitId: workbook.getId(), sheetId: sheet.getSheetId(), ranges: rules[0].ranges, types: ['checkbox'],
    }), /受保护/)
    assert.deepEqual(sheet.getDataValidations(), before)
  } finally {
    studio.dispose()
  }
})
