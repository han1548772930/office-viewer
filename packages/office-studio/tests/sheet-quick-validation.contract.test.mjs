import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let apiModule
let quickValidation

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
  assert.equal(dropdown.formula1, '"Red,Green, dark,It""s blue"')
  assert.equal(dropdown.renderMode, 'custom')

  const date = quickValidation.createQuickSheetDataValidationRule('date', ranges)
  assert.equal(date.type, 'date')
  assert.equal(date.operator, undefined)
  assert.equal(date.formula1, undefined)
  assert.throws(() => quickValidation.createQuickSheetDataValidationRule('dropdown', ranges), /至少需要一个选项/)
})

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
