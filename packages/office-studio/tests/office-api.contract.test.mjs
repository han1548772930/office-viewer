import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let apiModule

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

afterAll(async () => {
  await server?.close()
})

test('public FOffice API creates units, edits ranges, and tracks selection', () => {
  const studio = new apiModule.OfficeStudio()
  const officeApi = apiModule.FOffice.newAPI(studio)
  const created = []
  const subscription = officeApi.addEvent(officeApi.Event.WorkbookCreated, (event) => created.push(event.workbookId))

  const workbook = officeApi.createWorkbook({ name: 'Contract workbook' })
  const worksheet = workbook.getActiveSheet()
  assert.ok(worksheet)
  assert.equal(workbook.getName(), 'Contract workbook')
  assert.equal(officeApi.getActiveWorkbook()?.getId(), workbook.getId())
  assert.deepEqual(created, [workbook.getId()])

  const range = worksheet.getRange('A1:B2')
  assert.equal(range.getA1Notation(), 'A1:B2')
  range.setValues([[1, 'two'], [true, null]])
  assert.deepEqual(range.getValues(), [[1, 'two'], [true, null]])

  range.activate()
  assert.deepEqual(worksheet.getActiveRange()?.getRange(), {
    startRow: 0,
    startColumn: 0,
    endRow: 1,
    endColumn: 1,
  })
  assert.equal(studio.getSelection()?.unitId, workbook.getId())

  subscription.dispose()
  studio.dispose()
})

test('public API exposes cancellation and per-unit undo/redo', () => {
  const studio = new apiModule.OfficeStudio()
  const officeApi = apiModule.FOffice.newAPI(studio)
  const workbook = officeApi.getActiveWorkbook()
  assert.ok(workbook)
  const worksheet = workbook.getActiveSheet()
  assert.ok(worksheet)

  const range = worksheet.getRange('A1')
  range.setValue('before history')
  workbook.undo()
  assert.equal(range.getValue(), null)
  workbook.redo()
  assert.equal(range.getValue(), 'before history')

  const cancellation = officeApi.addEvent(officeApi.Event.BeforeCommandExecute, (event) => {
    if (event.id === 'sheet.command.set-workbook-name') event.cancel = true
  })
  assert.throws(() => workbook.setName('blocked'), /canceled/)
  cancellation.dispose()
  workbook.setName('renamed')
  assert.equal(workbook.getName(), 'renamed')

  studio.dispose()
})

test('public API creates and edits documents through the same facade', () => {
  const studio = new apiModule.OfficeStudio()
  const officeApi = apiModule.FOffice.newAPI(studio)
  const created = []
  const subscription = officeApi.addEvent(officeApi.Event.DocCreated, (event) => created.push(event.documentId))

  const document = officeApi.createDocument({ name: 'Contract document' })
  assert.equal(officeApi.getActiveDocument()?.getId(), document.getId())
  assert.equal(document.getName(), 'Contract document')
  assert.deepEqual(created, [document.getId()])

  document.insertText(0, 'Hello')
  assert.equal(document.getText(), 'Hello')
  document.undo()
  assert.equal(document.getText(), '')
  document.redo()
  assert.equal(document.getText(), 'Hello')
  assert.equal(document.getTextRange(1, 4).getText(), 'ell')

  subscription.dispose()
  studio.dispose()
})

test('partial workbook inputs retain their date system through formulas, history and reopening', () => {
  const studio = new apiModule.OfficeStudio({ workbook: { id: 'calendar-1904', date1904: true } })
  const api = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = api.getActiveWorkbook()
    assert.equal(workbook.getId(), 'calendar-1904')
    assert.equal(studio.getWorkbookModel(workbook.getId()).date1904, true)
    const cell = workbook.getActiveSheet().getRange('A1')
    cell.setValue('=DATE(1904,1,2)')
    assert.equal(cell.getValue(), 1)
    workbook.undo()
    assert.equal(cell.getValue(), null)
    workbook.redo()
    assert.equal(cell.getValue(), 1)
    const snapshot = JSON.parse(JSON.stringify(workbook.save()))
    assert.equal(snapshot.date1904, true)
    const reopened = api.createWorkbook({ ...snapshot, id: 'calendar-reopened' })
    assert.equal(studio.getWorkbookModel(reopened.getId()).date1904, true)
    reopened.getActiveSheet().getRange('A2').setValue('=DATE(1904,1,2)')
    assert.equal(reopened.getActiveSheet().getRange('A2').getValue(), 1)
    assert.equal(cell.getValue(), 1)
  } finally { studio.dispose() }
})

test('new workbook facades keep calendar settings isolated and default to the 1900 system', () => {
  const studio = new apiModule.OfficeStudio()
  const api = apiModule.FOffice.newAPI(studio)
  try {
    const original = api.getActiveWorkbook()
    const alternate = api.createWorkbook({ name: '1904 calendar', date1904: true })
    const standard = api.createWorkbook({ name: '1900 calendar', date1904: false })
    for (const [workbook, expected] of [[alternate, 1], [standard, 1463], [original, 1463]]) {
      const cell = workbook.getActiveSheet().getRange('A1')
      cell.setValue('=DATE(1904,1,2)')
      assert.equal(cell.getValue(), expected)
      assert.equal(workbook.save().date1904, expected === 1)
    }
  } finally { studio.dispose() }
})

test('worksheet data validation rejects API writes atomically and persists through history and snapshots', () => {
  const studio = new apiModule.OfficeStudio()
  const officeApi = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = officeApi.getActiveWorkbook()
    const sheet = workbook.getActiveSheet()
    const range = sheet.getRange('A1:A3')
    sheet.setDataValidation('A1:A3', {
      type: 'whole', operator: 'between', formula1: '1', formula2: '10', allowBlank: false,
      errorStyle: 'stop', errorTitle: 'Invalid quantity', error: 'Enter an integer from 1 to 10.',
    })
    const rule = range.getDataValidation()
    assert.equal(rule.type, 'whole')
    assert.equal(rule.ranges[0].endRow, 2)

    range.setValues([[2], [4], [8]])
    const beforeRejectedWrite = workbook.save()
    assert.throws(() => range.setValues([[2], [11], [8]]), /Enter an integer from 1 to 10/)
    assert.deepEqual(workbook.save().sheets[0].cells, beforeRejectedWrite.sheets[0].cells)

    workbook.undo()
    assert.deepEqual(range.getValues(), [[null], [null], [null]])
    workbook.redo()
    assert.deepEqual(range.getValues(), [[2], [4], [8]])

    sheet.setDataValidation('B1', { type: 'whole', operator: 'greaterThan', formula1: '10', errorStyle: 'warning' })
    sheet.getRange('B1').setValue(2)
    assert.equal(sheet.getDataValidationErrors('B1')[0].type, 'whole')
    const snapshot = JSON.parse(JSON.stringify(workbook.save()))
    assert.equal(snapshot.sheets[0].dataValidations.find(item => item.id === rule.id).formula1, '1')
    const reopened = officeApi.createWorkbook({ ...snapshot, id: 'validated-reopened' })
    const reopenedSheet = reopened.getActiveSheet()
    assert.equal(reopenedSheet.getDataValidation('A1').id, rule.id)
    assert.throws(() => reopenedSheet.getRange('A1').setValue(100), /Enter an integer from 1 to 10/)
    reopenedSheet.setDataValidation('A1:A3', null)
    assert.equal(reopenedSheet.getDataValidations().length, 1)
    reopenedSheet.setDataValidation('B1', null)
    assert.equal(reopenedSheet.getDataValidations().length, 0)
  } finally { studio.dispose() }
})

test('typed API cells and formula replacement keep content types through history and reopening', () => {
  const studio = new apiModule.OfficeStudio()
  const api = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = api.getActiveWorkbook()
    const sheet = workbook.getActiveSheet()
    const range = sheet.getRange('A1:D1')
    range.setValues([[{ v: '00123', t: 1 }, { v: '00123', t: 4 }, { v: '12.25', t: 2 }, { v: '0', t: 3 }]])
    assert.deepEqual(range.getValues(), [['00123', '00123', 12.25, false]])
    assert.deepEqual(range.getCellDatas()[0].map(cell => cell.dataType), ['text', 'force_string', 'number', 'boolean'])
    const error = sheet.getRange('A2')
    error.setFormula('=1/0')
    assert.equal(error.getCellData().cachedDataType, 'error')
    const previous = error.getCellData()
    error.setValue({ v: '00123', t: 1 })
    assert.equal(error.getValue(), '00123')
    assert.equal(error.getCellData().formula, null)
    assert.equal(error.getCellData().dataType, 'text')
    assert.equal(error.getCellData().cachedDataType, null)
    assert.equal(error.getCellData().metadata?.formulaError, undefined)
    workbook.undo()
    assert.deepEqual(error.getCellData(), previous)
    workbook.redo()
    const reopened = api.createWorkbook({ ...JSON.parse(JSON.stringify(workbook.save())), id: 'typed-api-reopened' })
    assert.deepEqual(reopened.getActiveSheet().getRange('A1:D1').getCellDatas(), range.getCellDatas())
    assert.deepEqual(reopened.getActiveSheet().getRange('A2').getCellData(), error.getCellData())
  } finally { studio.dispose() }
})

test('public number formats preserve scalars and caches in a single history step, including no-ops and reopen', () => {
  const studio = new apiModule.OfficeStudio({ workbook: { id: 'doc-prefixed-format-workbook', calculationMode: 'manual' } })
  const api = apiModule.FOffice.newAPI(studio)
  const executed = []
  const subscription = api.addEvent(api.Event.CommandExecuted, event => executed.push(event))
  try {
    const workbook = api.getActiveWorkbook()
    const sheet = workbook.getActiveSheet()
    const range = sheet.getRange('A1:E1')
    range.setValues([[12.25, { v: '00123', t: 1 }, { f: '=A1/100', v: .1225, t: 2 }, { v: '0', t: 3 }, { f: '=UNSUPPORTED()', v: null }]])
    const before = range.getCellDatas()
    const content = cells => cells[0].map(({ value, formula, dataType, cachedDataType }) => ({ value, formula, dataType, cachedDataType }))
    const depth = () => ({ undo: studio.editor.history.units[workbook.getId()].undo.length, redo: studio.editor.history.units[workbook.getId()].redo.length })
    const patterns = [['0.0', '0.0%', '0.00%', 'General', '0.0']]
    range.setNumberFormats(patterns)
    assert.deepEqual(content(range.getCellDatas()), content(before))
    assert.deepEqual(range.getCellDatas()[0].map(cell => cell.display), ['12.3', '00123', '12.25%', 'FALSE', '=UNSUPPORTED()'])
    assert.deepEqual(depth(), { undo: 2, redo: 0 })
    assert.deepEqual(executed.filter(event => event.id === 'sheet.command.numfmt.set.numfmt').map(({ kind, unitId }) => ({ kind, unitId })), [{ kind: 'workbook', unitId: workbook.getId() }])
    range.setNumberFormats(patterns)
    assert.deepEqual(depth(), { undo: 2, redo: 0 })
    workbook.undo()
    assert.deepEqual(range.getCellDatas(), before)
    assert.deepEqual(depth(), { undo: 1, redo: 1 })
    workbook.redo()
    const after = range.getCellDatas()
    const reopened = api.createWorkbook({ ...JSON.parse(JSON.stringify(workbook.save())), id: 'format-api-reopened' })
    assert.deepEqual(reopened.getActiveSheet().getRange('A1:E1').getCellDatas(), after)
    assert.deepEqual(range.getCellDatas(), after)
  } finally { subscription.dispose(); studio.dispose() }
})

test('number format command cancellation, invalid batches and denied targets leave both model and history intact', () => {
  const studio = new apiModule.OfficeStudio({ currentUserId: 'guest', workbook: {
    id: 'atomic-format-workbook', sheets: [{ id: 'atomic-sheet', name: 'Atomic formats', cellData: { 0: { 0: { v: 12 }, 1: { v: 25 } } },
      rangeProtectionRules: [{ id: 'format-lock', permissionId: 'format-permission', ownerUserId: 'owner', allowedUserIds: [],
        editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [{ startRow: 0, endRow: 0, startCol: 1, endCol: 1 }] }],
    }],
  } })
  const api = apiModule.FOffice.newAPI(studio)
  let cancellation
  try {
    const workbook = api.getActiveWorkbook()
    const sheet = workbook.getActiveSheet()
    const before = JSON.parse(JSON.stringify(studio.getWorkbookModel(workbook.getId())))
    const revision = studio.getWorkbookRevision(workbook.getId())
    const unchanged = () => {
      assert.deepEqual(JSON.parse(JSON.stringify(studio.getWorkbookModel(workbook.getId()))), before)
      assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
      assert.equal(studio.editor.history.units[workbook.getId()].undo.length, 0)
    }
    cancellation = api.addEvent(api.Event.BeforeCommandExecute, event => { if (event.id === 'sheet.command.set-number-format') event.cancel = true })
    assert.throws(() => sheet.getRange('A1').setNumberFormat('0.00'), /canceled/)
    unchanged()
    cancellation.dispose(); cancellation = undefined
    assert.throws(() => sheet.getRange('A1:B1').setNumberFormat('0.00'), /保护/)
    unchanged()
    assert.throws(() => sheet.getRange('A2:B2').setNumberFormats([['0.00', '0"unfinished']]))
    unchanged()
    assert.throws(() => sheet.getRange('A2:B2').setValues([[1, Infinity]]), TypeError)
    unchanged()
    assert.throws(() => api.executeCommandSync('sheet.command.numfmt.set.numfmt', { unitId: workbook.getId(), subUnitId: sheet.getSheetId(), values: [
      { row: 1, col: 0, pattern: '0.00' }, { row: 1, col: 16_384, pattern: '0.00' },
    ] }), RangeError)
    unchanged()
    assert.throws(() => api.executeCommandSync('sheet.command.set-number-format', { unitId: workbook.getId(), sheetId: sheet.getSheetId(),
      ranges: [{ startRow: 1, endRow: 1, startCol: 0, endCol: 1 }], pattern: '0"a\nb"' }))
    unchanged()
  } finally { cancellation?.dispose(); studio.dispose() }
})

test('text error labels, quoted formulas and real errors stay distinct through API edits, history and reopen', () => {
  const studio = new apiModule.OfficeStudio()
  const api = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = api.getActiveWorkbook()
    const sheet = workbook.getActiveSheet()
    sheet.getRange('A1:E1').setValues([[{ v: '#DIV/0!', t: 1 }, { v: '#N/A', t: 4 }, { f: '="#DIV/0!"' }, 2, { f: '=1/0' }]])
    const results = sheet.getRange('A2:E2')
    results.setValues([['=A1', '=ISTEXT(A1)', '=ISERROR(A1)', '=SUM(A1:D1)', '=IFERROR(E1,"caught")']])
    assert.deepEqual(results.getValues(), [['#DIV/0!', true, false, 2, 'caught']])
    assert.equal(sheet.getRange('C1').getCellData().cachedDataType, 'string')
    assert.equal(sheet.getRange('C1').getCellData().metadata?.formulaError, undefined)
    assert.equal(sheet.getRange('E1').getCellData().cachedDataType, 'error')
    sheet.getRange('A1').setValue(7)
    assert.deepEqual(results.getValues(), [[7, false, false, 9, 'caught']])
    workbook.undo()
    assert.deepEqual(results.getValues(), [['#DIV/0!', true, false, 2, 'caught']])
    workbook.redo()
    const reopened = api.createWorkbook({ ...JSON.parse(JSON.stringify(workbook.save())), id: 'error-text-reopened' })
    assert.deepEqual(reopened.getActiveSheet().getRange('A1:E2').getCellDatas(), sheet.getRange('A1:E2').getCellDatas())
  } finally { studio.dispose() }
})

test('formatting preserves imported formula caches in automatic and manual calculation modes', () => {
  for (const calculationMode of ['automatic', 'manual']) {
    const studio = new apiModule.OfficeStudio({ workbook: { id: `cache-format-${calculationMode}`, calculationMode: 'manual' } })
    const api = apiModule.FOffice.newAPI(studio)
    try {
      const workbook = api.getActiveWorkbook()
      const sheet = workbook.getActiveSheet()
      const range = sheet.getRange('A1:F1')
      range.setValues([[12.25, { v: '00123', t: 4 }, { f: '=A1/100', v: .5, t: 2 },
        { f: '=UNSUPPORTED()', v: null }, { f: '=UNSUPPORTED()', v: '#DIV/0!', t: 1 }, { f: '=UNSUPPORTED()', v: '#DIV/0!', t: 'e' }]])
      const initial = JSON.parse(JSON.stringify(studio.getWorkbookModel(workbook.getId())))
      initial.calculationMode = calculationMode
      initial.sheets[0].cells['0:3'].display = 'source 25%'
      studio.editor.workbook.replace(initial)
      studio.editor.history.clear(workbook.getId())
      const before = range.getCellDatas()
      const content = cells => cells[0].map(({ value, formula, dataType, cachedDataType, metadata }) => ({ value, formula, dataType, cachedDataType, metadata }))
      range.setNumberFormat('0.000')
      assert.deepEqual(content(range.getCellDatas()), content(before), calculationMode)
      assert.deepEqual(range.getCellDatas()[0].map(cell => cell.display), ['12.250', '00123', '0.500', 'source 25%', '#DIV/0!', '#DIV/0!'])
      const revision = studio.getWorkbookRevision(workbook.getId())
      range.setNumberFormat('0.000')
      assert.equal(studio.editor.history.units[workbook.getId()].undo.length, 1)
      assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
      workbook.undo()
      assert.deepEqual(range.getCellDatas(), before)
      workbook.redo()
      const after = range.getCellDatas()
      const reopened = api.createWorkbook({ ...JSON.parse(JSON.stringify(workbook.save())), id: `cache-reopened-${calculationMode}` })
      assert.deepEqual(reopened.getActiveSheet().getRange('A1:F1').getCellDatas(), after)
    } finally { studio.dispose() }
  }
})
