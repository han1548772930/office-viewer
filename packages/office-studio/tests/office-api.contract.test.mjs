import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let apiModule

before(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

after(async () => {
  await server?.close()
})

test('public FUniver API creates units, edits ranges, and tracks selection', () => {
  const studio = new apiModule.OfficeStudio()
  const univer = apiModule.FUniver.newAPI(studio)
  const created = []
  const subscription = univer.addEvent(univer.Event.WorkbookCreated, (event) => created.push(event.workbookId))

  const workbook = univer.createWorkbook({ name: 'Contract workbook' })
  const worksheet = workbook.getActiveSheet()
  assert.ok(worksheet)
  assert.equal(workbook.getName(), 'Contract workbook')
  assert.equal(univer.getActiveWorkbook()?.getId(), workbook.getId())
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
  const univer = apiModule.FUniver.newAPI(studio)
  const workbook = univer.getActiveWorkbook()
  assert.ok(workbook)
  const worksheet = workbook.getActiveSheet()
  assert.ok(worksheet)

  const range = worksheet.getRange('A1')
  range.setValue('before history')
  workbook.undo()
  assert.equal(range.getValue(), null)
  workbook.redo()
  assert.equal(range.getValue(), 'before history')

  const cancellation = univer.addEvent(univer.Event.BeforeCommandExecute, (event) => {
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
  const univer = apiModule.FUniver.newAPI(studio)
  const created = []
  const subscription = univer.addEvent(univer.Event.DocCreated, (event) => created.push(event.documentId))

  const document = univer.createDocument({ name: 'Contract document' })
  assert.equal(univer.getActiveDocument()?.getId(), document.getId())
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
  const api = apiModule.FUniver.newAPI(studio)
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
  const api = apiModule.FUniver.newAPI(studio)
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
  const univer = apiModule.FUniver.newAPI(studio)
  try {
    const workbook = univer.getActiveWorkbook()
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
    const reopened = univer.createWorkbook({ ...snapshot, id: 'validated-reopened' })
    const reopenedSheet = reopened.getActiveSheet()
    assert.equal(reopenedSheet.getDataValidation('A1').id, rule.id)
    assert.throws(() => reopenedSheet.getRange('A1').setValue(100), /Enter an integer from 1 to 10/)
    reopenedSheet.setDataValidation('A1:A3', null)
    assert.equal(reopenedSheet.getDataValidations().length, 1)
    reopenedSheet.setDataValidation('B1', null)
    assert.equal(reopenedSheet.getDataValidations().length, 0)
  } finally { studio.dispose() }
})
