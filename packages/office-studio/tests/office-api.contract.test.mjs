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
