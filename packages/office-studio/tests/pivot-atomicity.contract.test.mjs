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

function setup() {
  const studio = new apiModule.OfficeStudio()
  const officeApi = apiModule.FOffice.newAPI(studio)
  const workbook = officeApi.getActiveWorkbook()
  const source = workbook.getActiveSheet()
  const target = workbook.insertSheet('Pivot')
  source.getRange('A1:C3').setValues([
    ['Region', 'Quarter', 'Sales'],
    ['East', 'Q1', 100],
    ['West', 'Q1', 80],
  ])
  return { studio, workbook, source, target }
}

function sourceInfo(workbook, source) {
  return {
    unitId: workbook.getId(),
    subUnitId: source.getSheetId(),
    sheetName: source.getSheetName(),
    range: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
  }
}

function anchorInfo(workbook, target, row = 0, col = 0) {
  return { unitId: workbook.getId(), subUnitId: target.getSheetId(), row, col }
}

test('preflights an occupied anchor before registering an empty pivot', async () => {
  const { studio, workbook, source, target } = setup()
  target.getRange('A1').setValue('occupied')

  await assert.rejects(
    () => workbook.addPivotTable(sourceInfo(workbook, source), apiModule.PositionTypeEnum.Existing, anchorInfo(workbook, target)),
    /overlaps existing content/,
  )
  assert.equal(workbook.getPivotTables().length, 0)
  assert.equal(studio.getWorkbookModel(workbook.getId()).pivotTables?.length ?? 0, 0)

  studio.dispose()
})

test('does not commit a failed field addition after an empty pivot anchor becomes occupied', async () => {
  const { studio, workbook, source, target } = setup()
  const pivot = await workbook.addPivotTable(
    sourceInfo(workbook, source),
    apiModule.PositionTypeEnum.Existing,
    anchorInfo(workbook, target),
  )
  const before = studio.getWorkbookModel(workbook.getId()).pivotTables[0]
  assert.deepEqual(before.fields, [])
  assert.equal(before.revision, 0)
  assert.equal(before.outputRange, undefined)

  target.getRange('A1').setValue('occupied after creation')
  await assert.rejects(() => pivot.addField(2, 'value'), /overlaps existing content/)

  const after = studio.getWorkbookModel(workbook.getId()).pivotTables[0]
  assert.deepEqual(after.fields, [])
  assert.equal(after.revision, before.revision)
  assert.equal(after.outputRange, undefined)
  assert.deepEqual(after.generatedCellKeys, [])
  assert.deepEqual(pivot.getData(), [])
  assert.equal(target.getRange('A1').getValue(), 'occupied after creation')

  // The rejected command must not add a history item: the first undo clears
  // the cell written before the failed field addition.
  workbook.undo()
  assert.equal(target.getRange('A1').getValue(), null)
  assert.equal(workbook.getPivotTables().length, 1)

  studio.dispose()
})
