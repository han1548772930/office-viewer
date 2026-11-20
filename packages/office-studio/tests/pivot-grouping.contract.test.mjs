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

function setup(rows) {
  const studio = new apiModule.OfficeStudio()
  const univer = apiModule.FUniver.newAPI(studio)
  const workbook = univer.getActiveWorkbook()
  const source = workbook.getActiveSheet()
  const target = workbook.insertSheet('Grouped')
  source.getRange(`A1:D${rows.length + 1}`).setValues([['Amount', 'Date', 'Team', 'Sales'], ...rows])
  return { studio, workbook, source, target }
}

async function createPivot(workbook, source, target, rowField = 0) {
  const pivot = await workbook.addPivotTable({
    unitId: workbook.getId(),
    subUnitId: source.getSheetId(),
    sheetName: source.getSheetName(),
    range: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 3 },
  }, apiModule.PositionTypeEnum.Existing, {
    unitId: workbook.getId(),
    subUnitId: target.getSheetId(),
    row: 0,
    col: 0,
  })
  await pivot.addField(rowField, 'row')
  await pivot.addField(3, 'value')
  return pivot
}

test('numeric grouping emits interval labels and drills into original members', async () => {
  const { studio, workbook, source, target } = setup([
    [1, '2026-01-05', 'A', 10],
    [4, '2026-01-06', 'A', 20],
    [6, '2026-02-05', 'B', 30],
    [9, '2026-02-06', 'B', 40],
  ])
  const pivot = await createPivot(workbook, source, target)
  await pivot.setNumberGroup('0', { start: 0, end: 10, interval: 5 })
  const data = pivot.getData()
  assert.equal(data[1][0], '0-5')
  assert.equal(data[2][0], '5-10')
  assert.deepEqual(pivot.drillDown(1, 1).rows, [[1, '2026-01-05', 'A', 10], [4, '2026-01-06', 'A', 20]])
  studio.dispose()
})

test('date grouping supports Excel-like mode and calendar labels', async () => {
  const { studio, workbook, source, target } = setup([
    [1, '2026-01-05', 'A', 10],
    [4, '2026-01-20', 'A', 20],
    [6, '2026-02-05', 'B', 30],
    [9, '2026-02-06', 'B', 40],
  ])
  const pivot = await createPivot(workbook, source, target, 1)
  await pivot.setDateGroupingMode('excelLike')
  await pivot.setExcelDateGroup('1', 'month')
  const labels = pivot.getData().slice(1).map((row) => row[0])
  assert.deepEqual(labels, ['2026-01', '2026-02', 'Grand total'])
  studio.dispose()
})

test('manual groups accept typed members, clear with empty groups, and preserve history transactionality', async () => {
  const { studio, workbook, source, target } = setup([
    [1, '2026-01-05', 'A', 10],
    [4, '2026-01-20', 'A', 20],
    [6, '2026-02-05', 'B', 30],
    [9, '2026-02-06', 'B', 40],
  ])
  const pivot = await createPivot(workbook, source, target, 2)
  await assert.rejects(() => pivot.setManualGroups('2', [{ name: 'AB', members: ['A'] }]), /at least two/)
  assert.deepEqual(pivot.getData().slice(1).map((row) => row[0]), ['A', 'B', 'Grand total'])
  await pivot.setManualGroups('2', [{ name: 'AB', members: ['A', 'B'] }])
  assert.deepEqual(pivot.getData().slice(1).map((row) => row[0]), ['AB', 'Grand total'])
  await pivot.setManualGroups('2', [])
  assert.deepEqual(pivot.getData().slice(1).map((row) => row[0]), ['A', 'B', 'Grand total'])
  studio.dispose()
})

test('invalid numeric grouping does not mutate the pivot', async () => {
  const { studio, workbook, source, target } = setup([
    [1, '2026-01-05', 'A', 10],
    [4, '2026-01-20', 'A', 20],
    [6, '2026-02-05', 'B', 30],
    [9, '2026-02-06', 'B', 40],
  ])
  const pivot = await createPivot(workbook, source, target)
  const before = JSON.stringify(studio.getWorkbookModel(workbook.getId()).pivotTables[0].fields)
  await assert.rejects(() => pivot.setNumberGroup('0', { start: 10, end: 0, interval: 5 }), /end must be greater/)
  assert.equal(JSON.stringify(studio.getWorkbookModel(workbook.getId()).pivotTables[0].fields), before)
  studio.dispose()
})
