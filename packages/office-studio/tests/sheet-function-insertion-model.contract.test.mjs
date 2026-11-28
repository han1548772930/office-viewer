import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred: pure planning contracts; these do not replace UI or Office reopening.
let server, sdk
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  sdk = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const range = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })
function fixture() {
  const sheet = sdk.createBlankSheet('Data')
  const put = (row, col, value) => sheet.cells[`${row}:${col}`] = sdk.createCell(row, col, value)
  return { sheet, put, plan: (selections, name = 'SUM') => sdk.planSheetFunctionInsertion(sheet, selections, name) }
}

test('single-cell insertion prefers the upper nearest numeric block and retains intervening text/blanks in its reference', () => {
  const h = fixture(); h.put(0, 1, 90); h.put(1, 1, 'break'); h.put(2, 1, 2); h.put(3, 1, 3); h.put(4, 1, 'label'); h.put(6, 0, 4)
  const before = JSON.stringify(h.sheet), plan = h.plan([{ range: range(6, 1) }], 'sum')
  assert.equal(plan.edit.formula, '=SUM(B3:B6)'); assert.deepEqual(plan.list, [])
  assert.equal(JSON.stringify(h.sheet), before)
})

test('left-hand references are used only when the column has no numeric cells; empty input stays unclosed', () => {
  const h = fixture(); h.put(4, 0, 2); h.put(4, 1, 3); h.put(4, 2, 'skip')
  assert.equal(h.plan([{ range: range(4, 4) }]).edit.formula, '=SUM(A5:D5)')
  assert.equal(h.plan([{ range: range(0, 0) }], 'COUNT').edit.formula, '=COUNT(')
})

test('a nonnumeric rectangle enters editing at its real primary rather than filling the rectangle', () => {
  const h = fixture(); h.put(0, 2, 7)
  const plan = h.plan([{ range: range(2, 1, 3, 2), primary: { row: 3, column: 2 } }])
  assert.equal(plan.edit.formula, '=SUM(C1:C3)'); assert.deepEqual(plan.edit.primary, { row: 3, column: 2 })
  assert.deepEqual(sdk.expandSheetFunctionInsertions(plan.list, plan.listOfRangeHasNumber), [])
})

test('rich body numeric text, declared scalar text, booleans and missing rich bodies use fixed number identity', () => {
  const numeric = sdk.createCell(0, 0, '12'), plainText = sdk.createCell(0, 0, '12'), declared = sdk.createCell(0, 0, '12')
  numeric.richText = [{ text: '12e+3' }]; declared.dataType = 'number'
  assert.equal(sdk.sheetFunctionNumberCell(numeric), true); assert.equal(sdk.sheetFunctionNumberCell(plainText), false)
  assert.equal(sdk.sheetFunctionNumberCell(declared), true); assert.equal(sdk.sheetFunctionNumberCell(sdk.createCell(0, 0, false)), false)
  numeric.metadata = { publicBody: { body: { dataStream: ' 12\r\n' } } }
  assert.equal(sdk.sheetFunctionNumberCell(numeric), true)
  numeric.metadata.publicBody = {}; assert.equal(sdk.sheetFunctionNumberCell(numeric), false)
  numeric.metadata.publicBody = { body: { dataStream: ' \r\n' } }; assert.equal(sdk.sheetFunctionNumberCell(numeric), false)
})

test('numeric row uses an absent final cell, but a materialized blank moves the result past it', () => {
  const h = fixture(); h.put(0, 0, 2); h.put(0, 1, 3)
  let plan = h.plan([{ range: range(0, 0, 0, 2) }])
  assert.deepEqual(plan.listOfRangeHasNumber[0].primary, { row: 0, column: 2 })
  assert.equal(plan.listOfRangeHasNumber[0].formula, '=SUM(A1:B1)')
  h.put(0, 2, null)
  plan = h.plan([{ range: range(0, 0, 0, 2) }])
  assert.deepEqual(plan.listOfRangeHasNumber[0].primary, { row: 0, column: 3 })
  assert.equal(plan.listOfRangeHasNumber[0].formula, '=SUM(A1:C1)')
})

test('numeric rectangles align every result column to the maximum first-absent row', () => {
  const h = fixture(); h.put(0, 0, 1); h.put(1, 0, 2); h.put(2, 0, null); h.put(3, 0, 'below'); h.put(0, 1, 4)
  const plan = h.plan([{ range: range(0, 0, 2, 1) }])
  assert.deepEqual(plan.listOfRangeHasNumber.map(item => item.primary), [{ row: 4, column: 0 }, { row: 4, column: 1 }])
  assert.deepEqual(plan.listOfRangeHasNumber.map(item => item.formula), ['=SUM(A1:A3)', '=SUM(B1:B3)'])
})

test('an empty result row already inside the numeric rectangle is excluded from every reference', () => {
  const h = fixture(); h.put(0, 0, 1); h.put(0, 1, 3)
  const plan = h.plan([{ range: range(0, 0, 2, 1) }])
  assert.deepEqual(plan.listOfRangeHasNumber.map(item => item.primary), [{ row: 2, column: 0 }, { row: 2, column: 1 }])
  assert.deepEqual(plan.listOfRangeHasNumber.map(item => item.formula), ['=SUM(A1:A2)', '=SUM(B1:B2)'])
})

test('empty one-dimensional selections expand relative formulas while preserving absolute references', () => {
  const h = fixture(); h.put(0, 0, 1); h.put(1, 0, 2)
  const plan = h.plan([{ range: range(2, 0, 2, 2) }])
  assert.equal(plan.edit, undefined)
  assert.deepEqual(sdk.expandSheetFunctionInsertions(plan.list, []).map(item => item.formula), ['=SUM(A1:A2)', '=SUM(B1:B2)', '=SUM(C1:C2)'])
  const expanded = sdk.expandSheetFunctionInsertions([{ range: range(2, 0, 2, 1), primary: { row: 2, column: 0 }, formula: '=SUM(A1:$A$2)' }], [])
  assert.deepEqual(expanded.map(item => item.formula), ['=SUM(A1:$A$2)', '=SUM(B1:$A$2)'])
})

test('multiple selections retain planned writes before the first selection with no reference', () => {
  const h = fixture(); h.put(0, 0, 1)
  const plan = h.plan([{ range: range(1, 0) }, { range: range(1, 4) }, { range: range(3, 0) }])
  assert.equal(plan.list.length, 1); assert.deepEqual(plan.edit.primary, { row: 1, column: 4 })
  assert.equal(plan.edit.formula, '=SUM(')
})

test('shared and numeric writes follow fixed last-write order; budget and coordinates reject before expansion', () => {
  const item = { range: range(0, 0, 0, 1), primary: { row: 0, column: 0 }, formula: '=A2' }
  const writes = sdk.expandSheetFunctionInsertions([item], [{ ...item, primary: { row: 0, column: 1 }, formula: '=99' }])
  assert.deepEqual(writes.map(item => item.formula), ['=A2', '=99'])
  assert.throws(() => sdk.expandSheetFunctionInsertions([item], [], 1), /预算/)
  assert.throws(() => sdk.expandSheetFunctionInsertions([], [{ ...item, primary: { row: -1, column: 0 } }]), /范围/)
  assert.throws(() => fixture().plan([{ range: range(1000, 0) }]), /工作表/)
  assert.throws(() => fixture().plan([{ range: range(0, 0) }], 'SUM(A1)'), /名称/)
})
