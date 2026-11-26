import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { createCell } from '../src/domain/model.ts'
import { calculatePivotReport, getPivotValuesPosition, pivotDrillDown, pivotOutputValues,
  setPivotValuesPosition } from '../src/features/pivot/services/pivot-engine.ts'
import { applyPivotChartEdit, getPivotChartReport } from '../src/features/pivot/services/pivot-chart-engine.ts'
import { transformPivotSourceReport } from '../src/features/pivot/services/pivot-source-lifecycle.ts'
import { PIVOT_AXIS_RECORDS, pivotAxisWorkbook } from './fixtures/pivot-axis-state-fixture.mjs'

// Unexecuted contracts only. No test/type/build/browser/WASM/benchmark is run while drafting this batch.
const clone = value => structuredClone(value)
const positions = ['row', 'column'], owners = ['owned', 'linked'], indexes = [0, 1, 2, 3]
const measureOrder = {
  row: ['SSSSSSSSAAAAAAAA', 'SSSSAAAASSSSAAAA', 'SSAASSAASSAASSAA', 'SASASASASASASASA'],
  column: ['SSSSAAAA', 'SSSSAAAA', 'SSAASSAA', 'SASASASA'],
}
const setup = (ownership = 'owned', options = {}) => {
  const workbook = pivotAxisWorkbook({ ownership, multipleValues: true, ...options }), chart = workbook.charts[0]
  return { workbook, chart, report: () => getPivotChartReport(workbook, chart),
    output: () => calculatePivotReport(workbook, getPivotChartReport(workbook, chart)),
    edit: payload => applyPivotChartEdit(workbook, chart, payload) }
}
const source = f => f.workbook.sheets.find(sheet => sheet.id === 'axis-data').cells
function sourceScope(f, row, column) {
  const filters = { ...row.fieldFilters, ...column.fieldFilters }
  return PIVOT_AXIS_RECORDS.filter(record => Object.entries(filters).every(([id, value]) => {
    const field = f.report().fields.find(field => field.id === id)
    return field && Object.is(record[field.sourceColumn - 2], value)
  }))
}
function assertActualAggregates(f, output = f.output()) {
  for (const row of output.rowAxis) for (const column of output.columnAxis) {
    const raw = sourceScope(f, row, column), sum = raw.reduce((total, record) => total + record[6], 0)
    const measure = f.report().fields.find(field => field.id === (row.valueFieldId ?? column.valueFieldId))
    const expected = raw.length ? measure.aggregation === 'average' ? sum / raw.length : sum : null
    assert.equal(output.values[row.index][column.index], expected,
      `${measure.id} ${JSON.stringify(row.labels)} ${JSON.stringify(column.labels)}`)
  }
}

for (const ownership of owners) for (const position of positions) for (const index of indexes) {
  test(`${ownership}/${position}/${index} inserts ΣValues between real levels without changing aggregate scopes`, () => {
    const f = setup(ownership), before = clone(source(f))
    assert.equal(f.edit({ type: 'updateValuePosition', position, index }), true)
    assert.deepEqual(getPivotValuesPosition(f.report()), { position, index })
    const output = f.output(), axis = position === 'row' ? output.rowAxis : output.columnAxis
    assert.equal(axis.map(entry => entry.valueFieldId === 'amount-value' ? 'S' : 'A').join(''), measureOrder[position][index])
    if (position === 'row') {
      const labels = ['Region', 'Product', 'Item']; labels.splice(index, 0, 'Values')
      assert.deepEqual(output.values[output.headerRows - 1].slice(0, 4), labels)
      assert.equal(output.headerRows, 3)
    } else {
      assert.equal(output.values[index][3], 'Amount')
      const dimensions = [2024, 'Q1', 'Retail']; dimensions.splice(index, 0, 'Amount')
      assert.deepEqual(output.values.slice(0, 4).map(row => row[3]), dimensions)
      assert.equal(output.headerRows, 4)
    }
    assertActualAggregates(f, output); assert.deepEqual(source(f), before)
    if (ownership === 'linked') assert.deepEqual(pivotOutputValues(f.workbook, f.report()), output.values)
    else assert.equal(f.report().generatedCellKeys, undefined)
  })
}

for (const position of positions) test(`${position} auto tail and an oversized raw index have the same effective layout`, () => {
  const f = setup(); f.edit({ type: 'updateValuePosition', position, index: 3 }); const tail = f.output()
  f.edit({ type: 'updateValuePosition', position })
  assert.deepEqual(getPivotValuesPosition(f.report()), { position, index: -1 }); assert.equal(f.report().options.valueIndex, undefined)
  assert.deepEqual(f.output(), tail)
  f.edit({ type: 'updateValuePosition', position, index: 100 })
  assert.equal(getPivotValuesPosition(f.report()).index, 100); assert.deepEqual(f.output(), tail)
})

for (const ownership of owners) test(`${ownership} absent special field keeps all measures and uses fixed column fallback`, () => {
  const f = setup(ownership); f.edit({ type: 'updateValuePosition', position: 'column', index: 1 }); const columns = f.output()
  f.edit({ type: 'updateValuePosition', position: 'none', index: 1 })
  assert.deepEqual(getPivotValuesPosition(f.report()), { position: 'none', index: 1 })
  assert.deepEqual(f.output(), columns); assert.equal(f.report().fields.filter(field => field.area === 'value').length, 2)
  assertActualAggregates(f)
})

for (const position of positions) test(`single measure ${position} retains the existing row preference and its raw index`, () => {
  const f = setup('owned', { multipleValues: false }); f.edit({ type: 'updateValuePosition', position, index: 1 })
  assert.equal(f.report().fields.filter(field => field.area === 'value').length, 1)
  assert.deepEqual(getPivotValuesPosition(f.report()), { position, index: 1 }); assertActualAggregates(f)
  // Compatibility: the current local row renderer still exposes its Values column for one measure.
  if (position === 'row') assert.equal(f.output().values[f.output().headerRows - 1][1], 'Values')
})

for (const position of positions) test(`${position} field moves retain the special gap and unchanged moves remain no-ops`, () => {
  const f = setup(); f.edit({ type: 'updateValuePosition', position, index: 1 })
  const first = f.report().fields.find(field => field.area === position && field.position === 0).id
  f.edit({ type: 'moveField', fieldId: first, fieldArea: position, index: 2 })
  assert.equal(getPivotValuesPosition(f.report()).index, 0)
  f.edit({ type: 'moveField', fieldId: first, fieldArea: position, index: 0 })
  assert.equal(getPivotValuesPosition(f.report()).index, 1)
  const before = clone(f.workbook); assert.equal(f.edit({ type: 'moveField', fieldId: first, fieldArea: position, index: 0 }), false)
  assert.deepEqual(f.workbook, before); assertActualAggregates(f)
})

for (const position of positions) test(`${position} removing dimensions after and before Values preserves its surviving gap`, () => {
  const f = setup(); f.edit({ type: 'updateValuePosition', position, index: 1 })
  const fields = f.report().fields.filter(field => field.area === position).sort((a, b) => a.position - b.position)
  f.edit({ type: 'removeFields', fieldIds: [fields[1].id] }); assert.equal(getPivotValuesPosition(f.report()).index, 1)
  f.edit({ type: 'removeFields', fieldIds: [fields[0].id] }); assert.equal(getPivotValuesPosition(f.report()).index, 0)
  assertActualAggregates(f)
})

for (const position of positions) test(`${position} inserting a real field at the Values gap places it before the special field`, () => {
  const f = setup(); f.edit({ type: 'updateValuePosition', position, index: 1 })
  const removed = clone(f.report().fields.find(field => field.area === position && field.position === 2))
  f.edit({ type: 'removeFields', fieldIds: [removed.id] }); assert.equal(getPivotValuesPosition(f.report()).index, 1)
  f.edit({ type: 'addField', dataFieldId: removed.sourceFieldId, fieldArea: position, index: 1 })
  assert.equal(getPivotValuesPosition(f.report()).index, 2); assertActualAggregates(f)
})

for (const position of positions) test(`source retirement before ${position} Values preserves the gap and local field identities`, () => {
  const f = setup(); f.edit({ type: 'updateValuePosition', position, index: 2 })
  const removed = f.report().fields.find(field => field.area === position && field.position === 0), before = clone(f.workbook)
  const next = transformPivotSourceReport(before, f.report(), { sheetId: 'axis-data',
    range: range => ({ ...range, endCol: range.endCol - 1 }),
    column: column => column === removed.sourceColumn ? undefined : column > removed.sourceColumn ? column - 1 : column,
    point: (row, col) => ({ row, col }),
  }, true)
  assert.equal(getPivotValuesPosition(next).index, 1); assert.ok(next.fields.every(field => field.id !== removed.id))
  assert.deepEqual(next.fields.map(field => field.id), f.report().fields.filter(field => field.id !== removed.id).map(field => field.id))
  assert.deepEqual(f.workbook, before)
})

for (const ownership of owners) test(`${ownership} full fields configuration retains or explicitly replaces position and index`, () => {
  const f = setup(ownership); f.edit({ type: 'updateValuePosition', position: 'row', index: 1 })
  assert.equal(f.edit({ type: 'setFieldsConfig', config: { fields: clone(f.report().fields) } }), false)
  assert.deepEqual(getPivotValuesPosition(f.report()), { position: 'row', index: 1 })
  f.edit({ type: 'setFieldsConfig', config: { fields: clone(f.report().fields), valuePosition: 'column', valueIndex: 2,
    options: { valuePosition: 'column', valueIndex: 2 } } })
  assert.deepEqual(getPivotValuesPosition(f.report()), { position: 'column', index: 2 }); assertActualAggregates(f)
  const before = clone(f.workbook)
  assert.equal(f.edit({ type: 'setFieldsConfig', config: { fields: clone(f.report().fields), valuePosition: 'column', valueIndex: 2 } }), false)
  assert.deepEqual(f.workbook, before)
})

for (const ownership of owners) for (const index of [-2, 1.5, Number.NaN, Number.POSITIVE_INFINITY, '1', null]) {
  test(`${ownership} invalid Values index ${typeof index}:${String(index)} changes no report, matrix, chart or cache`, () => {
    const f = setup(ownership), before = clone(f.workbook)
    assert.throws(() => f.edit({ type: 'updateValuePosition', position: 'row', index }), /value index/)
    assert.deepEqual(f.workbook, before)
  })
}

for (const invalid of [{ valueIndex: -2 }, { valuePosition: 'none', options: { valuePosition: 'row' } },
  { valueIndex: 0, options: { valueIndex: 1 } }, { options: { valueIndex: Number.POSITIVE_INFINITY } }]) {
  test(`invalid full Values configuration ${JSON.stringify(invalid)} refuses before publication`, () => {
    const f = setup(), before = clone(f.workbook)
    assert.throws(() => f.edit({ type: 'setFieldsConfig', config: { fields: clone(f.report().fields), ...invalid } }), /value index|位置有冲突/)
    assert.deepEqual(f.workbook, before)
  })
}

for (const ownership of owners) for (const position of positions) for (const index of indexes) {
  test(`${ownership}/${position}/${index} subtotals, grand totals and collapsed paths keep source scopes and drill-down`, () => {
    const f = setup(ownership, { totals: true }), before = clone(source(f))
    f.edit({ type: 'setCollapse', fieldId: 'region-label', collapse: true, item: 'North' })
    f.edit({ type: 'setCollapse', fieldId: 'quarter-label', collapse: true, item: 'Q1' })
    f.edit({ type: 'updateValuePosition', position, index })
    const output = f.output(); assertActualAggregates(f, output); assert.deepEqual(source(f), before)
    const row = output.rowAxis.find(entry => entry.kind === 'collapsed'), column = output.columnAxis.find(entry => entry.kind === 'collapsed')
    assert.ok(row); assert.ok(column)
    if (ownership === 'linked') {
      const details = pivotDrillDown(f.workbook, f.report(), row.index, column.index)
      assert.deepEqual(details.rows, sourceScope(f, row, column)); assert.equal(details.rows.length, 8)
    }
    assert.ok(output.rowAxis.some(entry => entry.kind === 'grand-total'))
    assert.ok(output.columnAxis.some(entry => entry.kind === 'grand-total'))
  })
}

test('a linked Values axis change that expands into occupied worksheet content leaves every resource untouched', () => {
  const f = setup('linked'), host = f.workbook.sheets.find(sheet => sheet.id === f.report().anchor.subUnitId)
  const row = f.report().anchor.row + 18, col = f.report().anchor.col + 7
  host.cells[`${row}:${col}`] = createCell(row, col, 'occupied')
  const before = clone(f.workbook)
  assert.throws(() => f.edit({ type: 'updateValuePosition', position: 'row', index: 0 }), /overlap|occupied/i)
  assert.deepEqual(f.workbook, before)
})

for (const position of ['none', ...positions]) for (const index of [-1, 0, 1, 3, 100]) {
  test(`JSON reopening retains configured ${position}/${index} and the exact calculated matrix`, () => {
    const f = setup(); f.edit({ type: 'updateValuePosition', position, index })
    const before = f.output(), reopened = workbookFromInput(JSON.parse(JSON.stringify(f.workbook))), chart = reopened.charts[0]
    const restored = getPivotChartReport(reopened, chart)
    assert.deepEqual(getPivotValuesPosition(restored), { position, index })
    assert.deepEqual(calculatePivotReport(reopened, restored), before)
  })
}

test('the Values getter returns a defensive snapshot and direct deferred setters share the same no-op state', () => {
  const f = setup(); setPivotValuesPosition(f.workbook, f.report(), 'row', 0, 'data-only')
  const state = getPivotValuesPosition(f.report()); state.index = 100; state.position = 'none'
  assert.deepEqual(getPivotValuesPosition(f.report()), { position: 'row', index: 0 })
  const before = clone(f.workbook); setPivotValuesPosition(f.workbook, f.report(), 'row', 0, 'data-only')
  assert.deepEqual(f.workbook, before)
})
