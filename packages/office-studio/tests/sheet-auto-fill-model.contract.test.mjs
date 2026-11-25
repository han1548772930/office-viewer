import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createSheetAutoFillPlan } from '../src/features/sheet-workbench/sheet-auto-fill.ts'
import { createSheetFillSeriesPattern, sheetFillSeriesValueAt } from '../src/features/sheet-workbench/sheet-fill-series.ts'
import { parseSheetFillCommandInput } from '../src/features/sheet-workbench/sheet-fill-command.ts'
import { resolveSheetFillTarget, sheetCopyFillRanges, sheetFillPointerTarget } from '../src/features/sheet-workbench/sheet-fill-target.ts'
import { conditionalFormatFormulaAt } from '../src/features/conditional-formatting/conditional-format-formulas.ts'
import { sheetValidationFormulaAt } from '../src/features/sheet-validation/validation-formulas.ts'

// Draft for the later unified verification pass. No execution evidence yet.
const range = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })
const copy = value => JSON.parse(JSON.stringify(value))
function workbook(values = []) {
  const model = createBlankWorkbook(); model.id = 'fill-model'; model.calculationMode = 'manual'
  const sheet = model.sheets[0]; sheet.id = 'data'; model.activeSheetId = sheet.id
  sheet.rowCount = 30; sheet.columnCount = 12
  for (let row = 0; row < values.length; row++) for (let col = 0; col < values[row].length; col++) {
    if (values[row][col] !== undefined) sheet.cells[`${row}:${col}`] = createCell(row, col, values[row][col])
  }
  return model
}
function sequence(values, indexes) {
  const pattern = createSheetFillSeriesPattern(values.map(value => ({ value })))
  return indexes.map(index => sheetFillSeriesValueAt(pattern, index))
}

test('single numeric defaults to copy while explicit series and multiple column lanes continue independently', () => {
  const model = workbook([[2, 10], [5, 20]]), before = copy(model)
  const single = createSheetAutoFillPlan(model, 'data', range(0, 0), range(0, 0, 3))
  assert.equal(single.result.applyType, 'COPY'); assert.equal(single.result.supportsSeries, true)
  assert.deepEqual([1, 2, 3].map(row => single.sheet.cells[`${row}:0`].value), [2, 2, 2])
  const series = createSheetAutoFillPlan(model, 'data', range(0, 0, 1, 1), range(0, 0, 4, 1))
  assert.equal(series.result.applyType, 'SERIES')
  assert.deepEqual([2, 3, 4].map(row => [series.sheet.cells[`${row}:0`].value, series.sheet.cells[`${row}:1`].value]), [[8, 30], [11, 40], [14, 50]])
  assert.deepEqual(copy(model), before)
})

test('number regression, geometric progression and reverse ordinals retain the fixed series contracts', () => {
  assert.deepEqual(sequence([1, 3, 8], [3, 4, 5]), [11, 14.5, 18])
  assert.deepEqual(sequence([2, 4, 8], [3, 4, -1]), [16, 32, 1])
  assert.deepEqual(sequence([1, 2, 5], [3, 4]), [6.66667, 8.66667])
  assert.deepEqual(sequence([10, 8], [-2, -1, 2]), [14, 12, 6])
})

test('last numeric segment, padding, Chinese numbers and cycling names extend without converting text to numbers', () => {
  assert.deepEqual(sequence(['order2-item008', 'order2-item010'], [2, 3]), ['order2-item012', 'order2-item014'])
  assert.deepEqual(sequence(['十', '十一'], [2, 3]), ['十二', '十三'])
  assert.deepEqual(sequence(['一', '二'], [5, 6, 7]), ['六', '日', '一'])
  assert.deepEqual(sequence(['五', '六'], [2, 3]), ['日', '一'])
  assert.deepEqual(sequence(['六', '五'], [2, 3]), ['四', '三'])
  assert.deepEqual(sequence(['六', '日'], [2, 3]), ['一', '二'])
  assert.deepEqual(sequence(['星期六'], [1, 2]), ['星期日', '星期一'])
  assert.deepEqual(sequence(['Dec', 'Jan'], [2, 3]), ['Feb', 'Mar'])
  assert.deepEqual(sequence(['腊月'], [1, 2]), ['正月', '二月'])
})

test('copy translates relative formulas, preserves absolute references and rich runs, and invalidates old formula caches', () => {
  const model = workbook([['label']]), sheet = model.sheets[0]
  const cell = createCell(0, 1, 9); cell.formula = '=A1+$C$1'; cell.dataType = 'formula'; cell.cachedDataType = 'number'
  cell.richText = [{ text: 'cached' }]; cell.metadata = { formulaError: '#VALUE!', vendor: { retained: true } }
  sheet.cells['0:1'] = cell
  sheet.cells['0:0'].richText = [{ text: 'la', font: { bold: true } }, { text: 'bel', vendor: 7 }]
  const plan = createSheetAutoFillPlan(model, 'data', range(0, 0, 0, 1), range(0, 0, 2, 1), { applyType: 'COPY' })
  assert.equal(plan.sheet.cells['2:1'].formula, '=A3+$C$1')
  assert.equal(plan.sheet.cells['2:1'].value, null); assert.equal(plan.sheet.cells['2:1'].cachedDataType, null)
  assert.equal(plan.sheet.cells['2:1'].richText, undefined); assert.equal(plan.sheet.cells['2:1'].metadata.formulaError, undefined)
  assert.deepEqual(plan.sheet.cells['2:1'].metadata.vendor, { retained: true })
  assert.deepEqual(plan.sheet.cells['2:0'].richText, sheet.cells['0:0'].richText)
})

test('only-format replaces the complete source style; no-format continues the sequence with the target style', () => {
  const model = workbook([[2], [4], [90], [91]]), sheet = model.sheets[0]
  sheet.cells['0:0'].style.font.bold = true; sheet.cells['0:0'].style.fill = '#ff0000'
  sheet.cells['1:0'].style.font.bold = true; sheet.cells['1:0'].style.fill = '#ff0000'
  sheet.cells['2:0'].style.font.italic = true; sheet.cells['2:0'].style.fill = '#00ff00'; sheet.cells['2:0'].style.numberFormat = '0.00'
  const format = createSheetAutoFillPlan(model, 'data', range(0, 0, 1), range(0, 0, 3), { applyType: 'ONLY_FORMAT' })
  assert.equal(format.sheet.cells['2:0'].value, 90); assert.deepEqual(format.sheet.cells['2:0'].style, sheet.cells['0:0'].style)
  const content = createSheetAutoFillPlan(model, 'data', range(0, 0, 1), range(0, 0, 3), { applyType: 'NO_FORMAT' })
  assert.equal(content.sheet.cells['2:0'].value, 6); assert.equal(content.sheet.cells['3:0'].value, 8)
  assert.deepEqual(content.sheet.cells['2:0'].style, sheet.cells['2:0'].style)
  assert.equal(content.sheet.cells['2:0'].display, '6.00')
})

test('filter-hidden rows are skipped without skipping manual or outline-hidden rows', () => {
  const model = workbook([[1], [2], [99], [88], [77], [66]]), sheet = model.sheets[0]
  sheet.autoFilter = { range: range(0, 0, 7), columns: {}, filteredRows: [1, 3] }
  sheet.hiddenRows = [4]; sheet.outline = { groups: [{ id: 'rows', axis: 'row', start: 4, end: 5, level: 1, collapsed: true }] }
  const plan = createSheetAutoFillPlan(model, 'data', range(0, 0, 1), range(0, 0, 5), { applyType: 'SERIES' })
  assert.equal(plan.sheet.cells['1:0'].value, 2); assert.equal(plan.sheet.cells['3:0'].value, 88)
  assert.deepEqual([2, 4, 5].map(row => plan.sheet.cells[`${row}:0`].value), [2, 3, 4])
  assert.deepEqual(plan.result.ranges, [range(2, 0), range(4, 0, 5)])
})

test('shrink clears only trailing content and preserves retained cells, formatting, validation and merges', () => {
  const model = workbook([[1], [2], [3], [4]]), sheet = model.sheets[0]
  sheet.cells['3:0'].style.fill = '#abcdef'; sheet.cells['3:0'].hyperlink = 'https://example.test'
  sheet.dataValidations = [{ id: 'rule', type: 'whole', operator: 'greaterThan', formula1: '0', ranges: [range(0, 0, 3)] }]
  const plan = createSheetAutoFillPlan(model, 'data', range(0, 0, 3), range(0, 0, 1))
  assert.equal(plan.target.clear, true); assert.deepEqual(plan.result.selection, range(0, 0, 1))
  assert.deepEqual([0, 1, 2, 3].map(row => plan.sheet.cells[`${row}:0`].value), [1, 2, null, null])
  assert.equal(plan.sheet.cells['3:0'].style.fill, '#abcdef'); assert.deepEqual(plan.sheet.dataValidations, sheet.dataValidations)
})

test('whole source merges are tiled; partial source/target and filtered merge holes reject without mutation', () => {
  const model = workbook([['merged']]), sheet = model.sheets[0]
  sheet.mergedRanges = [range(0, 0, 1, 1)]
  const plan = createSheetAutoFillPlan(model, 'data', range(0, 0, 1, 1), range(0, 0, 5, 1), { applyType: 'COPY' })
  assert.deepEqual(plan.sheet.mergedRanges, [range(0, 0, 1, 1), range(2, 0, 3, 1), range(4, 0, 5, 1)])
  const before = copy(model)
  assert.throws(() => createSheetAutoFillPlan(model, 'data', range(0, 0), range(0, 0, 3)), /合并/)
  assert.throws(() => createSheetAutoFillPlan(model, 'data', range(0, 0, 1, 1), range(0, 0, 4, 1), { applyType: 'COPY' }), /截断/)
  sheet.autoFilter = { range: range(0, 0, 8, 1), columns: {}, filteredRows: [3] }
  assert.throws(() => createSheetAutoFillPlan(model, 'data', range(0, 0, 1, 1), range(0, 0, 5, 1), { applyType: 'COPY' }), /筛选/)
  delete sheet.autoFilter; assert.deepEqual(copy(model), before)
})

test('copied validation rejects an invalid series atomically and retains warning issues for an accepted candidate', () => {
  const model = workbook([[1], [4]]), sheet = model.sheets[0]
  sheet.dataValidations = [{ id: 'limit', type: 'whole', operator: 'lessThanOrEqual', formula1: '4', ranges: [range(0, 0, 1)],
    allowBlank: false, showErrorMessage: true, errorStyle: 'stop', error: 'Too large', errorTitle: 'Limit' }]
  const before = copy(model), blocked = createSheetAutoFillPlan(model, 'data', range(0, 0, 1), range(0, 0, 3), { applyType: 'SERIES' })
  assert.equal(blocked.result.accepted, false); assert.equal(blocked.result.reason, 'validation'); assert.equal(blocked.sheet, undefined)
  assert.ok(blocked.result.issues.every(issue => issue.column === 0 && issue.severity === 'stop')); assert.deepEqual(copy(model), before)
  sheet.dataValidations[0].errorStyle = 'warning'
  const warning = createSheetAutoFillPlan(model, 'data', range(0, 0, 1), range(0, 0, 3), { applyType: 'SERIES' })
  assert.equal(warning.result.accepted, true); assert.equal(warning.result.issues.length, 2)
})

test('CF and validation project clipped forward/backward tiles in both axes while each mode preserves its resource contract', () => {
  const cases = [
    [range(3, 1, 4, 2), range(3, 1, 7, 2)], [range(4, 1, 5, 2), range(1, 1, 5, 2)],
    [range(3, 1, 4, 2), range(3, 1, 4, 5)], [range(3, 4, 4, 5), range(3, 1, 4, 5)],
  ]
  const contains = (ranges, row, col) => ranges.some(r => r.startRow <= row && row <= r.endRow && r.startCol <= col && col <= r.endCol)
  const formula = (row, col) => `=${String.fromCharCode(65 + col)}${row + 1}>0`
  for (const [source, target] of cases) {
    const model = workbook(), sheet = model.sheets[0]
    for (let row = source.startRow; row <= source.endRow; row++) for (let col = source.startCol; col <= source.endCol; col++) {
      sheet.cells[`${row}:${col}`] = createCell(row, col, 1)
    }
    sheet.conditionalFormats = [{ id: 'source-cf', type: 'expression', formula1: formula(source.startRow, source.startCol),
      operator: 'greaterThan', style: { fill: '#35a65f' }, priority: 1, stopIfTrue: false, ranges: [source] }]
    sheet.dataValidations = [{ id: 'source-dv', type: 'custom', formula1: formula(source.startRow, source.startCol), ranges: [source],
      allowBlank: false, showErrorMessage: true, errorStyle: 'stop' }]
    const before = copy(model)
    for (const applyType of ['COPY', 'ONLY_FORMAT', 'NO_FORMAT']) {
      const plan = createSheetAutoFillPlan(model, 'data', source, target, { applyType })
      assert.equal(plan.result.accepted, true)
      if (applyType === 'NO_FORMAT') assert.deepEqual(plan.sheet.conditionalFormats, sheet.conditionalFormats)
      if (applyType === 'ONLY_FORMAT') assert.deepEqual(plan.sheet.dataValidations, sheet.dataValidations)
      for (const r of plan.result.ranges) for (let row = r.startRow; row <= r.endRow; row++) for (let col = r.startCol; col <= r.endCol; col++) {
        if (applyType !== 'NO_FORMAT') {
          const rules = plan.sheet.conditionalFormats.filter(rule => contains(rule.ranges, row, col))
          assert.equal(rules.length, 1); assert.equal(conditionalFormatFormulaAt(rules[0].formula1, rules[0], row, col), formula(row, col))
        }
        if (applyType !== 'ONLY_FORMAT') {
          const rules = plan.sheet.dataValidations.filter(rule => contains(rule.ranges, row, col))
          assert.equal(rules.length, 1); assert.equal(sheetValidationFormulaAt(rules[0].formula1, rules[0], row, col), formula(row, col))
        }
      }
    }
    assert.deepEqual(copy(model), before)
  }
})

test('unreadable sources, protected targets and sync-budget overflow preserve the original workbook', () => {
  const model = workbook([[1], [2]]), sheet = model.sheets[0]
  sheet.rangeProtectionRules = [{ id: 'locked', permissionId: 'permission', ownerUserId: 'remote', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'othersCanView', ranges: [range(2, 0, 3)] }]
  const before = copy(model)
  assert.throws(() => createSheetAutoFillPlan(model, 'data', range(0, 0, 1), range(0, 0, 3), { userId: 'local-user' }), /保护/)
  assert.deepEqual(copy(model), before)
  sheet.rangeProtectionRules[0].ranges = [range(0, 0, 1)]; sheet.rangeProtectionRules[0].viewState = 'noOneElseCanView'
  assert.throws(() => createSheetAutoFillPlan(model, 'data', range(0, 0, 1), range(0, 0, 3), { userId: 'local-user' }), /读取/)
  delete sheet.rangeProtectionRules
  assert.throws(() => createSheetAutoFillPlan(model, 'data', range(0, 0), range(0, 0, 100_000)), /Worker/)
})

test('pointer, Facade geometry and copy-down/right share a single direction and reject invalid ranges/aliases', () => {
  const source = range(2, 2, 3, 4)
  assert.deepEqual(sheetFillPointerTarget(source, 0, 1), range(0, 2, 3, 4))
  assert.deepEqual(sheetFillPointerTarget(source, 2, 3), range(2, 2, 2, 4))
  assert.throws(() => resolveSheetFillTarget(source, range(0, 0, 5, 6)), /一个方向/)
  assert.equal(sheetCopyFillRanges(range(0, 2), 'down'), undefined)
  assert.deepEqual(sheetCopyFillRanges(range(4, 2), 'down'), { sourceRange: range(3, 2), targetRange: range(3, 2, 4) })
  assert.deepEqual(sheetCopyFillRanges(range(4, 2), 'right'), { sourceRange: range(4, 1), targetRange: range(4, 1, 4, 2) })
  const parsed = parseSheetFillCommandInput({ unitId: 'book', subUnitId: 'data', sourceRange: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    targetRange: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 0 }, applyType: 'NO_FORMAT' })
  assert.deepEqual(parsed.sourceRange, range(0, 0, 1)); assert.equal(parsed.applyType, 'NO_FORMAT')
  assert.throws(() => parseSheetFillCommandInput({ sheetId: 'a', subUnitId: 'b' }), /不一致/)
  assert.throws(() => parseSheetFillCommandInput({ applyType: 'unsupported' }), /模式/)
  assert.equal(parseSheetFillCommandInput({ type: 'NO_FORMAT' }).applyType, 'NO_FORMAT')
  assert.throws(() => parseSheetFillCommandInput({ type: 'COPY', applyType: 'SERIES' }), /不一致/)
})
