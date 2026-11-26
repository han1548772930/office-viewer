import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { assertSheetCellShiftEdit, inverseSheetCellShiftPosition, mapSheetCellShiftPosition, transformSheetCellShiftRange } from '../src/domain/sheet-cell-shift.ts'
import { assertSheetCellShiftConfirmationAllowed, planWorkbookCellShiftCommand, SHEET_CELL_SHIFT_COMMANDS, SHEET_CELL_SHIFT_CONFIRM_COMMANDS, SheetCellShiftDimensions,
  sheetCellShiftCommandInput, sheetCellShiftCommandRange, sheetCellShiftDirectionForCommand } from '../src/features/sheet-workbench/sheet-cell-shift-command.ts'
import { transformFormulaReferenceAreasForCellShift } from '../src/features/formula/formula-reference-areas.ts'
import { sheetClipboardChangedRuleRanges } from '../src/features/sheet-workbench/sheet-clipboard-changes.ts'
import { assertSheetClipboardWorkbookChangesEditable } from '../src/features/sheet-workbench/sheet-clipboard-permissions.ts'
import { refreshWorkbookCellStyles } from '../src/features/sheet-workbench/sheet-style-inheritance.ts'
import { normalizeSheetRangeThemeStyle } from '../src/features/sheet-workbench/sheet-range-theme-style.ts'

// Drafts only. No tests, type/build checks, dev/browser, Rust/WASM or benchmarks have been run.
const range = (startRow, endRow = startRow, startCol = 1, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
const apiRange = value => ({ startRow: value.startRow, endRow: value.endRow, startColumn: value.startCol, endColumn: value.endCol })
const cell = (row, col, value, extra = {}) => ({ ...createCell(row, col, value), authoredStyle: {}, ...extra })
function sample() {
  const model = createBlankWorkbook(); model.id = 'cell-shift'; model.calculationMode = 'manual'
  model.sheets = ['Data', 'Report'].map((name, index) => ({ ...createBlankSheet(name), id: index ? 'report' : 'data',
    rowCount: 30, columnCount: 12, defaultRowAutoHeight: false }))
  model.activeSheetId = 'data'; return model
}
const cf = (id, ranges, formula1 = '=A1>0', extra = {}) => ({ id, ranges, type: 'expression', operator: 'equal', formula1,
  style: { background: '#ffeeaa' }, priority: 1, stopIfTrue: false, ...extra })
const validation = (id, ranges, extra = {}) => ({ id, ranges, type: 'list', formula1: 'A1,B2', allowBlank: true, ...extra })
const table = (id, source, extra = {}) => ({ id, name: id, sheetId: 'data', range: source,
  columns: Array.from({ length: source.endCol - source.startCol + 1 }, (_, index) => ({ id: `${id}-column-${index}`, name: `Field ${index}` })),
  showHeaderRow: true, showTotalsRow: false, styleName: 'TableStyleMedium2', ...extra })
const chart = (id, source, extra = {}) => ({ id, sheetId: 'report', type: 'line', source: { sheetId: 'data', range: source, orientation: 'columns' },
  position: { x: 0, y: 0 }, size: { width: 240, height: 160 }, categoryField: 0, valueFields: [1],
  legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 2, ...extra })
const pivot = (extra = {}) => ({ id: 'pivot', source: { unitId: 'cell-shift', subUnitId: 'data', sheetName: 'Data', range: range(0, 5, 1, 3) },
  anchor: { unitId: 'cell-shift', subUnitId: 'report', row: 12, col: 5 }, fields: [
    { id: 'category', sourceColumn: 1, name: 'Category', area: 'row', position: 0 },
    { id: 'value', sourceColumn: 2, name: 'Value', area: 'value', aggregation: 'sum', position: 0 },
  ], options: { repeatRowLabels: false, showRowGrandTotal: true, showColGrandTotal: true, showRowSubtotals: true, showColSubtotals: true, valuePosition: 'column' },
  revision: 3, drillDownMap: { '12:5': { fieldFilters: {}, valueFieldId: 'value' } }, ...extra })

test('explicit primary ranges are normalized independently and win over a non-last selection', () => {
  const primary = range(1), last = range(5, 5, 3)
  const parsed = sheetCellShiftCommandInput({ range: apiRange(primary), ranges: [apiRange(primary), apiRange(last)] })
  assert.deepEqual(parsed.range, primary); assert.deepEqual(sheetCellShiftCommandRange(parsed, parsed.ranges, 1), primary)
  assert.deepEqual(sheetCellShiftCommandRange({}, [primary, last], 0), primary)
  assert.deepEqual(sheetCellShiftCommandRange({ selections: [primary, last] }, [primary, last], 0), last)
  assert.throws(() => sheetCellShiftCommandInput({ range: { ...primary, startRow: -1 }, ranges: [last] }), /范围/)
  assert.throws(() => sheetCellShiftCommandInput({ expectedHistoryGeneration: 0 }), /版本/)
  assert.throws(() => sheetCellShiftCommandInput({ sheetId: 'a', subUnitId: 'b' }), /标识/)
})

test('fixed Dimension values and all eight command IDs resolve to the same four edit directions', () => {
  assert.deepEqual(SheetCellShiftDimensions, { COLUMNS: 0, ROWS: 1 })
  for (const direction of ['right', 'down', 'left', 'up']) {
    assert.equal(sheetCellShiftDirectionForCommand(SHEET_CELL_SHIFT_COMMANDS[direction]), direction)
    assert.equal(sheetCellShiftDirectionForCommand(SHEET_CELL_SHIFT_CONFIRM_COMMANDS[direction]), direction)
  }
  assert.equal(sheetCellShiftDirectionForCommand('sheet.command.insert-row-before'), undefined)
})

test('all four directions move sparse content only inside the band and leave the source immutable', () => {
  const cases = [
    ['right', '2:5', '2:4', '2:7'], ['down', '5:2', '4:2', '7:2'],
    ['left', '2:5', undefined, '2:3'], ['up', '5:2', undefined, '3:2'],
  ]
  for (const [direction, tailKey, selectedKey, movedTail] of cases) {
    const model = sample(), sheet = model.sheets[0]
    sheet.cells['2:2'] = cell(2, 2, 'selected')
    const [row, col] = tailKey.split(':').map(Number); sheet.cells[tailKey] = cell(row, col, 'tail')
    sheet.cells['5:5'] = cell(5, 5, 'outside'); const before = cloneModel(model)
    const next = planWorkbookCellShiftCommand(model, 'data', range(1, 2, 1, 2), direction)
    assert.equal(next.workbook.sheets[0].cells[movedTail].value, 'tail')
    if (selectedKey) assert.equal(next.workbook.sheets[0].cells[selectedKey].value, 'selected')
    else assert.ok(!Object.values(next.workbook.sheets[0].cells).some(item => item.value === 'selected'))
    assert.equal(next.workbook.sheets[0].cells['5:5'].value, 'outside'); assert.deepEqual(model, before)
  }
})

test('inserted blanks retain only direct, catalog, reset and cell-theme sources while rich content follows the moved cell', () => {
  const model = sample(), sheet = model.sheets[0]
  model.styles = { direct: { font: { bold: true }, fill: '#eeddcc' } }
  sheet.cells['1:1'] = cell(1, 1, 'rich', { formula: '=A1', authoredStyle: { font: { italic: true } }, authoredStyleReference: 'direct',
    clearedStyleProperties: ['font.bold'], themeStyle: { fill: '#aabbcc' }, richText: [{ text: 'rich', cl: { rgb: '#123456' } }],
    hyperlink: '#Data!A1', metadata: { note: 'kept', vendor: { exact: true } }, cachedDataType: 's' })
  sheet.cells['1:2'] = cell(1, 2, 'plain'); refreshWorkbookCellStyles(model, undefined, false)
  const next = planWorkbookCellShiftCommand(model, 'data', range(1, 1, 1, 2), 'down').workbook.sheets[0], blank = next.cells['1:1'], moved = next.cells['2:1']
  assert.equal(blank.value, null); assert.equal(blank.formula, null)
  for (const key of ['richText', 'hyperlink', 'metadata', 'sparkline', 'cachedDataType']) assert.equal(blank[key], undefined)
  assert.deepEqual(blank.authoredStyle, { font: { italic: true } }); assert.equal(blank.authoredStyleReference, 'direct')
  assert.deepEqual(blank.clearedStyleProperties, ['font.bold']); assert.deepEqual(blank.themeStyle, { fill: '#aabbcc' })
  assert.equal(next.cells['1:2'], undefined); assert.equal(next.cells['2:2'].value, 'plain')
  assert.deepEqual(moved.richText, sheet.cells['1:1'].richText); assert.deepEqual(moved.metadata, sheet.cells['1:1'].metadata)
  assert.notEqual(moved.metadata, sheet.cells['1:1'].metadata)
})

test('virtual range-theme blanks remain sparse and fragmented geometry keeps the original header', () => {
  const model = sample(), sheet = model.sheets[0]
  model.rangeThemes = { pattern: normalizeSheetRangeThemeStyle({ name: 'pattern', headerRowStyle: { bg: { rgb: '#123456' } } }) }
  sheet.rangeThemeRules = [{ id: 'theme', themeName: 'pattern', range: range(0, 5, 0, 3) }]
  const next = planWorkbookCellShiftCommand(model, 'data', range(2, 3, 1, 2), 'down').workbook.sheets[0]
  assert.deepEqual(next.cells, {}); assert.equal(next.rangeThemeRules.length, 3)
  assert.ok(next.rangeThemeRules.filter(rule => rule.styleRange).every(rule => rule.styleRange.startCol === 0))
})

test('canvas capacity grows only when actual members need it and axis geometry, hides, outline and freeze stay fixed', () => {
  const model = sample(), sheet = model.sheets[0]; sheet.rowCount = 6
  sheet.rowHeights = { 5: 45 }; sheet.rowHiddenSpans = [{ start: 5, end: 5, value: true }]; sheet.frozenRows = 2
  sheet.outline = { groups: [{ id: 'group', axis: 'row', start: 4, end: 5, collapsed: true }], summaryBelow: true, summaryRight: true }
  sheet.cells['4:1'] = cell(4, 1, 'fits')
  const fits = planWorkbookCellShiftCommand(model, 'data', range(1), 'down').workbook.sheets[0]
  assert.equal(fits.rowCount, 6); assert.equal(fits.cells['5:1'].value, 'fits')
  sheet.cells['5:1'] = cell(5, 1, 'needs capacity')
  const grows = planWorkbookCellShiftCommand(model, 'data', range(1), 'down').workbook.sheets[0]
  assert.equal(grows.rowCount, 7); assert.deepEqual(grows.rowHeights, sheet.rowHeights); assert.deepEqual(grows.rowHiddenSpans, sheet.rowHiddenSpans)
  assert.deepEqual(grows.outline, sheet.outline); assert.equal(grows.frozenRows, 2)
})

test('Office edges reject real overflow and malformed sources while allowing an empty last-cell insertion', () => {
  const model = sample(), sheet = model.sheets[0]; sheet.rowCount = 1_048_576; sheet.columnCount = 16_384
  assert.equal(planWorkbookCellShiftCommand(model, 'data', range(0, 0, 16_383), 'right').workbook.sheets[0].columnCount, 16_384)
  sheet.cells['0:16383'] = cell(0, 16_383, 'edge'); const before = cloneModel(model)
  assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(0, 0, 16_383), 'right'), /上限/); assert.deepEqual(model, before)
  const invalid = sample(); invalid.sheets[0].cells.wrong = cell(1, 1, 'bad key')
  assert.throws(() => planWorkbookCellShiftCommand(invalid, 'data', range(1), 'down'), /来源/)
  invalid.sheets[0].cells = {}; invalid.sheets[0].rowCount = NaN
  assert.throws(() => planWorkbookCellShiftCommand(invalid, 'data', range(1), 'down'), /容量/)
  assert.throws(() => assertSheetCellShiftEdit({ axis: 'row', kind: 'insert', start: 1, count: 1, bandStart: -1, bandEnd: 2 }))
})

test('resource-only chart sources, pivot output and slicer anchors grow canvas capacity without adding authored cells', () => {
  const model = sample(), sheet = model.sheets[0]; sheet.rowCount = 12; sheet.columnCount = 8
  model.charts = [chart('source', range(0, 11, 1, 2))]
  model.pivotTables = [pivot({ source: { unitId: model.id, subUnitId: 'report', sheetName: 'Report', range: range(0, 5, 0, 2) },
    anchor: { unitId: model.id, subUnitId: 'data', row: 10, col: 5 }, outputRange: range(10, 11, 5, 6), generatedCellKeys: ['10:5', '11:6'] })]
  model.slicers = [{ id: 'slicer', unitId: model.id, hostSheetId: 'data', pivotTableId: 'pivot', sourceFieldId: 'category', type: 'slicer', caption: 'Field', name: 'Field',
    position: { row: 9, column: 5 }, size: { rows: 2, columns: 2 }, revision: 1,
    drawing: { anchorType: 'both', from: { row: 9, column: 5, offsetX: 0, offsetY: 0 }, to: { row: 11, column: 6, offsetX: 0, offsetY: 0 } } }]
  const chartSource = planWorkbookCellShiftCommand(model, 'data', range(1, 1, 1, 2), 'down').workbook
  assert.equal(chartSource.sheets[0].rowCount, 13); assert.deepEqual(chartSource.sheets[0].cells, {})
  const output = planWorkbookCellShiftCommand(model, 'data', range(8, 8, 5, 6), 'down').workbook
  assert.equal(output.sheets[0].rowCount, 13); assert.equal(output.pivotTables[0].outputRange.endRow, 12)
  assert.equal(output.slicers[0].drawing.to.row, 12); assert.deepEqual(output.sheets[0].cells, {})
  const before = cloneModel(model); model.pivotTables[0].generatedCellKeys = ['1048576:6']
  assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(1, 1, 1, 2), 'down'), /范围|上限|坐标/)
  model.pivotTables[0].generatedCellKeys = before.pivotTables[0].generatedCellKeys; assert.deepEqual(model, before)
})

test('vertical confirmation rejects filtered tail rows from sheet and table filters but permits manual hidden rows and filters above the start', () => {
  const model = sample(), sheet = model.sheets[0], edit = { axis: 'row', kind: 'insert', start: 2, count: 1, bandStart: 1, bandEnd: 2 }
  sheet.autoFilter = { range: range(0, 6, 0, 3), columns: {}, filteredRows: [5] }
  assert.throws(() => assertSheetCellShiftConfirmationAllowed(model, 'data', edit), /筛选隐藏行/)
  assert.doesNotThrow(() => assertSheetCellShiftConfirmationAllowed(model, 'data', { ...edit, axis: 'column' }))
  sheet.autoFilter.filteredRows = [1]; sheet.rowHiddenSpans = [{ start: 5, end: 5, value: true }]
  assert.doesNotThrow(() => assertSheetCellShiftConfirmationAllowed(model, 'data', edit))
  model.tables = [table('filtered', range(0, 6, 5, 7), { autoFilter: { range: range(0, 6, 5, 7), columns: {}, filteredRows: [6] } })]
  assert.throws(() => assertSheetCellShiftConfirmationAllowed(model, 'data', { ...edit, kind: 'delete' }), /筛选隐藏行/)
})

test('coordinate inversion distinguishes inserted blanks and exhausted deletion bands without clipping physical ranges', () => {
  const edit = { axis: 'row', kind: 'insert', start: 2, count: 2, bandStart: 1, bandEnd: 2 }
  assert.deepEqual(mapSheetCellShiftPosition(3, 1, edit), { row: 5, col: 1 }); assert.equal(inverseSheetCellShiftPosition(2, 1, edit), undefined)
  assert.deepEqual(inverseSheetCellShiftPosition(5, 1, edit), { row: 3, col: 1 }); assert.deepEqual(mapSheetCellShiftPosition(3, 3, edit), { row: 3, col: 3 })
  assert.equal(mapSheetCellShiftPosition(3, 1, { ...edit, kind: 'delete' }), undefined)
  const pieces = transformSheetCellShiftRange(range(0, 5, 0, 3), edit)
  assert.equal(pieces.length, 3); assert.ok(pieces.some(value => value.startCol === 1 && value.endRow === 7))
  assert.throws(() => transformSheetCellShiftRange(range(1_048_575), edit), /上限/)
})

test('quoted, external, 3D, whole-axis and structured tokens share one scanner and preserve the perpendicular-span rule', () => {
  const edit = { axis: 'column', kind: 'insert', start: 1, count: 1, bandStart: 0, bandEnd: 1 }
  const formula = '=SUM(Data!$B$1:C2)+"B1"+Table1[B1]+\'[Book.xlsx]Data\'!B1+Data:Report!B1+SUM(B:B)+SUM(1:2)'
  assert.equal(transformFormulaReferenceAreasForCellShift(formula, edit, 'Report', 'Data'),
    '=SUM(Data!$C$1:D2)+"B1"+Table1[B1]+\'[Book.xlsx]Data\'!B1+Data:Report!B1+SUM(B:B)+SUM(1:2)')
  assert.equal(transformFormulaReferenceAreasForCellShift('=SUM(B1:D3)+SUM(B1:D2)', edit, 'Data', 'Data'), '=SUM(B1:D3)+SUM(C1:E2)')
  const full = { ...edit, bandEnd: 1_048_575 }
  assert.equal(transformFormulaReferenceAreasForCellShift('=SUM(B:B)', full, 'Data', 'Data'), '=SUM(C:C)')
})

test('deletion contracts surviving references and produces REF only for exhausted referenced members', () => {
  const edit = { axis: 'row', kind: 'delete', start: 1, count: 2, bandStart: 1, bandEnd: 2 }
  assert.equal(transformFormulaReferenceAreasForCellShift('=B2+B4+SUM(B1:C5)+SUM(A1:D5)', edit, 'Data', 'Data'),
    '=#REF!+B2+SUM(B1:C3)+SUM(A1:D5)')
  assert.equal(transformFormulaReferenceAreasForCellShift('=SUM(B:B)+SUM(2:3)', edit, 'Data', 'Data'), '=SUM(B:B)+SUM(2:3)')
})

test('cross-sheet formulas, sparkline sources, internal rich hyperlinks and scoped names update with their dependency', () => {
  const model = sample(), report = model.sheets[1]
  report.cells['0:0'] = cell(0, 0, 'dependent', { formula: '=Data!B2+Report!B2', hyperlink: '#Data!B2',
    sparkline: { type: 'line', sourceRange: 'Data!B2:B4' }, richText: [{ text: 'jump', hyperlink: '#Data!B2' }] })
  model.metadata = { definedNames: [{ name: 'Source', refersTo: 'Data!$B$2' }], localDefinedNames: [
    { name: 'Local', sheet: 'Report', refersTo: 'B2' }, { name: 'Remote', sheet: 'Report', refersTo: 'Data!B2' },
  ] }
  const next = planWorkbookCellShiftCommand(model, 'data', range(1), 'down').workbook
  assert.equal(next.sheets[1].cells['0:0'].formula, '=Data!B3+Report!B2')
  assert.equal(next.sheets[1].cells['0:0'].sparkline.sourceRange, 'Data!B3:B5')
  assert.equal(next.sheets[1].cells['0:0'].hyperlink, '#Data!B3'); assert.equal(next.sheets[1].cells['0:0'].richText[0].hyperlink, '#Data!B3')
  assert.equal(next.metadata.definedNames[0].refersTo, 'Data!$B$3'); assert.equal(next.metadata.localDefinedNames[0].refersTo, 'B2')
  assert.equal(next.metadata.localDefinedNames[1].refersTo, 'Data!B3')
})

test('conditional and validation fragments rebase relative origins while list and checkbox labels stay literal', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.conditionalFormats = [cf('cf', [range(0, 5, 0, 3)])]
  sheet.dataValidations = [validation('list', [range(0, 5, 0, 3)]), validation('checkbox', [range(0, 5, 0, 3)], { type: 'checkbox', formula1: 'A1', formula2: 'B2' })]
  const next = planWorkbookCellShiftCommand(model, 'data', range(2, 2, 1, 2), 'down').workbook.sheets[0]
  assert.equal(next.conditionalFormats.length, 3); assert.equal(new Set(next.conditionalFormats.map(rule => rule.id)).size, 3)
  assert.ok(next.conditionalFormats.every(rule => rule.formulaAnchor && rule.formula1))
  assert.equal(next.dataValidations.length, 6)
  assert.ok(next.dataValidations.filter(rule => rule.type === 'list').every(rule => rule.formula1 === 'A1,B2'))
  assert.ok(next.dataValidations.filter(rule => rule.type === 'checkbox').every(rule => rule.formula1 === 'A1' && rule.formula2 === 'B2'))
})

test('fragment IDs and anchors do not broaden permissions into unchanged protected members outside the band', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.conditionalFormats = [cf('cf', [range(0, 5, 0, 3)])]; sheet.dataValidations = [validation('dv', [range(0, 5, 0, 3)])]
  sheet.rangeProtectionRules = [{ id: 'retained', ranges: [range(5, 5, 0)], allowedUserIds: ['owner'] }]
  const plan = planWorkbookCellShiftCommand(model, 'data', range(2, 2, 1, 2), 'down')
  assert.equal(plan.workbook.sheets[0].conditionalFormats.length, 3)
  assert.ok(plan.affected.data.every(value => value.startCol > 0 || value.endCol < 0 || value.endRow < 5 || value.startRow > 5))
})

test('predicate precedence differences cover their real intersection and repeated predicates retain per-cell order', () => {
  const model = sample(), before = model.sheets[0]
  const a = validation('a', [range(0, 2, 0, 1)], { formula1: 'a' }), b = validation('b', [range(1, 3, 1, 2)], { formula1: 'b' })
  before.dataValidations = [a, b]; const after = cloneModel(before); after.dataValidations = [b, a]
  assert.deepEqual(sheetClipboardChangedRuleRanges(before, after), [range(1, 2, 1)])
  before.rangeProtectionRules = [{ id: 'overlap', ranges: [range(1, 1, 1)], allowedUserIds: ['owner'] }]
  const candidate = cloneModel(model); candidate.sheets[0].dataValidations = [b, a]
  assert.throws(() => assertSheetClipboardWorkbookChangesEditable(model, candidate, 'local-user'), /保护/)
  const repeated = cloneModel(before); repeated.dataValidations = [a, b, { ...a, id: 'a2' }, { ...b, id: 'b2' }]
  const reordered = cloneModel(repeated); reordered.dataValidations = [a, b, { ...b, id: 'b2' }, { ...a, id: 'a2' }]
  assert.deepEqual(sheetClipboardChangedRuleRanges(repeated, reordered), [range(1, 2, 1)])
  const geometry = cloneModel(before); geometry.dataValidations = [validation('a1', [range(0, 0, 1)], { formula1: 'a' }), b,
    validation('a2', [range(3, 3, 1)], { formula1: 'a' })]
  const moved = cloneModel(geometry); [moved.dataValidations[0].ranges, moved.dataValidations[2].ranges] = [moved.dataValidations[2].ranges, moved.dataValidations[0].ranges]
  assert.deepEqual(sheetClipboardChangedRuleRanges(geometry, moved), [range(3, 3, 1)])
})

test('owners can shift their protected bands and remote protected dependencies reject without touching the source', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.rangeProtectionRules = [{ id: 'owned', ranges: [range(8)], allowedUserIds: ['owner'], ownerUserId: 'owner' }]
  assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(1), 'down'), /保护/)
  assert.deepEqual(planWorkbookCellShiftCommand(model, 'data', range(1), 'down', 'owner').workbook.sheets[0].rangeProtectionRules[0].ranges, [range(9)])
  delete sheet.rangeProtectionRules; model.sheets[1].cells['0:0'] = cell(0, 0, null, { formula: '=Data!B2' })
  model.sheets[1].rangeProtectionRules = [{ id: 'remote', ranges: [range(0, 0, 0)], allowedUserIds: ['owner'] }]
  const before = cloneModel(model); assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(1), 'down'), /保护/); assert.deepEqual(model, before)
})

test('partial merges are fully prepared before confirmation and require permission across the complete dismantled range', () => {
  const model = sample(), sheet = model.sheets[0]; sheet.mergedRanges = [range(1, 2, 0, 2)]
  sheet.cells['1:0'] = cell(1, 0, 'master'); const before = cloneModel(model)
  const pending = planWorkbookCellShiftCommand(model, 'data', range(1, 1, 1), 'down')
  assert.equal(pending.requiresConfirmation, true); assert.deepEqual(pending.mergeRanges, sheet.mergedRanges)
  assert.deepEqual(pending.workbook.sheets[0].mergedRanges, []); assert.deepEqual(model, before)
  assert.equal(planWorkbookCellShiftCommand(model, 'data', range(1, 1, 1), 'down', 'local-user', true).requiresConfirmation, false)
  sheet.rangeProtectionRules = [{ id: 'master', ranges: [range(1, 1, 0)], allowedUserIds: ['owner'] }]
  assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(1, 1, 1), 'down'), /保护/)
})

test('fully moved and completely deleted merges need no dismantle prompt', () => {
  const model = sample(), sheet = model.sheets[0]; sheet.mergedRanges = [range(2, 3, 1, 2)]
  const moved = planWorkbookCellShiftCommand(model, 'data', range(1, 1, 1, 2), 'down')
  assert.equal(moved.requiresConfirmation, false); assert.deepEqual(moved.workbook.sheets[0].mergedRanges, [range(3, 4, 1, 2)])
  const removed = planWorkbookCellShiftCommand(model, 'data', range(2, 3, 1, 2), 'up')
  assert.equal(removed.requiresConfirmation, false); assert.deepEqual(removed.workbook.sheets[0].mergedRanges, [])
})

test('continuous filter and table contracts update as a whole and reject a partial split atomically', () => {
  const model = sample(), sheet = model.sheets[0], source = range(0, 5, 0, 3)
  sheet.autoFilter = { range: source, columns: { 1: { kind: 'values', selectedItems: ['x'] } }, filteredRows: [4] }
  const before = cloneModel(model); assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(2, 2, 1, 2), 'down'), /连续/); assert.deepEqual(model, before)
  const whole = planWorkbookCellShiftCommand(model, 'data', range(2, 2, 0, 3), 'down').workbook.sheets[0].autoFilter
  assert.deepEqual(whole.range, range(0, 6, 0, 3)); assert.deepEqual(whole.filteredRows, [5])
  delete sheet.autoFilter; model.tables = [table('table', source)]
  assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(2, 2, 1, 2), 'down'), /连续/)
  assert.deepEqual(planWorkbookCellShiftCommand(model, 'data', range(2, 2, 0, 3), 'down').workbook.tables[0].range, range(0, 6, 0, 3))
})

test('chart source fields and revisions follow touched dependencies while unrelated charts keep their revisions', () => {
  const model = sample(); model.charts = [chart('changed', range(0, 5, 1, 2)), chart('untouched', range(0, 5, 5, 6))]
  const next = planWorkbookCellShiftCommand(model, 'data', range(0, 5, 2), 'right').workbook
  assert.deepEqual(next.charts[0].source.range, range(0, 5, 1, 3)); assert.deepEqual(next.charts[0].valueFields, [2])
  assert.equal(next.charts[0].revision, 3); assert.equal(next.charts[1].revision, 2)
  model.sheets[1].rangeProtectionRules = [{ id: 'chart-host', ranges: [range(0, 8, 0, 5)], allowedUserIds: ['owner'] }]
  assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(0, 5, 2), 'right'), /保护/)
})

test('pivot fields, generated keys, output rectangles and drill-down invalidation use the same structure edit', () => {
  const model = sample(); model.pivotTables = [pivot()]
  const source = planWorkbookCellShiftCommand(model, 'data', range(0, 5, 2), 'right').workbook.pivotTables[0]
  assert.equal(source.fields[1].sourceColumn, 3); assert.equal(source.revision, 4); assert.deepEqual(source.drillDownMap, {})
  model.pivotTables = [pivot({ source: { unitId: model.id, subUnitId: 'report', sheetName: 'Report', range: range(0, 5, 0, 2) },
    anchor: { unitId: model.id, subUnitId: 'data', row: 10, col: 5 }, outputRange: range(10, 11, 5, 6), generatedCellKeys: ['10:5', '11:6'] })]
  const output = planWorkbookCellShiftCommand(model, 'data', range(8, 8, 5, 6), 'down').workbook.pivotTables[0]
  assert.deepEqual(output.anchor, { unitId: model.id, subUnitId: 'data', row: 11, col: 5 })
  assert.deepEqual(output.outputRange, range(11, 12, 5, 6)); assert.deepEqual(output.generatedCellKeys, ['11:5', '12:6'])
  assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(8, 8, 6), 'down'), /连续/)
})

test('slicer drawing anchors and table field dependencies move or disappear with their owning resources', () => {
  const model = sample(); model.tables = [table('table', range(0, 5, 0, 2))]
  model.slicers = [{ id: 'slicer', unitId: model.id, hostSheetId: 'data', tableId: 'table', sourceFieldId: 'table-column-1', type: 'slicer', caption: 'Field', name: 'Field',
    position: { row: 8, column: 5 }, size: { rows: 3, columns: 2 }, revision: 1,
    drawing: { anchorType: 'both', from: { row: 8, column: 5, offsetX: 3, offsetY: 4 }, to: { row: 11, column: 6, offsetX: 7, offsetY: 8 } } }]
  const moved = planWorkbookCellShiftCommand(model, 'data', range(7, 7, 5, 6), 'down').workbook.slicers[0]
  assert.equal(moved.position.row, 9); assert.equal(moved.drawing.from.row, 9); assert.equal(moved.drawing.to.row, 12)
  assert.equal(moved.drawing.from.offsetY, 4); assert.ok(moved.layout.height > 0); assert.equal(moved.revision, 2)
  const removed = planWorkbookCellShiftCommand(model, 'data', range(0, 5, 1), 'left').workbook
  assert.equal(removed.tables[0].columns.length, 2); assert.deepEqual(removed.slicers, [])
})

test('stop validation rejects the complete candidate before a partial-merge confirmation is published', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.mergedRanges = [range(1, 2, 0, 1)]; sheet.cells['1:1'] = cell(1, 1, -1)
  sheet.dataValidations = [{ id: 'positive', ranges: [range(1, 4)], type: 'whole', operator: 'greaterThanOrEqual', formula1: '0', allowBlank: true,
    showErrorMessage: true, errorStyle: 'stop', error: '必须非负' }]
  const before = cloneModel(model); assert.throws(() => planWorkbookCellShiftCommand(model, 'data', range(1), 'down'), /非负|验证/); assert.deepEqual(model, before)
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
function runtime(t, model = sample()) {
  const studio = new OfficeStudio({ workbook: model, sheetTextMeasure: text => [...text].length * 8 })
  t.onTestFinished(() => studio.dispose()); const api = FOffice.newAPI(studio), book = api.getActiveWorkbook()
  return { studio, api, book }
}
const depth = (studio, id = studio.activeWorkbookId) => studio.editor.history.units[id]?.undo.length ?? 0

test('Facade insertCells and deleteCells honor fixed Dimension values and one content history entry per operation', t => {
  const model = sample(); model.sheets[0].cells['1:1'] = cell(1, 1, 'move')
  const { studio, api, book } = runtime(t, model), sheet = book.getActiveSheet(), count = depth(studio)
  assert.equal(api.Enum.Dimension.COLUMNS, 0); assert.equal(api.Enum.Dimension.ROWS, 1)
  assert.equal(sheet.getRange('B2').insertCells(api.Enum.Dimension.COLUMNS), undefined)
  assert.equal(sheet.getRange('C2').getValue(), 'move'); assert.equal(depth(studio), count + 1)
  sheet.getRange('B2').deleteCells(api.Enum.Dimension.COLUMNS); assert.equal(sheet.getRange('B2').getValue(), 'move')
  sheet.getRange('B2').insertCells(api.Enum.Dimension.ROWS); assert.equal(sheet.getRange('B3').getValue(), 'move')
  sheet.getRange('B2').deleteCells(api.Enum.Dimension.ROWS); assert.equal(sheet.getRange('B2').getValue(), 'move')
  assert.equal(depth(studio), count + 4)
  assert.throws(() => sheet.getRange('B2').insertCells(2), /方向/); assert.equal(depth(studio), count + 4)
})

test('registered commands prioritize explicit primary ranges and Facade can edit an inactive worksheet', t => {
  const model = sample(); model.sheets[0].cells['1:1'] = cell(1, 1, 'active'); model.sheets[1].cells['1:1'] = cell(1, 1, 'inactive')
  const { studio, book } = runtime(t, model)
  studio.executeCommandSync(SHEET_CELL_SHIFT_COMMANDS.down, { unitId: book.getId(), subUnitId: 'data', range: apiRange(range(1)), ranges: [range(1), range(5, 5, 3)] })
  assert.equal(book.getActiveSheet().getRange('B3').getValue(), 'active'); assert.equal(book.getActiveSheet().getRange('B2').getValue(), null)
  book.getSheetBySheetId('report').getRange('B2').insertCells(SheetCellShiftDimensions.ROWS)
  assert.equal(book.getSheetBySheetId('report').getRange('B3').getValue(), 'inactive'); assert.equal(book.getActiveSheet().getSheetId(), 'data')
})

test('implicit command dispatch follows the active non-last selection and records before/after selection in content history', t => {
  const model = sample(); model.sheets[0].cells['1:1'] = cell(1, 1, 'primary'); model.sheets[0].cells['5:3'] = cell(5, 3, 'last')
  const { studio, book } = runtime(t, model)
  studio.setSelections(book.getId(), 'data', [apiRange(range(1)), apiRange(range(5, 5, 3))], { activeRangeIndex: 0 })
  const before = cloneModel(studio.getSelectionState(book.getId(), 'data'))
  assert.equal(studio.executeCommandSync(SHEET_CELL_SHIFT_COMMANDS.down), true)
  assert.equal(book.getActiveSheet().getRange('B3').getValue(), 'primary'); assert.equal(book.getActiveSheet().getRange('D6').getValue(), 'last')
  assert.equal(studio.getSelectionState(book.getId(), 'data').activeRangeIndex, 0)
  book.undo(); assert.deepEqual(studio.getSelectionState(book.getId(), 'data'), before)
  book.redo(); assert.equal(book.getActiveSheet().getRange('B3').getValue(), 'primary')
})

test('merge confirmation captures content/history versions and creates no history until the complete request is accepted', t => {
  const model = sample(); model.sheets[0].mergedRanges = [range(1, 2, 0, 2)]; model.sheets[0].cells['1:0'] = cell(1, 0, 'master')
  const { studio, book } = runtime(t, model), requests = [], structures = [], count = depth(studio)
  studio.events.addEvent('SheetCellShiftConfirmationRequested', event => requests.push(cloneModel(event)))
  studio.events.addEvent('SheetStructureChanged', event => structures.push(cloneModel(event)))
  assert.equal(studio.executeCommandSync(SHEET_CELL_SHIFT_CONFIRM_COMMANDS.down, { range: range(1) }), false)
  const request = requests[0]; assert.equal(depth(studio), count); assert.deepEqual(request.mergeRanges, [range(1, 2, 0, 2)])
  assert.equal(request.input.expectedRevision, studio.getWorkbookRevision(book.getId()))
  assert.equal(request.input.expectedHistoryGeneration, studio.editor.history.captureSavepoint(book.getId()).generation)
  assert.equal(studio.executeCommandSync(request.commandId, { ...request.input, dismantlePartialMerges: true }), true)
  assert.equal(depth(studio), count + 1); assert.equal(structures[0].edits[0].bandStart, 1)
  book.undo(); assert.deepEqual(studio.getWorkbookModel(book.getId()).sheets[0].mergedRanges, [range(1, 2, 0, 2)])
})

test('stale revisions or history generations reject confirmation without publishing or replacing newer content', t => {
  const model = sample(); model.sheets[0].mergedRanges = [range(1, 2, 0, 2)]
  const { studio, book } = runtime(t, model), requests = []
  studio.events.addEvent('SheetCellShiftConfirmationRequested', event => requests.push(cloneModel(event)))
  studio.executeCommandSync(SHEET_CELL_SHIFT_COMMANDS.down, { range: range(1) }); const request = requests[0]
  book.getActiveSheet().getRange('A10').setValue('newer'); const before = cloneModel(book.save()), count = depth(studio)
  assert.throws(() => studio.executeCommandSync(request.commandId, { ...request.input, dismantlePartialMerges: true }), /已更改/)
  assert.deepEqual(book.save(), before); assert.equal(depth(studio), count)
  assert.throws(() => studio.executeCommandSync(request.commandId, { ...request.input, dismantlePartialMerges: true,
    expectedRevision: studio.getWorkbookRevision(book.getId()), expectedHistoryGeneration: studio.editor.history.captureSavepoint(book.getId()).generation + 1 }), /已更改/)
  assert.equal(depth(studio), count)
})

test('a UI host constraint rejects inactive targets while the same explicit command remains valid without that constraint', t => {
  const model = sample(); model.sheets[1].cells['1:1'] = cell(1, 1, 'report')
  const { studio, book } = runtime(t, model), request = { unitId: book.getId(), sheetId: 'report', range: range(1) }
  assert.throws(() => studio.executeCommandSync(SHEET_CELL_SHIFT_COMMANDS.down, { ...request, host: { unitId: book.getId(), sheetId: 'report' } }), /宿主/)
  assert.equal(depth(studio), 0); assert.equal(studio.executeCommandSync(SHEET_CELL_SHIFT_COMMANDS.down, request), true)
  assert.equal(book.getSheetBySheetId('report').getRange('B3').getValue(), 'report')
})

test('registered vertical confirm commands reject filtered tails before merge events or history while explicit core commands retain their range contract', t => {
  const model = sample(); model.sheets[0].autoFilter = { range: range(0, 6, 0, 3), columns: {}, filteredRows: [5] }
  model.sheets[0].cells['1:1'] = cell(1, 1, 'move')
  const { studio, book } = runtime(t, model), requests = [], before = cloneModel(book.save()), count = depth(studio)
  studio.events.addEvent('SheetCellShiftConfirmationRequested', event => requests.push(event))
  for (const id of [SHEET_CELL_SHIFT_CONFIRM_COMMANDS.down, SHEET_CELL_SHIFT_CONFIRM_COMMANDS.up]) {
    assert.throws(() => studio.executeCommandSync(id, { range: range(1, 1, 0, 3) }), /筛选隐藏行/)
    assert.deepEqual(book.save(), before); assert.equal(depth(studio), count)
  }
  assert.deepEqual(requests, [])
  assert.equal(studio.executeCommandSync(SHEET_CELL_SHIFT_COMMANDS.down, { range: range(1, 1, 0, 3) }), true)
  assert.equal(book.getActiveSheet().getRange('B3').getValue(), 'move')
})

test('observer reentry retains the committed shift and its later edit as two separate content records', t => {
  const model = sample(); model.sheets[0].cells['1:1'] = cell(1, 1, 'move')
  const { studio, book } = runtime(t, model); let reentered = false
  studio.events.addEvent('SheetStructureChanged', () => {
    if (reentered) return; reentered = true; book.getActiveSheet().getRange('D10').setValue('observer')
  })
  studio.executeCommandSync(SHEET_CELL_SHIFT_COMMANDS.down, { range: range(1) })
  assert.equal(book.getActiveSheet().getRange('B3').getValue(), 'move'); assert.equal(book.getActiveSheet().getRange('D10').getValue(), 'observer')
  assert.equal(depth(studio), 2); book.undo(); assert.equal(book.getActiveSheet().getRange('D10').getValue(), null)
  assert.equal(book.getActiveSheet().getRange('B3').getValue(), 'move'); book.undo(); assert.equal(book.getActiveSheet().getRange('B2').getValue(), 'move')
})

test('internal save/reopen retains shifted author sources, formulas, rules and history without materializing virtual theme cells', t => {
  const model = sample(), sheet = model.sheets[0]
  sheet.cells['1:1'] = cell(1, 1, 'saved', { authoredStyle: { font: { italic: true } }, metadata: { note: 'kept' } })
  model.sheets[1].cells['0:0'] = cell(0, 0, null, { formula: '=Data!B2' })
  sheet.rangeThemeRules = [{ id: 'theme', themeName: 'default', range: range(0, 5, 0, 3) }]
  sheet.dataValidations = [validation('list', [range(0, 5, 0, 3)])]
  const { studio, api, book } = runtime(t, model), before = cloneModel(book.save())
  studio.executeCommandSync(SHEET_CELL_SHIFT_COMMANDS.down, { range: range(1, 1, 1, 2) })
  const saved = cloneModel(book.save()), reopened = api.createWorkbook({ ...saved, id: 'shift-reopened' })
  assert.equal(reopened.getActiveSheet().getRange('B3').getValue(), 'saved'); assert.equal(reopened.getSheetBySheetId('report').getRange('A1').getFormula(), '=Data!B3')
  assert.deepEqual(studio.getWorkbookModel(reopened.getId()).sheets[0].rangeThemeRules, saved.sheets[0].rangeThemeRules)
  assert.equal(Object.keys(saved.sheets[0].cells).length, 2); book.undo(); assert.deepEqual(book.save().sheets, before.sheets)
})
