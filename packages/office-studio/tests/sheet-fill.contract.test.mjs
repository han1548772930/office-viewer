import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'vitest'
import { createSheetFillPlan, detectSheetFillDownEndRow } from '../src/features/sheet-workbench/sheet-fill.ts'
import { fillPreviewRange, hitSelectionFillHandle, SHEET_SELECTION_STROKE_WIDTH } from '../src/rendering/spreadsheet/selection.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'

test('selection fill handle has an independent hit target above the cell beneath it', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'sheet', name: 'Sheet', cells: {}, rowCount: 20, columnCount: 10,
    defaultRowHeight: 24, defaultColumnWidth: 80,
  })
  const viewport = { width: 900, height: 500, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  const range = { startRow: 2, endRow: 6, startColumn: 2, endColumn: 5 }
  const endpoint = geometry.cellRect(range.endRow, range.endColumn, viewport, false)
  assert.deepEqual(
    hitSelectionFillHandle(geometry, [range], viewport, endpoint.x + endpoint.width, endpoint.y + endpoint.height),
    range,
  )
  assert.deepEqual(
    hitSelectionFillHandle(geometry, [range], viewport, endpoint.x + endpoint.width + 5, endpoint.y + endpoint.height + 5),
    range,
  )
  assert.equal(hitSelectionFillHandle(geometry, [range], viewport, endpoint.x + endpoint.width + 8, endpoint.y + endpoint.height), undefined)
})

test('fill drag repeats selected cells and shifts relative formula references per destination', () => {
  const source = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }
  const sourceCells = [[{ value: 'alpha' }], [{ formula: '=A1+$B$2', value: '', bold: true }]]
  const plan = createSheetFillPlan(source, { row: 4, column: 0 }, sourceCells)
  assert.deepEqual(plan.range, { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 })
  assert.deepEqual(plan.cellPatches.map(({ row, column, cell }) => ({ row, column, formula: cell.formula, value: cell.value, bold: cell.bold })), [
    { row: 2, column: 0, formula: undefined, value: 'alpha', bold: undefined },
    { row: 3, column: 0, formula: '=A3+$B$2', value: '', bold: true },
    { row: 4, column: 0, formula: undefined, value: 'alpha', bold: undefined },
  ])
})

test('series mode extends numeric values and date-serial values while copy mode repeats the source', () => {
  const source = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }
  const sourceCells = [[{ value: '2', rawValue: 2 }], [{ value: '5', rawValue: 5 }]]
  const series = createSheetFillPlan(source, { row: 4, column: 0 }, sourceCells)
  assert.deepEqual(series.cellPatches.map(({ cell }) => cell.rawValue), [8, 11, 14])

  const dates = createSheetFillPlan(source, { row: 3, column: 0 }, [
    [{ value: '45291', rawValue: 45291, numberFormat: 'date' }],
    [{ value: '45292', rawValue: 45292, numberFormat: 'date' }],
  ])
  assert.deepEqual(dates.cellPatches.map(({ cell }) => cell.rawValue), [45293, 45294])

  const copy = createSheetFillPlan(source, { row: 3, column: 0 }, sourceCells, { applyType: 'COPY' })
  assert.deepEqual(copy.cellPatches.map(({ cell }) => cell.rawValue), [2, 5])
})

test('series mode repeats irregular source patterns and keeps formula translation', () => {
  const irregular = createSheetFillPlan(
    { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    { row: 5, column: 0 },
    [[{ value: '1', rawValue: 1 }], [{ value: '3', rawValue: 3 }], [{ value: '8', rawValue: 8 }]],
  )
  assert.deepEqual(irregular.cellPatches.map(({ cell }) => cell.rawValue), [1, 3, 8])

  const formulas = createSheetFillPlan(
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    { row: 3, column: 0 },
    [[{ formula: '=A1+1', value: '' }], [{ formula: '=A2+1', value: '' }]],
  )
  assert.deepEqual(formulas.cellPatches.map(({ cell }) => cell.formula), ['=A3+1', '=A4+1'])
})

test('only-format copies source formatting and preserves destination values', () => {
  const plan = createSheetFillPlan(
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { row: 1, column: 0 },
    [[{ value: 'source', bold: true, background: '#ff0000', border: { bottom: { style: 'solid' } } }]],
    { applyType: 'ONLY_FORMAT', targetCells: [[{ value: 'source' }], [{ value: 'keep', formula: '=B2', italic: true }]] },
  )
  assert.deepEqual(plan.cellPatches.map(({ cell }) => ({ value: cell.value, formula: cell.formula, bold: cell.bold, italic: cell.italic, background: cell.background })), [
    { value: 'keep', formula: '=B2', bold: true, italic: true, background: '#ff0000' },
  ])
})

test('no-format copies source values and formulas while preserving destination formatting', () => {
  const plan = createSheetFillPlan(
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { row: 1, column: 0 },
    [[{ formula: '=A1+$B$1', value: '', bold: true, background: '#ff0000' }]],
    { applyType: 'NO_FORMAT', targetCells: [[{ value: 'source' }], [{ value: 'keep', italic: true, background: '#00ff00' }]] },
  )
  assert.deepEqual(plan.cellPatches.map(({ cell }) => ({ value: cell.value, formula: cell.formula, bold: cell.bold, italic: cell.italic, background: cell.background })), [
    { value: '', formula: '=A2+$B$1', bold: undefined, italic: true, background: '#00ff00' },
  ])
})

test('double-click fill uses left adjacent data first, falls back right, and stops before filled cells', () => {
  const source = { startRow: 0, endRow: 1, startColumn: 1, endColumn: 1 }
  const cells = {
    A3: { value: 'left' }, A4: { value: 'left' }, A5: { value: 'left' }, A6: { value: 'left' },
    B6: { value: 'already filled' },
    C3: { value: 'right' }, C4: { value: 'right' }, C5: { value: 'right' },
  }
  assert.equal(detectSheetFillDownEndRow(source, { cells, rowCount: 10, columnCount: 5 }), 4)
  assert.equal(detectSheetFillDownEndRow(source, { cells: { C3: { value: 'right' }, C4: { value: 'right' }, C5: { value: 'right' } }, rowCount: 10, columnCount: 5 }), 4)
  assert.equal(detectSheetFillDownEndRow(source, { cells: {}, rowCount: 10, columnCount: 5 }), undefined)
})

test('fill drag in any direction translates formulas and does not rewrite the source selection', () => {
  const source = { startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 }
  const plan = createSheetFillPlan(source, { row: 0, column: 2 }, [[{ formula: '=C3', value: '' }]])
  assert.deepEqual(plan.range, { startRow: 0, endRow: 2, startColumn: 2, endColumn: 2 })
  assert.deepEqual(plan.cellPatches.map(({ row, cell }) => [row, cell.formula]), [[0, '=C1'], [1, '=C2']])
})

test('fill preview expands the source selection toward the dragged cell in every direction', () => {
  const source = { startRow: 2, endRow: 3, startColumn: 2, endColumn: 4 }
  assert.deepEqual(fillPreviewRange(source, { row: 6, column: 4 }), {
    startRow: 2, endRow: 6, startColumn: 2, endColumn: 4,
  })
  assert.deepEqual(fillPreviewRange(source, { row: 3, column: 0 }), {
    startRow: 2, endRow: 3, startColumn: 0, endColumn: 4,
  })
  assert.deepEqual(fillPreviewRange(source, { row: 0, column: 1 }), {
    startRow: 0, endRow: 3, startColumn: 1, endColumn: 4,
  })
  assert.equal(fillPreviewRange(source, { row: 2, column: 3 }), undefined)
})

test('scene draws the fill preview with the 1px selection stroke at every zoom level', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'sheet', name: 'Sheet', cells: {}, rowCount: 20, columnCount: 10,
    defaultRowHeight: 24, defaultColumnWidth: 80,
  })
  const preview = { startRow: 2, endRow: 5, startColumn: 1, endColumn: 3 }
  for (const zoom of [0.75, 1, 1.5]) {
    const scene = createSpreadsheetScene(geometry, { width: 900, height: 500, scrollLeft: 0, scrollTop: 0, zoom }, {
      activeCell: '', selectedCells: [], fillPreviewRange: preview,
    })
    const outline = scene.commands.filter(command => command.kind === 'line' && command.color === '#246bfd')
    assert.equal(outline.length, 4)
    assert.ok(outline.every(command => command.width === SHEET_SELECTION_STROKE_WIDTH && command.clip))
  }
})

test('Canvas prioritizes the handle target and commits the fill event after pointer drag', () => {
  const canvas = readFileSync(new URL('../src/components/spreadsheet/SpreadsheetCanvas.vue', import.meta.url), 'utf8')
  const workbench = readFileSync(new URL('../src/components/ExcelWorkbench.vue', import.meta.url), 'utf8')
  assert.match(canvas, /hitSelectionFillHandle\(geometry\.value, props\.selectedRanges/)
  assert.match(canvas, /fillPreviewRange: fillPreview\.value/)
  assert.match(canvas, /fillPreview\.value = fillPreviewRange\(fillDrag\.source, fillDrag\.target\)/)
  assert.match(canvas, /fillPreview\.value = undefined/)
  assert.match(canvas, /emit\('fill-selection', source, target\)/)
  assert.match(canvas, /detectSheetFillDownEndRow\(source, props\.sheet\)/)
  assert.match(canvas, /emit\('fill-selection', source, \{ row: targetRow, column: source\.endColumn \}, 'SERIES'\)/)
  assert.match(workbench, /@fill-selection="fillSelection"/)
  assert.match(workbench, /createSheetFillPlan\(source, target, sourceCells, \{ applyType, targetCells \}\)/)
})
