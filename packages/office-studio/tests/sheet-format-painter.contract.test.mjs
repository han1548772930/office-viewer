import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createFormatPainterPatches } from '../src/features/sheet-workbench/sheet-format-painter.ts'
import { previewCustomNumberFormat, validateCustomNumberFormat } from '../src/features/sheet-workbench/custom-number-format.ts'

const address = (row, column) => `${column}:${row}`

test('format painter tiles source styles and preserves destination values and formulas', () => {
  const source = {
    range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    cells: [[{ value: 'source', bold: true, background: '#ff0', border: { top: { color: '#000' } } }, { value: 'source2', italic: true }]],
  }
  const targets = [{ startRow: 2, endRow: 2, startColumn: 3, endColumn: 6 }]
  const existing = {
    '3:2': { value: 'keep', formula: '=A1', bold: false },
    '4:2': { value: 'also keep', numberFormat: 'date' },
  }
  const patches = createFormatPainterPatches(source, targets, existing, address)

  assert.equal(patches.length, 4)
  assert.equal(patches[0].cell.value, 'keep')
  assert.equal(patches[0].cell.formula, '=A1')
  assert.equal(patches[0].cell.bold, true)
  assert.equal(patches[1].cell.italic, true)
  assert.equal(patches[2].cell.bold, true)
  assert.notEqual(patches[0].cell.border, source.cells[0][0].border)
  assert.notEqual(patches[0].cell.border.top, source.cells[0][0].border.top)
})

test('custom number formats validate Excel sections and show a preview', () => {
  assert.equal(validateCustomNumberFormat('[Red]#,##0.00;[Blue]-#,##0.00;"-"').pattern, '[Red]#,##0.00;[Blue]-#,##0.00;"-"')
  assert.equal(previewCustomNumberFormat('0.00%', 0.125), '12.50%')
  assert.equal(validateCustomNumberFormat('yyyy/mm/dd').pattern, 'yyyy/mm/dd')
})

test('custom number format validation rejects malformed patterns', () => {
  for (const pattern of ['', '0;0;0;0;0', '0"broken', '0[Red', 'text-only']) {
    assert.throws(() => validateCustomNumberFormat(pattern))
  }
})
