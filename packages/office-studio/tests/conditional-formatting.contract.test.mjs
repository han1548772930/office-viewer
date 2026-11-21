import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import {
  addConditionalFormatRule,
  conditionalFormatStyleAt,
  deleteConditionalFormatRule,
  evaluateConditionalFormatRule,
  moveConditionalFormatRule,
  updateConditionalFormatRule,
} from '../src/features/conditional-formatting/conditional-formatting.ts'

const sheet = { id: 'sheet-a', name: 'Sheet1', visible: true, cells: {}, mergedRanges: [], columnWidths: {}, rowHeights: {}, frozenRows: 0, frozenCols: 0, showGridlines: true }
const baseRule = {
  id: 'rule-a', type: 'cellIs', ranges: [{ startRow: 0, startCol: 0, endRow: 2, endCol: 1 }],
  operator: 'between', formula1: '10', formula2: '20', style: { fill: '#ff0000', bold: true }, priority: 1, stopIfTrue: false,
}

test('cell-is comparisons evaluate numeric scalars and exclude text/error values', () => {
  assert.equal(evaluateConditionalFormatRule(baseRule, 10), true)
  assert.equal(evaluateConditionalFormatRule(baseRule, 20), true)
  assert.equal(evaluateConditionalFormatRule(baseRule, 9), false)
  assert.equal(evaluateConditionalFormatRule(baseRule, '15'), false)
  assert.equal(evaluateConditionalFormatRule(baseRule, { kind: 'error', code: '#N/A' }), false)
})

test('rule priority composes distinct properties and stop-if-true halts lower-priority matches', () => {
  const rules = [
    baseRule,
    { ...baseRule, id: 'rule-b', operator: 'greaterThan', formula1: '12', formula2: undefined, style: { fontColor: '#00aa00', bold: false }, priority: 2, stopIfTrue: false },
  ]
  assert.deepEqual(conditionalFormatStyleAt(rules, 0, 0, 15), { fill: '#ff0000', bold: true, fontColor: '#00aa00' })
  assert.deepEqual(conditionalFormatStyleAt([{ ...rules[0], stopIfTrue: true }, rules[1]], 0, 0, 15), { fill: '#ff0000', bold: true })
  assert.deepEqual(conditionalFormatStyleAt(rules, 4, 0, 15), {})
})

test('rule create, update, priority move, delete and validation return immutable sheet snapshots', () => {
  const first = addConditionalFormatRule(sheet, {
    ranges: [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }], operator: 'greaterThan', formula1: '=5', style: { fill: '#aabbcc' },
  }, 'rule-a')
  assert.equal(sheet.conditionalFormats, undefined)
  assert.equal(first.conditionalFormats[0].formula1, '5')
  const second = addConditionalFormatRule(first, {
    ranges: [{ startRow: 2, startCol: 0, endRow: 2, endCol: 0 }], operator: 'equal', formula1: '9', style: { fontColor: '#112233' },
  }, 'rule-b')
  const moved = moveConditionalFormatRule(second, 'rule-b', 0)
  assert.deepEqual(moved.conditionalFormats.map(rule => [rule.id, rule.priority]), [['rule-b', 1], ['rule-a', 2]])
  const updated = updateConditionalFormatRule(moved, 'rule-a', {
    ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }], operator: 'between', formula1: '3', formula2: '7', style: { bold: true }, stopIfTrue: true,
  })
  assert.equal(updated.conditionalFormats[1].stopIfTrue, true)
  assert.deepEqual(deleteConditionalFormatRule(updated, 'rule-b').conditionalFormats.map(rule => [rule.id, rule.priority]), [['rule-a', 1]])
  assert.throws(() => addConditionalFormatRule(sheet, {
    ranges: [{ startRow: 1, startCol: 1, endRow: 0, endCol: 2 }], operator: 'between', formula1: 'x', style: {},
  }, 'invalid'), /valid cell ranges/)
})

test('parser conditional-format payload normalizes into the editable sheet model', () => {
  const workbook = normalizeWorkbook({
    schemaVersion: 1, parserVersion: 'test', format: 'xlsx', fileName: 'conditional.xlsx', size: 10, model: 'spreadsheet',
    payload: { modelType: 'spreadsheet', data: { date1904: false, activeSheet: 0, sheets: [{
      name: 'Scores', cells: [], conditionalFormats: [{
        id: 'conditional-1-1', type: 'cellIs', ranges: [{ startRow: 0, startCol: 1, endRow: 4, endCol: 1 }],
        operator: 'between', formula1: '80', formula2: '100', style: { fill: '#00AA55', fontColor: '#FFFFFF', bold: true },
        priority: 1, stopIfTrue: true,
      }],
    }] } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: ['conditionalFormatting'] },
  })
  assert.deepEqual(workbook.sheets[0].conditionalFormats, [{
    id: 'conditional-1-1', type: 'cellIs', ranges: [{ startRow: 0, startCol: 1, endRow: 4, endCol: 1 }],
    operator: 'between', formula1: '80', formula2: '100', style: { fill: '#00AA55', fontColor: '#FFFFFF', bold: true },
    priority: 1, stopIfTrue: true,
  }])
})
