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

test('visual rules resolve explicit thresholds, formulas, negative data bars and ordered icon boundaries', () => {
  const range = [{ startRow: 0, startCol: 0, endRow: 3, endCol: 0 }]
  const scope = [-100, -50, 0, 100].map((value, row) => ({ row, column: 0, value }))
  const dataBar = {
    ...baseRule, id: 'bar', type: 'dataBar', ranges: range, config: {
      min: { type: 'formula', value: '=-100' }, max: { type: 'num', value: 100 },
      positiveColor: '#008000', nativeColor: '#ff0000', isGradient: true, isShowValue: false,
    },
  }
  const bar = conditionalFormatStyleAt([dataBar], 0, 0, -100, [], scope, undefined, formula => formula === '=-100' ? -100 : undefined)
  assert.deepEqual(bar.dataBar, { fraction: 0.5, start: 0.5, color: '#ff0000', showValue: false, gradient: true, negative: true })
  const icons = ['3Arrows', '3Arrows', '3Arrows'].map((iconType, index) => ({
    value: { type: 'num', value: [67, 33, 0][index] }, operator: 'greaterThanOrEqual', iconType, iconId: String(index),
  }))
  const iconRule = { ...baseRule, id: 'icons', type: 'iconSet', ranges: range, config: { iconSet: '3Arrows', icons } }
  assert.equal(conditionalFormatStyleAt([iconRule], 0, 0, 80, [], scope).icon.iconId, '0')
  assert.equal(conditionalFormatStyleAt([iconRule], 1, 0, 50, [], scope).icon.iconId, '1')
  assert.equal(conditionalFormatStyleAt([iconRule], 2, 0, 5, [], scope).icon.iconId, '2')
  const reversed = conditionalFormatStyleAt([{ ...iconRule, config: { ...iconRule.config, reverse: true } }], 0, 0, 80, [], scope)
  assert.equal(reversed.icon.iconId, '2')
})

test('color scales support two or three threshold stops and ignore invalid formula thresholds', () => {
  const range = [{ startRow: 0, startCol: 0, endRow: 2, endCol: 0 }]
  const scope = [0, 50, 100].map((value, row) => ({ row, column: 0, value }))
  const rule = { ...baseRule, id: 'scale', type: 'colorScale', ranges: range, config: { thresholds: [
    { value: { type: 'min' }, color: '#000000' },
    { value: { type: 'percentile', value: 50 }, color: '#808080' },
    { value: { type: 'max' }, color: '#ffffff' },
  ] } }
  assert.equal(conditionalFormatStyleAt([rule], 1, 0, 50, [], scope).fill, '#808080')
  const formulaScale = { ...rule, config: { thresholds: [
    { value: { type: 'formula', value: '=NA()' }, color: '#000000' },
    { value: { type: 'max' }, color: '#ffffff' },
  ] } }
  assert.deepEqual(conditionalFormatStyleAt([formulaScale], 1, 0, 50, [], scope), {})
})

test('text rules handle case-insensitive matching, empty cells, and spreadsheet error values', () => {
  const textRule = { ...baseRule, type: 'containsText', subType: 'text', operator: 'beginsWith', formula1: 'sales', value: 'sales' }
  assert.equal(evaluateConditionalFormatRule(textRule, 'SALES total'), true)
  assert.equal(evaluateConditionalFormatRule(textRule, 'total sales'), false)
  assert.equal(evaluateConditionalFormatRule({ ...textRule, operator: 'containsBlanks' }, '   '), true)
  assert.equal(evaluateConditionalFormatRule({ ...textRule, operator: 'containsErrors' }, '#DIV/0!'), true)
  assert.equal(evaluateConditionalFormatRule({ ...textRule, operator: 'notContainsErrors' }, '#DIV/0!'), false)
})
