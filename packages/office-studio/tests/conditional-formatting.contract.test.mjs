import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { evaluateFormula } from '../src/features/formula/formula-engine.ts'
import { conditionalFormatFormulaAt } from '../src/features/conditional-formatting/conditional-format-formulas.ts'
import { conditionalFormatEditorFormula, mergeConditionalFormatRuleDraft, moveConditionalFormatRuleDraft,
  prependConditionalFormatRuleDraft, removeConditionalFormatRuleDraft } from '../src/features/conditional-formatting/conditional-format-rule-editor.ts'
import {
  addConditionalFormatRule,
  conditionalFormatStyleAt,
  deleteConditionalFormatRule,
  evaluateConditionalFormatRule,
  moveConditionalFormatRule,
  updateConditionalFormatRule,
  validateConditionalFormatDraft,
} from '../src/features/conditional-formatting/conditional-formatting.ts'
import {
  createConditionalIconConfig,
  conditionalIconSequence,
  validateConditionalIconConfig,
} from '../src/features/conditional-formatting/icon-set-config.ts'

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

test('icon set defaults use Univer ordered numeric thresholds and a complementary final bucket', () => {
  assert.deepEqual(createConditionalIconConfig('3Arrows').map(icon => [icon.operator, icon.value, icon.iconId]), [
    ['greaterThan', { type: 'num', value: 20 }, '0'],
    ['greaterThan', { type: 'num', value: 10 }, '1'],
    ['lessThanOrEqual', { type: 'num', value: Number.MAX_SAFE_INTEGER }, '2'],
  ])
  assert.deepEqual(createConditionalIconConfig('5Rating').map(icon => icon.value.value), [40, 30, 20, 10, Number.MAX_SAFE_INTEGER])
  assert.throws(() => createConditionalIconConfig('unknown'), /有效的图标集/)
})

test('icon set buckets handle exclusive and inclusive boundaries, mixed threshold types and formulas', () => {
  const range = [{ startRow: 0, startCol: 0, endRow: 3, endCol: 0 }]
  const icons = [
    { operator: 'greaterThanOrEqual', value: { type: 'num', value: 80 }, iconType: '3Arrows', iconId: '0' },
    { operator: 'greaterThan', value: { type: 'num', value: 40 }, iconType: '3Arrows', iconId: '1' },
    { operator: 'ignored', value: { type: 'min' }, iconType: '3Arrows', iconId: '2' },
  ]
  const rule = { ...baseRule, type: 'iconSet', ranges: range, config: { icons } }
  const iconAt = (value, values = [], formulaValue) => conditionalFormatStyleAt([rule], 0, 0, value, [], values.map((item, row) => ({ row, column: 0, value: item })), undefined, formulaValue).icon?.iconId
  assert.equal(iconAt(80), '0')
  assert.equal(iconAt(79), '1')
  assert.equal(iconAt(40), '2')
  assert.equal(iconAt(39), '2')

  const rangedRule = { ...rule, config: { icons: [
    { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 75 }, iconType: '3Arrows', iconId: '0' },
    { operator: 'greaterThan', value: { type: 'percentile', value: 50 }, iconType: '3Arrows', iconId: '1' },
    icons[2],
  ] } }
  const rangedAt = value => conditionalFormatStyleAt([rangedRule], 0, 0, value, [], [0, 10, 20, 30].map((item, row) => ({ row, column: 0, value: item }))).icon?.iconId
  assert.equal(rangedAt(22.5), '0')
  assert.equal(rangedAt(20), '1')
  assert.equal(rangedAt(15), '2')

  const formulaRule = { ...rule, config: { icons: [
    { operator: 'greaterThanOrEqual', value: { type: 'formula', value: '=A1' }, iconType: '3Arrows', iconId: '0' },
    icons[1], icons[2],
  ] } }
  assert.equal(conditionalFormatStyleAt([formulaRule], 0, 0, 50, [], [], undefined, formula => formula === '=A1' ? 60 : undefined).icon?.iconId, '1')
  assert.equal(conditionalFormatStyleAt([formulaRule], 0, 0, 70, [], [], undefined, formula => formula === '=A1' ? 60 : undefined).icon?.iconId, '0')
  assert.equal(conditionalFormatStyleAt([formulaRule], 0, 0, 70, [], [], undefined, () => undefined).icon, undefined)
})

test('icon-set reverse mirrors icon choices while preserving bucket thresholds', () => {
  const icons = [
    { operator: 'greaterThanOrEqual', value: { type: 'num', value: 67 }, iconType: '3Arrows', iconId: '0' },
    { operator: 'greaterThan', value: { type: 'num', value: 33 }, iconType: 'EMPTY_ICON_TYPE', iconId: '' },
    { operator: 'ignored', value: { type: 'unsupported', value: 'ignored' }, iconType: '3Flags', iconId: '2' },
  ]
  const config = { icons, reverse: true }
  assert.deepEqual(conditionalIconSequence(config).map(icon => [icon.iconType, icon.iconId, icon.value]), [
    ['3Flags', '2', { type: 'num', value: 67 }],
    ['EMPTY_ICON_TYPE', '', { type: 'num', value: 33 }],
    ['3Arrows', '0', { type: 'unsupported', value: 'ignored' }],
  ])
  assert.equal(conditionalFormatStyleAt([
    { ...baseRule, type: 'iconSet', config },
  ], 0, 0, 80).icon.iconId, '2')
  assert.equal(conditionalFormatStyleAt([
    { ...baseRule, type: 'iconSet', config },
  ], 0, 0, 50).icon.iconId, '')
  assert.equal(conditionalFormatStyleAt([
    { ...baseRule, type: 'iconSet', config },
  ], 0, 0, 10).icon.iconId, '0')
})

test('icon-set validation checks visible buckets and permits a no-icon entry', () => {
  const config = { icons: [
    { operator: 'greaterThanOrEqual', value: { type: 'num', value: 67 }, iconType: '3Arrows', iconId: '0' },
    { operator: 'greaterThan', value: { type: 'num', value: 33 }, iconType: 'EMPTY_ICON_TYPE', iconId: '' },
    { operator: 'invalid-final-bucket-operator', value: { type: 'ignored', value: '' }, iconType: '3Flags', iconId: '2' },
  ] }
  assert.doesNotThrow(() => validateConditionalIconConfig(config))
  assert.throws(() => validateConditionalIconConfig({ icons: [
    { operator: 'greaterThanOrEqual', value: { type: 'num', value: 33 }, iconType: '3Arrows', iconId: '0' },
    { operator: 'greaterThanOrEqual', value: { type: 'num', value: 67 }, iconType: '3Arrows', iconId: '1' },
    config.icons[2],
  ] }), /从高到低/)
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

test('formula rules return false when no formula evaluator is connected', () => {
  const formulaRule = { ...baseRule, type: 'expression', operator: undefined, formula1: '=A1>0' }
  assert.equal(evaluateConditionalFormatRule(formulaRule, 1), false)
  assert.equal(evaluateConditionalFormatRule(formulaRule, 1, [], formula => formula === '=A1>0'), true)
})

test('numeric operands evaluate actual formulas from their preserved origin and reject nonnumeric results', () => {
  const workbook = createBlankWorkbook(), target = workbook.sheets[0]
  target.cells['0:0'] = createCell(0, 0, 5)
  target.cells['0:1'] = createCell(0, 1, 8)
  target.cells['1:0'] = createCell(1, 0, 9)
  const rule = { ...baseRule, ranges: [{ startRow: 0, startCol: 0, endRow: 3, endCol: 0 }], formulaAnchor: { row: 1, col: 0 }, formula1: '=A1', formula2: '=$B$1+1' }
  const at = (formula, row) => evaluateFormula(conditionalFormatFormulaAt(formula, rule, row, 0), workbook, target.id, { formulaCell: { row, column: 0 } })
  assert.equal(evaluateConditionalFormatRule(rule, 7, [], undefined, formula => at(formula, 0)), false)
  assert.equal(evaluateConditionalFormatRule(rule, 7, [], undefined, formula => at(formula, 1)), true)
  assert.equal(evaluateConditionalFormatRule(rule, 7, [], undefined, formula => at(formula, 2)), false)
  for (const formula1 of ['=NA()', '="4"', '=TRUE', '=0x10']) {
    assert.equal(evaluateConditionalFormatRule({ ...rule, operator: 'notEqual', formula1, formula2: undefined }, 7, [], undefined, formula => at(formula, 1)), false, formula1)
  }
})

test('numeric drafts accept Office formula syntax and reject malformed or JavaScript-only number literals', () => {
  const draft = { ranges: [{ startRow: 1, startCol: 0, endRow: 3, endCol: 0 }], operator: 'between', formula1: '=A1+$B$1', formula2: '=ROW()+10', style: { fill: '#123456' }, formulaAnchor: { row: 1, col: 0 } }
  assert.doesNotThrow(() => validateConditionalFormatDraft(draft))
  const created = addConditionalFormatRule(sheet, draft, 'formula-operands').conditionalFormats[0]
  assert.equal(created.formula1, 'A1+$B$1')
  assert.equal(created.formula2, 'ROW()+10')
  assert.deepEqual(created.formulaAnchor, { row: 1, col: 0 })
  for (const formula1 of ['', '=', 'SUM(', '0x10', '0b11', 'Infinity', '1e999']) assert.throws(() => validateConditionalFormatDraft({ ...draft, formula1 }), /finite number or valid formula/, formula1)
  assert.throws(() => validateConditionalFormatDraft({ ...draft, formula2: '=SUM(' }), /second number or formula/)
  assert.doesNotThrow(() => validateConditionalFormatDraft({ ...draft, formula1: '-1.2e3', formula2: '=25%' }))
})

test('editing numeric ranges preserves recoverable reference underflow and safely rebases later ranges', () => {
  const original = { ...sheet, conditionalFormats: [{ ...baseRule, operator: 'greaterThan', formula1: 'A1', formula2: undefined, ranges: [{ startRow: 1, startCol: 0, endRow: 4, endCol: 0 }], metadata: { vendor: { keep: ['original'] } } }] }
  const draft = { ranges: [{ startRow: 0, startCol: 0, endRow: 3, endCol: 0 }], operator: 'greaterThan', formula1: 'A1', style: { fill: '#334455' } }
  const earlier = updateConditionalFormatRule(original, 'rule-a', draft).conditionalFormats[0]
  assert.equal(earlier.formula1, 'A1')
  assert.deepEqual(earlier.formulaAnchor, { row: 1, col: 0 })
  assert.equal(conditionalFormatFormulaAt(earlier.formula1, earlier, 0, 0), '#REF!')
  assert.equal(conditionalFormatFormulaAt(earlier.formula1, earlier, 1, 0), 'A1')
  const later = updateConditionalFormatRule({ ...sheet, conditionalFormats: [earlier] }, 'rule-a', { ...draft, ranges: [{ startRow: 3, startCol: 0, endRow: 6, endCol: 0 }] }).conditionalFormats[0]
  assert.equal(later.formula1, 'A3')
  assert.equal(later.formulaAnchor, undefined)
  later.metadata.vendor.keep.push('edited')
  assert.deepEqual(original.conditionalFormats[0].metadata, { vendor: { keep: ['original'] } })
})

test('rule editor retains exact unedited formula text and normalizes only a changed expression', () => {
  const imported = '=UNSUPPORTED([VendorColumn])'
  assert.equal(conditionalFormatEditorFormula('UNSUPPORTED([VendorColumn])', imported), imported)
  assert.equal(conditionalFormatEditorFormula('  =a1+$B$2  ', '=a1+$B$2'), '=a1+$B$2')
  assert.equal(conditionalFormatEditorFormula('  =A1>8  ', '=A1>5'), 'A1>8')
  assert.equal(conditionalFormatEditorFormula('  =ROW()+1  '), 'ROW()+1')
})

test('rule editor preserves compatible imported discriminators and changes incompatible text kinds', () => {
  const original = { ...baseRule, type: 'highlightCell', subType: 'formula', formula1: '=A1>5' }
  const edited = mergeConditionalFormatRuleDraft(original, { ...original, type: 'expression', style: { bold: true } })
  assert.equal(edited.type, 'highlightCell'); assert.equal(edited.subType, 'formula')
  const number = mergeConditionalFormatRuleDraft(baseRule, { ...baseRule, subType: 'number' })
  assert.equal(number.subType, undefined)
  const changedSubtype = mergeConditionalFormatRuleDraft({ ...baseRule, type: 'highlightCell', subType: 'number' }, { ...baseRule, type: 'highlightCell', subType: 'text' })
  assert.equal(changedSubtype.subType, 'text')
  const text = mergeConditionalFormatRuleDraft({ ...baseRule, type: 'beginsWith', formula1: 'A1' }, { ...baseRule, type: 'containsText', subType: 'text', operator: 'containsText', formula1: 'A1' })
  assert.equal(text.type, 'containsText'); assert.equal(text.subType, 'text')
})

test('rule-list drafts retain untouched priority slots and deeply isolate prepend, removal, moves and no-ops', () => {
  const first = { ...baseRule, priority: 4, formulaAnchor: { row: 1, col: 0 }, metadata: { keep: ['first'] }, config: { custom: { keep: ['first'] } } }
  const second = { ...baseRule, id: 'second', priority: 9, metadata: { keep: ['second'] } }
  const untouched = { ...baseRule, id: 'untouched', priority: 17, metadata: { keep: ['untouched'] } }
  const original = [first, second, untouched], draft = { ...baseRule, id: 'new', priority: 99, metadata: { keep: ['new'] } }
  const order = rules => rules.map(rule => [rule.id, rule.priority])
  const added = prependConditionalFormatRuleDraft(original, draft)
  assert.deepEqual(order(added), [['new', 4], ['rule-a', 4], ['second', 9], ['untouched', 17]])
  assert.deepEqual(added.slice(1), original); assert.equal(draft.priority, 99)
  assert.equal(prependConditionalFormatRuleDraft([], draft)[0].priority, 1)
  const removed = removeConditionalFormatRuleDraft(original, 'rule-a')
  assert.deepEqual(removed, [second, untouched])
  const moved = moveConditionalFormatRuleDraft(original, 'rule-a', 1)
  assert.deepEqual(order(moved), [['second', 4], ['rule-a', 9], ['untouched', 17]])
  assert.deepEqual(moveConditionalFormatRuleDraft(moved, 'rule-a', -1), original)
  const tied = moveConditionalFormatRuleDraft([first, { ...second, priority: 4 }, untouched], 'rule-a', 1)
  assert.deepEqual(order(tied), [['second', 4], ['rule-a', 4], ['untouched', 17]])
  const noOp = moveConditionalFormatRuleDraft(original, 'missing', 1)
  assert.deepEqual(noOp, original)
  assert.deepEqual(moveConditionalFormatRuleDraft(original, 'rule-a', -1), original)
  assert.deepEqual(moveConditionalFormatRuleDraft(original, 'untouched', 1), original)
  assert.deepEqual(removeConditionalFormatRuleDraft(original, 'missing'), original)
  added[0].metadata.keep.push('added draft'); added[1].metadata.keep.push('added existing')
  removed[0].metadata.keep.push('removed snapshot'); moved[1].formulaAnchor.row = 2
  tied[2].metadata.keep.push('tied snapshot'); noOp[0].config.custom.keep.push('no-op snapshot')
  assert.deepEqual(draft.metadata.keep, ['new']); assert.deepEqual(first.metadata.keep, ['first'])
  assert.deepEqual(second.metadata.keep, ['second']); assert.deepEqual(untouched.metadata.keep, ['untouched'])
  assert.deepEqual(first.formulaAnchor, { row: 1, col: 0 }); assert.deepEqual(first.config.custom.keep, ['first'])
  assert.deepEqual(order(original), [['rule-a', 4], ['second', 9], ['untouched', 17]])
})

test('rule editor preserves unknown configuration, threshold and metadata fields while resetting known style fields', () => {
  const previous = { ...baseRule, type: 'dataBar', formulaAnchor: { row: 1, col: 0 }, metadata: { vendor: { keep: ['source'] } }, config: {
    custom: { formulaLikeData: 'A1', nested: ['source'] }, min: { type: 'num', value: 0, vendor: { keep: true } }, max: { type: 'num', value: 100 }, positiveColor: '#123456',
  }, style: { fill: '#ff0000', bold: true, vendor: { retained: true } } }
  const draft = { ...baseRule, type: 'dataBar', formulaAnchor: { row: 1, col: 0 }, formula2: undefined, style: { fill: undefined, bold: false }, config: { min: { type: 'formula', value: '=A1' }, max: { type: 'num', value: 80 }, positiveColor: '#334455' } }
  const edited = mergeConditionalFormatRuleDraft(previous, draft)
  assert.deepEqual(edited.style, { bold: false, vendor: { retained: true } })
  assert.deepEqual(edited.config.min, { type: 'formula', value: '=A1', vendor: { keep: true } })
  assert.equal(edited.formula2, undefined)
  assert.deepEqual(edited.formulaAnchor, { row: 1, col: 0 })
  assert.deepEqual(edited.config.custom, previous.config.custom)
  edited.config.custom.nested.push('edit'); edited.metadata.vendor.keep.push('edit'); edited.config.min.vendor.keep = false
  assert.deepEqual(previous.config.custom.nested, ['source'])
  assert.deepEqual(previous.metadata.vendor.keep, ['source'])
  assert.equal(previous.config.min.vendor.keep, true)
})

test('color-scale edits preserve each endpoint extension when toggling the middle stop and clear stale operand values', () => {
  const previous = { ...baseRule, type: 'colorScale', config: { thresholds: [
    { color: '#000000', value: { type: 'num', value: 0, vendor: 'minimum' }, vendor: { endpoint: 'minimum' } },
    { color: '#888888', value: { type: 'num', value: 50, vendor: 'middle' }, vendor: { endpoint: 'middle' } },
    { color: '#ffffff', value: { type: 'num', value: 100, vendor: 'maximum' }, vendor: { endpoint: 'maximum' } },
  ] } }
  const edited = mergeConditionalFormatRuleDraft(previous, { ...baseRule, type: 'colorScale', config: { thresholds: [
    { color: '#112233', value: { type: 'min' } }, { color: '#aabbcc', value: { type: 'max' } },
  ] } })
  assert.deepEqual(edited.config.thresholds.map(stop => [stop.vendor.endpoint, stop.value.vendor, stop.value.type, stop.value.value]), [
    ['minimum', 'minimum', 'min', undefined], ['maximum', 'maximum', 'max', undefined],
  ])
  assert.equal(previous.config.thresholds.length, 3)
})

test('rule priority and deletion snapshots isolate nested opaque configuration and formula origins', () => {
  const first = { ...baseRule, formulaAnchor: { row: 1, col: 0 }, metadata: { source: ['kept'] }, config: { custom: { retained: true } } }
  const second = { ...baseRule, id: 'second', priority: 2 }
  const original = { ...sheet, conditionalFormats: [first, second] }
  const moved = moveConditionalFormatRule(original, 'rule-a', 1)
  moved.conditionalFormats[1].formulaAnchor.row = 2
  moved.conditionalFormats[1].config.custom.retained = false
  moved.conditionalFormats[1].metadata.source.push('changed')
  assert.deepEqual(first.formulaAnchor, { row: 1, col: 0 })
  assert.equal(first.config.custom.retained, true)
  assert.deepEqual(first.metadata.source, ['kept'])
  const deleted = deleteConditionalFormatRule(original, 'second')
  deleted.conditionalFormats[0].metadata.source.push('deleted snapshot')
  assert.deepEqual(first.metadata.source, ['kept'])
})

test('numeric rule validation rejects worksheet and origin overflow while accepting their exact upper limits', () => {
  const draft = { ranges: [{ startRow: 1_048_575, startCol: 16_383, endRow: 1_048_575, endCol: 16_383 }], operator: 'equal', formula1: '=ROW()+COLUMN()', style: { bold: true }, formulaAnchor: { row: 1_048_575, col: 16_383 } }
  assert.doesNotThrow(() => validateConditionalFormatDraft(draft))
  assert.throws(() => validateConditionalFormatDraft({ ...draft, formulaAnchor: { row: 1_048_576, col: 0 } }), /origin is out of range/)
  assert.throws(() => validateConditionalFormatDraft({ ...draft, formulaAnchor: { row: 0, col: -1 } }), /origin is out of range/)
  assert.throws(() => validateConditionalFormatDraft({ ...draft, ranges: [{ ...draft.ranges[0], endCol: 16_384 }] }), /valid cell ranges/)
})

function conditionalEnvelope(rules) {
  return { schemaVersion: 1, parserVersion: 'cf-extension-test', format: 'xlsx', fileName: 'extensions.xlsx', size: 10, model: 'spreadsheet',
    payload: { modelType: 'spreadsheet', data: { sheets: [{ name: 'Data', cells: [], conditionalFormats: rules }] } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: ['conditionalFormatting'] } }
}

test('conditional normalization deep-clones unknown root, style, visual stop and icon fields without changing the formula origin', () => {
  const raw = { ...baseRule, type: 'colorScale', formulaAnchor: { row: 1, col: 0 }, value: [10, 20], vendor: { root: ['source'] },
    style: { fill: '#112233', vendor: { style: ['source'] } }, metadata: { vendor: ['source'] }, config: {
      custom: { config: ['source'] }, min: { type: 'formula', value: '=A1', vendor: { min: ['source'] } }, max: { type: 'max', vendor: { max: ['source'] } },
      thresholds: [{ color: '#112233', value: { type: 'min', vendor: { value: ['source'] } }, vendor: { stop: ['source'] } },
        { color: '#aabbcc', value: { type: 'max' } }], icons: [{ iconType: '3Arrows', iconId: '0', operator: 'greaterThan',
        value: { type: 'formula', value: '=A1', vendor: { value: ['source'] } }, vendor: { icon: ['source'] } }],
    } }
  const normalized = normalizeWorkbook(conditionalEnvelope([raw])).sheets[0].conditionalFormats[0]
  assert.deepEqual(normalized.formulaAnchor, raw.formulaAnchor); assert.deepEqual(normalized.vendor, raw.vendor)
  assert.deepEqual(normalized.config.custom, raw.config.custom); assert.deepEqual(normalized.config.min.vendor, raw.config.min.vendor)
  assert.deepEqual(normalized.config.thresholds[0].vendor, raw.config.thresholds[0].vendor)
  assert.deepEqual(normalized.config.icons[0].vendor, raw.config.icons[0].vendor)
  normalized.vendor.root.push('changed'); normalized.style.vendor.style.push('changed'); normalized.metadata.vendor.push('changed')
  normalized.config.min.vendor.min.push('changed'); normalized.config.max.vendor.max.push('changed'); normalized.config.custom.config.push('changed')
  normalized.config.thresholds[0].value.vendor.value.push('changed'); normalized.config.thresholds[0].vendor.stop.push('changed')
  normalized.config.icons[0].value.vendor.value.push('changed'); normalized.config.icons[0].vendor.icon.push('changed'); normalized.value[0] = 99
  assert.deepEqual(raw.vendor.root, ['source']); assert.deepEqual(raw.style.vendor.style, ['source']); assert.deepEqual(raw.metadata.vendor, ['source'])
  assert.deepEqual(raw.config.min.vendor.min, ['source']); assert.deepEqual(raw.config.max.vendor.max, ['source'])
  assert.deepEqual(raw.config.custom.config, ['source']); assert.deepEqual(raw.config.thresholds[0].vendor.stop, ['source'])
  assert.deepEqual(raw.config.thresholds[0].value.vendor.value, ['source']); assert.deepEqual(raw.config.icons[0].vendor.icon, ['source'])
  assert.deepEqual(raw.config.icons[0].value.vendor.value, ['source']); assert.deepEqual(raw.value, [10, 20])
})

test('conditional normalization rejects malformed or partly invalid ranges instead of changing the origin by filtering', () => {
  const valid = { startRow: 0, startCol: 0, endRow: 4, endCol: 0 }, invalid = { ...valid, endRow: 1_048_576 }
  for (const ranges of [null, [], [invalid], [invalid, valid], [valid, invalid], [{ ...valid, startCol: -1 }]]) {
    assert.throws(() => normalizeWorkbook(conditionalEnvelope([{ ...baseRule, ranges }])), /范围无效/)
  }
  const upper = { startRow: 1_048_575, endRow: 1_048_575, startCol: 16_383, endCol: 16_383 }
  assert.deepEqual(normalizeWorkbook(conditionalEnvelope([{ ...baseRule, ranges: [upper] }])).sheets[0].conditionalFormats[0].ranges, [upper])
})

test('conditional model/view adapters isolate opaque configuration even when a view omits the rule field', () => {
  const workbook = createBlankWorkbook(), source = { ...baseRule, formulaAnchor: { row: 1, col: 0 }, vendor: { root: ['source'] },
    style: { fill: '#112233', vendor: { style: ['source'] } }, config: { custom: { values: ['source'] } }, metadata: { vendor: ['source'] } }
  workbook.sheets[0].conditionalFormats = [source]
  const view = toExcelWorkbookView(workbook), restored = applyExcelWorkbookView(workbook, view)
  view.sheets[0].conditionalFormats[0].config.custom.values.push('view')
  restored.sheets[0].conditionalFormats[0].vendor.root.push('restored')
  assert.deepEqual(source.config.custom.values, ['source']); assert.deepEqual(source.vendor.root, ['source'])
  delete view.sheets[0].conditionalFormats
  const retained = applyExcelWorkbookView(workbook, view).sheets[0].conditionalFormats[0]
  retained.formulaAnchor.row = 99; retained.style.vendor.style.push('fallback'); retained.metadata.vendor.push('fallback')
  assert.deepEqual(source.formulaAnchor, { row: 1, col: 0 }); assert.deepEqual(source.style.vendor.style, ['source']); assert.deepEqual(source.metadata.vendor, ['source'])
})
