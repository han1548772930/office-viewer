import assert from 'node:assert/strict'
import { test } from 'vitest'
import { reactive } from 'vue'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { conditionalFormatStyleAt } from '../src/features/conditional-formatting/conditional-formatting.ts'
import { conditionalFormatFormulaAt } from '../src/features/conditional-formatting/conditional-format-formulas.ts'
import { evaluateFormula } from '../src/features/formula/formula-engine.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import {
  assertConditionalFormatPainterSource, captureConditionalFormatPainterSource, createConditionalFormatPainterPlan,
} from '../src/features/sheet-workbench/sheet-format-painter-conditional.ts'
import {
  applySheetFormatPainterCommand, applySheetViewFormatPainterCommand, captureSheetFormatPainterSource,
  captureSheetViewFormatPainterSource,
} from '../src/features/sheet-workbench/sheet-format-painter.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'

const range = (startRow, startCol, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })
const contains = (area, row, col) => row >= area.startRow && row <= area.endRow && col >= area.startCol && col <= area.endCol
const cells = ranges => {
  const result = new Set()
  for (const area of ranges) for (let row = area.startRow; row <= area.endRow; row++) for (let col = area.startCol; col <= area.endCol; col++) result.add(`${row}:${col}`)
  return [...result].sort()
}
const rule = (id = 'source-cf', ranges = [range(0, 0)], extra = {}) => ({
  id, type: 'cellIs', ranges, operator: 'greaterThan', formula1: '5', style: { fill: '#fa6473', bold: true }, priority: 1, stopIfTrue: false, ...extra,
})
const capture = (rules, sampled = range(0, 0)) => captureConditionalFormatPainterSource('source-sheet', rules, sampled)
const plan = (source, current = undefined, sampled = range(0, 0), targets = [range(3, 3)], sheetId = 'target-sheet') => createConditionalFormatPainterPlan(sheetId, current, source, sampled, targets)
const expressionAt = (entry, row, col) => {
  return conditionalFormatFormulaAt(entry.formula1, entry, row, col)
}

test('conditional capture preserves full rule anchors and clones only intersecting rules and their nested data', () => {
  const original = [rule('outside', [range(10, 10)]), rule('source-cf', [range(0, 0, 4, 4)], {
    type: 'dataBar', config: { min: { type: 'formula', value: '=A1' }, max: { type: 'num', value: 100 }, isGradient: true }, metadata: { vendor: { keep: 'original' } },
  })]
  const source = capture(reactive(original), range(2, 2))
  assert.deepEqual(source.rules.map(entry => entry.id), ['source-cf']); assert.deepEqual(source.rules[0].ranges, [range(0, 0, 4, 4)])
  original[1].config.min.value = '=Z99'; original[1].metadata.vendor.keep = 'later edit'; original[1].style.fill = '#ffffff'
  assert.equal(source.rules[0].config.min.value, '=A1'); assert.equal(source.rules[0].metadata.vendor.keep, 'original'); assert.equal(source.rules[0].style.fill, '#fa6473')
})

test('the complete painter source deeply freezes conditional rules without sampling source values', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, 'source content excluded')
  sheet.conditionalFormats = [rule('source-cf', [range(0, 0)], { metadata: { vendor: { keep: true } } })]
  const source = captureSheetFormatPainterSource(reactive(sheet), range(0, 0))
  assert.equal(Object.isFrozen(source.conditional.rules[0].metadata.vendor), true)
  assert.equal(JSON.stringify(source).includes('source content excluded'), false)
  assert.throws(() => { source.conditional.rules[0].style.fill = '#000000' }, TypeError)
})

test('an empty sampled rule stack clears destination rule coverage while retaining every outside cell', () => {
  const original = [rule('target-cf', [range(2, 2, 6, 6)]), rule('untouched', [range(9, 9)])]
  const before = cloneModel(original), target = range(3, 3, 4, 4), result = plan(capture([]), original, range(0, 0), [target])
  assert.equal(result.changed, true); assert.deepEqual(original, before)
  const expected = cells(before[0].ranges).filter(key => { const [row, col] = key.split(':').map(Number); return !contains(target, row, col) })
  assert.deepEqual(cells(result.rules.find(entry => entry.id === 'target-cf').ranges), expected)
  assert.deepEqual(result.rules.find(entry => entry.id === 'untouched'), before[1])
})

test('fully covered destination rules are removed and an already empty destination is a no-op', () => {
  const result = plan(capture([]), [rule('target-cf', [range(3, 3)])])
  assert.equal(result.changed, true); assert.deepEqual(result.rules, [])
  assert.deepEqual(plan(capture([])), { changed: false, rules: undefined })
  assert.deepEqual(plan(capture([]), []), { changed: false, rules: [] })
})

test('legacy style-only snapshots and empty targets retain the existing conditional boundary', () => {
  const current = [rule('target-cf', [range(3, 3)])]
  assert.deepEqual(plan(undefined, current), { changed: false, rules: current })
  assert.deepEqual(plan(capture([]), current, range(0, 0), []), { changed: false, rules: current })
})

test('a single target cell expands the sampled rule pattern without changing destination data', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.conditionalFormats = [rule('source-cf', [range(0, 1)])]
  sheet.cells['3:3'] = { ...createCell(3, 3, 7), formula: '=7', dataType: 'formula', cachedDataType: 'number', metadata: { keep: true } }
  sheet.cells['3:4'] = { ...createCell(3, 4, 'target'), richText: [{ text: 'target', style: { italic: true } }] }
  const source = captureSheetFormatPainterSource(sheet, range(0, 0, 1, 1)), before = cloneModel(sheet.cells)
  const result = applySheetFormatPainterCommand(sheet, source, [range(3, 3)])
  assert.deepEqual(result.ranges, [range(3, 3, 4, 4)]); assert.equal(result.contentChanged, false)
  assert.deepEqual(sheet.cells, before); assert.deepEqual(cells(sheet.conditionalFormats[0].ranges), ['0:1', '3:4'])
})

test('partial source tiles retain conditional fragments at trailing rows and columns', () => {
  const sampled = range(0, 0, 1, 1), source = capture([rule('source-cf', [range(0, 0, 1, 0)])], sampled)
  const result = plan(source, undefined, sampled, [range(3, 3, 5, 7)])
  assert.deepEqual(cells(result.rules[0].ranges), ['3:3', '3:5', '3:7', '4:3', '4:5', '4:7', '5:3', '5:5', '5:7'])
  assert.ok(result.rules[0].ranges.every(area => area.endRow <= 5 && area.endCol <= 7))
})

test('later selections replace the earlier conditional tiling phase at their overlap', () => {
  const sampled = range(0, 0, 1, 1), targets = [range(3, 3, 5, 5), range(4, 4, 6, 6)]
  const source = capture([rule('source-cf', [range(0, 0)])], sampled), result = plan(source, undefined, sampled, targets), expected = []
  for (let row = 3; row <= 6; row++) for (let col = 3; col <= 6; col++) {
    const owner = [...targets].reverse().find(area => contains(area, row, col))
    if (owner && (row - owner.startRow) % 2 === 0 && (col - owner.startCol) % 2 === 0) expected.push(`${row}:${col}`)
  }
  assert.deepEqual(cells(result.rules[0].ranges), expected.sort())
})

test('adjacent and overlapping sampled rule ranges yield unique destination coverage and stable repainting', () => {
  const sampled = range(0, 0, 1, 2), source = capture([rule('source-cf', [range(0, 0, 1, 1), range(0, 1, 1, 2)])], sampled)
  const result = plan(source, undefined, sampled, [range(3, 3, 4, 5)])
  assert.deepEqual(result.rules[0].ranges, [range(3, 3, 4, 5)])
  const repeated = plan(source, result.rules, sampled, [range(3, 3, 4, 5)])
  assert.equal(repeated.changed, false); assert.deepEqual(repeated.rules, result.rules)
})

test('range replacement preserves arbitrary small unions without duplicated coverage', () => {
  let seed = 741
  const next = () => { seed = (Math.imul(seed, 1_664_525) + 1_013_904_223) >>> 0; return seed % 7 }
  for (let scenario = 0; scenario < 80; scenario++) {
    const areas = Array.from({ length: 4 }, () => { const r1 = next(), r2 = next(), c1 = next(), c2 = next(); return range(Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2)) })
    const r1 = next(), r2 = next(), c1 = next(), c2 = next(), target = range(Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2))
    const result = plan(capture([]), [rule('target-cf', areas)], range(0, 0), [target])
    const expected = cells(areas).filter(key => { const [row, col] = key.split(':').map(Number); return !contains(target, row, col) })
    const actual = result.rules?.[0]?.ranges ?? []
    assert.deepEqual(cells(actual), expected)
    if (result.changed) assert.equal(actual.reduce((count, area) => count + (area.endRow - area.startRow + 1) * (area.endCol - area.startCol + 1), 0), expected.length)
  }
})

test('whole-sheet destination coverage is clipped geometrically at a small target', () => {
  const current = [rule('whole-sheet', [range(0, 0, 1_048_575, 16_383)])]
  const result = plan(capture([]), current, range(0, 0), [range(4, 4)])
  assert.ok(result.rules[0].ranges.length <= 4)
  assert.equal(result.rules[0].ranges.some(area => contains(area, 4, 4)), false)
  assert.equal(result.rules[0].ranges.some(area => contains(area, 1_048_575, 16_383)), true)
})

test('painting the source back onto itself preserves its original rule list and range representation', () => {
  const sampled = range(0, 0, 1, 1), current = [rule('source-cf', [range(0, 0, 0, 1), range(1, 0, 1, 1)]), rule('outside', [range(10, 10)], { priority: 2 })]
  const result = plan(capture(current, sampled), current, sampled, [sampled], 'source-sheet')
  assert.equal(result.changed, false); assert.deepEqual(result.rules, current)
})

test('same-sheet paints expand equivalent source rule IDs while retaining priority and unrelated rule order', () => {
  const current = [rule('outside', [range(10, 10)], { priority: 3 }), rule('source-cf', [range(0, 0)], { priority: 2 })]
  const source = capture(current), result = plan(source, current, range(0, 0), [range(3, 3)], 'source-sheet')
  assert.deepEqual(result.rules.map(entry => entry.id), ['outside', 'source-cf'])
  assert.equal(result.rules[1].priority, 2); assert.deepEqual(cells(result.rules[1].ranges), ['0:0', '3:3'])
  assert.equal(plan(source, result.rules, range(0, 0), [range(3, 3)], 'source-sheet').changed, false)
})

test('cross-sheet copies have independent IDs and preserve the complete visual configuration', () => {
  const original = rule('source-cf', [range(0, 0)], { type: 'dataBar', stopIfTrue: true, metadata: { vendor: { keep: true } }, config: {
    min: { type: 'num', value: -100 }, max: { type: 'num', value: 100 }, isGradient: true, positiveColor: '#008800', nativeColor: '#cc0000', isShowValue: false,
  } })
  const source = capture([original]), result = plan(source)
  assert.notEqual(result.rules[0].id, original.id); assert.deepEqual(result.rules[0].config, original.config)
  assert.equal(result.rules[0].stopIfTrue, true); assert.deepEqual(result.rules[0].metadata, original.metadata)
  result.rules[0].metadata.vendor.keep = false; assert.equal(source.rules[0].metadata.vendor.keep, true)
})

test('repainting an independent destination retains its ID, original list slot and exact rule snapshot', () => {
  const source = capture([rule()]), first = plan(source, [rule('outside', [range(10, 10)], { priority: 20 })])
  const current = [first.rules[1], first.rules[0]], repeated = plan(source, current)
  assert.equal(repeated.changed, false); assert.deepEqual(repeated.rules, current)
})

test('later source edits keep the current source rule and apply the immutable sampled definition to targets', () => {
  const original = [rule()], source = capture(original)
  const current = cloneModel(original); current[0].style.fill = '#ffffff'
  const result = plan(source, current, range(0, 0), [range(3, 3)], 'source-sheet')
  assert.equal(result.rules[0].style.fill, '#ffffff'); assert.deepEqual(result.rules[0].ranges, [range(0, 0)])
  assert.equal(result.rules[1].style.fill, '#fa6473'); assert.deepEqual(result.rules[1].ranges, [range(3, 3)])
})

test('a changed source stack is copied together so mixed live reuse cannot reverse its target precedence', () => {
  const original = [rule('first', [range(0, 0)], { priority: 1, style: { fill: '#ff0000' }, stopIfTrue: true }), rule('second', [range(0, 0)], { priority: 2, style: { fill: '#00ff00' } })]
  const source = capture(original), current = cloneModel(original); current[0].style.fill = '#0000ff'
  const result = plan(source, current, range(0, 0), [range(3, 3)], 'source-sheet')
  assert.deepEqual(conditionalFormatStyleAt(result.rules, 3, 3, 10), { fill: '#ff0000' })
  assert.deepEqual(conditionalFormatStyleAt(result.rules, 0, 0, 10), { fill: '#0000ff' })
  assert.equal(result.rules.filter(entry => entry.ranges.some(area => contains(area, 3, 3))).length, 2)
})

test('source priority edits and equal-priority list reordering retain the sampled target ordering', () => {
  for (const reorder of [false, true]) {
    const original = [rule('first', [range(0, 0)], { priority: 1, style: { fill: '#ff0000' } }), rule('second', [range(0, 0)], { priority: reorder ? 1 : 2, style: { fill: '#00ff00' } })]
    const source = capture(original), current = cloneModel(original)
    if (reorder) current.reverse(); else current[0].priority = 5
    const result = plan(source, current, range(0, 0), [range(3, 3)], 'source-sheet')
    assert.equal(conditionalFormatStyleAt(result.rules, 3, 3, 10).fill, '#ff0000')
  }
})

test('copied expressions shift relative coordinates and retain absolute markers, sheet qualifiers and string literals', () => {
  const sampled = range(1, 1, 2, 2), original = rule('expression', [range(0, 0, 2, 2)], { type: 'expression',
    formula1: '=A1+$B1+C$1+$D$2+\'Other Sheet\'!A1+Sheet2!B$2+IF(A1="A1",1,0)',
  })
  const result = plan(capture([original], sampled), undefined, sampled, [range(3, 3, 4, 4)])
  assert.equal(result.rules[0].formula1, '=D4+$B4+F$1+$D$2+\'Other Sheet\'!D4+Sheet2!E$2+IF(D4="A1",1,0)')
  assert.equal(expressionAt(result.rules[0], 4, 4), '=E5+$B5+G$1+$D$2+\'Other Sheet\'!E5+Sheet2!F$2+IF(E5="A1",1,0)')
})

test('sampling inside a rule whose anchor is outside the brush retains each copied cell expression', () => {
  const sampled = range(2, 2), source = capture([rule('expression', [range(0, 0, 4, 4)], { type: 'expression', formula1: '=B2>A1' })], sampled)
  const result = plan(source, undefined, sampled, [range(5, 5, 6, 5)])
  assert.equal(result.rules[0].formula1, '=G7>F6'); assert.equal(expressionAt(result.rules[0], 6, 5), '=G8>F7')
})

test('copying to earlier coordinates uses the rule anchor rather than the sampled corner', () => {
  const sampled = range(4, 4), source = capture([rule('expression', [range(3, 3, 4, 4)], { type: 'expression', formula1: '=F6+$F6+F$6+$F$6' })], sampled)
  const result = plan(source, undefined, sampled, [range(1, 1)])
  assert.equal(result.rules[0].formula1, '=D4+$F4+D$6+$F$6')
})

test('destination clipping reanchors expressions while retaining every surviving cell meaning', () => {
  const original = rule('target-expression', [range(3, 3, 5, 5)], { type: 'expression', formula1: '=D4>$A$1' })
  const result = plan(capture([]), [original], range(0, 0), [range(3, 3, 3, 4)])
  const clipped = result.rules[0]
  for (const key of cells(clipped.ranges)) { const [row, col] = key.split(':').map(Number); assert.equal(expressionAt(clipped, row, col), expressionAt(original, row, col)) }
  assert.equal(clipped.formula1, '=F4>$A$1')
})

test('same-sheet extension to an earlier anchor retains source and target expression semantics and no-op repainting', () => {
  const sampled = range(3, 3), current = [rule('expression', [sampled], { type: 'expression', formula1: '=D4>5' })], source = capture(current, sampled)
  const result = plan(source, current, sampled, [range(1, 1)], 'source-sheet')
  assert.equal(result.rules.length, 1); assert.equal(result.rules[0].formula1, '=B2>5')
  assert.equal(expressionAt(result.rules[0], 3, 3), '=D4>5')
  assert.equal(plan(source, result.rules, sampled, [range(1, 1)], 'source-sheet').changed, false)
})

test('an invalid copied relative reference does not turn the still-valid source formula into REF', () => {
  const sampled = range(1, 1), current = [rule('expression', [sampled], { type: 'expression', formula1: '=A1>5' })]
  const result = plan(capture(current, sampled), current, sampled, [range(0, 0)], 'source-sheet')
  assert.equal(result.rules.length, 1); assert.equal(result.rules[0].formula1, '=A1>5'); assert.deepEqual(result.rules[0].formulaAnchor, { row: 1, col: 1 })
  assert.deepEqual(cells(result.rules[0].ranges), ['0:0', '1:1'])
  assert.equal(expressionAt(result.rules[0], 0, 0), '=#REF!>5'); assert.equal(expressionAt(result.rules[0], 1, 1), '=A1>5')
})

test('visual formulas shift only formula thresholds while retaining flags, colors, icon IDs and unknown metadata', () => {
  const source = capture([rule('visual', [range(0, 0)], { type: 'iconSet', config: {
    min: { type: 'formula', value: '=A1' }, max: { type: 'formula', value: '=$B$2' },
    thresholds: [{ value: { type: 'formula', value: '=C1' }, color: '#112233' }, { value: { type: 'num', value: 'A1' }, color: '#ffffff' }],
    icons: [{ value: { type: 'formula', value: '=A2' }, operator: 'greaterThan', iconType: '3Arrows', iconId: '0' }, { value: { type: 'min' }, iconType: 'EMPTY_ICON_TYPE', iconId: '' }],
    isShowValue: false, reverse: true, isGradient: true, iconSet: '3Arrows', custom: { keep: ['A1'] },
  }, metadata: { vendor: { keep: true } } })])
  const copied = plan(source).rules[0]
  assert.equal(copied.config.min.value, '=D4'); assert.equal(copied.config.max.value, '=$B$2')
  assert.equal(copied.config.thresholds[0].value.value, '=F4'); assert.equal(copied.config.thresholds[1].value.value, 'A1')
  assert.equal(copied.config.icons[0].value.value, '=D5'); assert.equal(copied.config.icons[1].iconType, 'EMPTY_ICON_TYPE')
  assert.equal(copied.config.reverse, true); assert.equal(copied.config.isShowValue, false); assert.equal(copied.config.isGradient, true)
  assert.deepEqual(copied.config.custom, { keep: ['A1'] }); assert.deepEqual(copied.metadata, { vendor: { keep: true } })
})

test('numeric comparison formulas and formula-subtype values are copied with their relative references', () => {
  const numeric = rule('numeric', [range(0, 0)], { operator: 'between', formula1: '=A1', formula2: '=$B2' })
  const expression = rule('expression', [range(0, 0)], { type: 'highlightCell', subType: 'formula', formula1: '', value: '=A1>0', priority: 2 })
  const result = plan(capture([numeric, expression]))
  assert.equal(result.rules[0].formula1, '=D4'); assert.equal(result.rules[0].formula2, '=$B5')
  assert.equal(result.rules[1].formula1, ''); assert.equal(result.rules[1].value, '=D4>0')
})

test('text, date, rank and duplicate rule literals that resemble addresses remain literal data', () => {
  const source = capture(['containsText', 'timePeriod', 'top10', 'duplicateValues'].map((type, index) => rule(type, [range(0, 0)], {
    type, subType: type === 'containsText' ? 'text' : undefined, formula1: 'A1', value: type === 'top10' ? 10 : 'A1', priority: index + 1,
  })))
  const result = plan(source)
  assert.ok(result.rules.every(entry => entry.formula1 === 'A1'))
  assert.equal(result.rules[0].value, 'A1'); assert.equal(result.rules[2].value, 10)
})

test('copied source priorities and stop-if-true preserve distinct-property style composition', () => {
  const source = capture([rule('second', [range(0, 0)], { priority: 8, style: { fill: '#00ff00', italic: true } }), rule('first', [range(0, 0)], { priority: 2, style: { fill: '#ff0000', bold: true }, stopIfTrue: true })])
  const result = plan(source)
  assert.deepEqual(conditionalFormatStyleAt(result.rules, 3, 3, 10), { fill: '#ff0000', bold: true })
  assert.equal(result.rules[0].stopIfTrue, true); assert.ok(result.rules[0].priority < result.rules[1].priority)
})

test('conditional-only model edits report changes without content mutations or blank cell creation', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]; sheet.conditionalFormats = [rule()]
  const source = captureSheetFormatPainterSource(sheet, range(0, 0)), result = applySheetFormatPainterCommand(sheet, source, [range(3, 3)])
  assert.equal(result.applied, true); assert.equal(result.changed, true); assert.equal(result.contentChanged, false); assert.deepEqual(sheet.cells, {})
})

test('view and domain painters retain the same rule stack through model adaptation', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.conditionalFormats = [rule('expression', [range(0, 0)], { type: 'expression', formula1: '=A1>5', metadata: { keep: true } })]
  sheet.cells['3:3'] = createCell(3, 3, 9)
  const view = toExcelWorkbookView(workbook), modelSource = captureSheetFormatPainterSource(sheet, range(0, 0)), viewSource = captureSheetViewFormatPainterSource(view.sheets[0], range(0, 0))
  assert.deepEqual(cloneModel(viewSource), cloneModel(modelSource))
  applySheetFormatPainterCommand(sheet, modelSource, [range(3, 3)]); applySheetViewFormatPainterCommand(view.sheets[0], viewSource, [range(3, 3)])
  const adapted = applyExcelWorkbookView(workbook, view)
  assert.deepEqual(adapted.sheets[0].conditionalFormats, sheet.conditionalFormats)
  assert.equal(adapted.sheets[0].cells['3:3'].value, 9)
})

test('invalid conditional snapshots reject before any style or merge content changes', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.mergedRanges = [range(0, 0, 0, 1)]; sheet.cells['3:3'] = createCell(3, 3, 1); sheet.cells['3:4'] = createCell(3, 4, 2)
  const source = captureSheetFormatPainterSource(sheet, range(0, 0)), before = cloneModel(sheet)
  const malformed = [
    value => { value.conditional.rules = [rule('invalid', [range(-1, 0)])] },
    value => { value.conditional.rules = [rule(), rule()] },
    value => { value.conditional.rules = [rule('invalid', [range(0, 0)], { config: { icons: [null] } })] },
    value => { value.conditional.rules = [rule('invalid', [range(0, 0)], { config: { min: { type: 'num', value: NaN } } })] },
    value => { value.conditional.rules = [rule('invalid', [range(0, 0)], { metadata: { bad: () => 1 } })] },
    value => { const metadata = {}; metadata.self = metadata; value.conditional.rules = [rule('invalid', [range(0, 0)], { metadata })] },
  ]
  for (const mutate of malformed) { const invalid = cloneModel(source); mutate(invalid); assert.throws(() => applySheetFormatPainterCommand(sheet, invalid, [range(3, 3)], { mergeConfirmed: true })); assert.deepEqual(sheet, before) }
})

test('conditional source validation rejects over-budget rule ranges before model application', () => {
  assert.throws(() => assertConditionalFormatPainterSource({ sheetId: 'source-sheet', rules: [rule('too-many', Array.from({ length: 100_001 }, () => range(0, 0)))] }), /最多同步处理/)
})

test('range work limits reject an oversized rule stack atomically before painting a target', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.conditionalFormats = Array.from({ length: 2500 }, (_, index) => rule('source-' + index, [range(0, 0)], { priority: index + 1 }))
  const source = captureSheetFormatPainterSource(sheet, range(0, 0)), before = cloneModel(sheet)
  assert.throws(() => applySheetFormatPainterCommand(sheet, source, [range(3, 3)]), /同步预算/)
  assert.deepEqual(sheet, before)
})

test('the Canvas scene evaluates copied relative expressions against destination values and updates dynamically', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.rowCount = 20; sheet.columnCount = 10
  sheet.cells['0:0'] = createCell(0, 0, 10); sheet.cells['3:3'] = createCell(3, 3, 7); sheet.cells['4:3'] = createCell(4, 3, 2)
  sheet.conditionalFormats = [rule('expression', [range(0, 0)], { type: 'expression', formula1: '=A1>=5', style: { fill: '#1122bb' } })]
  const source = captureSheetFormatPainterSource(sheet, range(0, 0)); applySheetFormatPainterCommand(sheet, source, [range(3, 3, 4, 3)])
  const viewport = { width: 900, height: 400, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  const fills = () => {
    const geometry = new SpreadsheetGeometry(toExcelWorkbookView(workbook).sheets[0]), rect = geometry.cellRect(3, 3, viewport)
    const scene = createSpreadsheetScene(geometry, viewport, { activeCell: 'J10', selectedCells: [] })
    return scene.commands.filter(command => command.kind === 'rect' && command.x === rect.x && command.y === rect.y).map(command => command.fill)
  }
  assert.ok(fills().includes('#1122bb')); sheet.cells['3:3'].value = 4; assert.equal(fills().includes('#1122bb'), false)
  assert.equal(sheet.cells['0:0'].value, 10); assert.equal(sheet.cells['3:3'].style.fill, null)
})

test('multi-cell row underflow retains the original formula and a valid reference for every later target row', () => {
  const sampled = range(1, 0, 4, 0), current = [rule('source', [sampled], { type: 'expression', formula1: '=A1>5' })], before = cloneModel(current)
  const copied = plan(capture(current, sampled), undefined, sampled, [range(0, 0, 3, 0)]).rules[0]
  assert.deepEqual(current, before); assert.deepEqual(copied.ranges, [range(0, 0, 3, 0)])
  assert.equal(copied.formula1, '=A1>5'); assert.deepEqual(copied.formulaAnchor, { row: 1, col: 0 })
  assert.deepEqual([0, 1, 2, 3].map(row => expressionAt(copied, row, 0)), ['=#REF!>5', '=A1>5', '=A2>5', '=A3>5'])
})

test('multi-cell column underflow recovers each later column while preserving absolute dimensions', () => {
  const sampled = range(0, 1, 0, 4), current = [rule('source', [sampled], { type: 'expression', formula1: '=A1+$B$1>5' })]
  const copied = plan(capture(current, sampled), undefined, sampled, [range(0, 0, 0, 3)]).rules[0]
  assert.deepEqual(copied.formulaAnchor, { row: 0, col: 1 })
  assert.deepEqual([0, 1, 2, 3].map(col => expressionAt(copied, 0, col)), ['=#REF!+$B$1>5', '=A1+$B$1>5', '=B1+$B$1>5', '=C1+$B$1>5'])
})

test('clipping away a target rule origin preserves later valid cells when the new first area overflows a column', () => {
  const original = rule('target', [range(0, 0), range(0, 16_383), range(1, 0, 2, 0)], { type: 'expression', formula1: '=B1>5' })
  const clipped = plan(capture([]), [original], range(9, 0), [range(0, 0)]).rules[0]
  assert.deepEqual(clipped.formulaAnchor, { row: 0, col: 0 }); assert.equal(clipped.priority, original.priority)
  assert.equal(expressionAt(clipped, 0, 16_383), '=#REF!>5')
  assert.equal(expressionAt(clipped, 1, 0), '=B2>5'); assert.equal(expressionAt(clipped, 2, 0), '=B3>5')
  assert.deepEqual(cells(clipped.ranges), ['0:16383', '1:0', '2:0'])
})

test('upper row boundaries remain local to each target cell rather than erasing a valid preceding row', () => {
  const sampled = range(0, 0), source = capture([rule('source', [sampled], { type: 'expression', formula1: '=A2>5' })])
  const copied = plan(source, undefined, sampled, [range(1_048_574, 0, 1_048_575, 0)]).rules[0]
  assert.equal(expressionAt(copied, 1_048_574, 0), '=A1048576>5')
  assert.equal(expressionAt(copied, 1_048_575, 0), '=#REF!>5')
})

test('opposite reference overflows retain independently valid IFERROR branches and one target rule', () => {
  const sampled = range(1, 0), current = [rule('source', [sampled], { type: 'expression', formula1: '=IFERROR(A1,0)+IFERROR(A1048576,0)>5' })]
  const copied = plan(capture(current, sampled), undefined, sampled, [range(0, 0), range(1_048_575, 0)]).rules[0]
  assert.deepEqual(copied.formulaAnchor, { row: 1, col: 0 })
  assert.equal(expressionAt(copied, 0, 0), '=IFERROR(#REF!,0)+IFERROR(A1048575,0)>5')
  assert.equal(expressionAt(copied, 1_048_575, 0), '=IFERROR(A1048575,0)+IFERROR(#REF!,0)>5')
})

test('partial-overflow visual thresholds retain the complete rule statistics without splitting its ranges', () => {
  const sampled = range(1, 0, 4, 0), current = [rule('source', [sampled], { type: 'dataBar', config: { min: { type: 'formula', value: '=A1' }, max: { type: 'max' }, positiveColor: '#123456' } })]
  const result = plan(capture(current, sampled), undefined, sampled, [range(0, 0, 3, 0)])
  assert.equal(result.rules.length, 1); const copied = result.rules[0]
  assert.deepEqual(copied.ranges, [range(0, 0, 3, 0)]); assert.deepEqual(copied.formulaAnchor, { row: 1, col: 0 })
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], values = [4, 50, 1000, 10]
  values.forEach((value, row) => { sheet.cells[`${row}:0`] = createCell(row, 0, value) })
  const scope = values.map((value, row) => ({ row, column: 0, value }))
  const at = row => conditionalFormatStyleAt([copied], row, 0, values[row], [], scope, undefined, (formula, entry) => evaluateFormula(conditionalFormatFormulaAt(formula, entry, row, 0), workbook, sheet.id, { formulaCell: { row, column: 0 } }))
  assert.equal(at(0).dataBar, undefined)
  assert.ok(Math.abs(at(1).dataBar.fraction - (50 - 4) / (1000 - 4)) < 1e-12)
})

test('same-sheet partial-overflow painting reuses the source stack and repeated lowercase formulas add no changes', () => {
  const sampled = range(1, 0, 4, 0), current = [rule('source', [sampled], { type: 'expression', formula1: '=a1>5' })], source = capture(current, sampled)
  const result = plan(source, current, sampled, [range(0, 0, 3, 0)], 'source-sheet')
  assert.equal(result.rules.length, 1); assert.equal(result.rules[0].id, 'source'); assert.equal(result.rules[0].formula1, '=a1>5')
  assert.equal(plan(source, result.rules, sampled, [range(0, 0, 3, 0)], 'source-sheet').changed, false)
})

test('explicit formula origins survive view mapping and are cloned independently in both directions', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.conditionalFormats = [rule('source', [range(0, 0, 3, 0)], { type: 'expression', formula1: '=A1>5', formulaAnchor: { row: 1, col: 0 } })]
  const view = toExcelWorkbookView(workbook), reopened = applyExcelWorkbookView(workbook, view)
  assert.deepEqual(reopened.sheets[0].conditionalFormats, sheet.conditionalFormats)
  view.sheets[0].conditionalFormats[0].formulaAnchor.row = 3
  assert.equal(sheet.conditionalFormats[0].formulaAnchor.row, 1); assert.equal(reopened.sheets[0].conditionalFormats[0].formulaAnchor.row, 1)
})

test('invalid explicit origins reject before model changes and parser normalization preserves valid origins', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], source = captureSheetFormatPainterSource(sheet, range(0, 0)), before = cloneModel(sheet)
  for (const origin of [null, {}, { row: -1, col: 0 }, { row: 1.5, col: 0 }, { row: 0, col: 16_384 }, { row: 1_048_576, col: 0 }, { row: NaN, col: 0 }]) {
    const invalid = cloneModel(source); invalid.conditional.rules = [rule('invalid', [range(0, 0)], { formulaAnchor: origin })]
    assert.throws(() => applySheetFormatPainterCommand(sheet, invalid, [range(2, 2)])); assert.deepEqual(sheet, before)
  }
  const payload = origin => ({ schemaVersion: 1, parserVersion: 'test', format: 'xlsx', fileName: 'origin.xlsx', size: 1, model: 'spreadsheet',
    payload: { modelType: 'spreadsheet', data: { sheets: [{ name: 'Data', cells: [], conditionalFormats: [rule('origin', [range(0, 0, 3, 0)], { type: 'expression', formulaAnchor: origin })] }] } }, diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] } })
  assert.deepEqual(normalizeWorkbook(payload({ row: 1, col: 0 })).sheets[0].conditionalFormats[0].formulaAnchor, { row: 1, col: 0 })
  assert.throws(() => normalizeWorkbook(payload({ row: -1, col: 0 })), /公式基准无效/)
})

test('Canvas paints valid later rows after a partial-overflow copy and responds to referenced value changes', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], sampled = range(1, 0, 4, 0)
  sheet.rowCount = 20; sheet.columnCount = 10
  ;[10, 2, 7, 9].forEach((value, row) => { sheet.cells[`${row}:0`] = createCell(row, 0, value) })
  sheet.conditionalFormats = [rule('source', [sampled], { type: 'expression', formula1: '=A1>=5', style: { fill: '#1122bb' } })]
  applySheetFormatPainterCommand(sheet, captureSheetFormatPainterSource(sheet, sampled), [range(0, 0, 3, 0)])
  const viewport = { width: 900, height: 400, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  const painted = () => {
    const geometry = new SpreadsheetGeometry(toExcelWorkbookView(workbook).sheets[0]), scene = createSpreadsheetScene(geometry, viewport, { activeCell: 'J10', selectedCells: [] })
    return [0, 1, 2, 3].map(row => { const rect = geometry.cellRect(row, 0, viewport); return scene.commands.some(command => command.kind === 'rect' && command.x === rect.x && command.y === rect.y && command.fill === '#1122bb') })
  }
  assert.deepEqual(painted(), [false, true, false, true]); sheet.cells['0:0'].value = 1
  assert.deepEqual(painted(), [false, false, false, true])
})
