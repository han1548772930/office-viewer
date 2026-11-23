import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import { conditionalFormatOfficeBaseline, normalizeConditionalFormatOfficeSource } from '../src/features/conditional-formatting/conditional-format-office-source.ts'
import { applySheetConditionalFormatsCommand, assertSheetConditionalFormatRuleSnapshot } from '../src/features/conditional-formatting/conditional-format-command.ts'
import { conditionalFormatStyleAt, evaluateConditionalFormatRule } from '../src/features/conditional-formatting/conditional-formatting.ts'
import { mapConditionalFormatFormulas, reanchorConditionalFormatRule } from '../src/features/conditional-formatting/conditional-format-formulas.ts'
import { conditionalFormatAverageEditorOperator, conditionalFormatAverageRuleOperator, conditionalFormatEditorText, mergeConditionalFormatRuleDraft } from '../src/features/conditional-formatting/conditional-format-rule-editor.ts'
import { createConditionalIconConfig, validateConditionalThreshold } from '../src/features/conditional-formatting/icon-set-config.ts'

const range = (startRow = 0, startCol = 0, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })
const rule = (changes = {}) => ({ id: 'office-local', type: 'office', ranges: [range(1, 1, 4, 1)],
  operator: 'futureCondition', formula1: '=A1>0', style: { bold: false }, priority: 17, stopIfTrue: true, ...changes })
const source = (baseline = rule()) => ({ version: 1, format: 'xlsx', originalId: baseline.id, storage: 'base',
  ruleXml: '<cfRule xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" type="futureCondition" priority="1"><formula>A1&gt;0</formula><extLst><ext uri="vendor"><v:payload xmlns:v="urn:vendor" reference="A1"/></ext></extLst></cfRule>',
  groupXml: '<conditionalFormatting xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" sqref="B2:B5"/>',
  baseline: conditionalFormatOfficeBaseline(baseline) })
function imported(conditionalFormats) {
  return normalizeWorkbook({ schemaVersion: 1, parserVersion: 'draft', format: 'xlsx', fileName: 'cf-source.xlsx', size: 10, model: 'spreadsheet',
    payload: { modelType: 'spreadsheet', data: { sheets: [{ name: 'Data', cells: [], conditionalFormats }] } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: ['conditionalFormatting'] } })
}

test('baseline keeps inspectable fields and normalizes icon identity without retaining provenance or vendor JSON', () => {
  const original = rule({ metadata: { vendor: true }, vendor: { root: 1 }, officeSource: { nested: true },
    style: { bold: true, vendor: { width: 2 } }, config: { vendor: 3, icons: [{ iconType: '3Arrows', iconId: 2,
      value: { type: 'num', value: 40, vendor: 4 }, vendor: 5 }] } })
  const baseline = conditionalFormatOfficeBaseline(original)
  assert.equal(baseline.id, undefined); assert.equal(baseline.officeSource, undefined); assert.equal(baseline.metadata, undefined)
  assert.deepEqual(baseline.style, { bold: true })
  assert.deepEqual(baseline.config.icons, [{ iconType: '3Arrows', iconId: '2', value: { type: 'num', value: 40 } }])
  baseline.ranges[0].endRow = 99; baseline.config.icons[0].value.value = 90
  assert.equal(original.ranges[0].endRow, 4); assert.equal(original.config.icons[0].value.value, 40)
})

test('source normalization preserves raw XML and source extensions with an isolated known baseline', () => {
  const original = { ...source(), vendor: { archive: ['original'] }, baseline: { ...source().baseline, vendor: 'uninspectable' } }
  const retained = normalizeConditionalFormatOfficeSource(original)
  assert.equal(retained.ruleXml, original.ruleXml); assert.equal(retained.baseline.vendor, undefined)
  retained.vendor.archive.push('copy'); retained.baseline.ranges[0].startRow = 3
  assert.deepEqual(original.vendor.archive, ['original']); assert.equal(original.baseline.ranges[0].startRow, 1)
})

test('invalid baseline scalar, style, formula origin and advanced bar fields reject at the model boundary', () => {
  for (const baseline of [rule({ formula1: 9 }), rule({ priority: 2_147_483_648 }), rule({ stopIfTrue: 'false' }),
    rule({ style: { bold: 'true' } }), rule({ style: { fontColor: 'red' } }), rule({ formulaAnchor: { row: -1, col: 0 } }),
    rule({ config: { minLength: 60, maxLength: 40 } }), rule({ config: { border: 'true' } }),
    rule({ config: { direction: 'vertical' } }), rule({ config: { axisColor: '#xyzxyz' } }),
    rule({ config: { icons: [{ value: { type: 'num', value: { hidden: 2 } } }] } })]) {
    assert.throws(() => normalizeConditionalFormatOfficeSource({ ...source(), baseline }), TypeError)
  }
})

test('parser normalization retains office rules, original IDs, sparse priorities and vendor fields across repeated snapshots', () => {
  const original = rule({ officeSource: source(), vendor: { retained: ['source'] } })
  const first = imported([original]).sheets[0].conditionalFormats[0]
  const second = imported([first]).sheets[0].conditionalFormats[0]
  assert.equal(second.type, 'office'); assert.equal(second.id, original.id); assert.equal(second.priority, 17)
  assert.equal(second.officeSource.ruleXml, original.officeSource.ruleXml); assert.deepEqual(second.vendor, original.vendor)
  assert.equal(second.officeSource.baseline.officeSource, undefined)
  second.officeSource.baseline.ranges[0].endRow = 99; second.vendor.retained.push('copy')
  assert.equal(original.officeSource.baseline.ranges[0].endRow, 4); assert.deepEqual(original.vendor.retained, ['source'])
})

test('an office rule never invokes expression or numeric formula callbacks and cannot stop known lower rules', () => {
  const opaque = rule({ officeSource: source(), subType: 'formula', style: { fill: '#ff0000' } })
  let evaluated = 0
  assert.equal(evaluateConditionalFormatRule(opaque, 9, [9], () => { evaluated++; return true }, () => { evaluated++; return 0 }), false)
  const known = { ...rule(), id: 'known', type: 'cellIs', operator: 'greaterThan', formula1: '5', priority: 18, style: { fill: '#336699' } }
  assert.deepEqual(conditionalFormatStyleAt([opaque, known], 1, 1, 9, [9], [], () => { evaluated++; return true }), { fill: '#336699' })
  assert.equal(evaluated, 0)
})

test('office range, style and order edits use the real command while condition edits reject atomically', () => {
  const sheet = createBlankWorkbook().sheets[0], original = rule({ officeSource: source() })
  sheet.conditionalFormats = [original]
  const changed = { ...original, ranges: [range(8, 3, 11, 3)], priority: 4, style: { bold: true }, stopIfTrue: false }
  const plan = applySheetConditionalFormatsCommand(sheet, [changed], 'local-user')
  assert.equal(plan.changed, true); assert.deepEqual(plan.ranges, [...original.ranges, ...changed.ranges])
  const before = JSON.stringify(sheet.conditionalFormats)
  for (const condition of [{ formula1: '=A1>2' }, { operator: 'otherCondition' }, { config: { min: { type: 'num', value: 2 } } }]) {
    assert.throws(() => applySheetConditionalFormatsCommand(sheet, [{ ...changed, ...condition }], 'local-user'), /条件需要保留/)
    assert.equal(JSON.stringify(sheet.conditionalFormats), before)
  }
  assert.equal(applySheetConditionalFormatsCommand(sheet, [], 'local-user').changed, true)
})

test('unknown XML is cloned unchanged by formula mapping and reanchoring retains its origin', () => {
  const original = rule({ officeSource: source() }), references = []
  const mapped = mapConditionalFormatFormulas(original, formula => { references.push(formula); return '=CHANGED()' })
  assert.deepEqual(references, []); assert.equal(mapped.formula1, original.formula1)
  const moved = reanchorConditionalFormatRule(original, [range(9, 3, 12, 3)])
  assert.deepEqual(moved.formulaAnchor, { row: 1, col: 1 }); assert.equal(moved.officeSource.ruleXml, original.officeSource.ruleXml)
  moved.officeSource.baseline.ranges[0].startRow = 99; assert.equal(original.officeSource.baseline.ranges[0].startRow, 1)
})

test('text editor preserves an unchanged imported predicate and rebuilds text when switching condition types', () => {
  const previous = rule({ type: 'containsText', subType: 'text', value: 'A1', formula1: '=ISNUMBER(SEARCH("A1",B2))' })
  assert.deepEqual(conditionalFormatEditorText('A1', previous), { formula1: previous.formula1, value: 'A1' })
  assert.deepEqual(conditionalFormatEditorText('A2', previous), { formula1: 'A2', value: 'A2' })
  assert.deepEqual(conditionalFormatEditorText('A1', rule({ type: 'expression', value: 'A1' })), { formula1: 'A1', value: 'A1' })
})

test('standard text predicates move their references while matching A1 literals and vendor data stay literal', () => {
  const original = rule({ type: 'containsText', subType: 'text', value: 'A1', formula1: '=ISNUMBER(SEARCH("A1",B2))',
    officeSource: source(), config: { vendor: { ref: 'B2' } } })
  const moved = reanchorConditionalFormatRule(original, [range(2, 2, 5, 2)])
  assert.equal(moved.formula1, '=ISNUMBER(SEARCH("A1",C3))'); assert.equal(moved.value, 'A1')
  assert.equal(moved.config.vendor.ref, 'B2'); assert.equal(moved.officeSource.ruleXml, original.officeSource.ruleXml)
  const literal = mapConditionalFormatFormulas({ ...original, formula1: 'A1', value: 'A1' }, () => 'changed')
  assert.equal(literal.formula1, 'A1'); assert.equal(literal.value, 'A1')
})

test('advanced bar fields remain in editor snapshots and invalid lengths, flags or colors reject as command input', () => {
  const original = rule({ type: 'dataBar', operator: '', formula1: '', stopIfTrue: false, config: {
    min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#336699', nativeColor: '#aa3344',
    minLength: 10, maxLength: 95, border: true, borderColor: '#112233', negativeBorderColor: '#445566',
    axisColor: '#000000', direction: 'rightToLeft', axisPosition: 'middle', vendor: { retained: true } } })
  const edited = mergeConditionalFormatRuleDraft(original, { ...original, config: { positiveColor: '#778899' } })
  assert.equal(edited.config.borderColor, '#112233'); assert.equal(edited.config.direction, 'rightToLeft')
  assert.deepEqual(edited.config.vendor, { retained: true }); assertSheetConditionalFormatRuleSnapshot(edited, original)
  for (const config of [{ minLength: 101 }, { border: 1 }, { axisColor: '#zzzzzz' }, { axisPosition: 'outside' }]) {
    assert.throws(() => assertSheetConditionalFormatRuleSnapshot({ ...edited, config: { ...edited.config, ...config } }, original))
  }
})

test('inclusive and not-equal average rules exclude nonfinite values instead of treating them as matches', () => {
  const average = rule({ type: 'average', subType: 'average', operator: 'greaterThanOrEqual', formula1: '', stopIfTrue: false })
  assert.equal(evaluateConditionalFormatRule(average, 2, [1, 2, 3]), true)
  assert.equal(evaluateConditionalFormatRule({ ...average, operator: 'lessThanOrEqual' }, 2, [1, 2, 3]), true)
  assert.equal(evaluateConditionalFormatRule({ ...average, operator: 'notEqual' }, 1, [1, 2, 3]), true)
  assert.equal(evaluateConditionalFormatRule({ ...average, operator: 'notEqual' }, 2, [1, 2, 3]), false)
  for (const value of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) assert.equal(evaluateConditionalFormatRule(average, value, [1, 2, 3]), false)
})

test('average editor comparisons retain their condition and upstream subtype through a style command', () => {
  const cases = [
    ['greaterThan', [false, false, true]], ['greaterThanOrEqual', [false, true, true]],
    ['lessThan', [true, false, false]], ['lessThanOrEqual', [true, true, false]],
    ['equal', [false, true, false]], ['notEqual', [true, false, true]],
  ]
  for (const type of ['average', 'highlightCell']) for (const [operator, matches] of cases) {
    const previous = rule({ type, subType: 'average', operator, formula1: '', stopIfTrue: false })
    const sheet = createBlankWorkbook().sheets[0]
    sheet.conditionalFormats = [previous]
    const editorOperator = conditionalFormatAverageEditorOperator(previous.operator)
    assert.ok(editorOperator)
    const next = mergeConditionalFormatRuleDraft(previous, { ...previous, type: 'average',
      operator: conditionalFormatAverageRuleOperator(editorOperator), style: { bold: true } })
    const plan = applySheetConditionalFormatsCommand(sheet, [next], 'local-user')
    assert.equal(plan.changed, true)
    assert.equal(sheet.conditionalFormats[0].type, type)
    assert.equal(sheet.conditionalFormats[0].subType, 'average')
    assert.equal(sheet.conditionalFormats[0].operator, operator)
    assert.equal(sheet.conditionalFormats[0].value, undefined)
    assert.deepEqual([1, 2, 3].map(value => evaluateConditionalFormatRule(sheet.conditionalFormats[0], value, [1, 2, 3])), matches)
  }
  assert.equal(conditionalFormatAverageRuleOperator('isBottom'), undefined)
  assert.equal(conditionalFormatAverageEditorOperator('between'), undefined)
})

test('automatic data-bar boundaries remain editable, include zero and reject invalid contexts atomically', () => {
  const original = rule({ type: 'dataBar', ranges: [range(0, 0, 2, 0)], operator: '', formula1: '', stopIfTrue: false,
    config: { min: { type: 'autoMin' }, max: { type: 'autoMax' }, positiveColor: '#336699', nativeColor: '#aa3344',
      isShowValue: false, vendor: { retained: true } } })
  const sheet = imported([original]).sheets[0], previous = sheet.conditionalFormats[0]
  const edited = mergeConditionalFormatRuleDraft(previous, { ...previous, config: { positiveColor: '#778899' } })
  assert.equal(applySheetConditionalFormatsCommand(sheet, [edited], 'local-user').changed, true)
  const current = sheet.conditionalFormats[0]
  assert.deepEqual(current.config.min, { type: 'autoMin' }); assert.deepEqual(current.config.max, { type: 'autoMax' })
  assert.deepEqual(current.config.vendor, { retained: true })
  for (const type of ['autoMin', 'autoMax']) {
    assert.doesNotThrow(() => validateConditionalThreshold({ type }, true))
    assert.throws(() => validateConditionalThreshold({ type }), /阈值类型无效/)
  }
  const cases = [
    { values: [5, 10], index: 0, fraction: 0.5, start: 0, negative: false },
    { values: [-10, -5], index: 1, fraction: 0.5, start: 1, negative: true },
    { values: [-10, -5, 10], index: 1, fraction: 0.25, start: 0.5, negative: true },
  ]
  for (const example of cases) {
    const scope = [...example.values.map((value, row) => ({ row, column: 0, value })), { row: 3, column: 0, value: 1_000_000 }]
    const visual = conditionalFormatStyleAt(sheet.conditionalFormats, example.index, 0, example.values[example.index], [...example.values, 1_000_000], scope).dataBar
    assert.ok(visual)
    assert.equal(visual.fraction, example.fraction); assert.equal(visual.start, example.start); assert.equal(visual.negative, example.negative)
    assert.equal(visual.color, example.negative ? '#aa3344' : '#778899'); assert.equal(visual.showValue, false)
  }
  assert.equal(conditionalFormatStyleAt(sheet.conditionalFormats, 0, 0, 0, [0, 0]).dataBar, undefined)
  const regular = { ...current, config: { ...current.config, min: { type: 'min' }, max: { type: 'max' } } }
  assert.equal(conditionalFormatStyleAt([regular], 0, 0, 5, [5, 10]).dataBar.fraction, 0)

  const invalidRules = [
    { ...current, config: { ...current.config, min: { type: 'autoMax' } } },
    { ...current, config: { ...current.config, max: { type: 'autoMin' } } },
    { ...current, type: 'colorScale', config: { ...current.config, thresholds: [
      { value: { type: 'min' }, color: '#112233' }, { value: { type: 'max' }, color: '#aabbcc' },
    ] } },
  ]
  for (const type of ['autoMin', 'autoMax']) {
    invalidRules.push({ ...current, type: 'colorScale', config: { thresholds: [
      { value: { type }, color: '#112233' }, { value: { type: 'max' }, color: '#aabbcc' },
    ] } })
    const icons = createConditionalIconConfig('3Arrows')
    icons[0].value = { type }
    invalidRules.push({ ...current, type: 'iconSet', config: { iconSet: '3Arrows', icons } })
  }
  const before = JSON.stringify(sheet.conditionalFormats)
  for (const invalid of invalidRules) {
    assert.throws(() => applySheetConditionalFormatsCommand(sheet, [invalid], 'local-user'), /自动阈值|阈值类型无效/)
    assert.equal(JSON.stringify(sheet.conditionalFormats), before)
  }
})
