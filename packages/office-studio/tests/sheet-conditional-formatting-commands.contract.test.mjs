import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { applySheetConditionalFormatsCommand, assertSheetConditionalFormatRuleSnapshot, planSheetConditionalFormatsCommand,
  SET_SHEET_CONDITIONAL_FORMATS_COMMAND as command, sheetConditionalFormatsCommandInput } from '../src/features/conditional-formatting/conditional-format-command.ts'
import { conditionalFormatFormulaAt, reanchorConditionalFormatRule } from '../src/features/conditional-formatting/conditional-format-formulas.ts'
import { conditionalFormatRuleListForScope, mergeConditionalFormatRuleDraft, moveConditionalFormatRuleDraft,
  prependConditionalFormatRuleDraft, removeConditionalFormatRuleDraft } from '../src/features/conditional-formatting/conditional-format-rule-editor.ts'
import { conditionalFormatStyleAt, evaluateConditionalFormatRule } from '../src/features/conditional-formatting/conditional-formatting.ts'
import { evaluateFormula } from '../src/features/formula/formula-engine.ts'
import {
  QUICK_CONDITIONAL_FORMAT_COMMANDS as quick, QUICK_CONDITIONAL_ICON_SETS, quickConditionalFormatRules, quickConditionalFormatSelection,
} from '../src/features/conditional-formatting/quick-conditional-formatting.ts'

const range = (startRow, startCol, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })
const rule = (id = 'cf-first', ranges = [range(0, 0, 3, 0)], priority = 1) => ({ id, type: 'cellIs', ranges, priority,
  operator: 'greaterThan', formula1: '5', style: { fill: '#336699' }, stopIfTrue: false })
const protection = ranges => [{ id: 'cf-locked', permissionId: 'cf-locked-p', ownerUserId: 'owner', allowedUserIds: [],
  editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges }]
function sample() {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'cf-edit-contract'; workbook.activeSheetId = sheet.id = 'cf-edit-main'; workbook.calculationMode = 'manual'
  sheet.rowCount = 200; sheet.columnCount = 26
  for (const [row, value] of [[0, 7], [1, 2], [2, 9], [3, 6]]) sheet.cells[`${row}:0`] = createCell(row, 0, value)
  sheet.cells['10:13'] = { ...createCell(10, 13, 999), formula: '=1+1', dataType: 'formula', cachedDataType: 'number', display: '999' }
  return workbook
}

test('one rule-list edit plans both old and new scopes and changes the conditional result without base-cell styling', () => {
  const sheet = sample().sheets[0], original = cloneModel(sheet.cells)
  const first = applySheetConditionalFormatsCommand(sheet, [rule()], 'local-user')
  assert.equal(first.changed, true); assert.deepEqual(first.ranges, [range(0, 0, 3, 0)])
  assert.equal(conditionalFormatStyleAt(sheet.conditionalFormats, 0, 0, 7).fill, '#336699')
  const moved = { ...rule(), ranges: [range(5, 2, 8, 2)], style: { fill: '#996633' } }
  const plan = applySheetConditionalFormatsCommand(sheet, [moved], 'local-user')
  assert.deepEqual(plan.ranges, [range(0, 0, 3, 0), range(5, 2, 8, 2)])
  assert.deepEqual(conditionalFormatStyleAt(sheet.conditionalFormats, 0, 0, 7), {})
  assert.equal(conditionalFormatStyleAt(sheet.conditionalFormats, 5, 2, 7).fill, '#996633')
  assert.deepEqual(sheet.cells, original)
})

test('unchanged property order and distinct-priority list order are no-ops with no snapshot replacement', () => {
  const sheet = sample().sheets[0], first = rule(), second = rule('second', [range(8, 8)], 2)
  sheet.conditionalFormats = [second, first]
  const original = sheet.conditionalFormats
  const reordered = [{ style: first.style, formula1: first.formula1, ...first }, second]
  const plan = applySheetConditionalFormatsCommand(sheet, reordered, 'local-user')
  assert.equal(plan.changed, false); assert.deepEqual(plan.ranges, []); assert.equal(sheet.conditionalFormats, original)
  plan.rules[0].style.fill = '#ffffff'; assert.equal(first.style.fill, '#336699')
})

test('tied priorities preserve stable stack order and a reorder changes the actual higher-priority fill', () => {
  const sheet = sample().sheets[0], first = { ...rule(), stopIfTrue: true }, second = { ...rule('second'), style: { fill: '#993366' } }
  sheet.conditionalFormats = [first, second]
  assert.equal(conditionalFormatStyleAt(sheet.conditionalFormats, 0, 0, 7).fill, '#336699')
  const plan = applySheetConditionalFormatsCommand(sheet, [second, first], 'local-user')
  assert.equal(plan.changed, true); assert.deepEqual(plan.ranges, [range(0, 0, 3, 0)])
  assert.deepEqual(sheet.conditionalFormats.map(item => item.id), ['second', 'cf-first'])
  assert.equal(conditionalFormatStyleAt(sheet.conditionalFormats, 0, 0, 7).fill, '#993366')
})

test('deleting or inserting an unrelated rule does not treat shifted protected entries as edits', () => {
  const sheet = sample().sheets[0], locked = rule('locked', [range(9, 9)], 2)
  sheet.conditionalFormats = [rule(), locked]; sheet.rangeProtectionRules = protection(locked.ranges)
  const removed = applySheetConditionalFormatsCommand(sheet, [locked], 'local-user')
  assert.deepEqual(removed.ranges, [range(0, 0, 3, 0)]); assert.equal(removed.changed, true)
  const added = applySheetConditionalFormatsCommand(sheet, [rule('new', [range(5, 2)]), locked], 'local-user')
  assert.deepEqual(added.ranges, [range(5, 2)]); assert.equal(added.changed, true)
  assert.deepEqual(sheet.conditionalFormats[1], locked)
})

test('range removal and movement validate the old protected scope and reject atomically', () => {
  for (const next of [[], [{ ...rule(), ranges: [range(6, 3)] }]]) {
    const sheet = sample().sheets[0]; sheet.conditionalFormats = [rule()]; sheet.rangeProtectionRules = protection([range(2, 0)])
    const original = cloneModel(sheet)
    assert.throws(() => applySheetConditionalFormatsCommand(sheet, next, 'local-user'), /保护/)
    assert.deepEqual(sheet, original)
  }
})

test('new protected targets and tied-priority changes fail before any rule is written', () => {
  for (const reorder of [false, true]) {
    const sheet = sample().sheets[0], first = rule(), second = rule('locked', [range(9, 9)])
    sheet.conditionalFormats = [first, second]; sheet.rangeProtectionRules = protection(second.ranges)
    const original = cloneModel(sheet), next = reorder ? [second, first] : [{ ...first, ranges: second.ranges }, second]
    assert.throws(() => applySheetConditionalFormatsCommand(sheet, next, 'local-user'), /保护/)
    assert.deepEqual(sheet, original)
  }
})

test('worksheet viewing and editing permissions are enforced even for an unchanged rule list', () => {
  for (const viewState of ['othersCanView', 'noOneElseCanView']) {
    const sheet = sample().sheets[0]; sheet.conditionalFormats = [rule()]
    sheet.worksheetProtection = { id: 'sheet-lock', permissionId: 'sheet-lock-p', unitType: 'worksheet', ownerUserId: 'owner',
      allowedUserIds: [], editState: 'onlyMe', viewState, description: '' }
    const original = cloneModel(sheet)
    assert.throws(() => applySheetConditionalFormatsCommand(sheet, [rule()], 'local-user'), /权限/)
    assert.deepEqual(sheet, original)
  }
})

test('partial-reference edits retain one complete rule, explicit origin and per-cell valid results', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  const original = { ...rule(), type: 'expression', formula1: '=A1>5', ranges: [range(1, 0, 4, 0)], metadata: { vendor: { kept: ['origin'] } } }
  sheet.conditionalFormats = [original]
  const moved = reanchorConditionalFormatRule(original, [range(0, 0, 3, 0)])
  const result = applySheetConditionalFormatsCommand(sheet, [moved], 'local-user'), current = sheet.conditionalFormats[0]
  assert.equal(result.changed, true); assert.equal(sheet.conditionalFormats.length, 1)
  assert.deepEqual(current.formulaAnchor, { row: 1, col: 0 }); assert.equal(current.formula1, '=A1>5')
  assert.deepEqual([0, 1, 2, 3].map(row => evaluateConditionalFormatRule(current, sheet.cells[`${row}:0`].value, [], formula =>
    evaluateFormula(conditionalFormatFormulaAt(formula, current, row, 0), workbook, sheet.id, { formulaCell: { row, column: 0 } }) === true)), [false, true, false, true])
  moved.metadata.vendor.kept.push('caller changed'); assert.deepEqual(current.metadata.vendor.kept, ['origin'])
})

test('unknown rule, style, config and threshold fields survive isolated command snapshots', () => {
  const sheet = sample().sheets[0]
  const source = { ...rule(), type: 'dataBar', formulaAnchor: { row: 1, col: 0 }, vendor: { root: ['keep'] },
    style: { fill: '#336699', vendor: { font: ['keep'] } }, config: { min: { type: 'formula', value: '=A1', vendor: { min: ['keep'] } },
      max: { type: 'num', value: 20 }, positiveColor: '#336699', custom: { keep: ['config'] } }, metadata: { keep: ['metadata'] } }
  applySheetConditionalFormatsCommand(sheet, [source], 'local-user')
  source.vendor.root.push('changed'); source.config.min.vendor.min.push('changed'); source.config.custom.keep.push('changed')
  const stored = sheet.conditionalFormats[0]
  assert.deepEqual(stored.vendor.root, ['keep']); assert.deepEqual(stored.config.min.vendor.min, ['keep'])
  assert.deepEqual(stored.config.custom.keep, ['config']); assert.deepEqual(stored.formulaAnchor, { row: 1, col: 0 })
  const view = toExcelWorkbookView({ ...sample(), sheets: [sheet] }), reopened = applyExcelWorkbookView({ ...sample(), sheets: [sheet] }, view).sheets[0]
  view.sheets[0].conditionalFormats[0].style.vendor.font.push('view'); reopened.conditionalFormats[0].metadata.keep.push('reopened')
  assert.deepEqual(stored.style.vendor.font, ['keep']); assert.deepEqual(stored.metadata.keep, ['metadata'])
})

test('invalid formulas, origins, priorities, duplicate ids and visual configurations reject a whole edit', () => {
  const malformed = [
    { ...rule(), formula1: '=SUM(' }, { ...rule(), operator: 'between' }, { ...rule(), priority: 0 },
    { ...rule(), formulaAnchor: { row: 1_048_576, col: 0 } }, { ...rule(), ranges: [range(0, 16_384)] },
    { ...rule(), subType: 'not-a-condition' }, { ...rule(), type: 'expression', formula1: '=AND(' },
    { ...rule(), type: 'dataBar', config: { min: { type: 'formula', value: '=SUM(' } } },
    { ...rule(), type: 'dataBar', config: { min: null } }, { ...rule(), type: 'colorScale', config: { thresholds: [] } },
    { ...rule(), type: 'iconSet', config: { icons: [] } }, { ...rule(), type: 'top10', value: 101, isPercent: true },
  ]
  const sheet = sample().sheets[0], before = cloneModel(sheet)
  for (const invalid of malformed) {
    assert.throws(() => applySheetConditionalFormatsCommand(sheet, [rule('valid'), invalid], 'local-user'))
    assert.deepEqual(sheet, before)
  }
  assert.throws(() => planSheetConditionalFormatsCommand(sheet, [rule(), rule()]), /重复/)
  assert.doesNotThrow(() => assertSheetConditionalFormatRuleSnapshot({ ...rule(), ranges: [range(1_048_575, 16_383)], formulaAnchor: { row: 1_048_575, col: 16_383 } }))
})

test('unedited imported formula grammar is retained while changed malformed formulas are rejected', () => {
  for (const kind of [{ type: 'expression' }, { type: 'highlightCell', subType: 'formula' }, { type: 'cellIs' }]) {
    const sheet = sample().sheets[0], imported = { ...rule(), ...kind, formula1: '=UNSUPPORTED([VendorColumn])', vendor: { preserved: true } }
    sheet.conditionalFormats = [imported]
    const changedStyle = { ...imported, style: { fill: '#112233' } }
    assert.doesNotThrow(() => applySheetConditionalFormatsCommand(sheet, [changedStyle], 'local-user'))
    assert.equal(sheet.conditionalFormats[0].formula1, imported.formula1)
    assert.throws(() => applySheetConditionalFormatsCommand(sheet, [{ ...changedStyle, formula1: '=SUM(' }], 'local-user'), /公式无效/)
  }
})

test('command input validates explicit ownership, aliases and optimistic revision without mutating parameters', () => {
  for (const input of [undefined, null, [], {}, { rules: null }, { rules: [], unitId: '' }, { rules: [], sheetId: 'a', subUnitId: 'b' },
    { rules: [], expectedRevision: -1 }, { rules: [], expectedRevision: 1.5 }, { rules: [], expectedHistoryGeneration: 0 },
    { rules: [], expectedHistoryGeneration: -1 }, { rules: [], expectedHistoryGeneration: 1.5 }, { rules: [], expectedHistoryGeneration: null },
    { rules: [], expectedUserId: '' }, { rules: [], expectedUserId: '   ' }, { rules: [], expectedUserId: 12 }, { rules: [], host: 'unknown' }]) {
    assert.throws(() => sheetConditionalFormatsCommandInput(input))
  }
  const input = { unitId: 'workbook', subUnitId: 'sheet', rules: [rule()], expectedRevision: 0,
    expectedHistoryGeneration: 1, expectedUserId: 'local-user', host: 'workbench' }
  assert.equal(sheetConditionalFormatsCommandInput(input), input)
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
function withStudio(run, workbook = sample()) {
  const studio = new OfficeStudio({ workbook, currentUserId: 'local-user' })
  try { return run(studio, FOffice.newAPI(studio)) } finally { studio.dispose() }
}
const undoCount = (studio, id = studio.activeWorkbookId) => studio.editor.history.units[id]?.undo.length ?? 0
const input = (studio, rules, extra = {}) => ({ unitId: studio.activeWorkbookId, sheetId: studio.editor.workbook.model.sheets[0].id, rules, ...extra })

test('quick visual rules render fixed bar colors and the three ranked icon sets across one sparse scope', () => {
  const scope = [range(0, 0, 3, 0), range(3, 3)]
  const [bar] = quickConditionalFormatRules([], scope, quick.dataBar, {})
  assert.deepEqual(bar.config, {
    min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#ffbe38', nativeColor: '#abd91a', isGradient: false, isShowValue: true,
  })
  const values = [-100, 0, 100]
  const negative = conditionalFormatStyleAt([bar], 0, 0, -100, values).dataBar
  assert.deepEqual(negative, { color: '#abd91a', fraction: .5, start: .5, negative: true, gradient: false, showValue: true })
  assert.equal(conditionalFormatStyleAt([bar], 3, 3, 100, values).dataBar.color, '#ffbe38')
  for (const iconSet of QUICK_CONDITIONAL_ICON_SETS) {
    const rules = quickConditionalFormatRules([bar], scope, quick.iconSet, { iconSet })
    assert.equal(rules[0].config.iconSet, iconSet)
    assert.deepEqual(rules[0].config.icons.slice(0, 2).map(icon => [icon.operator, icon.value]),
      [['greaterThanOrEqual', { type: 'percent', value: 67 }], ['greaterThanOrEqual', { type: 'percent', value: 33 }]])
    assert.deepEqual([100, 67, 33, 32].map(value => conditionalFormatStyleAt(rules, 0, 0, value, [0, 100]).icon?.iconId), ['0', '0', '1', '2'])
    assert.deepEqual(quickConditionalFormatSelection(rules, scope, { row: 0, col: 0 }),
      { hasDataBar: true, hasIconSet: true, activeIconSet: iconSet, canClearDataBar: true, canClearIconSet: true, canClearRange: true, canClearWorksheet: true })
    assert.deepEqual(conditionalFormatStyleAt(rules, 2, 2, 100, [0, 100]), {})
  }
})

test('type-specific clear preserves remaining coverage and rebases threshold formulas', () => {
  const scope = [range(1, 0, 4, 0), range(6, 2)]
  const bars = quickConditionalFormatRules([rule('other', scope, 3)], scope, quick.dataBar, { min: { type: 'formula', value: '=A2' } })
  const rules = quickConditionalFormatRules(bars, scope, quick.iconSet, {})
  const after = quickConditionalFormatRules(rules, [range(1, 0), range(6, 2)], quick.clear, { types: ['dataBar'] })
  assert.deepEqual(after.find(entry => entry.type === 'dataBar').ranges, [range(2, 0, 4, 0)])
  assert.equal(after.find(entry => entry.type === 'dataBar').config.min.value, '=A3')
  assert.deepEqual(after.filter(entry => entry.type !== 'dataBar'), rules.filter(entry => entry.type !== 'dataBar'))
  const all = quickConditionalFormatRules(after, scope, quick.clear, {})
  assert.deepEqual(all, [])
})

test('whole-sheet clear removes all rule types without requiring a selected range', () => {
  const scope = [range(0, 0, 3, 0)]
  const rules = quickConditionalFormatRules([rule('expression', [range(10, 13)])], scope, quick.dataBar, {})
  const before = cloneModel(rules)
  assert.deepEqual(quickConditionalFormatRules(rules, [], quick.clearWorksheet, {}), [])
  assert.deepEqual(rules, before)
  assert.equal(quickConditionalFormatRules([], [], quick.clearWorksheet, {}).length, 0)
})

test('management clear availability includes nonvisual rules and all old protected scopes', () => {
  const current = rule('current', [range(0, 0, 3, 0)]), outside = rule('outside', [range(10, 13)], 2)
  const canEdit = ranges => ranges.every(area => area.startCol !== 13)
  const selected = quickConditionalFormatSelection([current, outside], [range(0, 0)], { row: 0, col: 0 }, canEdit)
  assert.equal(selected.canClearRange, true)
  assert.equal(selected.canClearWorksheet, false)
  assert.equal(selected.canClearDataBar, false)
  assert.equal(selected.canClearIconSet, false)
  assert.equal(quickConditionalFormatSelection([current, outside], [range(10, 13)], { row: 10, col: 13 }, canEdit).canClearRange, false)
  assert.equal(quickConditionalFormatSelection([current], [], undefined, canEdit).canClearWorksheet, true)
})

test('registered whole-sheet clear uses one reversible transaction and retains cells and snapshot reopen', () => {
  const source = sample(), rules = [rule(), rule('other-type', [range(10, 13)], 2)]
  source.sheets[0].conditionalFormats = rules
  withStudio((studio, api) => {
    const workbook = api.getActiveWorkbook(), worksheet = workbook.getActiveSheet(), before = cloneModel(studio.editor.workbook.model.sheets[0].cells)
    assert.equal(api.executeCommandSync(quick.clearWorksheet), true)
    assert.deepEqual(worksheet.getConditionalFormats(), [])
    assert.equal(undoCount(studio), 1)
    workbook.undo(); assert.deepEqual(worksheet.getConditionalFormats(), rules)
    workbook.redo(); assert.deepEqual(worksheet.getConditionalFormats(), [])
    const revision = studio.getWorkbookRevision(workbook.getId())
    assert.equal(api.executeCommandSync(quick.clearWorksheet), false)
    assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
    assert.equal(undoCount(studio), 1)
    assert.deepEqual(studio.editor.workbook.model.sheets[0].cells, before)
    const reopened = api.createWorkbook({ ...workbook.save(), id: 'cf-sheet-cleared-reopened' })
    assert.deepEqual(reopened.getActiveSheet().getConditionalFormats(), [])
  }, source)
})

test('whole-sheet clear rejects any protected old scope and isolates explicit inactive targets', () => {
  const source = sample(), locked = rule('locked', [range(10, 13)], 2)
  source.sheets[0].conditionalFormats = [rule(), locked]
  source.sheets[0].rangeProtectionRules = protection(locked.ranges)
  withStudio((studio, api) => {
    const original = cloneModel(studio.editor.workbook.model), revision = studio.getWorkbookRevision(source.id)
    assert.throws(() => api.executeCommandSync(quick.clearWorksheet, { ranges: [range(0, 0)] }), /保护/)
    assert.deepEqual(studio.editor.workbook.model, original)
    assert.equal(studio.getWorkbookRevision(source.id), revision)
    assert.equal(undoCount(studio), 0)
    const other = sample(); other.id = 'cf-clear-other'; other.activeSheetId = other.sheets[0].id = 'cf-clear-other-sheet'
    other.sheets[0].conditionalFormats = [rule()]
    studio.createWorkbook(other, false)
    assert.equal(api.executeCommandSync(quick.clearWorksheet, { unitId: other.id, subUnitId: other.sheets[0].id, ranges: [] }), true)
    assert.deepEqual(studio.editor.workbook.model, original)
    assert.equal(undoCount(studio), 0)
    assert.equal(undoCount(studio, other.id), 1)
    assert.deepEqual(studio.getWorkbookModel(other.id).sheets[0].conditionalFormats, [])
  }, source)
})

test('registered quick commands share permission checks, exact rule history and snapshot reopen', () => withStudio((studio, api) => {
  const workbook = api.getActiveWorkbook(), worksheet = workbook.getActiveSheet()
  const before = cloneModel(studio.getWorkbookModel(workbook.getId()).sheets[0].cells)
  const target = { unitId: workbook.getId(), subUnitId: worksheet.getSheetId(), ranges: [{ startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 }] }
  assert.equal(api.executeCommandSync(quick.dataBar, target), true)
  assert.equal(api.executeCommandSync(quick.iconSet, { ...target, iconSet: '3TrafficLights1' }), true)
  assert.equal(undoCount(studio), 2)
  const rules = worksheet.getConditionalFormats()
  assert.deepEqual(rules.map(rule => rule.type), ['iconSet', 'dataBar'])
  assert.equal(api.executeCommandSync(quick.clear, { ...target, ranges: [range(0, 1)], types: ['iconSet'] }), false)
  assert.equal(api.executeCommandSync(quick.clear, { ...target, ranges: [range(1, 0)], types: ['iconSet'] }), true)
  const cleared = worksheet.getConditionalFormats()
  assert.deepEqual(cleared[0].ranges, [range(0, 0), range(2, 0, 3, 0)])
  assert.equal(undoCount(studio), 3)
  workbook.undo(); assert.deepEqual(worksheet.getConditionalFormats(), rules)
  workbook.redo(); assert.deepEqual(worksheet.getConditionalFormats(), cleared)
  assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].cells, before)
  const reopened = api.createWorkbook({ ...workbook.save(), id: 'cf-quick-reopened' })
  assert.deepEqual(reopened.getActiveSheet().getConditionalFormats(), cleared)
}))

test('quick command rejects stale and protected rule footprints without publishing a partial change', () => {
  const workbook = sample(), sheet = workbook.sheets[0], scope = [range(0, 0, 3, 0)]
  sheet.conditionalFormats = quickConditionalFormatRules([], scope, quick.dataBar, {})
  sheet.rangeProtectionRules = protection([range(2, 0)])
  withStudio((studio, api) => {
    const worksheet = api.getActiveWorkbook().getActiveSheet()
    const before = worksheet.getConditionalFormats()
    const target = { unitId: studio.activeWorkbookId, subUnitId: worksheet.getSheetId() }
    assert.throws(() => api.executeCommandSync(quick.clear, { ...target, ranges: [range(0, 0)], types: ['dataBar'] }), /保护/)
    assert.throws(() => api.executeCommandSync(quick.iconSet, { ...target, ranges: scope, iconSet: '3Stars' }), /保护/)
    assert.throws(() => api.executeCommandSync(quick.dataBar, { ...target, ranges: [range(0, 0)], expectedRevision: 100 }), /已更改/)
    assert.deepEqual(worksheet.getConditionalFormats(), before)
    assert.equal(undoCount(studio), 0)
  }, workbook)
})

test('scoped edits and list clear submit one complete reversible snapshot while retaining hidden protected rules', () => {
  const first = { ...rule('first', [range(0, 0, 3, 0), range(6, 2, 8, 2)], 4), metadata: { source: ['first'] } }
  const hidden = { ...rule('hidden', [range(10, 13)], 9), metadata: { source: ['hidden'] } }
  const last = rule('last', [range(1, 1, 2, 1)], 13)
  const original = [first, hidden, last], selection = range(0, 0, 3, 1)
  for (const action of ['edit', 'clear']) {
    const source = sample()
    source.sheets[0].conditionalFormats = cloneModel(original)
    source.sheets[0].rangeProtectionRules = protection(hidden.ranges)
    const visible = conditionalFormatRuleListForScope(original, 'selection', selection)
    assert.deepEqual(visible.map(rule => rule.id), ['first', 'last'])
    const next = action === 'clear' ? removeConditionalFormatRuleDraft(original, visible.map(rule => rule.id))
      : original.map(rule => rule.id === first.id ? mergeConditionalFormatRuleDraft(rule, { ...rule, formula1: '8' }) : rule)
    const view = toExcelWorkbookView(source)
    const local = applySheetConditionalFormatsCommand(view.sheets[0], next, 'local-user')
    assert.deepEqual(local.ranges, action === 'clear' ? [...first.ranges, ...last.ranges] : first.ranges)
    assert.deepEqual(local.rules.find(rule => rule.id === hidden.id), hidden)
    withStudio((studio, api) => {
      const workbook = api.getActiveWorkbook(), worksheet = workbook.getActiveSheet()
      const cells = cloneModel(studio.editor.workbook.model.sheets[0].cells), events = [], executed = []
      studio.events.addEvent('SheetValueChanged', event => events.push(event))
      studio.events.addEvent('CommandExecuted', event => executed.push(event.id))
      assert.equal(studio.executeCommandSync(command, input(studio, next)), true)
      assert.deepEqual(worksheet.getConditionalFormats(), local.rules)
      assert.equal(undoCount(studio), 1)
      assert.deepEqual(events, [])
      assert.deepEqual(executed, [command])
      assert.deepEqual(studio.editor.workbook.model.sheets[0].cells, cells)
      workbook.undo()
      assert.deepEqual(worksheet.getConditionalFormats(), original)
      workbook.redo()
      assert.deepEqual(worksheet.getConditionalFormats(), local.rules)
      const reopened = api.createWorkbook({ ...workbook.save(), id: 'cf-scoped-' + action + '-reopened' })
      assert.deepEqual(reopened.getActiveSheet().getConditionalFormats(), local.rules)
      assert.deepEqual(studio.getWorkbookModel(reopened.getId()).sheets[0].cells, cells)
    }, source)
  }
})

test('scoped list clear checks every range of a matched rule before removing any rule', () => {
  const first = rule('multiple', [range(0, 0, 3, 0), range(14, 9)], 4)
  const hidden = rule('hidden', [range(10, 13)], 9)
  const source = sample()
  source.sheets[0].conditionalFormats = [first, hidden]
  source.sheets[0].rangeProtectionRules = protection([first.ranges[1]])
  const visible = conditionalFormatRuleListForScope([first, hidden], 'selection', range(0, 0))
  const cleared = removeConditionalFormatRuleDraft([first, hidden], visible.map(rule => rule.id))
  assert.deepEqual(cleared, [hidden])
  const view = toExcelWorkbookView(source), beforeView = cloneModel(view)
  assert.throws(() => applySheetConditionalFormatsCommand(view.sheets[0], cleared, 'local-user'), /保护/)
  assert.deepEqual(view, beforeView)
  withStudio(studio => {
    const before = cloneModel(studio.editor.workbook.model), revision = studio.getWorkbookRevision(source.id)
    assert.throws(() => studio.executeCommandSync(command, input(studio, cleared)), /保护/)
    assert.deepEqual(studio.editor.workbook.model, before)
    assert.equal(studio.getWorkbookRevision(source.id), revision)
    assert.equal(undoCount(studio), 0)
  }, source)
})

test('scoped moves use global priority slots and reject crossed hidden protection including tied priorities', () => {
  for (const priorities of [[4, 9, 13, 21], [7, 7, 7, 21]]) {
    const first = rule('first', [range(0, 0, 3, 0)], priorities[0])
    const hidden = { ...rule('hidden', [range(10, 13)], priorities[1]), metadata: { source: ['hidden'] } }
    const last = rule('last', [range(1, 1, 2, 1)], priorities[2])
    const untouched = rule('untouched', [range(20, 2)], priorities[3])
    const original = [first, hidden, last, untouched]
    const visible = conditionalFormatRuleListForScope(original, 'selection', range(0, 0, 3, 1))
    const moved = moveConditionalFormatRuleDraft(original, visible[0].id, 1, visible[1].id)
    assert.deepEqual(moved.map(rule => [rule.id, rule.priority]), [
      ['hidden', priorities[0]], ['last', priorities[1]], ['first', priorities[2]], ['untouched', priorities[3]],
    ])
    for (const locked of [false, true]) {
      const source = sample()
      source.sheets[0].conditionalFormats = cloneModel(original)
      source.sheets[0].rangeProtectionRules = locked ? protection(hidden.ranges) : []
      withStudio((studio, api) => {
        const before = cloneModel(studio.editor.workbook.model), revision = studio.getWorkbookRevision(source.id)
        if (locked) {
          assert.throws(() => studio.executeCommandSync(command, input(studio, moved)), /保护/)
          assert.deepEqual(studio.editor.workbook.model, before)
          assert.equal(studio.getWorkbookRevision(source.id), revision)
          assert.equal(undoCount(studio), 0)
          return
        }
        assert.equal(studio.executeCommandSync(command, input(studio, moved)), true)
        assert.deepEqual(api.getActiveWorkbook().getActiveSheet().getConditionalFormats(), moved)
        assert.equal(undoCount(studio), 1)
        api.getActiveWorkbook().undo()
        assert.deepEqual(api.getActiveWorkbook().getActiveSheet().getConditionalFormats(), original)
        api.getActiveWorkbook().redo()
        assert.deepEqual(api.getActiveWorkbook().getActiveSheet().getConditionalFormats(), moved)
      }, source)
    }
  }
})

test('registered edits share command events, one history transaction, cached values and internal snapshot reopen', () => withStudio((studio, api) => {
  const beforeCells = cloneModel(studio.editor.workbook.model.sheets[0].cells), valueEvents = [], commands = []
  studio.events.addEvent('SheetValueChanged', event => valueEvents.push(event))
  studio.events.addEvent('CommandExecuted', event => commands.push(event.id))
  const source = { ...rule(), formulaAnchor: { row: 1, col: 0 }, metadata: { vendor: ['keep'] } }
  assert.equal(studio.executeCommandSync(command, input(studio, [source])), true)
  assert.equal(undoCount(studio), 1); assert.deepEqual(commands, [command]); assert.deepEqual(valueEvents, [])
  assert.deepEqual(studio.editor.workbook.model.sheets[0].cells, beforeCells)
  source.metadata.vendor.push('caller changed'); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats[0].metadata.vendor, ['keep'])
  api.getActiveWorkbook().undo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats ?? [], [])
  api.getActiveWorkbook().redo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats[0].formulaAnchor, { row: 1, col: 0 })
  const saved = api.getActiveWorkbook().save(), reopened = api.createWorkbook({ ...saved, id: 'cf-edit-reopened' })
  assert.deepEqual(studio.getWorkbookModel(reopened.getId()).sheets[0].conditionalFormats[0].metadata.vendor, ['keep'])
  assert.deepEqual(studio.getWorkbookModel(reopened.getId()).sheets[0].cells, beforeCells)
}))

test('registered no-op retains revision, history and model while publishing the accepted command', () => {
  const workbook = sample(); workbook.sheets[0].conditionalFormats = [rule()]
  withStudio(studio => {
    const before = cloneModel(studio.editor.workbook.model), revision = studio.getWorkbookRevision(workbook.id), executed = []
    studio.events.addEvent('CommandExecuted', event => executed.push(event.id))
    assert.equal(studio.executeCommandSync(command, input(studio, [cloneModel(rule())], { expectedRevision: revision })), false)
    assert.deepEqual(studio.editor.workbook.model, before); assert.equal(studio.getWorkbookRevision(workbook.id), revision)
    assert.equal(undoCount(studio), 0); assert.deepEqual(executed, [command])
  }, workbook)
})

test('registered deletion, priority move and range changes are one reversible rule-list transaction', () => {
  const workbook = sample(), first = rule(), second = { ...rule('second', [range(0, 0, 3, 0)], 2), style: { fontColor: '#993366' } }
  workbook.sheets[0].conditionalFormats = [first, second]
  withStudio((studio, api) => {
    const next = [{ ...second, priority: 1, ranges: [range(5, 2, 9, 2)] }]
    assert.equal(studio.executeCommandSync(command, input(studio, next)), true)
    assert.equal(undoCount(studio), 1); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, next)
    api.getActiveWorkbook().undo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, [first, second])
    api.getActiveWorkbook().redo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, next)
  }, workbook)
})

test('registered range removal checks old protected scopes and permits untouched protected rules', () => {
  const workbook = sample(), locked = rule('locked', [range(9, 9)], 2)
  workbook.sheets[0].conditionalFormats = [rule(), locked]; workbook.sheets[0].rangeProtectionRules = protection(locked.ranges)
  withStudio(studio => {
    const before = cloneModel(studio.editor.workbook.model)
    assert.throws(() => studio.executeCommandSync(command, input(studio, [])), /保护/)
    assert.deepEqual(studio.editor.workbook.model, before); assert.equal(undoCount(studio), 0)
    assert.equal(studio.executeCommandSync(command, input(studio, [locked])), true)
    assert.equal(undoCount(studio), 1); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, [locked])
  }, workbook)
})

test('panel-list actions preserve unrelated protected priorities while crossing a protected rule rejects atomically', () => {
  const actions = [
    { next: rules => prependConditionalFormatRuleDraft(rules, rule('new', [range(6, 6)], 99)), ranges: [range(6, 6)] },
    { next: rules => removeConditionalFormatRuleDraft(rules, 'cf-first'), ranges: [range(0, 0, 3, 0)] },
    { next: rules => moveConditionalFormatRuleDraft(rules, 'cf-first', 1), ranges: [range(0, 0, 3, 0), range(3, 3, 4, 3)] },
  ]
  for (const [firstPriority, secondPriority, lockedPriority] of [[1, 2, 3], [7, 7, 13]]) {
    const first = rule('cf-first', [range(0, 0, 3, 0)], firstPriority)
    const second = rule('second', [range(3, 3, 4, 3)], secondPriority)
    const locked = { ...rule('locked', [range(10, 13)], lockedPriority), formulaAnchor: { row: 1, col: 0 }, metadata: { keep: ['locked'] } }
    const original = [first, second, locked]
    const workbook = sample(); workbook.sheets[0].conditionalFormats = original; workbook.sheets[0].rangeProtectionRules = protection(locked.ranges)
    for (const action of actions) {
      const next = action.next(original), view = toExcelWorkbookView(workbook)
      const local = applySheetConditionalFormatsCommand(view.sheets[0], next, 'local-user')
      assert.equal(local.changed, true); assert.deepEqual(local.ranges, action.ranges)
      assert.deepEqual(view.sheets[0].conditionalFormats.find(rule => rule.id === locked.id), locked)
      withStudio((studio, api) => {
        const beforeCells = cloneModel(studio.editor.workbook.model.sheets[0].cells), valueEvents = []
        studio.events.addEvent('SheetValueChanged', event => valueEvents.push(event))
        assert.equal(studio.executeCommandSync(command, input(studio, next)), true)
        assert.equal(undoCount(studio), 1); assert.deepEqual(valueEvents, [])
        assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, view.sheets[0].conditionalFormats)
        assert.deepEqual(studio.editor.workbook.model.sheets[0].cells, beforeCells)
        api.getActiveWorkbook().undo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, original)
        api.getActiveWorkbook().redo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, view.sheets[0].conditionalFormats)
      }, cloneModel(workbook))
    }
    const crossed = moveConditionalFormatRuleDraft(original, 'second', 1), view = toExcelWorkbookView(workbook), beforeView = cloneModel(view)
    assert.throws(() => applySheetConditionalFormatsCommand(view.sheets[0], crossed, 'local-user'), /保护/)
    assert.deepEqual(view, beforeView)
    withStudio(studio => {
      const before = cloneModel(studio.editor.workbook.model), revision = studio.getWorkbookRevision(workbook.id)
      assert.throws(() => studio.executeCommandSync(command, input(studio, crossed)), /保护/)
      assert.deepEqual(studio.editor.workbook.model, before); assert.equal(undoCount(studio), 0)
      assert.equal(studio.getWorkbookRevision(workbook.id), revision)
    }, cloneModel(workbook))
  }
})

test('registered optimistic revisions reject stale drafts and changes made in BeforeCommandExecute', () => withStudio(studio => {
  const revision = studio.getWorkbookRevision(studio.activeWorkbookId)
  studio.commitWorkbook(studio.activeWorkbookId, 'Other edit', model => { model.name = 'changed' }, { recalculate: false })
  const before = cloneModel(studio.editor.workbook.model)
  assert.throws(() => studio.executeCommandSync(command, input(studio, [rule()], { expectedRevision: revision })), /已更改/)
  assert.deepEqual(studio.editor.workbook.model, before); assert.equal(undoCount(studio), 1)
  const currentRevision = studio.getWorkbookRevision(studio.activeWorkbookId)
  const listener = studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === command) studio.commitWorkbook(studio.activeWorkbookId, 'Listener edit', model => { model.name = 'listener changed' }, { recalculate: false })
  })
  assert.throws(() => studio.executeCommandSync(command, input(studio, [rule()], { expectedRevision: currentRevision })), /已更改/)
  listener.dispose(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats ?? [], [])
  assert.equal(undoCount(studio), 2)
}))

test('BeforeCommandExecute cancellation preserves model and history and allows an explicit retry', () => withStudio(studio => {
  const before = cloneModel(studio.editor.workbook.model), executed = []
  studio.events.addEvent('CommandExecuted', event => executed.push(event.id))
  const listener = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === command) event.cancel = true })
  assert.throws(() => studio.executeCommandSync(command, input(studio, [rule()])), /canceled/)
  assert.deepEqual(studio.editor.workbook.model, before); assert.deepEqual(executed, []); assert.equal(undoCount(studio), 0)
  listener.dispose(); assert.equal(studio.executeCommandSync(command, input(studio, [rule()])), true)
  assert.equal(undoCount(studio), 1)
}))

test('explicit inactive workbook and subUnitId edits isolate their history and do not change the current workbook', () => withStudio(studio => {
  const other = sample(); other.id = 'cf-edit-other'; other.activeSheetId = other.sheets[0].id = 'cf-other-sheet'
  studio.createWorkbook(other, false)
  const original = cloneModel(studio.editor.workbook.model)
  assert.equal(studio.executeCommandSync(command, { unitId: other.id, subUnitId: other.sheets[0].id, rules: [rule()] }), true)
  assert.deepEqual(studio.editor.workbook.model, original); assert.equal(undoCount(studio), 0); assert.equal(undoCount(studio, other.id), 1)
  assert.deepEqual(studio.getWorkbookModel(other.id).sheets[0].conditionalFormats, [rule()])
}))

test('standalone views and registered edits use the same rule planning and scope permissions', () => withStudio(studio => {
  const view = toExcelWorkbookView(studio.editor.workbook.model), draft = { ...rule(), formulaAnchor: { row: 1, col: 0 }, metadata: { source: ['keep'] } }
  const local = applySheetConditionalFormatsCommand(view.sheets[0], [draft], 'local-user')
  assert.equal(studio.executeCommandSync(command, input(studio, [draft])), local.changed)
  const expected = studio.editor.workbook.model.sheets[0].conditionalFormats
  assert.deepEqual(view.sheets[0].conditionalFormats, expected)
  const reopened = applyExcelWorkbookView(sample(), view)
  assert.deepEqual(reopened.sheets[0].conditionalFormats, expected)
}))

test('registered full-list and quick commands validate history, user and host guards before no-op or change', () => {
  const source = sample(); source.sheets[0].conditionalFormats = [rule()]
  for (const id of [command, quick.dataBar, quick.clearWorksheet]) withStudio((studio, api) => {
    const unitId = source.id, sheetId = source.sheets[0].id, before = cloneModel(studio.getWorkbookModel(unitId))
    const context = { unitId, sheetId, host: 'workbench', expectedRevision: studio.getWorkbookRevision(unitId),
      expectedHistoryGeneration: studio.editor.history.captureSavepoint(unitId).generation, expectedUserId: studio.currentUserId }
    const params = { ...context, rules: [rule()], ranges: [range(0, 0)] }
    for (const invalid of [{ expectedHistoryGeneration: 0 }, { expectedUserId: '' }, { host: 'unknown' }]) {
      assert.throws(() => api.executeCommandSync(id, { ...params, ...invalid }), TypeError)
    }
    for (const stale of [{ expectedHistoryGeneration: context.expectedHistoryGeneration + 1 }, { expectedUserId: 'other-user' }]) {
      assert.throws(() => api.executeCommandSync(id, { ...params, ...stale }), /已更改/)
    }
    assert.deepEqual(studio.getWorkbookModel(unitId), before); assert.equal(undoCount(studio), 0)
    studio.editor.history.clear(unitId)
    assert.throws(() => api.executeCommandSync(id, params), /已更改/)
    const current = { ...params, expectedHistoryGeneration: studio.editor.history.captureSavepoint(unitId).generation }
    assert.equal(api.executeCommandSync(id, current), id !== command)
    assert.equal(undoCount(studio), id === command ? 0 : 1)
  }, cloneModel(source))
})

test('standalone rule replacement requires viewing every old and new range before touching the full list', () => {
  const original = rule('multiple', [range(0, 0, 3, 0), range(14, 9)], 4)
  const moved = { ...original, ranges: [range(6, 2), range(18, 4)] }
  for (const denied of [...original.ranges, ...moved.ranges]) {
    const sheet = sample().sheets[0]; sheet.conditionalFormats = [cloneModel(original)]
    sheet.rangeProtectionRules = protection([denied]).map(rule => ({ ...rule, viewState: 'noOneElseCanView' }))
    const before = cloneModel(sheet)
    assert.throws(() => applySheetConditionalFormatsCommand(sheet, [moved], 'local-user'), /查看/)
    assert.deepEqual(sheet, before)
  }
})

test('registered rule replacement views and edits all old and new scopes and excludes an untouched private rule', () => {
  const source = sample(), original = rule('multiple', [range(0, 0, 3, 0), range(14, 9)], 4)
  const hidden = rule('private', [range(10, 13)], 9), moved = { ...original, ranges: [range(6, 2), range(18, 4)] }
  source.sheets[0].conditionalFormats = [original, hidden]
  source.sheets[0].rangeProtectionRules = protection(hidden.ranges).map(rule => ({ ...rule, viewState: 'noOneElseCanView' }))
  withStudio((studio, api) => {
    const viewed = [], edited = [], view = studio.assertCanViewRange.bind(studio), edit = studio.assertCanEditRange.bind(studio)
    studio.assertCanViewRange = (id, sheetId, scope) => { viewed.push(cloneModel(scope)); view(id, sheetId, scope) }
    studio.assertCanEditRange = (id, sheetId, scope) => { edited.push(cloneModel(scope)); edit(id, sheetId, scope) }
    const beforeCells = cloneModel(studio.getWorkbookModel(source.id).sheets[0].cells)
    assert.equal(api.executeCommandSync(command, input(studio, [moved, hidden])), true)
    const expected = [...original.ranges, ...moved.ranges].map(scope => JSON.stringify(scope)).sort()
    for (const scopes of [viewed, edited]) {
      assert.deepEqual([...new Set(scopes.map(scope => JSON.stringify(scope)))].sort(), expected)
      for (const scope of expected) assert.ok(scopes.filter(value => JSON.stringify(value) === scope).length >= 3)
    }
    assert.deepEqual(studio.getWorkbookModel(source.id).sheets[0].conditionalFormats, [moved, hidden])
    assert.deepEqual(studio.getWorkbookModel(source.id).sheets[0].cells, beforeCells); assert.equal(undoCount(studio), 1)
  }, source)
})

test('revoking view or edit after planning, in the mutator or before view preparation rejects the whole transaction', () => {
  for (const pass of [1, 2, 3]) for (const viewState of ['othersCanView', 'noOneElseCanView']) {
    const source = sample(), original = rule('multiple', [range(0, 0, 3, 0), range(14, 9)], 4)
    source.sheets[0].conditionalFormats = [original]
    withStudio((studio, api) => {
      let calls = 0, revoked
      const check = studio.assertCanViewRange.bind(studio), revision = studio.getWorkbookRevision(source.id), executed = []
      studio.events.addEvent('CommandExecuted', event => executed.push(event.id))
      studio.assertCanViewRange = (id, sheetId, scope) => {
        if (scope.startRow === 14 && ++calls === pass) {
          studio.getWorkbookModel(id).sheets[0].rangeProtectionRules = protection([scope]).map(rule => ({ ...rule, viewState }))
          revoked = cloneModel(studio.getWorkbookModel(id))
        }
        check(id, sheetId, scope)
      }
      assert.throws(() => api.executeCommandSync(command, input(studio, [{ ...original, formula1: '8' }])), /保护/)
      assert.equal(calls, pass); assert.deepEqual(studio.getWorkbookModel(source.id), revoked)
      assert.equal(studio.getWorkbookRevision(source.id), revision); assert.equal(undoCount(studio), 0); assert.deepEqual(executed, [])
    }, source)
  }
})

test('candidate, mutator and prepared view recheck model identity, sheet identity, raw rules, user and history', () => {
  const changes = [
    studio => { studio.editor.workbook.replace(cloneModel(studio.getWorkbookModel(studio.activeWorkbookId))) },
    studio => { const model = studio.getWorkbookModel(studio.activeWorkbookId); model.sheets[0] = cloneModel(model.sheets[0]) },
    studio => { studio.getWorkbookModel(studio.activeWorkbookId).sheets[0].conditionalFormats[0].formula1 = '6' },
    studio => { Object.defineProperty(studio, 'currentUserId', { configurable: true, value: 'other-user' }) },
    studio => { studio.editor.history.clear(studio.activeWorkbookId) },
  ]
  for (const stage of ['candidate', 'mutator', 'prepareView']) for (const change of changes) {
    const source = sample(); source.sheets[0].conditionalFormats = [rule()]
    withStudio((studio, api) => {
      let changed = false, external, externalHistory, externalRevision
      const injectChange = () => {
        if (changed) return
        changed = true; change(studio); external = cloneModel(studio.getWorkbookModel(source.id))
        externalHistory = undoCount(studio); externalRevision = studio.getWorkbookRevision(source.id)
      }
      if (stage === 'candidate') {
        const check = studio.assertCanEditWorksheet.bind(studio)
        studio.assertCanEditWorksheet = (id, sheetId) => { injectChange(); check(id, sheetId) }
      } else {
        const commit = studio.commitWorkbook.bind(studio)
        studio.commitWorkbook = (id, label, mutate, options) => commit(id, label,
          draft => { if (stage === 'mutator') injectChange(); mutate(draft) },
          { ...options, prepareView: candidate => { if (stage === 'prepareView') injectChange(); options.prepareView(candidate) } })
      }
      const executed = []; studio.events.addEvent('CommandExecuted', event => executed.push(event.id))
      assert.throws(() => api.executeCommandSync(command, input(studio, [{ ...rule(), formula1: '8' }])), /已更改/)
      assert.equal(changed, true); assert.deepEqual(studio.getWorkbookModel(source.id), external)
      assert.equal(undoCount(studio), externalHistory); assert.equal(studio.getWorkbookRevision(source.id), externalRevision)
      assert.deepEqual(executed, [])
    }, source)
  }
})

test('workbench host guards reject inactive targets while explicit API targets retain their independent histories', () => withStudio((studio, api) => {
  const source = sample(); source.id = 'cf-inactive-host'; source.activeSheetId = source.sheets[0].id = 'inactive-sheet'
  studio.createWorkbook(source, false)
  const before = cloneModel(studio.getWorkbookModel(source.id)), params = { unitId: source.id, sheetId: source.sheets[0].id, rules: [rule()], host: 'workbench' }
  assert.throws(() => api.executeCommandSync(command, params), /切换/)
  assert.throws(() => api.executeCommandSync(quick.dataBar, { ...params, ranges: [range(0, 0)] }), /切换/)
  assert.deepEqual(studio.getWorkbookModel(source.id), before); assert.equal(undoCount(studio, source.id), 0)
  assert.equal(api.executeCommandSync(command, { ...params, host: undefined }), true)
  assert.equal(undoCount(studio), 0); assert.equal(undoCount(studio, source.id), 1)
}))
