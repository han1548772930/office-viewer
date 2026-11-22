import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { applySheetConditionalFormatsCommand, assertSheetConditionalFormatRuleSnapshot, planSheetConditionalFormatsCommand,
  SET_SHEET_CONDITIONAL_FORMATS_COMMAND as command, sheetConditionalFormatsCommandInput } from '../src/features/conditional-formatting/conditional-format-command.ts'
import { conditionalFormatFormulaAt, reanchorConditionalFormatRule } from '../src/features/conditional-formatting/conditional-format-formulas.ts'
import { moveConditionalFormatRuleDraft, prependConditionalFormatRuleDraft, removeConditionalFormatRuleDraft } from '../src/features/conditional-formatting/conditional-format-rule-editor.ts'
import { conditionalFormatStyleAt, evaluateConditionalFormatRule } from '../src/features/conditional-formatting/conditional-formatting.ts'
import { evaluateFormula } from '../src/features/formula/formula-engine.ts'

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
    { rules: [], expectedRevision: -1 }, { rules: [], expectedRevision: 1.5 }]) assert.throws(() => sheetConditionalFormatsCommandInput(input))
  const input = { unitId: 'workbook', subUnitId: 'sheet', rules: [rule()], expectedRevision: 0 }
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
