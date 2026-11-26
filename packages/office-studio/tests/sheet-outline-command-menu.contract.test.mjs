import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetOutlineState, getSheetOutlineUpdate } from '../src/features/sheet-outline/sheet-outline.ts'
import { getSheetOutlinePromptChoices, outlineSelectionLabel, resolveSheetOutlineCommand as resolve,
  sheetOutlineCommandInput, SHEET_OUTLINE_CORE_COMMAND_IDS, SHEET_OUTLINE_UI_COMMAND_IDS } from '../src/features/sheet-outline/sheet-outline-command.ts'
import { createSheetOutlineContextMenuItem, createSheetOutlineRibbonItem, SHEET_OUTLINE_SHORTCUTS } from '../src/features/sheet-outline/sheet-outline-menu.ts'
import { sheetOutlineCommandForShortcut } from '../src/components/spreadsheet/use-sheet-outline-workbench.ts'
import { ribbonInvocation } from '../src/platform/ribbon/ribbon-schema.ts'

// Command and state drafts, not an executed feature acceptance report.
const selection = { startRow: 2, endRow: 4, startCol: 3, endCol: 5 }
const group = (id, axis, start, end, collapsed = false) => ({ id, axis, start, end, collapsed })
const context = (overrides = {}) => ({ unitId: 'book', sheetId: 'sheet', userId: 'visitor',
  sheet: { rowCount: 20, columnCount: 10, outline: createSheetOutlineState() }, selections: [selection], ...overrides })
const guard = ranges => ({ id: 'guard', ownerUserId: 'owner', permissionId: 'guard', editState: 'onlyOwnerCanEdit', allowedUserIds: [], ranges })

test('the four core commands resolve inclusive ranges and stable public identities', () => {
  const current = context(), added = resolve(SHEET_OUTLINE_CORE_COMMAND_IDS[0], { axis: 'row', start: 2, end: 4 }, current)
  current.sheet.outline = getSheetOutlineUpdate(current.sheet, { ...added.input.action, groupId: 'public' }, current.userId)
  assert.deepEqual(added.input.action, { axis: 'row', start: 2, end: 4, kind: 'group' })
  assert.deepEqual(resolve(SHEET_OUTLINE_CORE_COMMAND_IDS[1], { outlineId: 'public' }, current).input.action,
    { kind: 'remove-group', axis: 'row', groupId: 'public' })
  assert.deepEqual(resolve(SHEET_OUTLINE_CORE_COMMAND_IDS[2], { outlineId: 'public', collapsed: true }, current).input.action,
    { kind: 'set-collapsed', axis: 'row', groupId: 'public', collapsed: true })
  const clear = resolve(SHEET_OUTLINE_CORE_COMMAND_IDS[3], { axis: 'row', start: 2, end: 4 }, current)
  assert.deepEqual(getSheetOutlineUpdate(current.sheet, clear.input.action, current.userId).groups, [])
})

test('every selected row or column UI command maps to the real range operation', () => {
  const current = context()
  for (const id of SHEET_OUTLINE_UI_COMMAND_IDS.filter(id => id.includes('selected-rows') || id.includes('selected-columns'))) {
    const result = resolve(id, {}, current), axis = id.includes('rows') ? 'row' : 'column'
    assert.deepEqual(result.input.action, { axis, start: axis === 'row' ? 2 : 3, end: axis === 'row' ? 4 : 5,
      kind: id.includes('clear-selected') ? 'remove' : 'group' })
  }
})

test('ordinary single selections prompt for an axis while whole-row and whole-column shortcuts execute directly', () => {
  const current = context(), prompt = resolve('sheet.command.outline-selected-dimension-shortcut', {}, current)
  assert.equal(prompt.kind, 'prompt'); assert.deepEqual(prompt.choices.map(choice => choice.axis), ['row', 'column'])
  for (const [rangeType, axis] of [[1, 'row'], [2, 'column']]) {
    const direct = resolve('sheet.command.outline-selected-dimension-shortcut', {}, context({ selections: [{ ...selection, rangeType }] }))
    assert.equal(direct.kind, 'change'); assert.equal(direct.input.action.axis, axis)
  }
})

test('dimension UI selectors honor explicit axis values and avoid guessing an axis', () => {
  assert.equal(resolve('sheet.command.outline-selected-dimension', {}, context()), undefined)
  assert.equal(resolve('sheet.command.clear-selected-dimension-outlines', {}, context()), undefined)
  assert.equal(resolve('sheet.command.outline-selected-dimension', { value: 'column' }, context()).input.action.axis, 'column')
  assert.equal(resolve('sheet.command.clear-selected-dimension-outlines', { value: 'row' }, context()).input.action.kind, 'remove')
})

test('multiple selections suppress shortcuts and all corresponding menu entries', () => {
  const current = context({ selections: [selection, { ...selection, startRow: 7, endRow: 8 }] })
  assert.equal(resolve('sheet.command.outline-selected-dimension-shortcut', {}, current), undefined)
  assert.deepEqual(getSheetOutlinePromptChoices(current, 'group'), [])
  assert.equal(createSheetOutlineRibbonItem(current).hidden, true)
  assert.equal(createSheetOutlineContextMenuItem(current, 'row').hidden, true)
})

test('axis prompt capability distinguishes an editable row range from protected full columns', () => {
  const current = context()
  current.sheet.rangeProtectionRules = [guard([{ startRow: 12, endRow: 12, startCol: 4, endCol: 4 }])]
  const choices = getSheetOutlinePromptChoices(current, 'group')
  assert.equal(choices.find(choice => choice.axis === 'row').disabled, false)
  assert.equal(choices.find(choice => choice.axis === 'column').disabled, true)
  assert.match(choices.find(choice => choice.axis === 'column').reason, /受保护/)
})

test('crossing group previews and clear no-ops are reflected in selector children', () => {
  const current = context({ selections: [{ ...selection, startRow: 6, endRow: 12 }] })
  current.sheet.outline.groups = [group('outer', 'row', 2, 8)]
  const item = createSheetOutlineContextMenuItem(current, 'row')
  assert.equal(item.children[0].disabled, true); assert.match(item.children[0].disabledReason, /交叉/)
  assert.equal(item.children[1].disabled, true)
  current.selections = [{ ...selection, startRow: 2, endRow: 8 }]
  assert.equal(createSheetOutlineContextMenuItem(current, 'row').children[1].disabled, false)
})

test('worksheet protection disables menus and the same command model rejects execution', () => {
  const current = context()
  current.sheet.worksheetProtection = { id: 'sheet-guard', unitType: 'worksheet', ownerUserId: 'owner', permissionId: 'guard',
    editState: 'onlyMe', viewState: 'othersCanView', allowedUserIds: [] }
  assert.equal(createSheetOutlineRibbonItem(current).disabled, true)
  assert.equal(createSheetOutlineContextMenuItem(current, 'column').disabled, true)
  const result = resolve('sheet.command.outline-selected-rows', {}, current)
  assert.throws(() => getSheetOutlineUpdate(current.sheet, result.input.action, current.userId), /受保护/)
})

test('selector children and keyboard descriptors preserve fixed command routing', () => {
  const current = context(); current.sheet.outline.groups = [group('existing', 'row', 2, 4)]
  const item = createSheetOutlineRibbonItem(current)
  assert.deepEqual(ribbonInvocation(item, 'group'), { commandId: 'sheet.command.outline-selected-dimension-shortcut', params: {} })
  assert.deepEqual(ribbonInvocation(item, 'ungroup'), { commandId: 'sheet.command.clear-selected-dimension-outlines-shortcut', params: {} })
  assert.deepEqual(SHEET_OUTLINE_SHORTCUTS.map(item => [item.key, item.altKey, item.shiftKey, item.macKey, item.macPrimary]),
    [['ArrowRight', true, true, 'k', true], ['ArrowLeft', true, true, 'j', true]])
  const event = { key: 'ArrowRight', altKey: true, shiftKey: true, ctrlKey: false, metaKey: false, isComposing: false, keyCode: 0 }
  for (const shortcut of SHEET_OUTLINE_SHORTCUTS) {
    assert.equal(sheetOutlineCommandForShortcut({ ...event, key: shortcut.key }), shortcut.commandId)
    for (const modifier of ['ctrlKey', 'metaKey']) {
      const macEvent = { ...event, key: shortcut.macKey.toUpperCase(), altKey: false, [modifier]: true }
      assert.equal(sheetOutlineCommandForShortcut(macEvent, true), shortcut.commandId)
      assert.equal(sheetOutlineCommandForShortcut({ ...macEvent, shiftKey: false }, true), undefined)
      assert.equal(sheetOutlineCommandForShortcut(macEvent, false), undefined)
      assert.equal(sheetOutlineCommandForShortcut({ ...macEvent, isComposing: true }, true), undefined)
    }
  }
})

test('command resolution forwards the captured owner and rejects ambiguous or invalid ownership fields', () => {
  const params = { axis: 'row', start: 2, end: 4, unitId: 'book', subUnitId: 'sheet', expectedRevision: 4,
    expectedHistoryGeneration: 7, expectedUserId: 'visitor', host: 'workbench' }
  const input = resolve('sheet.command.add-dimension-outline', params, context()).input
  assert.equal(input.expectedRevision, 4); assert.equal(input.expectedHistoryGeneration, 7); assert.equal(input.expectedUserId, 'visitor'); assert.equal(input.host, 'workbench')
  assert.throws(() => resolve('sheet.command.add-dimension-outline', { ...params, subUnitId: 'other' }, context()), /不匹配/)
  assert.throws(() => sheetOutlineCommandInput({ expectedHistoryGeneration: 0 }), /正整数/)
  assert.throws(() => sheetOutlineCommandInput({ subUnitId: 'first', sheetId: 'second' }), /不一致/)
  assert.throws(() => sheetOutlineCommandInput({ collapsed: 'false' }), /布尔/)
})

test('local summary and level commands use atomic model actions and preserve the untouched axis', () => {
  const current = context(); current.sheet.outline.groups = [group('row', 'row', 2, 4), group('column', 'column', 2, 4)]
  const summary = resolve('sheet.command.set-outline-summary-settings', { summaryAbove: true, summaryLeft: true }, current)
  assert.equal(getSheetOutlineUpdate(current.sheet, summary.input.action, current.userId).summaryLeft, true)
  const level = resolve('sheet.command.set-outline-level', { axis: 'row', level: 1 }, current)
  assert.deepEqual(getSheetOutlineUpdate(current.sheet, level.input.action, current.userId).groups.map(item => item.collapsed), [true, false])
})

test('selection labels use inclusive one-based row and spreadsheet column names', () => {
  assert.equal(outlineSelectionLabel({ axis: 'row', start: 0, end: 2 }), '组合第 1–3 行')
  assert.equal(outlineSelectionLabel({ axis: 'column', start: 25, end: 27 }, 'remove'), '取消分组 Z–AB 列')
})
