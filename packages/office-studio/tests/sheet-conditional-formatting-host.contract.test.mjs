import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, ref, shallowReactive } from 'vue'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { applySheetConditionalFormatsCommand } from '../src/features/conditional-formatting/conditional-format-command.ts'
import { moveConditionalFormatRuleDraft } from '../src/features/conditional-formatting/conditional-format-rule-editor.ts'
import { useSheetConditionalFormattingWorkbench } from '../src/components/spreadsheet/use-sheet-conditional-formatting-workbench.ts'

// Implementation-phase drafts. These tests and the SSR harness remain unexecuted.
let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = (startRow, startCol, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })
const rule = (id, ranges, priority) => ({ id, ranges, priority, type: 'cellIs', operator: 'greaterThan', formula1: '5',
  style: { fill: '#336699' }, stopIfTrue: false, metadata: { retained: [id] } })
function host(standalone = false, configure = () => {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'conditional-host'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'
  sheet.rowCount = 50; sheet.columnCount = 20
  sheet.cells['0:0'] = createCell(0, 0, 7)
  sheet.cells['10:13'] = { ...createCell(10, 13, 99), formula: '=1+1', dataType: 'formula', cachedDataType: 'number', display: '99' }
  sheet.conditionalFormats = [rule('first', [area(0, 0, 3, 0), area(14, 9)], 4), rule('hidden', [area(10, 13)], 9), rule('last', [area(3, 3, 4, 3)], 13)]
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other); configure(sheet)
  const studio = standalone ? undefined : new OfficeStudio({ workbook: model, currentUserId: 'local-user' })
  const initial = toExcelWorkbookView(studio?.getWorkbookModel(model.id) ?? model), scope = effectScope(), open = ref(false)
  const state = shallowReactive({ studio, model: initial, sheet: initial.sheets[0], sheetId: 'data', active: true, blocked: false,
    userId: 'local-user', inputSessionId: undefined, ranges: [area(0, 0)], defaultRange: 'A1' })
  const refreshed = [], focused = [], echoes = []
  const options = { open, state: () => state, refresh: (api, id, current) => {
    if (!current()) return
    refreshed.push(id); state.model = toExcelWorkbookView(api.getWorkbookModel(id)); state.sheet = state.model.sheets[0]
  }, applyStandalone: (rules, current, expectEcho) => {
    assert.equal(current(), true)
    const next = cloneModel(state.model), changed = applySheetConditionalFormatsCommand(next.sheets[0], rules, state.userId).changed
    if (changed) {
      const snapshot = cloneModel(next); echoes.push(snapshot); expectEcho(snapshot)
      state.model = next; state.sheet = next.sheets[0]; refreshed.push(next.id)
    }
    return changed
  }, canRestoreFocus: () => !state.blocked, focus: () => focused.push(state.model.id) }
  const controller = scope.run(() => useSheetConditionalFormattingWorkbench(options))
  const currentModel = () => studio?.getWorkbookModel(model.id) ?? state.model
  return { studio, state, scope, open, options, controller, refreshed, focused, echoes, model: currentModel,
    book: studio ? FOffice.newAPI(studio).getActiveWorkbook() : undefined,
    history: () => studio?.editor.history.units[model.id]?.undo.length ?? 0,
    dispose: () => { scope.stop(); studio?.dispose() } }
}
function edited(owner) { return owner.rules.map(rule => rule.id === 'first' ? { ...cloneModel(rule), formula1: '8' } : cloneModel(rule)) }
const command = 'sheet.command.set-conditional-formats'

test('opening copies the authoritative full list even when the view projection is older', async () => {
  const h = host()
  try {
    h.state.sheet.conditionalFormats = []
    const before = cloneModel(h.model())
    assert.equal(h.controller.openPanel(), true)
    const owner = h.controller.context.value
    assert.equal(owner.rules.length, 3); owner.rules[0].metadata.retained.push('draft')
    assert.deepEqual(h.model(), before); assert.equal(h.controller.openPanel('create'), false)
    owner.setOpen(false); owner.closed(); await nextTick()
    assert.equal(h.history(), 0); assert.deepEqual(h.refreshed, []); assert.deepEqual(h.focused, ['conditional-host'])
  } finally { h.dispose() }
})

test('one accepted edit refreshes once and returns focus only after its close callback', async () => {
  const h = host()
  try {
    const cells = cloneModel(h.model().sheets[0].cells)
    h.controller.openPanel(); const owner = h.controller.context.value
    assert.deepEqual(owner.submit(edited(owner)), { accepted: true, changed: true })
    assert.equal(h.history(), 1); assert.deepEqual(h.refreshed, ['conditional-host']); assert.deepEqual(h.focused, [])
    const rules = cloneModel(h.model().sheets[0].conditionalFormats)
    owner.setOpen(false); owner.closed(); await nextTick(); assert.deepEqual(h.focused, ['conditional-host'])
    owner.closed(); await nextTick(); assert.equal(h.focused.length, 1)
    h.book.undo(); assert.equal(h.model().sheets[0].conditionalFormats[0].formula1, '5')
    h.book.redo(); assert.deepEqual(h.model().sheets[0].conditionalFormats, rules); assert.deepEqual(h.model().sheets[0].cells, cells)
  } finally { h.dispose() }
})

test('an accepted no-op reports changed=false and a competing input prevents delayed focus', async () => {
  const h = host()
  try {
    const before = cloneModel(h.model())
    h.controller.openPanel(); const owner = h.controller.context.value
    assert.deepEqual(owner.submit(owner.rules), { accepted: true, changed: false })
    assert.deepEqual(h.model(), before); assert.equal(h.history(), 0); assert.deepEqual(h.refreshed, [])
    owner.setOpen(false); h.state.blocked = true; owner.closed(); await nextTick(); assert.deepEqual(h.focused, [])
    h.state.blocked = false; h.controller.openPanel(); const second = h.controller.context.value
    second.setOpen(false); second.closed(); h.controller.openPanel(); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('workbook, worksheet, user, input session and activation changes end the old owner', async () => {
  for (const change of [state => { state.model = { ...state.model, id: 'replacement' } }, state => { state.sheetId = 'other' },
    state => { state.userId = 'other-user' }, state => { state.inputSessionId = 'new-cell-input' }, state => { state.active = false; state.active = true }]) {
    const h = host()
    try {
      h.controller.openPanel(); const owner = h.controller.context.value; change(h.state)
      assert.equal(h.open.value, false); assert.equal(h.controller.context.value, undefined)
      assert.equal(owner.submit(edited(owner)).accepted, false)
      owner.closed(); await nextTick(); assert.deepEqual(h.focused, []); assert.equal(h.history(), 0)
    } finally { h.dispose() }
  }
})

test('content, cleared history and same-ID model replacements cannot receive an old full-list draft', async () => {
  for (const change of [h => h.book.getActiveSheet().getRange('B2').setValue('external'),
    h => h.studio.editor.history.clear('conditional-host'), h => h.studio.editor.workbook.replace(cloneModel(h.model()))]) {
    const h = host()
    try {
      h.controller.openPanel(); const owner = h.controller.context.value; change(h)
      const before = cloneModel(h.model()), count = h.history()
      assert.equal(owner.submit(edited(owner)).accepted, false)
      assert.equal(h.open.value, false); assert.deepEqual(h.model(), before); assert.equal(h.history(), count)
      owner.closed(); await nextTick(); assert.deepEqual(h.focused, [])
    } finally { h.dispose() }
  }
})

test('command cancellation preserves the draft, permits retry and disposal removes late actions', async () => {
  const h = host()
  try {
    h.controller.openPanel(); const owner = h.controller.context.value
    const cancel = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === command) event.cancel = true })
    assert.match(owner.submit(edited(owner)).error, /已取消/)
    assert.equal(h.open.value, true); assert.equal(h.controller.context.value, owner); assert.equal(h.history(), 0)
    cancel.dispose(); assert.equal(owner.submit(edited(owner)).changed, true)
    h.scope.stop(); assert.equal(owner.submit(edited(owner)).accepted, false)
    owner.closed(); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('a host leaving and returning during view refresh still invalidates the owner', async () => {
  const h = host()
  try {
    h.options.refresh = () => { h.state.active = false; h.state.active = true }
    h.controller.openPanel(); const owner = h.controller.context.value
    assert.deepEqual(owner.submit(edited(owner)), { accepted: true, changed: true })
    assert.equal(h.history(), 1); assert.equal(h.open.value, false); assert.equal(h.controller.context.value, undefined)
    owner.setOpen(false); owner.closed(); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('a settled listener edit cannot be mistaken for the drawer commit or steal focus', async () => {
  const h = host()
  try {
    const subscription = h.studio.events.addEvent('CommandExecuted', event => {
      if (event.id === command) h.book.getActiveSheet().getRange('B2').setValue('listener edit')
    })
    h.controller.openPanel(); const owner = h.controller.context.value
    assert.equal(owner.submit(edited(owner)).changed, true); subscription.dispose()
    assert.equal(h.history(), 2); assert.equal(h.open.value, false); assert.deepEqual(h.refreshed, [])
    assert.equal(h.model().sheets[0].cells['1:1'].value, 'listener edit')
    owner.closed(); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('rule capability covers all ranges and planner priority changes while untouched hidden rules survive', () => {
  const h = host(false, sheet => { sheet.rangeProtectionRules = [{ id: 'locked', permissionId: 'lock-p', ownerUserId: 'owner',
    allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', ranges: [area(10, 13)], description: '' }] })
  try {
    h.controller.openPanel(); const owner = h.controller.context.value
    assert.equal(owner.canEditRule(owner.rules[0]), true); assert.equal(owner.canEditRule(owner.rules[1]), false)
    assert.equal(owner.canApplyRules(edited(owner)), true)
    assert.equal(owner.canApplyRules(owner.rules.filter(rule => rule.id !== 'hidden')), false)
    assert.equal(owner.canApplyRules(moveConditionalFormatRuleDraft(owner.rules, 'first', 1, 'last')), false)
    h.state.blocked = true; assert.equal(owner.submit(edited(owner)).accepted, false); assert.equal(h.history(), 0)
  } finally { h.dispose() }
})

test('range preview is independent from selection/history and denied if any preview scope is unreadable', () => {
  const h = host(false, sheet => { sheet.rangeProtectionRules = [{ id: 'private', permissionId: 'private-p', ownerUserId: 'owner',
    allowedUserIds: [], editState: 'onlyMe', viewState: 'noOneElseCanView', ranges: [area(14, 9)], description: '' }] })
  try {
    const before = cloneModel(h.model()), selections = cloneModel(h.state.ranges)
    h.controller.openPanel(); const owner = h.controller.context.value
    owner.preview([area(0, 0)]); assert.deepEqual(h.controller.previewRanges.value, [area(0, 0)])
    owner.preview(owner.rules[0].ranges); assert.deepEqual(h.controller.previewRanges.value, [])
    owner.preview([area(3, 3)]); h.state.blocked = true; assert.deepEqual(h.controller.previewRanges.value, [])
    h.state.blocked = false; owner.setOpen(false); assert.deepEqual(h.controller.previewRanges.value, [])
    assert.deepEqual(h.model(), before); assert.deepEqual(h.state.ranges, selections); assert.equal(h.history(), 0)
  } finally { h.dispose() }
})

test('standalone edits share permissions and no-op semantics while external replacement rejects old callbacks', () => {
  const h = host(true)
  try {
    h.controller.openPanel(); const owner = h.controller.context.value
    assert.deepEqual(owner.submit(owner.rules), { accepted: true, changed: false })
    assert.deepEqual(owner.submit(edited(owner)), { accepted: true, changed: true }); assert.equal(h.refreshed.length, 1)
    assert.deepEqual(owner.submit(h.model().sheets[0].conditionalFormats), { accepted: true, changed: false })
    h.state.model = cloneModel(h.state.model); const before = cloneModel(h.state.model)
    assert.equal(owner.submit(edited(owner)).accepted, false); assert.deepEqual(h.state.model, before)
  } finally { h.dispose() }
})

test('callbacks from an older closing drawer cannot close or preview a newly opened drawer', async () => {
  const h = host()
  try {
    h.controller.openPanel(); const first = h.controller.context.value; first.setOpen(false)
    h.controller.openPanel(); const second = h.controller.context.value
    first.setOpen(false); first.closed(); first.preview([area(4, 4)]); await nextTick()
    assert.equal(h.open.value, true); assert.equal(h.controller.context.value, second)
    assert.deepEqual(h.controller.previewRanges.value, []); assert.deepEqual(h.focused, [])
    assert.equal(first.submit(edited(first)).accepted, false); assert.equal(h.history(), 0)
  } finally { h.dispose() }
})

test('standalone focus survives exactly its published view echo and equal external replacements still invalidate it', async () => {
  for (const external of [false, true]) {
    const h = host(true)
    try {
      h.controller.openPanel(); const owner = h.controller.context.value
      assert.deepEqual(owner.submit(edited(owner)), { accepted: true, changed: true })
      owner.setOpen(false)
      const incoming = external ? cloneModel(h.echoes[0]) : h.echoes[0]
      h.controller.receiveView(incoming, () => { h.state.model = cloneModel(incoming); h.state.sheet = h.state.model.sheets[0] })
      if (external) {
        assert.equal(h.controller.context.value, undefined); assert.equal(owner.submit(edited(owner)).accepted, false)
      } else assert.equal(h.controller.context.value, owner)
      owner.closed(); await nextTick(); assert.deepEqual(h.focused, external ? [] : ['conditional-host'])
      assert.equal(h.refreshed.length, 1); assert.equal(h.echoes.length, 1)
    } finally { h.dispose() }
  }
})
