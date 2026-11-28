import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, ref, shallowReactive } from 'vue'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createSheetDataValidationRule } from '../src/features/sheet-validation/validation-engine.ts'
import { SHEET_DATA_VALIDATION_COMMANDS } from '../src/features/sheet-validation/validation-rule-command.ts'
import { useSheetDataValidationWorkbench } from '../src/components/spreadsheet/use-sheet-data-validation-workbench.ts'

// Implementation-phase owner contracts. No SSR server, test runner or browser has been started.
let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = (row = 0, column = 0, endRow = row, endCol = column) => ({ startRow: row, endRow, startCol: column, endCol })
function host(configure = () => {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'validation-panel-host'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'
  sheet.rowCount = 100; sheet.columnCount = 30; sheet.cells['0:0'] = createCell(0, 0, 60)
  sheet.dataValidations = [createSheetDataValidationRule({ id: 'date', type: 'date', metadata: { showTime: true, retained: ['owner'] } }, [area()]),
    createSheetDataValidationRule({ id: 'list', type: 'listMultiple', formula1: '"Red,Green"', formula2: '#ff0000,' }, [area(3, 0)])]
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other); configure(model, sheet)
  const studio = new OfficeStudio({ workbook: model, currentUserId: 'local-user' }), book = FOffice.newAPI(studio).getActiveWorkbook()
  const state = shallowReactive({ studio, model: toExcelWorkbookView(studio.getWorkbookModel(model.id)), sheetId: 'data', userId: 'local-user', active: true, blocked: false,
    inputSessionId: undefined, inputOwner: 'validation-panel-owner', ranges: [area()], defaultRange: 'A1' })
  const open = ref(false), scope = effectScope(), refreshed = [], focused = []
  const options = { open, state: () => state, refresh: (studio, unitId, current) => { if (current()) { refreshed.push(unitId); state.model = toExcelWorkbookView(studio.getWorkbookModel(unitId)) } },
    focus: () => focused.push(state.model.id), canRestoreFocus: () => !state.blocked }
  const controller = scope.run(() => useSheetDataValidationWorkbench(options)), current = () => studio.getWorkbookModel(model.id)
  return { studio, book, state, open, scope, options, controller, refreshed, focused, current, count: () => studio.editor.history.units[model.id]?.undo.length ?? 0,
    show: (type, id) => { assert.equal(controller.openPanel(type, id), true); return controller.context.value }, dispose: () => { scope.stop(); studio.dispose() } }
}

test('rule management reads authoritative metadata and captures one owner, range default, input and history session', async () => {
  const h = host()
  try {
    h.state.model.sheets[0].dataValidations[0].metadata.showTime = false
    const before = cloneModel(h.current()), owner = h.show(undefined, 'date')
    assert.equal(owner.rules[0].metadata.showTime, true); assert.equal(owner.defaultRange, 'A1'); assert.equal(owner.inputOwner, 'validation-panel-owner')
    assert.equal(owner.initialRuleId, 'date'); assert.equal(owner.canAdd, true); assert.equal(owner.canRemoveAll, true); assert.equal(owner.disabled, false)
    owner.setOpen(false); owner.closed(); await nextTick(); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0); assert.deepEqual(h.focused, ['validation-panel-host'])
  } finally { h.dispose() }
})

test('hidden ranges remove entire rule entries from the visible list and prevent accidental clear-all', () => {
  const h = host((_model, sheet) => {
    sheet.dataValidations[1].ranges.push(area(3, 12)); sheet.rangeProtectionRules = [{ id: 'private', ranges: [area(3, 12)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'othersCanEdit' }]
  })
  try {
    const before = cloneModel(h.current()), owner = h.show(); assert.deepEqual(owner.rules.map(rule => rule.id), ['date']); assert.equal(owner.canRemoveAll, false)
    assert.equal(owner.canEditRule(cloneModel(h.current().sheets[0].dataValidations[1])), false)
    assert.equal(owner.showRule('list'), false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    owner.setOpen(false); owner.closed(); assert.equal(h.controller.openPanel(undefined, 'list'), false)
  } finally { h.dispose() }
})

test('a visible but protected rule cannot be edited or removed, while a legal new target remains available', () => {
  const h = host((_model, sheet) => { sheet.rangeProtectionRules = [{ id: 'read-only', ranges: [area(3, 0)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'othersCanView', editState: 'onlyMe' }] })
  try {
    const before = cloneModel(h.current()), owner = h.show(), rule = owner.rules.find(rule => rule.id === 'list')
    assert.equal(owner.rules.length, 2); assert.equal(owner.canEditRule(rule), false); assert.equal(owner.canRemoveAll, false); assert.equal(owner.canAdd, true)
    assert.equal(owner.submit({ action: 'remove', ruleId: 'list' }).accepted, false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    assert.equal(owner.disabled, false); assert.equal(h.open.value, true)
  } finally { h.dispose() }
})

test('one owned rule edit updates visible rules and host view once and keeps the owner open across another edit', () => {
  const h = host()
  try {
    const owner = h.show(undefined, 'date'), rule = { ...cloneModel(owner.rules[0]), metadata: { showTime: false, retained: ['owner'] } }
    assert.deepEqual(owner.submit({ action: 'replace', rule }), { accepted: true, changed: true }); assert.equal(owner.isCurrent(), true); assert.equal(h.open.value, true)
    assert.equal(h.controller.context.value, owner); assert.equal(owner.rules.find(rule => rule.id === 'date').metadata.showTime, false); assert.equal(h.count(), 1); assert.deepEqual(h.refreshed, ['validation-panel-host'])
    assert.deepEqual(owner.submit({ action: 'replace', rule: cloneModel(owner.rules.find(rule => rule.id === 'date')) }), { accepted: true, changed: false }); assert.equal(h.count(), 1)
    assert.deepEqual(owner.submit({ action: 'remove', ruleId: 'list' }), { accepted: true, changed: true }); assert.equal(h.count(), 2); assert.equal(owner.rules.length, 1)
    assert.equal(owner.canRemoveAll, true); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('normal command cancellation keeps the owner and draft context live for retry without a content history record', () => {
  const h = host()
  try {
    const owner = h.show(), before = cloneModel(h.current()), subscription = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === SHEET_DATA_VALIDATION_COMMANDS.remove) event.cancel = true })
    const rejected = owner.submit({ action: 'remove', ruleId: 'list' }); assert.equal(rejected.accepted, false); assert.match(rejected.error, /取消.*草稿/)
    assert.equal(owner.isCurrent(), true); assert.equal(h.open.value, true); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    subscription.dispose(); assert.equal(owner.submit({ action: 'remove', ruleId: 'list' }).accepted, true); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test('Canvas range picking changes selection and add capability without replacing the current rule owner or default', () => {
  const h = host((_model, sheet) => { sheet.rangeProtectionRules = [{ id: 'denied', ranges: [area(7, 0)], ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }] })
  try {
    const owner = h.show(); h.state.ranges = [area(7, 0)]; h.studio.setSelections(h.state.model.id, 'data', [{ startRow: 7, endRow: 7, startColumn: 0, endColumn: 0 }])
    assert.equal(h.controller.context.value, owner); assert.equal(owner.isCurrent(), true); assert.equal(owner.canAdd, false); assert.equal(owner.defaultRange, 'A1')
    h.state.ranges = [area(2, 0)]; assert.equal(owner.canAdd, true); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test.each(['sheet', 'workbook', 'user', 'input', 'activation'])('%s leaving and returning permanently invalidates the old rule owner and every captured callback', async kind => {
  const h = host()
  try {
    const owner = h.show(), before = cloneModel(h.current())
    if (kind === 'sheet') { h.state.sheetId = 'other'; h.state.sheetId = 'data' }
    else if (kind === 'workbook') { h.state.model = { ...h.state.model, id: 'peer' }; h.state.model = toExcelWorkbookView(h.current()) }
    else if (kind === 'user') { h.state.userId = 'other'; h.state.userId = 'local-user' }
    else if (kind === 'input') { h.state.inputSessionId = 'new-session'; h.state.inputSessionId = undefined }
    else { h.state.active = false; h.state.active = true }
    assert.equal(h.open.value, false); assert.equal(h.controller.context.value, undefined); assert.equal(owner.submit({ action: 'remove', ruleId: 'list' }).accepted, false)
    const next = h.show(); owner.setOpen(false); owner.closed(); await nextTick(); assert.equal(h.controller.context.value, next); assert.equal(h.open.value, true)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test.each(['history', 'same-ID model', 'rule edit', 'unrelated content'])('%s changes close the full-list owner without claiming stale rule edits', async kind => {
  const h = host()
  try {
    const owner = h.show()
    if (kind === 'history') h.studio.editor.history.clear(h.state.model.id)
    else if (kind === 'same-ID model') h.studio.editor.workbook.replace(cloneModel(h.current()))
    else if (kind === 'rule edit') h.current().sheets[0].dataValidations[0].metadata.showTime = false
    else h.book.getActiveSheet().getRange('H10').setValue('External')
    await nextTick(); assert.equal(h.open.value, false); assert.equal(h.controller.context.value, undefined)
    assert.equal(owner.submit({ action: 'remove', ruleId: 'list' }).accepted, false); owner.closed(); await nextTick(); assert.deepEqual(h.focused, [])
    assert.equal(h.current().sheets[0].dataValidations.some(rule => rule.id === 'list'), true)
  } finally { h.dispose() }
})

test('observer edits after rule publication keep their own history and prevent stale host refresh or after-close focus', async () => {
  const h = host()
  try {
    const owner = h.show(), subscription = h.studio.events.addEvent('CommandExecuted', event => { if (event.id === SHEET_DATA_VALIDATION_COMMANDS.remove) h.book.getActiveSheet().getRange('H10').setValue('Observer') })
    assert.equal(owner.submit({ action: 'remove', ruleId: 'list' }).accepted, true); subscription.dispose()
    assert.equal(h.count(), 2); assert.equal(h.current().sheets[0].cells['9:7'].value, 'Observer'); assert.equal(h.controller.context.value, undefined)
    assert.deepEqual(h.refreshed, []); owner.closed(); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('leaving and returning during view publication ends the owner and cannot restore its focus', async () => {
  const h = host()
  try {
    const owner = h.show(); h.options.refresh = () => { h.state.active = false; h.state.active = true }
    assert.equal(owner.submit({ action: 'remove', ruleId: 'list' }).accepted, true); assert.equal(h.count(), 1); assert.equal(h.controller.context.value, undefined)
    owner.closed(); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('reference preview reads only permitted source cells and a blocked host cannot open or submit rules', () => {
  const h = host((_model, sheet) => { sheet.cells['0:3'] = createCell(0, 3, 'Secret'); sheet.rangeProtectionRules = [{ id: 'private-source', ranges: [area(0, 3)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'othersCanEdit' }] })
  try {
    h.state.blocked = true; assert.equal(h.controller.openPanel(), false); h.state.blocked = false
    const owner = h.show(), before = cloneModel(h.current()); assert.deepEqual(owner.listOptions('=$D$1'), [])
    h.state.blocked = true; assert.equal(owner.disabled, true); assert.equal(owner.canAdd, false); assert.equal(owner.canRemoveAll, false)
    assert.equal(owner.submit({ action: 'remove', ruleId: 'list' }).accepted, false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('composition keeps Ant inputs editable but prevents rule commands, navigation and close until the owned input finishes', async () => {
  const h = host()
  try {
    const owner = h.show(), before = cloneModel(h.current()); owner.setComposing(true)
    assert.equal(owner.composing, true); assert.equal(owner.disabled, false); assert.equal(owner.canAdd, false); assert.equal(owner.canRemoveAll, false)
    assert.equal(owner.canEditRule(owner.rules[0]), false); assert.equal(owner.showRule('date'), false)
    owner.showRules(); assert.equal(owner.navigationRevision, 0); owner.setOpen(false); assert.equal(h.open.value, true)
    assert.equal(owner.submit({ action: 'remove', ruleId: 'list' }).accepted, false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    owner.setComposing(false); assert.equal(owner.showRule('date'), true); assert.equal(owner.initialRuleId, 'date'); assert.equal(owner.navigationRevision, 1)
    owner.showRules(); assert.equal(owner.initialRuleId, undefined); assert.equal(owner.navigationRevision, 2)
    owner.setOpen(false); owner.closed(); const next = h.show(); next.setComposing(true); owner.setComposing(false); await nextTick()
    assert.equal(next.composing, true); assert.equal(h.open.value, true); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})
