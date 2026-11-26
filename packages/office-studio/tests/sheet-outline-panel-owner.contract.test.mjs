import assert from 'node:assert/strict'
import { effectScope, markRaw, nextTick, reactive } from 'vue'
import { test } from 'vitest'
import { EventHub } from '../src/api/events.ts'
import { CHANGE_SHEET_OUTLINE_COMMAND, createSheetOutlineState, getSheetOutlineUpdate } from '../src/features/sheet-outline/sheet-outline.ts'
import { applySheetOutlinePanelChange, captureSheetOutlinePanelOwner, prepareSheetOutlinePanelChange,
  sheetOutlinePanelChangePublished, sheetOutlinePanelOwnerCurrent } from '../src/features/sheet-outline/sheet-outline-panel-owner.ts'
import { useSheetOutlineDialog } from '../src/features/sheet-outline/use-sheet-outline-dialog.ts'

// Deterministic host fixtures cover draft ownership; real Studio history has a separate integration draft.
function fixture() {
  const runtime = reactive({ userId: 'user', revision: 0, historyGeneration: 1, activeWorkbookId: 'book', activeKind: 'workbook',
    workbook: { id: 'book', activeSheetId: 'sheet', sheets: [{ id: 'sheet', rowCount: 20, columnCount: 10, outline: createSheetOutlineState() },
      { id: 'other-sheet', rowCount: 20, columnCount: 10 }] } })
  const events = new EventHub(), calls = [], target = { workbookId: 'book', sheetId: 'sheet' }, subscriptions = new Set()
  const subscribe = events.addEvent.bind(events)
  events.addEvent = (name, listener) => {
    const subscription = subscribe(name, listener), tracked = { dispose() { subscription.dispose(); subscriptions.delete(tracked) } }
    subscriptions.add(tracked); return tracked
  }
  let onExecute = () => {}, accepted = true
  const studio = markRaw({
    get currentUserId() { return runtime.userId }, get activeWorkbookId() { return runtime.activeWorkbookId }, events,
    editor: { state: runtime, get workbook() { return { model: runtime.workbook } }, history: { captureSavepoint() { return { generation: runtime.historyGeneration } } } },
    getWorkbookModel(id) { return runtime.workbook?.id === id ? runtime.workbook : undefined }, getWorkbookRevision() { return runtime.revision },
    executeCommandSync(id, input) {
      calls.push({ id, input })
      if (!accepted) return false
      assert.equal(id, CHANGE_SHEET_OUTLINE_COMMAND)
      assert.equal(input.expectedRevision, runtime.revision); assert.equal(input.expectedHistoryGeneration, runtime.historyGeneration)
      assert.equal(input.expectedUserId, runtime.userId)
      const sheet = runtime.workbook.sheets.find(sheet => sheet.id === input.sheetId), outline = getSheetOutlineUpdate(sheet, input.action, runtime.userId)
      if (!outline) return false
      sheet.outline = outline; runtime.revision += 1; events.emit('CommandExecuted', { id, params: input, unitId: input.unitId })
      onExecute(); return true
    },
  })
  return { runtime, studio, target, calls, subscriptions, owner: () => captureSheetOutlinePanelOwner(studio, target),
    onExecute(callback) { onExecute = callback }, accept(value) { accepted = value } }
}
const action = { kind: 'group', axis: 'row', start: 2, end: 4, groupId: 'stable-draft-id' }
function dialog(f, callbacks = {}) {
  const host = reactive({ studio: f.studio, ...f.target, open: true, draftKey: 'draft' }), scope = effectScope(), closed = [], changed = []
  const controls = scope.run(() => useSheetOutlineDialog({ get studio() { return host.studio }, get workbookId() { return host.workbookId },
    get sheetId() { return host.sheetId }, get open() { return host.open }, get draftKey() { return host.draftKey },
    close: (owner, committed) => { closed.push({ owner, committed }); host.open = false; callbacks.close?.() },
    changed: () => { changed.push(true); callbacks.changed?.(host) } }))
  return { host, scope, controls, closed, changed }
}

test('a captured owner is detached from caller target mutations and retains content/history/user identity', () => {
  const f = fixture(), owner = f.owner(); f.target.sheetId = 'changed-by-caller'
  assert.deepEqual(owner.target, { workbookId: 'book', sheetId: 'sheet' })
  assert.equal(owner.userId, 'user'); assert.equal(owner.historyGeneration, 1); assert.equal(owner.revision, 0)
  assert.equal(sheetOutlinePanelOwnerCurrent(owner), true)
})

test('preview and publication share group ID and forward every owner field', () => {
  const f = fixture(), prepared = prepareSheetOutlinePanelChange(f.owner(), action)
  assert.equal(prepared.after.groups[0].id, action.groupId); assert.equal(prepared.input.action.groupId, action.groupId)
  assert.equal(prepared.input.expectedRevision, 0); assert.equal(prepared.input.expectedHistoryGeneration, 1)
  assert.equal(prepared.input.expectedUserId, 'user'); assert.equal(prepared.input.host, 'workbench')
  assert.deepEqual(f.runtime.workbook.sheets[0].outline.groups, [])
  assert.equal(applySheetOutlinePanelChange(prepared), true); assert.equal(sheetOutlinePanelChangePublished(prepared), true)
})

test('stale revision, history reset and user change refuse a prepared command before dispatch', () => {
  for (const key of ['revision', 'historyGeneration', 'userId']) {
    const f = fixture(), prepared = prepareSheetOutlinePanelChange(f.owner(), action)
    f.runtime[key] = key === 'userId' ? 'other-user' : f.runtime[key] + 1
    assert.equal(applySheetOutlinePanelChange(prepared), false); assert.equal(f.calls.length, 0)
  }
})

test('active kind, workbook, worksheet and disposal each invalidate a panel owner', () => {
  for (const change of [f => { f.runtime.activeKind = 'document' }, f => { f.runtime.activeWorkbookId = 'other-book' },
    f => { f.runtime.workbook.activeSheetId = 'other-sheet' }, f => { f.runtime.workbook = undefined }]) {
    const f = fixture(), owner = f.owner(); change(f); assert.equal(sheetOutlinePanelOwnerCurrent(owner), false)
  }
})

test('own publication accepts exactly one revision and rejects reentrant extra changes', () => {
  const f = fixture(), prepared = prepareSheetOutlinePanelChange(f.owner(), action)
  f.onExecute(() => { f.runtime.revision += 1 })
  assert.equal(applySheetOutlinePanelChange(prepared), true)
  assert.equal(sheetOutlinePanelChangePublished(prepared), false)
  assert.equal(sheetOutlinePanelOwnerCurrent(prepared.owner, true), false)
})

test('cancelled host and no-op draft changes do not dispatch an edit command', () => {
  const f = fixture(), owner = f.owner(), prepared = prepareSheetOutlinePanelChange(owner, action)
  assert.equal(applySheetOutlinePanelChange(prepared, () => false), false)
  assert.equal(prepareSheetOutlinePanelChange(owner, { kind: 'summary-settings', summaryAbove: false, summaryLeft: false }), undefined)
  assert.equal(f.calls.length, 0)
  f.accept(false)
  assert.equal(applySheetOutlinePanelChange(prepared), false); assert.equal(sheetOutlinePanelChangePublished(prepared), false)
})

test('permission changes without a content revision are rechecked immediately before dispatch', () => {
  const f = fixture(), prepared = prepareSheetOutlinePanelChange(f.owner(), action)
  f.runtime.workbook.sheets[0].rangeProtectionRules = [{ id: 'guard', permissionId: 'guard', ownerUserId: 'other-user',
    editState: 'onlyOwnerCanEdit', allowedUserIds: [], ranges: [{ startRow: 3, endRow: 3, startCol: 0, endCol: 9 }] }]
  assert.throws(() => applySheetOutlinePanelChange(prepared), /受保护/); assert.equal(f.calls.length, 0)
})

test('ordinary dialog apply closes once with its own committed owner and one owned command', () => {
  const f = fixture(), d = dialog(f)
  try {
    assert.equal(d.controls.apply(action), true); assert.equal(d.closed.length, 1); assert.equal(d.closed[0].committed, true)
    assert.equal(d.changed.length, 1); assert.equal(f.calls.length, 1); assert.equal(d.host.open, false)
  } finally { d.scope.stop() }
})

test('closing and reopening during command execution leaves the replacement dialog open', () => {
  const f = fixture(), d = dialog(f)
  f.onExecute(() => { d.host.open = false; d.host.open = true })
  try {
    assert.equal(d.controls.apply(action), false); assert.equal(d.closed.length, 0); assert.equal(d.changed.length, 0)
    assert.equal(d.host.open, true); assert.equal(d.controls.owner.value.revision, 1); assert.equal(d.controls.busy.value, false)
  } finally { d.scope.stop() }
})

test('a changed observer can replace the draft without the old apply closing it', () => {
  const f = fixture(), d = dialog(f, { changed: host => { host.draftKey = 'replacement' } })
  try {
    assert.equal(d.controls.apply(action), true); assert.equal(d.closed.length, 0); assert.equal(d.host.open, true)
    assert.equal(d.controls.owner.value.revision, 1); assert.equal(d.controls.busy.value, false)
  } finally { d.scope.stop() }
})

test('external edits leave the old draft stale until explicit reread', async () => {
  const f = fixture(), d = dialog(f)
  try {
    const other = prepareSheetOutlinePanelChange(f.owner(), { ...action, axis: 'column', groupId: 'other' })
    applySheetOutlinePanelChange(other); await nextTick()
    assert.equal(d.controls.stale.value, true); assert.equal(d.controls.apply(action), false)
    d.controls.read(); assert.equal(d.controls.stale.value, false); assert.equal(d.controls.owner.value.revision, 1)
  } finally { d.scope.stop() }
})

test('scope disposal during publication suppresses completion and releases subscriptions', () => {
  const f = fixture(), d = dialog(f); f.onExecute(() => d.scope.stop())
  assert.equal(d.controls.apply(action), false); assert.equal(d.closed.length, 0); assert.equal(d.changed.length, 0)
  assert.equal(f.subscriptions.size, 0)
})
