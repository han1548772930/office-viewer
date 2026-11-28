import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, ref, shallowReactive } from 'vue'
import { createBlankSheet, createBlankWorkbook } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { ComputeError } from '../src/compute/errors.ts'
import { createSheetDataValidationRule } from '../src/features/sheet-validation/validation-engine.ts'
import { useSheetDataValidationWorkbench } from '../src/components/spreadsheet/use-sheet-data-validation-workbench.ts'
import { useSheetQuickValidationWorkbench } from '../src/components/spreadsheet/use-sheet-quick-validation-workbench.ts'

// Unexecuted owner/command contracts use controllable compute transport. Algorithm and real Worker/browser drafts remain separate.
let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = { startRow: 0, endRow: 79, startCol: 0, endCol: 79 }
const computeProgress = { phase: 'computing', processed: 3200, total: 6400, unit: 'cells' }

function transport({ abortImmediately = true } = {}) {
  const jobs = [], waiters = new Map()
  const client = { forgetDocument() {}, setDocumentRevision() {}, submit(task, options) {
    const captured = structuredClone(task, { transfer: [task.ranges.buffer, task.merges.buffer, task.positions.buffer, task.actions.buffer] })
    let resolve, reject, settled = false
    const promise = new Promise((yes, no) => { resolve = yes; reject = no })
    const finish = (callback, value) => {
      if (settled) return
      settled = true; options.signal?.removeEventListener('abort', aborted); callback(value)
    }
    const aborted = () => { if (abortImmediately) finish(reject, new ComputeError('COMPUTE_CANCELLED', '计算已取消')) }
    const requestId = `owned-bulk-${jobs.length + 1}`
    const job = { captured, options, requestId, promise,
      progress(value = computeProgress) { try { options.onProgress?.(value) } catch { /* A late transport observer cannot settle content. */ } },
      complete() { finish(resolve, { type: 'result', requestId, documentId: options.documentId, modelRevision: options.modelRevision,
        output: { kind: 'checkboxCells', runs: new Uint32Array([0, 79, 0, 79, 1]), changedCells: 6400 },
        metrics: { elapsedMs: 1, inputBytes: 16, estimatedPeakBytes: 1, boundary: {}, transport: {} } }) },
      cancel: aborted,
    }
    options.signal?.addEventListener('abort', aborted, { once: true })
    if (options.signal?.aborted) aborted()
    jobs.push(job); waiters.get(jobs.length - 1)?.(job); waiters.delete(jobs.length - 1)
    return job
  } }
  return { client, jobs, waitFor(index = 0) { return jobs[index] ? Promise.resolve(jobs[index]) : new Promise(resolve => waiters.set(index, resolve)) } }
}

function host(transportOptions = {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'owned-bulk'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'
  sheet.rowCount = 80; sheet.columnCount = 80; sheet.cells = {}; sheet.dataValidations = []
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other)
  const compute = transport(transportOptions), studio = new OfficeStudio({ workbook: model, currentUserId: 'local-user', computeClient: compute.client })
  const book = FOffice.newAPI(studio).getActiveWorkbook(), scope = effectScope(), open = ref(false)
  const state = shallowReactive({ studio, model: toExcelWorkbookView(studio.getWorkbookModel(model.id)), sheetId: 'data', userId: 'local-user', active: true,
    blocked: false, inputOwner: 'owned-bulk-panel', inputSessionId: undefined, ranges: [area], defaultRange: 'A1:CB80' })
  const refreshed = [], focused = [], openedRules = [], notices = [], events = []
  const refresh = (source, unitId, current) => {
    if (current()) { refreshed.push(unitId); state.model = toExcelWorkbookView(source.getWorkbookModel(unitId)) }
  }
  const panel = scope.run(() => useSheetDataValidationWorkbench({ open, state: () => state, refresh,
    focus: () => focused.push(state.model.id), canRestoreFocus: () => !state.blocked }))
  const quick = scope.run(() => useSheetQuickValidationWorkbench({
    state: () => ({ studio: state.studio, unitId: state.model.id, sheetId: state.sheetId, userId: state.userId,
      active: state.active, blocked: state.blocked, ranges: state.ranges }),
    translator: () => (_key, fallback) => fallback, owner: () => open.value ? panel.context.value : undefined,
    refresh, openRule: ruleId => openedRules.push(ruleId), notify: message => notices.push(message),
  }))
  const subscription = studio.events.addEvent('SheetValueChanged', value => events.push(cloneModel(value)))
  const next = createSheetDataValidationRule({ id: 'bulk-rule', type: 'checkbox', formula1: 'Yes', formula2: 'No' }, [area])
  return { studio, book, compute, scope, state, open, panel, quick, next, refreshed, focused, openedRules, notices, events,
    current: () => studio.getWorkbookModel(model.id), count: () => studio.editor.history.units[model.id]?.undo.length ?? 0,
    show() { assert.equal(panel.openPanel('checkbox'), true); return panel.context.value },
    dispose() { scope.stop(); subscription.dispose(); studio.dispose() },
  }
}

test('an async panel owner disables authoring, rejects duplicate submission and commits one visible rule and value transaction', async () => {
  const h = host()
  try {
    const before = cloneModel(h.current()), owner = h.show(), task = owner.submitAsync({ action: 'set', rule: h.next })
    const job = await h.compute.waitFor()
    assert.equal(owner.pending, true); assert.equal(owner.disabled, true); assert.equal(owner.canAdd, false); assert.equal(owner.canRemoveAll, false)
    assert.equal((await owner.submitAsync({ action: 'set', rule: h.next })).accepted, false)
    owner.setOpen(false); assert.equal(h.open.value, true); assert.equal(h.compute.jobs.length, 1); assert.deepEqual(h.current(), before)
    job.progress(); assert.equal(owner.progress.phase, 'computing'); assert.equal(owner.progress.processed, 3200)
    job.complete(); assert.deepEqual(await task, { accepted: true, changed: true })
    assert.equal(owner.pending, false); assert.equal(owner.progress, undefined); assert.equal(owner.isCurrent(), true); assert.equal(owner.disabled, false)
    assert.deepEqual(owner.rules.map(rule => rule.id), ['bulk-rule']); assert.equal(h.count(), 1); assert.equal(h.events.length, 1)
    assert.deepEqual(h.refreshed, ['owned-bulk']); assert.ok(Object.values(h.current().sheets[0].cells).every(cell => cell.value === 'No'))
  } finally { h.dispose() }
})

test('panel cancellation clears progress, preserves the owner and allows an actual retry without an earlier content record', async () => {
  const h = host()
  try {
    const before = cloneModel(h.current()), owner = h.show(), first = owner.submitAsync({ action: 'set', rule: h.next }), job = await h.compute.waitFor()
    job.progress(); owner.cancel(); const rejected = await first
    assert.equal(rejected.accepted, false); assert.match(rejected.error, /取消.*草稿/); assert.equal(job.options.signal.aborted, true)
    assert.equal(owner.pending, false); assert.equal(owner.progress, undefined); assert.equal(owner.isCurrent(), true); assert.equal(h.open.value, true)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0); assert.deepEqual(h.events, [])
    const retry = owner.submitAsync({ action: 'set', rule: h.next }), second = await h.compute.waitFor(1)
    job.progress(); assert.equal(owner.progress.phase, 'classifying'); second.complete(); assert.equal((await retry).accepted, true)
    assert.equal(h.count(), 1); assert.equal(h.events.length, 1)
  } finally { h.dispose() }
})

test.each(['sheet', 'workbook', 'user', 'input', 'activation'])('%s changes abort an old panel task and late completion cannot refresh, close or focus its replacement', async kind => {
  const h = host({ abortImmediately: false })
  try {
    const before = cloneModel(h.current()), owner = h.show(), task = owner.submitAsync({ action: 'set', rule: h.next }), job = await h.compute.waitFor()
    if (kind === 'sheet') { h.state.sheetId = 'other'; h.state.sheetId = 'data' }
    else if (kind === 'workbook') { h.state.model = { ...h.state.model, id: 'peer' }; h.state.model = toExcelWorkbookView(h.current()) }
    else if (kind === 'user') { h.state.userId = 'peer'; h.state.userId = 'local-user' }
    else if (kind === 'input') { h.state.inputSessionId = 'replacement'; h.state.inputSessionId = undefined }
    else { h.state.active = false; h.state.active = true }
    assert.equal(job.options.signal.aborted, true); assert.equal(h.panel.context.value, undefined); assert.equal(h.open.value, false)
    const next = h.show(); job.progress(); job.complete(); assert.equal((await task).accepted, false)
    owner.setOpen(false); owner.closed(); await nextTick()
    assert.equal(h.panel.context.value, next); assert.equal(h.open.value, true); assert.equal(next.pending, false); assert.equal(next.progress, undefined)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0); assert.deepEqual(h.events, []); assert.deepEqual(h.refreshed, []); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('a closed-panel quick checkbox owns progress and duplicate suppression and refreshes one actual publication', async () => {
  const h = host()
  try {
    const before = cloneModel(h.current()), task = h.quick.apply('checkbox'), job = await h.compute.waitFor()
    assert.equal(h.quick.busy.value, true); await h.quick.apply('checkbox'); assert.equal(h.compute.jobs.length, 1); assert.deepEqual(h.current(), before)
    job.progress(); assert.equal(h.quick.progress.value.phase, 'computing'); assert.equal(h.quick.progress.value.processed, 3200)
    job.complete(); await task
    assert.equal(h.quick.busy.value, false); assert.equal(h.quick.progress.value, undefined); assert.equal(h.count(), 1); assert.equal(h.events.length, 1)
    assert.equal(Object.keys(h.current().sheets[0].cells).length, 6400); assert.ok(Object.values(h.current().sheets[0].cells).every(cell => cell.value === 0))
    assert.deepEqual(h.refreshed, ['owned-bulk']); assert.deepEqual(h.openedRules, []); assert.deepEqual(h.notices, [])
  } finally { h.dispose() }
})

test('quick cancellation remains available with an open pending panel and never removes its draft owner', async () => {
  const h = host()
  try {
    const before = cloneModel(h.current()), owner = h.show(), task = h.quick.apply('checkbox'), job = await h.compute.waitFor()
    assert.equal(h.quick.busy.value, true); assert.equal(owner.pending, true); job.progress(); assert.equal(owner.progress.phase, 'computing')
    h.quick.cancel(); await task
    assert.equal(job.options.signal.aborted, true); assert.equal(h.quick.busy.value, false); assert.equal(owner.pending, false); assert.equal(owner.isCurrent(), true)
    assert.equal(h.open.value, true); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0); assert.deepEqual(h.events, [])
    assert.equal(h.notices.length, 1); assert.match(h.notices[0], /取消.*草稿/)
  } finally { h.dispose() }
})

test.each(['activation', 'blocked', 'user', 'scope'])('quick %s invalidation suppresses late progress, notifications and view publication', async kind => {
  const h = host({ abortImmediately: false })
  try {
    const before = cloneModel(h.current()), task = h.quick.apply('checkbox'), job = await h.compute.waitFor()
    if (kind === 'activation') { h.state.active = false; h.state.active = true }
    else if (kind === 'blocked') { h.state.blocked = true; h.state.blocked = false }
    else if (kind === 'user') { h.state.userId = 'peer'; h.state.userId = 'local-user' }
    else h.scope.stop()
    assert.equal(job.options.signal.aborted, true); job.progress(); job.complete(); await task
    assert.equal(h.quick.busy.value, false); assert.equal(h.quick.progress.value, undefined); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    assert.deepEqual(h.events, []); assert.deepEqual(h.refreshed, []); assert.deepEqual(h.openedRules, []); assert.deepEqual(h.notices, [])
  } finally { h.dispose() }
})

test('scope disposal aborts a pending panel command and prevents content, focus and view callbacks', async () => {
  const h = host()
  try {
    const before = cloneModel(h.current()), owner = h.show(), task = owner.submitAsync({ action: 'set', rule: h.next }), job = await h.compute.waitFor()
    h.scope.stop(); assert.equal(job.options.signal.aborted, true); assert.equal((await task).accepted, false)
    job.progress(); job.complete(); owner.closed(); await nextTick()
    assert.equal(h.panel.context.value, undefined); assert.equal(h.open.value, false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    assert.deepEqual(h.events, []); assert.deepEqual(h.refreshed, []); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})
