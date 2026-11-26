import assert from 'node:assert/strict'
import { test } from 'vitest'
import { SheetDrawingMoveHistory, SHEET_DRAWING_MOVE_HISTORY_WINDOW } from '../src/features/drawing-groups/drawing-move-history.ts'

// Deferred receipt contracts; this class stores no content snapshots or independent undo stack.
const context = (extra = {}) => ({ unitId: 'book', sheetId: 'data', userId: 'local', selectionKey: '["chart:c","image:a"]',
  selectionGeneration: 2, contentRevision: 4, savepoint: { unitId: 'book', revision: 5, generation: 1 }, savedRevision: 1, now: 1000, ...extra })

test('the exact applied receipt can be amended through the inclusive fixed one-second window', () => {
  const history = new SheetDrawingMoveHistory(), first = context(); history.remember(first)
  assert.equal(SHEET_DRAWING_MOVE_HISTORY_WINDOW, 1000)
  assert.deepEqual(history.amendment(context({ now: 2000 }), () => true), first.savepoint)
})

test('each actual committed move renews the window from its newest receipt', () => {
  const history = new SheetDrawingMoveHistory(); history.remember(context())
  const next = context({ now: 1900, contentRevision: 5, savepoint: { unitId: 'book', revision: 6, generation: 1 } })
  history.remember(next)
  assert.deepEqual(history.amendment({ ...next, now: 2900 }, () => true), next.savepoint)
})

test('expired windows and clocks that move backwards start a fresh history entry', () => {
  const history = new SheetDrawingMoveHistory(); history.remember(context())
  assert.equal(history.amendment(context({ now: 2001 }), () => true), undefined)
  assert.equal(history.amendment(context({ now: 999 }), () => true), undefined)
})

test('unit, sheet, user, selection owner, content, saved and reopened-history changes separate move bursts', () => {
  for (const extra of [{ unitId: 'other' }, { sheetId: 'other' }, { userId: 'peer' }, { selectionKey: '["image:a"]' },
    { selectionGeneration: 3 }, { contentRevision: 5 }, { savedRevision: 2 },
    { savepoint: { unitId: 'book', revision: 6, generation: 1 } }, { savepoint: { unitId: 'book', revision: 5, generation: 2 } }]) {
    const history = new SheetDrawingMoveHistory(); history.remember(context())
    assert.equal(history.amendment(context(extra), () => true), undefined, JSON.stringify(extra))
  }
})

test('an acknowledged current savepoint is preserved as a separately undoable boundary', () => {
  const history = new SheetDrawingMoveHistory(), saved = context({ savedRevision: 5 }); history.remember(saved)
  assert.equal(history.amendment(saved, () => true), undefined)
})

test('a displaced top entry or a redo branch cannot be amended even with an otherwise matching receipt', () => {
  const history = new SheetDrawingMoveHistory(); history.remember(context())
  assert.equal(history.amendment(context(), () => false), undefined)
})

test('receipts are copied and lifecycle clearing discards only the grouping session', () => {
  const history = new SheetDrawingMoveHistory(), original = context(); history.remember(original)
  original.savepoint.revision = 123
  const receipt = history.amendment(context(), () => true); assert.equal(receipt.revision, 5)
  receipt.revision = 234
  assert.equal(history.amendment(context(), () => true).revision, 5)
  history.clear(); assert.equal(history.amendment(context(), () => true), undefined)
})
