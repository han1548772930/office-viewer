import assert from 'node:assert/strict'
import { test } from 'vitest'
import { canResumeSheetFillGesture, sheetFillGestureSelectionMatches, sheetFillGestureStillCurrent, snapshotSheetFillGesture } from '../src/features/sheet-workbench/sheet-fill-gesture.ts'

const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
const editing = () => ({ sheetId: 'first-sheet', workbookRevision: 7, selectedKey: 'A1',
  selectedRanges: [{ ...range }], editingKey: 'A1', editSessionId: 'edit-one' })
const ready = () => ({ ...editing(), workbookRevision: 8, editingKey: null, editSessionId: undefined })
const committed = { sheetId: 'first-sheet', workbookRevision: 8 }

test('fill gesture keeps an immutable selection across an in-place caller mutation', () => {
  const origin = editing(), captured = snapshotSheetFillGesture(origin)
  origin.selectedRanges[0].endRow = 3
  origin.selectedRanges.push({ ...range, startColumn: 4, endColumn: 4 })
  assert.deepEqual(captured.selectedRanges, [range])
  assert.equal(sheetFillGestureSelectionMatches(captured, origin), false)
})

test('fill resumes only at the acknowledged post-commit revision and unchanged selection', () => {
  assert.equal(canResumeSheetFillGesture(editing(), committed, ready()), true)
  assert.equal(canResumeSheetFillGesture(editing(), committed, { ...ready(), workbookRevision: 9 }), false)
  assert.equal(canResumeSheetFillGesture(editing(), { ...committed, workbookRevision: 7 }, ready()), false)
  assert.equal(canResumeSheetFillGesture(editing(), { ...committed, sheetId: 'second-sheet' }, ready()), false)
})

test('an unchanged-value commit and a standalone host can acknowledge their unchanged revision', () => {
  assert.equal(canResumeSheetFillGesture(editing(), { ...committed, workbookRevision: 7 }, { ...ready(), workbookRevision: 7 }), true)
  assert.equal(canResumeSheetFillGesture({ ...editing(), workbookRevision: undefined },
    { sheetId: 'first-sheet' }, { ...ready(), workbookRevision: undefined }), true)
})

test('composition or failed confirmation leaving any input session active cannot resume a fill', () => {
  assert.equal(canResumeSheetFillGesture(editing(), committed, { ...ready(), editingKey: 'A1' }), false)
  assert.equal(canResumeSheetFillGesture(editing(), committed, { ...ready(), editSessionId: 'edit-one' }), false)
  assert.equal(canResumeSheetFillGesture(editing(), committed, { ...ready(), editSessionId: 'replacement-edit' }), false)
})

test('same-sheet observer navigation, range edits, and sheet switches invalidate the original press', () => {
  for (const changed of [
    { ...ready(), selectedKey: 'E5' },
    { ...ready(), selectedRanges: [{ ...range, endRow: 1 }] },
    { ...ready(), selectedRanges: [{ ...range, rangeType: 1 }] },
    { ...ready(), selectedRanges: [{ ...range }, { ...range, startColumn: 4, endColumn: 4 }] },
    { ...ready(), sheetId: 'second-sheet' },
  ]) assert.equal(canResumeSheetFillGesture(editing(), committed, changed), false)
})

test('an active drag requires its original content revision, logical selection, and closed input session', () => {
  const origin = snapshotSheetFillGesture(ready())
  assert.equal(sheetFillGestureStillCurrent(origin, ready()), true)
  assert.equal(sheetFillGestureStillCurrent(origin, { ...ready(), workbookRevision: 9 }), false)
  assert.equal(sheetFillGestureStillCurrent(origin, { ...ready(), selectedKey: 'B1' }), false)
  assert.equal(sheetFillGestureStillCurrent(origin, { ...ready(), editingKey: 'A1', editSessionId: 'new-edit' }), false)
})
