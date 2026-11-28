import assert from 'node:assert/strict'
import { test } from 'vitest'
import { OfficeStudio } from '../src/api/office-studio.ts'
import { FDocument } from '../src/api/facade.ts'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph } from '../src/features/document/model/document-model.ts'
import { documentTextFields, setDocumentTextFields } from '../src/features/document/model/document-text-fields.ts'
import { createDocumentSnapshotHistory, prepareDocumentHistoryCacheRebase, prepareDocumentSnapshotHistoryAmend } from '../src/features/document/services/document-history-cache.ts'

// Drafts for the implementation batch; no validation has been run.
const timestamp = new Date(2030, 5, 15, 13, 24, 0).getTime()
const point = offset => ({ blockId: 'body', offset })
const selection = offset => ({ anchor: point(offset), focus: point(offset) })

function fixture() {
  const model = createBlankDocument()
  model.id = 'field-history'
  const block = createParagraph('oldTail'); block.id = 'body'
  setDocumentTextFields(block, [{ id: 'date', field: 'DATE', startOffset: 0, endOffset: 3,
    instruction: 'DATE \\@ "yyyy-MM-dd"', cachedResult: 'old' }])
  model.blocks = [block]
  const studio = new OfficeStudio({ document: model })
  return { studio, document: new FDocument(studio, model.id), id: model.id }
}

test('on-open updates live, undo and redo caches without changing history or its savepoint', () => {
  const { studio, document, id } = fixture()
  try {
    document.executeCommand({ id: 'doc.command.insert-text', text: '!' }, selection(0))
    document.setName('Renamed')
    assert.equal(document.undo(), true)
    document.save()
    const stack = studio.editor.history.units[id]
    const before = { undo: stack.undo.length, redo: stack.redo.length, revision: stack.revision,
      savedRevision: stack.savedRevision, generation: stack.generation, dirty: studio.editor.history.isDirty(id) }
    const caret = selection(5)
    document.setInputSelection(caret)
    document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
    const result = document.refreshFields({ locale: 'en-US', timestamp, onOpen: true, selection: caret })
    assert.equal(result.changed, true)
    assert.equal(result.selectionCurrent, true)
    assert.equal(result.selection.focus.offset, 12)
    assert.equal(document.getInputMarks().bold, true)
    assert.equal(document.getText(), '!2030-06-15Tail')
    assert.deepEqual({ undo: stack.undo.length, redo: stack.redo.length, revision: stack.revision,
      savedRevision: stack.savedRevision, generation: stack.generation, dirty: studio.editor.history.isDirty(id) }, before)
    assert.equal(document.getModel().modified, false)
    assert.equal(document.undo(), true)
    assert.equal(document.getText(), '2030-06-15Tail')
    assert.equal(documentTextFields(document.getModel().blocks[0])[0].endOffset, 10)
    assert.equal(document.redo(), true)
    assert.equal(document.getText(), '!2030-06-15Tail')
    assert.equal(document.redo(), true)
    assert.equal(document.getName(), 'Renamed')
    assert.equal(document.getText(), '!2030-06-15Tail')
  } finally { studio.dispose() }
})

test('already-current live caches still rebase historical field results on open', () => {
  const { studio, document } = fixture()
  try {
    assert.equal(document.refreshFields({ locale: 'en-US', timestamp }).changed, true)
    const before = studio.editor.history.units[document.id].undo.length
    assert.equal(document.refreshFields({ locale: 'en-US', timestamp, onOpen: true }).changed, false)
    assert.equal(studio.editor.history.units[document.id].undo.length, before)
    assert.equal(document.undo(), true)
    assert.equal(document.getText(), '2030-06-15Tail')
  } finally { studio.dispose() }
})

test('a failed cache projection cannot partially rewrite retained history snapshots', () => {
  const before = { ...createBlankDocument(), id: 'cache-before', name: 'Before' }
  const after = { ...before, name: 'After' }
  let published
  const callbacks = createDocumentSnapshotHistory(before, after, model => { published = model })
  let projections = 0
  assert.throws(() => prepareDocumentHistoryCacheRebase([callbacks], snapshot => {
    snapshot.name = 'Projected'
    if (++projections === 2) throw new Error('invalid cached field')
    return snapshot
  }), /invalid cached field/)
  callbacks.undo(); assert.equal(published.name, 'Before')
  callbacks.redo(); assert.equal(published.name, 'After')
})

test('derived cache publication retains redo and prevents nested content-history edits', () => {
  const { studio, document, id } = fixture()
  try {
    document.setName('First'); document.setName('Second'); document.undo()
    const stack = studio.editor.history.units[id], savepoint = studio.editor.history.captureSavepoint(id)
    studio.editor.history.publishDerived(id, () => {
      assert.throws(() => document.setName('Nested'), /历史事务正在提交/)
    })
    assert.equal(stack.redo.length, 1)
    assert.deepEqual(studio.editor.history.captureSavepoint(id), savepoint)
    assert.equal(document.getName(), 'First')
  } finally { studio.dispose() }
})

test('a TOC history amendment retains the original undo pair and rebases the final redo cache', () => {
  const before = { ...createBlankDocument(), name: 'Before' }, inserted = { ...before, name: 'Inserted' }, settled = { ...before, name: 'Settled' }
  let published
  const entry = createDocumentSnapshotHistory(before, inserted, value => { published = value })
  const amendment = prepareDocumentSnapshotHistoryAmend(entry, settled, value => { published = value })
  entry.redo(); assert.equal(published.name, 'Inserted')
  amendment.redo(); assert.equal(published.name, 'Settled')
  const merged = { ...entry, redo: amendment.redo }
  prepareDocumentHistoryCacheRebase([merged], value => ({ ...value, name: value.name + ':cached' }))()
  merged.undo(); assert.equal(published.name, 'Before:cached')
  merged.redo(); assert.equal(published.name, 'Settled:cached')
})

test('an amendment rejected before publication or throwing during publication cannot alter retained snapshot caches', () => {
  const before = { ...createBlankDocument(), name: 'Before' }, after = { ...before, name: 'Inserted' }
  let published
  const entry = createDocumentSnapshotHistory(before, after, value => { published = value })
  prepareDocumentSnapshotHistoryAmend(entry, { ...after, name: 'Never published' }, value => { published = value })
  entry.redo(); assert.equal(published.name, 'Inserted')
  const throwing = prepareDocumentSnapshotHistoryAmend(entry, { ...after, name: 'Failed' }, () => { throw new Error('publish refused') })
  assert.throws(() => throwing.redo(), /publish refused/)
  prepareDocumentHistoryCacheRebase([entry], value => ({ ...value, name: value.name + ':cached' }))()
  entry.undo(); assert.equal(published.name, 'Before:cached')
  entry.redo(); assert.equal(published.name, 'Inserted:cached')
})
