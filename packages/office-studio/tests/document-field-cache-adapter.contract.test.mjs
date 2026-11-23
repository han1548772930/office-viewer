import assert from 'node:assert/strict'
import { test } from 'vitest'
import { computed } from 'vue'
import { OfficeStudio } from '../src/api/office-studio.ts'
import { FDocument } from '../src/api/facade.ts'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph } from '../src/features/document/model/document-model.ts'
import { documentTextFields, setDocumentTextFields } from '../src/features/document/model/document-text-fields.ts'
import { reserveDocumentFieldOpenRefresh, releaseDocumentFieldOpenRefresh } from '../src/features/document/services/document-field-open-refresh.ts'

// Contract drafts; execution belongs to the unified validation phase.
const timestamp = new Date(2030, 5, 15, 13, 24, 0).getTime()
const caret = offset => ({ anchor: { blockId: 'body', offset }, focus: { blockId: 'body', offset } })
function fixture() {
  const model = { ...createBlankDocument(), id: 'adapter-document', modified: false }
  const paragraph = { ...createParagraph('oldTail'), id: 'body' }
  setDocumentTextFields(paragraph, [{ id: 'date', field: 'DATE', startOffset: 0, endOffset: 3,
    instruction: 'DATE \\@ "yyyy-MM-dd"', cachedResult: 'old' }])
  model.blocks = [paragraph]
  const studio = new OfficeStudio({ document: model })
  studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id), model }
}
function request(adapter, selection = caret(5)) {
  const state = adapter.getState()
  return { unitId: state.unitId, expectedRevision: state.modelRevision, expectedHistoryGeneration: state.historyGeneration,
    inputGeneration: state.inputGeneration, selection, locale: 'en-US', timestamp }
}
const text = model => model.blocks.map(block => block.runs.map(run => run.text).join('')).join('\n')

test('standalone cache adapter maps the caret and keeps typing marks, dirty, savepoint and both history branches', () => {
  const { studio, document } = fixture()
  try {
    document.setName('First'); document.setName('Second'); document.undo(); document.save()
    document.setInputSelection(caret(5))
    document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret(5))
    const adapter = document.getFieldCacheAdapter(), stack = studio.editor.history.units[document.id]
    const history = { undo: stack.undo.length, redo: stack.redo.length, revision: stack.revision,
      saved: stack.savedRevision, generation: stack.generation, dirty: studio.editor.history.isDirty(document.id) }
    const result = adapter.refresh(request(adapter))
    assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(result.selectionCurrent, true)
    assert.equal(text(result.model), '2030-06-15Tail')
    assert.deepEqual(result.selection, caret(12))
    assert.equal(document.getInputMarks().bold, true)
    assert.equal(result.model.modified, false)
    assert.deepEqual({ undo: stack.undo.length, redo: stack.redo.length, revision: stack.revision,
      saved: stack.savedRevision, generation: stack.generation, dirty: studio.editor.history.isDirty(document.id) }, history)
    assert.equal(document.undo(), true); assert.equal(text(document.getModel()), '2030-06-15Tail')
    assert.equal(document.redo(), true); assert.equal(document.getName(), 'First')
    assert.equal(document.redo(), true); assert.equal(document.getName(), 'Second')
    assert.equal(documentTextFields(document.getModel().blocks[0])[0].endOffset, 10)
  } finally { studio.dispose() }
})

test('an opening host with a different local caret does not replace the authoritative input session or its marks', () => {
  for (const [offset, mapped] of [[1, 1], [5, 12]]) {
    const { studio, document } = fixture()
    try {
      document.setInputSelection(caret(offset))
      document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret(offset))
      const adapter = document.getFieldCacheAdapter(), before = document.getInputState()
      const result = adapter.refresh(request(adapter, caret(0)))
      assert.equal(result.accepted, true); assert.equal(result.changed, true)
      assert.equal(result.selectionCurrent, false)
      assert.deepEqual(result.selection, caret(0))
      assert.deepEqual(document.getInputState().location.selection, caret(mapped))
      assert.equal(document.getInputState().valid, true)
      assert.equal(document.getInputState().generation, before.generation + 1)
      assert.equal(document.getInputMarks().bold, true)
      assert.equal(document.getModel().modified, false)
      assert.equal(studio.editor.history.units[document.id].undo.length, 0)
    } finally { studio.dispose() }
  }
})

test('an authoritative caret in a merged tail paragraph moves with the cache while local selection remains independent', () => {
  const model = { ...createBlankDocument(), id: 'adapter-document', modified: false }
  const first = { ...createParagraph('old'), id: 'body' }, tail = { ...createParagraph('Tail'), id: 'tail' }
  setDocumentTextFields(first, [{ id: 'date', field: 'DATE', instruction: 'DATE \\@ "yyyy-MM-dd"',
    startOffset: 0, endOffset: 0, endBlockId: tail.id, cachedResult: 'old\n' }])
  model.blocks = [first, tail]
  const studio = new OfficeStudio({ document: model }), document = new FDocument(studio, model.id)
  studio.activateDocument(model.id)
  try {
    const hostCaret = { anchor: { blockId: 'tail', offset: 2 }, focus: { blockId: 'tail', offset: 2 } }
    document.setInputSelection(hostCaret)
    document.executeCommand({ id: 'doc.command.format-text', patch: { italic: true } }, hostCaret)
    const adapter = document.getFieldCacheAdapter(), result = adapter.refresh(request(adapter, caret(0)))
    assert.equal(result.changed, true); assert.equal(result.selectionCurrent, false)
    assert.equal(document.getModel().blocks.length, 1)
    assert.equal(document.getText(), '2030-06-15Tail')
    assert.deepEqual(document.getInputState().location.selection, caret(12))
    assert.equal(document.getInputMarks().italic, true)
    assert.equal(document.getModel().modified, false)
  } finally { studio.dispose() }
})

test('Studio and Facade wrappers share cache identity and one reservation for each history generation', () => {
  const { studio, document } = fixture()
  try {
    const first = document.getFieldCacheAdapter(), second = new FDocument(studio, document.id).getFieldCacheAdapter()
    assert.equal(first, second); assert.equal(first, studio.getDocumentFieldCacheAdapter(document.id))
    assert.equal(first.identity, document.regionalSettings)
    const generation = first.getState().historyGeneration
    assert.equal(reserveDocumentFieldOpenRefresh(first.identity, document.id, generation), true)
    assert.equal(reserveDocumentFieldOpenRefresh(second.identity, document.id, generation), false)
    studio.editor.history.clear(document.id)
    const reopenedGeneration = first.getState().historyGeneration
    assert.notEqual(reopenedGeneration, generation)
    assert.equal(reserveDocumentFieldOpenRefresh(second.identity, document.id, reopenedGeneration), true)
    assert.equal(releaseDocumentFieldOpenRefresh(first.identity, document.id, generation), false)
    assert.equal(releaseDocumentFieldOpenRefresh(second.identity, document.id, reopenedGeneration), true)
  } finally { studio.dispose() }
})

test('stale revision, history or input requests reject before command observation and publication', () => {
  const { studio, document } = fixture()
  try {
    const adapter = document.getFieldCacheAdapter(), input = request(adapter), commands = []
    studio.events.addEvent('BeforeCommandExecute', event => commands.push(event.id))
    for (const patch of [{ unitId: 'other' }, { expectedRevision: input.expectedRevision + 1 },
      { expectedHistoryGeneration: input.expectedHistoryGeneration + 1 }, { inputGeneration: 999 }]) {
      assert.deepEqual(adapter.refresh({ ...input, ...patch }), { accepted: false, changed: false })
    }
    assert.deepEqual(commands, []); assert.equal(document.getText(), 'oldTail')
  } finally { studio.dispose() }
})

test('BeforeCommandExecute history reentry is checked again even when the content revision is unchanged', () => {
  const { studio, document } = fixture()
  try {
    const adapter = document.getFieldCacheAdapter(), input = request(adapter)
    studio.events.addEvent('BeforeCommandExecute', event => {
      if (event.id === 'doc.command.refresh-header-footer-fields') studio.editor.history.clear(document.id)
    })
    assert.throws(() => adapter.refresh(input), /历史或打开代数/)
    assert.equal(document.getText(), 'oldTail')
    assert.equal(studio.editor.history.units[document.id].undo.length, 0)
  } finally { studio.dispose() }
})

test('a committed cache remains an accepted change if a command observer closes the document', () => {
  const { studio, document } = fixture()
  try {
    const adapter = document.getFieldCacheAdapter(), input = request(adapter)
    studio.events.addEvent('CommandExecuted', event => {
      if (event.id === 'doc.command.refresh-header-footer-fields') studio.disposeUnit(document.id)
    })
    const result = adapter.refresh(input)
    assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(result.selectionCurrent, false)
    assert.equal(text(result.model), '2030-06-15Tail')
    assert.equal(result.historyGeneration, input.expectedHistoryGeneration)
    assert.equal(adapter.getState(), undefined)
  } finally { studio.dispose() }
})

test('navigation from a post-command observer invalidates only selection restoration', () => {
  const { studio, document } = fixture()
  try {
    document.setInputSelection(caret(5))
    const adapter = document.getFieldCacheAdapter()
    studio.events.addEvent('CommandExecuted', event => {
      if (event.id === 'doc.command.refresh-header-footer-fields') document.setInputSelection(caret(0))
    })
    const result = adapter.refresh(request(adapter))
    assert.equal(result.changed, true); assert.equal(result.selectionCurrent, false)
    assert.deepEqual(document.getInputState().location.selection, caret(0))
    assert.equal(document.getText(), '2030-06-15Tail')
  } finally { studio.dispose() }
})

test('inactive document cache revisions are observable without deep traversal of its model', () => {
  const { studio, document } = fixture()
  try {
    const adapter = document.getFieldCacheAdapter(), observed = computed(() => adapter.getState()?.modelRevision)
    studio.createDocument({ id: 'another-document' })
    const before = observed.value
    const result = adapter.refresh(request(adapter))
    assert.equal(result.changed, true)
    assert.equal(observed.value, before + 1)
    assert.equal(observed.value, result.modelRevision)
    assert.equal(adapter.refresh(request(adapter)).changed, false)
  } finally { studio.dispose() }
})
