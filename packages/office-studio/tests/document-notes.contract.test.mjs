import assert from 'node:assert/strict'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, createTable, runsForTarget, tableForBlock } from '../src/features/document/model/document-model.ts'

// Updated contract drafts. This batch has not run tests, a browser or Office file acceptance.

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const { deleteDocumentNote, insertDocumentNote, normalizeDocumentNotes, noteReferencePoint, prepareDocumentNoteMutation, updateDocumentNote } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-notes.ts')
const { documentNoteInsertState } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-note-command-state.ts')
const { OfficeStudio, FDocument } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts')
afterAll(() => vite.close())

test('footnote and endnote references, definitions, navigation, and deletion share document state', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  paragraph.runs[0].text = 'reference text'
  const selected = { anchor: { blockId: paragraph.id, offset: 10 }, focus: { blockId: paragraph.id, offset: 10 } }

  const insertedFootnote = insertDocumentNote(model, selected, 'footnote')
  updateDocumentNote(model, 'footnote', insertedFootnote.note.id, 'Footnote body')
  assert.equal(insertedFootnote.host.kind, 'note')
  assert.equal(insertedFootnote.selection.focus.blockId, model.footnotes[insertedFootnote.note.id].content.blocks[0].id)
  const footnotePoint = noteReferencePoint(model, model.footnotes[insertedFootnote.note.id])
  const next = { ...footnotePoint, offset: footnotePoint.offset + 1 }
  const insertedEndnote = insertDocumentNote(model, { anchor: next, focus: next }, 'endnote')
  updateDocumentNote(model, 'endnote', insertedEndnote.note.id, 'Endnote body')

  assert.equal(model.footnotes[insertedFootnote.note.id].text, 'Footnote body')
  assert.equal(model.endnotes[insertedEndnote.note.id].text, 'Endnote body')
  assert.deepEqual(noteReferencePoint(model, insertedFootnote.note), { blockId: paragraph.id, offset: 10 })
  assert.equal(model.blocks[0].runs.some(run => run.metadata?.note?.noteId === insertedFootnote.note.id), true)
  assert.equal(deleteDocumentNote(model, 'footnote', insertedFootnote.note.id), true)
  assert.equal(model.footnotes[insertedFootnote.note.id], undefined)
  assert.equal(model.blocks[0].runs.some(run => run.metadata?.note?.noteId === insertedFootnote.note.id), false)
})

test('cross-paragraph insertion preserves selected text and uses the logical end in either selection direction', () => {
  for (const reverse of [false, true]) {
  const model = createBlankDocument()
  const first = model.blocks[0]
  first.runs[0].text = 'First paragraph'
  const second = createParagraph('Second paragraph')
  model.blocks.push(second)
  const start = { blockId: first.id, offset: 2 }, end = { blockId: second.id, offset: 6 }
  const inserted = insertDocumentNote(model, reverse ? { anchor: end, focus: start } : { anchor: start, focus: end }, 'footnote')
  assert.deepEqual(noteReferencePoint(model, inserted.note), end)
  assert.equal(model.blocks[0].runs.map(run => run.text).join(''), 'First paragraph')
  assert.equal(model.blocks[1].runs.filter(run => !run.metadata?.note).map(run => run.text).join(''), 'Second paragraph')
  assert.equal(inserted.host.noteId, inserted.note.id)
  }
})

test('nonempty note insertion shares the logical end permission with Ribbon, history and save while preserving a locked starting paragraph', () => {
  for (const kind of ['footnote', 'endnote']) {
    for (const reverse of [false, true]) {
      const source = createBlankDocument(); source.id = 'note-insert-permission-' + kind + '-' + reverse
      const first = source.blocks[0]; first.runs[0].text = 'First paragraph 😀'
      const second = createParagraph('Second paragraph 😀'); source.blocks.push(second)
      source.contentControls = [{ id: 'locked-start', type: 'richText', blockId: first.id, startOffset: 0,
        endOffset: first.runs[0].text.length, lock: 'contentLocked' }]
      source.permissionRules = [{ id: 'deny-start', scope: 'paragraph', targetId: first.id, action: 'edit', effect: 'deny', userIds: ['editor'] }]
      const start = { blockId: first.id, offset: 2 }, end = { blockId: second.id, offset: 6 }
      const selection = reverse ? { anchor: end, focus: start } : { anchor: start, focus: end }
      const permission = { documentId: source.id, userId: 'editor' }, before = structuredClone(source)
      assert.equal(documentNoteInsertState(source, selection, { permission }).available, true)
      const plan = prepareDocumentNoteMutation(source, { type: 'insert', kind, selection }, permission)
      const property = kind === 'footnote' ? 'footnotes' : 'endnotes'
      assert.equal(plan.changed, true); assert.deepEqual(source, before)
      assert.deepEqual(plan.model.blocks[0], before.blocks[0]); assert.deepEqual(plan.model.contentControls, before.contentControls)
      assert.deepEqual(noteReferencePoint(plan.model, plan.model[property][plan.noteId]), end)
      assert.equal(plan.model.blocks[1].runs.filter(run => !run.metadata?.note).map(run => run.text).join(''), 'Second paragraph 😀')

      const studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studio.activateDocument(source.id)
      const document = new FDocument(studio, source.id)
      try {
        document.setInputSelection(selection)
        const original = structuredClone(document.getModel())
        const result = document.executeNoteCommand({ type: 'insert', kind, selection }), after = structuredClone(document.getModel())
        assert.equal(result.changed, true); assert.equal(result.selectionCurrent, true)
        assert.deepEqual(after.blocks[0], original.blocks[0]); assert.deepEqual(after.contentControls, original.contentControls)
        assert.deepEqual(noteReferencePoint(after, after[property][result.noteId]), end)
        assert.deepEqual(document.getInputState().location.host, { kind: 'note', noteKind: kind, noteId: result.noteId })
        assert.equal(studio.editor.history.units[source.id].undo.length, 1)
        assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), original)
        assert.equal(document.redo(), true); assert.deepEqual(document.getModel(), after)
        const saved = structuredClone(document.saveModel()), reopened = new OfficeStudio({ document: saved, currentUserId: 'editor' })
        try { assert.deepEqual(new FDocument(reopened, saved.id).getModel(), saved) } finally { reopened.dispose() }

        document.setPermissionRules([{ id: 'deny-end', scope: 'paragraph', targetId: second.id, action: 'edit', effect: 'deny', userIds: ['editor'] }])
        document.setInputSelection(selection)
        const denied = structuredClone(document.getModel()), revision = document.getModelRevision(), history = studio.editor.history.units[source.id]
        const undo = history.undo.length, redo = history.redo.length, input = structuredClone(document.getInputState())
        assert.equal(documentNoteInsertState(denied, selection, { permission }).available, false)
        assert.throws(() => document.executeNoteCommand({ type: 'insert', kind, selection }), /permission denied/i)
        assert.deepEqual(document.getModel(), denied); assert.equal(document.getModelRevision(), revision)
        assert.equal(studio.editor.history.units[source.id].undo.length, undo)
        assert.equal(studio.editor.history.units[source.id].redo.length, redo); assert.deepEqual(document.getInputState(), input)
      } finally { studio.dispose() }
    }
  }
})

test('insertion rejects a split UTF-16 character without changing any content', () => {
  const model = createBlankDocument(); model.blocks[0].runs[0].text = 'A😀B'
  const before = structuredClone(model), point = { blockId: model.blocks[0].id, offset: 2 }
  assert.throws(() => insertDocumentNote(model, { anchor: point, focus: point }, 'footnote'), /UTF-16/)
  assert.deepEqual(model, before)
})

test('parser note collections are normalized and linked to their inline reference runs', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  paragraph.runs[0].text = '1'
  paragraph.runs[0].metadata = { noteReference: { type: 'footnote', noteId: 'n-1' } }
  const notes = normalizeDocumentNotes({ footnotes: { 'n-1': { text: 'Imported note' } } }, model.blocks)

  assert.equal(notes.footnotes['n-1'].text, 'Imported note')
  assert.equal(notes.footnotes['n-1'].referenceBlockId, paragraph.id)
  assert.equal(paragraph.runs[0].metadata.note.noteId, 'n-1')
})

test('Univer note records are separated by kind and retain data-stream paragraph text', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  const footnoteReference = createRun('1', { metadata: { noteReference: { type: 'footnote', noteId: 'footnote-1' } } })
  const endnoteReference = createRun('1', { metadata: { note: { type: 'endnote', noteId: 'endnote-1' } } })
  paragraph.runs = [footnoteReference, endnoteReference]

  const notes = normalizeDocumentNotes({
    notes: {
      footnoteKey: { noteId: 'footnote-1', type: 'footnote', body: { dataStream: 'First paragraph\rSecond paragraph\r\n\u0000' } },
      endnoteKey: { noteId: 'endnote-1', type: 'endnote', body: { dataStream: 'Endnote body\r\n' } },
      ignored: { noteId: 'ignored', type: 'comment', body: { dataStream: 'Comment body\r\n' } },
    },
  }, model.blocks)

  assert.deepEqual(Object.keys(notes.footnotes), ['footnote-1'])
  assert.deepEqual(Object.keys(notes.endnotes), ['endnote-1'])
  assert.equal(notes.footnotes['footnote-1'].text, 'First paragraph\nSecond paragraph')
  assert.equal(notes.endnotes['endnote-1'].text, 'Endnote body')
  assert.equal(notes.footnotes['footnote-1'].referenceRunId, footnoteReference.id)
  assert.equal(notes.endnotes['endnote-1'].referenceRunId, endnoteReference.id)
})

test('conflicting definitions for the same note kind and identity are rejected instead of discarding one body', () => {
  const model = createBlankDocument()
  const before = structuredClone(model.blocks)
  assert.throws(() => normalizeDocumentNotes({
    footnotes: [{ id: 'shared', text: 'Parser definition' }],
    notes: {
      sharedKey: { noteId: 'shared', type: 'footnote', body: { dataStream: 'Fallback definition\r\n' } },
      extraKey: { noteId: 'extra', type: 'footnote', body: { dataStream: 'Extra definition\r\n' } },
    },
  }, model.blocks), /冲突|重复|conflict/i)
  assert.deepEqual(model.blocks, before)
})

test('deleting an imported orphan definition never removes unrelated text', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  paragraph.runs[0].text = 'Keep this text'
  model.footnotes = { orphan: { id: 'orphan', text: 'Unlinked body', referenceBlockId: '', referenceRunId: '' } }

  assert.equal(deleteDocumentNote(model, 'footnote', 'orphan'), true)
  assert.equal(model.blocks[0].runs[0].text, 'Keep this text')
  assert.equal(model.footnotes.orphan, undefined)
})

test('reference navigation returns a point that resolves to its direct table-cell paragraph', () => {
  const model = createBlankDocument()
  const table = createTable(2, 2)
  const prefix = createParagraph('Another paragraph')
  const reference = createParagraph('Before ')
  const referenceRun = createRun('1', { metadata: { note: { type: 'endnote', noteId: 'cell-note' } } })
  reference.runs.push(referenceRun)
  table.rows[1].cells[1].blocks = [prefix, reference]
  const tableBlock = { ...createParagraph(''), type: 'table', metadata: { table } }
  model.blocks = [tableBlock]
  const note = { id: 'cell-note', text: 'Cell note', referenceBlockId: reference.id, referenceRunId: referenceRun.id }

  const point = noteReferencePoint(model, note)

  assert.deepEqual(point, { blockId: tableBlock.id, row: 1, col: 1, cellBlockId: reference.id, offset: 7 })
  assert.equal(runsForTarget(model, point)[1].id, referenceRun.id)
})

test('renumbering preserves legacy references identified only by run id', () => {
  const model = createBlankDocument()
  const first = createParagraph('1')
  const second = createParagraph('2')
  model.blocks = [first, second]
  model.footnotes = {
    first: { id: 'first', text: 'First', referenceBlockId: first.id, referenceRunId: first.runs[0].id },
    second: { id: 'second', text: 'Second', referenceBlockId: second.id, referenceRunId: second.runs[0].id },
  }

  assert.equal(deleteDocumentNote(model, 'footnote', 'first'), true)

  assert.equal(model.blocks.find(block => block.id === second.id).runs[0].text, '1')
  assert.deepEqual(noteReferencePoint(model, model.footnotes.second), { blockId: second.id, offset: 0 })
})

test('insertion enters the actual note body while callers can advance after a complete two-digit reference', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  let selection = { anchor: { blockId: paragraph.id, offset: 0 }, focus: { blockId: paragraph.id, offset: 0 } }
  let inserted
  for (let index = 0; index < 10; index++) {
    inserted = insertDocumentNote(model, selection, 'footnote')
    assert.equal(inserted.host.noteId, inserted.note.id)
    assert.equal(inserted.selection.focus.blockId, inserted.note.content.blocks[0].id)
    const point = noteReferencePoint(model, inserted.note)
    const reference = model.blocks[0].runs.find(run => run.id === inserted.note.referenceRunId)
    const after = { ...point, offset: point.offset + reference.text.length }
    selection = { anchor: after, focus: after }
  }

  assert.equal(model.blocks[0].runs.map(run => run.text).join(''), '12345678910')
  assert.deepEqual(selection.focus, { blockId: paragraph.id, offset: 11 })
  assert.deepEqual(noteReferencePoint(model, inserted.note), { blockId: paragraph.id, offset: 9 })
})

test('deleting a note removes its reference from a nested table paragraph', () => {
  const model = createBlankDocument()
  const nestedReference = createParagraph('1')
  nestedReference.runs[0].metadata = { note: { type: 'footnote', noteId: 'nested' } }
  const laterReference = createParagraph('2')
  laterReference.runs[0].metadata = { note: { type: 'footnote', noteId: 'later' } }
  const nestedTable = {
    ...createParagraph(''),
    type: 'table',
    metadata: { table: { rows: [{ id: 'row-1', cells: [{ id: 'cell-1', blocks: [nestedReference, laterReference], style: {}, rowSpan: 1, colSpan: 1 }] }] } },
  }
  model.blocks[0].type = 'table'
  model.blocks[0].metadata = {
    table: { rows: [{ id: 'outer-row', cells: [{ id: 'outer-cell', blocks: [nestedTable], style: {}, rowSpan: 1, colSpan: 1 }] }] },
  }
  model.footnotes = {
    nested: { id: 'nested', text: 'Nested note', referenceBlockId: nestedReference.id, referenceRunId: nestedReference.runs[0].id },
    later: { id: 'later', text: 'Later note', referenceBlockId: laterReference.id, referenceRunId: laterReference.runs[0].id },
  }

  // The nested reference exists, but DocumentPoint has no nested cell path.
  assert.equal(noteReferencePoint(model, model.footnotes.nested), undefined)
  assert.equal(noteReferencePoint(model, model.footnotes.later), undefined)
  assert.equal(deleteDocumentNote(model, 'footnote', 'nested'), true)
  const liveNested = tableForBlock(tableForBlock(model.blocks[0]).rows[0].cells[0].blocks[0]).rows[0].cells[0].blocks
  assert.equal(liveNested[0].runs.some(run => run.metadata?.note?.noteId === 'nested'), false)
  assert.deepEqual(liveNested[0].runs.map(run => run.text), [''])
  assert.equal(liveNested[1].runs[0].text, '1')
  assert.equal(model.footnotes.nested, undefined)
})
