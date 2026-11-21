import assert from 'node:assert/strict'
import { test, after } from 'node:test'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph } from '../src/features/document/model/document-model.ts'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const { deleteDocumentNote, insertDocumentNote, normalizeDocumentNotes, noteReferencePoint, updateDocumentNote } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-notes.ts')
after(() => vite.close())

test('footnote and endnote references, definitions, navigation, and deletion share document state', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  paragraph.runs[0].text = 'reference text'
  const selected = { anchor: { blockId: paragraph.id, offset: 10 }, focus: { blockId: paragraph.id, offset: 10 } }

  const insertedFootnote = insertDocumentNote(model, selected, 'footnote')
  updateDocumentNote(model, 'footnote', insertedFootnote.note.id, 'Footnote body')
  const insertedEndnote = insertDocumentNote(model, insertedFootnote.selection, 'endnote')
  updateDocumentNote(model, 'endnote', insertedEndnote.note.id, 'Endnote body')

  assert.equal(model.footnotes[insertedFootnote.note.id].text, 'Footnote body')
  assert.equal(model.endnotes[insertedEndnote.note.id].text, 'Endnote body')
  assert.deepEqual(noteReferencePoint(model, insertedFootnote.note), { blockId: paragraph.id, offset: 10 })
  assert.equal(paragraph.runs.some(run => run.metadata?.note?.noteId === insertedFootnote.note.id), true)
  assert.equal(deleteDocumentNote(model, 'footnote', insertedFootnote.note.id), true)
  assert.equal(model.footnotes[insertedFootnote.note.id], undefined)
  assert.equal(paragraph.runs.some(run => run.metadata?.note?.noteId === insertedFootnote.note.id), false)
})

test('note insertion rejects selections crossing paragraph or cell boundaries without partial mutation', () => {
  const model = createBlankDocument()
  const before = structuredClone(model)
  const first = model.blocks[0]
  const second = createParagraph('Second paragraph')
  model.blocks.push(second)
  assert.throws(() => insertDocumentNote(model, { anchor: { blockId: first.id, offset: 0 }, focus: { blockId: second.id, offset: 0 } }, 'footnote'))
  before.blocks.push(second)
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

test('deleting an imported orphan definition never removes unrelated text', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  paragraph.runs[0].text = 'Keep this text'
  model.footnotes = { orphan: { id: 'orphan', text: 'Unlinked body', referenceBlockId: '', referenceRunId: '' } }

  assert.equal(deleteDocumentNote(model, 'footnote', 'orphan'), true)
  assert.equal(paragraph.runs[0].text, 'Keep this text')
  assert.equal(model.footnotes.orphan, undefined)
})
