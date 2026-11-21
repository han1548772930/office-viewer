import assert from 'node:assert/strict'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, createTable, runsForTarget } from '../src/features/document/model/document-model.ts'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const { deleteDocumentNote, insertDocumentNote, normalizeDocumentNotes, noteReferencePoint, updateDocumentNote } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-notes.ts')
afterAll(() => vite.close())

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

test('explicit parser note collections take precedence when merging Univer note records', () => {
  const model = createBlankDocument()
  const notes = normalizeDocumentNotes({
    footnotes: [{ id: 'shared', text: 'Parser definition' }],
    notes: {
      sharedKey: { noteId: 'shared', type: 'footnote', body: { dataStream: 'Fallback definition\r\n' } },
      extraKey: { noteId: 'extra', type: 'footnote', body: { dataStream: 'Extra definition\r\n' } },
    },
  }, model.blocks)

  assert.deepEqual(Object.keys(notes.footnotes), ['shared', 'extra'])
  assert.equal(notes.footnotes.shared.text, 'Parser definition')
  assert.equal(notes.footnotes.extra.text, 'Extra definition')
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

  assert.equal(second.runs[0].text, '1')
  assert.deepEqual(noteReferencePoint(model, model.footnotes.second), { blockId: second.id, offset: 0 })
})

test('inserting a two-digit note keeps the caret after the full reference', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  let selection = { anchor: { blockId: paragraph.id, offset: 0 }, focus: { blockId: paragraph.id, offset: 0 } }
  let inserted
  for (let index = 0; index < 10; index++) {
    inserted = insertDocumentNote(model, selection, 'footnote')
    selection = inserted.selection
  }

  assert.equal(paragraph.runs.map(run => run.text).join(''), '12345678910')
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
  assert.equal(nestedReference.runs.some(run => run.metadata?.note?.noteId === 'nested'), false)
  assert.deepEqual(nestedReference.runs.map(run => run.text), [''])
  assert.equal(laterReference.runs[0].text, '1')
  assert.equal(model.footnotes.nested, undefined)
})
