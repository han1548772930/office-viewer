import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-13/14/16/17 contract drafts. Execution is deferred to unified verification.
let server, domain, modelUtils, structure, fields, notes, appearance, codeSource, codeBlocks
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-structured-blocks', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;[domain, modelUtils, structure, fields, notes, appearance, codeSource, codeBlocks] = await Promise.all([
    'domain/model', 'features/document/model/document-model', 'features/document/commands/document-structure-commands',
    'features/document/model/document-text-fields', 'features/document/notes/document-notes',
    'features/document/model/document-structured-appearance', 'features/document/model/document-code-source', 'features/document/code-blocks/code-blocks',
  ].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterAll(async () => { await server?.close() })
const copy = value => structuredClone(value)
const point = (blockId, offset = 0) => ({ blockId, offset })
const selection = (anchor, focus = anchor) => ({ anchor, focus })
const paragraph = (id, text = '') => ({ ...modelUtils.createParagraph(text), id })
const documentWith = blocks => ({ ...domain.createBlankDocument(), id: 'structured-contract', blocks })
const current = (model, id) => modelUtils.findDocumentBlock(model, id)
function quote(id, children) { return { ...paragraph(id), type: 'quote', runs: [], children } }
function code(id, children) { return { ...paragraph(id), type: 'code', runs: [], children, code: { language: 'typescript', showLineNumbers: true, tabSize: 4, wrap: true } } }
function execute(model, command, range = selection(point(modelUtils.textTargets(model)[0].blockId))) {
  return structure.executeDocumentStructureCommand(model, range, command, { documentId: model.id, userId: 'editor', defaultAllow: true })
}
const dateField = (id, startOffset, endOffset, extra = {}) => ({ id, field: 'DATE', instruction: 'DATE', cachedResult: 'old', startOffset, endOffset, ...extra })
function insertColumns(model, count = 2) {
  execute(model, { id: 'doc.command.insert-column-group', columnGroupId: 'columns', columnCount: count,
    columnIds: Array.from({ length: count }, (_, index) => 'column-' + index), afterBlockId: model.blocks[0].id })
  return current(model, 'columns')
}

test('wrapping respects either selection direction and an exclusive next-paragraph boundary', () => {
  for (const reverse of [false, true]) {
    const model = documentWith([paragraph('one', 'A😀B'), paragraph('two', 'tail')]), original = copy(model.blocks)
    const range = selection(point('one', 1), point('two', 0))
    if (reverse) [range.anchor, range.focus] = [range.focus, range.anchor]
    execute(model, { id: 'doc.command.wrap-structured-block', blockId: 'quote', blockType: 'quote' }, range)
    assert.deepEqual(current(model, 'quote').children, [original[0]])
    assert.deepEqual(current(model, 'two'), original[1])
    execute(model, { id: 'doc.command.unwrap-structured-block', blockId: 'quote' }, range)
    assert.deepEqual(model.blocks, original)
  }
})

test('quote reads imported paragraph-rule color and applies text color to actual runs', () => {
  const child = paragraph('inside', 'text')
  child.metadata.univerBodyParagraphSource = { version: 1, paragraph: { paragraphStyle: { borderLeft: { color: { rgb: '#987654' } } } } }
  const model = documentWith([quote('quote', [child])])
  assert.equal(appearance.documentQuoteLineColor(model.blocks[0]), '#987654')
  execute(model, { id: 'doc.command.update-quote-style', blockId: 'quote', patch: { lineColor: '#123456', textColor: '#aabbcc' } })
  assert.equal(appearance.documentQuoteLineColor(current(model, 'quote')), '#123456')
  assert.equal(current(model, 'inside').runs[0].color, '#aabbcc')
  assert.equal(modelUtils.runText(current(model, 'inside').runs), 'text')
})

test('callout configuration changes preserve mixed text colors until color is explicitly changed', () => {
  const child = paragraph('inside')
  child.runs = [modelUtils.createRun('red', { color: '#ff0000', bold: true }), modelUtils.createRun('blue', { color: '#0000ff', italic: true })]
  const model = documentWith([child]), original = copy(child.runs)
  execute(model, { id: 'doc.command.wrap-structured-block', blockType: 'callout', blockId: 'callout', blockIds: ['inside'] })
  execute(model, { id: 'doc.command.update-callout', blockId: 'callout', patch: { icon: '💡', showIcon: false, borderOpacity: 0.4 } })
  assert.deepEqual(current(model, 'inside').runs, original)
  execute(model, { id: 'doc.command.update-callout', blockId: 'callout', patch: {}, textColor: '#334455' })
  assert.deepEqual(current(model, 'inside').runs.map(run => [run.text, run.color]), [['red', '#334455'], ['blue', '#334455']])
  assert.equal(current(model, 'inside').runs[0].bold, true)
  assert.equal(current(model, 'inside').runs[1].italic, true)
})

test('deleting a wrapper rejects locked fields, inline controls and complete block controls without publishing a candidate', () => {
  for (const setup of [
    model => fields.setDocumentTextFields(current(model, 'inside'), [dateField('locked', 0, 3, { metadata: { source: { locked: true } } })]),
    model => { model.contentControls = [{ id: 'inline', type: 'text', blockId: 'inside', startOffset: 0, endOffset: 3, lock: 'sdtLocked' }] },
    model => { model.contentControls = [{ id: 'block', type: 'group', blockId: 'quote', startOffset: 0, endOffset: 0, blockRange: { blockIds: ['quote'] }, lock: 'sdtLocked' }] },
  ]) {
    const model = documentWith([quote('quote', [paragraph('inside', 'old')]), paragraph('outside', 'after')]); setup(model)
    const before = copy(model)
    assert.throws(() => execute(model, { id: 'doc.command.remove-structured-block', blockId: 'quote' }, selection(point('inside', 1))))
    assert.deepEqual(model, before)
  }
})

test('deleting a wrapper cleans controls, comment anchors and note stories, renumbers surviving references and moves the input point', () => {
  const model = documentWith([quote('quote', [paragraph('inside', 'old')]), paragraph('outside', 'after')])
  const removedNote = notes.insertDocumentNote(model, selection(point('inside', 3)), 'footnote').note.id
  const retainedNote = notes.insertDocumentNote(model, selection(point('outside', 5)), 'footnote').note.id
  model.contentControls = [{ id: 'inline', type: 'text', blockId: 'inside', startOffset: 0, endOffset: 3, lock: 'unlocked' }]
  const anchor = (blockId, text) => ({ blockId, startOffset: 0, endOffset: text.length, text })
  model.comments = { thread: { id: 'thread', anchor: [anchor('inside', 'old'), anchor('outside', 'after')], referenceText: 'old\nafter', messages: [], createdAt: '2026-11-25', updatedAt: '2026-11-25' } }
  const result = execute(model, { id: 'doc.command.remove-structured-block', blockId: 'quote' }, selection(point('inside', 1)))
  assert.equal(current(model, 'inside'), undefined)
  assert.deepEqual(model.contentControls, [])
  assert.deepEqual(model.comments.thread.anchor, [anchor('outside', 'after')])
  assert.equal(model.comments.thread.referenceText, 'after')
  assert.equal(model.footnotes[removedNote], undefined)
  assert.equal(model.footnotes[retainedNote].referenceBlockId, 'outside')
  assert.equal(current(model, 'outside').runs.find(run => run.metadata?.note?.noteId === retainedNote).text, '1')
  assert.deepEqual(result, selection(point('outside')))
})

test('deleting the only content container creates a real default paragraph and valid caret', () => {
  const model = documentWith([quote('quote', [paragraph('inside', 'text')])])
  const result = execute(model, { id: 'doc.command.remove-structured-block', blockId: 'quote' }, selection(point('inside', 2)))
  assert.equal(model.blocks.length, 1); assert.equal(model.blocks[0].type, 'paragraph')
  assert.equal(modelUtils.runText(model.blocks[0].runs), '')
  assert.deepEqual(result, selection(point(model.blocks[0].id)))
})

test('cross-paragraph fields retain surviving endpoints; a locked result rejects the deletion atomically', () => {
  for (const locked of [false, true]) {
    const model = documentWith([paragraph('before', 'A'), quote('quote', [paragraph('inside', 'BB')]), paragraph('after', 'C')])
    fields.setDocumentTextFields(current(model, 'before'), [dateField('cross', 0, 2, { endBlockId: 'inside', cachedResult: 'A\nBB', metadata: { source: { locked } } })])
    const before = copy(model), command = { id: 'doc.command.remove-structured-block', blockId: 'quote' }
    if (locked) { assert.throws(() => execute(model, command)); assert.deepEqual(model, before) }
    else {
      execute(model, command)
      const field = fields.documentTextFields(current(model, 'before'))[0]
      assert.equal(field.endBlockId, undefined); assert.equal(field.endOffset, 1); assert.equal(field.cachedResult, 'A')
    }
  }
})

test('column creation preserves custom identities for all 2–5 sizes and rejects invalid identities and counts atomically', () => {
  for (const count of [2, 3, 4, 5]) {
    const model = documentWith([paragraph('body')]), group = insertColumns(model, count)
    assert.deepEqual(group.columns.columns.map(column => [column.id, column.widthRatio]), Array.from({ length: count }, (_, index) => ['column-' + index, 1]))
  }
  for (const options of [{ columnCount: 1 }, { columnCount: 6 }, { columnCount: 2, columnIds: ['same', 'same'] },
    { columnCount: 2, columnIds: ['columns', 'other'] }, { columnCount: 2, widthRatios: [1, 0] }]) {
    const model = documentWith([paragraph('body')]), before = copy(model)
    assert.throws(() => execute(model, { id: 'doc.command.insert-column-group', columnGroupId: 'columns', afterBlockId: 'body', ...options }))
    assert.deepEqual(model, before)
  }
})

test('multiline column insert creates actual paragraphs in the target container and retains the requested first ID', () => {
  const model = documentWith([paragraph('body')]), group = insertColumns(model)
  group.columns.columns[0].blocks = [quote('nested', [paragraph('existing', 'tail')])]
  execute(model, { id: 'doc.command.insert-column-paragraph', blockId: 'columns', columnId: 'column-0', index: 0, paragraphId: 'inserted', text: 'A\r\n😀\rC' })
  const content = current(model, 'nested').children
  assert.equal(content[0].id, 'inserted')
  assert.deepEqual(content.map(block => modelUtils.runText(block.runs)), ['A', '😀', 'C', 'tail'])
  const before = copy(model)
  assert.throws(() => execute(model, { id: 'doc.command.insert-column-paragraph', blockId: 'columns', columnId: 'column-0', index: 1, paragraphId: 'inserted', text: 'duplicate' }))
  assert.deepEqual(model, before)
})

test('two columns cannot be deleted individually and destructive deletion removes content while keeping remaining IDs', () => {
  const model = documentWith([paragraph('body')]); insertColumns(model, 3)
  const removed = current(model, 'columns').columns.columns[1].blocks[0].id
  execute(model, { id: 'doc.command.remove-column', blockId: 'columns', columnId: 'column-1', keepContents: false }, selection(point(removed)))
  assert.equal(current(model, removed), undefined)
  assert.deepEqual(current(model, 'columns').columns.columns.map(column => column.id), ['column-0', 'column-2'])
  const before = copy(model)
  assert.throws(() => execute(model, { id: 'doc.command.remove-column', blockId: 'columns', columnId: 'column-0', keepContents: false }))
  assert.deepEqual(model, before)
})

test('moving columns cannot reverse a field spanning their contents', () => {
  const model = documentWith([paragraph('body')]), group = insertColumns(model, 3)
  group.columns.columns[0].blocks = [paragraph('left', 'A')]; group.columns.columns[1].blocks = [paragraph('right', 'B')]
  fields.setDocumentTextFields(current(model, 'left'), [dateField('cross', 0, 1, { endBlockId: 'right', cachedResult: 'A\nB' })])
  const before = copy(model)
  assert.throws(() => execute(model, { id: 'doc.command.move-column', blockId: 'columns', columnId: 'column-1', index: 0 }), /字段/)
  assert.deepEqual(model, before)
})

test('code indent plans are descending, handle emoji and exclude a selected next-line start', () => {
  const source = 'A😀\nB\nC'
  assert.deepEqual(codeSource.documentCodeIndentEdits(source, 1, 6, 1, 4), [
    { startOffset: 4, endOffset: 4, text: '\t' }, { startOffset: 0, endOffset: 0, text: '\t' },
  ])
  assert.deepEqual(codeSource.documentCodeIndentEdits('    A\n\tB\n  C', 0, 12, -1, 4), [
    { startOffset: 9, endOffset: 11, text: '' }, { startOffset: 6, endOffset: 7, text: '' }, { startOffset: 0, endOffset: 4, text: '' },
  ])
  assert.deepEqual(codeSource.documentCodeIndentEdits(source, 3, 3, 1, 4), [{ startOffset: 3, endOffset: 3, text: '\t' }])
  assert.throws(() => codeSource.documentCodeIndentEdits(source, 2, 2, 1, 4), /范围/)
})

test('code indentation preserves reverse selections and rich-text marks in one outer candidate', () => {
  const first = paragraph('first', 'a😀b'), last = paragraph('last', '  c')
  first.runs[0].bold = true
  const model = documentWith([code('code', [first, last])]), originalRuns = copy(first.runs)
  const range = selection(point('last', 3), point('first', 1))
  const indented = execute(model, { id: 'doc.command.indent-code-block', blockId: 'code', direction: 1 }, range)
  assert.equal(codeBlocks.codeBlockData(current(model, 'code')).source, '\ta😀b\n\t  c')
  assert.deepEqual(indented, selection(point('last', 4), point('first', 2)))
  assert.ok(current(model, 'first').runs.some(run => run.bold && run.text.includes('a😀b')))
  const restored = execute(model, { id: 'doc.command.indent-code-block', blockId: 'code', direction: -1 }, indented)
  assert.equal(codeBlocks.codeBlockData(current(model, 'code')).source, 'a😀b\n  c')
  assert.deepEqual(restored, range)
  assert.equal(current(model, 'first').runs.some(run => run.id === originalRuns[0].id && run.bold), true)
})

test('a later locked outdent rejects all previously planned code edits', () => {
  const model = documentWith([code('code', [paragraph('first', '  A'), paragraph('last', '  B')])])
  fields.setDocumentTextFields(current(model, 'first'), [dateField('locked', 0, 3, { cachedResult: '  A', metadata: { source: { locked: true } } })])
  const before = copy(model)
  assert.throws(() => execute(model, { id: 'doc.command.indent-code-block', blockId: 'code', direction: -1 }, selection(point('first', 0), point('last', 3))))
  assert.deepEqual(model, before)
})
