import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred DOC-11/DOC-22 import contracts. These drafts have not been run.
let server, model, projection, native, normalizer
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/native-controls-contract', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;[model, projection, native, normalizer] = await Promise.all([
    'domain/model', 'parser/documentTextProjection', 'parser/documentContentControls', 'parser/documentNormalizer',
  ].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterAll(async () => { await server?.close() })
const copy = value => structuredClone(value)
const source = (id, text = '', children = []) => ({ id, type: 'paragraph', content: text,
  details: { inlines: [{ type: 'text', text }] }, children })
const dto = (id, properties, range = {}) => ({ id, partName: 'word/document.xml', story: 'body',
  sourceXml: '<w:sdt><w:sdtContent/></w:sdt>', properties, textRanges: [], ...range })
function formulaSource() {
  const inlines = [{ type: 'text', text: '😀A' }, { type: 'form_field', text: 'x^2', blockId: 'formula-one' },
    { type: 'text', text: ' between ' }, { type: 'form_field', text: 'y+1', blockId: 'formula-two' },
    { type: 'text', text: ' tail' }]
  return { id: 'formula-p', type: 'paragraph', content: '😀Ax^2 between y+1 tail', details: { inlines },
    children: ['formula-one', 'formula-two'].map(id => ({ id, type: 'form_field',
      details: { field: { tag: 'office-latex:v1:inline' } } })) }
}
function projectedBlock(raw = formulaSource()) {
  return { id: raw.id, type: 'paragraph', style: {}, metadata: { source: copy(raw) },
    runs: raw.details.inlines.map((inline, index) => ({ id: 'run-' + index,
      text: inline.type === 'form_field' ? '\uFFFC' : inline.text,
      metadata: { source: copy(inline), ...(inline.type === 'form_field' ? { latexFormula: { latex: inline.text, mode: 'inline', version: 1 } } : {}) } })) }
}
function envelope(blocks, nativeContentControls, comments) {
  return { format: 'docx', fileName: 'native-contract.docx', payload: { data: { documentId: 'native-contract', blocks, nativeContentControls, comments } } }
}

test('two formula atoms map UTF-16 boundaries without mutating literal source or splitting emoji', () => {
  const block = projectedBlock(), original = copy(block.metadata.source)
  const mapping = projection.parsedDocumentTextProjection(block)
  assert.equal(mapping.sourceText, '😀Ax^2 between y+1 tail')
  assert.equal(mapping.targetText, '😀A\uFFFC between \uFFFC tail')
  assert.deepEqual([0, 2, 3, 6, 15, 18, 23].map(offset => mapping.mapBoundary(offset)), [0, 2, 3, 4, 13, 14, 19])
  for (const offset of [1, 4, 5, 16, 17, -1, 24]) assert.equal(mapping.mapBoundary(offset), undefined, String(offset))
  assert.deepEqual(block.metadata.source, original)
})

test('native controls and comments after two formulas use the same projected paragraph coordinates', () => {
  const raw = formulaSource(), input = envelope([raw], [dto('tail-control', { type: 'text' },
    { textRanges: [{ blockId: raw.id, startOffset: 18, endOffset: 23 }] })],
  [{ id: 7, text: 'Review tail', targets: [{ story: 'body', nodeId: raw.id, start: 18, end: 23 }] }])
  const original = copy(input), result = normalizer.normalizeDocument(input)
  assert.equal(result.blocks[0].runs.map(run => run.text).join(''), '😀A\uFFFC between \uFFFC tail')
  assert.deepEqual([result.contentControls[0].startOffset, result.contentControls[0].endOffset], [14, 19])
  assert.equal(result.contentControls[0].value, ' tail')
  const anchor = result.comments['docx-comment-7'].anchor[0]
  assert.deepEqual([anchor.startOffset, anchor.endOffset, anchor.text], [14, 19, ' tail'])
  assert.deepEqual(input, original)
})

test('cross-paragraph field ends are projected independently and retained source collections stay literal', () => {
  const first = projectedBlock(), last = projectedBlock({ ...formulaSource(), id: 'last-p' })
  first.metadata.source.details.documentTextFields = { version: 1, fields: [
    { id: 'cross-field', startOffset: 6, endOffset: 23, endBlockId: last.id, instruction: 'REF target' },
  ] }
  const original = copy(first.metadata.source.details.documentTextFields)
  projection.projectParsedDocumentFieldRanges([first, last])
  const field = first.metadata.documentTextFields.fields[0]
  assert.deepEqual([field.startOffset, field.endOffset, field.endBlockId], [4, 19, 'last-p'])
  assert.deepEqual(first.metadata.source.details.documentTextFields, original)
})

test('ambiguous formula interiors keep a whole native subtree unmapped and reject editable field ranges', () => {
  const document = model.createBlankDocument(), block = projectedBlock()
  document.blocks = [block]
  const parent = dto('parent', { type: 'group' }, { textRanges: [{ blockId: block.id, startOffset: 4, endOffset: 6 }] })
  const child = dto('child', { type: 'text' }, { parentId: 'parent', textRanges: [{ blockId: block.id, startOffset: 6, endOffset: 15 }] })
  const original = copy([parent, child]), hydration = native.normalizeNativeDocumentContentControls(document, [parent, child])
  assert.equal(hydration.hydratedCount, 0)
  assert.deepEqual(new Set(hydration.unmapped.map(item => item.control.id)), new Set(['parent', 'child']))
  assert.deepEqual([parent, child], original)
  block.metadata.source.details.documentTextFields = { version: 1, fields: [{ id: 'bad', startOffset: 4, endOffset: 6 }] }
  assert.throws(() => projection.projectParsedDocumentFieldRanges([block]), RangeError)
})

test('an empty formula source has no guessed ownership at its expanded atom boundary', () => {
  const raw = formulaSource(); raw.details.inlines[1].text = ''
  const block = projectedBlock(raw), mapping = projection.parsedDocumentTextProjection(block)
  assert.equal(mapping.mapBoundary(3), undefined)
  assert.equal(mapping.mapBoundary(12), 13)
})

test('block SDTs include physical paragraph images inside real table cells and wrappers keep picture identity', () => {
  const image = { id: 'figure', type: 'figure', details: { src: 'data:image/png;base64,AA==', widthPx: 12, heightPx: 8 } }
  const paragraph = source('cell-p', '', [image])
  const table = { id: 'table', type: 'table', details: { rowCount: 1, columnCount: 1 },
    children: [{ id: 'cell', type: 'table_cell', details: { rowIndex: 0, columnIndex: 0, rowSpan: 1, columnSpan: 1 }, children: [paragraph] }] }
  const controls = [dto('wrapper', { type: 'group' }, { blockIds: ['table'], hasObjectContent: true, objectContentCount: 1 }),
    dto('picture', { type: 'picture' }, { parentId: 'wrapper', blockIds: ['cell-p'], hasObjectContent: true, objectContentCount: 1 })]
  const result = normalizer.normalizeDocument(envelope([table], controls))
  const blocks = result.blocks[0].metadata.table.rows[0].cells[0].blocks
  assert.deepEqual(blocks.map(block => block.id), ['cell-p', 'figure'])
  assert.equal(blocks[1].metadata.officeSourceParagraphId, 'cell-p')
  assert.equal(result.metadata.nativeContentControlSource.hydratedCount, 2)
  const picture = result.contentControls.find(control => control.id === 'picture')
  assert.deepEqual(picture.blockRange.blockIds, ['cell-p', 'figure'])
  assert.equal(picture.metadata.pictureBlockId, 'figure')
  assert.equal(picture.metadata.pictureAnchorBlockId, 'cell-p')
  assert.equal(picture.value, image.details.src)
  assert.equal(blocks[1].metadata.contentControlId, 'picture')
})

test('a native picture with incomplete objects stays retained; empty native placeholders remain editable', () => {
  const document = model.createBlankDocument(); document.blocks = [projectedBlock({ id: 'p', details: { inlines: [{ type: 'text', text: '' }] } })]
  const missing = dto('missing', { type: 'picture' }, { blockIds: ['p'], hasObjectContent: true, objectContentCount: 1 })
  assert.equal(native.normalizeNativeDocumentContentControls(document, [missing]).hydratedCount, 0)
  const placeholder = dto('placeholder', { type: 'picture', showingPlaceholder: true },
    { textRanges: [{ blockId: 'p', startOffset: 0, endOffset: 0 }] })
  assert.equal(native.normalizeNativeDocumentContentControls(document, [placeholder]).hydratedCount, 1)
  assert.equal(document.contentControls[0].value, '')
})
