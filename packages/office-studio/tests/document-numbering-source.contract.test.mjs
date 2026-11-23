// Draft contracts for the final unified verification phase; no WASM is run.
import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

let server, OfficeStudio, FDocument, normalizeDocument, validateDocumentOfficeSource
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;({ normalizeDocument } = await server.ssrLoadModule('/packages/office-studio/src/parser/documentNormalizer.ts'))
  ;({ validateDocumentOfficeSource } = await server.ssrLoadModule('/packages/office-studio/src/parser/documentOfficeSource.ts'))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
function source() {
  return { version: 1, numbering: { partName: 'word/numbering.xml',
    xml: `<w:numbering xmlns:w="${WORD}"><w:num w:numId="7"><w:abstractNumId w:val="2"/></w:num></w:numbering>`,
    relationshipsXml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="picture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/bullet.png"/></Relationships>',
    resources: [{ partName: 'word/media/bullet.png', contentType: 'image/png', dataBase64: 'AQID' }] } }
}
function envelope(officeSource) {
  return { schemaVersion: 1, parserVersion: 'draft', format: 'docx', fileName: 'numbering.docx', size: 1, model: 'document',
    payload: { modelType: 'document', data: { documentId: 'numbering-source', officeSource, blocks: [
      { id: 'paragraph', type: 'paragraph', content: 'Native numbered content', details: { inlines: [{ type: 'text', text: 'Native numbered content' }] } },
    ] } }, diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] } }
}
function open(model) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}

test('real numbering XML, relationships and binary media remain provenance through an edit, undo/redo and snapshot reopening', () => {
  const provenance = source(), original = copy(provenance)
  const model = normalizeDocument(envelope(provenance))
  assert.deepEqual(model.metadata.source.officeSource, original)
  const { document } = open(model)
  const selection = { anchor: { blockId: 'paragraph', offset: 0 }, focus: { blockId: 'paragraph', offset: 0 } }
  assert.equal(document.executeCommand({ id: 'doc.command.insert-text', text: '编辑😀' }, selection).changed, true)
  const edited = copy(document.getModel())
  assert.deepEqual(edited.metadata.source.officeSource, original)
  assert.equal(document.undo(), true)
  assert.deepEqual(document.getModel().metadata.source.officeSource, original)
  assert.equal(document.redo(), true)
  const reopened = open(copy(document.save()))
  assert.deepEqual(reopened.document.getModel().metadata.source.officeSource, original)
  assert.deepEqual(provenance, original)
})

test('numbering source validation rejects unsupported fields, duplicate canonical paths and malformed Base64 without changing the DTO', () => {
  const cases = [
    value => { value.numbering.oldNumberingEngine = {} },
    value => { value.numbering.partName = '../numbering.xml' },
    value => { value.numbering.resources[0].partName = 'word\\media\\bullet.png' },
    value => { value.numbering.resources[0].partName = 'https://example.test/bullet.png' },
    value => { value.numbering.resources.push({ ...value.numbering.resources[0], partName: '/word/media/bullet.png' }) },
    value => { value.numbering.resources[0].contentType = 'image/png\ninvalid' },
    value => { value.numbering.resources[0].extraBinaryField = 'unmapped' },
    value => { value.numbering.resources[0].dataBase64 = 'AR==' },
    value => { value.numbering.resources[0].dataBase64 = 'AQJ=' },
    value => { value.numbering.resources[0].dataBase64 = 'AQ==' + '\n' },
    value => { value.numbering.xml = '\uD800' },
    value => { value.numbering.resources = null },
  ]
  for (const mutate of cases) {
    const value = source(); mutate(value); const original = copy(value)
    assert.throws(() => validateDocumentOfficeSource(value, 'docx'))
    assert.deepEqual(value, original)
  }
})

test('numbering source budgets include decoded media and ODT cannot claim a native Word numbering part', () => {
  const empty = source(); delete empty.numbering.resources; delete empty.numbering.relationshipsXml
  validateDocumentOfficeSource(empty, 'docx')
  const origin = { format: 'odt', synthesized: false, resources: [] }
  assert.throws(() => validateDocumentOfficeSource({ ...source(), origin }, 'odt'), TypeError)
  const bytes = 16 * 1024 * 1024
  const tooLarge = source(); tooLarge.numbering.xml = 'x'.repeat(bytes + 1)
  assert.throws(() => validateDocumentOfficeSource(tooLarge, 'docx'), RangeError)
  const combined = source(); combined.numbering.xml = 'x'.repeat(bytes); combined.numbering.relationshipsXml = 'x'.repeat(bytes)
  assert.throws(() => validateDocumentOfficeSource(combined, 'docx'), RangeError)
  const count = source(); count.numbering.resources = Array.from({ length: 1025 }, (_, index) => ({ partName: `word/media/${index}.png`, contentType: 'image/png', dataBase64: 'AQ==' }))
  assert.throws(() => validateDocumentOfficeSource(count, 'docx'), TypeError)
})
