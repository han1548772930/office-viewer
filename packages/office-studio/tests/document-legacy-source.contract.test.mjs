// Drafts only: not executed in the 2026-11-22 continuation. These public API
// snapshot contracts complement native Office-byte drafts; they do not run WASM.
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
const docOrigin = (offset = 32, dataHex = '00AbFF') => ({ format: 'doc', synthesized: true, resources: [{ kind: 'binary', streamName: '1Table', offset, dataHex }] })
const odtStyles = '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:extra="urn:test"><extra:unmodelled value="保留😀"/></office:document-styles>'
const odtContent = '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body/></office:document-content>'
const odtOrigin = () => ({ format: 'odt', synthesized: false, resources: [{ kind: 'xml', partName: 'styles.xml', xml: odtStyles }, { kind: 'xml', partName: 'content.xml', xml: odtContent }] })
const source = format => ({ version: 1, origin: format === 'doc' ? docOrigin() : odtOrigin(),
  ...(format === 'doc' ? { styles: { partName: 'word/styles.xml', xml: '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>' } } : {}) })

function envelope(format, officeSource, fontSize = format === 'odt' ? 18 : 16) {
  return { schemaVersion: 1, parserVersion: 'draft', format, fileName: `defaults.${format}`, size: 1, model: 'document',
    payload: { modelType: 'document', data: { documentId: `legacy-${format}`, officeSource,
      defaultStyle: { version: 1, text: { fontFamilyAscii: 'Georgia', fontFamilyEastAsia: '宋体', fontFamilyCs: 'Amiri', fontSizePt: fontSize, color: 'auto', bold: false, italic: true },
        paragraph: { align: 'right', indent: { leftTwips: 720 }, spacing: { beforeTwips: 0, afterTwips: 180, lineTwips: 288, lineRule: 'auto' } } },
      blocks: [{ id: 'leading', type: 'paragraph', content: 'First styled paragraph', details: { style: { align: 'center' }, inlines: [{ type: 'text', text: 'First styled paragraph',
        style: { fontFamilyAscii: 'Courier New', fontSizePt: 42, color: '#FF0000', bold: true }, sourceFormat: { unmodelled: 'retained outside provenance schema' } }] } },
      { id: 'next', type: 'paragraph', content: 'Next paragraph', details: { style: { align: 'left' }, inlines: [{ type: 'text', text: 'Next paragraph', style: { fontSizePt: 20 } }] } }],
      unmodelledPayload: { original: 'retained' } } }, diagnostics: { warnings: [], errors: [], stats: {}, capabilities: ['documentDefaultStyle'] } }
}

function open(model) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}

for (const format of ['doc', 'odt']) test(`${format} defaults drive public clear, undo/redo, internal save/reopen, and the next input while preserving source resources`, () => {
  const dto = envelope(format, source(format)), imported = normalizeDocument(dto)
  const originalSource = copy(dto.payload.data.officeSource)
  assert.deepEqual(imported.metadata.source.officeSource, originalSource)
  assert.deepEqual(imported.metadata.source.unmodelledPayload, { original: 'retained' })
  assert.equal(imported.blocks[0].runs[0].fontSize, 42)
  assert.equal(imported.defaultStyle.text.fontSize, format === 'doc' ? 16 : 18)
  assert.notEqual(imported.defaultStyle.text.fontSize, imported.blocks[0].runs[0].fontSize)
  const { studio, document } = open(imported), before = copy(document.getModel())
  const range = { anchor: { blockId: 'leading', offset: 0 }, focus: { blockId: 'leading', offset: 'First styled paragraph'.length } }
  const result = document.executeCommand({ id: 'doc.command.clear-formatting' }, range)
  assert.equal(result.changed, true)
  const cleared = copy(document.getModel()), run = cleared.blocks[0].runs[0]
  assert.deepEqual([run.fontFamily, run.eastAsiaFontFamily, run.fontSize, run.bold, run.italic], ['Georgia', '宋体', format === 'doc' ? 16 : 18, false, true])
  assert.deepEqual(cleared.blocks[0].style, cleared.defaultStyle.paragraph)
  assert.deepEqual(cleared.blocks[1], before.blocks[1])
  assert.deepEqual(cleared.metadata.source.officeSource, originalSource)
  assert.equal(studio.editor.history.units[document.id].undo.length, 1)
  assert.equal(document.executeCommand({ id: 'doc.command.clear-formatting' }, result.selection).changed, false)
  assert.equal(studio.editor.history.units[document.id].undo.length, 1)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), before)
  assert.equal(document.redo(), true); assert.deepEqual(document.getModel(), cleared)
  const saved = document.save(), reopened = open(copy(saved))
  assert.deepEqual(reopened.document.getModel().metadata.source.officeSource, originalSource)
  assert.deepEqual(reopened.document.getModel().defaultStyle, cleared.defaultStyle)
  assert.equal(reopened.document.executeCommand({ id: 'doc.command.clear-formatting' }, result.selection).changed, false)
  assert.equal(reopened.studio.editor.history.units[saved.id].undo.length, 0)
  assert.equal(reopened.document.executeCommand({ id: 'doc.command.insert-text', text: 'new input' }, result.selection).changed, true)
  const inserted = reopened.document.getModel().blocks[0].runs.find(run => run.text === 'new input')
  assert.deepEqual([inserted.fontFamily, inserted.fontSize, inserted.bold, inserted.italic], ['Georgia', format === 'doc' ? 16 : 18, false, true])
  assert.deepEqual(reopened.document.getModel().metadata.source.officeSource, originalSource)
})

test('a DOCX with historical ODT resources resolves only its native Word defaults and keeps historical diagnostics', () => {
  const history = odtOrigin(); history.resourceErrors = ['An old ODF property is unresolved.']
  const officeSource = { version: 1, styles: { partName: 'word/styles.xml', xml: '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>' }, legacyOrigins: [history] }
  const model = normalizeDocument(envelope('docx', officeSource, 17))
  assert.equal(model.defaultStyle.text.fontSize, 17)
  assert.equal(model.metadata.source.officeSource.origin, undefined)
  assert.deepEqual(model.metadata.source.officeSource.legacyOrigins, [history])
  assert.deepEqual(model.metadata.source.officeSource.styles, officeSource.styles)
  const invalid = envelope('docx', { ...officeSource, resourceErrors: ['Current Word style failed.'] }, 17)
  assert.throws(() => normalizeDocument(invalid), TypeError)
})

test('old DOCX DTOs without history and absent-source DOC/ODT DTOs remain compatible without fabricating imported defaults', () => {
  for (const value of [undefined, null, { version: 1 }, { version: 1, legacyOrigins: [] }]) validateDocumentOfficeSource(value, 'docx')
  for (const format of ['doc', 'odt']) {
    const dto = envelope(format, undefined); delete dto.payload.data.defaultStyle
    const model = normalizeDocument(dto)
    assert.equal(model.defaultStyle, undefined)
    assert.equal(model.metadata.source.officeSource, undefined)
  }
})

test('source format, binary ranges, UTF-8 names and unsupported source fields reject without modifying the DTO', () => {
  const invalid = [
    [{ version: 1 }, 'doc'],
    [{ version: 1, origin: odtOrigin() }, 'docx'],
    [{ ...source('odt'), styles: { partName: 'styles.xml', xml: odtStyles } }, 'odt'],
    [{ version: 1, origin: { ...docOrigin(), synthesized: false } }, 'doc'],
    [{ version: 1, origin: docOrigin(0xFFFFFFFF, '00') }, 'doc'],
    [{ version: 1, origin: docOrigin(0.5, '00') }, 'doc'],
    [{ version: 1, origin: docOrigin(0, 'x0') }, 'doc'],
    [{ version: 1, origin: { ...odtOrigin(), extraUnretained: true } }, 'odt'],
    [{ version: 1, legacyOrigins: null }, 'docx'],
    [{ version: 1, unretainedExtension: {} }, 'docx'],
    [{ version: 1, origin: { format: 'odt', synthesized: false, resources: [{ kind: 'xml', partName: '中'.repeat(342), xml: 'x' }] } }, 'odt'],
    [{ version: 1, origin: { format: 'odt', synthesized: false, resources: [{ kind: 'xml', partName: 'styles.xml', xml: '\uD800' }] } }, 'odt'],
  ]
  for (const [value, format] of invalid) {
    const before = copy(value)
    assert.throws(() => validateDocumentOfficeSource(value, format))
    assert.deepEqual(value, before)
  }
  validateDocumentOfficeSource({ version: 1, origin: { format: 'odt', synthesized: false, resources: [{ kind: 'xml', partName: '😀'.repeat(256), xml: 'x' }] } }, 'odt')
})

test('combined budgets use the deduplicated history identity and preserve the original DTO text', () => {
  const origin = docOrigin(0, '00Ab'), duplicate = docOrigin(0, '00ab')
  const value = { version: 1, origin, legacyOrigins: [duplicate] }
  const before = copy(value)
  validateDocumentOfficeSource(value, 'doc')
  assert.deepEqual(value, before)
  const resources = start => Array.from({ length: 600 }, (_, index) => ({ kind: 'binary', streamName: '1Table', offset: start + index, dataHex: '00' }))
  assert.throws(() => validateDocumentOfficeSource({ version: 1, legacyOrigins: [
    { format: 'doc', synthesized: true, resources: resources(0) }, { format: 'doc', synthesized: true, resources: resources(600) },
  ] }, 'docx'), RangeError)
  assert.throws(() => validateDocumentOfficeSource({ version: 1, legacyOrigins: Array.from({ length: 33 }, () => docOrigin()) }, 'docx'))
  const hex = '00'.repeat(8 * 1024 * 1024)
  assert.throws(() => validateDocumentOfficeSource({ version: 1, legacyOrigins: [docOrigin(0, hex), docOrigin(1, hex + '00')] }, 'docx'), RangeError)
})

test('serialized JSON budget includes control-character escape expansion', () => {
  const value = { version: 1, origin: { format: 'odt', synthesized: false,
    resources: [{ kind: 'xml', partName: 'styles.xml', xml: '\u0001'.repeat(11 * 1024 * 1024) }] } }
  assert.throws(() => validateDocumentOfficeSource(value, 'odt'), RangeError)
})
