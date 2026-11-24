import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-11/DOC-22 drafts. XML/XPath parsing itself is covered by the separate real-browser drafts.
// These drafts are not executed until the user's unified validation phase.
let server, boundary, graph, models, permissions
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-custom-xml-contract', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  boundary = await server.ssrLoadModule('/packages/office-studio/src/domain/document-custom-xml.ts')
  graph = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-custom-xml-validation.ts')
  models = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  permissions = await server.ssrLoadModule('/packages/office-studio/src/features/document/permissions/document-permissions.ts')
})
afterAll(async () => { await server?.close() })

const source = (partName = 'customXml/item.xml', patch = {}) => ({ partName, xml: '<root><value>old</value></root>', storeItemID: '{STORE}', ...patch })
const propType = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'
const relType = 'application/vnd.openxmlformats-package.relationships+xml'
const propRelationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'
const copy = value => JSON.parse(JSON.stringify(value))
const resource = (partName, contentType = 'application/octet-stream', dataBase64 = 'AA==') => ({ partName, contentType, dataBase64 })
const relation = (target, type = 'urn:office-studio:retained-resource', id = 'resource', external = false) => ({ id, target, type, external })

// Injected platform DTOs isolate graph planning; they do not simulate XPath or claim XML-parser coverage.
function inspector({ relationships = new Map(), properties = new Map(), decoded = new Map() } = {}) {
  const validated = []
  return { validated, validate(xml) { validated.push(xml) }, propertiesStoreItemID(xml) {
    if (!properties.has(xml)) throw new Error('The fixture did not define a properties XML document')
    return properties.get(xml)
  }, relationships(xml) {
    if (!relationships.has(xml)) throw new Error('The fixture did not define an OPC relationships document')
    return relationships.get(xml)
  }, decodeXmlResource(base64) {
    if (!decoded.has(base64)) throw new Error('The fixture did not define an encoded XML resource')
    return decoded.get(base64)
  } }
}

test('OPC names normalize spaces, UTF-8 escapes and case identity without changing source XML', () => {
  const original = source('customXml/名 前%20😀.xml')
  const normalized = boundary.normalizeDocumentCustomXmlParts([original])
  assert.equal(normalized[0].partName, 'customXml/%E5%90%8D%20%E5%89%8D%20%F0%9F%98%80.xml')
  assert.equal(boundary.canonicalDocumentCustomXmlPartName(normalized[0].partName), normalized[0].partName)
  assert.equal(boundary.documentCustomXmlPartKey('CustomXML/%6Eame.xml'), boundary.documentCustomXmlPartKey('customxml/name.XML'))
  assert.equal(normalized[0].xml, original.xml)
  assert.equal(original.partName, 'customXml/名 前%20😀.xml')
  assert.throws(() => boundary.normalizeDocumentCustomXmlParts([source('customXml/%6Eame.xml'), source('CUSTOMXML/NAME.XML', { storeItemID: '{OTHER}' })]))
})

test('OPC references resolve actual owner-relative URIs and reject encoded separators or traversal', () => {
  assert.equal(boundary.resolveDocumentCustomXmlTarget('customXml/nested/item.xml', '../media/图 1.bin'), 'customXml/media/%E5%9B%BE%201.bin')
  assert.equal(boundary.resolveDocumentCustomXmlTarget('word/document.xml', '/customXml/item%25.xml'), 'customXml/item%25.xml')
  assert.equal(boundary.documentCustomXmlRelationshipPart('customXml/a%20b.xml'), 'customXml/_rels/a%20b.xml.rels')
  assert.equal(boundary.documentCustomXmlRelationshipOwner('CUSTOMXML/_RELS/A%20B.XML.RELS'), 'CUSTOMXML/A%20B.XML')
  assert.equal(boundary.documentCustomXmlRelationshipOwner('_rels/.rels'), undefined)
  for (const target of ['../../escape.xml', 'item%2fescape.xml', 'item%5Cescape.xml', '%2e%2e/escape.xml', 'item%.xml', 'item%FF.xml', '//outside/item.xml', 'x//y.xml', 'item.xml?x=1']) {
    assert.throws(() => boundary.resolveDocumentCustomXmlTarget('customXml/item.xml', target), target)
  }
})

test('source arrays are dense and bounded and normalization isolates caller-owned objects', () => {
  assert.deepEqual(boundary.normalizeDocumentCustomXmlParts([]), [])
  assert.deepEqual(boundary.normalizeDocumentCustomXmlParts(undefined), [])
  const sparse = [source()]; sparse.length = 2
  for (const value of [null, {}, sparse, [source(undefined, { unknown: 1 })], [source('customXml/item.rels')], [source('[Content_Types].xml')]]) {
    assert.throws(() => boundary.normalizeDocumentCustomXmlParts(value))
  }
  const input = [source(undefined, { relatedParts: [resource('customXml/blob.bin')] })]
  const original = copy(input)
  const normalized = boundary.normalizeDocumentCustomXmlParts(input)
  normalized[0].relatedParts[0].partName = 'changed.bin'
  normalized[0].xml = '<changed/>'
  assert.deepEqual(input, original)
})

test('canonical base64 accepts empty resources and rejects malformed padding or nonzero trailing bits', () => {
  for (const [value, length] of [['', 0], ['Zg==', 1], ['AAA=', 2], ['AAAA', 3]]) assert.equal(boundary.documentCustomXmlResourceByteLength(value), length)
  for (const value of ['Zh==', 'AA=A', 'AB==', 'AAB=', 'Zg==\n', '====']) assert.throws(() => boundary.documentCustomXmlResourceByteLength(value), value)
})

test('XML output budgets include declaration correction and actual UTF-16/BOM bytes', () => {
  const xml = '<?xml version="1.0" encoding="UTF8"?><root>😀</root>'
  assert.equal(boundary.documentCustomXmlEncodedBytes(xml, 'utf8'), boundary.documentCustomXmlUtf8Bytes(xml) + 1)
  assert.equal(boundary.documentCustomXmlEncodedBytes(xml, 'utf16le', true), (xml.length + 4) * 2 + 2)
  const maximum = boundary.DOCUMENT_CUSTOM_XML_MAX_PART_BYTES
  const exact = '<r>' + 'x'.repeat(maximum / 2 - 7) + '</r>'
  assert.equal(exact.length * 2, maximum)
  boundary.normalizeDocumentCustomXmlParts([source(undefined, { xml: exact, encoding: 'utf16le' })])
  assert.throws(() => boundary.normalizeDocumentCustomXmlParts([source(undefined, { xml: exact, encoding: 'utf16le', byteOrderMark: true })]), error => error.code === 'DOC_CUSTOM_XML_RESOURCE_LIMIT')
  assert.throws(() => boundary.normalizeDocumentCustomXmlParts([source(undefined, { xml: '<r>\uD800</r>' })]), error => error.code === 'DOC_CUSTOM_XML_INVALID_UNICODE')
})

test('content types preserve valid spelling and reject malformed MIME fields', () => {
  const normalized = boundary.normalizeDocumentCustomXmlParts([source(undefined, { contentType: 'Application/XML' })])
  assert.equal(normalized[0].contentType, 'Application/XML')
  for (const value of ['unknown+xml', 'application/xml; charset=utf8', 'application/xml\n', 'application/xml\r\n', 'application/xmK', 'application//xml']) assert.equal(boundary.isDocumentCustomXmlContentType(value), false)
  assert.throws(() => boundary.normalizeDocumentCustomXmlParts([source(undefined, { contentType: 'application/octet-stream' })]))
})

test('all XML stores are inspected and a shared, cyclic resource graph is validated once by package identity', () => {
  const a = '<Relationships-a/>', b = '<Relationships-b/>', attachment = '<Relationships-attachment/>'
  const runtime = inspector({ relationships: new Map([[a, [relation('shared.bin'), relation('https://example.test/schema', 'urn:external', 'ext', true)]],
    [b, [relation('SHARED.bin')]], [attachment, [relation('item.xml')]]]), decoded: new Map([['YQ==', attachment]]) })
  const parts = [source('customXml/item.xml', { relationshipsXml: a, relatedParts: [resource('customXml/shared.bin'), resource('customXml/_rels/shared.bin.rels', relType, 'YQ==')] }),
    source('customXml/other.xml', { storeItemID: '{OTHER}', relationshipsXml: b })]
  const result = graph.validateDocumentCustomXmlParts(parts, runtime)
  assert.equal(result.length, 2)
  assert.deepEqual(runtime.validated, parts.map(part => part.xml))
  assert.equal(result[0].relationshipsXml, a)
  assert.equal(result[1].relatedParts, undefined)
})

test('missing graph targets or owners and conflicting relationship representations fail before publication', () => {
  const rels = '<Relationships/>', runtime = inspector({ relationships: new Map([[rels, [relation('missing.bin')]]]), decoded: new Map([['YQ==', rels]]) })
  assert.throws(() => graph.validateDocumentCustomXmlParts([source(undefined, { relationshipsXml: rels })], runtime), error => error.code === 'DOC_CUSTOM_XML_MISSING_PART')
  assert.throws(() => graph.validateDocumentCustomXmlParts([source(undefined, { relatedParts: [resource('customXml/_rels/missing.bin.rels', relType, 'YQ==')] })], runtime), error => error.code === 'DOC_CUSTOM_XML_MISSING_PART')
  const other = '<Relationships-other/>', conflict = inspector({ relationships: new Map([[rels, []], [other, []]]), decoded: new Map([['YQ==', other]]) })
  assert.throws(() => graph.validateDocumentCustomXmlParts([source(undefined, { relationshipsXml: rels,
    relatedParts: [resource('customXml/_rels/item.xml.rels', relType, 'YQ==')] })], conflict), error => error.code === 'DOC_CUSTOM_XML_RELATIONSHIP_CONFLICT')
})

test('properties identities are checked against the actual property XML, including inferred and duplicate stores', () => {
  const rels = '<Relationships-props/>', props = '<datastoreItem/>', runtime = inspector({ relationships: new Map([[rels, [relation('props.xml', propRelationship)]]]),
    properties: new Map([[props, '{STORE}']]), decoded: new Map([['YQ==', props]]) })
  const inferred = graph.validateDocumentCustomXmlParts([source(undefined, { storeItemID: undefined, relationshipsXml: rels,
    relatedParts: [resource('customXml/props.xml', propType, 'YQ==')] })], runtime)
  assert.equal(inferred[0].storeItemID, '{STORE}')
  assert.throws(() => graph.validateDocumentCustomXmlParts([source(undefined, { storeItemID: '{WRONG}', propertiesPartName: 'customXml/props.xml', propertiesXml: props })], runtime), error => error.code === 'DOC_CUSTOM_XML_INVALID_PROPERTIES')
  assert.throws(() => graph.validateDocumentCustomXmlParts([source(undefined, { storeItemID: undefined, relationshipsXml: rels,
    relatedParts: [resource('customXml/props.xml', propType, 'YQ==')] }), source('customXml/other.xml')], runtime), error => error.code === 'DOC_CUSTOM_XML_INVALID_PROPERTIES')
})

test('XML-only or binding-property edits target protected paragraphs even when their rendered text is unchanged', () => {
  const before = models.createBlankDocument(); before.id = 'xml-permissions'
  before.blocks[0].id = 'protected-p'
  before.customXmlParts = [source()]
  before.contentControls = [{ id: 'bound', type: 'richText', placement: 'inline', blockId: 'protected-p', startOffset: 0, endOffset: 0,
    dataBinding: { xpath: '/root/value', storeItemID: '{STORE}', cachedResult: 'old' } }]
  before.permissionRules = [{ id: 'deny-p', scope: 'paragraph', targetId: 'protected-p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  const after = copy(before); after.customXmlParts[0].xml = '<root><value>new</value></root>'
  const targets = permissions.permissionTargetsForChangedBlocks(before, after)
  assert(targets.some(target => target.scope === 'paragraph' && target.id === 'protected-p'))
  assert.throws(() => permissions.assertDocumentPermission(before.permissionRules, { documentId: before.id, userId: 'editor' }, 'edit', targets), error => error.code === 'DOCUMENT_PERMISSION_DENIED')
  const propertyEdit = copy(before); propertyEdit.contentControls[0].dataBinding.xpath = '/root/other'
  assert(permissions.permissionTargetsForChangedBlocks(before, propertyEdit).some(target => target.id === 'protected-p'))
  assert.deepEqual(after.blocks, before.blocks)
})
