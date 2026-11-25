import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// S19 drafts only. No browser, Office or performance acceptance is implied.
let server, domain, modelUtils, controls, interaction, sourceOrder, codec, api
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-parent', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  interaction = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/content-control-interaction.ts')
  sourceOrder = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/content-control-source-order.ts')
  codec = await server.ssrLoadModule('/packages/office-studio/src/api/document-body-codec.ts')
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })

const copy = value => structuredClone(value)
const order = value => ({ univerBodySdtOrder: { version: 1, order: value } })
const selection = (start, end = start, target = { blockId: 'p' }) =>
  ({ anchor: { ...target, offset: start }, focus: { ...target, offset: end } })
function fixture() {
  const source = domain.createBlankDocument(); source.id = 'parent-order'; source.blocks[0].id = 'p'; source.blocks[0].runs[0].text = 'NAME'
  source.contentControls = [
    { id: 'm-outer', type: 'group', blockId: 'p', startOffset: 0, endOffset: 4, metadata: order(30) },
    { id: 'z-inner', type: 'richText', blockId: 'p', startOffset: 0, endOffset: 4, metadata: order(10) },
    { id: 'a-middle', type: 'group', blockId: 'p', startOffset: 0, endOffset: 4, metadata: order(20) },
  ]
  return source
}
const parent = (source, id = 'z-inner', host, page) => interaction.documentContentControlParentSelection(source, id, host, page)
function rawBody() {
  return { dataStream: 'NAME\r\n', paragraphs: [{ startIndex: 4 }], customRanges: [
    { rangeId: 'link', rangeType: 0, startIndex: 0, endIndex: 3, properties: { url: 'https://example.test' } },
    ...['z-inner', 'a-middle', 'm-outer'].map((id, index) => ({ rangeId: id, rangeType: 2, startIndex: 0, endIndex: 3,
      properties: { kind: index ? 'group' : 'richText', placement: 'inline' } })),
  ] }
}

test('three equal imported ranges navigate by original source order despite local ID and array sorting', () => {
  const source = fixture(), before = copy(source)
  assert.equal(controls.selectedDocumentContentControlForSelection(source, selection(2)).id, 'z-inner')
  assert.deepEqual(parent(source), { controlId: 'a-middle', selection: selection(0, 4) })
  assert.deepEqual(parent(source, 'a-middle'), { controlId: 'm-outer', selection: selection(0, 4) })
  assert.equal(parent(source, 'm-outer'), undefined)
  source.contentControls.reverse()
  assert.equal(parent(source).controlId, 'a-middle')
  source.contentControls.reverse()
  assert.deepEqual(source, before)
})

test('explicit parent identity overrides equal-range provenance and an invalid explicit parent cannot fall back', () => {
  const source = fixture(), child = source.contentControls.find(control => control.id === 'z-inner')
  child.parentId = 'm-outer'
  assert.equal(parent(source).controlId, 'm-outer')
  source.contentControls.find(control => control.id === 'm-outer').metadata = order(0)
  assert.equal(parent(source).controlId, 'm-outer')
  child.parentId = 'missing'
  assert.equal(parent(source), undefined)
})

test('missing, invalid and duplicated provenance never manufacture an equal-range parent', () => {
  const source = fixture()
  source.contentControls.forEach(control => { delete control.metadata })
  assert.equal(parent(source), undefined)
  source.contentControls.forEach(control => { control.metadata = order(1) })
  assert.equal(sourceOrder.documentContentControlSourceOrders(source.contentControls).size, 0)
  assert.equal(parent(source), undefined)
  for (const value of [null, [], {}, { version: 2, order: 1 }, { version: 1, order: -1 },
    { version: 1, order: 1.5 }, { version: 1, order: '1' }, { version: 1, order: Number.MAX_SAFE_INTEGER + 1 }]) {
    assert.equal(sourceOrder.readDocumentContentControlSourceOrder(value), undefined)
  }
  assert.equal(sourceOrder.readDocumentContentControlSourceOrder({ version: 1, order: 0 }), 0)
  source.contentControls[0].startOffset = 0; source.contentControls[0].endOffset = 4
  source.contentControls[1].startOffset = 1; source.contentControls[1].endOffset = 3
  source.contentControls[2].startOffset = 1; source.contentControls[2].endOffset = 3
  assert.equal(parent(source).controlId, 'm-outer', 'strict containment is still a valid parent')
})

test('strict native note-boundary containment chooses the nearest parent even when visible offsets have equal widths', () => {
  const source = fixture()
  const note = id => ({ ...modelUtils.createRun(''), id, metadata: { note: { type: 'footnote', noteId: id } } })
  source.blocks[0].runs = [modelUtils.createRun('L'), note('n1'), note('n2'), modelUtils.createRun('R')]
  source.contentControls.forEach(control => { control.startOffset = 1; control.endOffset = 1; delete control.metadata })
  Object.assign(source.contentControls[0], { startNoteReferenceCount: 0, endNoteReferenceCount: 2 })
  Object.assign(source.contentControls[1], { startNoteReferenceCount: 1, endNoteReferenceCount: 1 })
  Object.assign(source.contentControls[2], { startNoteReferenceCount: 0, endNoteReferenceCount: 1 })
  assert.equal(parent(source).controlId, 'a-middle')
  assert.equal(parent(source, 'a-middle').controlId, 'm-outer')
})

test('body provenance stays independent of header controls whose source indices happen to match', () => {
  const source = fixture(), header = modelUtils.createParagraph('NAME'); header.id = 'header-p'
  source.sectionHeaderFooters = { first: { header: 'NAME', footer: '', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '',
    differentFirstPage: false, differentOddEven: false, linkToPrevious: false,
    storyContent: { header: { blocks: [header], runs: copy(header.runs), paragraphs: [{ startOffset: 0, endOffset: 4, style: {} }] } } } }
  source.contentControls.push(...source.contentControls.map(control => ({ ...copy(control), id: 'header-' + control.id, blockId: header.id })))
  const host = { kind: 'story', sectionId: 'first', story: 'header' }
  assert.equal(parent(source).controlId, 'a-middle')
  assert.equal(parent(source, 'header-z-inner', host).controlId, 'header-a-middle')
  assert.equal(controls.selectedDocumentContentControlForSelection(source, selection(2)).id, 'z-inner')
})

test('header and note table parents preserve their host coordinates and requested page affinity', () => {
  for (const host of [{ kind: 'story', sectionId: 'first', story: 'header' }, { kind: 'note', noteKind: 'footnote', noteId: 'note' }]) {
    const source = fixture(), block = source.blocks[0], table = modelUtils.createParagraph()
    table.id = 'host-table'; table.type = 'table'; table.runs = []
    table.metadata.table = { rows: [{ id: 'host-row', cells: [{ id: 'host-cell', blocks: [block], rowSpan: 1, colSpan: 1, style: {} }] }] }
    const content = { blocks: [table], runs: copy(block.runs), paragraphs: [{ startOffset: 0, endOffset: 4, style: {} }] }
    source.blocks = [modelUtils.createParagraph('Body')]
    if (host.kind === 'story') source.sectionHeaderFooters = { first: { header: 'NAME', footer: '', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '',
      differentFirstPage: false, differentOddEven: false, linkToPrevious: false, storyContent: { header: content } } }
    else source.footnotes = { note: { id: 'note', text: 'NAME', content } }
    const before = copy(source), target = { blockId: table.id, row: 0, col: 0, cellBlockId: block.id, pageAffinity: 5 }
    assert.deepEqual(parent(source, 'z-inner', host, 5), { controlId: 'a-middle', selection: selection(0, 4, target) })
    assert.deepEqual(source, before)
    for (const page of [-1, 0.5, Number.MAX_SAFE_INTEGER + 1]) assert.equal(parent(source, 'z-inner', host, page), undefined)
  }
})

test('table-row parent selections carry page affinity on both text points and the structural view', () => {
  const source = fixture(), paragraphs = [source.blocks[0], modelUtils.createParagraph('NEXT')]; paragraphs[1].id = 'p-next'
  const table = modelUtils.createParagraph(); table.id = 'row-table'; table.type = 'table'; table.runs = []
  table.metadata.table = { rows: paragraphs.map((block, index) => ({ id: 'r' + index,
    cells: [{ id: 'c' + index, blocks: [block], rowSpan: 1, colSpan: 1, style: {} }] })) }
  source.blocks = [table]
  source.contentControls = [
    { id: 'row-parent', type: 'group', blockId: 'p', startOffset: 0, endOffset: 4, tableRowRange: { tableId: table.id, rowIds: ['r0', 'r1'] } },
    { id: 'row-child', type: 'group', blockId: 'p-next', startOffset: 0, endOffset: 4, parentId: 'row-parent',
      tableRowRange: { tableId: table.id, rowIds: ['r1'] } },
  ]
  const result = parent(source, 'row-child', undefined, 7)
  assert.equal(result.controlId, 'row-parent'); assert.equal(result.selection.anchor.pageAffinity, 7)
  assert.equal(result.selection.focus.pageAffinity, 7); assert.equal(result.selection.tableSelection.pageAffinity, 7)
})

test('a value focus token cannot match a properties request at the same logical range and input generation', () => {
  const source = fixture(), selected = selection(0, 4)
  const input = { valid: true, modelRevision: 4, generation: 7, marks: {}, location: { kind: 'body', selection: selected } }
  const request = { source, controlId: 'z-inner', selection: selected, inputSession: input, mode: 'value' }
  const token = interaction.captureDocumentContentControlFocusToken(request, input)
  assert.ok(token); assert.equal(interaction.documentContentControlFocusTokenMatchesRequest(token, { ...request }), true)
  assert.equal(interaction.documentContentControlFocusTokenMatchesRequest(token, { ...request, mode: 'properties' }), false)
  assert.equal(interaction.documentContentControlFocusTokenMatchesRequest(token, { ...request, mode: undefined }), true)
})

test('public decoding records all custom-range indices and rejects malformed persisted ordering', () => {
  const decoded = codec.decodeDocumentBody(rawBody())
  assert.deepEqual(decoded.contentControls.map(control => [control.id, sourceOrder.documentContentControlSourceOrder(control)]),
    [['z-inner', 1], ['a-middle', 2], ['m-outer', 3]])
  for (const value of [null, [], {}, { version: 2, order: 1 }, { version: 1, order: -1 }, { version: 1, order: 1.5 }, { version: 1, order: '1' }]) {
    const body = rawBody(); body.customRanges[1].officeStudio = { version: 1, metadata: { univerBodySdtOrder: value } }
    assert.throws(() => codec.decodeDocumentBody(body), /来源顺序/)
  }
})

test('public facade save and reopen preserve implicit parent navigation after customRanges are reordered', () => {
  const decoded = codec.decodeDocumentBody(rawBody()), source = domain.createBlankDocument()
  source.id = 'public-parent-order'; source.blocks = decoded.content.blocks; source.contentControls = decoded.contentControls
  const studio = new api.OfficeStudio({ document: source }); studio.activateDocument(source.id)
  let reopened
  try {
    const document = new api.FDocument(studio, source.id), saved = copy(document.save())
    saved.body.customRanges.reverse()
    reopened = new api.OfficeStudio({ document: saved }); reopened.activateDocument(source.id)
    const model = new api.FDocument(reopened, source.id).getModel()
    assert.deepEqual(model.contentControls.map(control => [control.id, sourceOrder.documentContentControlSourceOrder(control)]).sort(),
      [['a-middle', 2], ['m-outer', 3], ['z-inner', 1]])
    assert.equal(parent(model).controlId, 'a-middle'); assert.equal(parent(model, 'a-middle').controlId, 'm-outer')
    assert.equal(parent(model, 'm-outer'), undefined)
    assert.equal(studio.editor.history.units[source.id]?.undo.length ?? 0, 0)
    assert.equal(reopened.editor.history.units[source.id]?.undo.length ?? 0, 0)
  } finally { studio.dispose(); reopened?.dispose() }
})
