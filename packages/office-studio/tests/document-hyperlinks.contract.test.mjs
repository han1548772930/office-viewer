import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-08 drafts; execution belongs to the unified validation phase.
let server, OfficeStudio, FDocument, domain, modelUtils, hyperlinks, codec, controls, hosts, publicBody, layout, scene, sections, segments, fields
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-hyperlinks', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, modelUtils, hyperlinks, codec, controls, hosts, publicBody, layout, scene, sections, segments, fields] = await Promise.all([
    'domain/model', 'features/document/model/document-model', 'features/document/services/document-hyperlinks',
    'api/document-body-codec', 'features/document/services/content-controls', 'features/document/services/document-block-host',
    'api/document-public-body', 'features/document/layout/document-layout', 'features/document/rendering/document-scene',
    'features/document/services/section-header-footer', 'api/document-segments', 'features/document/model/document-text-fields',
  ].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const range = (blockId, start = 0, end = start) => ({ anchor: { blockId, offset: start }, focus: { blockId, offset: end } })
const paragraph = (id, text) => ({ ...modelUtils.createParagraph(text), id })
const history = (studio, id) => studio.editor.history.units[id]?.undo.length ?? 0
const links = document => hyperlinks.documentHyperlinks(document.getModel())
const authoredRuns = model => model.blocks.flatMap(block => block.runs.map(run => ({ id: run.id, text: run.text, bold: run.bold,
  italic: run.italic, underline: run.underline, color: run.color, fontSize: run.fontSize, fontFamily: run.fontFamily })))
function open(model = { ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: [paragraph('p', 'Alpha'), paragraph('tail', 'Tail')] }) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}
function publicModel(body) {
  const decoded = codec.decodeDocumentBody(body)
  return { ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: decoded.content.blocks,
    metadata: decoded.content.metadata, contentControls: decoded.contentControls }
}
function linkedBody(properties = { url: 'https://old.example' }) {
  return { dataStream: 'AB\rCD\r\n', paragraphs: [{ startIndex: 2, paragraphId: 'p' }, { startIndex: 5, paragraphId: 'tail' }],
    sectionBreaks: [{ startIndex: 6, sectionId: 'first' }],
    textRuns: [{ st: 0, ed: 1, ts: { bl: 1, cl: { rgb: '#883311' } } }, { st: 1, ed: 2, ts: { it: 1 } },
      { st: 3, ed: 5, ts: { fs: 18 } }],
    customRanges: [{ rangeType: 0, rangeId: 'imported-link', startIndex: 0, endIndex: 4, properties },
      { rangeType: 1, rangeId: 'nested-field', startIndex: 0, endIndex: 0, properties: { fieldType: 'AUTHOR', instruction: 'AUTHOR', cachedResult: 'A' } }] }
}

test('one command boundary normalizes bare websites, mailboxes and internal targets and rejects unsupported protocols', () => {
  assert.equal(hyperlinks.normalizeDocumentHyperlinkTarget(' example.org/path '), 'https://example.org/path')
  assert.equal(hyperlinks.normalizeDocumentHyperlinkTarget('person@example.org'), 'mailto:person@example.org')
  assert.equal(hyperlinks.normalizeDocumentHyperlinkTarget('#heading-id'), '#heading-id')
  assert.equal(hyperlinks.normalizeDocumentHyperlinkTarget('https://example.org'), 'https://example.org')
  for (const value of ['', '#', 'javascript:alert(1)', 'https://', 'abc\0def']) assert.throws(() => hyperlinks.normalizeDocumentHyperlinkTarget(value))
})

test('the registered add command links the real half-open selection without changing author text formatting', () => {
  const model = { ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: [paragraph('p', 'A😀B'), paragraph('tail', 'Tail')] }
  Object.assign(model.blocks[0].runs[0], { color: '#883311', underline: false, bold: true, metadata: { authorHint: 'retained' } })
  const { studio, document } = open(model), events = []
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') events.push(event.id) })
  const selected = range('p', 1, 3)
  assert.equal(document.executeCommand({ id: 'docs.command.add-hyper-link', payload: 'example.org' }, selected).changed, true)
  const link = links(document)[0]
  assert.equal(link.label, '😀'); assert.equal(link.target, 'https://example.org'); assert.deepEqual(link.selection, selected)
  assert.equal(document.getText(), 'A😀B\nTail')
  for (const run of document.getModel().blocks[0].runs) {
    assert.equal(run.color, '#883311'); assert.equal(run.underline, false); assert.equal(run.bold, true)
    assert.equal(run.metadata.authorHint, 'retained')
  }
  assert.deepEqual(events, ['docs.command.add-hyper-link']); assert.equal(history(studio, document.id), 1)
  const saved = document.save(), reopened = open(saved).document
  assert.equal(links(reopened)[0].id, link.id); assert.equal(links(reopened)[0].label, '😀')
})

test('collapsed add, paragraph-only ranges and surrogate interiors reject without content or history changes', () => {
  const { studio, document } = open({ ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: [paragraph('p', 'A😀B'), paragraph('tail', 'Tail')] })
  const before = copy(document.getModel())
  assert.equal(hyperlinks.documentHyperlinkSelectionState(document.getModel(), range('p', 0)).available, false)
  for (const selected of [range('p', 0), range('p', 1, 2), { anchor: { blockId: 'p', offset: 4 }, focus: { blockId: 'tail', offset: 0 } }]) {
    assert.throws(() => document.executeCommand({ id: 'docs.command.add-hyper-link', payload: 'example.org' }, selected))
    assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id), 0)
  }
})

test('unchanged labels preserve mixed run identities, nested fields and paragraph metadata when only the target changes', () => {
  const { studio, document } = open(publicModel(linkedBody({ url: 'https://old.example', textStyleMode: 'text', textColorMode: 'text', authorHint: 7 })))
  const before = copy(document.getModel()), original = authoredRuns(before), link = links(document)[0]
  assert.equal(link.label, 'AB\nCD')
  assert.equal(document.executeCommand({ id: 'docs.command.update-hyper-link', linkId: link.id, payload: 'new.example', label: link.label }, range('p', 1)).changed, true)
  assert.deepEqual(authoredRuns(document.getModel()), original)
  assert.deepEqual(document.getModel().blocks.map(block => block.metadata), before.blocks.map(block => block.metadata))
  const output = document.getBody().customRanges.find(range => range.rangeId === link.id)
  assert.equal(output.properties.url, 'https://new.example'); assert.equal(output.properties.textStyleMode, 'text')
  assert.equal(output.properties.textColorMode, 'text'); assert.equal(output.properties.authorHint, 7)
  assert.equal(history(studio, document.id), 1)
  assert.equal(document.undo(), true); assert.deepEqual(authoredRuns(document.getModel()), original)
  assert.equal(links(document)[0].target, 'https://old.example'); assert.equal(document.redo(), true)
})

test('bookmark and heading targets import as editable links and target changes write exactly one destination property', () => {
  for (const properties of [{ bookmarkId: 'bookmark' }, { headingId: 'tail' }]) {
    const { document } = open(publicModel(linkedBody(properties))), link = links(document)[0]
    assert.equal(link.target, '#' + (properties.bookmarkId ?? properties.headingId))
    assert.equal(document.getBody().customRanges.find(range => range.rangeId === link.id).properties.url, undefined)
    document.executeCommand({ id: 'docs.command.update-hyper-link', linkId: link.id, payload: 'external.example', label: link.label }, range('p', 1))
    const external = document.getBody().customRanges.find(range => range.rangeId === link.id).properties
    assert.equal(external.url, 'https://external.example'); assert.equal(external.bookmarkId, undefined); assert.equal(external.headingId, undefined)
    document.executeCommand({ id: 'docs.command.update-hyper-link', linkId: link.id, payload: '#tail', label: link.label }, range('p', 1))
    const internal = document.getBody().customRanges.find(range => range.rangeId === link.id).properties
    assert.equal(internal.headingId, 'tail'); assert.equal(internal.url, undefined); assert.equal(internal.bookmarkId, undefined)
  }
})

test('delete removes the full stable link entity and old source metadata while preserving text, styles, undo and body reopen', () => {
  const { studio, document } = open(publicModel(linkedBody())), original = authoredRuns(document.getModel()), link = links(document)[0]
  assert.equal(document.executeCommand({ id: 'docs.command.delete-hyper-link', linkId: link.id }, range('p', 1)).changed, true)
  assert.deepEqual(authoredRuns(document.getModel()), original); assert.equal(links(document).length, 0)
  for (const run of document.getModel().blocks.flatMap(block => block.runs)) {
    assert.equal(run.href, undefined); assert.equal(run.metadata?.link, undefined); assert.equal(run.metadata?.univerBodyHyperlinkSource, undefined)
  }
  assert.equal(document.getBody().customRanges.some(range => range.rangeType === 0), false)
  const reopened = open(document.save()).document
  assert.equal(links(reopened).length, 0); assert.equal(reopened.getText(), 'AB\nCD')
  assert.equal(history(studio, document.id), 1); assert.equal(document.undo(), true); assert.equal(links(document)[0].id, link.id)
  assert.equal(document.redo(), true); assert.equal(links(document).length, 0)
})

test('overlapping additions retain separate prefix and suffix links through save and do not merge across unlinked text', () => {
  const { document } = open({ ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: [paragraph('p', 'abcde')] })
  document.executeCommand({ id: 'docs.command.add-hyper-link', payload: 'old.example' }, range('p', 0, 5))
  document.executeCommand({ id: 'docs.command.add-hyper-link', payload: 'new.example' }, range('p', 1, 3))
  const split = links(document)
  assert.deepEqual(split.map(link => link.label), ['a', 'bc', 'de']); assert.equal(new Set(split.map(link => link.id)).size, 3)
  const reopened = open(document.save()).document
  assert.deepEqual(links(reopened).map(link => [link.id, link.label]), split.map(link => [link.id, link.label]))
  reopened.executeCommand({ id: 'docs.command.delete-hyper-link', linkId: links(reopened)[1].id }, range('p', 2))
  const saved = reopened.getBody().customRanges.filter(range => range.rangeType === 0)
  assert.equal(saved.length, 2); assert.deepEqual(links(open(reopened.save()).document).map(link => link.label), ['a', 'de'])
})

test('capability and execution check the complete entity, including a locked or denied paragraph outside the caret', () => {
  for (const locked of [false, true]) {
    const model = publicModel(linkedBody())
    if (locked) controls.createDocumentContentControl(model, { id: 'lock', type: 'richText', blockId: 'tail', startOffset: 0, endOffset: 2, lock: 'contentLocked' })
    if (!locked) model.permissionRules = [{ id: 'deny-tail', scope: 'paragraph', targetId: 'tail', userIds: ['editor'], action: 'edit', effect: 'deny' }]
    const { studio, document } = open(model), before = copy(document.getModel()), link = links(document)[0], selected = range('p', 1)
    assert.equal(hyperlinks.documentHyperlinkSelectionState(document.getModel(), selected, { documentId: document.id, userId: 'editor' }).available, false)
    for (const command of [{ id: 'docs.command.update-hyper-link', linkId: link.id, payload: 'new.example', label: link.label },
      { id: 'docs.command.update-hyper-link', linkId: link.id, payload: 'new.example', label: 'Replacement' },
      { id: 'docs.command.delete-hyper-link', linkId: link.id }]) {
      assert.throws(() => document.executeCommand(command, selected))
      assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id), 0)
    }
  }
})

test('editing a label resolves its actual source control and mirrors bound peers in one history transaction', () => {
  const model = { ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: [paragraph('p', 'Name'), paragraph('peer', 'Name'), paragraph('tail', 'Tail')], contentControls: [] }
  for (const [id, blockId] of [['origin', 'p'], ['bound-peer', 'peer']]) controls.createDocumentContentControl(model, {
    id, type: 'richText', blockId, startOffset: 0, endOffset: 4, value: 'Name',
    dataBinding: { storeItemID: '{STORE}', xpath: '/person/name', prefixMappings: '', cachedResult: 'Name', dirty: false },
  })
  const { studio, document } = open(model)
  document.executeCommand({ id: 'docs.command.add-hyper-link', payload: 'example.org' }, range('p', 0, 4))
  const link = links(document)[0], before = copy(document.getModel()), count = history(studio, document.id)
  document.executeCommand({ id: 'docs.command.update-hyper-link', linkId: link.id, payload: link.target, label: '李😀', activeContentControlId: 'origin' }, range('tail', 1))
  assert.equal(document.getModel().blocks[0].runs.map(run => run.text).join(''), '李😀')
  assert.equal(document.getModel().blocks[1].runs.map(run => run.text).join(''), '李😀')
  for (const control of document.getModel().contentControls) {
    assert.equal(control.value, '李😀'); assert.equal(control.dataBinding.cachedResult, '李😀'); assert.equal(control.dataBinding.dirty, true)
  }
  assert.equal(history(studio, document.id), count + 1)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel().contentControls, before.contentControls)
  assert.equal(document.getText(), 'Name\nName\nTail'); assert.equal(document.redo(), true)
  const reopened = open(document.save()).document
  assert.equal(links(reopened).find(candidate => candidate.id === link.id).label, '李😀')
  assert.deepEqual(reopened.getModel().contentControls.map(control => control.value), ['李😀', '李😀'])
})

test('a story link classifies fresh body headings against the full document and navigates to body without a content edit', () => {
  const model = { ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: [paragraph('p', 'Body'), { ...paragraph('tail', 'Heading'), type: 'heading', level: 1 }] }
  const header = paragraph('header-p', 'Link')
  model.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: 'Link',
    storyContent: { header: { blocks: [header], runs: [], paragraphs: [] } } } }
  const { studio, document } = open(model), bodyBefore = copy(document.getModel().blocks)
  const host = { kind: 'story', sectionId: 'first', story: 'header' }
  document.executeCommand({ id: 'docs.command.add-hyper-link', payload: '#tail', host }, range('header-p', 0, 4))
  const scoped = hosts.documentModelForBlockHost(document.getModel(), host), link = hyperlinks.documentHyperlinks(scoped)[0]
  assert.equal(scoped.blocks[0].runs[0].metadata.link.headingId, 'tail')
  assert.deepEqual(document.getModel().blocks, bodyBefore)
  const beforeNavigation = copy(document.getModel()), count = history(studio, document.id)
  assert.deepEqual(publicBody.documentHyperlinkTargetSelection(document.getModel(), link.target), range('tail', 0))
  assert.deepEqual(document.getModel(), beforeNavigation); assert.equal(history(studio, document.id), count)
  assert.equal(document.undo(), true); assert.equal(hyperlinks.documentHyperlinks(hosts.documentModelForBlockHost(document.getModel(), host)).length, 0)
  assert.equal(document.redo(), true)
  const reopened = open(document.save()).document
  assert.equal(hyperlinks.documentHyperlinks(hosts.documentModelForBlockHost(reopened.getModel(), host))[0].id, link.id)
})

test('public bookmark and heading aliases resolve across paragraph and range markers with bookmark precedence and no mutation', () => {
  const body = { dataStream: 'A\r\x1fTarget\x1e\rTail\r\n',
    paragraphs: [{ startIndex: 1, paragraphId: 'p' }, { startIndex: 10, paragraphId: 'heading', paragraphStyle: { headingId: 'heading-alias' } },
      { startIndex: 15, paragraphId: 'tail' }], sectionBreaks: [{ startIndex: 16, sectionId: 'first' }],
    customRanges: [{ rangeType: 0, rangeId: 'link', startIndex: 0, endIndex: 0, properties: { bookmarkId: 'marked' } },
      { rangeType: 3, rangeId: 'marked', startIndex: 2, endIndex: 9, properties: {} },
      { rangeType: 3, rangeId: 'heading', startIndex: 11, endIndex: 14, properties: { bookmarkId: 'bookmark-alias' } }] }
  const model = publicModel(body), before = copy(model)
  assert.deepEqual(publicBody.documentHyperlinkTargetSelection(model, '#marked'), range('heading', 0))
  assert.deepEqual(publicBody.documentHyperlinkTargetSelection(model, '#heading'), range('tail', 0))
  assert.deepEqual(publicBody.documentHyperlinkTargetSelection(model, '#bookmark-alias'), range('tail', 0))
  assert.deepEqual(publicBody.documentHyperlinkTargetSelection(model, '#heading-alias'), range('heading', 0))
  assert.equal(publicBody.documentHyperlinkTargetSelection(model, '#missing'), undefined)
  assert.equal(publicBody.documentHyperlinkTargetSelection(model, 'https://example.org'), undefined)
  assert.deepEqual(model, before)
})

test('Canvas respects imported textStyleMode and textColorMode while link deletion restores author appearance', () => {
  for (const properties of [{}, { textColorMode: 'text' }, { textStyleMode: 'text' }]) {
    const model = publicModel({ dataStream: 'Link\r\n', paragraphs: [{ startIndex: 4, paragraphId: 'p' }],
      sectionBreaks: [{ startIndex: 5, sectionId: 'first' }], textRuns: [{ st: 0, ed: 4, ts: { cl: { rgb: '#883311' }, ul: { s: 0 } } }],
      customRanges: [{ rangeType: 0, rangeId: 'visual-link', startIndex: 0, endIndex: 3, properties: { url: 'https://example.org', ...properties } }] })
    const { document } = open(model)
    const draw = () => scene.sceneFromDocument(document.getModel(), layout.layoutDocument(document.getModel(), { measure: text => text.length * 8 }),
      { width: 1000, height: 700, scrollLeft: 0, scrollTop: 0, zoom: 100 },
      { selection: range('p'), focused: false, caretVisible: false, composing: false, images: new Map() })
    const commands = draw().commands, expectedColor = properties.textStyleMode === 'text' || properties.textColorMode === 'text' ? '#883311' : '#2563b6'
    assert.ok(commands.filter(command => command.kind === 'text').every(command => command.color === expectedColor))
    assert.equal(commands.filter(command => command.kind === 'line').length > 0, properties.textStyleMode !== 'text')
    document.executeCommand({ id: 'docs.command.delete-hyper-link', linkId: 'visual-link' }, range('p', 1))
    const removed = draw().commands
    assert.ok(removed.filter(command => command.kind === 'text').every(command => command.color === '#883311'))
    assert.equal(removed.filter(command => command.kind === 'line').length, 0)
  }
})

test('explicit public edit and click resolve the complete entity while the active caret is outside it, without content or history edits', () => {
  const model = publicModel(linkedBody())
  model.blocks.push(paragraph('outside', 'Outside'))
  const { studio, document } = open(model)
  studio.setDocumentInputSelection(document.id, range('outside', 2))
  const source = document.getModel(), before = copy(source), input = copy(document.getInputState()), count = history(studio, document.id)
  const current = document.getBody().customRanges.find(range => range.rangeType === 0), context = { documentId: document.id, userId: 'editor' }
  const request = { unitId: document.id, linkId: current.rangeId, startIndex: current.startIndex, endIndex: current.endIndex }
  for (const [id, params] of [['doc.operation.show-hyper-link-edit-popup', { link: request }],
    ['doc.operation.click-hyper-link', { unitId: document.id, linkId: current.rangeId }]]) {
    const resolved = publicBody.resolveDocumentHyperlinkOperation(source, id, params, context)
    assert.equal(resolved.kind, 'link'); assert.equal(resolved.source, source); assert.equal(resolved.model, source)
    assert.equal(resolved.link.id, 'imported-link'); assert.equal(resolved.link.label, 'AB\nCD')
    assert.deepEqual(resolved.link.selection, { anchor: { blockId: 'p', offset: 0 }, focus: { blockId: 'tail', offset: 2 } })
    assert.equal(resolved.host, undefined); assert.equal(resolved.segmentPage, undefined)
  }
  assert.deepEqual(document.getModel(), before); assert.deepEqual(document.getInputState(), input); assert.equal(history(studio, document.id), count)
})

test('public operation parsing rejects wrong identities, stale inclusive indices and malformed explicit requests without selection fallback', () => {
  const model = publicModel(linkedBody()), before = copy(model), current = publicBody.documentPublicBodyIndex(model).data.body.customRanges.find(range => range.rangeType === 0)
  const request = { unitId: model.id, linkId: current.rangeId, startIndex: current.startIndex, endIndex: current.endIndex }
  for (const patch of [{ unitId: 'other-document' }, { linkId: 'missing' }, { segmentId: 'missing' }, { segmentId: 4 },
    { startIndex: current.startIndex + 1 }, { endIndex: current.endIndex + 1 }, { startIndex: String(current.startIndex) },
    { segmentPage: -1 }, { segmentPage: 1.5 }]) {
    assert.equal(publicBody.resolveDocumentHyperlinkOperation(model, 'doc.operation.show-hyper-link-edit-popup', { link: { ...request, ...patch } }), undefined)
  }
  for (const params of [null, false, [], { link: null }, { link: 0 }, { link: {} }]) {
    assert.equal(publicBody.resolveDocumentHyperlinkOperation(model, 'doc.operation.show-hyper-link-edit-popup', params), undefined)
  }
  for (const params of [undefined, {}, { unitId: 'other-document', linkId: current.rangeId },
    { unitId: model.id, linkId: current.rangeId, segmentId: 'missing' }, { unitId: model.id, linkId: 'missing' }]) {
    assert.equal(publicBody.resolveDocumentHyperlinkOperation(model, 'doc.operation.click-hyper-link', params), undefined)
  }
  assert.deepEqual(publicBody.resolveDocumentHyperlinkOperation(model, 'doc.operation.show-hyper-link-edit-popup', undefined), { kind: 'selection' })
  assert.deepEqual(publicBody.resolveDocumentHyperlinkOperation(model, 'doc.operation.show-hyper-link-edit-popup', {}), { kind: 'selection' })
  const other = { ...copy(model), id: 'other-document' }
  assert.equal(publicBody.resolveDocumentHyperlinkOperation(other, 'doc.operation.click-hyper-link', { unitId: model.id, linkId: current.rangeId }), undefined)
  const ambiguous = copy(model)
  for (const noteKind of ['footnote', 'endnote']) {
    const reference = modelUtils.createRun('1', { metadata: { note: { kind: noteKind, noteId: 'shared-note' } } })
    ambiguous.blocks[0].runs.push(reference)
    ambiguous[noteKind === 'footnote' ? 'footnotes' : 'endnotes'] = { 'shared-note': { id: 'shared-note', text: noteKind,
      referenceBlockId: 'p', referenceRunId: reference.id, content: { blocks: [paragraph(noteKind + '-p', noteKind)], runs: [], paragraphs: [] } } }
  }
  assert.equal(publicBody.resolveDocumentHyperlinkOperation(ambiguous, 'doc.operation.click-hyper-link',
    { unitId: model.id, linkId: current.rangeId, segmentId: 'shared-note' }), undefined)
  assert.deepEqual(model, before)
})

test('public edit offsets count UTF-16, range delimiters and internal paragraph marks while excluding the final paragraph mark', () => {
  const { document } = open({ ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: [paragraph('p', 'A😀'), paragraph('empty', ''), paragraph('tail', 'BC')] })
  const selected = { anchor: { blockId: 'p', offset: 1 }, focus: { blockId: 'tail', offset: 2 } }
  document.executeCommand({ id: 'docs.command.add-hyper-link', payload: 'unicode.example' }, selected)
  const source = document.getModel(), index = publicBody.documentPublicBodyIndex(source), current = index.data.body.customRanges.find(range => range.rangeType === 0)
  const request = { unitId: source.id, linkId: current.rangeId, startIndex: current.startIndex, endIndex: current.endIndex }
  assert.equal(index.data.body.dataStream[current.startIndex], '\x1f'); assert.equal(index.data.body.dataStream[current.endIndex], '\x1e')
  assert.equal(index.data.body.dataStream[current.endIndex + 1], '\r')
  assert.equal(index.data.body.dataStream.slice(current.startIndex, current.endIndex + 1).replace(/[\x1f\x1e]/g, ''), '😀\r\rBC')
  assert.deepEqual(publicBody.documentPublicRangeToFlat(index, current.startIndex, current.endIndex + 1), { startOffset: 1, endOffset: 7 })
  const resolved = publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.show-hyper-link-edit-popup', { link: request })
  assert.deepEqual(resolved.link.selection, selected); assert.equal(resolved.link.label, '😀\n\nBC')
  for (const patch of [{ endIndex: current.endIndex + 1 }, { startIndex: current.startIndex + 2 }, { startIndex: 1, endIndex: 6 }]) {
    assert.equal(publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.show-hyper-link-edit-popup', { link: { ...request, ...patch } }), undefined)
  }
})

test('public story IDs and local aliases resolve the actual host while segmentPage remains a positioning hint and targets use source body', () => {
  const model = { ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: [paragraph('p', 'Body'), { ...paragraph('tail', 'Heading'), type: 'heading', level: 1 }] }
  const header = paragraph('header-p', 'Header link')
  header.runs[0].href = '#tail'
  model.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: 'Header link', footer: 'Footer',
    storyContent: { header: { blocks: [header], runs: [], paragraphs: [] }, footer: { blocks: [paragraph('footer-p', 'Footer')], runs: [], paragraphs: [] } } } }
  const source = open(model).document.getModel(), host = { kind: 'story', sectionId: 'first', story: 'header' }, before = copy(source)
  const segmentId = segments.documentSegmentIdForHost(source, host), current = publicBody.documentPublicBodyIndex(source, host).data.body.customRanges.find(range => range.rangeType === 0)
  const request = { unitId: source.id, linkId: current.rangeId, startIndex: current.startIndex, endIndex: current.endIndex }
  for (const segment of [segmentId, 'first:header']) for (const segmentPage of [0, 8]) {
    const resolved = publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.show-hyper-link-edit-popup', { link: { ...request, segmentId: segment, segmentPage } })
    assert.equal(resolved.kind, 'link'); assert.equal(resolved.source, source); assert.deepEqual(resolved.host, host); assert.equal(resolved.segmentPage, segmentPage)
    assert.equal(resolved.model.blocks[0].id, 'header-p'); assert.deepEqual(resolved.link.selection, range('header-p', 0, 11))
    assert.deepEqual(publicBody.documentHyperlinkTargetSelection(resolved.source, resolved.link.target), range('tail', 0))
  }
  const clicked = publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.click-hyper-link', { unitId: source.id, linkId: current.rangeId, segmentId })
  assert.deepEqual(clicked.host, host); assert.equal(clicked.link.target, '#tail')
  assert.equal(publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.click-hyper-link', { unitId: source.id, linkId: current.rangeId, segmentId: '' }), undefined)
  assert.equal(publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.click-hyper-link', { unitId: source.id, linkId: current.rangeId,
    segmentId: segments.documentSegmentIdForHost(source, { ...host, story: 'footer' }) }), undefined)
  assert.deepEqual(source, before)
})

test('explicit edit checks the complete link entity while click uses actual source View permission independently of Edit and content locks', () => {
  for (const locked of [false, true]) {
    const model = publicModel(linkedBody())
    if (locked) controls.createDocumentContentControl(model, { id: 'lock', type: 'richText', blockId: 'tail', startOffset: 0, endOffset: 2, lock: 'contentLocked' })
    if (!locked) model.permissionRules = [{ id: 'deny-tail-edit', scope: 'paragraph', targetId: 'tail', userIds: ['editor'], action: 'edit', effect: 'deny' }]
    const { studio, document } = open(model), source = document.getModel(), before = copy(source)
    const current = publicBody.documentPublicBodyIndex(source).data.body.customRanges.find(range => range.rangeType === 0), context = { documentId: source.id, userId: 'editor' }
    const request = { unitId: source.id, linkId: current.rangeId, startIndex: current.startIndex, endIndex: current.endIndex }
    assert.equal(publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.show-hyper-link-edit-popup', { link: request }, context), undefined)
    assert.equal(publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.click-hyper-link', { unitId: source.id, linkId: current.rangeId }, context).kind, 'link')
    assert.deepEqual(source, before); assert.equal(history(studio, source.id), 0)
    source.permissionRules = [...(source.permissionRules ?? []), { id: 'deny-tail-view', scope: 'paragraph', targetId: 'tail', userIds: ['editor'], action: 'view', effect: 'deny' }]
    assert.equal(publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.click-hyper-link', { unitId: source.id, linkId: current.rangeId }, context), undefined)
    assert.equal(history(studio, source.id), 0)
  }
})

test('native href runs share the first stable public range ID through complete-entity resolution, author styles and public body reopen', () => {
  const first = modelUtils.createRun('A', { href: 'https://native.example', bold: true, color: '#883311' }), second = modelUtils.createRun('😀', { href: 'https://native.example', italic: true })
  first.id = 'native-first'; second.id = 'native-second'
  const tail = modelUtils.createRun('Z', { href: 'https://native.example', fontSize: 18 }); tail.id = 'native-tail'
  const { studio, document } = open({ ...domain.createBlankDocument(), id: 'hyperlink-api',
    blocks: [{ ...paragraph('p', ''), runs: [first, second] }, { ...paragraph('tail', ''), runs: [tail] }] })
  const source = document.getModel(), original = authoredRuns(source), body = document.getBody(), current = body.customRanges.filter(range => range.rangeType === 0)
  assert.equal(current.length, 1); assert.equal(current[0].rangeId, 'hyperlink:native-first')
  for (const caret of [range('p', 1), range('tail', 0)]) {
    studio.setDocumentInputSelection(document.id, caret)
    const request = { unitId: source.id, linkId: current[0].rangeId, startIndex: current[0].startIndex, endIndex: current[0].endIndex }
    const resolved = publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.show-hyper-link-edit-popup', { link: request })
    assert.equal(resolved.link.id, current[0].rangeId); assert.equal(resolved.link.label, 'A😀\nZ')
    assert.deepEqual(resolved.link.selection, { anchor: { blockId: 'p', offset: 0 }, focus: { blockId: 'tail', offset: 1 } })
  }
  assert.equal(publicBody.resolveDocumentHyperlinkOperation(source, 'doc.operation.click-hyper-link', { unitId: source.id, linkId: 'hyperlink:native-second' }), undefined)
  const reopened = publicModel(body)
  assert.deepEqual(authoredRuns(reopened), original); assert.equal(hyperlinks.documentHyperlinks(reopened)[0].id, current[0].rangeId)
  assert.equal(publicBody.documentPublicBodyIndex(reopened).data.body.customRanges.filter(range => range.rangeType === 0)[0].rangeId, current[0].rangeId)
  assert.equal(history(studio, source.id), 0)
})

test('public href writing ignores empty runs, stops at real unlinked text and preserves explicitly distinct range identities', () => {
  const make = (id, text, rangeId) => {
    const run = modelUtils.createRun(text, { href: 'https://same.example' }); run.id = id
    if (rangeId) run.metadata = { univerBodyHyperlinkSource: { rangeId, rangeType: 0, properties: { url: run.href } } }
    return run
  }
  const model = { ...domain.createBlankDocument(), id: 'hyperlink-api', blocks: [{ ...paragraph('p', ''),
    runs: [make('first', 'A'), modelUtils.createRun(''), make('second', 'B'), modelUtils.createRun(' '), make('after-gap', 'C'),
      make('explicit-left', 'D', 'left-link'), make('explicit-right', 'E', 'right-link')] }] }
  const before = copy(model), entities = hyperlinks.documentHyperlinks(model), output = publicBody.documentPublicBodyIndex(model).data.body
  assert.deepEqual(entities.map(link => [link.id, link.label]), [['hyperlink:first', 'AB'], ['hyperlink:after-gap', 'C'], ['left-link', 'D'], ['right-link', 'E']])
  assert.deepEqual(output.customRanges.filter(range => range.rangeType === 0).map(range => range.rangeId), entities.map(link => link.id))
  const reopened = publicModel(output)
  assert.deepEqual(hyperlinks.documentHyperlinks(reopened).map(link => [link.id, link.label]), entities.map(link => [link.id, link.label]))
  assert.deepEqual(model, before)
})

test('all eight enclosing reference field types reject target edits, label edits, deletion and overlapping additions without model or history changes', () => {
  for (const fieldType of ['TOC', 'INDEX', 'TOA', 'REF', 'PAGEREF', 'NOTEREF', 'CITATION', 'BIBLIOGRAPHY']) {
    const body = linkedBody()
    body.customRanges[1] = { ...body.customRanges[1], endIndex: 4, properties: { fieldType, instruction: fieldType, cachedResult: 'AB\rCD' } }
    const { studio, document } = open(publicModel(body)), model = document.getModel(), before = copy(model), link = links(document)[0]
    const context = { documentId: document.id, userId: 'editor' }, state = hyperlinks.documentHyperlinkSelectionState(model, range('p', 1), context)
    assert.equal(state.available, false); assert.match(state.reason, /引用字段生成/); assert.equal(state.link.id, link.id)
    for (const command of [{ id: 'docs.command.update-hyper-link', linkId: link.id, payload: 'new.example', label: link.label },
      { id: 'docs.command.update-hyper-link', linkId: link.id, payload: 'new.example', label: 'Replacement' },
      { id: 'docs.command.delete-hyper-link', linkId: link.id }, { id: 'docs.command.add-hyper-link', payload: 'new.example' }]) {
      assert.throws(() => document.executeCommand(command, range('p', 0, 1)), /引用字段生成/)
      assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id), 0)
    }
    const current = document.getBody().customRanges.find(range => range.rangeType === 0)
    assert.equal(publicBody.resolveDocumentHyperlinkOperation(model, 'doc.operation.show-hyper-link-edit-popup', { link: {
      unitId: document.id, linkId: current.rangeId, startIndex: current.startIndex, endIndex: current.endIndex } }, context), undefined)
    assert.equal(publicBody.resolveDocumentHyperlinkOperation(model, 'doc.operation.click-hyper-link', { unitId: document.id, linkId: current.rangeId }, context).kind, 'link')
  }
})

test('referenceFieldLink uses current run properties or import provenance, and an explicit false overrides the old generated flag', () => {
  for (const layer of ['current', 'source', 'override']) {
    const model = publicModel(linkedBody({ url: 'https://old.example', ...(layer === 'current' ? {} : { referenceFieldLink: true }) }))
    if (layer === 'current') model.blocks.at(-1).runs.at(-1).metadata.link.referenceFieldLink = true
    for (const run of model.blocks.flatMap(block => block.runs)) {
      if (layer === 'source') delete run.metadata.link.referenceFieldLink
      if (layer === 'override') run.metadata.link.referenceFieldLink = false
    }
    const { studio, document } = open(model), before = copy(document.getModel()), link = links(document)[0]
    const state = hyperlinks.documentHyperlinkSelectionState(document.getModel(), range('p', 1))
    assert.equal(state.available, layer === 'override')
    const command = { id: 'docs.command.update-hyper-link', linkId: link.id, payload: 'new.example', label: link.label }
    if (layer === 'override') {
      document.executeCommand(command, range('p', 1)); assert.equal(history(studio, document.id), 1)
      assert.equal(document.getBody().customRanges.find(range => range.rangeId === link.id).properties.referenceFieldLink, false)
      continue
    }
    assert.throws(() => document.executeCommand(command, range('p', 1)), /引用字段生成/)
    assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id), 0)
  }
})

test('current native field ranges govern containment while ordinary AUTHOR, unknown fields and partially overlapping REF remain editable', () => {
  for (const [field, endBlockId, endOffset, blocked] of [['AUTHOR', 'tail', 2, false], ['CUSTOM', 'tail', 2, false],
    ['REF', undefined, 1, false], ['REF', 'tail', 2, true]]) {
    const model = publicModel(linkedBody())
    fields.setDocumentTextFields(model.blocks[0], [{ id: 'current-field', field, instruction: field, startOffset: 0, endOffset,
      ...(endBlockId ? { endBlockId } : {}) }])
    const { studio, document } = open(model), before = copy(document.getModel()), link = links(document)[0]
    assert.equal(hyperlinks.documentHyperlinkSelectionState(document.getModel(), range('tail', 1)).available, !blocked)
    const command = { id: 'docs.command.update-hyper-link', linkId: link.id, payload: 'new.example', label: link.label }
    if (blocked) {
      assert.throws(() => document.executeCommand(command, range('tail', 1)), /引用字段生成/)
      assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id), 0)
      continue
    }
    document.executeCommand(command, range('tail', 1)); assert.equal(history(studio, document.id), 1)
    assert.deepEqual(document.getModel().blocks.map(block => block.metadata), before.blocks.map(block => block.metadata))
    assert.deepEqual(authoredRuns(document.getModel()), authoredRuns(before))
    document.executeCommand({ id: 'docs.command.delete-hyper-link', linkId: link.id }, range('p', 1))
    assert.equal(document.getBody().customRanges.some(range => range.rangeType === 0), false)
    assert.equal(document.getBody().customRanges.find(range => range.rangeId === 'current-field').properties.fieldType, field)
  }
})

test('the shared selection query preserves full range and paragraph boundary matching independently of edit capability', () => {
  const model = publicModel(linkedBody({ url: 'https://old.example', referenceFieldLink: true }))
  model.blocks.push(paragraph('outside', 'Outside'))
  model.permissionRules = [{ id: 'deny-edit', scope: 'document', userIds: ['local-user'], action: 'edit', effect: 'deny' }]
  const before = copy(model)
  for (const selected of [range('p', 0), range('p', 2), range('tail', 2),
    { anchor: { blockId: 'tail', offset: 1 }, focus: { blockId: 'p', offset: 1 } }]) {
    assert.equal(hyperlinks.documentHyperlinkForSelection(model, selected).id, 'imported-link')
    assert.equal(hyperlinks.documentHyperlinkSelectionState(model, selected).available, false)
  }
  assert.equal(hyperlinks.documentHyperlinkForSelection(model, range('outside', 1)), undefined)
  assert.equal(hyperlinks.documentHyperlinkForSelection(model, { anchor: { blockId: 'p', offset: 1 }, focus: { blockId: 'outside', offset: 1 } }), undefined)
  assert.deepEqual(model, before)
})
