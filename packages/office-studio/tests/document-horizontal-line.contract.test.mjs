import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-02 drafts; execution belongs to the unified validation phase.
let server, OfficeStudio, FDocument, domain, modelUtils, borders, codec, controls, hosts, sections, layout, scene
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-horizontal-line', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, modelUtils, borders, codec, controls, hosts, sections, layout, scene] = await Promise.all([
    'domain/model', 'features/document/model/document-model', 'features/document/model/document-horizontal-line',
    'api/document-body-codec', 'features/document/services/content-controls', 'features/document/services/document-block-host',
    'features/document/services/section-header-footer', 'features/document/layout/document-layout',
    'features/document/rendering/document-scene',
  ].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const range = (blockId, start = 0, end = start) => ({ anchor: { blockId, offset: start }, focus: { blockId, offset: end } })
const paragraph = (id, text) => ({ ...modelUtils.createParagraph(text), id })
const history = (studio, id) => studio.editor.history.units[id]?.undo.length ?? 0
const text = block => modelUtils.runText(block.runs)
function open(model = { ...domain.createBlankDocument(), id: 'horizontal-line-api', blocks: [paragraph('p', 'ABCD'), paragraph('next', 'Tail')] }) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}
function publicModel(body, options = {}) {
  const decoded = codec.decodeDocumentBody(body, {}, options)
  return { ...domain.createBlankDocument(), id: 'horizontal-line-api', blocks: decoded.content.blocks,
    metadata: decoded.content.metadata, contentControls: decoded.contentControls }
}
function borderBody(border, text = '') {
  return { dataStream: text + '\rTail\r\n', paragraphs: [
    { startIndex: text.length, paragraphId: 'p', paragraphStyle: { borderBottom: border, spaceBelow: { v: 10 } } },
    { startIndex: text.length + 5, paragraphId: 'next' },
  ], sectionBreaks: [{ startIndex: text.length + 6, sectionId: 'first' }] }
}

test('the real horizontal-line command and existing alias split author text in one undoable transaction and persist the default border', () => {
  for (const id of ['doc.command.horizontal-line', 'doc.command.insert-divider']) {
    const model = { ...domain.createBlankDocument(), id: 'horizontal-line-api', blocks: [paragraph('p', 'ABCD'), paragraph('next', 'Tail')] }
    model.blocks[0].runs[0].bold = true
    const { studio, document } = open(model), before = copy(document.getModel()), events = []
    studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') events.push(event.id) })
    const result = document.executeCommand({ id }, range('p', 2)), blocks = document.getModel().blocks
    assert.equal(result.changed, true)
    assert.deepEqual(blocks.map(block => block.type), ['paragraph', 'divider', 'paragraph', 'paragraph'])
    assert.deepEqual(blocks.map(text), ['AB', '', 'CD', 'Tail'])
    assert.equal(blocks[0].runs[0].bold, true); assert.equal(blocks[2].runs[0].bold, true)
    assert.equal(blocks[1].style.spaceAfter, 10 * 4 / 3)
    assert.deepEqual(borders.documentHorizontalLineBorder(blocks[1]), borders.DOCUMENT_HORIZONTAL_LINE_BORDER)
    assert.deepEqual(events, [id]); assert.equal(history(studio, document.id), 1)
    const saved = document.save(), reopened = open(saved).document
    const written = document.getBody().paragraphs.find(paragraph => paragraph.paragraphId === blocks[1].id)
    assert.deepEqual(written.paragraphStyle.borderBottom, { width: 1, padding: 5, color: { rgb: '#CDD0D8' }, dashStyle: 1 })
    assert.equal(written.paragraphStyle.spaceBelow.v, 10)
    assert.equal(written.officeStudio.metadata?.documentDividerBorder, undefined)
    assert.deepEqual(reopened.getModel().blocks.map(text), ['AB', '', 'CD', 'Tail'])
    assert.deepEqual(borders.documentHorizontalLineBorder(reopened.getModel().blocks[1]), borders.DOCUMENT_HORIZONTAL_LINE_BORDER)
    assert.equal(document.undo(), true); assert.equal(document.getText(), before.blocks.map(text).join('\n'))
    assert.equal(document.redo(), true); assert.deepEqual(document.getModel().blocks.map(text), ['AB', '', 'CD', 'Tail'])
  }
})

test('insert-horizontal-line-bellow preserves the selected author text and inserts after the focus paragraph', () => {
  for (const selected of [range('p', 1, 3), range('p', 3, 1)]) {
    const { studio, document } = open()
    assert.equal(document.executeCommand({ id: 'doc.command.insert-horizontal-line-bellow' }, selected).changed, true)
    assert.deepEqual(document.getModel().blocks.map(text), ['ABCD', '', '', 'Tail'])
    assert.deepEqual(document.getModel().blocks.map(block => block.type), ['paragraph', 'divider', 'paragraph', 'paragraph'])
    assert.equal(history(studio, document.id), 1)
    assert.equal(document.undo(), true); assert.equal(document.getText(), 'ABCD\nTail')
    assert.equal(document.redo(), true); assert.deepEqual(document.getModel().blocks.map(text), ['ABCD', '', '', 'Tail'])
  }
})

test('empty borderBottom paragraphs import all public line styles and keep unknown border properties through body reopen', () => {
  for (const [dashStyle, style] of [[0, 'solid'], [1, 'solid'], [2, 'dot'], [3, 'dash']]) {
    const border = { width: 2.5, padding: 8, color: { rgb: '#a13579' }, dashStyle, authorHint: 'retained', offset: 7 }
    const { document } = open(publicModel(borderBody(border))), block = document.getModel().blocks[0]
    assert.equal(block.type, 'divider'); assert.equal(block.runs.length, 0)
    assert.deepEqual(borders.documentHorizontalLineBorder(block), { width: 2.5, padding: 8, color: '#a13579', style })
    const written = document.getBody().paragraphs[0].paragraphStyle.borderBottom
    assert.deepEqual(written, { ...border, dashStyle: dashStyle === 0 ? 1 : dashStyle })
    const reopened = publicModel(document.getBody())
    assert.equal(reopened.blocks[0].type, 'divider')
    assert.deepEqual(borders.documentHorizontalLineBorder(reopened.blocks[0]), borders.documentHorizontalLineBorder(block))
  }
})

test('a border-only paragraph remains a divider after heading and inherited list resolution while its list source stays serializable', () => {
  const body = borderBody({ width: 1, padding: 5, color: { rgb: '#CDD0D8' }, dashStyle: 1 })
  body.paragraphs[0].styleId = 'inherited-list'
  body.paragraphs[0].paragraphStyle.outlineLevel = 0
  const model = publicModel(body, { styles: { 'inherited-list': { bullet: { listType: 'CHECK_LIST', listId: 'counter', nestingLevel: 0 } } } })
  const { document } = open(model)
  assert.equal(document.getModel().blocks[0].type, 'divider')
  assert.equal(document.getBody().paragraphs[0].bullet.listId, 'counter')
  assert.equal(layout.layoutDocument(document.getModel(), { measure: text => text.length * 8 }).dividers.length, 1)
  assert.equal(publicModel(document.getBody()).blocks[0].type, 'divider')
})

test('nonempty paragraphs with a bottom border retain current public text even if old source metadata names a divider', () => {
  for (const source of [undefined, { version: 1, blockType: 'divider' }]) {
    const border = { width: 2, padding: 4, color: { rgb: '#a13579' }, dashStyle: 3, authorHint: 'retained' }
    const body = borderBody(border, 'Current text')
    if (source) body.paragraphs[0].officeStudio = source
    const { document } = open(publicModel(body))
    assert.equal(document.getModel().blocks[0].type, 'paragraph')
    assert.equal(text(document.getModel().blocks[0]), 'Current text')
    assert.deepEqual(document.getBody().paragraphs[0].paragraphStyle.borderBottom, border)
    assert.equal(text(publicModel(document.getBody()).blocks[0]), 'Current text')
  }
})

test('the public import boundary rejects negative widths, negative padding and unknown dash styles', () => {
  for (const border of [{ width: -1 }, { padding: -1 }, { dashStyle: 4 }, { dashStyle: 1.5 }]) {
    assert.throws(() => publicModel(borderBody(border)))
  }
})

test('body and cell layout share divider thickness, color and line style and Canvas draws scaled segments', () => {
  for (const [dashStyle, style] of [[1, 'solid'], [2, 'dot'], [3, 'dash']]) {
    const model = publicModel(borderBody({ width: 2, padding: 7, color: { rgb: '#a13579' }, dashStyle }))
    const table = modelUtils.createTable(1, 1), cellDivider = { ...copy(model.blocks[0]), id: 'cell-divider' }
    table.columnWidths = [160]
    table.rows[0].cells[0].blocks = [cellDivider, paragraph('cell-tail', 'Cell')]
    model.blocks.push({ ...paragraph('table', ''), type: 'table', runs: [], metadata: { table } })
    const before = copy(model), measured = layout.layoutDocument(model, { measure: text => text.length * 8 })
    assert.equal(measured.dividers.length, 2)
    for (const divider of measured.dividers) {
      assert.equal(divider.height, 2); assert.equal(divider.color, '#a13579'); assert.equal(divider.style, style)
    }
    const commands = scene.sceneFromDocument(model, measured,
      { width: 1400, height: 1400, scrollLeft: 0, scrollTop: 0, zoom: 150 },
      { selection: range('next'), focused: false, caretVisible: false, composing: false, images: new Map() }).commands
      .filter(command => command.kind === 'line' && command.color === '#a13579')
    assert.equal(commands.length === 2, style === 'solid')
    assert.ok(commands.length >= 2)
    assert.ok(commands.every(command => command.width === 3 && command.x2 > command.x1))
    assert.deepEqual(model, before)
  }
})

test('content locks and block edit permissions reject both insertion modes without content or history changes', () => {
  for (const locked of [false, true]) {
    const model = { ...domain.createBlankDocument(), id: 'horizontal-line-api', blocks: [paragraph('p', 'ABCD'), paragraph('next', 'Tail')] }
    if (locked) controls.createDocumentContentControl(model, { id: 'locked', type: 'richText', blockId: 'p', startOffset: 0, endOffset: 4, lock: 'contentLocked' })
    if (!locked) model.permissionRules = [{ id: 'deny-p', scope: 'block', targetId: 'p', userIds: ['editor'], action: 'edit', effect: 'deny' }]
    const { studio, document } = open(model), before = copy(document.getModel())
    for (const id of ['doc.command.horizontal-line', 'doc.command.insert-horizontal-line-bellow']) {
      assert.throws(() => document.executeCommand({ id, ...(locked ? { activeContentControlId: 'locked' } : {}) }, range('p', 1, 3)))
      assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id), 0)
    }
  }
})

test('story insertion edits the addressed header, keeps body content and restores the same divider through undo, redo and save', () => {
  const model = { ...domain.createBlankDocument(), id: 'horizontal-line-api', blocks: [paragraph('p', 'Body')] }
  model.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: 'Header',
    storyContent: { header: { blocks: [paragraph('header-p', 'Header')], runs: [], paragraphs: [] } } } }
  const { studio, document } = open(model), beforeBody = copy(document.getModel().blocks)
  const host = { kind: 'story', sectionId: 'first', story: 'header' }
  document.executeCommand({ id: 'doc.command.insert-horizontal-line-bellow', host }, range('header-p', 1, 4))
  const blocks = hosts.documentModelForBlockHost(document.getModel(), host).blocks
  assert.deepEqual(blocks.map(text), ['Header', '', ''])
  assert.equal(blocks[1].type, 'divider'); assert.deepEqual(document.getModel().blocks, beforeBody)
  assert.equal(history(studio, document.id), 1)
  assert.equal(document.undo(), true); assert.deepEqual(hosts.documentModelForBlockHost(document.getModel(), host).blocks.map(text), ['Header'])
  assert.equal(document.redo(), true)
  const reopened = open(document.save()).document, saved = hosts.documentModelForBlockHost(reopened.getModel(), host).blocks
  assert.equal(saved[1].id, blocks[1].id)
  assert.deepEqual(borders.documentHorizontalLineBorder(saved[1]), borders.DOCUMENT_HORIZONTAL_LINE_BORDER)
  assert.deepEqual(reopened.getModel().blocks, beforeBody)
})
