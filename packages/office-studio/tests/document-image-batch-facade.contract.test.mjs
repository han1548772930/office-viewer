import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred DOC-07 integration contracts. This batch did not start Vite or run tests.
let server, OfficeStudio, FDocument, FOffice, domain, models, atoms, imageModel, batch, layoutApi, layoutEngine, commands, controls, sections, stories, codec
const studios = new Set(), bindings = new Set(), copy = value => structuredClone(value)
const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg=='
const measurer = { measure: text => Array.from(text).length * 8 }, options = { measurer }
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-6, `${actual} != ${expected}`)
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-image-batch-facade', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument, FOffice } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, models, atoms, imageModel, batch, layoutApi, layoutEngine, commands, controls, sections, stories, codec] = await Promise.all([
    'domain/model', 'features/document/model/document-model', 'features/document/images/document-image-atom',
    'features/document/images/document-image-model', 'features/document/images/document-image-batch', 'api/document-image-layout',
    'features/document/layout/document-layout', 'api/document-image-commands', 'features/document/services/content-controls',
    'features/document/services/section-header-footer', 'features/document/model/document-story-blocks', 'api/document-body-codec',
  ].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const binding of bindings) binding.dispose(); bindings.clear(); for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const paragraph = (id, text = '') => ({ ...models.createParagraph(text), id })
const image = (id, x, y, layer = 0, patch = {}) => ({ ...paragraph(id), type: 'image', runs: [], metadata: { image: {
  src: png, alt: id, resourceId: id + '-resource', width: 100, height: 60, intrinsicWidth: 1, intrinsicHeight: 1, layer,
  wrappingStyle: 'inFrontOfText', positionH: { relativeFrom: 0, posOffset: x }, positionV: { relativeFrom: 0, posOffset: y }, ...patch,
} } })
const source = () => ({ ...domain.createBlankDocument(), id: 'batch-facade', blocks: [paragraph('p', 'Body'), image('a', 60, 170),
  image('b', 230, 220, 1, { flipX: true }), image('c', 410, 340, 2), image('d', 140, 480, 3), paragraph('tail', 'Tail')],
  page: { width: 640, height: 800, marginLeft: 40, marginRight: 40, marginTop: 40, marginBottom: 40 } })
function open(input = source()) {
  const studio = new OfficeStudio({ document: input, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(input.id)
  return { studio, document: new FDocument(studio, input.id) }
}
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })
const datum = (document, id, segmentId = '') => document.getImage(id, segmentId).getImage()
const order = document => document.getImages().map(image => ({ id: image.getId(), layer: image.getImage().layer ?? 0 })).sort((a, b) => a.layer - b.layer).map(image => image.id)
function bind(source, read) { const handle = layoutApi.bindDocumentImageLayout(source, read); bindings.add(handle); return handle }
function bindLive(document) {
  const model = document.getModel()
  return bind(model, () => ({ source: model, layout: layoutEngine.layoutDocument(model, measurer) }))
}

test('Facade alignment accepts local handles and IDs and creates one event/history before undo, redo and public snapshot reopening', () => {
  const { studio, document } = open(), before = copy(document.getModel()), events = []
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') events.push(event.id) })
  assert.equal(document.alignImages([document.getImage('a'), 'b', document.getImage('c')], '1', '', options), true)
  for (const id of ['a', 'b', 'c']) near(datum(document, id).positionH.posOffset, 60)
  assert.deepEqual(datum(document, 'd'), before.blocks.find(block => block.id === 'd').metadata.image)
  assert.deepEqual(events, ['doc.command.align-images']); assert.equal(history(studio, document.id).undo, 1)
  const aligned = copy(document.getModel())
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), before)
  assert.equal(document.redo(), true); assert.deepEqual(document.getModel(), aligned)
  const reopened = open(copy(document.save())).document
  for (const id of ['a', 'b', 'c']) near(datum(reopened, id).positionH.posOffset, 60)
  assert.equal(datum(reopened, 'b').flipX, true); assert.equal(datum(reopened, 'a').resourceId, 'a-resource')
})

test('batch translation, uniform scaling and rotation preserve selected identities, crops and flip data through the same history', () => {
  const input = source(); input.blocks[1].metadata.image.crop = { top: 0.1, right: 0.2, bottom: 0, left: 0 }
  const { studio, document } = open(input), before = copy(document.getModel())
  assert.equal(document.transformImages(['a', 'b', 'c'], { x: 12, y: 7, scale: 1.25, angle: 90 }, '', options), true)
  for (const id of ['a', 'b', 'c']) { near(datum(document, id).width, 125); near(datum(document, id).height, 75); near(datum(document, id).angle, 90) }
  assert.deepEqual(datum(document, 'a').crop, before.blocks[1].metadata.image.crop); assert.equal(datum(document, 'b').flipX, true)
  assert.equal(history(studio, document.id).undo, 1); assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), before)
  assert.equal(document.redo(), true); assert.equal(open(copy(document.save())).document.getImage('b').getFlipX(), true)
})

test('multi-object front/back arrangement preserves the original relative order rather than the handle argument order', () => {
  const { studio, document } = open()
  assert.equal(document.arrangeImages(['c', 'a'], 'front'), true); assert.deepEqual(order(document), ['b', 'd', 'a', 'c'])
  assert.equal(history(studio, document.id).undo, 1); assert.equal(document.undo(), true); assert.deepEqual(order(document), ['a', 'b', 'c', 'd'])
  assert.equal(document.arrangeImages(['c', 'a'], 'back'), true); assert.deepEqual(order(document), ['a', 'c', 'b', 'd'])
  assert.deepEqual(order(open(copy(document.save())).document), ['a', 'c', 'b', 'd'])
})

test('absolute arrangement floors/clamps a single index and rejects an absolute multi-object target without publishing', () => {
  const { studio, document } = open()
  assert.equal(document.arrangeImages(['a'], 2.9), true); assert.deepEqual(order(document), ['b', 'c', 'a', 'd'])
  assert.equal(document.undo(), true); assert.equal(document.arrangeImages(['a'], 100), true); assert.deepEqual(order(document), ['b', 'c', 'd', 'a'])
  const before = copy(document.getModel()), oldHistory = history(studio, document.id)
  assert.throws(() => document.arrangeImages(['a', 'b'], 1)); assert.throws(() => document.arrangeImages(['a'], NaN))
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), oldHistory)
})

test('a batch removal uses one content history and reopens the same surviving resources', () => {
  const { studio, document } = open(), before = copy(document.getModel())
  assert.equal(document.removeImages([document.getImage('b'), 'a']), true)
  assert.deepEqual(document.getImages().map(image => image.getId()), ['c', 'd']); assert.equal(history(studio, document.id).undo, 1)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), before)
  assert.equal(document.redo(), true); const reopened = open(copy(document.save())).document
  assert.equal(reopened.getImage('a'), null); assert.equal(datum(reopened, 'd').resourceId, 'd-resource')
})

test('removing two atoms from one paragraph preserves its surrounding UTF-16 text and public drawing IDs across undo and reopening', () => {
  const input = source(), p = paragraph('p')
  const atom = id => { const run = models.createRun(); atoms.setDocumentImageAtom(run, id, image(id, 0, 0).metadata.image); return run }
  p.runs = [models.createRun('L😀'), atom('a'), models.createRun('M'), atom('b'), models.createRun('R')]; input.blocks = [p, paragraph('tail')]
  const { studio, document } = open(input)
  assert.equal(document.removeImages(['a', 'b']), true); assert.equal(document.getParagraph('p').getText(), 'L😀MR')
  assert.equal(history(studio, document.id).undo, 1); assert.equal(document.undo(), true)
  assert.equal(document.getParagraph('p').getText(), 'L😀\uFFFCM\uFFFCR')
  assert.deepEqual(imageModel.documentImageEntries(document.getModel()).map(entry => entry.inline.startOffset), [3, 5])
  const reopened = open(copy(document.save())).document
  assert.deepEqual(reopened.getImages().map(image => image.getId()), ['a', 'b'])
})

test('no-op geometry and already-front peers leave content, undo/redo and savepoint generation unchanged', () => {
  const { studio, document } = open(), before = copy(document.getModel()), generation = document.getHistoryGeneration()
  assert.equal(document.transformImages(['a', 'b'], { x: 0, y: 0, scale: 1, angle: 360 }, '', options), false)
  assert.equal(document.alignImages(['a', 'b'], '0', '', options), false)
  assert.equal(document.alignImages(['a', 'b'], '7', '', options), false)
  assert.equal(document.arrangeImages(['c', 'd'], 'front'), false)
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 })
  assert.equal(document.getHistoryGeneration(), generation)
})

test('invalid target lists, transforms and read options fail before changing any member or history', () => {
  const { studio, document } = open(), before = copy(document.getModel()), oldHistory = history(studio, document.id)
  for (const ids of [[], ['a', 'a'], ['a', 'missing'], [null], [1], [''], Array.from({ length: 1001 }, (_, index) => 'id-' + index)]) {
    assert.throws(() => document.transformImages(ids, { x: 10 }, '', options)); assert.throws(() => document.removeImages(ids))
  }
  for (const transform of [{ scale: 0 }, { scale: Infinity }, { angle: NaN }, { x: 10, skew: 1 }, { origin: { x: 0, y: 0, other: 1 } }]) {
    assert.throws(() => document.transformImages(['a', 'b'], transform, '', options))
  }
  for (const readOptions of [null, { preferredPage: -1, measurer }, { preferredPage: 0.5, measurer }, { measurer: {} }, { measurer, stale: true }]) {
    assert.throws(() => document.alignImages(['a', 'b'], '1', '', readOptions))
  }
  assert.throws(() => document.alignImages(['a', 'b'], '9', '', options))
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), oldHistory)
})

test('handles from another Studio with the same document ID cannot select or delete local objects', () => {
  const local = open(), foreign = open(copy(local.document.getModel())), before = copy(local.document.getModel())
  for (const operation of [() => local.document.alignImages([local.document.getImage('a'), foreign.document.getImage('b')], '1', '', options),
    () => local.document.transformImages([foreign.document.getImage('a'), 'b'], { x: 10 }, '', options),
    () => local.document.arrangeImages([foreign.document.getImage('a')], 'front'), () => local.document.removeImages([foreign.document.getImage('a')])]) assert.throws(operation, /其他文档实例/)
  assert.deepEqual(local.document.getModel(), before); assert.equal(history(local.studio, local.document.id).undo, 0)
  assert.equal(foreign.document.transformImages(['a', 'b'], { x: 10 }, '', options), true); near(datum(local.document, 'a').positionH.posOffset, 60)
})

test('a locked later target rejects geometry and deletion atomically while leaving an unlocked image editable', () => {
  const input = source(); input.blocks.splice(2, 0, paragraph('lock-anchor', 'Locked'))
  controls.createDocumentContentControl(input, { id: 'lock', type: 'richText', placement: 'block', blockRange: { blockIds: ['lock-anchor', 'b'] },
    blockId: 'lock-anchor', startOffset: 0, endOffset: 6, lock: 'contentLocked' })
  const { studio, document } = open(input), before = copy(document.getModel())
  assert.throws(() => document.alignImages(['a', 'b'], '1', '', options), /锁定/)
  assert.throws(() => document.removeImages(['a', 'b']), /锁定/)
  assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id).undo, 0)
  assert.equal(document.transformImages(['a'], { x: 5 }, '', options), true); near(datum(document, 'a').positionH.posOffset, 65)
})

test('view and edit permissions are checked even on an unchanged alignment boundary image', () => {
  for (const action of ['view', 'edit']) {
    const input = source(); input.permissionRules = [{ id: 'deny-a', scope: 'entity', targetId: 'a', action, effect: 'deny', userIds: ['editor'] }]
    const { studio, document } = open(input), before = copy(document.getModel())
    assert.throws(() => document.alignImages(['a', 'b'], '1', '', options), /permission denied/)
    assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id).undo, 0)
  }
})

test('arrangement validates every affected peer before publishing a selected image at the front', () => {
  const input = source(); input.permissionRules = [{ id: 'deny-b', scope: 'entity', targetId: 'b', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  const { studio, document } = open(input), before = copy(document.getModel())
  assert.throws(() => document.arrangeImages(['a'], 'front'), /permission denied/)
  assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id).undo, 0)
})

test('header batches remain in their segment and keep IDs and positions through save/reopen', () => {
  const input = source(), header = stories.projectDocumentStoryBlocks({ blocks: [paragraph('header-p', 'Header'), image('h1', 20, 10), image('h2', 160, 30)],
    runs: [], paragraphs: [], metadata: {} }).content
  input.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: 'Header\uFFFC\uFFFC', storyContent: { header } } }
  const { studio, document } = open(input), segmentId = document.getSnapshot().documentStyle.defaultHeaderId, body = copy(document.getModel().blocks)
  assert.ok(segmentId); assert.throws(() => document.transformImages(['h1', 'h2'], { x: 10 }, '', options), /指定内容区/)
  assert.throws(() => document.removeImages([document.getImage('a')], segmentId), /其他文档实例或内容区/)
  assert.equal(document.transformImages([document.getImage('h1', segmentId), 'h2'], { x: 10 }, segmentId, options), true)
  assert.deepEqual(document.getModel().blocks, body); assert.equal(history(studio, document.id).undo, 1)
  const x = datum(document, 'h1', segmentId).positionH.posOffset, saved = document.save(), reopened = open(copy(saved)).document
  assert.deepEqual(reopened.getImages(segmentId).map(image => image.getId()), ['h1', 'h2']); near(datum(reopened, 'h1', segmentId).positionH.posOffset, x)
})

test('note segment batch removal updates the note body once and preserves the note reference through undo and reopening', () => {
  const { studio, document } = open(), noteId = document.insertFootnote(0, codec.encodeDocumentBody({ blocks: [paragraph('note-p', 'Note'), image('n1', 20, 10), image('n2', 160, 30)],
    runs: [], paragraphs: [], metadata: {} }))
  assert.ok(noteId); const oldHistory = history(studio, document.id)
  assert.equal(document.removeImages(['n1', 'n2'], noteId), true); assert.equal(document.getImages(noteId).length, 0)
  assert.equal(history(studio, document.id).undo, oldHistory.undo + 1); assert.equal(document.undo(), true)
  const reopened = open(copy(document.save())).document
  assert.deepEqual(reopened.getImages(noteId).map(image => image.getId()), ['n1', 'n2']); assert.ok(reopened.getModel().footnotes[noteId])
})

test('replacement layout owners survive old disposal and current disposal restores the detached measurer path', () => {
  const { document } = open(), model = document.getModel(), first = layoutEngine.layoutDocument(model, measurer), second = copy(first)
  const old = bind(model, () => ({ source: model, layout: first })), current = bind(model, () => ({ source: model, layout: second }))
  old.dispose(); assert.equal(layoutApi.readDocumentImageLayout(document, '', options), second)
  current.dispose(); let measured = 0
  const detached = layoutApi.readDocumentImageLayout(document, '', { measurer: { measure(text) { measured++; return measurer.measure(text) } } })
  assert.notEqual(detached, second); assert.ok(measured > 0)
  assert.throws(() => layoutApi.bindDocumentImageLayout(model, undefined)); assert.throws(() => layoutApi.bindDocumentImageLayout(null, () => undefined))
})

test('a live layout is borrowed without remeasuring and a source/host mismatch falls back to the authoritative layout', () => {
  const { document } = open(), model = document.getModel(), layout = layoutEngine.layoutDocument(model, measurer)
  bind(model, () => ({ source: model, layout }))
  assert.equal(layoutApi.readDocumentImageLayout(document, '', { measurer: { measure() { throw new Error('Unexpected measurement') } } }), layout)
  bind(model, () => ({ source: copy(model), layout })); assert.notEqual(layoutApi.readDocumentImageLayout(document, '', options), layout)
  bind(model, () => ({ source: model, host: { kind: 'note', noteId: 'foreign-note' }, layout }))
  assert.notEqual(layoutApi.readDocumentImageLayout(document, '', options), layout)
})

test('a reentrant layout callback cannot apply stale geometry after an independently committed content change', () => {
  const { studio, document } = open(), model = document.getModel(), layout = layoutEngine.layoutDocument(model, measurer), beforeA = datum(document, 'a')
  bind(model, () => { document.getImage('d').update({ alt: 'External change' }); return { source: model, layout } })
  assert.throws(() => document.transformImages(['a', 'b'], { x: 10 }, '', options), /版本已变化/)
  assert.deepEqual(datum(document, 'a'), beforeA); assert.equal(datum(document, 'd').alt, 'External change')
  assert.equal(history(studio, document.id).undo, 1)
})

test('BeforeCommandExecute cancellation reports false and preserves the entire batch and history', () => {
  const { studio, document } = open(), before = copy(document.getModel())
  studio.events.addEvent('BeforeCommandExecute', event => { if (event.kind === 'document' && event.id === 'doc.command.align-images') event.cancel = true })
  assert.equal(document.alignImages(['a', 'b'], '1', '', options), false)
  assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id).undo, 0)
})

test('public alignment and transform commands share the document history and expose the fixed string enum', () => {
  const { studio, document } = open(), api = FOffice.newAPI(studio)
  assert.deepEqual(api.Enum.DocumentImageAlignType, batch.DOCUMENT_IMAGE_ALIGN_TYPES)
  assert.deepEqual(Object.values(api.Enum.DocumentImageAlignType), ['0', '1', '2', '3', '4', '5', '6', '7', '8'])
  bindLive(document)
  assert.equal(studio.executeCommandSync(commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.align, { unitId: document.id, drawingIds: ['a', 'b', 'c'], alignType: api.Enum.DocumentImageAlignType.left, preferredPage: 0 }), true)
  bindLive(document)
  assert.equal(studio.executeCommandSync(commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.batchTransform, { unitId: document.id, drawingIds: ['a', 'b', 'c'], transform: { x: 10, y: 5 }, preferredPage: 0 }), true)
  for (const id of ['a', 'b', 'c']) near(datum(document, id).positionH.posOffset, 70)
  assert.equal(history(studio, document.id).undo, 2)
})

test('public stale revision/history, duplicate IDs and conflicting segments reject before invoking a layout provider', () => {
  const { studio, document } = open(), before = copy(document.getModel()); let reads = 0
  const model = document.getModel(); bind(model, () => { reads++; return { source: model, layout: layoutEngine.layoutDocument(model, measurer) } })
  const base = { unitId: document.id, drawingIds: ['a', 'b'], alignType: '1' }
  for (const request of [{ expectedRevision: document.getModelRevision() + 1 }, { expectedHistoryGeneration: document.getHistoryGeneration() + 1 },
    { drawingIds: ['a', 'a'] }, { segmentId: '', subUnitId: 'foreign' }, { unitId: 'missing' }]) assert.throws(() => studio.executeCommandSync(commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.align, { ...base, ...request }))
  assert.equal(reads, 0); assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id).undo, 0)
})
