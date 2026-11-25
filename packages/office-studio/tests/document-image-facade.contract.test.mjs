import assert from 'node:assert/strict'
import { File as NodeFile } from 'node:buffer'
import { afterAll, afterEach, beforeAll, beforeEach, test, vi } from 'vitest'
import { createServer } from 'vite'

// Deferred integration/IO drafts use the real Facade, command bus, source codec
// and history. Browser decoding is separately exercised by the browser draft.
let server, OfficeStudio, FDocument, domain, modelUtils, imageApi, commands, editors, codec, sections, storyBlocks, structureMenus, ribbon, toolbar
let pendingImages, deferDecode, naturalWidth, naturalHeight
const studios = new Set(), copy = value => structuredClone(value)
const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg=='
const bytes = Buffer.from(png.split(',')[1], 'base64')
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-image-facade', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, modelUtils, imageApi, commands, editors, codec, sections, storyBlocks, structureMenus, ribbon, toolbar] = await Promise.all(['domain/model', 'features/document/model/document-model',
    'api/document-images', 'api/document-image-commands', 'api/document-image-editors', 'api/document-body-codec',
    'features/document/services/section-header-footer', 'features/document/model/document-story-blocks', 'components/document/document-structure-ribbon',
    'components/document/document-ribbon', 'components/document/document-toolbar-state']
    .map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
beforeEach(() => {
  pendingImages = []; deferDecode = false; naturalWidth = 1; naturalHeight = 1
  vi.stubGlobal('File', NodeFile)
  vi.stubGlobal('Image', class {
    onload = null; onerror = null; complete = false; naturalWidth = 0; naturalHeight = 0
    get src() { return this.value ?? '' }
    set src(value) {
      this.value = value
      if (!value) { pendingImages = pendingImages.filter(image => image !== this); return }
      if (deferDecode) { if (!pendingImages.includes(this)) pendingImages.push(this) }
      else queueMicrotask(() => this.finish())
    }
    finish() { if (!this.value) return; this.complete = true; this.naturalWidth = naturalWidth; this.naturalHeight = naturalHeight; this.onload?.() }
  })
  vi.stubGlobal('FileReader', class {
    result = null; onload = null; onerror = null; onabort = null
    readAsArrayBuffer(file) { file.arrayBuffer().then(result => { if (!this.aborted) { this.result = result; this.onload?.() } }) }
    abort() { this.aborted = true; this.onabort?.() }
  })
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear(); vi.unstubAllGlobals() })
afterAll(async () => { await server?.close() })
const paragraph = (id, text = '') => ({ ...modelUtils.createParagraph(text), id })
const selection = (id = 'p', offset = 0) => ({ anchor: { blockId: id, offset }, focus: { blockId: id, offset } })
const picture = (id, layer = 0) => ({ ...paragraph(id), type: 'image', runs: [], metadata: { image: { src: png, alt: id, width: 100, height: 60,
  intrinsicWidth: 1, intrinsicHeight: 1, layer, wrappingStyle: 'inFrontOfText' } } })
function open(input) {
  const source = input ?? { ...domain.createBlankDocument(), id: 'image-api', blocks: [paragraph('p', 'ABCD'), paragraph('tail', 'Tail')] }
  const studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(source.id)
  return { studio, document: new FDocument(studio, source.id) }
}
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })
const waitForDecode = async (count = 1) => {
  for (let attempt = 0; attempt < 20 && pendingImages.length < count; attempt++) await Promise.resolve()
  assert.ok(pendingImages.length >= count)
}
const completeDecode = () => { for (const image of pendingImages.splice(0)) image.finish() }

test('Facade insertion, object query, formatting, crop and removal use canonical content commands and one history per change', async () => {
  const { studio, document } = open(), events = []
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') events.push(event.id) })
  const image = await document.insertImage({ source: png, imageSourceType: 'BASE64', width: 120, alt: 'Inserted' })
  assert.ok(image); assert.equal(image.getUnitId(), document.id); assert.equal(document.getImage(image.getId()).getId(), image.getId())
  assert.deepEqual(image.getSize(), { width: 120, height: 120 }); assert.equal(image.getSourceType(), 'BASE64')
  assert.equal(history(studio, document.id).undo, 1)
  assert.equal(image.setRotate(45), true); assert.equal(image.setCrop({ top: 0, right: 0.25, bottom: 0, left: 0 }), true)
  assert.equal(history(studio, document.id).undo, 3); assert.equal(document.undo(), true); assert.equal(image.getImage().crop, undefined)
  assert.equal(document.redo(), true); assert.equal(image.remove(), true); assert.equal(document.getImage(image.getId()), null)
  assert.equal(document.undo(), true); assert.equal(image.getAngle(), 45)
  assert.deepEqual(events.filter(id => id.includes('image')), ['doc.command.insert-image', 'doc.command.update-image', 'doc.command.update-image', 'doc.command.remove-image'])
})

test('public insertion returns Promise<boolean> and five source/range/identity guards leave history unchanged on failure', async () => {
  const { studio, document } = open(), insert = commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.insert
  const result = studio.executeCommand(insert, { unitId: document.id, source: png, imageSourceType: 'BASE64', width: 80 })
  assert.equal(typeof result.then, 'function'); assert.equal(await result, true)
  const before = copy(document.getModel()), oldHistory = history(studio, document.id)
  for (const request of [{ source: png, imageSourceType: 'URL' }, { source: 'uuid', imageSourceType: 'UUID' },
    { source: png, expectedRevision: 0 }, { source: png, expectedHistoryGeneration: 0 },
    { source: png, segmentId: 'missing' }, { source: png, segmentId: '', subUnitId: 'other' },
    { source: png, textRange: { startOffset: 2, endOffset: 1 } }]) await assert.rejects(() => studio.executeCommand(insert, { unitId: document.id, ...request }))
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), oldHistory)
})

test('transform batches convert srcRect after the final size notification and publish one history entry', () => {
  const { studio, document } = open({ ...domain.createBlankDocument(), id: 'transform-api', blocks: [paragraph('p'), picture('a'), picture('b', 1)] })
  const transform = commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.transform
  assert.equal(studio.executeCommandSync(transform, { unitId: document.id, drawings: [
    { drawingId: 'a', key: 'srcRect', value: { left: 20, right: 20 } }, { drawingId: 'a', key: 'size', value: { width: 160, height: 80 } },
    { drawingId: 'b', key: 'angle', value: 90 },
  ] }), true)
  assert.equal(document.getImage('a').getImage().crop.left, 0.1); assert.equal(document.getImage('b').getAngle(), 90)
  assert.equal(history(studio, document.id).undo, 1)
  const before = copy(document.getModel()), previous = history(studio, document.id)
  assert.throws(() => studio.executeCommandSync(transform, { unitId: document.id, drawings: [{ drawingId: 'a', key: 'angle', value: 45 }, { drawingId: 'b', key: 'size', value: { width: -1, height: 80 } }] }))
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous)
})

test('public numeric arrangement matches upstream order, preserves a selected group and uses one undo transaction', () => {
  const { studio, document } = open({ ...domain.createBlankDocument(), id: 'arrange-api', blocks: [paragraph('p'), ...['a', 'b', 'c', 'd'].map(picture)] })
  const arrange = commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.arrange
  assert.equal(studio.executeCommandSync(arrange, { unitId: document.id, drawingIds: ['c', 'a'], arrangeType: 2 }), true)
  assert.deepEqual(document.getImages().sort((a, b) => a.getImage().layer - b.getImage().layer).map(image => image.getId()), ['b', 'd', 'a', 'c'])
  assert.equal(history(studio, document.id).undo, 1); assert.equal(document.undo(), true)
  assert.equal(studio.executeCommandSync(arrange, { unitId: document.id, drawingIds: ['d'], zOrder: 1.9 }), true)
  assert.deepEqual(document.getImages().sort((a, b) => a.getImage().layer - b.getImage().layer).map(image => image.getId()), ['a', 'd', 'b', 'c'])
})

test('external abort and selection navigation cancel pending source decode before insertion or replacement', async () => {
  const { studio, document } = open(); deferDecode = true
  const abort = new AbortController(), before = copy(document.getModel())
  const insertion = document.insertImage({ source: png, signal: abort.signal }); await waitForDecode(); abort.abort()
  await assert.rejects(insertion, error => error.name === 'AbortError'); assert.deepEqual(document.getModel(), before)
  const next = document.insertImage({ source: png }); await waitForDecode(); document.setInputSelection(selection('tail', 1)); completeDecode()
  await assert.rejects(next, error => error.name === 'AbortError'); assert.equal(history(studio, document.id).undo, 0)
  assert.equal(document.getInputState().location.selection.focus.blockId, 'tail')
})

test('studio disposal cancels its IO owner while another studio with the same document ID keeps its own pending decode', async () => {
  const first = open(), second = open(); deferDecode = true
  const a = first.document.insertImage({ source: png }), b = second.document.insertImage({ source: png })
  await waitForDecode(2); first.studio.dispose(); studios.delete(first.studio)
  await assert.rejects(a, error => error.name === 'AbortError'); completeDecode()
  assert.ok(await b); assert.equal(second.document.getImages().length, 1)
})

test('File and URL reads embed actual source bytes, preserve original URL and reject mismatched decoded dimensions', async () => {
  const { document } = open()
  const local = await document.insertImage({ source: new File([bytes], 'local.png') })
  assert.equal(local.getImage().byteLength, bytes.length); assert.equal(local.getImage().mimeType, 'image/png')
  vi.stubGlobal('fetch', async () => new Response(bytes, { status: 200 }))
  const remote = await document.insertImage({ source: 'https://example.test/image.png', imageSourceType: 'URL' })
  assert.equal(remote.getSource(), 'https://example.test/image.png'); assert.equal(remote.getSourceType(), 'URL')
  assert.ok(remote.getImage().src.startsWith('data:image/png;base64,'))
  naturalWidth = 2
  await assert.rejects(() => imageApi.prepareDocumentImageSource({ id: 'bad', name: 'Bad', mimeType: 'image/png', source: png, byteLength: bytes.length, width: 1, height: 1 }), /尺寸/)
})

test('image editor registrations are studio/unit isolated and replaced or stale handles cannot close the current editor', async () => {
  const first = open(), second = open(), closed = [], opened = []
  first.document.setInputSelection(selection())
  const a = first.document.bindImageEditor({ open: target => { opened.push(['first', target]); return true }, close: () => { closed.push('first'); return true } })
  const b = second.document.bindImageEditor({ open: target => { opened.push(['second', target]); return true }, close: () => { closed.push('second'); return true } })
  assert.equal(await first.studio.executeCommand(commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.insert, { unitId: first.document.id }), true)
  assert.equal(opened.length, 1); assert.equal(opened[0][0], 'first'); assert.equal(history(first.studio, first.document.id).undo, 0)
  const latest = first.document.bindImageEditor({ open: () => true, close: () => { closed.push('latest'); return true } })
  assert.equal(a.isCurrent(), false); a.dispose(); assert.equal(latest.isCurrent(), true); assert.equal(b.isCurrent(), true)
  const stale = opened[0][1]; first.document.setInputSelection(selection('tail', 1))
  first.document.executeCommand({ id: 'doc.command.insert-text', text: 'x' }, selection('tail', 1))
  assert.equal(editors.requestDocumentImageEditorOpen(first.studio, stale), false)
})

test('the five fixed paragraph mounts share the real insert operation with their recorded ordering and below placement', async () => {
  const { studio, document } = open(), opened = []
  document.setInputSelection(selection())
  document.bindImageEditor({ open: target => { opened.push(target); return true }, close: () => true })
  const mounts = ['layout-below', 'content-insert', 'empty-paragraph', 'quick-insert', 'quick-insert-below']
  for (const mount of mounts) {
    const entries = structureMenus.documentStructuredMountedMenuItems({ paragraphId: 'p', disabled: false }, mount)
    const item = entries.find(item => item.id.endsWith('-image'))
    assert.ok(item); assert.equal(item.commandId, commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.insert)
    assert.equal(item.order, mount.startsWith('quick') ? 1 : 5)
    assert.equal(entries.indexOf(item), mount.startsWith('quick') ? 1 : entries.length - 1)
    assert.equal(await studio.executeCommand(item.commandId, { unitId: document.id, ...item.params }), true)
    assert.equal(opened.at(-1).paragraphMenuPlacement, ['layout-below', 'content-insert', 'quick-insert-below'].includes(mount) ? 'below' : undefined)
    assert.equal(opened.at(-1).mode, 'insert'); assert.deepEqual(opened.at(-1).selection, selection())
  }
  for (const mount of ['quick-edit', 'table-below']) assert.equal(structureMenus.documentStructuredMountedMenuItems({ paragraphId: 'p', disabled: false }, mount).some(item => item.id.endsWith('-image')), false)
  assert.equal(history(studio, document.id).undo, 0)
})

test('the regular media root keeps its grid/classic command and derives its disabled state from the editing context', () => {
  const { document } = open(), state = toolbar.documentToolbarState(document.getModel(), selection())
  const context = { state, canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false, paper: 'A4', orientation: 'portrait',
    zoom: 100, contentControlsAvailable: true }
  const find = tabs => tabs.flatMap(tab => tab.groups.flatMap(group => group.items)).find(item => item.id === 'doc-image')
  const root = find(ribbon.createDocumentRibbon(context))
  assert.equal(root.commandId, commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.insert); assert.equal(root.order, 0)
  assert.deepEqual(root.gridLayout, { row: 1, column: 4 }); assert.equal(!!root.disabled, false)
  assert.equal(find(ribbon.createDocumentRibbon({ ...context, editingDisabled: true })).disabled, true)
})

test('header and note segment image ownership is isolated across APIs, commands and public save/reopen', () => {
  const source = { ...domain.createBlankDocument(), id: 'segment-api', blocks: [paragraph('p', 'Body'), picture('body-image')] }
  const header = storyBlocks.projectDocumentStoryBlocks({ blocks: [paragraph('header-p'), picture('header-image')], runs: [], paragraphs: [], metadata: {} }).content
  source.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: '\uFFFC', storyContent: { header } } }
  const { studio, document } = open(source), headerId = document.getSnapshot().documentStyle.defaultHeaderId
  assert.ok(headerId); assert.equal(document.getImage('header-image'), null); assert.equal(document.getImages(headerId)[0].getId(), 'header-image')
  assert.throws(() => studio.executeCommandSync(commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.transform, { unitId: document.id, segmentId: '', drawings: [{ drawingId: 'header-image', key: 'angle', value: 45 }] }))
  assert.equal(document.getImage('header-image', headerId).setRotate(45), true)
  const noteId = document.insertFootnote(1, codec.encodeDocumentBody({ blocks: [paragraph('note-p'), picture('note-image')], runs: [], paragraphs: [], metadata: {} }))
  assert.ok(noteId); assert.equal(document.getImage('note-image', noteId).setRotate(20), true)
  const reopened = open(document.save()).document
  assert.equal(reopened.getImage('header-image', headerId).getAngle(), 45); assert.equal(reopened.getImage('note-image', noteId).getAngle(), 20)
  assert.equal(reopened.getImage('body-image').getAngle(), 0)
})
