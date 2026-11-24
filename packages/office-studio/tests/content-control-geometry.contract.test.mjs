import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-11 layout and Canvas-scene drafts; no browser or renderer acceptance is implied.
let server, models, modelUtils, controls, geometry, rendering, documentScene
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-geometry', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  models = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  geometry = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/content-control-geometry.ts')
  rendering = await server.ssrLoadModule('/packages/office-studio/src/features/document/rendering/content-control-scene.ts')
  documentScene = await server.ssrLoadModule('/packages/office-studio/src/features/document/rendering/document-scene.ts')
})
afterAll(async () => { await server?.close() })

function fixture(text = 'ABCDEFGH') {
  const model = models.createBlankDocument()
  model.id = 'sdt-geometry'; model.blocks[0].id = 'p'; model.blocks[0].runs[0].text = text; model.contentControls = []
  return model
}
function add(model, patch = {}) {
  return controls.createDocumentContentControl(model, { id: 'control', type: 'text', blockId: 'p', startOffset: 1, endOffset: 7, ...patch })
}
function line(text, patch = {}) {
  let offset = patch.start ?? 0, x = patch.x ?? 20
  const y = patch.y ?? 20, height = patch.height ?? 16
  const glyphs = [...text].map(character => {
    const result = { x, y, width: character.length * 10, height, offset, end: offset + character.length,
      text: character, run: modelUtils.createRun(character), font: '16px Aptos' }
    x += result.width; offset = result.end
    return result
  })
  return { x: patch.x ?? 20, y, width: x - (patch.x ?? 20), height, start: patch.start ?? 0, end: offset,
    target: { blockId: 'p' }, baseline: y + 13, page: 0, ...patch, glyphs }
}
function layout(lines, patch = {}) {
  return { width: 300, height: 424, pages: [{ index: 0, x: 0, y: 0, width: 300, height: 200 }, { index: 1, x: 0, y: 224, width: 300, height: 200 }],
    lines, cells: [], images: [], dividers: [], columnSeparators: [], headings: [], pageHeaderFooters: [], pageContent: [], ...patch }
}

test('wrapped ranges resolve one fragment per line with exact UTF-16 start/end and endpoint flags', () => {
  const model = fixture(); add(model)
  const result = geometry.documentContentControlGeometry(model, layout([line('ABCD'), line('EFGH', { start: 4, y: 40 })]))[0]
  assert.deepEqual(result.fragments.map(fragment => [fragment.x, fragment.y, fragment.width, fragment.startOffset, fragment.endOffset, fragment.first, fragment.last]),
    [[30, 20, 30, 1, 4, true, false], [20, 40, 30, 4, 7, false, true]])
  assert.deepEqual(result.selection, { anchor: { blockId: 'p', offset: 1 }, focus: { blockId: 'p', offset: 7 } })
})

test('an empty range at a wrap boundary belongs to the next line and remains hittable', () => {
  const model = fixture(); add(model, { startOffset: 4, endOffset: 4 })
  const result = geometry.documentContentControlGeometry(model, layout([line('ABCD'), line('EFGH', { start: 4, y: 40 })]))
  assert.equal(result[0].fragments.length, 1)
  assert.deepEqual([result[0].fragments[0].x, result[0].fragments[0].y, result[0].fragments[0].width], [20, 40, 12])
  assert.equal(geometry.hitTestDocumentContentControl(result, 25, 45)?.control.id, 'control')
  assert.equal(geometry.hitTestDocumentContentControl(result, 25, 25), undefined)
})

test('an empty end-of-paragraph range uses the last line when no next line exists', () => {
  const model = fixture('ABCD'); add(model, { startOffset: 4, endOffset: 4 })
  const fragment = geometry.documentContentControlGeometry(model, layout([line('ABCD')]))[0].fragments[0]
  assert.deepEqual([fragment.x, fragment.width, fragment.startOffset], [60, 12, 4])
})

test('supplementary characters use complete glyph bounds and UTF-16 model offsets', () => {
  const model = fixture('A😀B'); add(model, { startOffset: 1, endOffset: 3 })
  const result = geometry.documentContentControlGeometry(model, layout([line('A😀B')]))[0]
  assert.deepEqual([result.fragments[0].x, result.fragments[0].width, result.fragments[0].startOffset, result.fragments[0].endOffset], [30, 20, 1, 3])
})

test('equal-range children are drawn after their parents and hit testing chooses the deepest child', () => {
  const model = fixture('NAME')
  add(model, { id: 'z-parent', type: 'group', startOffset: 0, endOffset: 4 })
  add(model, { id: 'a-child', type: 'group', parentId: 'z-parent', startOffset: 0, endOffset: 4 })
  add(model, { id: 'b-grandchild', parentId: 'a-child', startOffset: 0, endOffset: 4 })
  const result = geometry.documentContentControlGeometry(model, layout([line('NAME')]))
  assert.deepEqual(result.map(item => item.control.id), ['z-parent', 'a-child', 'b-grandchild'])
  assert.equal(geometry.hitTestDocumentContentControl(result, 35, 25)?.control.id, 'b-grandchild')
})

test('table-cell controls keep the table owner coordinates and match their actual cell paragraph', () => {
  const model = fixture('CELL')
  const table = modelUtils.createParagraph(); table.id = 'table'; table.type = 'table'; table.runs = []
  table.metadata = { table: { rows: [{ id: 'row', cells: [{ id: 'cell', blocks: [model.blocks[0]], style: {}, rowSpan: 1, colSpan: 1 }] }] } }
  model.blocks = [table]
  controls.createDocumentContentControl(model, { id: 'cell-control', type: 'text', blockId: 'p', startOffset: 0, endOffset: 4 })
  const target = { blockId: 'table', row: 0, col: 0, cellBlockId: 'p' }
  const result = geometry.documentContentControlGeometry(model, layout([line('CELL', { target, clip: { x: 25, y: 20, width: 30, height: 16 } })]))[0]
  assert.deepEqual(result.target, target)
  assert.deepEqual(geometry.contentControlFragmentVisibleRect(result.fragments[0]), { x: 25, y: 20, width: 30, height: 16 })
  assert.deepEqual(result.selection.focus, { ...target, offset: 4 })
})

test('a line whose clip lies wholly outside its page cannot create an unbounded visible fragment', () => {
  const model = fixture('NAME'); add(model, { startOffset: 0, endOffset: 4 })
  const result = geometry.documentContentControlGeometry(model, layout([line('NAME', { clip: { x: 400, y: 20, width: 80, height: 16 } })]))
  assert.deepEqual(result, [])
})

test('partial clips constrain hit testing to the intersection with the physical page', () => {
  const model = fixture('NAME'); add(model, { startOffset: 0, endOffset: 4 })
  const result = geometry.documentContentControlGeometry(model, layout([line('NAME', { clip: { x: 40, y: 20, width: 500, height: 16 } })]))
  assert.deepEqual(geometry.contentControlFragmentVisibleRect(result[0].fragments[0]), { x: 40, y: 20, width: 20, height: 16 })
  assert.equal(geometry.hitTestDocumentContentControl(result, 30, 25), undefined)
  assert.equal(geometry.hitTestDocumentContentControl(result, 45, 25)?.control.id, 'control')
})

test('repeated-header replicas retain page affinity in the selection request', () => {
  const model = fixture('NAME'); add(model, { startOffset: 2, endOffset: 2 })
  const result = geometry.documentContentControlGeometry(model, layout([line('NAME', { repeatedHeader: true }), line('NAME', { page: 1, y: 244, repeatedHeader: true })]))[0]
  assert.deepEqual(result.fragments.map(fragment => fragment.page), [0, 1])
  const selection = geometry.contentControlSelectionAtFragment(result, result.fragments[1])
  assert.equal(selection.anchor.pageAffinity, 1); assert.equal(selection.focus.pageAffinity, 1)
  assert.equal(selection.focus.offset, 2)
})

test('picture controls use persistent image-block bounds instead of the empty text anchor', () => {
  const model = fixture(''); add(model, { type: 'picture', startOffset: 0, endOffset: 0, metadata: { pictureBlockId: 'image' } })
  const result = geometry.documentContentControlGeometry(model, layout([line('')], { images: [
    { blockId: 'image', x: 90, y: 50, width: 80, height: 60, page: 0, src: 'data:image/png;base64,AA==', alt: 'Picture', clip: { x: 100, y: 50, width: 70, height: 60 } },
  ] }))[0]
  assert.deepEqual([result.fragments[0].x, result.fragments[0].y, result.fragments[0].width, result.fragments[0].height], [90, 50, 80, 60])
  assert.equal(geometry.hitTestDocumentContentControl([result], 95, 60), undefined)
  assert.equal(geometry.hitTestDocumentContentControl([result], 120, 60)?.control.type, 'picture')
})

test('an image clip outside the physical page cannot expose a picture-control border', () => {
  const model = fixture(''); add(model, { type: 'picture', startOffset: 0, endOffset: 0, metadata: { pictureBlockId: 'image' } })
  const result = geometry.documentContentControlGeometry(model, layout([line('')], { images: [
    { blockId: 'image', x: 350, y: 20, width: 80, height: 60, page: 0, src: 'data:image/png;base64,AA==', alt: 'Clipped', clip: { x: 350, y: 20, width: 80, height: 60 } },
  ] }))
  assert.deepEqual(result, [])
})

test('hidden controls retain hit geometry while producing no Canvas boundary commands', () => {
  const model = fixture('NAME'); add(model, { appearance: 'hidden', startOffset: 0, endOffset: 4 })
  const source = layout([line('NAME')]), result = geometry.documentContentControlGeometry(model, source)
  const scene = { width: 600, height: 500, commands: [] }
  rendering.appendDocumentContentControlScene(scene, result, source, { width: 600, height: 500, scrollTop: 0, scrollLeft: 0, zoom: 100 },
    { activeControlId: 'control', primary: '#123456', border: '#999999' })
  assert.deepEqual(scene.commands, []); assert.equal(geometry.hitTestDocumentContentControl(result, 30, 25)?.control.id, 'control')
})

test('tag appearance paints only the true start and end brackets across wrapped fragments', () => {
  const model = fixture(); add(model, { appearance: 'tags', color: '00AAFF' })
  const source = layout([line('ABCD'), line('EFGH', { start: 4, y: 40 })]), result = geometry.documentContentControlGeometry(model, source)
  const scene = { width: 600, height: 500, commands: [] }
  rendering.appendDocumentContentControlScene(scene, result, source, { width: 600, height: 500, scrollTop: 0, scrollLeft: 0, zoom: 100 },
    { activeControlId: 'control', primary: '#123456', border: '#999999' })
  assert.equal(scene.commands.length, 6)
  assert.ok(scene.commands.every(command => command.kind === 'line' && command.color === '#00AAFF' && command.width === 1.5))
})

test('Canvas boundaries apply viewport zoom, scrolling and physical clipping', () => {
  const model = fixture('NAME'); add(model, { startOffset: 0, endOffset: 4 })
  const source = layout([line('NAME', { clip: { x: 25, y: 20, width: 30, height: 16 } })]), result = geometry.documentContentControlGeometry(model, source)
  const viewport = { width: 600, height: 500, scrollTop: 10, scrollLeft: 5, zoom: 150 }, offset = documentScene.sceneOffset(source, viewport)
  const scene = { width: 600, height: 500, commands: [] }
  rendering.appendDocumentContentControlScene(scene, result, source, viewport, { primary: '#123456', border: '#999999' })
  assert.equal(scene.commands.length, 1)
  assert.deepEqual([scene.commands[0].x, scene.commands[0].y, scene.commands[0].width, scene.commands[0].height], [offset.x + 30, offset.y + 30, 60, 24])
  assert.deepEqual(scene.commands[0].clip, { x: offset.x + 37.5, y: offset.y + 30, width: 45, height: 24 })
})
