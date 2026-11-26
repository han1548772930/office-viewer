import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, createTable, runText } from '../src/features/document/model/document-model.ts'
import { setDocumentImageAtom } from '../src/features/document/images/document-image-atom.ts'
import { documentImageEntries } from '../src/features/document/images/document-image-model.ts'
import { DOCUMENT_IMAGE_ALIGN_TYPES, alignDocumentImageBatch, assertDocumentImageBatchMembers, documentImageBatchBounds,
  documentImageBatchMembers, transformDocumentImageBatch } from '../src/features/document/images/document-image-batch.ts'
import { documentImageBox, documentImageRotationHandle } from '../src/features/document/layout/document-image-layout.ts'
import { emptyLayout, layoutDocument } from '../src/features/document/layout/document-layout.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { captureDocumentImageCanvasTarget, documentImageAtRotationHandle, documentImageBatchGestureResult, documentImageCanvasCommandMatchesTarget,
  documentImageCanvasTargetCurrent, documentImageGestureResult } from '../src/features/document/controllers/document-image-canvas.ts'
import { documentImagesClipboardData } from '../src/features/document/services/document-clipboard.ts'
import { createDocumentContentControl } from '../src/features/document/services/content-controls.ts'
import { sceneFromDocument, sceneOffset } from '../src/features/document/rendering/document-scene.ts'

// Deferred DOC-07 contracts. No test, renderer, Vite or browser was started in this batch.
const copy = value => structuredClone(value), near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-6, `${actual} != ${expected}`)
const src = 'data:image/png;base64,AA==', origin = { x: 13, y: 29 }, measurer = { measure: text => Array.from(text).length * 8 }
const paragraph = (id, text = '') => ({ ...createParagraph(text), id })
const image = (id, x, y, width = 90, height = 60, angle = 0) => ({ ...paragraph(id), type: 'image', runs: [], metadata: { image: {
  src, alt: id, resourceId: 'resource-' + id, width, height, angle, wrappingStyle: 'inFrontOfText',
  positionH: { relativeFrom: 0, posOffset: x }, positionV: { relativeFrom: 0, posOffset: y },
} } })
const document = () => ({ ...createBlankDocument(), id: 'image-batch', blocks: [paragraph('p', 'Body'), image('a', 50, 90),
  image('b', 250, 140, 120, 50, 90), image('c', 450, 220, 70, 40), paragraph('tail', 'Tail')],
  page: { width: 640, height: 800, marginLeft: 40, marginRight: 40, marginTop: 40, marginBottom: 40 } })
const selection = { anchor: { blockId: 'p', offset: 0 }, focus: { blockId: 'p', offset: 0 } }, permission = { documentId: 'image-batch', userId: 'editor' }
const datum = (model, id) => documentImageEntries(model).find(entry => entry.imageId === id).image
function controlledLayout(model) {
  const layout = emptyLayout()
  layout.pages = [{ ...origin, index: 0, pageNumber: 1, width: 640, height: 800, content: { x: 53, y: 69, width: 560, height: 720 } }]
  layout.images = documentImageEntries(model).map(entry => ({ blockId: entry.imageId, objectImage: true, src: entry.image.src, alt: entry.image.alt,
    x: origin.x + (entry.image.positionH?.posOffset ?? 50), y: origin.y + (entry.image.positionV?.posOffset ?? 90),
    ...documentImageBox(entry.image, 560, 720), angle: entry.image.angle, page: 0, wrappingStyle: entry.image.wrappingStyle }))
  return layout
}
const members = (model, ids = ['a', 'b', 'c']) => documentImageBatchMembers(model, controlledLayout(model), ids)
const execute = (model, command, selected = selection) => executeDocumentCommand(model, selected, command, permission)

test('all six alignments use visible rotated bounds and leave local dimensions, resource identity and the other axis intact', () => {
  const cases = [
    ['1', [50, 50, 50], [90, 140, 220]], ['2', [240, 260, 250], [90, 140, 220]], ['3', [430, 470, 450], [90, 140, 220]],
    ['4', [50, 250, 450], [90, 90, 90]], ['5', [50, 250, 450], [145, 115, 155]], ['6', [50, 250, 450], [200, 140, 220]],
  ]
  assert.deepEqual(Object.values(DOCUMENT_IMAGE_ALIGN_TYPES), ['0', '1', '2', '3', '4', '5', '6', '7', '8'])
  for (const [align, xs, ys] of cases) {
    const model = document(), before = ['a', 'b', 'c'].map(id => copy(datum(model, id)))
    execute(model, { id: 'doc.command.align-images', members: members(model), align })
    for (const [index, id] of ['a', 'b', 'c'].entries()) {
      const next = datum(model, id)
      near(next.positionH.posOffset, xs[index]); near(next.positionV.posOffset, ys[index])
      assert.deepEqual([next.width, next.height, next.angle, next.resourceId], [before[index].width, before[index].height, before[index].angle, before[index].resourceId])
    }
  }
})

test('horizontal distribution uses equal edge gaps and the full selection span even if the widest right edge belongs to the middle object', () => {
  const model = document(); model.blocks = [paragraph('p'), image('a', 0, 50, 10, 20), image('b', 20, 80, 200, 30), image('c', 40, 100, 10, 40)]
  const current = members(model), updates = alignDocumentImageBatch(current, '7'), byId = new Map(updates.map(update => [update.imageId, update]))
  const positions = current.map(member => byId.get(member.imageId)?.rectangle ?? member.rectangle)
  near(positions[0].x - origin.x, 0); near(positions[1].x - origin.x, 10); near(positions[2].x - origin.x, 210)
  near(positions[1].x - positions[0].x - positions[0].width, 0); near(positions[2].x - positions[1].x - positions[1].width, 0)
  assert.deepEqual(model.blocks.map(block => block.id), ['p', 'a', 'b', 'c'])
})

test('vertical and negative-gap distribution keep each image size and permit deliberate overlap', () => {
  const model = document(); model.blocks = [paragraph('p'), image('a', 0, 0, 80, 80), image('b', 60, 60, 80, 80), image('c', 90, 90, 80, 80)]
  for (const align of ['7', '8']) {
    const current = members(model), updates = alignDocumentImageBatch(current, align), byId = new Map(updates.map(update => [update.imageId, update]))
    const frames = current.map(member => byId.get(member.imageId)?.rectangle ?? member.rectangle), horizontal = align === '7'
    const value = frame => horizontal ? frame.x : frame.y, size = frame => horizontal ? frame.width : frame.height
    near(value(frames[1]) - value(frames[0]) - size(frames[0]), -35)
    near(value(frames[2]) - value(frames[1]) - size(frames[1]), -35)
    assert.ok(frames.every(frame => frame.width === 80 && frame.height === 80))
  }
})

test('default alignment, distribution of two objects and a complete rotation do not mutate any source image', () => {
  const model = document(), before = copy(model), current = members(model)
  assert.deepEqual(alignDocumentImageBatch(current, '0'), [])
  assert.deepEqual(alignDocumentImageBatch(current.slice(0, 2), '7'), [])
  assert.deepEqual(transformDocumentImageBatch(current, { angle: 360 }), [])
  execute(model, { id: 'doc.command.transform-images', members: current, transform: { x: 0, y: 0, scale: 1 } })
  assert.deepEqual(model, before)
})

test('a batch rotation moves centers around the common origin and retains each local crop frame and flip state', () => {
  const model = document(); Object.assign(datum(model, 'a'), { flipX: true, crop: { top: 0.1, right: 0, bottom: 0, left: 0 } })
  const before = copy(model), current = members(model), updates = transformDocumentImageBatch(current, { angle: 90, y: 100 })
  const expected = [[340, 85, 60, 90, 90], [260, 265, 120, 50, 180], [220, 475, 40, 70, 90]]
  for (const [index, update] of updates.entries()) {
    const [centerX, centerY, width, height, angle] = expected[index]
    near(update.rectangle.x + update.rectangle.width / 2 - origin.x, centerX); near(update.rectangle.y + update.rectangle.height / 2 - origin.y, centerY)
    near(update.rectangle.width, width); near(update.rectangle.height, height); near(update.patch.angle, angle)
  }
  assert.deepEqual(model, before)
  execute(model, { id: 'doc.command.transform-images', members: current, transform: { angle: 90, y: 100 } })
  assert.equal(datum(model, 'a').flipX, true); assert.deepEqual(datum(model, 'a').crop, before.blocks[1].metadata.image.crop)
})

test('uniform corner resize holds the opposite selection corner and applies the same scale to every image', () => {
  const model = document(), current = members(model), rectangle = documentImageBatchBounds(current)
  const target = captureDocumentImageCanvasTarget(model, 'c', selection, undefined, undefined, ['a', 'b', 'c'])
  const start = { x: rectangle.x + rectangle.width, y: rectangle.y + rectangle.height }
  const gesture = { target, members: current, rectangle, start, corner: 'se', pointerId: 9, zoom: 150 }
  const result = documentImageBatchGestureResult(gesture, rectangle.x + rectangle.width * 2, rectangle.y + rectangle.height * 2)
  near(result.transform.scale, 2); near(result.rectangle.x, rectangle.x); near(result.rectangle.y, rectangle.y)
  for (const [index, update] of result.updates.entries()) {
    near(update.patch.width, current[index].image.width * 2); near(update.patch.height, current[index].image.height * 2)
  }
  assert.deepEqual(documentImageBatchGestureResult(gesture, start.x, start.y).updates, [])
})

test('the default center remains valid when a legal image box extends beyond the top-left coordinate range', () => {
  const model = document(); model.blocks = [paragraph('p'), image('a', 99900, 90, 100000, 60), image('b', 99800, 180, 100000, 50)]
  const current = members(model, ['a', 'b']), before = copy(model)
  assertDocumentImageBatchMembers(model, current)
  const bounds = documentImageBatchBounds(current)
  assert.ok(bounds.x + bounds.width / 2 > 100000)
  assert.deepEqual(transformDocumentImageBatch(current, { x: 0, angle: 360 }), [])
  const updates = transformDocumentImageBatch(current, { x: -100 })
  near(updates[0].patch.positionH.posOffset, 99800); near(updates[1].patch.positionH.posOffset, 99700)
  assert.deepEqual(model, before)
  execute(model, { id: 'doc.command.transform-images', members: current, transform: { x: -100 } })
  near(datum(model, 'a').positionH.posOffset, 99800); near(datum(model, 'b').positionH.posOffset, 99700)
})

test('rotation handle hit testing keeps its screen distance at different zooms and snap uses fifteen-degree increments', () => {
  const model = document(), current = members(model), rectangle = documentImageBatchBounds(current)
  for (const zoom of [50, 100, 200]) {
    const handle = documentImageRotationHandle(rectangle, zoom)
    near(Math.hypot(handle.x - handle.stemX, handle.y - handle.stemY) * zoom / 100, 24)
    assert.equal(documentImageAtRotationHandle(rectangle, handle.x, handle.y, zoom), true)
    assert.equal(documentImageAtRotationHandle(rectangle, handle.x + 20 * 100 / zoom, handle.y, zoom), false)
  }
  const center = { x: rectangle.x + rectangle.width / 2, y: rectangle.y + rectangle.height / 2 }, start = documentImageRotationHandle(rectangle, 100)
  const radius = start.y - center.y, radians = 22 * Math.PI / 180
  const gesture = { target: captureDocumentImageCanvasTarget(model, 'c', selection), members: current, rectangle, start, rotate: true, pointerId: 3, zoom: 100 }
  const result = documentImageBatchGestureResult(gesture, center.x - radius * Math.sin(radians), center.y + radius * Math.cos(radians), true)
  near(result.transform.angle, 15)
})

test('single rotated-object gesture preserves its center and updates a page-relative position without changing the model during preview', () => {
  const model = document(), current = members(model, ['a'])[0], rectangle = current.rectangle, start = documentImageRotationHandle(rectangle, 100)
  const center = { x: rectangle.x + rectangle.width / 2, y: rectangle.y + rectangle.height / 2 }, before = copy(model)
  const result = documentImageGestureResult({ target: captureDocumentImageCanvasTarget(model, 'a', selection), image: current.image, rectangle,
    start, rotate: true, pointerId: 1, pageOrigin: origin, zoom: 100 }, center.x + Math.abs(start.y - center.y), center.y)
  near(result.patch.angle, 90); near(result.rectangle.x + result.rectangle.width / 2, center.x); near(result.rectangle.y + result.rectangle.height / 2, center.y)
  near(result.patch.positionH.posOffset, result.rectangle.x - origin.x); assert.deepEqual(model, before)
})

test('converting a scaled inline atom to a float retains the painted size and its original UTF-16 address', () => {
  const model = document(), run = createRun(), block = paragraph('p')
  setDocumentImageAtom(run, 'a', { src, alt: 'a', width: 800, height: 400, wrappingStyle: 'inline' })
  block.runs = [createRun('L'), run, createRun('R')]; model.blocks = [block, paragraph('tail')]
  const layout = controlledLayout(model); Object.assign(layout.images[0], { width: 200, height: 100, paintWidth: 200, paintHeight: 100 })
  const current = documentImageBatchMembers(model, layout, ['a']), anchor = copy(documentImageEntries(model)[0].inline)
  execute(model, { id: 'doc.command.transform-images', members: current, transform: { x: 10, y: 15 } })
  assert.deepEqual([datum(model, 'a').width, datum(model, 'a').height, datum(model, 'a').wrappingStyle], [200, 100, 'inFrontOfText'])
  const next = documentImageEntries(model)[0].inline
  assert.deepEqual([next.target, next.startOffset, next.endOffset, next.run.id], [anchor.target, anchor.startOffset, anchor.endOffset, anchor.run.id])
  assert.equal(runText(model.blocks[0].runs), 'L\uFFFCR')
})

test('invalid geometry, duplicate IDs, foreign pages and stale image content reject before publishing an earlier update', () => {
  const model = document(), current = members(model), before = copy(model)
  const invalid = [current.map(member => ({ ...member, rectangle: { ...member.rectangle, width: NaN } })), [current[0], current[0]],
    current.map((member, index) => index ? { ...member, rectangle: { ...member.rectangle, page: 1 } } : member),
    current.map((member, index) => index ? { ...member, image: { ...member.image, alt: 'old source' } } : member)]
  for (const value of invalid) {
    assert.throws(() => execute(model, { id: 'doc.command.transform-images', members: value, transform: { x: 10 } }))
    assert.deepEqual(model, before)
  }
  for (const transform of [{ scale: 0 }, { scale: Infinity }, { x: NaN }, { origin: { x: NaN, y: 0 } }, { skew: 1 }, { angle: 0, origin: { x: 0, y: 0, extra: true } }]) {
    assert.throws(() => execute(model, { id: 'doc.command.transform-images', members: current, transform }))
    assert.deepEqual(model, before)
  }
})

test('cross-container geometry is rejected while independent batch removal remains an atomic content operation', () => {
  const model = document(), table = createTable(1, 1); table.rows[0].cells[0].blocks = [image('nested', 30, 50)]
  model.blocks.push({ ...paragraph('table'), type: 'table', runs: [], metadata: { table } })
  const current = members(model, ['a', 'nested']), before = copy(model)
  assert.throws(() => assertDocumentImageBatchMembers(model, current), /同一内容容器/)
  assert.throws(() => execute(model, { id: 'doc.command.align-images', members: current, align: '1' }), /同一内容容器/)
  assert.deepEqual(model, before)
  execute(model, { id: 'doc.command.batch-images', commands: ['a', 'nested'].map(imageId => ({ id: 'doc.command.remove-image', imageId })) })
  assert.equal(documentImageEntries(model).some(entry => ['a', 'nested'].includes(entry.imageId)), false)
  assert.equal(documentImageEntries(model).length, 2)
})

test('every selected target is checked for edit and view permission even when an alignment leaves that particular object in place', () => {
  for (const action of ['edit', 'view']) {
    const model = document(); model.permissionRules = [{ id: 'deny-a', scope: 'paragraph', action, effect: 'deny', userIds: ['editor'], targetId: 'a' }]
    const before = copy(model)
    assert.throws(() => execute(model, { id: 'doc.command.align-images', members: members(model), align: '1' }), /permission denied/)
    assert.deepEqual(model, before)
  }
})

test('a locked later deletion cannot publish the removal of an earlier selected image', () => {
  const model = document()
  model.blocks.splice(2, 0, paragraph('lock-anchor', 'Locked'))
  createDocumentContentControl(model, { id: 'locked', type: 'richText', placement: 'block', blockRange: { blockIds: ['lock-anchor', 'b'] },
    blockId: 'lock-anchor', startOffset: 0, endOffset: 6, lock: 'controlLocked' })
  const before = copy(model)
  assert.throws(() => execute(model, { id: 'doc.command.batch-images', commands: ['a', 'b'].map(imageId => ({ id: 'doc.command.remove-image', imageId })) }), /锁定/)
  assert.deepEqual(model, before)
})

test('removing two atoms from one paragraph coordinates the live offsets once and preserves the surrounding text', () => {
  const model = document(), block = paragraph('p'), runs = []
  for (const [id, text] of [['a', 'L'], ['b', 'M']]) {
    runs.push(createRun(text)); const run = createRun(); setDocumentImageAtom(run, id, { src, alt: id, width: 40, height: 30 }); runs.push(run)
  }
  runs.push(createRun('R')); block.runs = runs; model.blocks = [block]
  const selected = { anchor: { blockId: 'p', offset: 5 }, focus: { blockId: 'p', offset: 5 } }
  const next = execute(model, { id: 'doc.command.batch-images', commands: ['a', 'b'].map(imageId => ({ id: 'doc.command.remove-image', imageId })) }, selected)
  assert.equal(runText(model.blocks[0].runs), 'LMR'); assert.equal(documentImageEntries(model).length, 0)
  assert.equal(next.focus.offset, 3); assert.equal(next.anchor.offset, 3)
})

test('selection captures include all object IDs and reject stale selection order, removed members and mismatched command targets', () => {
  const model = document(), ids = ['a', 'b'], session = { valid: true, generation: 7, modelRevision: 5, location: { kind: 'body', selection } }
  const target = captureDocumentImageCanvasTarget(model, 'b', selection, undefined, session, ids), current = { source: model, model, selection, session, requireSession: true, imageIds: ['a', 'b'] }
  ids.pop(); assert.deepEqual(target.imageIds, ['a', 'b']); assert.equal(documentImageCanvasTargetCurrent(target, current), true)
  assert.equal(documentImageCanvasTargetCurrent(target, { ...current, imageIds: ['b', 'a'] }), false)
  assert.equal(documentImageCanvasTargetCurrent(target, { ...current, model: { ...model, blocks: model.blocks.filter(block => block.id !== 'a') } }), false)
  assert.equal(documentImageCanvasCommandMatchesTarget({ target, command: { id: 'doc.command.remove-image', imageId: 'b' } }), false)
  assert.equal(documentImageCanvasCommandMatchesTarget({ target, command: { id: 'doc.command.batch-images', commands: ['a', 'b'].map(imageId => ({ id: 'doc.command.remove-image', imageId })) } }), true)
  assert.equal(documentImageCanvasCommandMatchesTarget({ target, command: { id: 'doc.command.arrange-images', imageIds: ['a', 'a'], arrange: 'front' } }), false)
})

test('multi-image clipboard projection carries all selected resources and alternate labels without copying unselected text or changing source IDs', () => {
  const model = document(), before = copy(model), data = documentImagesClipboardData(model, ['c', 'a'])
  assert.equal(data.text, 'c\na'); assert.equal((data.html.match(/<img\b/g) ?? []).length, 2)
  assert.equal(data.html.includes('Body'), false); assert.equal(data.html.includes('data-office-image='), true)
  assert.throws(() => documentImagesClipboardData(model, ['a', 'a'])); assert.throws(() => documentImagesClipboardData(model, ['missing']))
  assert.deepEqual(model, before)
})

test('the shared scene previews every member with the same local rotation and displays individual outlines plus one transform frame', () => {
  const model = document(), layout = layoutDocument(model, measurer), current = documentImageBatchMembers(model, layout, ['a', 'b', 'c']), before = copy(model)
  const updates = transformDocumentImageBatch(current, { x: 15, y: 20 }), viewport = { width: 1000, height: 1000, scrollLeft: 0, scrollTop: 0, zoom: 150 }
  const scene = sceneFromDocument(model, layout, viewport, { selection, focused: true, caretVisible: true, composing: false,
    images: new Map([[src, { complete: true, naturalWidth: 1, naturalHeight: 1 }]]), imageSelection: documentImageBatchBounds(current),
    imageSelections: updates.map(update => update.rectangle), imagePreviews: updates.map(update => ({ imageId: update.imageId, page: 0, rectangle: update.rectangle })) })
  const commands = scene.commands.filter(command => command.kind === 'image'), offset = sceneOffset(layout, viewport)
  assert.equal(commands.length, 3)
  for (const [index, command] of commands.entries()) {
    const rectangle = updates[index].rectangle
    near(command.x, offset.x + (rectangle.x + (rectangle.width - rectangle.paintWidth) / 2) * 1.5)
    near(command.y, offset.y + (rectangle.y + (rectangle.height - rectangle.paintHeight) / 2) * 1.5)
    near(command.rotation ?? 0, current[index].image.angle ?? 0)
  }
  assert.ok(scene.commands.filter(command => command.kind === 'line').length >= 17)
  assert.equal(scene.commands.some(command => command.kind === 'rect' && command.fill === '#24569c'), false)
  assert.deepEqual(model, before)
})

test('page-scoped previews do not move the same header drawing occurrence on another physical page', () => {
  const model = document(); model.blocks = [paragraph('p'), image('a', 50, 90)]
  const layout = layoutDocument(model, measurer), first = layout.images[0], page = layout.pages[0]
  const repeated = { ...first, page: 1, y: first.y + 800, repeatedHeader: true }
  layout.pages.push({ ...page, index: 1, y: page.y + 800, content: { ...page.content, y: page.content.y + 800 } })
  layout.images.push(repeated); layout.pageContent[1] = { ...layout.pageContent[0], images: [repeated], lines: [], cells: [], dividers: [], columnSeparators: [] }
  const viewport = { width: 1000, height: 1800, scrollLeft: 0, scrollTop: 0, zoom: 100 }, shift = sceneOffset(layout, viewport)
  const scene = sceneFromDocument(model, layout, viewport, { selection, focused: false, caretVisible: false, composing: false,
    images: new Map([[src, { complete: true, naturalWidth: 1, naturalHeight: 1 }]]), imagePreviews: [{ imageId: 'a', page: 0, rectangle: { ...first, x: first.x + 25 } }] })
  const commands = scene.commands.filter(command => command.kind === 'image')
  assert.equal(commands.length, 2); near(commands[0].x, shift.x + first.x + 25); near(commands[1].x, shift.x + repeated.x)
})
