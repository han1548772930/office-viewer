import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createTable } from '../src/features/document/model/document-model.ts'
import { documentImageBox, documentImageFrameCorners, documentImageObstacle, documentImageWrapBand } from '../src/features/document/layout/document-image-layout.ts'
import { layoutDocument } from '../src/features/document/layout/document-layout.ts'
import { measureDocumentBlockCollections } from '../src/features/document/layout/document-block-measurement.ts'
import { measureDocumentTable } from '../src/features/document/layout/document-table-measurement.ts'
import { captureDocumentImageCanvasTarget, documentImageAtPoint, documentImageCanvasTargetCurrent,
  documentImageGestureResult, documentImageResizeCorner } from '../src/features/document/controllers/document-image-canvas.ts'
import { sceneFromDocument, sceneOffset } from '../src/features/document/rendering/document-scene.ts'

// Deferred Canvas geometry, relative-host layout and scene contracts. No renderer was started.
const copy = value => structuredClone(value), near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} != ${expected}`)
const paragraph = (id, text = '') => {
  const block = { ...createParagraph(text), id }
  Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, spacingRule: 'exact', lineHeight: 20, wordWrap: true })
  return block
}
const image = patch => ({ src: 'data:image/png;base64,AA==', alt: 'Image', width: 100, height: 60, keepAspectRatio: true, ...patch })
const picture = (id, patch = {}) => ({ ...paragraph(id), type: 'image', runs: [], metadata: { image: image(patch) } })
const document = blocks => ({ ...createBlankDocument(), id: 'image-canvas', page: { width: 500, height: 600, marginLeft: 40, marginRight: 40, marginTop: 40, marginBottom: 40 }, blocks })
const selection = { anchor: { blockId: 'p', offset: 0 }, focus: { blockId: 'p', offset: 0 } }
const measurer = { measure: text => Array.from(text).length * 8 }
const frame = (data, x = 60, y = 80) => ({ blockId: 'image', src: data.src, alt: data.alt, ...documentImageBox(data, 420, 520), x, y, page: 0,
  angle: data.angle, wrappingStyle: data.wrappingStyle, objectImage: true })

test('crop uses source fractions while visible dimensions and rotated layout bounds remain independent', () => {
  const data = image({ crop: { top: 0.2, right: 0.3, bottom: 0.1, left: 0.1 }, angle: 90 })
  const box = documentImageBox(data, 420, 520)
  near(box.width, 60); near(box.height, 100); near(box.paintWidth, 100); near(box.paintHeight, 60)
  const large = image({ width: 1000, height: 600 })
  near(documentImageBox(large, 200, 100).paintHeight, 100)
  near(documentImageBox({ ...large, wrappingStyle: 'wrapSquare' }, 200, 100).paintWidth, 1000)
})

test('rotation hit testing excludes empty AABB corners and uses real transformed resize handles', () => {
  const data = image({ angle: 35, wrappingStyle: 'inFrontOfText' }), rectangle = frame(data)
  const model = document([picture('image', data), paragraph('p')]), layout = layoutDocument(model, measurer)
  layout.images = [rectangle]; layout.lines = []
  assert.equal(documentImageAtPoint(model, layout, rectangle.x + 1, rectangle.y + 1), undefined)
  assert.equal(documentImageAtPoint(model, layout, rectangle.x + rectangle.width / 2, rectangle.y + rectangle.height / 2).blockId, 'image')
  for (const corner of documentImageFrameCorners(rectangle)) assert.equal(documentImageResizeCorner(rectangle, corner.x, corner.y, 100), corner.corner)
  assert.equal(documentImageResizeCorner(rectangle, rectangle.x + rectangle.width / 2, rectangle.y + rectangle.height / 2, 100), undefined)
})

test('a rotated resize holds the opposite corner fixed and keeps model dimensions in the unrotated visible frame', () => {
  const data = image({ angle: 30, wrappingStyle: 'inFrontOfText' }), rectangle = frame(data), model = document([picture('image', data), paragraph('p')])
  const target = captureDocumentImageCanvasTarget(model, 'image', selection), start = documentImageFrameCorners(rectangle).find(point => point.corner === 'se')
  const gesture = { target, image: data, rectangle, start, pageOrigin: { x: 0, y: 0 }, corner: 'se', pointerId: 1, zoom: 100 }
  const before = documentImageFrameCorners(rectangle).find(point => point.corner === 'nw'), result = documentImageGestureResult(gesture, start.x + 40, start.y + 30)
  const after = documentImageFrameCorners(result.rectangle).find(point => point.corner === 'nw')
  near(after.x, before.x); near(after.y, before.y); near(result.patch.width / result.patch.height, data.width / data.height)
  assert.equal(result.patch.positionH.relativeFrom, 0); near(result.patch.positionH.posOffset, result.rectangle.x)
  assert.ok(result.patch.width > data.width); assert.equal(model.blocks[0].metadata.image.width, data.width)
  const unlocked = documentImageGestureResult(gesture, start.x + 40, start.y + 30, true)
  assert.ok(Math.abs(unlocked.patch.width / unlocked.patch.height - data.width / data.height) > 1e-4)
})

test('drag converts an inline picture to a page-positioned float while stale gesture owners are rejected', () => {
  const data = image(), rectangle = frame(data), model = document([picture('image', data), paragraph('p')])
  const session = { valid: true, generation: 2, modelRevision: 3, location: { kind: 'body', selection } }
  const target = captureDocumentImageCanvasTarget(model, 'image', selection, undefined, session), before = copy(model)
  const result = documentImageGestureResult({ target, image: data, rectangle, start: { x: 10, y: 20 }, pageOrigin: { x: 20, y: 30 }, pointerId: 1, zoom: 100 }, 35, 50)
  assert.equal(result.patch.wrappingStyle, 'inFrontOfText'); near(result.patch.positionH.posOffset, 65); near(result.patch.positionV.posOffset, 80)
  const current = { source: model, model, selection, session, requireSession: true }
  assert.equal(documentImageCanvasTargetCurrent(target, current), true)
  assert.equal(documentImageCanvasTargetCurrent(target, { ...current, source: copy(model) }), false)
  assert.equal(documentImageCanvasTargetCurrent(target, { ...current, session: { ...session, generation: 3 } }), false)
  assert.equal(documentImageCanvasTargetCurrent(target, { ...current, host: { kind: 'story', sectionId: 'first', story: 'header' } }), false)
  assert.equal(documentImageCanvasTargetCurrent(target, { ...current, disabled: true }), false)
  assert.deepEqual(model, before)
})

test('behind-text pictures give ordinary text clicks precedence and remain reachable with the explicit alternate hit mode', () => {
  const data = image({ wrappingStyle: 'behindText' }), rectangle = frame(data), model = document([picture('image', data), paragraph('p')]), layout = layoutDocument(model, measurer)
  layout.images = [rectangle]; layout.lines = [{ x: rectangle.x, y: rectangle.y, width: rectangle.width, height: 20 }]
  const x = rectangle.x + 20, y = rectangle.y + 10
  assert.equal(documentImageAtPoint(model, layout, x, y), undefined)
  assert.equal(documentImageAtPoint(model, layout, x, y, true).blockId, 'image')
})

test('wrapping distances, side selection and top/bottom obstacles determine a finite line band', () => {
  const data = image({ wrappingStyle: 'wrapSquare', wrapSide: 'right', distances: { top: 5, right: 10, bottom: 5, left: 10 } })
  const obstacle = documentImageObstacle(frame(data, 50, 0), data), area = { x: 0, y: 0, width: 300, height: 200 }
  const band = documentImageWrapBand(area, 0, 20, [obstacle])
  near(band.x, 160); near(band.width, 140)
  const below = documentImageWrapBand(area, 0, 20, [{ ...obstacle, wrappingStyle: 'wrapTopAndBottom' }])
  near(below.y, 65); near(below.width, 300)
})

test('relative table and auxiliary measurement rewraps only unplaced UTF-16 suffixes and expands to floated image bounds', () => {
  const text = '😀' + 'a'.repeat(90), blocks = [picture('float', { width: 120, height: 80, wrappingStyle: 'wrapSquare',
    positionH: { relativeFrom: 1, posOffset: 0 }, positionV: { relativeFrom: 1, posOffset: 0 } }), paragraph('p', text)]
  const measured = measureDocumentBlockCollections([blocks], { width: 240, markers: new Map(), measurer, measureTable: measureDocumentTable })
  const lines = measured.lines.filter(line => line.target.blockId === 'p')
  assert.ok(lines[0].x >= 120); assert.ok(lines.some(line => line.y >= 80 && line.x === 0))
  assert.equal(lines[0].start, 0); assert.equal(lines.at(-1).end, text.length)
  for (let index = 1; index < lines.length; index++) assert.equal(lines[index].start, lines[index - 1].end)
  assert.ok(lines.every(line => line.start !== 1 && line.end !== 1))
  assert.ok(Number.isFinite(measured.height)); assert.ok(measured.height >= 80)
  const finite = measureDocumentBlockCollections([[picture('float', { wrappingStyle: 'inFrontOfText', positionV: { relativeFrom: 0, align: 0 } })]],
    { width: 240, markers: new Map(), measurer, measureTable: measureDocumentTable })
  assert.ok(Number.isFinite(finite.images[0].y)); assert.ok(Number.isFinite(finite.height))
  const table = createTable(1, 1); table.rows[0].cells[0].blocks = blocks
  const nested = measureDocumentTable({ ...paragraph('table'), type: 'table', runs: [], metadata: { table } }, 240, new Map(), measurer)
  assert.ok(nested.lines.filter(line => line.target.cellBlockId === 'p').length > 1)
  assert.ok(nested.images[0].tableCellId); assert.ok(Number.isFinite(nested.height))
})

test('Canvas scene emits shared crop/rotation commands and bitmap/selection previews without modifying content', () => {
  const data = image({ angle: 25, crop: { top: 0.1, right: 0.2, bottom: 0.1, left: -0.1 }, wrappingStyle: 'inFrontOfText' })
  const model = document([paragraph('p'), picture('image', data)]), layout = layoutDocument(model, measurer), original = copy(model)
  const rectangle = layout.images.find(image => image.objectImage), preview = { ...rectangle, x: rectangle.x + 20, y: rectangle.y + 10 }
  const viewport = { width: 700, height: 700, scrollLeft: 0, scrollTop: 0, zoom: 150 }, bitmap = { complete: true, naturalWidth: 1, naturalHeight: 1 }
  const scene = sceneFromDocument(model, layout, viewport, { selection, focused: true, caretVisible: false, composing: false,
    images: new Map([[data.src, bitmap]]), imageSelection: preview, imagePreview: { imageId: 'image', rectangle: preview, wrappingStyle: 'inFrontOfText' } })
  const command = scene.commands.find(command => command.kind === 'image'), shift = sceneOffset(layout, viewport)
  assert.equal(command.rotation, 25); near(command.sourceRect.x, -0.1); near(command.sourceRect.y, 0.1)
  near(command.sourceRect.width, 0.9); near(command.sourceRect.height, 0.8)
  near(command.x, shift.x + (preview.x + (preview.width - preview.paintWidth) / 2) * 1.5)
  near(command.y, shift.y + (preview.y + (preview.height - preview.paintHeight) / 2) * 1.5)
  assert.ok(scene.commands.filter(command => command.kind === 'line').length >= 4)
  assert.deepEqual(model, original)
})
