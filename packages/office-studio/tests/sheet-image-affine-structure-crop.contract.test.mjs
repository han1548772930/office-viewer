import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffinePoint } from '../src/domain/drawing-affine.ts'
import { transformSheetImagesForAxisEdit, transformSheetImagesForCellShift,
  transformSheetImagesForGeometryChange } from '../src/features/sheet-images/sheet-image-structure.ts'
import { sheetImageFrame, sheetImageFootprint, sheetImageWorldFrameDelta,
  patchSheetFloatingImage, sheetImageWithDrawingGeometry } from '../src/features/sheet-images/sheet-image-placement.ts'
import { cropSheetImageToRatio } from '../src/features/sheet-images/sheet-image-crop.ts'
import { applySheetImageEditorCropRatio, createSheetImageEditorDraft, resizeSheetImageEditorDraft,
  sheetImagePatchFromDraft } from '../src/features/sheet-images/sheet-image-editor-state.ts'
import { dragSheetImageCrop, sheetImageUncroppedFrame, sheetImageUncroppedDrawingGeometry } from '../src/features/sheet-images/sheet-image-crop-session.ts'
import { moveSheetImage } from '../src/features/sheet-images/sheet-image-layout.ts'
import { alignSheetImageSelection, transformSheetImageSelection } from '../src/features/sheet-images/sheet-image-selection.ts'
import { sheetImageAffine, sheetImageAt, sheetImageCopy as copy, sheetImageTestAxes } from './fixtures/sheet-image-fixture.mjs'

// Deferred author/geometry contracts. No browser, test or benchmark has been executed.
const dimensions = { rowCount: 64, columnCount: 64 }
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} != ${expected}`)
function sameDirections(after, before) {
  for (const key of ['a', 'b', 'c', 'd']) near(after.drawingGeometry.matrix[key], before.drawingGeometry.matrix[key])
}
function samePoint(after, before) { near(after.x, before.x); near(after.y, before.y) }
const world = (image, x, y) => drawingAffinePoint(image.drawingGeometry.matrix, x, y)

test('inserting rows before an affine Position drawing translates e/f and leaves local geometry and author crop unchanged', () => {
  const axes = sheetImageTestAxes(), image = sheetImageAffine('position', axes), original = copy(image)
  const after = transformSheetImagesForAxisEdit([image], { kind: 'insert', axis: 'row', start: 3, count: 2 }, dimensions,
    { before: axes, after: axes })[0]
  sameDirections(after, image); near(after.drawingGeometry.matrix.f, image.drawingGeometry.matrix.f + 40)
  near(after.drawingGeometry.matrix.e, image.drawingGeometry.matrix.e)
  assert.equal(after.drawingGeometry.width, image.drawingGeometry.width); assert.deepEqual(after.crop, image.crop)
  assert.equal(after.anchor.row, image.anchor.row + 2); assert.deepEqual(image, original)
})

test('Both row insertion inside the two markers expands only the projected local height without discarding shear or reflection', () => {
  const axes = sheetImageTestAxes(), image = sheetImageAffine('both', axes, { anchorType: '1' }), before = sheetImageFrame(image, axes)
  const after = transformSheetImagesForAxisEdit([image], { kind: 'insert', axis: 'row', start: 5, count: 2 }, dimensions,
    { before: axes, after: axes })[0], frame = sheetImageFrame(after, axes)
  near(frame.top, before.top); near(frame.height, before.height + 40); near(frame.width, before.width)
  sameDirections(after, image); assert.deepEqual(after.size, image.size); assert.deepEqual(after.crop, image.crop)
  assert.equal(after.anchor.row, image.anchor.row); assert.equal(after.endAnchor.row, image.endAnchor.row + 2)
})

test('deleting all Both markers removes the drawing while a None drawing stays absolute', () => {
  const axes = sheetImageTestAxes(), both = sheetImageAffine('both', axes, { anchorType: '1' }), absolute = sheetImageAffine('absolute', axes, { anchorType: '2' }),
    start = both.anchor.row, count = both.endAnchor.row - start + 1
  const after = transformSheetImagesForAxisEdit([both, absolute], { kind: 'delete', axis: 'row', start, count }, dimensions,
    { before: axes, after: axes })
  assert.equal(after.length, 1); assert.deepEqual(after[0], absolute)
})

test('a wholly contained affine Both drawing follows a moved axis block while Position and None retain their markers', () => {
  const axes = sheetImageTestAxes(), both = sheetImageAffine('both', axes, { anchorType: '1' }), position = sheetImageAffine('position', axes), absolute = sheetImageAffine('absolute', axes, { anchorType: '2' })
  const after = transformSheetImagesForAxisEdit([both, position, absolute], { kind: 'move', axis: 'row', start: 3, count: 5, destination: 12 }, dimensions,
    { before: axes, after: axes })
  assert.equal(after[0].anchor.row, both.anchor.row + 4); near(after[0].drawingGeometry.matrix.f, both.drawingGeometry.matrix.f + 80)
  sameDirections(after[0], both); assert.deepEqual(after[1], position); assert.deepEqual(after[2], absolute)
})

test('axis copy assigns fresh identities and reprojects each affine copy without mutating the originals', () => {
  const axes = sheetImageTestAxes(), images = [sheetImageAffine('both', axes, { anchorType: '1' }), sheetImageAffine('absolute', axes, { anchorType: '2' })], original = copy(images)
  const after = transformSheetImagesForAxisEdit(images, { kind: 'copy', axis: 'row', start: 3, count: 5, destination: 12 }, dimensions,
    { before: axes, after: axes })
  assert.equal(after.length, 4); assert.equal(new Set(after.map(image => image.id)).size, 4)
  for (const [source, duplicate] of [[images[0], after[1]], [images[1], after[3]]]) {
    sameDirections(duplicate, source); near(duplicate.drawingGeometry.matrix.f, source.drawingGeometry.matrix.f + 180)
    assert.deepEqual(duplicate.crop, source.crop); assert.deepEqual(duplicate.resource, source.resource)
  }
  assert.deepEqual(images, original)
})

test('cell insertion reprojects an affine drawing contained in the perpendicular band and preserves a crossing drawing', () => {
  const axes = sheetImageTestAxes(), contained = sheetImageAffine('inside', axes, { anchorType: '1' }), crossing = sheetImageAffine('crossing', axes),
    edit = { kind: 'insert', axis: 'row', start: 5, count: 2, bandStart: 1, bandEnd: 2 }
  crossing.anchor.column = 0; crossing.drawingGeometry.matrix.e -= 64
  const after = transformSheetImagesForCellShift([contained, crossing], edit, dimensions, { before: axes, after: axes })
  near(sheetImageFrame(after[0], axes).height, sheetImageFrame(contained, axes).height + 40)
  sameDirections(after[0], contained); assert.deepEqual(after[1], crossing)
})

test('dimension changes use actual before/after axis offsets and Both marker spans', () => {
  const beforeAxes = sheetImageTestAxes(), afterAxes = sheetImageTestAxes([20, 20, 50, 20, 35, 35, 35]),
    position = sheetImageAffine('position', beforeAxes), both = sheetImageAffine('both', beforeAxes, { anchorType: '1' }),
    after = transformSheetImagesForGeometryChange([position, both], { before: beforeAxes, after: afterAxes })
  near(after[0].drawingGeometry.matrix.f, position.drawingGeometry.matrix.f + 30)
  const attached = copy(after[1]); delete attached.drawingGeometry
  const projected = sheetImageFrame(attached, afterAxes), actual = sheetImageFrame(after[1], afterAxes)
  for (const key of ['left', 'top', 'width', 'height']) near(actual[key], projected[key])
  sameDirections(after[1], both); assert.deepEqual(after[1].crop, both.crop)
})

test('the shared drawing host translation preserves a Both author after its projected dimensions changed', () => {
  const beforeAxes = sheetImageTestAxes(), afterAxes = sheetImageTestAxes([20, 20, 50, 20, 35, 35, 35]),
    image = transformSheetImagesForGeometryChange([sheetImageAffine('host-both', beforeAxes, { anchorType: '1' })],
      { before: beforeAxes, after: afterAxes })[0], geometry = copy(image.drawingGeometry)
  assert.notEqual(sheetImageFrame(image, afterAxes).height, image.size.height)
  geometry.matrix.e += 25; geometry.matrix.f += 18
  const translated = sheetImageWithDrawingGeometry(image, geometry, afterAxes), patched = patchSheetFloatingImage(image, { drawingGeometry: geometry }, afterAxes)
  for (const after of [translated, patched]) {
    assert.deepEqual(after.size, image.size); assert.deepEqual(after.crop, image.crop); assert.equal(after.rotation, image.rotation)
    assert.deepEqual(after.drawingGeometry, geometry)
    near(sheetImageFrame(after, afterAxes).left, sheetImageFrame(image, afterAxes).left + 25)
    near(sheetImageFrame(after, afterAxes).top, sheetImageFrame(image, afterAxes).top + 18)
  }
})

test('an explicit affine size patch remains a real resize instead of taking the author-preserving translation branch', () => {
  const axes = sheetImageTestAxes(), image = sheetImageAffine('geometry-resize', axes), geometry = copy(image.drawingGeometry)
  geometry.width *= 1.5
  const after = patchSheetFloatingImage(image, { drawingGeometry: geometry }, axes), frame = sheetImageFrame(after, axes)
  near(frame.width, sheetImageFrame(image, axes).width * 1.5); near(frame.height, sheetImageFrame(image, axes).height)
  near(after.size.width, image.size.width * 1.5); near(after.size.height, image.size.height)
  near(after.crop.left, image.crop.left * 1.5); near(after.crop.right, image.crop.right * 1.5)
  sameDirections(after, image); assert.deepEqual(after.drawingGeometry, geometry)
})

test('collapsed Both visibility retains the reversible author matrix and unhide restores the marker frame', () => {
  const axes = sheetImageTestAxes(), image = sheetImageAffine('both', axes, { anchorType: '1' }),
    hidden = sheetImageTestAxes([], [], [image.anchor.row, image.anchor.row + 1, image.endAnchor.row]), original = copy(image),
    collapsed = transformSheetImagesForGeometryChange([image], { before: axes, after: hidden })[0]
  assert.deepEqual(collapsed.drawingGeometry, original.drawingGeometry); assert.equal(sheetImageFrame(collapsed, hidden).height, 0)
  assert.deepEqual([sheetImageFootprint(collapsed, hidden).width, sheetImageFootprint(collapsed, hidden).height], [0, 0])
  assert.throws(() => cropSheetImageToRatio(collapsed, hidden, 1), /零尺寸|隐藏/)
  const restored = transformSheetImagesForGeometryChange([collapsed], { before: hidden, after: axes })[0]
  const frame = sheetImageFrame(restored, axes), before = sheetImageFrame(original, axes)
  for (const key of ['left', 'top', 'width', 'height']) near(frame[key], before[key])
  sameDirections(restored, original); assert.deepEqual(restored.crop, original.crop)
})

test('ratio crop preserves the complete uncropped source frame under affine shear and both source flips', () => {
  const axes = sheetImageTestAxes(), image = sheetImageAffine('crop', axes), before = sheetImageUncroppedFrame(image, axes),
    geometry = sheetImageUncroppedDrawingGeometry(image, axes), after = cropSheetImageToRatio(image, axes, 16 / 9),
    source = sheetImageUncroppedFrame(after, axes), afterGeometry = sheetImageUncroppedDrawingGeometry(after, axes)
  near(sheetImageFrame(after, axes).width / sheetImageFrame(after, axes).height, 16 / 9)
  for (const key of ['left', 'top', 'width', 'height']) near(source[key], before[key])
  near(afterGeometry.width, geometry.width); near(afterGeometry.height, geometry.height)
  for (const key of ['a', 'b', 'c', 'd', 'e', 'f']) near(afterGeometry.matrix[key], geometry.matrix[key])
  assert.equal(after.flipX, true); assert.equal(after.flipY, true)
})

test('an image property draft composes resize and successive crop ratios through the shared affine author path', () => {
  const axes = sheetImageTestAxes(), image = sheetImageAffine('editor-crop', axes), original = copy(image),
    draft = createSheetImageEditorDraft(image, 0, 0, axes)
  resizeSheetImageEditorDraft(draft, 'width', draft.width + 12)
  applySheetImageEditorCropRatio(draft, 16 / 9)
  const intermediate = patchSheetFloatingImage(image, sheetImagePatchFromDraft(draft), axes),
    source = sheetImageUncroppedDrawingGeometry(intermediate, axes)
  applySheetImageEditorCropRatio(draft, 1)
  const after = patchSheetFloatingImage(image, sheetImagePatchFromDraft(draft), axes),
    finalSource = sheetImageUncroppedDrawingGeometry(after, axes), frame = sheetImageFrame(after, axes)
  near(frame.width / frame.height, 1); near(frame.width, draft.width); near(frame.height, draft.height)
  for (const key of ['a', 'b', 'c', 'd', 'e', 'f']) near(finalSource.matrix[key], source.matrix[key])
  near(finalSource.width, source.width); near(finalSource.height, source.height)
  assert.equal(after.flipX, image.flipX); assert.equal(after.flipY, image.flipY); assert.deepEqual(image, original)
})

test('an affine west crop handle keeps the opposite edge and complete source geometry fixed', () => {
  const axes = sheetImageTestAxes(), image = sheetImageAffine('crop-drag', axes), delta = sheetImageWorldFrameDelta(image, 8, 0),
    source = sheetImageUncroppedDrawingGeometry(image, axes), after = dragSheetImageCrop(image, axes, delta.x, delta.y, 'w'),
    nextSource = sheetImageUncroppedDrawingGeometry(after, axes)
  samePoint(world(after, after.drawingGeometry.width, 0), world(image, image.drawingGeometry.width, 0))
  samePoint(world(after, after.drawingGeometry.width, after.drawingGeometry.height), world(image, image.drawingGeometry.width, image.drawingGeometry.height))
  for (const key of ['a', 'b', 'c', 'd', 'e', 'f']) near(nextSource.matrix[key], source.matrix[key])
  near(nextSource.width, source.width); near(nextSource.height, source.height)
})

test('affine resize uses displayed local directions so the reflected opposite edge remains fixed', () => {
  const axes = sheetImageTestAxes(), image = sheetImageAffine('resize', axes), delta = sheetImageWorldFrameDelta(image, 12, 0),
    after = moveSheetImage(axes, image, delta.x, delta.y, 'e')
  samePoint(world(after, 0, 0), world(image, 0, 0))
  samePoint(world(after, 0, after.drawingGeometry.height), world(image, 0, image.drawingGeometry.height))
  sameDirections(after, image); near(sheetImageFrame(after, axes).width, sheetImageFrame(image, axes).width + 12)
})

test('affine rotation preserves the world center and rejects a singular matrix patch before changing the author', () => {
  const axes = sheetImageTestAxes(), image = sheetImageAffine('rotate', axes), original = copy(image),
    after = patchSheetFloatingImage(image, { rotation: image.rotation + 45 }, axes)
  samePoint(world(after, after.drawingGeometry.width / 2, after.drawingGeometry.height / 2), world(image, image.drawingGeometry.width / 2, image.drawingGeometry.height / 2))
  assert.throws(() => patchSheetFloatingImage(image, { drawingGeometry: { ...image.drawingGeometry, matrix: { a: 1, b: 0, c: 1, d: 0, e: 0, f: 0 } } }, axes), /不可逆/)
  assert.deepEqual(image, original)
})

test('selection alignment uses frame edges while preserving each affine author and order', () => {
  const axes = sheetImageTestAxes(), first = sheetImageAffine('first', axes), second = sheetImageAt('second', axes, { left: 260, top: 140, width: 50, height: 30 }),
    original = copy([first, second]), after = alignSheetImageSelection([first, second], axes, '1')
  assert.deepEqual(after.map(image => image.id), ['first', 'second'])
  near(sheetImageFrame(after[0], axes).left, sheetImageFrame(after[1], axes).left)
  sameDirections(after[0], first); assert.deepEqual(after[0].crop, first.crop); assert.deepEqual([first, second], original)
})

test('rigid multi-image movement preserves relative positions and every affine corner delta', () => {
  const axes = sheetImageTestAxes(), first = sheetImageAffine('first', axes), second = sheetImageAffine('second', axes),
    after = transformSheetImageSelection([first, second], axes, 25, 18)
  for (let index = 0; index < after.length; index++) {
    const before = [first, second][index]; sameDirections(after[index], before)
    const oldCorner = world(before, 0, 0), corner = world(after[index], 0, 0)
    near(corner.x - oldCorner.x, 25); near(corner.y - oldCorner.y, 18); assert.deepEqual(after[index].crop, before.crop)
  }
})
