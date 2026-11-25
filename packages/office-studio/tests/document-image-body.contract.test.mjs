import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { documentImageFromDrawing, documentImageToDrawing } from '../src/api/document-image-body.ts'
import { decodeDocumentBody, encodeDocumentBody } from '../src/api/document-body-codec.ts'
import { DOCUMENT_BODY_TOKENS as T } from '../src/api/document-body-types.ts'
import { documentImageEntries, setDocumentImageEntry } from '../src/features/document/images/document-image-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'

// Deferred source-preservation and public save/reopen drafts; no Office sample was opened.
const copy = value => structuredClone(value)
const source = 'data:image/png;base64,AA=='
function drawing(id = 'image') {
  return { drawingId: id, drawingType: 0, source, imageSourceType: 'BASE64', description: 'Original', authorHint: 'drawing',
    layoutType: 4, wrapText: 2, distT: 4, distR: 5, distB: 6, distL: 7,
    srcRect: { left: 20, right: 20, top: 10, bottom: 10, authorHint: 'crop' },
    docTransform: { authorHint: 'transform', size: { width: 160, height: 80, authorHint: 'size' }, angle: 15,
      positionH: { relativeFrom: 0, align: 2, authorHint: 'horizontal' }, positionV: { relativeFrom: 3, percent: 10, authorHint: 'vertical' } },
    officeStudioImage: { authorHint: 'extension', intrinsicWidth: 400, intrinsicHeight: 200, resourceId: 'media', keepAspectRatio: true } }
}
function bodySource(drawings = { image: drawing() }, drawingsOrder = Object.keys(drawings)) {
  const ids = Object.keys(drawings), dataStream = ids.map(() => T.customBlock + T.paragraph).join('') + T.section
  return { body: { dataStream, customBlocks: ids.map((blockId, index) => ({ blockId, startIndex: index * 2, authorHint: 'block-' + blockId })),
    paragraphs: ids.map((_id, index) => ({ paragraphId: 'p-' + index, startIndex: index * 2 + 1 })), sectionBreaks: [{ startIndex: dataStream.length - 1, sectionId: 'first' }] },
    drawings, drawingsOrder }
}
const resources = data => ({ drawings: data.drawings, drawingsOrder: data.drawingsOrder })
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8, `${actual} != ${expected}`)

test('native pixel crop imports as fractions without changing visible size and unchanged export retains all source hints', () => {
  const raw = drawing(), before = copy(raw), image = documentImageFromDrawing(raw, source, 160, 80, 2)
  assert.deepEqual([image.width, image.height, image.angle, image.wrappingStyle, image.wrapSide, image.layer], [160, 80, 15, 'wrapSquare', 'right', 2])
  assert.deepEqual(image.crop, { top: 0.1, right: 0.1, bottom: 0.1, left: 0.1 })
  const encoded = documentImageToDrawing(image, raw)
  assert.deepEqual(encoded.srcRect, raw.srcRect); assert.equal(encoded.layoutType, 4)
  assert.equal(encoded.authorHint, 'drawing'); assert.equal(encoded.docTransform.authorHint, 'transform')
  assert.equal(encoded.docTransform.size.authorHint, 'size'); assert.equal(encoded.officeStudioImage.authorHint, 'extension')
  assert.equal(encoded.docTransform.positionH.authorHint, 'horizontal'); assert.deepEqual(raw, before)
})

test('resized visible frames recalculate native crop pixels and changed positioning drops only obsolete mode values', () => {
  const raw = drawing(), image = documentImageFromDrawing(raw, source, 160, 80, 0)
  image.width = 320; image.height = 160; image.positionH = { relativeFrom: 3, posOffset: 12 }; image.positionV = { relativeFrom: 0, align: 4 }
  const encoded = documentImageToDrawing(image, raw)
  near(encoded.srcRect.left, 40); near(encoded.srcRect.top, 20); assert.equal(encoded.srcRect.authorHint, 'crop')
  assert.deepEqual(encoded.docTransform.positionH, { relativeFrom: 3, posOffset: 12, authorHint: 'horizontal' })
  assert.deepEqual(encoded.docTransform.positionV, { relativeFrom: 0, align: 4, authorHint: 'vertical' })
  assert.deepEqual([encoded.docTransform.size.width, encoded.docTransform.size.height], [320, 160])
})

test('clearing crop emits a native null while signed exterior crop and local extension survive reopening', () => {
  const raw = drawing(), image = documentImageFromDrawing(raw, source, 160, 80, 0)
  image.crop = { top: 0, right: 0, bottom: 0, left: 0 }
  assert.equal(documentImageToDrawing(image, raw).srcRect, null)
  image.crop = { top: 0, right: 0, bottom: 0, left: -0.2 }
  const written = documentImageToDrawing(image, raw), restored = documentImageFromDrawing(written, source, 160, 80, 0)
  near(written.srcRect.left, -160 / 6); assert.deepEqual(restored.crop, image.crop)
  assert.equal(written.officeStudioImage.authorHint, 'extension')
})

test('unsupported crop properties remain retained and unmodified instead of being presented as implemented geometry', () => {
  const raw = drawing(); raw.srcRect = { l: 1000, polygonHint: [1, 2, 3] }
  const image = documentImageFromDrawing(raw, source, 160, 80, 0)
  assert.equal(image.crop, undefined)
  assert.deepEqual(documentImageToDrawing(image, raw).srcRect, raw.srcRect)
  raw.srcRect = { left: -160 }
  assert.equal(documentImageFromDrawing(raw, source, 160, 80, 0).crop, undefined)
  assert.deepEqual(documentImageToDrawing(documentImageFromDrawing(raw, source, 160, 80, 0), raw).srcRect, raw.srcRect)
})

test('public body reorders only image slots and retains an unknown drawing at its original stacking slot', () => {
  const original = bodySource({ a: drawing('a'), shape: { drawingId: 'shape', drawingType: 1, source: 'shape-source', authorHint: 'unmodeled' }, b: drawing('b') }, ['a', 'shape', 'b'])
  const decoded = decodeDocumentBody(original.body, resources(original)), model = { ...createBlankDocument(), id: 'image-body', blocks: decoded.content.blocks,
    metadata: decoded.content.metadata }, selectionId = decoded.content.blocks.find(block => block.type === 'paragraph').id
  const selection = { anchor: { blockId: selectionId, offset: 0 }, focus: { blockId: selectionId, offset: 0 } }
  executeDocumentCommand(model, selection, { id: 'doc.command.arrange-images', imageIds: ['a'], arrange: 'front' }, { documentId: model.id, userId: 'editor' })
  const saved = encodeDocumentBody({ ...decoded.content, blocks: model.blocks })
  assert.deepEqual(saved.drawingsOrder, ['b', 'shape', 'a']); assert.equal(saved.drawings.shape.authorHint, 'unmodeled')
  const reopened = decodeDocumentBody(saved.body, resources(saved))
  assert.deepEqual(documentImageEntries({ blocks: reopened.content.blocks }).sort((a, b) => a.image.layer - b.image.layer).map(entry => entry.imageId), ['b', 'a'])
})

test('body save/reopen retains image identity, visible dimensions, native crop, resource metadata and deletes removed drawings', () => {
  const original = bodySource(), decoded = decodeDocumentBody(original.body, resources(original)), first = documentImageEntries({ blocks: decoded.content.blocks })[0]
  setDocumentImageEntry(first, { ...first.image, alt: 'Updated', angle: 45, width: 240, height: 120 })
  const saved = encodeDocumentBody(decoded.content), reopened = decodeDocumentBody(saved.body, resources(saved)), current = documentImageEntries({ blocks: reopened.content.blocks })[0]
  assert.equal(current.imageId, 'image'); assert.equal(current.image.alt, 'Updated')
  assert.deepEqual([current.image.width, current.image.height, current.image.angle, current.image.intrinsicWidth, current.image.resourceId], [240, 120, 45, 400, 'media'])
  assert.deepEqual(current.image.crop, first.image.crop); assert.equal(saved.drawings.image.srcRect.authorHint, 'crop')
  if (current.inline) current.block.runs.splice(current.block.runs.indexOf(current.inline.run), 1)
  else reopened.content.blocks.splice(reopened.content.blocks.indexOf(current.block), 1)
  const removed = encodeDocumentBody(reopened.content)
  assert.equal(removed.drawings?.image, undefined); assert.equal(removed.drawingsOrder?.includes('image') ?? false, false)
})
