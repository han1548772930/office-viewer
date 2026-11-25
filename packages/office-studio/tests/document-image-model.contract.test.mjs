import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createTable, findDocumentBlock } from '../src/features/document/model/document-model.ts'
import { arrangeDocumentImageOrder, documentImageCropFromSourceRect, documentImageEntries, documentImageSourceRectFromCrop,
  normalizeDocumentImage, normalizeDocumentImageCrop, normalizeDocumentImagePosition, patchDocumentImage } from '../src/features/document/images/document-image-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { assertDocumentImageEditable } from '../src/features/document/commands/document-image-commands.ts'
import { createDocumentContentControl, replacePictureContentControlImage } from '../src/features/document/services/content-controls.ts'

// Deferred DOC-07 contracts. Written only; this batch has not executed them.
const copy = value => structuredClone(value)
const source = 'data:image/png;base64,AA=='
const image = patch => ({ src: source, alt: 'Image', width: 160, height: 80, intrinsicWidth: 400, intrinsicHeight: 200, ...patch })
const paragraph = (id, text = '') => ({ ...createParagraph(text), id })
const picture = (id, patch = {}) => ({ ...paragraph(id), type: 'image', runs: [], metadata: { image: image(patch) } })
const document = blocks => ({ ...createBlankDocument(), id: 'image-contract', blocks: [paragraph('p', 'Text'), ...blocks, paragraph('tail', 'Tail')] })
const selected = { anchor: { blockId: 'p', offset: 0 }, focus: { blockId: 'p', offset: 0 } }
const permission = { documentId: 'image-contract', userId: 'editor' }
const execute = (model, command) => executeDocumentCommand(model, selected, command, permission)
const data = (model, id) => findDocumentBlock(model, id).metadata.image
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8, `${actual} != ${expected}`)
const order = model => documentImageEntries(model).sort((a, b) => (a.image.layer ?? 0) - (b.image.layer ?? 0)).map(entry => entry.block.id)

test('Univer CSS-pixel srcRect offsets use the visible frame and round-trip signed outcropping', () => {
  const crop = documentImageCropFromSourceRect({ left: 20, right: 20, top: 10, bottom: 10 }, 160, 80)
  assert.deepEqual(crop, { top: 0.1, right: 0.1, bottom: 0.1, left: 0.1 })
  assert.deepEqual(documentImageSourceRectFromCrop(crop, 160, 80), { top: 10, right: 20, bottom: 10, left: 20 })
  const outside = documentImageCropFromSourceRect({ left: -20, right: 0 }, 160, 80)
  near(outside.left, -20 / 140); near(documentImageSourceRectFromCrop(outside, 160, 80).left, -20)
  assert.deepEqual(documentImageCropFromSourceRect(null, 160, 80), { top: 0, right: 0, bottom: 0, left: 0 })
  assert.equal(documentImageSourceRectFromCrop({ top: 0, right: 0, bottom: 0, left: 0 }, 160, 80), null)
})

test('invalid crop, source-frame, position and resource budgets reject without changing their inputs', () => {
  for (const crop of [{ top: 0, right: 0.8, bottom: 0, left: 0.2 }, { top: NaN, right: 0, bottom: 0, left: 0 },
    { top: 0, right: 0, bottom: 0, left: -1001 }, { top: 0, right: 0, bottom: 0, left: 0, unknown: 1 }]) assert.throws(() => normalizeDocumentImageCrop(crop))
  assert.throws(() => documentImageCropFromSourceRect({ left: -160 }, 160, 80))
  assert.throws(() => documentImageCropFromSourceRect({ l: 20 }, 160, 80))
  assert.throws(() => normalizeDocumentImagePosition({ relativeFrom: 0, posOffset: 1, align: 2 }))
  assert.throws(() => normalizeDocumentImagePosition({ relativeFrom: 8 }))
  for (const patch of [{ width: Infinity }, { height: 0 }, { src: 'javascript:bad' }, { alt: 'x'.repeat(32769) },
    { byteLength: 5 * 1024 * 1024 + 1 }, { layer: 1000001 }]) assert.throws(() => normalizeDocumentImage(image(patch)))
})

test('single-axis resize honors the current ratio while explicit dimensions and unlocked resize remain authoritative', () => {
  const original = image(), before = copy(original)
  assert.deepEqual([patchDocumentImage(original, { width: 320 }).width, patchDocumentImage(original, { width: 320 }).height], [320, 160])
  assert.equal(patchDocumentImage(original, { height: 40 }).width, 80)
  assert.equal(patchDocumentImage(original, { width: 320, keepAspectRatio: false }).height, 80)
  assert.equal(patchDocumentImage(original, { width: 320, height: 90 }).height, 90)
  assert.equal(patchDocumentImage(original, { angle: -30 }).angle, 330)
  assert.throws(() => patchDocumentImage(original, { src: 'https://example.test/other.png' }), /异步资源读取/)
  assert.deepEqual(original, before)
})

test('image discovery traverses decorated blocks, columns, cell contents and merged continuation contents', () => {
  const table = createTable(2, 1), cell = table.rows[0].cells[0]
  cell.rowSpan = 2; cell.blocks = [picture('cell')]; table.rows[1].cells = []
  cell.style.nativeContinuationContent = { 1: { version: 1, ownerCellId: cell.id, rowId: table.rows[1].id, colSpan: 1, blocks: [picture('continued')] } }
  const model = document([{ ...paragraph('quote'), type: 'quote', runs: [], children: [picture('quoted')] },
    { ...paragraph('columns'), type: 'columns', runs: [], columns: { columns: [{ id: 'left', blocks: [picture('column')] }] } },
    { ...paragraph('table'), type: 'table', runs: [], metadata: { table } }])
  assert.deepEqual(new Set(documentImageEntries(model).map(entry => entry.block.id)), new Set(['quoted', 'column', 'cell', 'continued']))
  execute(model, { id: 'doc.command.update-image', imageId: 'continued', patch: { alt: 'Continuation' } })
  assert.equal(data(model, 'continued').alt, 'Continuation')
})

test('update, replacement, reset and removal preserve the visible box, replace resource identity and keep legal text anchors', () => {
  const model = document([picture('a', { angle: 20, layer: 3, wrappingStyle: 'wrapSquare', crop: { top: 0.1, right: 0, bottom: 0, left: 0 },
    positionH: { relativeFrom: 0, posOffset: 12 } })])
  execute(model, { id: 'doc.command.update-image', imageId: 'a', patch: { width: 240 } })
  assert.deepEqual([data(model, 'a').width, data(model, 'a').height], [240, 120])
  execute(model, { id: 'doc.command.replace-image', imageId: 'a', image: image({ src: 'data:image/jpeg;base64,AA==', intrinsicWidth: 600, intrinsicHeight: 300, resourceId: 'new' }) })
  assert.deepEqual([data(model, 'a').width, data(model, 'a').height, data(model, 'a').angle, data(model, 'a').layer], [240, 120, 20, 3])
  assert.equal(data(model, 'a').resourceId, 'new'); assert.equal(data(model, 'a').positionH.posOffset, 12)
  assert.deepEqual(data(model, 'a').crop, { top: 0, right: 0, bottom: 0, left: 0 })
  execute(model, { id: 'doc.command.reset-image-size', imageId: 'a' })
  assert.deepEqual([data(model, 'a').width, data(model, 'a').height, data(model, 'a').angle], [600, 300, 0])
  assert.deepEqual(execute(model, { id: 'doc.command.remove-image', imageId: 'a' }), selected)
  assert.equal(documentImageEntries(model).length, 0)
})

test('an invalid later batch item cannot publish an earlier mutation or bypass source preparation', () => {
  const model = document([picture('a'), picture('b')]), before = copy(model)
  assert.throws(() => execute(model, { id: 'doc.command.batch-images', commands: [
    { id: 'doc.command.update-image', imageId: 'a', patch: { angle: 90 } },
    { id: 'doc.command.update-image', imageId: 'b', patch: { width: -1 } },
  ] }))
  assert.deepEqual(model, before)
  assert.throws(() => execute(model, { id: 'doc.command.batch-images', commands: [{ id: 'doc.command.update-image', imageId: 'a', patch: { src: source } }] }), /异步资源读取/)
  assert.deepEqual(model, before)
})

test('adjacent and absolute arrangements preserve selected order and follow the upstream zero-based clamp', () => {
  const ids = ['a', 'b', 'c', 'd', 'e']
  assert.deepEqual(arrangeDocumentImageOrder(ids, ['c', 'a'], 'forward'), ['b', 'a', 'd', 'c', 'e'])
  assert.deepEqual(arrangeDocumentImageOrder(ids, ['c', 'a'], 'backward'), ['a', 'c', 'b', 'd', 'e'])
  assert.deepEqual(arrangeDocumentImageOrder(ids, ['c', 'a'], 'front'), ['b', 'd', 'e', 'a', 'c'])
  assert.deepEqual(arrangeDocumentImageOrder(ids, ['c', 'a'], 'back'), ['a', 'c', 'b', 'd', 'e'])
  assert.deepEqual(arrangeDocumentImageOrder(ids, ['e'], 1.9), ['a', 'e', 'b', 'c', 'd'])
  assert.equal(arrangeDocumentImageOrder(ids, ['c'], -1)[0], 'c')
  assert.equal(arrangeDocumentImageOrder(ids, ['c'], Infinity).at(-1), 'c')
  assert.equal(arrangeDocumentImageOrder(ids, ['c'], NaN)[0], 'c')
  assert.throws(() => arrangeDocumentImageOrder(ids, ['a', 'a'], 'front'))
  assert.throws(() => arrangeDocumentImageOrder(ids, ['a', 'b'], 1))
})

test('multi-image arrangement is atomic across changed peers and rejects unrelated containers', () => {
  const model = document(['a', 'b', 'c', 'd'].map((id, layer) => picture(id, { layer })))
  execute(model, { id: 'doc.command.arrange-images', imageIds: ['c', 'a'], arrange: 'front' })
  assert.deepEqual(order(model), ['b', 'd', 'a', 'c'])
  const nested = picture('nested'); model.blocks.push({ ...paragraph('quote'), type: 'quote', runs: [], children: [nested] })
  const before = copy(model)
  assert.throws(() => execute(model, { id: 'doc.command.arrange-images', imageIds: ['a', 'nested'], arrange: 'front' }), /同一容器/)
  assert.deepEqual(model, before)
  model.permissionRules = [{ id: 'deny-peer', scope: 'paragraph', action: 'edit', effect: 'deny', userIds: ['editor'], targetId: 'd' }]
  const denied = copy(model)
  assert.throws(() => execute(model, { id: 'doc.command.arrange-images', imageIds: ['a'], arrange: 'back' }), /permission denied/)
  assert.deepEqual(model, denied)
})

test('picture content-control locks apply to the real object and changed inline mirrors preserve unrelated controls and anchors', () => {
  const model = { ...createBlankDocument(), id: 'image-contract', blocks: [paragraph('anchor', 'P'), paragraph('p', 'Text')] }
  createDocumentContentControl(model, { id: 'control', type: 'picture', blockId: 'anchor', startOffset: 0, endOffset: 1,
    placement: 'block', blockRange: { blockIds: ['anchor'] } })
  const id = replacePictureContentControlImage(model, 'control', image())
  for (const lock of ['contentLocked', 'sdtContentLocked']) {
    model.contentControls.find(control => control.id === 'control').lock = lock
    const before = copy(model)
    assert.throws(() => assertDocumentImageEditable(model, id), /锁定/)
    assert.throws(() => execute(model, { id: 'doc.command.update-image', imageId: id, patch: { alt: 'blocked' } }), /锁定/)
    assert.deepEqual(model, before)
  }
  model.contentControls.find(control => control.id === 'control').lock = 'controlLocked'
  execute(model, { id: 'doc.command.update-image', imageId: id, patch: { alt: 'Allowed' } })
  assert.equal(data(model, id).alt, 'Allowed')
  const before = copy(model)
  assert.throws(() => execute(model, { id: 'doc.command.remove-image', imageId: id }), /锁定/)
  assert.deepEqual(model, before)

  const inline = document([])
  createDocumentContentControl(inline, { id: 'inline-picture', type: 'picture', placement: 'inline', blockId: 'p', startOffset: 1, endOffset: 2,
    metadata: { nativePictureHint: { retained: true } } })
  const inlineId = replacePictureContentControlImage(inline, 'inline-picture', image())
  createDocumentContentControl(inline, { id: 'unrelated', type: 'text', blockId: 'tail', startOffset: 0, endOffset: 4,
    metadata: { retained: ['original'] } })
  const control = () => inline.contentControls.find(item => item.id === 'inline-picture')
  const range = () => ({ blockId: control().blockId, startOffset: control().startOffset, endOffset: control().endOffset })
  const anchor = range(), unrelated = copy(inline.contentControls.find(item => item.id === 'unrelated'))
  control().lock = 'contentLocked'
  assert.throws(() => assertDocumentImageEditable(inline, inlineId), /锁定/)
  control().lock = 'unlocked'
  execute(inline, { id: 'doc.command.update-image', imageId: inlineId, patch: { alt: 'Inline edited', width: 240 } })
  assert.deepEqual(control().metadata.pictureImage, data(inline, inlineId))
  assert.deepEqual(inline.contentControls.find(item => item.id === 'unrelated'), unrelated)
  execute(inline, { id: 'doc.command.remove-image', imageId: inlineId })
  assert.deepEqual(range(), anchor)
  assert.equal(control().metadata.pictureBlockId, undefined); assert.equal(control().metadata.pictureImage, undefined)
  assert.deepEqual(control().metadata.nativePictureHint, { retained: true })
  assert.equal(control().value, ''); assert.equal(control().showingPlaceholder, true)
  assert.deepEqual(inline.contentControls.find(item => item.id === 'unrelated'), unrelated)
})

test('explicit wrapping changes remove only the stale native layout discriminator', () => {
  const block = picture('a', { wrappingStyle: 'wrapSquare' })
  block.metadata.univerBodyDrawingSource = { layoutType: 4, polygonHint: { points: [1, 2] } }
  const model = document([block])
  execute(model, { id: 'doc.command.update-image', imageId: 'a', patch: { angle: 10 } })
  assert.equal(findDocumentBlock(model, 'a').metadata.univerBodyDrawingSource.layoutType, 4)
  execute(model, { id: 'doc.command.update-image', imageId: 'a', patch: { wrappingStyle: 'wrapTopAndBottom' } })
  assert.equal(findDocumentBlock(model, 'a').metadata.univerBodyDrawingSource.layoutType, undefined)
  assert.deepEqual(findDocumentBlock(model, 'a').metadata.univerBodyDrawingSource.polygonHint, { points: [1, 2] })
})
