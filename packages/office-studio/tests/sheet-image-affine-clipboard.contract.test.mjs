import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffineBounds, drawingAffineFrame } from '../src/domain/drawing-affine.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../src/rendering/spreadsheet/layout.ts'
import { sheetClipboardObjectAxes, sheetClipboardRangeBounds } from '../src/features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { captureSheetClipboard, SHEET_CLIPBOARD_SERIALIZED_BYTE_BUDGET } from '../src/features/sheet-workbench/sheet-clipboard-data.ts'
import { createSheetClipboardTargetGeometry } from '../src/features/sheet-workbench/sheet-clipboard-target.ts'
import { assertSheetImageClipboardEntryGeometry, captureSheetClipboardImages, captureSheetImageClipboardEntry,
  pasteSheetClipboardImages } from '../src/features/sheet-images/sheet-image-clipboard.ts'
import { assertSheetImageObjectClipboard, captureSheetImageObjects, planPasteSheetImageObjects, sheetImageObjectClipboardContent,
  sheetImageObjectClipboardFromHtml, sheetImageObjectPasteRanges } from '../src/features/sheet-images/sheet-image-object-clipboard.ts'
import { sheetImageFrame } from '../src/features/sheet-images/sheet-image-placement.ts'
import { transformSheetImagesForGeometryChange } from '../src/features/sheet-images/sheet-image-structure.ts'
import { sheetImageAffine, sheetImageAt, sheetImageCopy as copy, sheetImageFixture, sheetImagePng } from './fixtures/sheet-image-fixture.mjs'

// Written only. No test, browser or Office file has been run for this batch.
const range = { startRow: 0, endRow: 8, startCol: 0, endCol: 5 }
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} != ${expected}`)
function prepared() {
  const fixture = sheetImageFixture(); fixture.sheet.images = [sheetImageAffine('a', fixture.axes)]
  return fixture
}
function authorHtml(clipboard) {
  return `<div data-office-image-clipboard-id="${clipboard.id}" data-office-image-clipboard-model="${encodeURIComponent(JSON.stringify(clipboard))}"></div>`
}

test('affine capture retains shear, reflection, crop, flips and opaque metadata in a detached author', () => {
  const { workbook, sheet } = prepared(), image = sheet.images[0], before = copy(image)
  const entry = captureSheetImageClipboardEntry(workbook, sheet, image)
  assert.deepEqual(entry.image, before)
  const frame = drawingAffineFrame(image.drawingGeometry), bounds = drawingAffineBounds(image.drawingGeometry)
  assert.deepEqual(entry.frame, { ...frame, left: frame.left + GRID_HEADER_WIDTH, top: frame.top + GRID_HEADER_HEIGHT })
  assert.deepEqual(entry.bounds, { ...bounds, left: bounds.left + GRID_HEADER_WIDTH, top: bounds.top + GRID_HEADER_HEIGHT })
  entry.image.drawingGeometry.matrix.e++; entry.image.metadata.opaqueAuthor.retained.push('changed')
  assert.deepEqual(image, before)
})

test('affine clipboard rejects independently changed frames and four-corner bounds', () => {
  const { workbook, sheet } = prepared(), entry = captureSheetImageClipboardEntry(workbook, sheet, sheet.images[0])
  const wrongFrame = copy(entry); wrongFrame.frame.left += 1
  const wrongBounds = copy(entry); wrongBounds.bounds.width += 1
  assert.throws(() => assertSheetImageClipboardEntryGeometry(wrongFrame), /显示框/)
  assert.throws(() => assertSheetImageClipboardEntryGeometry(wrongBounds), /范围/)
  const oldFlavour = copy(entry); delete oldFlavour.frame
  assert.doesNotThrow(() => assertSheetImageClipboardEntryGeometry(oldFlavour))
})

test('legacy rotation clipboard keeps its frame when a previous clipboard omitted the optional frame field', () => {
  const { workbook, sheet, axes } = sheetImageFixture(), image = sheetImageAt('legacy', axes, undefined, { rotation: 33 })
  const entry = captureSheetImageClipboardEntry(workbook, sheet, image); delete entry.frame
  assert.doesNotThrow(() => assertSheetImageClipboardEntryGeometry(entry))
  assert.equal(entry.image.drawingGeometry, undefined)
})

test('object paste translates world origins once and preserves all local affine components on a different sheet', () => {
  const { workbook, sheet, destination, axes } = prepared()
  sheet.images.push(sheetImageAt('b', axes, { left: 260, top: 140, width: 50, height: 30 }, { rotation: -20 }))
  destination.rowHeights = { 0: 36, 1: 42 }; destination.columnWidths = { 0: 180 }
  const clipboard = captureSheetImageObjects(workbook, sheet, ['a', 'b'], 'editor'), before = copy(workbook),
    target = { startRow: 10, endRow: 10, startCol: 6, endCol: 6 }, origin = sheetClipboardRangeBounds(workbook, destination, target),
    plan = planPasteSheetImageObjects(workbook, destination, target, clipboard),
    deltaX = origin.left - clipboard.origin.left, deltaY = origin.top - clipboard.origin.top
  assert.deepEqual(workbook, before); assert.equal(plan.images.length, 2)
  assert.equal(new Set(plan.images.map(image => image.id)).size, 2)
  assert.ok(plan.images.every(image => !sheet.images.some(old => old.id === image.id)))
  const original = clipboard.entries[0].image, pasted = plan.images[0]
  for (const key of ['a', 'b', 'c', 'd']) assert.equal(pasted.drawingGeometry.matrix[key], original.drawingGeometry.matrix[key])
  assert.equal(pasted.drawingGeometry.width, original.drawingGeometry.width); assert.equal(pasted.drawingGeometry.height, original.drawingGeometry.height)
  near(pasted.drawingGeometry.matrix.e, original.drawingGeometry.matrix.e + deltaX)
  near(pasted.drawingGeometry.matrix.f, original.drawingGeometry.matrix.f + deltaY)
  for (const key of ['crop', 'flipX', 'flipY', 'metadata', 'resource']) assert.deepEqual(pasted[key], original[key])
  assert.deepEqual(plan.ranges, sheetImageObjectPasteRanges(workbook, destination, target, clipboard))
})

test('mutating a paste plan cannot alter the clipboard or source media metadata', () => {
  const { workbook, sheet, destination } = prepared(), clipboard = captureSheetImageObjects(workbook, sheet, ['a'], 'editor'), before = copy(clipboard)
  const plan = planPasteSheetImageObjects(workbook, destination, range, clipboard)
  plan.images[0].resource.name = 'different'; plan.images[0].crop.left = 99; plan.images[0].drawingGeometry.matrix.a = 2
  assert.deepEqual(clipboard, before); assert.equal(sheet.images[0].resource.name, 'sample.png')
})

test('copying a Both drawing after row geometry changes preserves the original size and crop author basis', () => {
  const { workbook, sheet, destination, axes } = sheetImageFixture()
  sheet.images = [sheetImageAffine('resized-both', axes, { anchorType: '1' })]
  const start = sheet.images[0].anchor.row, end = sheet.images[0].endAnchor.row
  sheet.rowHeights = Object.fromEntries(Array.from({ length: end - start + 1 }, (_, index) => [start + index, 48]))
  const afterAxes = sheetClipboardObjectAxes(workbook, sheet)
  sheet.images = transformSheetImagesForGeometryChange(sheet.images, { before: axes, after: afterAxes })
  const image = sheet.images[0], frame = sheetImageFrame(image, afterAxes)
  assert.notEqual(frame.height, image.size.height)
  const clipboard = captureSheetImageObjects(workbook, sheet, [image.id], 'editor'),
    plan = planPasteSheetImageObjects(workbook, destination, range, clipboard), pasted = plan.images[0]
  assert.deepEqual(pasted.size, image.size); assert.deepEqual(pasted.crop, image.crop)
  assert.equal(pasted.rotation, image.rotation)
  assert.equal(pasted.drawingGeometry.width, image.drawingGeometry.width)
  assert.equal(pasted.drawingGeometry.height, image.drawingGeometry.height)
  for (const key of ['a', 'b', 'c', 'd']) assert.equal(pasted.drawingGeometry.matrix[key], image.drawingGeometry.matrix[key])
  const restored = sheetImageObjectClipboardFromHtml(sheetImageObjectClipboardContent(clipboard).html)
  assert.deepEqual(restored.entries[0].image.size, image.size); assert.deepEqual(restored.entries[0].image.crop, image.crop)
})

test('internal HTML restores the complete author after broker disposal without duplicating base64 media', () => {
  const { workbook, sheet } = prepared(), clipboard = captureSheetImageObjects(workbook, sheet, ['a'], 'editor')
  const { html, plain } = sheetImageObjectClipboardContent(clipboard)
  assert.equal(plain, ''); assert.equal(html.split(sheetImagePng).length - 1, 1)
  const descriptor = JSON.parse(decodeURIComponent(/data-office-image-clipboard-model="([^"]+)"/.exec(html)[1]))
  assert.equal(descriptor.entries[0].image.resource.source, undefined)
  const recovered = sheetImageObjectClipboardFromHtml(html)
  assert.deepEqual(recovered, clipboard); recovered.entries[0].image.crop.left++
  assert.notDeepEqual(recovered.entries[0].image.crop, clipboard.entries[0].image.crop)
})

test('HTML media restores escaped attribute characters exactly once', () => {
  const { workbook, sheet } = prepared(), source = 'https://example.test/image.png?q=<value>&t="abc"&literal=&amp;'
  sheet.images[0].resource.source = source
  const clipboard = captureSheetImageObjects(workbook, sheet, ['a'], 'editor')
  assert.equal(sheetImageObjectClipboardFromHtml(sheetImageObjectClipboardContent(clipboard).html).entries[0].image.resource.source, source)
})

test('ordinary external IMG content remains on the external image preparation route', () => {
  assert.equal(sheetImageObjectClipboardFromHtml(`<img src="${sheetImagePng}">`), undefined)
  assert.equal(sheetImageObjectClipboardFromHtml('<table><tr><td>text</td></tr></table>'), undefined)
})

test('malformed author encoding and descriptor versions reject before a paste model is created', () => {
  assert.throws(() => sheetImageObjectClipboardFromHtml('<div data-office-image-clipboard-model="%GG"></div>'), /无法读取/)
  const { workbook, sheet } = prepared(), clipboard = captureSheetImageObjects(workbook, sheet, ['a'], 'editor')
  clipboard.version = 2; assert.throws(() => sheetImageObjectClipboardFromHtml(authorHtml(clipboard)), /版本/)
})

test('missing and duplicate media indexes cannot substitute sources in an internal descriptor', () => {
  const { workbook, sheet } = prepared(), { html } = sheetImageObjectClipboardContent(captureSheetImageObjects(workbook, sheet, ['a'], 'editor'))
  assert.throws(() => sheetImageObjectClipboardFromHtml(html.replace(/<img\b[^>]*>/, '')), /缺少实际媒体/)
  assert.throws(() => sheetImageObjectClipboardFromHtml(html.replace('</div>', `<img data-office-image-index="0" src="${sheetImagePng}"></div>`)), /媒体身份/)
  assert.throws(() => sheetImageObjectClipboardFromHtml(html.replace('data-office-image-index="0"', 'data-office-image-index="1"')), /媒体身份/)
})

test('author tokens, matrices and actual image bytes are validated instead of trusting an internal attribute name', () => {
  const { workbook, sheet } = prepared(), clipboard = captureSheetImageObjects(workbook, sheet, ['a'], 'editor'), { html } = sheetImageObjectClipboardContent(clipboard)
  assert.throws(() => sheetImageObjectClipboardFromHtml(html.replace(`data-office-image-clipboard-id="${clipboard.id}"`, 'data-office-image-clipboard-id="other"')), /标识/)
  assert.throws(() => sheetImageObjectClipboardFromHtml(html.replace(sheetImagePng, 'data:image/png;base64,AAAA')))
  clipboard.entries[0].image.drawingGeometry.matrix.d = -0.06
  assert.throws(() => sheetImageObjectClipboardFromHtml(authorHtml(clipboard)), /矩阵|范围|显示框/)
})

test('the previous complete-source author descriptor remains readable without indexed IMG tags', () => {
  const { workbook, sheet } = prepared(), clipboard = captureSheetImageObjects(workbook, sheet, ['a'], 'editor')
  assert.deepEqual(sheetImageObjectClipboardFromHtml(authorHtml(clipboard)), clipboard)
})

test('internal author batches reject duplicate identities and an over-limit object count before planning insertion', () => {
  const { workbook, sheet } = prepared(), clipboard = captureSheetImageObjects(workbook, sheet, ['a'], 'editor'), duplicate = copy(clipboard)
  duplicate.entries.push(copy(duplicate.entries[0]))
  assert.throws(() => assertSheetImageObjectClipboard(duplicate), /重复/)
  const tooMany = copy(clipboard); tooMany.entries = Array.from({ length: 1001 }, () => copy(clipboard.entries[0]))
  assert.throws(() => assertSheetImageObjectClipboard(tooMany), /批次数量/)
})

test('declared remote resources cannot bypass the 32 MiB author batch budget', () => {
  const { workbook, sheet, axes } = sheetImageFixture()
  sheet.images = Array.from({ length: 7 }, (_, index) => {
    const image = sheetImageAffine('remote-' + index, axes)
    image.resource.source = 'https://example.test/image-' + index + '.png'
    image.resource.byteLength = 5 * 1024 * 1024
    return image
  })
  assert.throws(() => captureSheetImageObjects(workbook, sheet, sheet.images.map(image => image.id), 'editor'), /32 MiB/)
})

test('an oversized HTML clipboard rejects before decoding its author descriptor or allocating media', () => {
  const html = '<div data-office-image-clipboard-model="%GG">' + ' '.repeat(SHEET_CLIPBOARD_SERIALIZED_BYTE_BUDGET) + '</div>'
  assert.throws(() => sheetImageObjectClipboardFromHtml(html), /内存预算/)
})

test('range copy excludes None drawings while explicit object copy supports their absolute placement', () => {
  const { workbook, sheet, axes } = prepared(); sheet.images.push(sheetImageAt('absolute', axes, undefined, { anchorType: '2' }))
  assert.deepEqual(captureSheetClipboardImages(workbook, sheet, range, Array.from({ length: 9 }, (_, row) => row)).map(entry => entry.image.id), ['a'])
  assert.equal(captureSheetImageObjects(workbook, sheet, ['absolute'], 'editor').entries[0].image.anchorType, '2')
})

test('range copy uses all transformed corners and rejects an image crossing the selected physical rectangle', () => {
  const { workbook, sheet } = prepared()
  assert.equal(captureSheetClipboardImages(workbook, sheet, { startRow: 0, endRow: 2, startCol: 0, endCol: 1 }, [0, 1, 2]).length, 0)
})

test('range copy excludes images whose occupied rows were removed from the visible row projection', () => {
  const { workbook, sheet } = prepared(), rows = Array.from({ length: 9 }, (_, row) => row).filter(row => row !== 4)
  assert.equal(captureSheetClipboardImages(workbook, sheet, range, rows).length, 0)
})

test('a Both image collapsed by hidden rows is skipped by range copy and rejected by object capture', () => {
  const { workbook, sheet, axes } = sheetImageFixture(); sheet.images = [sheetImageAt('both', axes, undefined, { anchorType: '1' })]
  sheet.hiddenRows = Array.from({ length: sheet.rowCount }, (_, row) => row)
  assert.equal(sheetImageFrame(sheet.images[0], sheetClipboardObjectAxes(workbook, sheet)).height, 0)
  assert.equal(captureSheetClipboardImages(workbook, sheet, range, Array.from({ length: 9 }, (_, row) => row)).length, 0)
  assert.throws(() => captureSheetImageObjects(workbook, sheet, ['both'], 'editor'), /隐藏|不可复制/)
})

test('range paste carries affine media inside the detached candidate without changing the source', () => {
  const { workbook, sheet, destination } = prepared(), before = copy(workbook), snapshot = captureSheetClipboard(workbook,
    { unitId: workbook.id, sheetId: sheet.id, range }, { userId: 'editor' }), candidate = copy(workbook),
    target = { unitId: workbook.id, sheetId: destination.id, range: { startRow: 12, endRow: 12, startCol: 6, endCol: 6 } },
    targetSheet = candidate.sheets[1], geometry = createSheetClipboardTargetGeometry(candidate, targetSheet, target.range, snapshot)
  pasteSheetClipboardImages(workbook, candidate, target, snapshot, geometry, 'all')
  assert.deepEqual(workbook, before); assert.equal(targetSheet.images.length, 1)
  const image = targetSheet.images[0], origin = sheetClipboardRangeBounds(candidate, targetSheet, geometry.ranges[0])
  for (const key of ['a', 'b', 'c', 'd']) assert.equal(image.drawingGeometry.matrix[key], sheet.images[0].drawingGeometry.matrix[key])
  near(image.drawingGeometry.matrix.e, sheet.images[0].drawingGeometry.matrix.e + origin.left - snapshot.sourceOrigin.left)
  near(image.drawingGeometry.matrix.f, sheet.images[0].drawingGeometry.matrix.f + origin.top - snapshot.sourceOrigin.top)
})

test('values, formulas, formats and column-width paste omit all floating images', () => {
  const { workbook, sheet, destination } = prepared(), snapshot = captureSheetClipboard(workbook,
    { unitId: workbook.id, sheetId: sheet.id, range }, { userId: 'editor' })
  for (const mode of ['values', 'formulas', 'formats', 'columnWidths']) {
    const candidate = copy(workbook), target = { unitId: workbook.id, sheetId: destination.id, range },
      geometry = createSheetClipboardTargetGeometry(candidate, candidate.sheets[1], range, snapshot)
    pasteSheetClipboardImages(workbook, candidate, target, snapshot, geometry, mode)
    assert.deepEqual(candidate, workbook, mode)
  }
})

test('range cut transfers the original image identity only inside the candidate', () => {
  const { workbook, sheet, destination } = prepared(), before = copy(workbook), snapshot = captureSheetClipboard(workbook,
    { unitId: workbook.id, sheetId: sheet.id, range }, { userId: 'editor', copyType: 'cut' }), candidate = copy(workbook),
    target = { unitId: workbook.id, sheetId: destination.id, range: { startRow: 12, endRow: 12, startCol: 6, endCol: 6 } },
    geometry = createSheetClipboardTargetGeometry(candidate, candidate.sheets[1], target.range, snapshot)
  pasteSheetClipboardImages(workbook, candidate, target, snapshot, geometry, 'all', { sourceSheetId: sheet.id })
  assert.equal(candidate.sheets[0].images.length, 0); assert.equal(candidate.sheets[1].images[0].id, 'a'); assert.deepEqual(workbook, before)
})
