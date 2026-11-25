import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, runText } from '../src/features/document/model/document-model.ts'
import { decodeDocumentBody, encodeDocumentBody } from '../src/api/document-body-codec.ts'
import { DOCUMENT_BODY_TOKENS as T } from '../src/api/document-body-types.ts'
import { documentImageAtom, setDocumentImageAtom } from '../src/features/document/images/document-image-atom.ts'
import { documentImageEntries, setDocumentImageEntry } from '../src/features/document/images/document-image-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'

// Source/body continuation drafts. No file or browser validation has run.
const src = 'data:image/png;base64,AA=='
const drawing = (id = 'picture', layoutType = 0) => ({ drawingId: id, drawingType: 0, source: src, description: id, layoutType,
  docTransform: { size: { width: 100, height: 60 }, angle: 20, flipX: true, flipY: false, authorHint: 'transform' },
  srcRect: { left: 10, right: 10, top: 0, bottom: 0, authorHint: 'crop' }, authorHint: 'drawing' })
const native = (layoutType = 0) => ({ body: { dataStream: 'L😀' + T.customBlock + 'R' + T.paragraph + T.section,
  customBlocks: [{ startIndex: 3, blockId: 'picture', authorHint: 'custom' }], paragraphs: [{ startIndex: 5, paragraphId: 'p' }],
  sectionBreaks: [{ startIndex: 6, sectionId: 'first' }] }, drawings: { picture: drawing('picture', layoutType) }, drawingsOrder: ['picture'] })
const resources = data => ({ drawings: data.drawings, drawingsOrder: data.drawingsOrder })
const decode = data => decodeDocumentBody(data.body, resources(data))
const entries = content => documentImageEntries({ blocks: content.blocks })

test('native inline drawing remains one UTF-16 atom inside the original paragraph and retains both flips', () => {
  const original = native(), before = structuredClone(original), decoded = decode(original)
  assert.equal(decoded.content.blocks.length, 1); assert.equal(decoded.content.blocks[0].id, 'p')
  assert.equal(runText(decoded.content.blocks[0].runs), 'L😀\uFFFCR')
  const entry = entries(decoded.content)[0]
  assert.equal(entry.imageId, 'picture'); assert.deepEqual([entry.inline.startOffset, entry.inline.endOffset], [3, 4])
  assert.deepEqual([entry.image.flipX, entry.image.flipY], [true, false]); assert.deepEqual(original, before)
})

test('native floating drawing keeps its paragraph atom rather than splitting either adjacent text fragment', () => {
  const decoded = decode(native(3)), entry = entries(decoded.content)[0]
  assert.equal(decoded.content.blocks.length, 1); assert.equal(entry.image.wrappingStyle, 'wrapSquare')
  assert.equal(entry.block.id, 'p'); assert.equal(runText(entry.block.runs), 'L😀\uFFFCR')
  const saved = encodeDocumentBody(decoded.content)
  assert.equal(saved.body.customBlocks[0].officeStudio.imageAtom, true)
  assert.equal(saved.drawings.picture.authorHint, 'drawing'); assert.equal(saved.drawings.picture.docTransform.authorHint, 'transform')
})

test('repeated drawing identities reject at the source and model write boundaries', () => {
  const original = native(); original.body.dataStream = T.customBlock + T.customBlock + T.paragraph + T.section
  original.body.customBlocks = [{ startIndex: 0, blockId: 'picture' }, { startIndex: 1, blockId: 'picture' }]
  original.body.paragraphs = [{ startIndex: 2, paragraphId: 'p' }]; original.body.sectionBreaks = [{ startIndex: 3, sectionId: 'first' }]
  assert.throws(() => decode(original), /标识重复/)
  const a = createRun(), b = createRun(), image = { src, alt: 'Picture', width: 100, height: 60 }
  setDocumentImageAtom(a, 'same', image); setDocumentImageAtom(b, 'same', image)
  assert.throws(() => encodeDocumentBody({ blocks: [{ ...createParagraph(), runs: [a, b] }], runs: [], paragraphs: [], metadata: {} }), /身份重复|标识重复/)
})

test('saved atom transforms reopen with the same drawing ID and removing it deletes only its drawing resource', () => {
  const decoded = decode(native()), entry = entries(decoded.content)[0]
  setDocumentImageEntry(entry, { ...entry.image, flipX: false, flipY: true, width: 150, height: 90 })
  const saved = encodeDocumentBody(decoded.content), reopened = decode(saved), current = entries(reopened.content)[0]
  assert.equal(current.imageId, 'picture'); assert.deepEqual([current.image.flipX, current.image.flipY], [false, true])
  assert.deepEqual([current.image.width, current.image.height], [150, 90]); assert.equal(saved.drawings.picture.srcRect.authorHint, 'crop')
  current.block.runs.splice(current.block.runs.indexOf(current.inline.run), 1)
  const removed = encodeDocumentBody(reopened.content)
  assert.equal(removed.drawings?.picture, undefined); assert.equal(runText(reopened.content.blocks[0].runs), 'L😀R')
})

test('changing wrapping preserves the atom save contract and reopening keeps the same text endpoints', () => {
  const decoded = decode(native()), document = { ...createBlankDocument(), id: 'body-inline', blocks: decoded.content.blocks, metadata: decoded.content.metadata }
  executeDocumentCommand(document, { anchor: { blockId: 'p', offset: 0 }, focus: { blockId: 'p', offset: 0 } },
    { id: 'doc.command.update-image', imageId: 'picture', patch: { wrappingStyle: 'inFrontOfText', positionV: { relativeFrom: 0, posOffset: 40 } } },
    { documentId: document.id, userId: 'editor' })
  const saved = encodeDocumentBody({ ...decoded.content, blocks: document.blocks }), reopened = decode(saved)
  assert.equal(reopened.content.blocks.length, 1); assert.equal(entries(reopened.content)[0].image.wrappingStyle, 'inFrontOfText')
  assert.equal(runText(reopened.content.blocks[0].runs), 'L😀\uFFFCR'); assert.equal(saved.body.customBlocks[0].startIndex, 3)
})

test('multiple images in one paragraph write independent custom markers and preserve their arranged layers', () => {
  const a = createRun(), b = createRun(); setDocumentImageAtom(a, 'a', { src, alt: 'A', width: 40, height: 30, layer: 1, flipX: true })
  setDocumentImageAtom(b, 'b', { src, alt: 'B', width: 30, height: 20, layer: 0, flipY: true })
  const content = { blocks: [{ ...createParagraph(), id: 'p', runs: [createRun('L'), a, createRun('M'), b, createRun('R')] }], runs: [], paragraphs: [], metadata: {} }
  const saved = encodeDocumentBody(content), restored = decode(saved)
  assert.deepEqual(saved.body.customBlocks.map(block => [block.blockId, block.startIndex]), [['a', 1], ['b', 3]])
  assert.deepEqual(saved.drawingsOrder, ['b', 'a']); assert.deepEqual(entries(restored.content).map(entry => entry.imageId), ['a', 'b'])
  assert.equal(documentImageAtom(entries(restored.content)[0].inline.run).image.flipX, true)
})
