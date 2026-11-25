import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, createTable, runText } from '../src/features/document/model/document-model.ts'
import { documentImageAtom, setDocumentImageAtom } from '../src/features/document/images/document-image-atom.ts'
import { documentImageEntries, normalizeDocumentImage, patchDocumentImage } from '../src/features/document/images/document-image-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { copyDocumentBlocksForPaste } from '../src/features/document/services/document-block-paste.ts'
import { applyDocumentPasteMode, documentImageClipboardData, documentSelectionClipboardData, documentTableClipboardData } from '../src/features/document/services/document-clipboard.ts'
import { clearDocumentFormattingRuns } from '../src/features/document/services/document-clear-formatting.ts'
import { createDocumentContentControl } from '../src/features/document/services/content-controls.ts'
import { permissionTargetsForChangedBlocks, permissionTargetsForImage, permissionTargetsForSelection, resolveDocumentPermission } from '../src/features/document/permissions/document-permissions.ts'
import { createDocumentTableViewSelection } from '../src/features/document/model/document-table-selection.ts'
import { documentTableGrid, documentTableWholeRange } from '../src/features/document/model/document-table-grid.ts'
import { documentTableTextMatrix } from '../src/features/document/model/document-table-text.ts'

// DOC-07 continuation drafts. These model/command/clipboard contracts have not run.
const source = 'data:image/png;base64,AA=='
const image = patch => ({ src: source, alt: 'Picture', width: 100, height: 60, intrinsicWidth: 200, intrinsicHeight: 120, ...patch })
const atom = (id = 'picture', patch = {}) => { const run = createRun(); setDocumentImageAtom(run, id, image(patch)); return run }
const paragraph = (id, runs) => ({ ...createParagraph(), id, runs })
const model = blocks => ({ ...createBlankDocument(), id: 'inline-images', blocks })
const range = (start, end = start, blockId = 'p') => ({ anchor: { blockId, offset: start }, focus: { blockId, offset: end } })
const context = { documentId: 'inline-images', userId: 'editor' }
const execute = (document, selection, command) => executeDocumentCommand(document, selection, command, context)
const embedded = document => documentImageEntries(document)[0]

test('flip input validates booleans and changes the atom through the ordinary object command', () => {
  for (const flipX of [1, 'true', null]) assert.throws(() => normalizeDocumentImage(image({ flipX })))
  assert.deepEqual([patchDocumentImage(image(), { flipX: true, flipY: false }).flipX, patchDocumentImage(image(), { flipY: true }).flipY], [true, true])
  const document = model([paragraph('p', [createRun('L'), atom(), createRun('R')])])
  execute(document, range(0), { id: 'doc.command.update-image', imageId: 'picture', patch: { flipX: true, flipY: true } })
  assert.equal(embedded(document).image.flipX, true); assert.equal(embedded(document).image.flipY, true)
  assert.equal(runText(document.blocks[0].runs), 'L\uFFFCR')
})

test('replacement keeps the old flips and reset restores both flips with the intrinsic dimensions', () => {
  const document = model([paragraph('p', [atom('picture', { flipX: true, flipY: true })])])
  execute(document, range(0), { id: 'doc.command.replace-image', imageId: 'picture', image: image({ src: 'data:image/jpeg;base64,AA==', intrinsicWidth: 300, intrinsicHeight: 180 }) })
  assert.deepEqual([embedded(document).image.flipX, embedded(document).image.flipY], [true, true])
  execute(document, range(0), { id: 'doc.command.reset-image-size', imageId: 'picture' })
  assert.deepEqual([embedded(document).image.flipX, embedded(document).image.flipY, embedded(document).image.width, embedded(document).image.height], [false, false, 300, 180])
})

test('insertion replaces a UTF-16 range with one real atom and typing beside it preserves its identity', () => {
  const document = model([{ ...createParagraph('ABCD'), id: 'p' }]), style = structuredClone(document.blocks[0].style)
  const inserted = execute(document, range(2, 3), { id: 'doc.command.insert-image', imageId: 'picture', image: image({ flipX: true }) })
  assert.equal(document.blocks.length, 1); assert.equal(runText(document.blocks[0].runs), 'AB\uFFFCD')
  assert.deepEqual(document.blocks[0].style, style); assert.equal(inserted.focus.offset, 3)
  assert.deepEqual([embedded(document).inline.startOffset, embedded(document).inline.endOffset], [2, 3])
  execute(document, inserted, { id: 'doc.command.insert-text', text: 'x' })
  assert.equal(runText(document.blocks[0].runs), 'AB\uFFFCxD'); assert.equal(embedded(document).imageId, 'picture')
  assert.ok(document.blocks[0].runs.filter(run => !documentImageAtom(run)).every(run => !run.metadata?.imageAtom))
})

test('Backspace and Delete each remove exactly one image atom and preserve adjacent text', () => {
  for (const [offset, direction] of [[2, -1], [1, 1]]) {
    const document = model([paragraph('p', [createRun('L'), atom(), createRun('R')])])
    const result = execute(document, range(offset), { id: 'doc.command.delete', direction })
    assert.equal(runText(document.blocks[0].runs), 'LR'); assert.equal(documentImageEntries(document).length, 0)
    assert.equal(result.focus.offset, 1); assert.equal(document.blocks.length, 1)
  }
})

test('object removal transforms the active caret and enclosing content-control ranges with the same text mutation', () => {
  const document = model([paragraph('p', [createRun('L'), atom(), createRun('Right')])])
  createDocumentContentControl(document, { id: 'control', type: 'richText', blockId: 'p', startOffset: 0, endOffset: 7 })
  const result = execute(document, range(5), { id: 'doc.command.remove-image', imageId: 'picture' })
  assert.equal(result.focus.offset, 4); assert.equal(runText(document.blocks[0].runs), 'LRight')
  assert.equal(document.contentControls.find(control => control.id === 'control').endOffset, 6)
})

test('multiple atoms in the same paragraph and a table cell keep separate image IDs and source addresses', () => {
  const table = createTable(1, 1); table.rows[0].cells[0].blocks = [paragraph('cell-p', [createRun('X'), atom('cell-picture')])]
  const document = model([paragraph('p', [atom('first'), createRun('XY'), atom('second')]),
    { ...createParagraph(), id: 'table', type: 'table', runs: [], metadata: { table } }])
  const entries = documentImageEntries(document)
  assert.deepEqual(entries.map(entry => entry.imageId), ['first', 'second', 'cell-picture'])
  assert.deepEqual(entries.slice(0, 2).map(entry => [entry.block.id, entry.inline.startOffset]), [['p', 0], ['p', 3]])
  assert.equal(entries[2].inline.target.blockId, 'table'); assert.equal(entries[2].inline.target.cellBlockId, 'cell-p')
})

test('each paste owns a fresh atom ID while source pixels, crop and flip values stay intact', () => {
  const original = [paragraph('p', [atom('picture', { flipY: true, crop: { top: 0, right: 0.1, bottom: 0, left: -0.1 } })])], before = structuredClone(original)
  const first = copyDocumentBlocksForPaste(original), second = copyDocumentBlocksForPaste(original)
  const a = documentImageEntries({ blocks: first })[0], b = documentImageEntries({ blocks: second })[0]
  assert.notEqual(a.imageId, 'picture'); assert.notEqual(a.imageId, b.imageId); assert.notEqual(a.inline.run.id, b.inline.run.id)
  assert.deepEqual(a.image, image({ flipY: true, crop: { top: 0, right: 0.1, bottom: 0, left: -0.1 } })); assert.deepEqual(original, before)
  const duplicate = [paragraph('p', [atom('same'), atom('same')])]
  assert.throws(() => copyDocumentBlocksForPaste(duplicate), /重复.*图片|图片.*重复/)
  const invalid = atom(); invalid.text = 'ordinary text'
  assert.throws(() => copyDocumentBlocksForPaste([paragraph('invalid', [invalid])]), /无效.*原子/)
})

test('rich copying includes a real IMG and destination formatting retains the object while text mode emits its alt text', () => {
  const document = model([paragraph('p', [createRun('L'), atom('picture', { flipX: true }), createRun('R')])])
  const copied = documentSelectionClipboardData(document, range(0, 3)), object = documentImageClipboardData(document, 'picture')
  assert.equal(copied.text, 'LPictureR'); assert.match(copied.html, /<img /); assert.match(copied.html, /flipX/)
  assert.equal(object.text, 'Picture'); assert.match(object.html, /data-office-image="v1"/)
  const destination = { paragraphStyle: document.blocks[0].style, textStyle: { fontFamily: 'Arial', fontSize: 20 } }
  const rich = applyDocumentPasteMode(document.blocks, 'destination', destination)
  assert.equal(documentImageEntries({ blocks: rich })[0].image.flipX, true)
  assert.equal(rich[0].runs[0].fontFamily, 'Arial')
  const plain = applyDocumentPasteMode(document.blocks, 'text', destination)
  assert.equal(runText(plain[0].runs), 'LPictureR'); assert.equal(documentImageEntries({ blocks: plain }).length, 0)
})

test('an image entity lock protects the atom and leaves insertion into adjacent ordinary text available', () => {
  const document = model([paragraph('p', [createRun('L'), atom(), createRun('R')])])
  document.permissionRules = [{ id: 'locked', scope: 'entity', targetId: 'picture', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  assert.equal(resolveDocumentPermission(document.permissionRules, context, 'edit', permissionTargetsForSelection(document, range(0, 1))).allowed, true)
  assert.equal(resolveDocumentPermission(document.permissionRules, context, 'edit', permissionTargetsForSelection(document, range(1, 2))).allowed, false)
  assert.ok(permissionTargetsForImage(document, 'picture').some(target => target.scope === 'paragraph' && target.id === 'p'))
  const before = structuredClone(document)
  execute(document, range(0), { id: 'doc.command.insert-text', text: 'x' })
  assert.equal(embedded(document).imageId, 'picture')
  assert.equal(permissionTargetsForChangedBlocks(before, document).some(target => target.scope === 'entity' && target.id === 'picture'), false)
  const unchanged = structuredClone(document)
  assert.throws(() => execute(document, range(0), { id: 'doc.command.update-image', imageId: 'picture', patch: { flipX: true } }), /permission denied/i)
  assert.deepEqual(document, unchanged)
})

test('clearing text formatting retains the selected image atom, resource metadata and stable drawing identity', () => {
  const runs = [createRun('L', { bold: true }), atom('picture', { flipY: true }), createRun('R', { italic: true })]
  const cleared = clearDocumentFormattingRuns(runs, 0, 3, createBlankDocument().defaultStyle.text)
  assert.equal(documentImageAtom(cleared[1]).imageId, 'picture'); assert.equal(documentImageAtom(cleared[1]).image.flipY, true)
  assert.equal(runText(cleared), 'L\uFFFCR'); assert.equal(cleared[0].bold, false)
})

test('table rich copy and text-only paste project image alt text without changing the source UTF-16 model', () => {
  const table = createTable(1, 2)
  table.rows[0].cells[0].blocks = [paragraph('cell-p', [createRun('L'), atom('table-image'), createRun('R')])]
  table.rows[0].cells[1].blocks = [{ ...createParagraph(), id: 'quote', type: 'quote', runs: [],
    children: [paragraph('other-p', [createRun('Other'), atom('nested-image')])] }]
  const document = model([{ ...createParagraph(), id: 'table', type: 'table', runs: [], metadata: { table } }])
  const view = createDocumentTableViewSelection(document, 'table', documentTableWholeRange(documentTableGrid(table))).tableSelection
  const before = structuredClone(document), data = documentTableClipboardData(document, view)
  assert.equal(data.text, 'LPictureR\tOtherPicture'); assert.match(data.html, /<img /)
  assert.equal(documentTableTextMatrix(table)[0][0], 'L\uFFFCR')
  const plain = applyDocumentPasteMode(document.blocks, 'text')
  assert.equal(runText(plain[0].runs), data.text); assert.deepEqual(document, before)
})

test('text replacement removes the picture control resource mirror in the same atom mutation', () => {
  const document = model([paragraph('p', [createRun('L'), atom(), createRun('R')])])
  createDocumentContentControl(document, { id: 'picture-control', type: 'picture', blockId: 'p', startOffset: 1, endOffset: 2,
    metadata: { pictureBlockId: 'picture', pictureImage: image(), custom: 'kept' } })
  const result = execute(document, range(1, 2), { id: 'doc.command.insert-text', text: 'X' })
  const control = document.contentControls.find(control => control.id === 'picture-control')
  assert.equal(runText(document.blocks[0].runs), 'LXR'); assert.equal(result.focus.offset, 2)
  assert.equal(documentImageEntries(document).length, 0); assert.equal(control.metadata.pictureBlockId, undefined)
  assert.equal(control.metadata.pictureImage, undefined); assert.equal(control.metadata.custom, 'kept')
  assert.equal(control.showingPlaceholder, true)
})
