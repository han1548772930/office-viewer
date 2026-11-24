import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-11 complete-block drafts. No test, browser, Office or performance acceptance has been executed.
let server, domain, model, controls, repeating, ranges, rows, fields, insertion, interaction, permissions, hosts, editor, layout, pictureAnchors, geometry, api
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-blocks', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  model = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  repeating = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/repeating-content-control-blocks.ts')
  ranges = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/content-control-block-range.ts')
  rows = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/content-control-table-rows.ts')
  fields = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-fields.ts')
  insertion = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/content-control-insertion.ts')
  interaction = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/content-control-interaction.ts')
  permissions = await server.ssrLoadModule('/packages/office-studio/src/features/document/permissions/content-control-permissions.ts')
  hosts = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-block-host.ts')
  editor = await server.ssrLoadModule('/packages/office-studio/src/api/document-editor-commands.ts')
  layout = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
  pictureAnchors = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-picture-input-anchor.ts')
  geometry = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/content-control-geometry.ts')
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
const permission = { documentId: 'block-control', userId: 'editor' }
function paragraph(id, text = '') { const block = model.createParagraph(text); block.id = id; block.runs[0].id = id + '-run'; return block }
function tableBlock(id = 'table') {
  return { ...paragraph(id), type: 'table', runs: [], metadata: { table: {
    rows: Array.from({ length: 2 }, (_, row) => ({ id: id + '-r' + row,
      cells: Array.from({ length: 2 }, (_, column) => ({ id: id + '-c' + row + '-' + column,
        blocks: [paragraph(id + '-p' + row + '-' + column, 'Cell ' + row + column)], style: {}, rowSpan: 1, colSpan: 1, columnIndex: column })) })),
    columnWidths: [90, 90], headerRowCount: 0,
  } } }
}
function fixture() {
  const value = domain.createBlankDocument(); value.id = 'block-control'
  value.blocks = [paragraph('before', 'Before'), paragraph('a', 'A'), tableBlock(), paragraph('b', 'B'), paragraph('after', 'After')]
  value.contentControls = []
  const input = { type: 'repeatingSection', placement: 'block', blockRange: { blockIds: ['a', 'table', 'b'] }, blockId: 'a', startOffset: 0, endOffset: 1 }
  controls.createDocumentContentControl(value, { ...input, id: 'section' })
  controls.createDocumentContentControl(value, { ...input, type: 'repeatingSectionItem', id: 'item', parentId: 'section' })
  return value
}
const control = (value, id = 'item') => value.contentControls?.find(control => control.id === id)
const selected = value => ranges.documentContentControlBlocksSelection(value, control(value).blockRange)
const target = (value, id) => model.textTargets(value).find(point => (point.cellBlockId ?? point.blockId) === id)
const table = value => model.tableForBlock(model.findDocumentBlock(value, 'table'))
const storyContent = blocks => ({ blocks, runs: copy(blocks.flatMap(block => block.runs)), paragraphs: blocks.map(block => ({ startOffset: 0, endOffset: model.runText(block.runs).length, style: {} })) })
const measurer = { measure: text => [...text].length * 7 }

test('complete-block insertion follows the live container and rejects partial text or a different cell container', () => {
  const value = fixture(), full = selected(value), input = insertion.documentContentControlInputForSelection(value, 'repeatingSection', full, 'new')
  assert.deepEqual(input.blockRange, { blockIds: ['a', 'table', 'b'] }); assert.equal(input.placement, 'block')
  const reverse = { anchor: copy(full.focus), focus: copy(full.anchor) }
  assert.deepEqual(insertion.documentContentControlInputForSelection(value, 'group', reverse).blockRange, input.blockRange)
  const partial = copy(full); partial.anchor.offset = 1
  assert.throws(() => insertion.documentContentControlInputForSelection(value, 'richText', partial), /完整内容/)
  const first = target(value, 'table-p0-0'), last = target(value, 'table-p0-1')
  assert.throws(() => insertion.documentContentControlInputForSelection(value, 'group', {
    anchor: { ...first, offset: 0 }, focus: { ...last, offset: model.runText(model.runsForTarget(value, last)).length },
  }), /容器/)
  assert.throws(() => ranges.resolveDocumentContentControlBlocks(value, { blockIds: ['b', 'table', 'a'] }), /顺序|不连续|连续/)
})

test('actual block duplicate allocates a complete nested tree and remaps row controls, fields and picture ownership', () => {
  const value = fixture(), rowInput = { type: 'repeatingSection', placement: 'row', tableRowRange: { tableId: 'table', rowIds: ['table-r0'] },
    blockId: 'table-p0-0', startOffset: 0, endOffset: 7, parentId: 'item' }
  controls.createDocumentContentControl(value, { ...rowInput, id: 'row-section' })
  controls.createDocumentContentControl(value, { ...rowInput, type: 'repeatingSectionItem', id: 'row-item', parentId: 'row-section' })
  fields.setDocumentTextFields(model.findDocumentBlock(value, 'a'), [{ id: 'field', field: 'REF', startOffset: 0, endOffset: 1, endBlockId: 'b', cachedResult: 'A\nB', instruction: 'REF retained' }])
  const anchor = table(value).rows[0].cells[1].blocks[0], picture = { ...paragraph('picture'), type: 'image', runs: [],
    metadata: { image: { src: 'data:image/png;base64,AA==', width: 20, height: 20, alt: 'Nested image' }, anchorBlockId: anchor.id, anchorOffset: 0, contentControlId: 'picture-control' } }
  table(value).rows[0].cells[1].blocks.push(picture)
  controls.createDocumentContentControl(value, { id: 'picture-control', type: 'picture', blockId: anchor.id, startOffset: 0, endOffset: 7,
    parentId: 'row-item', metadata: { pictureBlockId: 'picture', pictureAnchorBlockId: anchor.id } })
  const before = copy(value), plan = repeating.prepareDocumentRepeatingContentControlBlocksMutation(value, 'item', 'duplicate', permission)
  assert.deepEqual(value, before)
  const copied = control(plan.model, plan.controlId), copiedIds = copied.blockRange.blockIds
  assert.equal(copiedIds.length, 3); assert.equal(copied.parentId, 'section'); assert.ok(copiedIds.every(id => !['a', 'table', 'b'].includes(id)))
  assert.deepEqual(plan.model.blocks.map(block => block.id), ['before', 'a', 'table', 'b', ...copiedIds, 'after'])
  assert.deepEqual(control(plan.model, 'section').blockRange.blockIds, ['a', 'table', 'b', ...copiedIds])
  const row = control(plan.model, plan.identityMap.control.get('row-item'))
  assert.equal(row.tableRowRange.tableId, plan.identityMap.block.get('table')); assert.equal(row.tableRowRange.rowIds[0], plan.identityMap.row.get('table-r0'))
  assert.equal(rows.documentContentControlTableRowSelection(plan.model, row.tableRowRange).tableSelection.kind, 'row')
  const copiedField = fields.documentTextFields(model.findDocumentBlock(plan.model, plan.identityMap.block.get('a')))[0]
  assert.notEqual(copiedField.id, 'field'); assert.equal(copiedField.endBlockId, plan.identityMap.block.get('b')); assert.equal(copiedField.instruction, 'REF retained')
  const copiedPicture = control(plan.model, plan.identityMap.control.get('picture-control'))
  assert.equal(copiedPicture.metadata.pictureBlockId, plan.identityMap.block.get('picture'))
  assert.equal(model.findDocumentBlock(plan.model, copiedPicture.metadata.pictureBlockId).metadata.contentControlId, copiedPicture.id)
  assert.deepEqual(interaction.documentContentControlSelectionForId(plan.model, copied.id), plan.selection)
})

test('block copy and deletion include typed note resources, note controls and independent thread messages', () => {
  const value = fixture(), a = model.findDocumentBlock(value, 'a'), noteBlock = paragraph('note-p', 'Note')
  a.runs[0].text = '1'; a.runs[0].metadata = { note: { type: 'footnote', noteId: '7' }, noteReference: { footnoteId: 7 } }
  value.footnotes = { '7': { id: '7', text: 'Note', referenceBlockId: 'a', referenceRunId: a.runs[0].id, content: storyContent([noteBlock]) } }
  controls.createDocumentContentControl(value, { id: 'note-group', type: 'group', blockId: noteBlock.id, startOffset: 0, endOffset: 4 })
  controls.createDocumentContentControl(value, { id: 'note-child', type: 'richText', parentId: 'note-group', blockId: noteBlock.id, startOffset: 0, endOffset: 4 })
  value.comments = { thread: { id: 'thread', anchor: [{ blockId: 'a', startOffset: 0, endOffset: 1, text: '1' },
    { blockId: noteBlock.id, startOffset: 0, endOffset: 4, text: 'Note' }], referenceText: '1\nNote',
    messages: [{ id: 'message', authorId: 'editor', authorName: 'Editor', text: 'Comment', createdAt: '2026-11-25T00:00:00Z' }], resolved: false, createdAt: '2026-11-25T00:00:00Z' } }
  const plan = repeating.prepareDocumentRepeatingContentControlBlocksMutation(value, 'item', 'duplicate', permission)
  const noteId = plan.identityMap.note.get('footnotes:7'), copiedNote = plan.model.footnotes[noteId]
  assert.notEqual(noteId, '7'); assert.equal(copiedNote.referenceBlockId, plan.identityMap.block.get('a')); assert.equal(copiedNote.referenceRunId, plan.identityMap.run.get('a-run'))
  const child = control(plan.model, plan.identityMap.control.get('note-child'))
  assert.equal(child.blockId, plan.identityMap.block.get('note-p')); assert.equal(child.parentId, plan.identityMap.control.get('note-group'))
  const thread = plan.model.comments[plan.identityMap.comment.get('thread')]
  assert.equal(thread.anchor.length, 2); assert.notEqual(thread.messages[0].id, 'message'); assert.equal(thread.messages[0].text, 'Comment')
  const deleted = repeating.prepareDocumentRepeatingContentControlBlocksMutation(plan.model, plan.controlId, 'delete', permission)
  assert.equal(deleted.model.footnotes[noteId], undefined); assert.equal(control(deleted.model, child.id), undefined); assert.equal(deleted.model.comments[thread.id], undefined)
  assert.ok(deleted.model.footnotes['7']); assert.ok(deleted.model.comments.thread)
})

test('locked groups admit a real child edit and preserve their full outside text, structure and formatting', () => {
  const value = fixture()
  controls.createDocumentContentControl(value, { id: 'group', type: 'group', placement: 'block', blockRange: { blockIds: value.blocks.map(block => block.id) },
    blockId: 'before', startOffset: 0, endOffset: 6, lock: 'contentLocked' })
  control(value, 'section').parentId = 'group'; model.findDocumentBlock(value, 'before').runs[0].bold = true
  const before = copy(value), plan = repeating.prepareDocumentRepeatingContentControlBlocksMutation(value, 'item', 'duplicate', permission)
  assert.deepEqual(model.findDocumentBlock(plan.model, 'before'), model.findDocumentBlock(before, 'before'))
  assert.deepEqual(model.findDocumentBlock(plan.model, 'after'), model.findDocumentBlock(before, 'after'))
  assert.equal(control(plan.model, 'group').lock, 'contentLocked')
  const deleted = repeating.prepareDocumentRepeatingContentControlBlocksMutation(plan.model, plan.controlId, 'delete', permission)
  assert.deepEqual(deleted.model.blocks.map(block => block.id), before.blocks.map(block => block.id))
  const point = target(value, 'table-p1-1'), selection = { anchor: { ...point, offset: 0 }, focus: { ...point, offset: 7 } }
  const edited = copy(value)
  editor.applyDocumentEditorCommand(edited, { unitId: edited.id, selection, command: { id: 'doc.command.insert-text', text: '输入😀', activeContentControlId: 'item' } }, permission)
  assert.equal(model.runText(model.findDocumentBlock(edited, 'table-p1-1').runs), '输入😀')
  assert.deepEqual(model.findDocumentBlock(edited, 'before'), model.findDocumentBlock(before, 'before'))
  const locked = copy(value); control(locked).lock = 'contentLocked'; const untouched = copy(locked)
  assert.throws(() => controls.duplicateRepeatingSectionItem(locked, 'item', permission), /锁定/); assert.deepEqual(locked, untouched)
})

test('scoped permission checks address the real migrated note, nested table and image instead of its stale source block', () => {
  const value = fixture(), a = model.findDocumentBlock(value, 'a'), noteTable = tableBlock('note-table')
  a.runs[0].metadata = { note: { type: 'footnote', noteId: 'note' }, officeStoryRunSourceId: 'old-reference-run' }; a.runs[0].text = '1'
  value.footnotes = { note: { id: 'note', text: '', referenceBlockId: 'before', referenceRunId: 'old-reference-run', content: storyContent([noteTable]) } }
  const addressed = permissions.documentContentControlPermissionTargets(value, control(value))
  assert.ok(addressed.some(target => target.scope === 'entity' && target.id === 'footnote:note'))
  assert.ok(addressed.some(target => target.scope === 'paragraph' && target.id === 'note-table-p1-1'))
  assert.ok(addressed.some(target => target.scope === 'entity' && target.id === 'note-table'))
  value.permissionRules = [{ id: 'deny-note', scope: 'entity', targetId: 'footnote:note', action: 'copy', effect: 'deny', userIds: ['editor'] }]
  const before = copy(value)
  assert.throws(() => controls.duplicateRepeatingSectionItem(value, 'item', permission), /权限|permission|无权/); assert.deepEqual(value, before)
})

test('multiline full-block values retain the first paragraph source, replace the actual tree and unwrap temporary content atomically', () => {
  const value = fixture(); value.contentControls = []
  model.findDocumentBlock(value, 'a').style.align = 'right'; model.findDocumentBlock(value, 'a').metadata.retainedSource = { key: 'source' }
  controls.createDocumentContentControl(value, { id: 'text', type: 'richText', placement: 'block', blockRange: { blockIds: ['a', 'table', 'b'] },
    blockId: 'a', startOffset: 0, endOffset: 1, temporary: true })
  controls.setDocumentContentControlValue(value, 'text', '输入😀\nSecond')
  assert.equal(control(value, 'text'), undefined)
  assert.equal(value.blocks.length, 4); assert.equal(model.runText(value.blocks[1].runs), '输入😀'); assert.equal(model.runText(value.blocks[2].runs), 'Second')
  assert.equal(value.blocks[1].style.align, 'right'); assert.deepEqual(value.blocks[1].metadata.retainedSource, { key: 'source' })
  assert.equal(value.blocks[0].id, 'before'); assert.equal(value.blocks.at(-1).id, 'after')
})

function pictureFixture(temporary = false) {
  const value = domain.createBlankDocument(); value.id = 'block-control'; value.blocks = [paragraph('picture-anchor', 'P'), paragraph('following', 'Following')]; value.contentControls = []
  controls.createDocumentContentControl(value, { id: 'picture-control', type: 'picture', blockId: 'picture-anchor', startOffset: 0, endOffset: 1,
    placement: 'block', blockRange: { blockIds: ['picture-anchor'] }, temporary })
  return value
}
const image = { src: 'data:image/png;base64,AA==', alt: 'Picture', width: 80, height: 30 }

test('block picture replacements own one empty input bridge and image, including temporary unwrap and locked shared-owner rejection', () => {
  for (const temporary of [false, true]) {
    const value = pictureFixture(temporary), pictureId = controls.replacePictureContentControlImage(value, 'picture-control', image)
    assert.equal(value.blocks.length, 3); assert.equal(model.runText(value.blocks[0].runs), '')
    const picture = model.findDocumentBlock(value, pictureId)
    assert.equal(picture.metadata.pictureInputAnchor, true); assert.equal(picture.metadata.anchorBlockId, value.blocks[0].id); assert.equal(picture.metadata.anchorOffset, 0)
    if (temporary) { assert.equal(control(value, 'picture-control'), undefined); assert.equal(picture.metadata.contentControlId, undefined) }
    else {
      assert.deepEqual(control(value, 'picture-control').blockRange.blockIds, [value.blocks[0].id, pictureId])
      value.contentControls.push({ ...copy(control(value, 'picture-control')), id: 'locked-picture', parentId: 'picture-control', lock: 'contentLocked' })
      const before = copy(value)
      assert.throws(() => controls.replacePictureContentControlImage(value, 'picture-control', { ...image, alt: 'Changed' }), /锁定子控件|多个/)
      assert.deepEqual(value, before)
    }
  }
})

test('dangling, non-image and outside-range picture pointers fail without publishing any partial candidate', () => {
  for (const pointer of ['missing', 'following', '', 4]) {
    const value = pictureFixture(); control(value, 'picture-control').metadata = { pictureBlockId: pointer }
    const before = copy(value); assert.throws(() => controls.replacePictureContentControlImage(value, 'picture-control', image), /身份|实际图片/); assert.deepEqual(value, before)
  }
  const value = pictureFixture(), outside = { ...paragraph('outside-picture'), type: 'image', runs: [], metadata: { image } }
  value.blocks.push(outside); control(value, 'picture-control').metadata = { pictureBlockId: outside.id }
  const before = copy(value); assert.throws(() => controls.replacePictureContentControlImage(value, 'picture-control', image), /范围之外/); assert.deepEqual(value, before)
})

test('picture input bridges share physical layout, image hit bounds and tiny-image caret limits after an unwrap', () => {
  const value = pictureFixture(true); controls.replacePictureContentControlImage(value, 'picture-control', { ...image, height: 1 })
  const rendered = layout.layoutDocument(value, measurer), picture = rendered.images.find(image => image.blockId === value.blocks[1].id)
  const line = rendered.lines.find(line => line.target.blockId === value.blocks[0].id), following = rendered.lines.find(line => line.target.blockId === 'following')
  assert.equal(line.page, picture.page); assert.equal(line.start, 0); assert.equal(line.end, 0)
  assert.ok(line.y >= picture.y && line.y + line.height <= picture.y + picture.height)
  const caret = layout.caretRect(rendered, { blockId: value.blocks[0].id, offset: 0 })
  assert.ok(caret && caret.y >= picture.y && caret.y + caret.height <= picture.y + picture.height)
  assert.ok(following.y >= picture.y + picture.height); assert.equal(pictureAnchors.documentPictureInputAnchors(value.blocks).get(picture.blockId).id, value.blocks[0].id)
  const separate = copy(value); separate.blocks[0].runs[0].text = 'Actual text'
  assert.equal(pictureAnchors.documentPictureInputAnchors(separate.blocks).size, 0)
  fields.setDocumentTextFields(value.blocks[0], [{ id: 'empty-field', field: 'DATE', startOffset: 0, endOffset: 0 }])
  assert.equal(pictureAnchors.documentPictureInputAnchors(value.blocks).size, 0)
  const persistent = pictureFixture(); controls.replacePictureContentControlImage(persistent, 'picture-control', image)
  const content = layout.layoutDocument(persistent, measurer), imageBounds = content.images[0]
  const regions = geometry.documentContentControlGeometry(persistent, content), hit = geometry.hitTestDocumentContentControl(regions,
    imageBounds.x + imageBounds.width / 2, imageBounds.y + imageBounds.height / 2)
  assert.equal(hit.control.id, 'picture-control')
  assert.deepEqual(hit.selection, interaction.documentContentControlSelectionForId(persistent, 'picture-control'))
})

test('full-block controls in a stored header use the same resource copy, logical selection and permission scopes', () => {
  const value = fixture(), content = value.blocks.slice(1, 4), body = [paragraph('body', 'Body')]
  value.blocks = body; value.sectionHeaderFooters = { first: { header: '', storyContent: { header: storyContent(content) } } }
  const host = { kind: 'story', sectionId: 'first', story: 'header' }, source = copy(value), oldSelection = interaction.documentContentControlSelectionForId(value, 'item', host)
  const copied = controls.duplicateRepeatingSectionItem(value, 'item', permission)
  const request = { source, controlId: 'item', selection: oldSelection, host, command: { id: 'doc.command.duplicate-repeating-section-item', params: { controlId: 'item' } } }
  const resolved = interaction.documentContentControlSelectionAfterRepeatingItemEdit(request, value)
  assert.equal(resolved.controlId, copied.id); assert.deepEqual(value.blocks, body)
  const scoped = hosts.documentModelForBlockHost(value, host), addressed = permissions.documentContentControlPermissionTargets(scoped, control(scoped, copied.id), value)
  assert.ok(addressed.some(target => target.scope === 'entity' && target.id === 'first:header'))
  assert.equal(value.sectionHeaderFooters.first.storyContent.header.blocks.length, 6)
})

test('registered block insertion, copy, delete, undo/redo and snapshot reopen share document content history', () => {
  const value = fixture(); value.contentControls = []
  const studio = new api.OfficeStudio({ document: value, currentUserId: 'editor' }); studio.activateDocument(value.id)
  const document = new api.FDocument(studio, value.id)
  try {
    const selection = ranges.documentContentControlBlocksSelection(document.getModel(), { blockIds: ['a', 'table', 'b'] })
    document.setInputSelection(selection); const state = document.getInputState(), before = copy(document.getModel())
    const input = insertion.documentContentControlInputForSelection(before, 'repeatingSection', selection, 'inserted')
    assert.equal(document.executeContentControlCommand('doc.command.add-content-control', { control: input, host: undefined, expectedRevision: state.modelRevision, inputGeneration: state.generation }), true)
    const item = insertion.documentContentControlSelectionAfterInsertion(before, document.getModel(), input)
    assert.equal(document.getContentControl(item.controlId).duplicateRepeatingItem(), true)
    const copied = document.getModel().contentControls.find(control => control.parentId === 'inserted' && control.id !== item.controlId)
    assert.equal(document.getContentControl(copied.id).deleteRepeatingItem(), true)
    assert.equal(studio.editor.history.units[document.id].undo.length, 3); assert.equal(document.getModel().blocks.length, 5)
    document.undo(); assert.equal(document.getModel().blocks.length, 8); document.redo(); assert.equal(document.getModel().blocks.length, 5)
    const saved = copy(document.save()), reopened = new api.OfficeStudio({ document: saved, currentUserId: 'editor' })
    try { assert.deepEqual(new api.FDocument(reopened, saved.id).getModel().contentControls, saved.contentControls) } finally { reopened.dispose() }
  } finally { studio.dispose() }
})
