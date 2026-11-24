import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-11 repeating-row drafts. These have not been executed; browser/Office/large-sample acceptance is deferred.
let server, domain, model, rows, controls, repeating, selections, insertion, interaction, continuations, fields, mutations, editor, geometry, groupLocks, api
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-rows', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  model = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  rows = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/content-control-table-rows.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  repeating = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/repeating-content-control-rows.ts')
  selections = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-table-selection.ts')
  insertion = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/content-control-insertion.ts')
  interaction = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/content-control-interaction.ts')
  continuations = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-table-continuations.ts')
  fields = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-fields.ts')
  mutations = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/table-mutations.ts')
  editor = await server.ssrLoadModule('/packages/office-studio/src/api/document-editor-commands.ts')
  geometry = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/content-control-geometry.ts')
  groupLocks = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/content-control-group-lock.ts')
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
const permission = { documentId: 'row-control', userId: 'editor' }
function paragraph(id, text) { const block = model.createParagraph(text); block.id = id; block.runs[0].id = id + '-run'; return block }
function tableBlock(id, count = 4, columns = 2) {
  const table = { rows: Array.from({ length: count }, (_, r) => ({ id: id + '-r' + r,
    cells: Array.from({ length: columns }, (_, c) => ({ id: id + '-c' + r + '-' + c, blocks: [paragraph(id + '-p' + r + '-' + c, 'R' + r + 'C' + c)],
      style: {}, rowSpan: 1, colSpan: 1, columnIndex: c })) })), columnWidths: Array.from({ length: columns }, () => 100), headerRowCount: 0 }
  return { ...paragraph(id, ''), type: 'table', runs: [], metadata: { table } }
}
function fixture(start = 1, end = 2, count = 4) {
  const result = domain.createBlankDocument(); result.id = 'row-control'
  result.blocks = [paragraph('before', 'Before'), tableBlock('table', count), paragraph('after', 'After')]; result.contentControls = []
  const range = { tableId: 'table', rowIds: table(result).rows.slice(start, end + 1).map(row => row.id) }
  const input = { type: 'repeatingSection', placement: 'row', blockId: 'table-p' + start + '-0', startOffset: 0, endOffset: 4, tableRowRange: range }
  controls.createDocumentContentControl(result, { ...input, id: 'section' })
  controls.createDocumentContentControl(result, { ...input, type: 'repeatingSectionItem', id: 'item', parentId: 'section' })
  return result
}
const table = value => model.tableForBlock(model.findDocumentBlock(value, 'table'))
const control = (value, id = 'item') => value.contentControls.find(control => control.id === id)
const target = (value, id) => model.textTargets(value).find(point => (point.cellBlockId ?? point.blockId) === id)
const rowSelection = value => rows.documentContentControlTableRowSelection(value, control(value).tableRowRange)
function fullRowSelection(value, start = 1, end = 2) {
  return selections.createDocumentTableViewSelection(value, 'table', { startRow: start, endRow: end, startColumn: 0, endColumn: 1 }, 'row')
}
function mergedFixture() {
  const value = fixture(), owner = table(value).rows[1].cells[0]; owner.rowSpan = 2
  table(value).rows[2].cells.splice(0, 1)
  owner.style.nativeContinuationContent = { 1: { version: 1, ownerCellId: owner.id, rowId: 'table-r2', colSpan: 1,
    blocks: [paragraph('continued', 'Continuation')], source: { retained: 'native tc' } } }
  owner.style.nativeContinuations = { 1: { ownerCellId: owner.id, rowId: 'table-r2', colSpan: 1, retained: 'native properties' } }
  return value
}

test('row ownership includes every physical cell and distinguishes foreign story/note tables', () => {
  const value = fixture(), item = control(value)
  assert.equal(rows.documentContentControlContainsBlockRange(value, item, 'table-p2-1', 0, 4), true)
  assert.equal(rows.documentContentControlContainsBlockRange(value, item, 'table-p3-1', 0, 4), false)
  const foreign = { ...item, id: 'foreign', tableRowRange: { tableId: 'note-table', rowIds: ['note-row'] } }
  assert.equal(rows.documentContentControlContainsBlockRange(value, foreign, 'table-p1-0', 0, 4), false)
  assert.equal(rows.documentContentControlContains(value, foreign, item), false)
  assert.throws(() => rows.resolveDocumentContentControlTableRows(value, { tableId: 'table', rowIds: ['table-r2', 'table-r1'] }), /顺序|不连续/)
})

test('whole-row insertion uses persistent row IDs and rejects forged, partial-column and object-only selections', () => {
  const value = fixture(), selected = fullRowSelection(value), input = insertion.documentContentControlInputForSelection(value, 'repeatingSection', selected, 'new-section')
  assert.deepEqual(input.tableRowRange, { tableId: 'table', rowIds: ['table-r1', 'table-r2'] })
  assert.equal(input.placement, 'row'); assert.equal(input.blockId, 'table-p1-0'); assert.equal(input.endOffset, 4)
  const forged = copy(selected); forged.focus.col = 9
  assert.throws(() => insertion.documentContentControlInputForSelection(value, 'repeatingSection', forged), /真实结构选区/)
  assert.throws(() => insertion.documentContentControlInputForSelection(value, 'repeatingSection', selections.createDocumentTableViewSelection(value, 'table',
    { startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 }, 'range')), /完整行/)
  assert.throws(() => insertion.documentContentControlInputForSelection(value, 'text', selected), /单元格正文/)
  assert.throws(() => insertion.documentContentControlInputForSelection(value, 'repeatingSectionItem', selected), /不能单独插入/)
  const objectOnly = fixture(0, 0, 1)
  for (const cell of table(objectOnly).rows[0].cells) cell.blocks = [{ ...paragraph(cell.id + '-image', ''), type: 'image', runs: [], metadata: { image: { src: 'data:image/png;base64,AA==', width: 20, height: 20, alt: '' } } }]
  assert.throws(() => insertion.documentContentControlInputForSelection(objectOnly, 'repeatingSection', fullRowSelection(objectOnly, 0, 0)), /合法的输入段落/)
})

test('two-row duplicate creates independent physical rows and IDs while preserving source and enclosing ranges', () => {
  const value = fixture(), before = copy(value), result = repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'duplicate', permission)
  assert.deepEqual(value, before)
  const copied = control(result.model, result.controlId), newRows = copied.tableRowRange.rowIds
  assert.equal(newRows.length, 2); assert.ok(newRows.every(id => !table(before).rows.some(row => row.id === id)))
  assert.deepEqual(table(result.model).rows.map(row => row.id), ['table-r0', 'table-r1', 'table-r2', ...newRows, 'table-r3'])
  assert.deepEqual(control(result.model, 'section').tableRowRange.rowIds, ['table-r1', 'table-r2', ...newRows])
  assert.deepEqual(control(result.model).tableRowRange, control(before).tableRowRange)
  assert.equal(copied.parentId, 'section'); assert.equal(result.selection.tableSelection.kind, 'row')
  const oldIds = new Set(model.textTargets(before).map(point => point.cellBlockId ?? point.blockId))
  assert.ok(model.textTargets(result.model).filter(point => !oldIds.has(point.cellBlockId ?? point.blockId)).length >= 4)
  assert.equal(table(result.model).rows[3].cells[0].blocks[0].runs[0].text, 'R1C0')
})

test('internal merged rows copy continuation bodies, source properties and nested-table identities', () => {
  const value = mergedFixture(), owner = table(value).rows[1].cells[0]
  continuations.documentTableContinuationEntries(owner)[0].content.blocks.push(tableBlock('nested', 1, 1))
  const result = repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'duplicate', permission), copiedOwner = table(result.model).rows[3].cells[0]
  const continued = continuations.documentTableContinuationEntries(copiedOwner)[0]
  assert.equal(copiedOwner.rowSpan, 2); assert.equal(continued.content.ownerCellId, copiedOwner.id)
  assert.equal(continued.content.rowId, table(result.model).rows[4].id); assert.notEqual(continued.content.blocks[0].id, 'continued')
  assert.equal(continued.content.blocks[0].runs[0].text, 'Continuation'); assert.notEqual(continued.content.blocks[1].id, 'nested')
  assert.equal(copiedOwner.style.nativeContinuations[1].rowId, continued.content.rowId)
  assert.equal(copiedOwner.style.nativeContinuations[1].retained, 'native properties')
})

test('a locked row group permits only its real repeating child and preserves all outside physical rows', () => {
  const value = fixture()
  controls.createDocumentContentControl(value, { id: 'group', type: 'group', placement: 'row',
    tableRowRange: { tableId: 'table', rowIds: table(value).rows.map(row => row.id) }, blockId: 'table-p0-0', startOffset: 0, endOffset: 4, lock: 'contentLocked' })
  control(value, 'section').parentId = 'group'; const before = copy(value)
  const plan = repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'duplicate', permission)
  assert.equal(groupLocks.documentContentControlGroupEditAllowed(before, plan.model, control(before, 'group'), { activeControlId: 'item' }), true)
  assert.deepEqual(table(plan.model).rows[0], table(before).rows[0]); assert.deepEqual(table(plan.model).rows.at(-1), table(before).rows.at(-1))
  const altered = copy(plan.model); table(altered).rows[0].cells[0].blocks[0].runs[0].text = 'Changed outside'
  assert.equal(groupLocks.documentContentControlGroupEditAllowed(before, altered, control(before, 'group'), { activeControlId: 'item' }), false)
  const deleted = repeating.prepareDocumentRepeatingContentControlRowsMutation(plan.model, plan.controlId, 'delete', permission)
  assert.deepEqual(table(deleted.model).rows, table(before).rows)
  const forged = copy(before); control(forged, 'section').parentId = 'missing'
  assert.throws(() => repeating.prepareDocumentRepeatingContentControlRowsMutation(forged, 'item', 'duplicate', permission), /父|所属/)
})

test('locked-group input can edit actual continuation text without changing an outside merged owner', () => {
  const value = mergedFixture()
  controls.createDocumentContentControl(value, { id: 'group', type: 'group', placement: 'row',
    tableRowRange: { tableId: 'table', rowIds: table(value).rows.map(row => row.id) }, blockId: 'table-p0-0', startOffset: 0, endOffset: 4, lock: 'contentLocked' })
  control(value, 'section').parentId = 'group'; const before = copy(value), point = target(value, 'continued')
  editor.applyDocumentEditorCommand(value, { unitId: value.id, selection: { anchor: { ...point, offset: 0 }, focus: { ...point, offset: 12 } },
    command: { id: 'doc.command.insert-text', text: '续格😀', activeContentControlId: 'item' } }, permission)
  assert.equal(model.runText(model.findDocumentBlock(value, 'continued').runs), '续格😀')
  assert.deepEqual(table(value).rows[0], table(before).rows[0]); assert.deepEqual(table(value).rows[3], table(before).rows[3])
  assert.equal(table(value).rows[1].cells[0].rowSpan, 2)
})

test('a recursive note lock is checked before row deletion and leaves the caller untouched', () => {
  const value = fixture(), referenceBlock = table(value).rows[1].cells[0].blocks[0], parent = paragraph('parent-note-p', '1'), nested = paragraph('nested-note-p', 'Locked')
  referenceBlock.runs[0].metadata = { note: { type: 'footnote', noteId: 'parent-note' } }
  parent.runs[0].metadata = { note: { type: 'endnote', noteId: 'nested-note' } }
  value.footnotes = { 'parent-note': { id: 'parent-note', referenceBlockId: referenceBlock.id, referenceRunId: referenceBlock.runs[0].id, text: '1',
    content: { blocks: [parent], runs: copy(parent.runs), paragraphs: [{ startOffset: 0, endOffset: 1, style: {} }] } } }
  value.endnotes = { 'nested-note': { id: 'nested-note', referenceBlockId: parent.id, referenceRunId: parent.runs[0].id, text: 'Locked',
    content: { blocks: [nested], runs: copy(nested.runs), paragraphs: [{ startOffset: 0, endOffset: 6, style: {} }] } } }
  controls.createDocumentContentControl(value, { id: 'nested-note-lock', type: 'richText', blockId: nested.id, startOffset: 0, endOffset: 6, lock: 'contentLocked' })
  const before = copy(value)
  assert.throws(() => controls.deleteRepeatingSectionItem(value, 'item', permission), /锁定/); assert.deepEqual(value, before)
})

test('copy remaps fields, child controls, pictures, comments and note bodies without a phantom shell table', () => {
  const value = fixture(), first = table(value).rows[1].cells[0].blocks[0], last = table(value).rows[2].cells[0].blocks[0]
  fields.setDocumentTextFields(first, [{ id: 'field', field: 'REF', startOffset: 0, endOffset: 4, endBlockId: last.id, instruction: 'REF retained', cachedResult: 'R1C0\nR2C0' }])
  const picture = { ...paragraph('picture', ''), type: 'image', runs: [], metadata: { image: { src: 'data:image/png;base64,AA==', alt: 'Picture', width: 20, height: 20 }, anchorBlockId: 'table', contentControlId: 'image-control' } }
  table(value).rows[1].cells[1].blocks.push(picture)
  controls.createDocumentContentControl(value, { id: 'image-control', type: 'picture', blockId: 'table-p1-1', startOffset: 0, endOffset: 4,
    parentId: 'item', metadata: { pictureBlockId: 'picture', anchorBlockId: 'table' } })
  first.runs[0].metadata = { note: { type: 'footnote', noteId: 'note' } }
  const noteBlock = paragraph('note-p', 'Note'), noteImage = { ...picture, id: 'note-image', metadata: { ...picture.metadata, anchorBlockId: 'note-p' } }
  value.footnotes = { note: { id: 'note', text: 'Note', referenceBlockId: first.id, referenceRunId: first.runs[0].id,
    content: { blocks: [noteBlock, noteImage], runs: copy(noteBlock.runs), paragraphs: [{ startOffset: 0, endOffset: 4, style: {} }] } } }
  value.comments = { thread: { id: 'thread', anchor: [{ blockId: first.id, startOffset: 0, endOffset: 4, text: 'R1C0' }], referenceText: 'R1C0',
    messages: [{ id: 'message', authorId: 'editor', authorName: 'Editor', text: 'Comment', createdAt: '2026-11-25T00:00:00Z' }], resolved: false, createdAt: '2026-11-25T00:00:00Z' } }
  const result = repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'duplicate', permission), copiedItem = control(result.model, result.controlId)
  const copiedPicture = result.model.contentControls.find(control => control.type === 'picture' && control.id !== 'image-control')
  assert.equal(copiedPicture.parentId, copiedItem.id); assert.notEqual(copiedPicture.metadata.pictureBlockId, 'picture'); assert.equal(copiedPicture.metadata.anchorBlockId, 'table')
  assert.equal(model.findDocumentBlock(result.model, copiedPicture.metadata.pictureBlockId).metadata.anchorBlockId, 'table')
  const copiedField = fields.documentTextFieldRanges(result.model).find(range => range.field.id !== 'field')
  assert.notEqual(copiedField.field.endBlockId, last.id); assert.equal(copiedField.endParagraph.text, 'R2C0')
  const copiedNote = Object.values(result.model.footnotes).find(note => note.id !== 'note')
  assert.notEqual(copiedNote.content.blocks[0].id, 'note-p'); assert.equal(copiedNote.content.blocks[1].metadata.anchorBlockId, copiedNote.content.blocks[0].id)
  assert.equal(model.findDocumentBlock(result.model, copiedNote.referenceBlockId).runs[0].metadata.note.noteId, copiedNote.id)
  const copiedThread = Object.values(result.model.comments).find(thread => thread.id !== 'thread')
  assert.notEqual(copiedThread.messages[0].id, 'message'); assert.notEqual(copiedThread.anchor[0].blockId, first.id)
})

test('delete removes real physical rows and resolves the nearest live caret without stale table coordinates', () => {
  const value = fixture(), result = repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'delete', permission)
  assert.deepEqual(table(result.model).rows.map(row => row.id), ['table-r0', 'table-r3']); assert.equal(control(result.model), undefined)
  assert.equal(control(result.model, 'section'), undefined); assert.equal(model.findDocumentBlock(result.model, 'table-p1-0'), undefined)
  assert.equal(result.selection.focus.row, 1); assert.equal(result.selection.focus.cellBlockId, 'table-p3-0')
  assert.equal(model.sameTarget(result.selection.focus, target(result.model, 'table-p3-0')), true)
})

test('deleting all item rows removes the table and selects a surviving paragraph', () => {
  const value = fixture(0, 1, 2), result = repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'delete', permission)
  assert.equal(model.findDocumentBlock(result.model, 'table'), undefined); assert.equal(result.selection.focus.blockId, 'after')
  assert.equal(result.selection.tableSelection, undefined)
})

test('source Copy, affected paragraph Edit, ancestor/descendant locks and section flags reject atomic candidates', () => {
  const cases = [
    value => { value.permissionRules = [{ id: 'no-copy', scope: 'document', action: 'copy', effect: 'deny', userIds: ['editor'] }] },
    value => { value.permissionRules = [{ id: 'second-cell', scope: 'paragraph', targetId: 'table-p2-1', action: 'edit', effect: 'deny', userIds: ['editor'] }] },
    value => { control(value, 'section').lock = 'contentLocked' },
    value => { control(value).lock = 'sdtLocked' },
    value => { control(value, 'section').repeatingSection = { doNotAllowInsertDelete: true } },
  ]
  for (const setup of cases) { const value = fixture(); setup(value); const before = copy(value)
    assert.throws(() => repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'duplicate', permission)); assert.deepEqual(value, before) }
  const value = fixture(); controls.createDocumentContentControl(value, { id: 'locked-child', type: 'richText', parentId: 'item',
    blockId: 'table-p2-1', startOffset: 0, endOffset: 4, lock: 'sdtLocked' })
  assert.throws(() => repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'delete', permission), /锁定/)
})

test('locked fields in rows and referenced notes block deletion, and half a cross-item field cannot be copied', () => {
  for (const where of ['row', 'note']) { const value = fixture(), block = table(value).rows[1].cells[0].blocks[0]
    const fieldBlock = where === 'row' ? block : paragraph('locked-note', 'Locked')
    fields.setDocumentTextFields(fieldBlock, [{ id: 'locked-field', field: 'REF', startOffset: 0, endOffset: 4, locked: true }])
    if (where === 'note') { block.runs[0].metadata = { note: { type: 'footnote', noteId: 'note' } }; value.footnotes = { note: { id: 'note', text: 'Locked', referenceBlockId: block.id, referenceRunId: block.runs[0].id,
      content: { blocks: [fieldBlock], runs: copy(fieldBlock.runs), paragraphs: [{ startOffset: 0, endOffset: 6, style: {} }] } } } }
    const before = copy(value); assert.throws(() => repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'delete', permission), /锁定字段/); assert.deepEqual(value, before)
  }
  const value = fixture(); fields.setDocumentTextFields(table(value).rows[1].cells[0].blocks[0], [{ id: 'across', field: 'REF', startOffset: 0, endOffset: 4, endBlockId: 'table-p3-0' }])
  assert.throws(() => repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'duplicate', permission), /跨项目字段范围/)
})

test('deletion removes unreferenced note controls and comment segments in the same candidate', () => {
  const value = fixture(), block = table(value).rows[1].cells[0].blocks[0], noteBlock = paragraph('note-p', 'Note')
  block.runs[0].metadata = { note: { type: 'footnote', noteId: 'note' } }
  value.footnotes = { note: { id: 'note', text: 'Note', referenceBlockId: block.id, referenceRunId: block.runs[0].id,
    content: { blocks: [noteBlock], runs: copy(noteBlock.runs), paragraphs: [{ startOffset: 0, endOffset: 4, style: {} }] } } }
  value.contentControls.push({ id: 'note-control', type: 'richText', blockId: noteBlock.id, startOffset: 0, endOffset: 4, lock: 'unlocked' })
  value.comments = { noteThread: { id: 'noteThread', anchor: [{ blockId: noteBlock.id, startOffset: 0, endOffset: 4, text: 'Note' }], referenceText: 'Note', messages: [], resolved: false, createdAt: '2026-11-25T00:00:00Z' } }
  const result = repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'delete', permission)
  assert.equal(result.model.footnotes.note, undefined); assert.equal(control(result.model, 'note-control'), undefined); assert.equal(result.model.comments.noteThread, undefined)
})

test('regular row deletion reanchors a wrapperLocked multirow control instead of dropping it with its first paragraph', () => {
  const value = fixture(); control(value).lock = 'sdtLocked'; control(value, 'section').lock = 'sdtLocked'
  const result = mutations.prepareDocumentTableMutation(value, fullRowSelection(value, 1, 1), { id: 'doc.command.table-delete-rows', tableId: 'table',
    range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 } }, { blockId: 'table', range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 } }, permission)
  assert.deepEqual(control(result.model).tableRowRange.rowIds, ['table-r2']); assert.equal(control(result.model).blockId, 'table-p2-0')
  assert.equal(control(result.model).lock, 'sdtLocked'); assert.equal(control(result.model, 'section').blockId, 'table-p2-0')
})

test('ordinary second-cell input unwraps temporary rows and keeps wrapperLocked rows across text replacement', () => {
  for (const lock of ['unlocked', 'sdtLocked']) { const value = fixture(); control(value).temporary = true; control(value).lock = lock
    const point = target(value, 'table-p2-1'), selected = { anchor: { ...point, offset: 0 }, focus: { ...point, offset: 4 } }
    editor.applyDocumentEditorCommand(value, { unitId: value.id, selection: selected, command: { id: 'doc.command.insert-text', text: '输入😀' } }, permission)
    assert.equal(model.runText(model.findDocumentBlock(value, 'table-p2-1').runs), '输入😀')
    assert.equal(!!control(value), lock === 'sdtLocked'); assert.equal(control(value, 'section').tableRowRange.rowIds.length, 2)
  }
})

test('ambiguous duplicate results, reordered physical rows and stale parents cannot publish a structural selection', () => {
  const value = fixture(), result = repeating.prepareDocumentRepeatingContentControlRowsMutation(value, 'item', 'duplicate', permission)
  const duplicate = copy(control(result.model, result.controlId)); duplicate.id = 'ambiguous'; result.model.contentControls.push(duplicate)
  assert.equal(rows.documentContentControlTableRowsMutationSelection(value, result.model, 'item', 'duplicate'), undefined)
  const stale = copy(value); control(stale, 'section').tableRowRange.rowIds.reverse()
  assert.equal(interaction.documentContentControlParentSelection(stale, 'item'), undefined)
  const reordered = copy(value); table(reordered).rows.reverse()
  assert.throws(() => rows.reconcileDocumentContentControlTableRows(value, reordered), /反转|不连续/)
})

test('row geometry uses stable physical metrics on each page and preserves structural page affinity', () => {
  const value = fixture(), page = (index, y) => ({ index, x: 0, y, width: 300, height: 200 })
  const source = { width: 300, height: 424, pages: [page(0, 0), page(1, 224)], lines: [], cells: [], images: [], dividers: [], columnSeparators: [], headings: [], pageHeaderFooters: [], pageContent: [],
    tableFragments: [{ tableId: 'table', x: 20, y: 160, width: 200, height: 40, page: 0, rows: table(value).rows.map((row, index) => ({ id: row.id, row: index, y: 120 + index * 40, height: 40 })) },
      { tableId: 'table', x: 20, y: 244, width: 200, height: 40, page: 1, rows: table(value).rows.map((row, index) => ({ id: row.id, row: index, y: 164 + index * 40, height: 40 })) }] }
  const result = geometry.documentContentControlGeometry(value, source), item = result.find(item => item.control.id === 'item')
  assert.deepEqual(item.fragments.map(fragment => [fragment.page, fragment.x, fragment.y, fragment.width, fragment.height]), [[0, 20, 160, 200, 40], [1, 20, 244, 200, 40]])
  assert.equal(geometry.hitTestDocumentContentControl(result, 200, 175).control.id, 'item')
  const selected = geometry.contentControlSelectionAtFragment(item, item.fragments[1])
  assert.deepEqual(selected.tableSelection.range, fullRowSelection(value).tableSelection.range); assert.equal(selected.tableSelection.pageAffinity, 1)
  assert.equal(selected.anchor.pageAffinity, 1); assert.equal(selected.focus.pageAffinity, 1); assert.equal(item.selection.tableSelection.pageAffinity, undefined)
})

test('registered row add, duplicate, undo/redo and snapshot reopen share one content history', () => {
  const value = fixture(); value.contentControls = []
  const studio = new api.OfficeStudio({ document: value, currentUserId: 'editor' }); studio.activateDocument(value.id)
  const document = new api.FDocument(studio, value.id)
  try {
    const input = insertion.documentContentControlInputForSelection(document.getModel(), 'repeatingSection', fullRowSelection(document.getModel()), 'inserted-section')
    document.setInputSelection(fullRowSelection(document.getModel())); const session = document.getInputState()
    assert.equal(document.executeContentControlCommand('doc.command.add-content-control', { control: input, host: undefined, expectedRevision: session.modelRevision, inputGeneration: session.generation }), true)
    const inserted = insertion.documentContentControlSelectionAfterInsertion(value, document.getModel(), input)
    assert.equal(inserted.selection.tableSelection.kind, 'row'); assert.equal(document.getContentControl(inserted.controlId).duplicateRepeatingItem(), true)
    assert.equal(table(document.getModel()).rows.length, 6); assert.equal(studio.editor.history.units[document.id].undo.length, 2)
    document.undo(); assert.equal(table(document.getModel()).rows.length, 4); document.redo(); assert.equal(table(document.getModel()).rows.length, 6)
    const saved = copy(document.save()), reopened = new api.OfficeStudio({ document: saved, currentUserId: 'editor' })
    try { assert.deepEqual(new api.FDocument(reopened, saved.id).getModel().contentControls, saved.contentControls) } finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('registered add cancellation and stale generation leave content and automatic-item selection unchanged', () => {
  const value = fixture(); value.contentControls = []
  const studio = new api.OfficeStudio({ document: value, currentUserId: 'editor' }); studio.activateDocument(value.id)
  const document = new api.FDocument(studio, value.id), before = copy(document.getModel())
  const input = insertion.documentContentControlInputForSelection(before, 'repeatingSection', fullRowSelection(before), 'cancelled-section')
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.add-content-control') event.cancel = true })
  try {
    const selectionBefore = copy(document.getInputState())
    assert.throws(() => document.executeContentControlCommand('doc.command.add-content-control', { control: input }), /Command was canceled/)
    assert.deepEqual(document.getModel(), before); assert.equal(insertion.documentContentControlSelectionAfterInsertion(before, document.getModel(), input), undefined)
    assert.deepEqual(document.getInputState(), selectionBefore); assert.equal(studio.editor.history.units[document.id]?.undo.length ?? 0, 0)
    cancel.dispose(); document.setInputSelection(fullRowSelection(document.getModel())); const session = document.getInputState(); document.setInputSelection(model.collapsedSelection({ ...target(document.getModel(), 'before'), offset: 0 }))
    assert.throws(() => document.executeContentControlCommand('doc.command.add-content-control', { control: input, host: undefined, expectedRevision: session.modelRevision, inputGeneration: session.generation }), /会话已更改/)
    assert.deepEqual(document.getModel(), before); assert.equal(studio.editor.history.units[document.id]?.undo.length ?? 0, 0)
  } finally { cancel.dispose(); studio.dispose() }
})
