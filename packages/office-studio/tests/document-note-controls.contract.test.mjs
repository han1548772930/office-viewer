import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-11/DOC-12 contract drafts. No tests, browser, native IO or performance acceptance were run in this batch.
let server, domain, model, controls, notes, content, settings, fields, stories, hosts
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-note-controls', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  model = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  notes = await server.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-note-editing.ts')
  content = await server.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-note-content.ts')
  settings = await server.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-note-settings.ts')
  fields = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-fields.ts')
  stories = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-story-blocks.ts')
  hosts = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-block-host.ts')
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
function paragraph(id, text = '') { const block = model.createParagraph(text); block.id = id; block.runs[0].id = id + '-run'; return block }
const story = blocks => stories.projectDocumentStoryBlocks({ blocks, runs: [], paragraphs: [] }, 'display').content
const byId = (value, id) => value.contentControls.find(control => control.id === id)
const point = value => ({ blockId: value.blocks[0].id, offset: model.runText(value.blocks[0].runs).length })
const selected = value => { const at = point(value); return { anchor: at, focus: at } }
function tree() {
  const source = domain.createBlankDocument(), imageAnchor = paragraph('image-anchor'), imageTail = paragraph('image-tail')
  const image = { ...paragraph('image'), type: 'image', runs: [], metadata: {
    image: { src: 'data:image/png;base64,AA==', width: 28, height: 18, alt: 'Source image' },
    anchorBlockId: 'image-anchor', tailBlockId: 'image-tail', contentControlId: 'picture',
    provenance: { contentControlId: 'source-only', anchorBlockId: 'source-only' },
  } }
  const table = { ...paragraph('table'), type: 'table', runs: [], metadata: { table: { columnWidths: [160], headerRowCount: 0,
    rows: [
      { id: 'row-0', cells: [{ id: 'cell-0', blocks: [paragraph('cell-p0', 'Cell zero'), imageAnchor, image, imageTail], style: {}, rowSpan: 1, colSpan: 1 }] },
      { id: 'row-1', cells: [{ id: 'cell-1', blocks: [paragraph('cell-p1', 'Cell one')], style: {}, rowSpan: 1, colSpan: 1 }] },
    ],
  } } }
  source.blocks = [paragraph('head', 'A😀'), table, paragraph('tail', 'Tail')]; source.contentControls = []
  fields.setDocumentTextFields(source.blocks[0], [{ id: 'source-field', field: 'REF', startOffset: 0, endOffset: 4,
    endBlockId: 'tail', cachedResult: 'A😀\nTail', instruction: 'REF retained', metadata: { rawInstruction: 'REF retained' } }])
  const add = input => controls.createDocumentContentControl(source, input)
  add({ id: 'root', type: 'group', placement: 'block', blockId: 'head', startOffset: 0, endOffset: 3,
    blockRange: { blockIds: ['head', 'table', 'tail'] }, rawSdtPr: '<w:sdtPr><w:tag w:val="source-root"/></w:sdtPr>' })
  const blockInput = { placement: 'block', blockId: 'cell-p0', startOffset: 0, endOffset: 9, blockRange: { blockIds: ['table'] } }
  add({ ...blockInput, id: 'block-section', type: 'repeatingSection', parentId: 'root' })
  add({ ...blockInput, id: 'block-item', type: 'repeatingSectionItem', parentId: 'block-section' })
  const rowInput = { placement: 'row', blockId: 'cell-p0', startOffset: 0, endOffset: 9, tableRowRange: { tableId: 'table', rowIds: ['row-0'] } }
  add({ ...rowInput, id: 'row-section', type: 'repeatingSection', parentId: 'block-item' })
  add({ ...rowInput, id: 'row-item', type: 'repeatingSectionItem', parentId: 'row-section' })
  add({ id: 'picture', type: 'picture', blockId: 'image-anchor', startOffset: 0, endOffset: 0, parentId: 'row-item',
    metadata: { pictureBlockId: 'image', pictureAnchorBlockId: 'image-anchor', pictureTailBlockId: 'image-tail' } })
  add({ id: 'child', type: 'richText', blockId: 'cell-p1', startOffset: 0, endOffset: 8, parentId: 'block-item' })
  model.findDocumentBlock(source, 'cell-p1').runs[0].metadata = { contentControlId: 'child', provenance: { contentControlId: 'source-only' } }
  return { content: story(source.blocks), contentControls: copy(source.contentControls) }
}
function assertCopiedTree(source, prepared) {
  const target = { ...domain.createBlankDocument(), blocks: prepared.content.blocks, contentControls: prepared.contentControls }
  controls.validateDocumentContentControlRanges(target)
  // Structural positions provide the actual copy correspondence; no IDs are predicted from allocation order.
  const root = prepared.contentControls.find(control => control.blockRange?.blockIds.length === 3)
  const blockSection = prepared.contentControls.find(control => control.type === 'repeatingSection' && control.blockRange?.blockIds.length === 1)
  const blockItem = prepared.contentControls.find(control => control.type === 'repeatingSectionItem' && control.blockRange)
  const rowSection = prepared.contentControls.find(control => control.type === 'repeatingSection' && control.tableRowRange)
  const rowItem = prepared.contentControls.find(control => control.type === 'repeatingSectionItem' && control.tableRowRange)
  const picture = prepared.contentControls.find(control => control.type === 'picture'), child = prepared.contentControls.find(control => control.type === 'richText')
  for (const [id, control] of [['root', root], ['block-section', blockSection], ['block-item', blockItem], ['row-section', rowSection], ['row-item', rowItem], ['picture', picture], ['child', child]]) {
    assert.notEqual(control.id, id)
  }
  assert.equal(blockSection.parentId, root.id); assert.equal(blockItem.parentId, blockSection.id)
  assert.equal(rowSection.parentId, blockItem.id); assert.equal(rowItem.parentId, rowSection.id)
  assert.equal(picture.parentId, rowItem.id); assert.equal(child.parentId, blockItem.id)
  const tableBlock = prepared.content.blocks[1], table = model.tableForBlock(tableBlock), cell = table.rows[0].cells[0]
  assert.deepEqual(root.blockRange.blockIds, prepared.content.blocks.map(block => block.id))
  assert.deepEqual(blockSection.blockRange.blockIds, [tableBlock.id]); assert.equal(rowSection.tableRowRange.tableId, tableBlock.id)
  assert.deepEqual(rowItem.tableRowRange.rowIds, [table.rows[0].id]); assert.notEqual(table.rows[0].id, 'row-0'); assert.notEqual(cell.id, 'cell-0')
  assert.equal(picture.blockId, cell.blocks[1].id); assert.equal(picture.metadata.pictureBlockId, cell.blocks[2].id)
  assert.equal(picture.metadata.pictureTailBlockId, cell.blocks[3].id); assert.equal(cell.blocks[2].metadata.anchorBlockId, cell.blocks[1].id)
  assert.equal(cell.blocks[2].metadata.tailBlockId, cell.blocks[3].id); assert.equal(cell.blocks[2].metadata.contentControlId, picture.id)
  assert.deepEqual(cell.blocks[2].metadata.provenance, source.content.blocks[1].metadata.table.rows[0].cells[0].blocks[2].metadata.provenance)
  assert.equal(table.rows[1].cells[0].blocks[0].runs[0].metadata.contentControlId, child.id)
  assert.notEqual(prepared.content.blocks[0].runs[0].id, 'head-run')
  const field = fields.documentTextFields(prepared.content.blocks[0])[0]
  assert.notEqual(field.id, 'source-field'); assert.equal(field.endBlockId, prepared.content.blocks[2].id)
  assert.equal(field.instruction, 'REF retained'); assert.equal(root.rawSdtPr, source.contentControls.find(control => control.id === 'root').rawSdtPr)
  return target
}

test('block, row, field, picture and wrapper identities use the actual shared copy map without changing the source', () => {
  const source = tree(), before = copy(source), value = domain.createBlankDocument()
  const prepared = content.prepareDocumentNoteContentWithControls(value, 'footnote', 'note', source.content, '', true, source.contentControls)
  assertCopiedTree(source, prepared); assert.deepEqual(source, before)
  prepared.content.blocks[0].runs[0].text = 'Changed'; assert.deepEqual(source, before)
})

test('reusing a source for multiple notes allocates independent bodies and controls, including a shared ID across kinds', () => {
  const source = tree(), value = domain.createBlankDocument(), before = copy(source)
  for (const [kind, noteId] of [['footnote', 'shared'], ['endnote', 'shared'], ['footnote', 'another']]) {
    const inserted = notes.insertDocumentNote(value, selected(value), kind, { noteId, ...source })
    assert.equal(inserted.host.noteId, noteId)
  }
  const all = [value.footnotes.shared, value.endnotes.shared, value.footnotes.another]
  assert.equal(new Set(all.map(note => note.content.blocks[0].id)).size, 3)
  assert.equal(value.contentControls.length, source.contentControls.length * 3)
  assert.equal(new Set(value.contentControls.map(control => control.id)).size, value.contentControls.length)
  assert.deepEqual(source, before); controls.validateDocumentContentControlRanges(value)
})

test('note replacement removes only old owned wrappers and comment segments in its atomic candidate', () => {
  const source = tree(), value = domain.createBlankDocument(), inserted = notes.insertDocumentNote(value, selected(value), 'footnote', { noteId: 'replace', ...source })
  const oldIds = hosts.documentModelForBlockHost(value, inserted.host).contentControls.map(control => control.id)
  const oldBlock = value.footnotes.replace.content.blocks[0].id, body = value.blocks[0]
  controls.createDocumentContentControl(value, { id: 'foreign', type: 'richText', blockId: body.id, startOffset: 0, endOffset: 0 })
  value.comments = { mixed: { id: 'mixed', anchor: [{ blockId: oldBlock, startOffset: 0, endOffset: 1, text: 'A' },
    { blockId: body.id, startOffset: 0, endOffset: 0, text: '' }], referenceText: 'A\n', messages: [], resolved: false, createdAt: '2026-11-25T00:00:00Z' } }
  assert.equal(notes.patchDocumentNote(value, 'footnote', 'replace', source), true)
  assert.equal(value.contentControls.some(control => oldIds.includes(control.id)), false); assert.ok(byId(value, 'foreign'))
  assert.deepEqual(value.comments.mixed.anchor.map(segment => segment.blockId), [body.id]); assert.equal(value.comments.mixed.referenceText, '')
  assertCopiedTree(source, { content: value.footnotes.replace.content, contentControls: hosts.documentModelForBlockHost(value, inserted.host).contentControls })
})

test('source foreign ranges, missing wrappers, invalid UTF-16 and a later invalid decoration never publish partial content', () => {
  for (const mutate of [
    source => { source.contentControls.find(control => control.id === 'child').parentId = 'foreign-parent' },
    source => { source.contentControls.find(control => control.id === 'root').blockRange.blockIds.push('foreign-block') },
    source => { source.contentControls.find(control => control.id === 'row-item').tableRowRange.rowIds = ['foreign-row'] },
    source => { source.contentControls.find(control => control.id === 'picture').metadata.pictureBlockId = 'foreign-image' },
    source => { source.contentControls.push({ id: 'bad-utf16', type: 'richText', blockId: 'head', startOffset: 2, endOffset: 3, parentId: 'root' }) },
    source => { source.contentControls.find(control => control.id === 'child').metadata = { officeStoryLockSource: true } },
  ]) { const source = tree(), value = domain.createBlankDocument(); mutate(source); const before = copy(value)
    assert.throws(() => notes.insertDocumentNote(value, selected(value), 'footnote', source)); assert.deepEqual(value, before) }
  const source = tree(), value = domain.createBlankDocument(), before = copy(value), bad = copy(source.contentControls)
  bad.find(control => control.id === 'child').blockId = 'foreign-block'
  assert.throws(() => notes.insertDocumentNote(value, selected(value), 'footnote', { ...source,
    settings: { separator: source.content, continuationSeparator: source.content },
    decorationContentControls: { separator: source.contentControls, continuationSeparator: bad } }))
  assert.deepEqual(value, before)
  assert.throws(() => content.prepareDocumentNoteContent(value, 'footnote', 'missing', source.content), /控件锚点|内容控件/)
})

test('locked story content, wrapper locks and locked fields reject replacement without removing any original resources', () => {
  for (const lock of ['story', 'wrapper', 'field']) {
    const source = tree(), value = domain.createBlankDocument(), inserted = notes.insertDocumentNote(value, selected(value), 'endnote', { noteId: 'locked', ...source })
    const note = value.endnotes.locked
    if (lock === 'story') note.content.locks = [{ id: 'lock', startOffset: 0, endOffset: 1, lock: 'contentLocked' }]
    if (lock === 'wrapper') value.contentControls.find(control => control.id === hosts.documentModelForBlockHost(value, inserted.host).contentControls[0].id).lock = 'sdtLocked'
    if (lock === 'field') fields.setDocumentTextFields(note.content.blocks[0], fields.documentTextFields(note.content.blocks[0]).map(field => ({ ...field, locked: true })))
    const before = copy(value); assert.throws(() => notes.patchDocumentNote(value, 'endnote', 'locked', source), /锁|locked/i); assert.deepEqual(value, before)
  }
})

test('decoration ownership payload is transient, same-story settings preserve IDs, and reset retains dormant provenance', () => {
  const source = tree(), value = domain.createBlankDocument()
  value.noteSettings = { footnote: { metadata: { source: { xml: '<w:footnotePr/>' }, officeDormantNoteDecorations: { version: 1, contents: [] } } } }
  assert.equal(settings.setDocumentNoteSettings(value, 'footnote', { startNumber: 3, separator: source.content }, undefined, undefined,
    { separator: source.contentControls }), true)
  const captured = copy(value), current = settings.getDocumentNoteSettings(value, 'footnote')
  const wrappers = hosts.documentModelForBlockHost(value, { kind: 'noteDecoration', noteKind: 'footnote', decoration: 'separator' }).contentControls
  assert.equal(settings.setDocumentNoteSettings(value, 'footnote', current, undefined, undefined, { separator: wrappers }), false)
  assert.deepEqual(value, captured); assert.equal(value.noteSettings.footnote.decorationContentControls, undefined)
  assert.equal(settings.setDocumentNoteSettings(value, 'footnote', null), true)
  assert.deepEqual(value.noteSettings.footnote.metadata, captured.noteSettings.footnote.metadata); assert.equal(value.noteSettings.footnote.separator, undefined)
  assert.equal(value.contentControls.length, 0)
  for (const args of [[{}, { separator: [] }], [null, { separator: [] }]]) {
    const before = copy(value); assert.throws(() => settings.setDocumentNoteSettings(value, 'footnote', args[0], undefined, undefined, args[1])); assert.deepEqual(value, before)
  }
})
