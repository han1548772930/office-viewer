import assert from 'node:assert/strict'
import { test } from 'vitest'
import { OfficeStudio } from '../src/api/office-studio.ts'
import { FDocument } from '../src/api/facade.ts'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection, createParagraph, createTable, runText, textTargets } from '../src/features/document/model/document-model.ts'
import { documentHeadingCollapseState, documentHeadingVisibility, normalizeDocumentHeadingSelection, setDocumentHeadingCollapsed,
  SET_DOCUMENT_HEADING_COLLAPSED_COMMAND, visibleDocumentTextTargets } from '../src/features/document/headings/heading-collapse.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { layoutDocument } from '../src/features/document/layout/document-layout.ts'
import { paragraphListMarkers } from '../src/features/document/model/document-lists.ts'

// Implementation drafts only. No test, frontend or browser validation executed.
const copy = value => structuredClone(value)
const paragraph = (id, text = id) => ({ ...createParagraph(text), id })
const heading = (id, level = 1) => ({ ...paragraph(id), type: 'heading', level })
const caret = (id, offset = 0) => collapsedSelection({ blockId: id, offset })
const measurer = { measure: text => text.length * 7 }
function fixture() {
  const model = createBlankDocument(); model.id = 'heading-fold'; model.metadata = { documentFlavor: 2, documentFlavorContract: 'univer' }
  model.blocks = [paragraph('before'), heading('parent'), paragraph('intro'), heading('child', 3), paragraph('child-body'),
    heading('sibling', 2), paragraph('sibling-body'), heading('next'), paragraph('next-body')]
  return { model, context: { documentId: model.id, userId: 'editor' } }
}
function studioFixture() {
  const { model } = fixture(), studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}

test('fold sections stop at the next same/higher heading and retain independent child fold flags without changing source text', () => {
  const { model, context } = fixture(), text = model.blocks.map(block => runText(block.runs)), ids = model.blocks.map(block => block.id)
  setDocumentHeadingCollapsed(model, caret('before'), 'child', true, context)
  setDocumentHeadingCollapsed(model, caret('before'), 'parent', true, context)
  const source = copy(model), visibility = documentHeadingVisibility(model)
  assert.deepEqual(visibility.blocks.map(block => block.id), ['before', 'parent', 'next', 'next-body'])
  assert.equal(visibility.hiddenOwners.get('child-body'), 'parent')
  assert.equal(visibility.sectionById.get('child').collapsed, true)
  assert.equal(visibility.sections.find(section => section.paragraphId === 'parent').endIndex, 7)
  assert.throws(() => setDocumentHeadingCollapsed(model, caret('before'), 'child', false, context), /上层/)
  assert.deepEqual(model, source)
  setDocumentHeadingCollapsed(model, caret('before'), 'parent', false, context)
  const expanded = documentHeadingVisibility(model)
  assert.equal(expanded.hiddenOwners.get('child-body'), 'child')
  assert.equal(expanded.hiddenOwners.has('sibling-body'), false)
  assert.deepEqual(model.blocks.map(block => runText(block.runs)), text)
  assert.deepEqual(model.blocks.map(block => block.id), ids)
})

test('collapsing a region intersecting a forward/backward selection moves to the visible owner, while view normalization maps endpoints', () => {
  for (const backward of [false, true]) {
    const { model, context } = fixture()
    const endpoints = [{ blockId: 'before', offset: 1 }, { blockId: 'next', offset: 2 }]
    const selection = { anchor: endpoints[backward ? 1 : 0], focus: endpoints[backward ? 0 : 1] }
    const next = setDocumentHeadingCollapsed(model, selection, 'parent', true, context)
    assert.deepEqual(next, caret('parent', 'parent'.length))
    const normalized = normalizeDocumentHeadingSelection(model, { anchor: { blockId: 'child-body', offset: 2 }, focus: { blockId: 'next', offset: 0 } })
    assert.deepEqual(normalized.anchor, { blockId: 'parent', offset: 'parent'.length })
    assert.deepEqual(normalized.focus, { blockId: 'next', offset: 0 })
    assert.deepEqual(visibleDocumentTextTargets(model).map(target => target.blockId), ['before', 'parent', 'next', 'next-body'])
  }
})

test('cell and column heading flows remain independent and use canonical nested text targets', () => {
  const { model, context } = fixture(), table = createTable(1, 2)
  table.rows[0].cells[0].blocks = [heading('cell-left'), paragraph('cell-left-body')]
  table.rows[0].cells[1].blocks = [heading('cell-right'), paragraph('cell-right-body')]
  const tableBlock = { ...paragraph('table'), type: 'table', runs: [], metadata: { table } }
  const columns = { ...paragraph('columns'), type: 'columns', runs: [], columns: { version: 1, layout: 'fixed', responsive: 'shrink', gap: 20,
    columns: [{ id: 'left', widthRatio: 1, blocks: [heading('col-left'), paragraph('col-left-body')] },
      { id: 'right', widthRatio: 1, blocks: [heading('col-right'), paragraph('col-right-body')] }] } }
  model.blocks = [tableBlock, columns, paragraph('outside')]
  const before = copy(model)
  setDocumentHeadingCollapsed(model, caret('outside'), 'cell-left', true, context)
  setDocumentHeadingCollapsed(model, caret('outside'), 'col-left', true, context)
  const visibility = documentHeadingVisibility(model)
  assert.deepEqual(visibility.sectionById.get('cell-left').target, { blockId: 'table', row: 0, col: 0, cellBlockId: 'cell-left' })
  assert.equal(visibility.hiddenOwners.get('cell-left-body'), 'cell-left')
  assert.equal(visibility.hiddenOwners.has('cell-right-body'), false)
  assert.equal(visibility.hiddenOwners.get('col-left-body'), 'col-left')
  assert.equal(visibility.hiddenOwners.has('col-right-body'), false)
  const layout = layoutDocument(model, measurer), rendered = new Set(layout.lines.map(line => line.target.cellBlockId ?? line.target.blockId))
  assert.equal(rendered.has('cell-left-body'), false); assert.equal(rendered.has('cell-right-body'), true)
  assert.equal(rendered.has('col-left-body'), false); assert.equal(rendered.has('col-right-body'), true)
  assert.equal(model.blocks[0].metadata.table.rows[0].cells[0].blocks.length, before.blocks[0].metadata.table.rows[0].cells[0].blocks.length)
  assert.equal(model.blocks[1].columns.columns[0].blocks.length, before.blocks[1].columns.columns[0].blocks.length)
})

test('derived Canvas layout hides text, table, images and paragraphs while preserving authored objects and source numbering', () => {
  const { model, context } = fixture(), hiddenList = paragraph('list-hidden', 'hidden item'), visibleList = paragraph('list-visible', 'visible item')
  hiddenList.type = 'numbered'; hiddenList.list = { id: 'shared-list', preset: 'ORDER_LIST' }
  visibleList.type = 'numbered'; visibleList.list = { id: 'shared-list', preset: 'ORDER_LIST' }
  const image = { ...paragraph('image'), type: 'image', runs: [], metadata: { image: { id: 'photo', src: 'data:image/png;base64,AA==', width: 40, height: 20, alt: 'photo' } } }
  const table = { ...paragraph('hidden-table'), type: 'table', runs: [], metadata: { table: createTable(1, 1) } }
  model.blocks = [heading('parent'), hiddenList, image, table, heading('next'), visibleList]
  const markers = paragraphListMarkers(model.blocks), source = copy(model)
  assert.equal(markers.get('list-visible').ordinal, 2)
  setDocumentHeadingCollapsed(model, caret('parent'), 'parent', true, context)
  const layout = layoutDocument(model, measurer)
  assert.equal(layout.lines.some(line => ['list-hidden', 'hidden-table'].includes(line.target.blockId)), false)
  assert.equal(layout.images.some(item => item.blockId === 'image'), false)
  assert.equal((layout.tableFragments ?? []).some(item => item.tableId === 'hidden-table'), false)
  const visible = layout.lines.find(line => line.target.blockId === 'list-visible')
  assert.equal(visible.marker, markers.get('list-visible')?.marker)
  assert.deepEqual(model.blocks.map(block => block.id), source.blocks.map(block => block.id))
  assert.deepEqual(model.blocks.find(block => block.id === 'image'), source.blocks.find(block => block.id === 'image'))
})

test('flags are inert outside modern mode and on nonheadings, and empty/duplicate/hidden headings do not authorize edits', () => {
  const { model, context } = fixture()
  model.blocks[0].metadata.headingCollapsed = true
  model.blocks.at(-1).metadata.headingCollapsed = true
  model.blocks.find(block => block.id === 'parent').metadata.headingCollapsed = true
  model.metadata.documentFlavor = 1
  assert.equal(documentHeadingVisibility(model).blocks, model.blocks)
  assert.deepEqual(documentHeadingVisibility(model).sections, [])
  assert.equal(documentHeadingCollapseState(model, 'parent', true, context).available, false)
  model.metadata.documentFlavor = 2
  assert.equal(documentHeadingCollapseState(model, 'before', false, context).available, false)
  assert.equal(documentHeadingCollapseState(model, 'child', false, context).available, false)
  model.blocks.push(heading('empty'))
  assert.equal(documentHeadingCollapseState(model, 'empty', true, context).available, false)
  model.blocks.push(heading('next'))
  assert.equal(documentHeadingCollapseState(model, 'next', true, context).available, false)
})

test('the command enforces heading edit ACL and content locks before touching the candidate', () => {
  const { model, context } = fixture()
  model.permissionRules = [{ id: 'deny-heading', scope: 'entity', targetId: 'parent', action: 'edit', effect: 'deny', userIds: [] }]
  const before = copy(model)
  assert.equal(documentHeadingCollapseState(model, 'parent', true, context).available, false)
  assert.throws(() => setDocumentHeadingCollapsed(model, caret('intro'), 'parent', true, context))
  assert.deepEqual(model, before)
  model.permissionRules = []
  model.contentControls = [{ id: 'locked-heading', type: 'richText', blockId: 'parent', startOffset: 0, endOffset: 6, lock: 'contentLocked' }]
  const locked = copy(model)
  assert.equal(documentHeadingCollapseState(model, 'parent', true, context).available, false)
  assert.throws(() => setDocumentHeadingCollapsed(model, caret('intro'), 'parent', true, context), /locked/)
  assert.deepEqual(model, locked)
})

test('Enter expands before splitting and boundary Delete/Backspace open hidden content instead of deleting it', () => {
  for (const direction of [-1, 1]) {
    const { model, context } = fixture()
    setDocumentHeadingCollapsed(model, caret('before'), 'parent', true, context)
    const source = model.blocks.map(block => runText(block.runs))
    const selected = direction === 1 ? caret('parent', 6) : caret('next', 0)
    const result = executeDocumentCommand(model, selected, { id: 'doc.command.delete', direction }, context)
    assert.deepEqual(result, selected)
    assert.equal(model.blocks.find(block => block.id === 'parent').metadata.headingCollapsed, false)
    assert.deepEqual(model.blocks.map(block => runText(block.runs)), source)
  }
  const { model, context } = fixture()
  setDocumentHeadingCollapsed(model, caret('before'), 'parent', true, context)
  executeDocumentCommand(model, caret('parent', 6), { id: 'doc.command.insert-text', text: '\n' }, context)
  assert.equal(model.blocks.find(block => block.id === 'parent').metadata.headingCollapsed, false)
  assert.ok(model.blocks.some(block => block.id === 'child-body'))
  assert.equal(documentHeadingVisibility(model).hiddenOwners.has('child-body'), false)
})

test('Studio uses one content history, restores the hidden pre-fold selection on undo and the visible owner on redo', () => {
  const { studio, document } = studioFixture()
  try {
    document.setInputSelection(caret('child-body', 2))
    const before = copy(document.getModel()), original = copy(document.getInputState().location)
    assert.equal(document.setHeadingCollapsed('parent', true), true)
    assert.deepEqual(document.getInputState().location.selection, caret('parent', 6))
    assert.equal(studio.editor.history.units[document.id].undo.length, 1)
    const after = copy(document.getModel())
    assert.equal(document.setHeadingCollapsed('parent', true), false)
    assert.equal(studio.editor.history.units[document.id].undo.length, 1)
    document.undo(); assert.deepEqual(document.getModel(), before); assert.deepEqual(document.getInputState().location, original)
    document.redo(); assert.deepEqual(document.getModel(), after); assert.deepEqual(document.getInputState().location.selection, caret('parent', 6))
  } finally { studio.dispose() }
})

test('public snapshot and model reopening keep nested flags and every hidden paragraph', () => {
  const { studio, document } = studioFixture()
  let reopened
  try {
    document.setInputSelection(caret('before'))
    document.setHeadingCollapsed('child', true); document.setHeadingCollapsed('parent', true)
    const snapshot = copy(document.save())
    assert.equal(snapshot.body.paragraphs.find(paragraph => paragraph.paragraphId === 'parent').headingCollapsed, true)
    assert.equal(snapshot.body.paragraphs.find(paragraph => paragraph.paragraphId === 'child').headingCollapsed, true)
    reopened = new OfficeStudio({ document: snapshot, currentUserId: 'editor' }); reopened.activateDocument(document.id)
    const next = new FDocument(reopened, document.id)
    assert.equal(next.getModel().blocks.some(block => block.id === 'child-body'), true)
    assert.equal(documentHeadingVisibility(next.getModel()).hiddenOwners.get('child-body'), 'parent')
    next.setInputSelection(caret('child-body', 2))
    assert.deepEqual(next.getInputState().location.selection, caret('parent', 6))
    assert.equal(next.setHeadingCollapsed('parent', false), true)
    assert.equal(documentHeadingVisibility(next.getModel()).hiddenOwners.get('child-body'), 'child')
  } finally { reopened?.dispose(); studio.dispose() }
})

test('BeforeCommand cancellation/navigation cannot redirect a captured fold to a new document or input generation', () => {
  const { studio, document } = studioFixture()
  try {
    document.setInputSelection(caret('intro', 2))
    const before = copy(document.getModel())
    let canceled = true
    const listener = studio.events.addEvent('BeforeCommandExecute', event => {
      if (event.id !== SET_DOCUMENT_HEADING_COLLAPSED_COMMAND) return
      if (canceled) event.cancel = true
      else document.setInputSelection(caret('next'))
    })
    assert.throws(() => document.setHeadingCollapsed('parent', true), /canceled/)
    canceled = false
    assert.throws(() => document.setHeadingCollapsed('parent', true), /已更改/)
    assert.deepEqual(document.getModel(), before)
    assert.equal(studio.editor.history.units[document.id]?.undo.length ?? 0, 0)
    listener.dispose()
  } finally { studio.dispose() }
})
