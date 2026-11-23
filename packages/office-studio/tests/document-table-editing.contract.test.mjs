import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createPinia } from 'pinia'
import { createServer } from 'vite'

// DOC-06 contract drafts. Execute only during the agreed unified verification phase.
let server, domain, model, grid, selections, navigation, commands, mutations, style, continuation, blockPaste, clipboard, fields, search, input, indexed, facade, tableApi, OfficeStudio
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-table-editing', optimizeDeps: { noDiscovery: true }, server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  model = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  grid = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-table-grid.ts')
  selections = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-table-selection.ts')
  navigation = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-table-navigation.ts')
  commands = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
  mutations = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/table-mutations.ts')
  style = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-table-style.ts')
  continuation = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-table-continuations.ts')
  blockPaste = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-block-paste.ts')
  clipboard = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-clipboard.ts')
  fields = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-fields.ts')
  search = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-search.ts')
  input = await server.ssrLoadModule('/packages/office-studio/src/features/document/stores/document-input-store.ts')
  indexed = await server.ssrLoadModule('/packages/office-studio/src/api/document-indexed-input.ts')
  facade = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts')
  tableApi = await server.ssrLoadModule('/packages/office-studio/src/api/document-tables.ts')
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const paragraph = (id, text, marks = {}) => ({ ...model.createParagraph(text, marks), id })
const rect = (startRow, startColumn, endRow = startRow, endColumn = startColumn) => ({ startRow, endRow, startColumn, endColumn })
const sourceOf = (document, id = 'table') => model.tableForBlock(model.findDocumentBlock(document, id))
const ownerAt = (document, row, column, id = 'table') => grid.documentTableGrid(sourceOf(document, id)).cellAt(row, column)
const view = (document, range, kind = 'range', pageAffinity) => selections.createDocumentTableViewSelection(document, 'table', range, kind, pageAffinity)
const execute = (document, selection, command) => commands.executeDocumentCommand(document, selection, command, { documentId: document.id, userId: 'editor' })
function tableBlock(id, rows, columns, prefix = id) {
  const table = model.createTable(rows, columns)
  table.rows.forEach((row, r) => {
    row.id = `${prefix}-row-${r}`
    row.cells.forEach((cell, c) => { cell.id = `${prefix}-cell-${r}-${c}`; cell.blocks = [paragraph(`${prefix}-p-${r}-${c}`, `${r},${c}`)] })
  })
  table.columnWidths = Array.from({ length: columns }, () => 100)
  return { ...paragraph(id, ''), type: 'table', runs: [], metadata: { table } }
}
function fixture(rows = 3, columns = 3) {
  return { ...domain.createBlankDocument(), id: 'table-contract', blocks: [paragraph('before', 'Before'), tableBlock('table', rows, columns), paragraph('after', 'After')], contentControls: [], modified: false }
}
function continuationFixture() {
  const document = fixture(2, 2), table = sourceOf(document), owner = table.rows[0].cells[0]
  owner.rowSpan = 2
  table.rows[1].cells.splice(0, 1)
  owner.style.nativeContinuationContent = { 1: { version: 1, ownerCellId: owner.id, rowId: table.rows[1].id, colSpan: 1,
    blocks: [paragraph('continued', 'Continuation')], source: { retained: 'tc' } } }
  owner.style.nativeContinuations = { 1: { ownerCellId: owner.id, rowId: table.rows[1].id, colSpan: 1, retained: 'native-properties' } }
  return document
}
function open(document) {
  const studio = new OfficeStudio({ document: copy(document), currentUserId: 'editor' })
  studios.add(studio); studio.activateDocument(document.id)
  return { studio, document: new facade.FDocument(studio, document.id), table: new tableApi.FDocumentTable(new facade.FDocument(studio, document.id), 'table') }
}

test('covered physical rows and logical anchors form a read-only rectangular grid', () => {
  const document = fixture(3, 3), table = sourceOf(document), owner = table.rows[0].cells[0]
  owner.rowSpan = 3; owner.colSpan = 3
  table.rows[0].cells = [owner]; table.rows[1].cells = []; table.rows[2].cells = []
  const before = copy(table), resolved = grid.documentTableGrid(table)
  assert.equal(resolved.anchors.length, 1)
  assert.equal(resolved.cellAt(2, 2).cell.id, owner.id)
  assert.equal(resolved.storedCell(2, 0), undefined)
  assert.deepEqual(grid.expandDocumentTableRange(resolved, rect(2, 2)), rect(0, 0, 2, 2))
  assert.deepEqual(table, before)
})

test('grid and paste budgets reject before manufacturing large cell collections', () => {
  const block = tableBlock('huge', 1, 1), table = model.tableForBlock(block), owner = table.rows[0].cells[0]
  owner.rowSpan = 500; owner.colSpan = 512
  table.rows = Array.from({ length: 500 }, (_, r) => ({ id: `huge-row-${r}`, cells: r === 0 ? [owner] : [] }))
  table.columnWidths = []
  assert.equal(grid.documentTableGrid(table).anchors.length, 1)
  assert.throws(() => blockPaste.copyDocumentBlocksForPaste([block]), /单次单元格预算/)
  const saved = copy(block)
  owner.colSpan = 513
  assert.throws(() => grid.documentTableGrid(table), /512 列/)
  owner.colSpan = 512
  assert.deepEqual(block, saved)
  const document = fixture(1, 1), original = copy(document)
  assert.throws(() => execute(document, model.collapsedSelection({ blockId: 'before', offset: 0 }), { id: 'doc.command.insert-table', rows: 600, cols: 512 }), /预算/)
  assert.deepEqual(document, original)
})

test('object-only table selection does not insert text and cloned view ranges are detached', () => {
  const document = fixture(1, 2), table = sourceOf(document)
  for (const [index, cell] of table.rows[0].cells.entries()) cell.blocks = [{ ...paragraph(`image-${index}`, ''), type: 'image', runs: [], metadata: { image: { src: 'data:image/png;base64,AA==', width: 10, height: 10 } } }]
  const before = copy(document), selected = view(document, rect(0, 0, 0, 1), 'table')
  assert.equal(selected.anchor.blockId, 'before')
  assert.equal(model.selectionCollapsed(selected), false)
  assert.deepEqual(selections.documentTableSelectedTextTargets(document, selected.tableSelection), [])
  const detached = selections.cloneDocumentSelection(selected)
  detached.tableSelection.range.endColumn = 0
  assert.equal(selected.tableSelection.range.endColumn, 1)
  assert.deepEqual(document, before)
})

test('Shift navigation expands and contracts a logical range while skipping a merged owner', () => {
  const document = fixture(2, 4), table = sourceOf(document)
  table.rows[0].cells[1].colSpan = 2; table.rows[0].cells.splice(2, 1)
  const start = view(document, rect(0, 0), 'cell')
  const grown = navigation.navigateDocumentTableSelection(document, start, 'right', true)
  assert.deepEqual(grown.selection.tableSelection.range, rect(0, 0, 0, 2))
  const shrunk = navigation.navigateDocumentTableSelection(document, grown.selection, 'left', true, false, grown.cursor)
  assert.deepEqual(shrunk.selection.tableSelection.range, rect(0, 0))
  const skipped = navigation.navigateDocumentTableSelection(document, grown.selection, 'right', false, false, grown.cursor)
  assert.deepEqual(skipped.selection.tableSelection.range, rect(0, 3))
})

test('structural format, clear, TSV and search operate on actual cells, including continuation paragraphs', () => {
  const document = continuationFixture(), selected = view(document, rect(0, 0, 1, 0), 'column')
  const untouched = copy(ownerAt(document, 0, 1).cell)
  execute(document, selected, { id: 'doc.command.format-text', patch: { bold: true } })
  assert.equal(model.findDocumentBlock(document, 'continued').runs.every(run => run.bold), true)
  assert.equal(model.findDocumentBlock(document, 'table-p-0-0').runs.every(run => run.bold), true)
  assert.deepEqual(ownerAt(document, 0, 1).cell, untouched)
  const payload = clipboard.documentTableClipboardData(document, selected.tableSelection)
  assert.match(payload.text, /Continuation/)
  assert.match(payload.html, /data-office-table-continuation="v1"/)
  assert.equal(search.findDocumentMatches(document, '1', { selection: selected }).length, 0)
  assert.deepEqual(search.findDocumentMatches(document, 'Continuation', { selection: selected })[0].target,
    { blockId: 'table', row: 0, col: 0, cellBlockId: 'continued' })
  execute(document, selected, { id: 'doc.command.table-set-range-text', mode: 'clear' })
  assert.equal(model.runText(model.runsForTarget(document, { blockId: 'table', row: 0, col: 0, cellBlockId: 'table-p-0-0' })), '')
  assert.deepEqual(ownerAt(document, 0, 1).cell, untouched)
})

test('mixed row, margin and border patches preserve every unspecified property, including undefined keys', () => {
  const document = fixture(2, 2), table = sourceOf(document)
  table.rows[0].height = 35; table.rows[0].heightRule = 'exact'; table.rows[0].allowSplit = false
  table.rows[1].height = 55; table.rows[1].heightRule = 'atLeast'; table.rows[1].allowSplit = true
  for (const [r, row] of table.rows.entries()) for (const [c, cell] of row.cells.entries()) {
    cell.style.margin = { top: 2 + r, right: 4 + c, bottom: 6 + r, left: 8 + c }
    cell.style.borders = { top: { width: 1 + c, style: c ? 'dash' : 'dot', color: c ? '#00aa00' : '#aa0000' } }
  }
  execute(document, view(document, rect(0, 0, 1, 1)), { id: 'doc.command.table-resize-row', height: 70, preserveHeightRule: true, allowSplit: undefined })
  assert.deepEqual(sourceOf(document).rows.map(row => [row.height, row.heightRule, row.allowSplit]), [[70, 'exact', false], [70, 'atLeast', true]])
  const marginBefore = sourceOf(document).rows.flatMap(row => row.cells).map(cell => copy(cell.style.margin))
  execute(document, view(document, rect(0, 0, 1, 1)), { id: 'doc.command.table-set-cell-margin', marginPatch: { left: 3, top: undefined } })
  assert.deepEqual(sourceOf(document).rows.flatMap(row => row.cells).map((cell, index) => cell.style.margin), marginBefore.map(margin => ({ ...margin, left: 3 })))
  execute(document, view(document, rect(0, 0, 0, 1)), { id: 'doc.command.table-set-border', preset: 'top', borderPatch: { color: '#123456', width: undefined } })
  assert.deepEqual(sourceOf(document).rows[0].cells.map(cell => cell.style.borders.top), [
    { width: 1, style: 'dot', color: '#123456' }, { width: 2, style: 'dash', color: '#123456' },
  ])
})

test('shared border updates include an outside neighbour in permission checks and reject without partial publication', () => {
  const document = fixture(1, 2), selected = view(document, rect(0, 0), 'cell')
  document.permissionRules = [{ id: 'neighbour', scope: 'paragraph', targetId: 'table-p-0-1', action: 'edit', effect: 'deny', userIds: [] }]
  const before = copy(document)
  assert.throws(() => execute(document, selected, { id: 'doc.command.table-set-border', preset: 'right', border: { width: 3, style: 'dash', color: '#f00' } }), /permission denied/)
  assert.deepEqual(document, before)
})

test('continuation input and indexed addressing change its own body without flattening it into the owner', () => {
  const document = continuationFixture(), first = ownerAt(document, 0, 0).cell.blocks[0], firstBefore = copy(first)
  const target = model.textTargets(document).find(candidate => candidate.cellBlockId === 'continued')
  const caret = model.collapsedSelection({ ...target, offset: 4 })
  execute(document, caret, { id: 'doc.command.insert-text', text: '中😀' })
  const owner = ownerAt(document, 0, 0).cell, continued = continuation.documentTableContinuationEntries(owner)[0]
  assert.equal(model.runText(continued.content.blocks[0].runs), 'Cont中😀inuation')
  assert.deepEqual(owner.blocks[0], firstBefore)
  assert.equal(continued.content.rowId, sourceOf(document).rows[1].id)
  const absolute = fields.documentTextFieldParagraphs(document).find(entry => entry.block.id === 'continued').absoluteStart
  assert.deepEqual(indexed.documentSelectionAtBodyRange(document, absolute + 4, absolute + 4).focus, { ...target, offset: 4 })
})

test('unmerge restores independent continuation content into the physical row without duplicating paragraphs', () => {
  const document = continuationFixture(), before = copy(document), bodyIds = fields.documentTextFieldParagraphs(document).map(entry => entry.block.id)
  execute(document, view(document, rect(0, 0, 1, 0)), { id: 'doc.command.table-unmerge-cells' })
  assert.equal(ownerAt(document, 1, 0).cell.blocks[0].id, 'continued')
  assert.equal(ownerAt(document, 0, 0).cell.blocks[0].id, 'table-p-0-0')
  assert.equal(sourceOf(document).rows.every(row => row.cells.every(cell => cell.rowSpan === 1 && cell.colSpan === 1)), true)
  const ids = fields.documentTextFieldParagraphs(document).map(entry => entry.block.id)
  assert.equal(ids.filter(id => id === 'continued').length, 1)
  assert.equal(bodyIds.every(id => ids.includes(id)), true)
  assert.equal(model.runText(model.findDocumentBlock(document, 'continued').runs), model.runText(model.findDocumentBlock(before, 'continued').runs))
})

test('insert and delete rekey continuation offsets by stable row identity and preserve the borrowed snapshot', () => {
  const document = continuationFixture(), before = copy(document), continuedRow = sourceOf(document).rows[1].id
  execute(document, view(document, rect(0, 0)), { id: 'doc.command.table-insert-row', row: 0, position: 'after' })
  let owner = ownerAt(document, 0, 0).cell
  assert.equal(owner.rowSpan, 3)
  assert.deepEqual(continuation.documentTableContinuationEntries(owner).map(entry => [entry.offset, entry.content.rowId]), [[2, continuedRow]])
  execute(document, view(document, rect(1, 0, 1, 1), 'row'), { id: 'doc.command.table-delete-rows', range: rect(1, 0, 1, 1) })
  owner = ownerAt(document, 0, 0).cell
  assert.deepEqual(continuation.documentTableContinuationEntries(owner).map(entry => [entry.offset, entry.content.rowId]), [[1, continuedRow]])
  assert.equal(model.runText(model.findDocumentBlock(before, 'continued').runs), 'Continuation')
  const saved = copy(document)
  execute(document, view(document, rect(1, 1), 'cell'), { id: 'doc.command.table-delete-rows', range: rect(1, 0, 1, 1) })
  assert.equal(model.findDocumentBlock(document, 'continued'), undefined)
  assert.equal(model.findDocumentBlock(saved, 'continued').id, 'continued')
})

test('copy remaps continuation, nested table and cross-paragraph field identities without changing its input', () => {
  const document = continuationFixture(), cell = ownerAt(document, 0, 0).cell
  cell.style.nativeContinuationContent[1].blocks.push(paragraph('continued-end', 'End'), tableBlock('nested', 1, 1))
  fields.setDocumentTextFields(cell.style.nativeContinuationContent[1].blocks[0], [{ id: 'field', field: 'DATE', instruction: 'DATE \\@ "yyyy-MM-dd"', cachedResult: 'Continuation\nEnd', startOffset: 0, endOffset: 3, endBlockId: 'continued-end' }])
  const source = [model.findDocumentBlock(document, 'table')], saved = copy(source), result = blockPaste.copyDocumentBlocksForPaste(source)
  assert.deepEqual(source, saved)
  const table = model.tableForBlock(result[0]), copied = continuation.documentTableContinuationEntries(table.rows[0].cells[0])[0]
  assert.equal(copied.content.ownerCellId, table.rows[0].cells[0].id)
  assert.equal(copied.content.rowId, table.rows[1].id)
  assert.notEqual(copied.content.blocks[0].id, 'continued')
  const field = fields.documentTextFieldRanges({ blocks: result })[0]
  assert.notEqual(field.field.id, 'field')
  assert.equal(field.field.endBlockId, copied.content.blocks[1].id)
  assert.notEqual(copied.content.blocks[2].id, 'nested')
})

test('a continuation identity mismatch is rejected before copy can silently discard source content', () => {
  const document = continuationFixture(), cell = ownerAt(document, 0, 0).cell
  cell.style.nativeContinuationContent[1].rowId = 'wrong-row'
  const before = copy(document)
  assert.throws(() => blockPaste.copyDocumentBlocksForPaste([model.findDocumentBlock(document, 'table')]), /稳定行身份/)
  assert.deepEqual(document, before)
})

test('rich paste rebuilds a complete source topology while keeping surviving destination anchor IDs', () => {
  const document = fixture(2, 2), table = sourceOf(document), oldId = table.rows[0].cells[0].id
  table.rows[0].cells[0].colSpan = 2; table.rows[0].cells.pop()
  const source = tableBlock('source', 2, 2), nested = tableBlock('nested-source', 1, 1)
  model.tableForBlock(source).rows[0].cells[1].blocks.push(nested)
  const savedSource = copy(source), savedDestination = copy(document)
  const planned = mutations.prepareDocumentTablePasteMutation(document, view(document, rect(0, 0)), [source], { documentId: document.id, userId: 'editor' })
  assert.deepEqual(document, savedDestination); assert.deepEqual(source, savedSource)
  const result = grid.documentTableGrid(sourceOf(planned.model))
  assert.equal(result.anchors.length, 4)
  assert.equal(result.cellAt(0, 0).cell.id, oldId)
  assert.notEqual(result.cellAt(0, 1).cell.id, oldId)
  assert.equal(result.cellAt(0, 1).cell.blocks.some(block => block.type === 'table'), true)
  assert.deepEqual(planned.selection.tableSelection.range, rect(0, 0, 1, 1))
  assert.equal(planned.selection.tableSelection.pageAffinity, undefined)
})

test('multi-cell source formatting and destination formatting are evaluated for each actual destination cell', () => {
  const document = fixture(2, 2), source = tableBlock('source', 2, 2), table = sourceOf(document)
  table.rows.forEach((row, r) => row.cells.forEach((cell, c) => {
    cell.style.fill = r ? '#ddffdd' : '#ddddff'; cell.style.margin = { top: 1 + r, right: 2 + c, bottom: 3 + r, left: 4 + c }
    cell.blocks[0].style.align = c ? 'right' : 'center'
    cell.blocks[0].runs = [model.createRun('Target', { fontSize: 11 + r + c, fontFamily: c ? 'Georgia' : 'Arial', color: '#123456' })]
  }))
  const incoming = model.tableForBlock(source)
  incoming.style = { fill: '#ffcccc', cellMargin: { top: 8, right: 9, bottom: 10, left: 11 } }
  incoming.rows.forEach(row => row.cells.forEach(cell => { cell.blocks[0].runs = [model.createRun('Source', { fontSize: 40, italic: true, color: '#ff0000' })] }))
  const sourceResult = mutations.prepareDocumentTablePasteMutation(document, view(document, rect(0, 0)), [source], { documentId: document.id, userId: 'editor' }, 'source').model
  const destinationResult = mutations.prepareDocumentTablePasteMutation(document, view(document, rect(0, 0)), [source], { documentId: document.id, userId: 'editor' }, 'destination').model
  for (let r = 0; r < 2; r++) for (let c = 0; c < 2; c++) {
    const original = ownerAt(document, r, c).cell, matched = ownerAt(destinationResult, r, c).cell, formatted = ownerAt(sourceResult, r, c).cell
    assert.equal(style.documentTableCellFill(sourceOf(sourceResult), formatted), '#ffcccc')
    assert.equal(formatted.blocks[0].runs[0].fontSize, 40)
    assert.equal(matched.style.fill, original.style.fill)
    assert.deepEqual(matched.style.margin, original.style.margin)
    assert.equal(matched.blocks[0].style.align, original.blocks[0].style.align)
    assert.equal(matched.blocks[0].runs[0].fontSize, 11 + r + c)
    assert.equal(matched.blocks[0].runs[0].fontFamily, c ? 'Georgia' : 'Arial')
    assert.equal(matched.blocks[0].runs[0].italic, true)
  }
})

test('a single-cell rich paste keeps an existing destination rowspan and colspan', () => {
  const document = continuationFixture(), cell = ownerAt(document, 0, 0).cell, originalId = cell.id
  const result = execute(document, view(document, rect(0, 0)), { id: 'doc.command.paste-blocks', blocks: [paragraph('paste', 'Rich', { bold: true })] })
  const current = ownerAt(document, 0, 0).cell
  assert.equal(current.id, originalId); assert.equal(current.rowSpan, 2); assert.equal(current.colSpan, 1)
  assert.equal(model.runText(current.blocks[0].runs), 'Rich')
  assert.equal(current.style.nativeContinuationContent, undefined)
  assert.deepEqual(result.tableSelection.range, rect(0, 0, 1, 0))
})

test('paste overflow and a footprint cutting an existing merge leave both source and destination unchanged', () => {
  for (const kind of ['overflow', 'merge-cut']) {
    const document = fixture(2, 3), table = sourceOf(document)
    if (kind === 'merge-cut') { table.rows[0].cells[1].colSpan = 2; table.rows[0].cells.pop() }
    const source = tableBlock('source', kind === 'overflow' ? 3 : 1, 2), before = copy(document), sourceBefore = copy(source)
    assert.throws(() => execute(document, view(document, rect(0, 0)), { id: 'doc.command.paste-blocks', blocks: [source] }), /超过现有表格|不能切开已有合并/)
    assert.deepEqual(document, before); assert.deepEqual(source, sourceBefore)
  }
})

test('rich replacement rejects zero-length control and field locks even when visible text is unchanged', () => {
  for (const kind of ['control', 'field']) {
    const document = fixture(1, 1), block = ownerAt(document, 0, 0).cell.blocks[0]
    block.runs = [model.createRun('')]
    if (kind === 'control') document.contentControls = [{ id: 'locked', type: 'text', blockId: block.id, startOffset: 0, endOffset: 0, lock: 'contentLocked' }]
    else fields.setDocumentTextFields(block, [{ id: 'locked', field: 'DATE', startOffset: 0, endOffset: 0, locked: true }])
    const before = copy(document)
    assert.throws(() => execute(document, view(document, rect(0, 0)), { id: 'doc.command.paste-blocks', blocks: [paragraph('empty-source', '', { bold: true })] }), /锁定/)
    assert.deepEqual(document, before)
  }
})

test('rich paste permissions use the changed footprint within a broader structural selection', () => {
  const document = fixture(3, 3), before = copy(ownerAt(document, 2, 2).cell)
  document.permissionRules = [{ id: 'outside', scope: 'paragraph', targetId: 'table-p-2-2', action: 'edit', effect: 'deny', userIds: [] }]
  execute(document, view(document, rect(0, 0, 2, 2), 'table'), { id: 'doc.command.paste-blocks', blocks: [paragraph('source', 'Only first')] })
  assert.equal(model.runText(ownerAt(document, 0, 0).cell.blocks[0].runs), 'Only first')
  assert.deepEqual(ownerAt(document, 2, 2).cell, before)
})

test('dropdown option identity remains distinct for duplicate labels and validates before text publication', () => {
  const document = fixture(1, 1), selected = view(document, rect(0, 0))
  execute(document, selected, { id: 'doc.command.table-set-column-type', column: 0, config: { type: 'dropdown', dropdownOptions: [
    { id: 'first', label: 'Same', color: '#f00' }, { id: 'second', label: 'Same', color: '#00f' },
  ] } })
  execute(document, selected, { id: 'doc.command.table-set-cell-text', row: 0, column: 0, text: 'Same', optionId: 'first' })
  execute(document, selected, { id: 'doc.command.table-set-cell-text', row: 0, column: 0, text: 'Same', optionId: 'second' })
  assert.equal(ownerAt(document, 0, 0).cell.style.selectedDropdownOptionId, 'second')
  const before = copy(document)
  assert.throws(() => execute(document, selected, { id: 'doc.command.table-set-cell-text', row: 0, column: 0, text: 'Mismatch', optionId: 'second' }), /下拉选项/)
  assert.deepEqual(document, before)
})

test('field, content-control and comment targets retain paragraph identity during table moves', () => {
  const document = fixture(3, 2), block = ownerAt(document, 1, 0).cell.blocks[0]
  fields.setDocumentTextFields(block, [{ id: 'field', field: 'DATE', startOffset: 0, endOffset: 3, cachedResult: '1,0' }])
  document.contentControls = [{ id: 'control', type: 'text', blockId: block.id, startOffset: 0, endOffset: 3 }]
  document.comments = { thread: { id: 'thread', anchor: [{ blockId: block.id, startOffset: 0, endOffset: 3, text: '1,0' }], referenceText: '1,0', messages: [], resolved: false, createdAt: '2026-11-24T00:00:00Z', updatedAt: '2026-11-24T00:00:00Z' } }
  execute(document, view(document, rect(1, 0, 1, 1), 'row'), { id: 'doc.command.table-move-rows', fromRow: 1, count: 1, toRow: 2, position: 'after' })
  assert.equal(ownerAt(document, 2, 0).cell.blocks[0].id, block.id)
  assert.equal(fields.documentTextFieldRanges(document)[0].startParagraph.block.id, block.id)
  assert.equal(document.contentControls[0].blockId, block.id)
  assert.equal(document.comments.thread.anchor[0].blockId, block.id)
  assert.equal(search.findDocumentMatches(document, '1,0')[0].target.row, 2)
})

test('input replacement rebases structural views by cell ID and clears derived page affinity', () => {
  const document = fixture(3, 2), previous = copy(document), selected = view(document, rect(1, 0, 1, 1), 'row', 2)
  const store = input.useDocumentInputStore(createPinia())
  store.select(document.id, { kind: 'body', selection: selected }, 4)
  execute(document, selected, { id: 'doc.command.table-move-rows', fromRow: 1, count: 1, toRow: 2, position: 'after' })
  store.invalidate(document.id)
  assert.equal(store.revalidateModel(document.id, document, previous), true)
  const state = store.get(document.id)
  assert.deepEqual(state.location.selection.tableSelection.range, rect(2, 0, 2, 1))
  assert.equal(state.location.selection.tableSelection.pageAffinity, undefined)
  assert.equal(state.location.selection.focus.pageAffinity, undefined)
  assert.equal(state.valid, false)
  assert.deepEqual(store.getMarks(document.id, 5), {})
})

test('continuation caret relocation keeps a guarded refresh source without automatically restoring stale marks', () => {
  const document = continuationFixture(), previous = copy(document), store = input.useDocumentInputStore(createPinia())
  const target = model.textTargets(document).find(candidate => candidate.cellBlockId === 'continued')
  const location = { kind: 'body', selection: model.collapsedSelection({ ...target, offset: 2, pageAffinity: 1 }) }
  store.format(document.id, location, { bold: true }, 10)
  const source = store.get(document.id)
  execute(document, view(document, rect(0, 0)), { id: 'doc.command.table-insert-row', row: 0, position: 'before' })
  store.invalidate(document.id); store.revalidateModel(document.id, document, previous)
  const next = store.get(document.id).location
  assert.equal(next.selection.focus.cellBlockId, 'continued')
  assert.equal(next.selection.focus.row, 1)
  assert.deepEqual(store.getMarks(document.id, 11), {})
  assert.equal(store.completeUnchangedSelection(document.id, source, 11, next), false)
  assert.equal(store.completeUnchangedSelection(document.id, { ...source, generation: source.generation + 1 }, 11, next, true), false)
  assert.equal(store.completeUnchangedSelection(document.id, source, 11, next, true), true)
  assert.equal(store.getMarks(document.id, 11, next).bold, true)
})

test('deleting the selected table removes stale structural context and leaves a legal body fallback', () => {
  const document = fixture(1, 1), previous = copy(document), store = input.useDocumentInputStore(createPinia()), selected = view(document, rect(0, 0), 'table')
  store.select(document.id, { kind: 'body', selection: selected }, 1)
  execute(document, selected, { id: 'doc.command.table-delete' })
  store.invalidate(document.id); store.revalidateModel(document.id, document, previous)
  const next = store.get(document.id)
  assert.equal(next.location.selection.tableSelection, undefined)
  assert.equal(model.textTargets(document).some(target => model.sameTarget(target, next.location.selection.focus)), true)
  assert.equal(next.valid, false)
})

test('Facade handles follow current snapshots through one content transaction, undo, redo and JSON reopening', () => {
  const opened = open(continuationFixture()), document = opened.document, table = opened.table, cell = table.getCell(0, 0)
  const original = copy(document.getModel()), originalId = cell.getId()
  assert.equal(table.insertRowBefore(0), true)
  assert.equal(cell.getId(), originalId); assert.equal(cell.getRowIndex(), 1)
  assert.equal(opened.studio.editor.history.units[document.id].undo.length, 1)
  document.undo()
  assert.deepEqual(document.getModel(), original); assert.equal(cell.getRowIndex(), 0)
  document.redo()
  assert.equal(cell.getRowIndex(), 1)
  const saved = document.save(), reopened = open(saved)
  assert.deepEqual(reopened.document.getModel().blocks, saved.blocks)
  assert.equal(reopened.table.getCell(1, 0).getText().includes('Continuation'), true)
  const detached = table.getSource(); detached.rows[1].cells[0].blocks[0].runs[0].text = 'Outside'
  assert.equal(cell.getText().includes('Outside'), false)
})
