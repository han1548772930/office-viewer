import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred contracts for actual layout, view projection and checked Canvas command targets. Not executed in this batch.
let server, domain, modelUtils, layout, structuralLayout, scroll, canvas, input, ribbon, languages, tableMeasurement, tokenizer, codeFonts, scene
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-structured-canvas', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;[domain, modelUtils, layout, structuralLayout, scroll, canvas, input, ribbon, languages, tableMeasurement, tokenizer, codeFonts, scene] = await Promise.all([
    'domain/model', 'features/document/model/document-model', 'features/document/layout/document-layout',
    'features/document/layout/document-structural-layout', 'features/document/layout/document-code-scroll',
    'features/document/controllers/document-structured-canvas', 'features/document/model/document-input-session',
    'components/document/document-structure-ribbon', 'features/document/code-blocks/document-code-languages',
    'features/document/layout/document-table-measurement',
    'features/document/code-blocks/code-blocks',
    'features/document/code-blocks/document-code-fonts', 'features/document/rendering/document-scene',
  ].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterAll(async () => { await server?.close() })
const copy = value => structuredClone(value)
const measurer = { measure: text => Array.from(text).length * 8 }
const close = (actual, expected, tolerance = 0.03) => assert.ok(Math.abs(actual - expected) < tolerance, `${actual} != ${expected}`)
function paragraph(id, text = '') {
  const block = { ...modelUtils.createParagraph(text, { fontSize: 12 }), id }
  Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, lineHeight: 20, spacingRule: 'exact', wordWrap: true })
  return block
}
const documentWith = blocks => ({ ...domain.createBlankDocument(), id: 'structured-canvas', blocks,
  page: { width: 300, height: 160, marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20 } })
const code = (id, children, wrap = false) => ({ ...paragraph(id), type: 'code', runs: [], children,
  code: { language: 'typescript', showLineNumbers: true, tabSize: 4, wrap } })
const columns = (id, extra = {}) => ({ ...paragraph(id), type: 'columns', runs: [], columns: { version: 1, gap: 16, layout: 'fixed', responsive: 'shrink',
  columns: ['left', 'middle', 'right'].map((name, index) => ({ id: id + '-' + name, widthRatio: index + 1, blocks: [paragraph(id + '-' + name + '-p')] })), ...extra } })
function columnContainer(group, width = 360) {
  return { blockId: 'columns', kind: 'columns', x: 10, y: 20, width, height: 72, page: 0, first: true, last: true,
    columnGroup: { columnIds: group.columns.map(column => column.id), ...structuralLayout.documentColumnWidths(group, width), ...structuralLayout.documentColumnMetrics(group) } }
}
const allItems = items => items.flatMap(item => [item, ...(item.children ? allItems(item.children) : [])])
const itemById = (items, id) => allItems(items).find(item => item.id === id)

test('column defaults and explicit dimensions affect real body origins and reserve subsequent physical pages', () => {
  const block = columns('columns'), model = documentWith([block, paragraph('tail', 'Tail')]), source = copy(model)
  const result = layout.layoutDocument(model, measurer), container = result.containers.find(value => value.blockId === block.id)
  assert.deepEqual(structuralLayout.documentColumnMetrics(block.columns), { horizontalPadding: 8, minHeight: 72 })
  assert.equal(container.columnGroup.horizontalPadding, 8); assert.equal(container.columnGroup.minHeight, 72)
  const first = result.lines.find(line => line.target.blockId === 'columns-left-p')
  close(first.x, container.x + 8); assert.ok(container.height >= 72); assert.deepEqual(model, source)
  Object.assign(block.columns, { horizontalPadding: 12, minHeight: 320 })
  const largeSource = copy(model), large = layout.layoutDocument(model, measurer), frames = large.containers.filter(value => value.blockId === block.id)
  assert.ok(frames.length >= 3); assert.ok(new Set(frames.map(frame => frame.page)).size >= 3)
  assert.ok(frames.reduce((sum, frame) => sum + frame.height, 0) >= 320)
  assert.equal(frames[0].first, true); assert.equal(frames.at(-1).last, true)
  const tail = large.lines.find(line => line.target.blockId === 'tail'), last = frames.at(-1)
  assert.ok(tail.page > last.page || tail.page === last.page && tail.y >= last.y + last.height - 1e-6)
  assert.deepEqual(model, largeSource)
})

test('table-cell columns use the same padding/minimums and retain geometry through real pagination', () => {
  const group = columns('cell-columns', { horizontalPadding: 11, minHeight: 320 }), table = modelUtils.createTable(1, 1)
  table.style = { layout: 'fixed', width: 220, cellMargin: { top: 4, right: 4, bottom: 4, left: 4 } }; table.columnWidths = [220]
  table.rows[0].cells[0].blocks = [group]
  const block = { ...paragraph('table'), type: 'table', runs: [], metadata: { table } }
  const measured = tableMeasurement.measureDocumentTable(block, 260, new Map(), measurer), decoration = measured.containers.find(value => value.blockId === group.id)
  assert.equal(decoration.columnGroup.horizontalPadding, 11); assert.equal(decoration.columnGroup.minHeight, 320)
  close(measured.lines.find(line => line.target.cellBlockId === 'cell-columns-left-p').x, decoration.x + 11)
  assert.ok(decoration.height >= 320); assert.ok(measured.height >= 328)
  const model = documentWith([block, paragraph('tail', 'After table')]), original = copy(model), result = layout.layoutDocument(model, measurer)
  assert.ok(result.tableFragments.filter(fragment => fragment.tableId === 'table').length >= 3)
  assert.ok(new Set(result.containers.filter(container => container.blockId === group.id).map(container => container.page)).size >= 3)
  assert.deepEqual(model, original)
})

test('code projection preserves the source, shares lines across indexes, and aligns moved glyphs with hit testing', () => {
  const model = documentWith([code('code', [paragraph('long', 'a😀b'.repeat(120)), paragraph('short', 'x'), paragraph('empty')])])
  const source = layout.layoutDocument(model, measurer), before = copy(source), regions = scroll.documentCodeScrollRegions(source)
  const region = regions.find(region => region.container.blockId === 'code'), offset = Math.min(region.maxScroll, 160)
  assert.ok(offset > 0)
  const projected = scroll.projectDocumentCodeScroll(source, { [region.key]: offset }), raw = source.lines.find(line => line.target.blockId === 'long'), moved = projected.lines.find(line => line.target.blockId === 'long')
  assert.deepEqual(source, before); assert.notEqual(moved, raw); close(moved.x, raw.x - offset)
  moved.glyphs.forEach((glyph, index) => close(glyph.x, raw.glyphs[index].x - offset))
  assert.equal(moved.codeBlock.gutterX, raw.x); assert.equal(moved.codeBlock.gutterClip.x, region.container.x)
  assert.ok(moved.codeBlock.gutterClip.width >= moved.clip.width)
  assert.equal(projected.pageContent[moved.page].lines.find(line => line.target.blockId === 'long'), moved)
  assert.equal(layout.lineForPoint(projected, { blockId: 'long', offset: 0 }), moved)
  const visible = moved.glyphs.find(glyph => glyph.x >= region.clip.x && glyph.x + glyph.width <= region.clip.x + region.clip.width)
  assert.ok(visible)
  const hit = layout.hitTestDocument(projected, visible.x + visible.width / 4, visible.y + visible.height / 2)
  assert.equal(hit.blockId, 'long'); assert.equal(hit.offset, visible.offset)
  for (const id of ['short', 'empty']) {
    const line = projected.lines.find(line => line.target.blockId === id)
    assert.ok(line.x < region.clip.x)
    const point = layout.hitTestDocument(projected, region.clip.x + 4, line.y + line.height / 2)
    assert.equal(point.blockId, id); assert.equal(point.offset, id === 'short' ? 1 : 0)
  }
  const farther = scroll.projectDocumentCodeScroll(source, { [region.key]: offset + 80 })
  close(farther.lines.find(line => line.target.blockId === 'long').x, raw.x - offset - 80)
  assert.equal(scroll.documentCodeScrollRegions(source).find(value => value.key === region.key).maxScroll, region.maxScroll)
})

test('code caret scroll uses raw glyphs and clamping keeps every requested view offset within the block', () => {
  const text = 'abcdefgh'.repeat(100), source = layout.layoutDocument(documentWith([code('code', [paragraph('long', text)])]), measurer)
  const region = scroll.documentCodeScrollRegions(source)[0], trailing = scroll.documentCodeCaretScrollOffset(source, { blockId: 'long', offset: text.length }, {})
  assert.ok(trailing.offset > 0); assert.equal(trailing.key, region.key)
  const view = scroll.projectDocumentCodeScroll(source, { [trailing.key]: trailing.offset }), caret = layout.caretRect(view, { blockId: 'long', offset: text.length })
  assert.ok(caret.x >= region.clip.x); assert.ok(caret.x + caret.width <= region.clip.x + region.clip.width)
  const leading = scroll.documentCodeCaretScrollOffset(source, { blockId: 'long', offset: 0 }, { [region.key]: region.maxScroll })
  assert.equal(leading.offset, 0)
  assert.deepEqual(scroll.clampDocumentCodeScrollOffsets(source, { [region.key]: Infinity, missing: 99 }), {})
  assert.equal(scroll.clampDocumentCodeScrollOffsets(source, { [region.key]: -2 })[region.key], 0)
  assert.equal(scroll.clampDocumentCodeScrollOffsets(source, { [region.key]: region.maxScroll + 500 })[region.key], region.maxScroll)
  assert.ok(scroll.documentCodeScrollRegions(layout.layoutDocument(documentWith([code('wrapped', [paragraph('wrap', text)], true)]), measurer)).every(region => region.maxScroll === 0))
})

test('code keys isolate body/header/footer hosts while fragments of one block share their largest extent', () => {
  const body = layout.layoutDocument(documentWith([code('code', Array.from({ length: 10 }, (_, index) => paragraph('p-' + index, index === 9 ? 'x'.repeat(200) : 'x')))]), measurer)
  const regions = scroll.documentCodeScrollRegions(body)
  assert.ok(regions.length > 1); assert.equal(new Set(regions.map(region => region.key)).size, 1)
  assert.equal(new Set(regions.map(region => region.maxScroll)).size, 1); assert.ok(regions[0].maxScroll > 0)
  const container = regions[0].container, header = { kind: 'story', sectionId: 'first', story: 'header' }, footer = { ...header, story: 'footer' }
  assert.notEqual(scroll.documentCodeScrollKey(container), scroll.documentCodeScrollKey({ ...container, host: header }))
  assert.notEqual(scroll.documentCodeScrollKey({ ...container, host: header }), scroll.documentCodeScrollKey({ ...container, host: footer }))
  const headerLines = body.lines.map(line => ({ ...line, host: header })), headerContainers = body.containers.map(container => ({ ...container, host: header }))
  const combined = { ...body, auxiliaryContent: [{ x: 0, y: 0, width: body.width, height: body.height, page: 0, host: header, sourceHost: header,
    lines: headerLines, containers: headerContainers, cells: [], images: [], dividers: [], tableFragments: [] }] }
  const projected = scroll.projectDocumentCodeScroll(combined, { [regions[0].key]: 100 })
  assert.equal(projected.auxiliaryContent[0].lines[0].x, headerLines[0].x)
  close(projected.lines[0].x, body.lines[0].x - 100)
})

test('column boundary resizing preserves adjacent weights, untouched weights and minimum widths', () => {
  const group = columns('columns').columns, container = columnContainer(group), before = copy(group)
  const resized = canvas.resizeDocumentColumnBoundary(group, container, 0, 20)
  assert.ok(resized); close(resized.widths[0], container.columnGroup.widths[0] + 20)
  close(resized.widthRatios[0] + resized.widthRatios[1], 3, 1e-8); assert.equal(resized.widthRatios[2], 3)
  close(resized.guideX, container.x + resized.widths[0] + container.columnGroup.gap / 2); assert.deepEqual(group, before)
  group.columns[0].minWidth = 100; group.columns[1].minWidth = 50
  const constrained = canvas.resizeDocumentColumnBoundary(group, columnContainer(group), 0, 10_000)
  assert.ok(constrained); assert.ok(constrained.widths[0] >= 100); assert.ok(constrained.widths[1] >= 50)
  assert.equal(constrained.widthRatios[2], 3); close(constrained.widthRatios[0] + constrained.widthRatios[1], 3, 1e-8)
  assert.equal(canvas.resizeDocumentColumnBoundary(group, { ...container, columnGroup: { ...container.columnGroup, stacked: true } }, 0, 20), undefined)
  assert.equal(canvas.resizeDocumentColumnBoundary(group, { ...container, columnGroup: { ...container.columnGroup, columnIds: ['foreign'] } }, 0, 20), undefined)
  for (const delta of [NaN, Infinity]) assert.equal(canvas.resizeDocumentColumnBoundary(group, container, 0, delta), undefined)
  for (const index of [-1, 2, 0.5]) assert.equal(canvas.resizeDocumentColumnBoundary(group, container, index, 0), undefined)
})

test('Canvas targets clone the selection and reject old revisions, generations, hosts and input locations', () => {
  const source = documentWith([paragraph('p', 'abc')]), selection = { anchor: { blockId: 'p', offset: 1 }, focus: { blockId: 'p', offset: 1 } }
  const session = { location: input.documentInputLocation(selection), marks: {}, generation: 7, modelRevision: 3, valid: true }
  const target = canvas.captureDocumentStructuredCanvasTarget(source, selection, undefined, session), current = patch => canvas.documentStructuredCanvasTargetCurrent(target, { source, selection, session, requireSession: true, ...patch })
  assert.equal(current({}), true); assert.notEqual(target.selection, selection); assert.notEqual(target.inputSession, session)
  assert.equal(current({ source: copy(source) }), false); assert.equal(current({ disabled: true }), false)
  assert.equal(current({ session: { ...session, modelRevision: 4 } }), false); assert.equal(current({ session: { ...session, generation: 8 } }), false)
  assert.equal(current({ session: { ...session, valid: false } }), false); assert.equal(current({ session: undefined }), false)
  assert.equal(current({ host: { kind: 'story', sectionId: 'first', story: 'header' } }), false)
  const moved = { anchor: { blockId: 'p', offset: 2 }, focus: { blockId: 'p', offset: 2 } }
  assert.equal(current({ selection: moved }), false)
  delete source.id
  const standalone = canvas.captureDocumentStructuredCanvasTarget(source, selection), matches = patch => canvas.documentStructuredCanvasTargetCurrent(standalone, { source, selection, ...patch })
  assert.equal(matches({}), true); assert.equal(matches({ source: copy(source) }), false); assert.equal(matches({ requireSession: true }), false)
  source.blocks[0].runs[0].text = 'changed'; assert.equal(matches({}), false)
})

test('paragraph and floating menus parse actual insert, language and column actions from the shared catalogue', () => {
  const menus = ribbon.documentStructuredParagraphMenuItems({ paragraphId: 'p', disabled: false })
  assert.equal(menus.length, 2)
  for (const count of [2, 3, 4, 5]) {
    const current = itemById(menus, 'doc-context-insert-columns-current-' + count), below = itemById(menus, 'doc-context-insert-columns-below-' + count)
    assert.deepEqual(ribbon.documentStructureCommandFromMenu(current.commandId, current.params), { id: 'doc.command.wrap-column-group', columnCount: count })
    assert.deepEqual(ribbon.documentStructureCommandFromMenu(below.commandId, below.params), { id: 'doc.command.insert-column-group', columnCount: count, afterBlockId: 'p' })
  }
  for (const type of ['quote', 'callout', 'code']) {
    const below = itemById(menus, 'doc-context-insert-' + type + '-below')
    assert.deepEqual(ribbon.documentStructureCommandFromMenu(below.commandId, below.params), { id: 'doc.command.insert-structured-block', blockType: type, afterBlockId: 'p' })
  }
  const options = languages.documentStructuredCodeLanguageOptions('rust')
  assert.equal(options.length, 21); assert.equal(languages.documentStructuredCodeLanguageOptions('legacy-language').at(-1).value, 'legacy-language')
  const codeItems = ribbon.documentStructureContextItems({ blockId: 'code', type: 'code', disabled: false,
    code: { language: 'rust', wrap: false, showLineNumbers: true, tabSize: 4 } })
  assert.equal(itemById(codeItems, 'doc-context-code-language').children.length, 21)
  const language = itemById(codeItems, 'doc-context-code-language-typescript')
  assert.deepEqual(ribbon.documentStructureCommandFromMenu(language.commandId, language.params), { id: 'doc.command.update-code-settings', blockId: 'code', patch: { language: 'typescript' } })
  const group = columns('columns').columns, items = ribbon.documentStructureContextItems({ blockId: 'columns', type: 'columns', disabled: false, columns: group, columnId: group.columns[1].id })
  const equal = itemById(items, 'doc-context-column-equal'), remove = itemById(items, 'doc-context-column-delete')
  assert.deepEqual(ribbon.documentStructureCommandFromMenu(equal.commandId, equal.params), { id: 'doc.command.update-column-group', blockId: 'columns', widthRatios: [1, 1, 1] })
  assert.equal(ribbon.documentStructureCommandFromMenu(remove.commandId, remove.params).keepContents, false)
  const two = { ...group, columns: group.columns.slice(0, 2) }
  assert.equal(itemById(ribbon.documentStructureContextItems({ blockId: 'columns', type: 'columns', disabled: false, columns: two, columnId: two.columns[0].id }), 'doc-context-column-delete').disabled, true)
  for (const params of [{ columnCount: 1 }, { columnCount: 6 }, { columnCount: '2' }]) assert.equal(ribbon.documentStructureCommandFromMenu('doc.command.wrap-column-group', params), undefined)
  assert.equal(ribbon.documentStructureCommandFromMenu('doc.command.update-column-group', { blockId: 'columns', widthRatios: [1, -1] }), undefined)
})

test('public plaintext and markup aliases use the existing plain and HTML grammar without rewriting stored language values', () => {
  const plain = 'const a = 1', markup = '<div class="item">Text</div>'
  assert.deepEqual(tokenizer.tokenizeCode(plain, 'plaintext'), tokenizer.tokenizeCode(plain, 'plain'))
  assert.deepEqual(tokenizer.tokenizeCode(markup, 'markup'), tokenizer.tokenizeCode(markup, 'html'))
  assert.equal(languages.documentStructuredCodeLanguageOptions('markup').find(option => option.value === 'markup').value, 'markup')
})

test('Code font mutation changes actual glyph measurement and Canvas scene fonts while preserving logical targets', () => {
  const model = documentWith([code('code', [paragraph('inside', 'abcd')])]), original = copy(model)
  const fontMeasurer = { measure: (text, font) => Array.from(text).length * (font.includes('Menlo') ? 10 : 8) }
  const before = layout.layoutDocument(model, fontMeasurer), lineBefore = before.lines.find(line => line.target.blockId === 'inside')
  assert.ok(lineBefore)
  codeFonts.setDocumentCodeFontFamily(model, model.blocks[0], 'Menlo')
  const result = layout.layoutDocument(model, fontMeasurer), line = result.lines.find(line => line.target.blockId === 'inside')
  assert.ok(line.glyphs.every(glyph => glyph.font.includes('Menlo')))
  close(line.glyphs[0].x, lineBefore.glyphs[0].x); assert.ok(line.glyphs.at(-1).x > lineBefore.glyphs.at(-1).x)
  assert.deepEqual(line.target, lineBefore.target)
  const hit = layout.hitTestDocument(result, line.glyphs[1].x + 2, line.y + line.height / 2)
  assert.equal(hit.blockId, 'inside'); assert.equal(hit.offset, line.glyphs[1].offset)
  const draw = scene.sceneFromDocument(model, result, { width: 600, height: 600, scrollLeft: 0, scrollTop: 0, zoom: 100 },
    { selection: { anchor: { blockId: 'inside', offset: 0 }, focus: { blockId: 'inside', offset: 0 } },
      focused: true, caretVisible: false, composing: false, images: new Map() })
  const glyphCommands = draw.commands.filter(command => command.kind === 'text' && ['a', 'b', 'c', 'd'].includes(command.text))
  assert.equal(glyphCommands.length, 4); assert.ok(glyphCommands.every(command => command.font.includes('Menlo')))
  assert.equal(modelUtils.runText(model.blocks[0].children[0].runs), modelUtils.runText(original.blocks[0].children[0].runs))
  assert.deepEqual(model.blocks[0].children[0].runs.map(run => run.id), original.blocks[0].children[0].runs.map(run => run.id))
})
