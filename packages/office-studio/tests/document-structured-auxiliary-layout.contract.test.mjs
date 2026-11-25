import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred integration contracts for shared table/auxiliary measurement. No runtime acceptance is implied.
let server, domain, modelUtils, layout, auxiliary, blockMeasurement, tableMeasurement, sections, storyBlocks, structuralBlocks, textFields
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-structured-auxiliary-layout', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;[domain, modelUtils, layout, auxiliary, blockMeasurement, tableMeasurement, sections, storyBlocks, structuralBlocks, textFields] = await Promise.all([
    'domain/model', 'features/document/model/document-model', 'features/document/layout/document-layout',
    'features/document/layout/document-auxiliary-layout', 'features/document/layout/document-block-measurement',
    'features/document/layout/document-table-measurement', 'features/document/services/section-header-footer',
    'features/document/model/document-story-blocks', 'features/document/model/document-structural-blocks',
    'features/document/model/document-text-fields',
  ].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterAll(async () => { await server?.close() })
const copy = value => structuredClone(value)
const measurer = { measure: text => Array.from(text).length * 8 }
function paragraph(id, text = '') {
  const block = { ...modelUtils.createParagraph(text, { fontSize: 12 }), id }
  Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, lineHeight: 20, spacingRule: 'exact', wordWrap: true })
  return block
}
const content = blocks => ({ blocks, runs: [], paragraphs: [], metadata: {} })
const code = (id, text = 'let x = 1') => ({ ...paragraph(id), type: 'code', runs: [], children: [paragraph(id + '-p', text)],
  code: { language: 'rust', showLineNumbers: true, tabSize: 4, wrap: false } })
const columns = id => ({ ...paragraph(id), type: 'columns', runs: [], columns: { version: 1, gap: 16, horizontalPadding: 9, minHeight: 96,
  layout: 'fixed', responsive: 'shrink', columns: ['left', 'right'].map(name => ({ id: id + '-' + name, widthRatio: 1,
    blocks: [paragraph(id + '-' + name + '-p', name)] })) } })
function tableBlock(id, blocks) {
  const table = modelUtils.createTable(1, 1)
  table.rows[0].cells[0].id = id + '-cell'; table.rows[0].cells[0].blocks = blocks
  table.style = { layout: 'fixed', width: 240, cellMargin: { top: 4, right: 4, bottom: 4, left: 4 } }; table.columnWidths = [240]
  return { ...paragraph(id), type: 'table', runs: [], metadata: { table } }
}
function documentWithStories(headerBlocks, footerBlocks = []) {
  const model = { ...domain.createBlankDocument(), id: 'structured-auxiliary', blocks: [paragraph('body', 'Body')],
    page: { width: 600, height: 800, marginLeft: 48, marginRight: 48, marginTop: 260, marginBottom: 220 } }
  const header = storyBlocks.projectDocumentStoryBlocks(content(headerBlocks)).content,
    footer = storyBlocks.projectDocumentStoryBlocks(content(footerBlocks)).content
  model.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: modelUtils.runText(header.runs), footer: modelUtils.runText(footer.runs),
    storyContent: { header, footer } } }
  return model
}

test('top-level header structures produce real child text, decoration geometry and host-aware input targets', () => {
  const blocks = [{ ...paragraph('quote'), type: 'quote', runs: [], children: [paragraph('quote-p', 'Quoted text')] },
    { ...paragraph('callout'), type: 'callout', runs: [], children: [paragraph('callout-p', 'Callout text')], callout: structuralBlocks.defaultDocumentCalloutConfig() },
    code('code', 'x'.repeat(100)), columns('columns')]
  const model = documentWithStories(blocks), original = copy(model), host = { kind: 'story', sectionId: 'first', story: 'header' },
    bounds = { x: 48, y: 16, width: 504, height: 260 }
  const measured = auxiliary.layoutDocumentAuxiliaryBlocks(model, host, bounds, 2, 3, 7, measurer)
  assert.deepEqual(new Set(measured.containers.map(container => container.blockId)), new Set(['quote', 'callout', 'code', 'columns']))
  for (const id of ['quote-p', 'callout-p', 'code-p', 'columns-left-p', 'columns-right-p']) {
    const lines = measured.lines.filter(line => line.target.blockId === id)
    assert.ok(lines.length); assert.ok(lines.every(line => line.page === 2 && line.host.story === 'header'))
    assert.ok(lines.every(line => line.target.row === undefined && line.target.cellBlockId === undefined))
  }
  assert.equal(measured.lines.filter(line => line.target.blockId === 'code-p').length, 1)
  const codeLine = measured.lines.find(line => line.target.blockId === 'code-p')
  assert.equal(codeLine.codeBlock.containerId, 'code'); assert.equal(codeLine.codeBlock.language, 'rust')
  assert.equal(measured.containers.find(container => container.blockId === 'code').codeWrap, false)
  const group = measured.containers.find(container => container.blockId === 'columns')
  assert.equal(group.columnGroup.horizontalPadding, 9); assert.equal(group.columnGroup.minHeight, 96); assert.ok(group.height >= 96)
  assert.equal(measured.lines.find(line => line.target.blockId === 'columns-left-p').x, group.x + 9)
  assert.ok(measured.containers.every(container => container.host.story === 'header' && container.page === 2 && container.clip.y >= bounds.y))
  assert.deepEqual(model, original)
})

test('auxiliary top-level tables retain structured containers, cell addresses and their intersected host clips', () => {
  const table = tableBlock('header-table', [code('cell-code', 'c'.repeat(100)), columns('cell-columns')]), model = documentWithStories([table]), original = copy(model)
  const host = { kind: 'story', sectionId: 'first', story: 'header' }, bounds = { x: 48, y: 16, width: 504, height: 220 }
  const output = auxiliary.layoutDocumentAuxiliaryBlocks(model, host, bounds, 0, 1, 1, measurer)
  assert.equal(output.tableFragments[0].tableId, 'header-table'); assert.equal(output.tableFragments[0].rows[0].y, output.tableFragments[0].y)
  const line = output.lines.find(line => line.target.cellBlockId === 'cell-code-p')
  assert.ok(line); assert.equal(line.target.blockId, 'header-table'); assert.equal(line.target.row, 0); assert.equal(line.target.col, 0)
  assert.equal(line.tableCellId, 'header-table-cell'); assert.equal(line.codeBlock.containerId, 'cell-code')
  for (const id of ['cell-code', 'cell-columns']) {
    const container = output.containers.find(container => container.blockId === id)
    assert.ok(container); assert.equal(container.host.story, 'header'); assert.equal(container.tableDepth, 1)
    assert.ok(container.clip.x >= bounds.x); assert.ok(container.clip.y >= bounds.y)
    assert.ok(container.clip.x + container.clip.width <= bounds.x + bounds.width)
    assert.ok(container.clip.y + container.clip.height <= bounds.y + bounds.height)
  }
  const group = output.containers.find(container => container.blockId === 'cell-columns')
  assert.equal(output.lines.find(line => line.target.cellBlockId === 'cell-columns-left-p').x, group.x + 9)
  assert.deepEqual(model, original)
})

test('footer structures align their measured height to the bottom without changing source content', () => {
  const model = documentWithStories([paragraph('header-p', 'Header')], [columns('footer-columns'), code('footer-code')]), original = copy(model)
  const host = { kind: 'story', sectionId: 'first', story: 'footer' }, bounds = { x: 48, y: 500, width: 504, height: 220 }
  const output = auxiliary.layoutDocumentAuxiliaryBlocks(model, host, bounds, 0, 1, 1, measurer, true)
  assert.equal(output.y + output.height, bounds.y + bounds.height)
  assert.ok(output.lines.every(line => line.host.story === 'footer')); assert.ok(output.containers.every(container => container.host.story === 'footer'))
  assert.deepEqual(model, original)
})

test('physical host layout keeps header and footer containers with the exact lines used for caret and hit testing', () => {
  const model = documentWithStories([code('header-code')], [columns('footer-columns')]), original = copy(model), result = layout.layoutDocument(model, measurer)
  const header = { kind: 'story', sectionId: 'first', story: 'header' }, footer = { ...header, story: 'footer' }
  const headerLayout = auxiliary.documentLayoutForBlockHost(result, header), footerLayout = auxiliary.documentLayoutForBlockHost(result, footer)
  assert.ok(headerLayout.containers.some(container => container.blockId === 'header-code'))
  assert.ok(footerLayout.containers.some(container => container.blockId === 'footer-columns'))
  assert.ok(headerLayout.containers.every(container => container.host.story === 'header')); assert.ok(footerLayout.containers.every(container => container.host.story === 'footer'))
  assert.ok(headerLayout.lines.every(line => line.target.blockId !== 'body' && line.host.story === 'header'))
  const line = headerLayout.lines.find(line => line.target.blockId === 'header-code-p'), container = headerLayout.containers.find(container => container.blockId === 'header-code')
  assert.equal(headerLayout.pageContent[line.page].lines.find(candidate => candidate === line), line)
  assert.equal(headerLayout.pageContent[container.page].containers.find(candidate => candidate === container), container)
  const glyph = line.glyphs[0], point = layout.hitTestDocument(headerLayout, glyph.x + glyph.width / 4, line.y + line.height / 2)
  assert.equal(point.blockId, 'header-code-p'); assert.equal(point.offset, 0)
  assert.ok(layout.caretRect(headerLayout, point)); assert.deepEqual(model, original)
})

test('PAGE display inside a structured header keeps cached source offsets and leaves the document unchanged', () => {
  const child = paragraph('page-p', '99'), quote = { ...paragraph('page-quote'), type: 'quote', runs: [], children: [child] }
  textFields.setDocumentTextFields(child, [{ id: 'page', field: 'PAGE', cachedResult: '99', startOffset: 0, endOffset: 2 }])
  const model = documentWithStories([quote]), original = copy(model), host = { kind: 'story', sectionId: 'first', story: 'header' }
  const output = auxiliary.layoutDocumentAuxiliaryBlocks(model, host, { x: 48, y: 16, width: 504, height: 220 }, 0, 3, 7, measurer)
  const line = output.lines.find(line => line.target.blockId === 'page-p')
  assert.equal(line.glyphs.map(glyph => glyph.text).join(''), '3'); assert.equal(line.start, 0); assert.equal(line.end, 2)
  assert.equal(line.glyphs[0].offset, 0); assert.equal(line.glyphs[0].end, 2)
  assert.deepEqual(model, original)
})

test('shared line translation preserves code gutter, clips and marker coordinates without touching the input', () => {
  const measured = blockMeasurement.measureDocumentBlockCollections([[code('code')]], { width: 300, markers: new Map(), measurer, measureTable: tableMeasurement.measureDocumentTable })
  const line = measured.lines[0]
  line.codeBlock.gutterX = line.x; line.codeBlock.gutterClip = { x: 0, y: 0, width: 300, height: 52 }; line.clip = { x: 44, y: 0, width: 244, height: 52 }
  line.markerStyle = { x: 2, font: '12px Arial', color: '#112233' }
  const before = copy(line), shifted = blockMeasurement.translateDocumentLine(line, 20, 30)
  assert.deepEqual(line, before); assert.equal(shifted.codeBlock.gutterX, line.codeBlock.gutterX + 20)
  assert.deepEqual(shifted.codeBlock.gutterClip, { x: 20, y: 30, width: 300, height: 52 })
  assert.equal(shifted.clip.x, line.clip.x + 20); assert.equal(shifted.clip.y, line.clip.y + 30)
  assert.equal(shifted.markerStyle.x, line.markerStyle.x + 20); assert.equal(shifted.glyphs[0].x, line.glyphs[0].x + 20)
})
