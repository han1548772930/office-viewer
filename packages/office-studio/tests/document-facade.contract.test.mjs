import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let facade
let documentModel
let apiModule

before(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  facade = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-facade.ts')
  documentModel = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

after(async () => {
  await server?.close()
})

function paragraph(id, text) {
  const block = documentModel.createBlankDocument().blocks[0]
  block.id = id
  block.runs[0].text = text
  return block
}

function modelWith(blocks) {
  return { ...documentModel.createBlankDocument(), id: 'doc-test', blocks }
}

test('flattens text in model order and maps table-cell paragraphs back to their block', () => {
  const tableBlock = paragraph('table-1', '')
  tableBlock.type = 'table'
  tableBlock.runs = []
  tableBlock.metadata.table = {
    rows: [{
      id: 'row-1',
      cells: [{
        id: 'cell-1',
        blocks: [paragraph('cell-paragraph', 'inside')],
        style: {},
        rowSpan: 1,
        colSpan: 1,
      }],
    }],
  }
  const model = modelWith([paragraph('before', 'Alpha'), tableBlock, paragraph('after', 'Omega')])

  const body = facade.flattenDocumentBody(model)
  assert.equal(body.text, 'Alpha\ninside\nOmega')
  assert.deepEqual(body.segments.map(({ blockId, cellBlockId, startOffset, endOffset }) => ({ blockId, cellBlockId, startOffset, endOffset })), [
    { blockId: 'before', cellBlockId: undefined, startOffset: 0, endOffset: 5 },
    { blockId: 'table-1', cellBlockId: 'cell-paragraph', startOffset: 6, endOffset: 12 },
    { blockId: 'after', cellBlockId: undefined, startOffset: 13, endOffset: 18 },
  ])
  assert.equal(facade.findDocumentParagraph(model, { blockId: 'table-1' })?.paragraphId, 'cell-paragraph')
  assert.equal(facade.findDocumentParagraph(model, 'Omega')?.blockId, 'after')
})

test('reads half-open ranges and rejects malformed offsets', () => {
  const model = modelWith([paragraph('one', 'Alpha'), paragraph('two', 'Omega')])
  assert.deepEqual(facade.readDocumentRange(model, 3, 8), { startOffset: 3, endOffset: 8, text: 'ha\nOm' })
  assert.throws(() => facade.readDocumentRange(model, -1, 0), RangeError)
  assert.throws(() => facade.readDocumentRange(model, 4, 3), RangeError)
  assert.throws(() => facade.readDocumentRange(model, 0, 12), RangeError)
  assert.throws(() => facade.readDocumentRange(model, 1.5, 2), RangeError)
})

test('replaces across paragraphs without mutating the source and normalizes inserted line endings', () => {
  const source = modelWith([paragraph('one', 'alpha'), paragraph('two', 'omega')])
  const before = structuredClone(source)

  const result = facade.replaceDocumentRange(source, 2, 9, 'X\r\nY')

  assert.equal(facade.documentBodyText(result), 'alX\nYga')
  assert.deepEqual(source, before)
  assert.equal(result.modified, true)
  assert.equal(result.wordCount, 6)
})

test('inserts into a document with no text targets and preserves non-text blocks', () => {
  const image = paragraph('image-1', '')
  image.type = 'image'
  image.runs = []
  image.metadata.image = { src: 'data:image/png;base64,AA==', alt: 'image', width: 1, height: 1 }
  const source = modelWith([image])

  const result = facade.insertDocumentText(source, 0, 'Caption\ncontinued')

  assert.equal(facade.documentBodyText(result), 'Caption\ncontinued')
  assert.deepEqual(result.blocks.map((block) => block.type), ['paragraph', 'paragraph', 'image'])
  assert.equal(result.blocks[0].runs[0].text, 'Caption')
  assert.equal(result.blocks[2].id, 'image-1')
  assert.equal(result.modified, true)
})

test('empty insertion and empty deletion are no-ops', () => {
  const source = modelWith([paragraph('one', 'Alpha')])
  const originalRunId = source.blocks[0].runs[0].id

  const inserted = facade.insertDocumentText(source, 2, '')
  const deleted = facade.deleteDocumentRange(source, 2, 2)

  assert.equal(inserted.modified, false)
  assert.equal(deleted.modified, false)
  assert.equal(inserted.blocks[0].runs[0].id, originalRunId)
  assert.equal(deleted.blocks[0].runs[0].id, originalRunId)
})

test('rejects ranges that span independent table cells without changing the source', () => {
  const tableBlock = paragraph('table-1', '')
  tableBlock.type = 'table'
  tableBlock.runs = []
  tableBlock.metadata.table = {
    rows: [{
      id: 'row-1',
      cells: ['left', 'right'].map((text, index) => ({
        id: `cell-${index}`,
        blocks: [paragraph(`cell-paragraph-${index}`, text)],
        style: {},
        rowSpan: 1,
        colSpan: 1,
      })),
    }],
  }
  const model = modelWith([tableBlock])
  const before = structuredClone(model)

  assert.throws(() => facade.replaceDocumentRange(model, 1, 6, ''), RangeError)
  assert.deepEqual(model, before)
})

test('rejects ranges that would remove an intervening non-text block', () => {
  const divider = paragraph('divider-1', '')
  divider.type = 'divider'
  divider.runs = []
  const model = modelWith([paragraph('before', 'alpha'), divider, paragraph('after', 'omega')])
  const before = structuredClone(model)

  assert.throws(() => facade.replaceDocumentRange(model, 2, 9, ''), RangeError)
  assert.deepEqual(model, before)
})

test('creates detached saved snapshots with sanitized portable names', () => {
  const source = modelWith([paragraph('one', 'Alpha')])
  source.modified = true

  const saved = facade.saveDocumentSnapshot(source, '  Report:/Q4.  ')

  assert.equal(saved.name, 'Report--Q4')
  assert.equal(saved.modified, false)
  assert.notEqual(saved, source)
  saved.blocks[0].runs[0].text = 'changed'
  assert.equal(source.blocks[0].runs[0].text, 'Alpha')
  assert.equal(facade.safeDocumentName('CON.txt'), 'document-CON.txt')
  assert.equal(facade.safeDocumentName('x'.repeat(179) + '.y').endsWith('.'), false)
})

test('reads and writes text styles through the public range facade with undo and redo', () => {
  const first = paragraph('first', 'ABCD')
  first.runs = [
    { ...first.runs[0], id: 'run-plain', text: 'AB' },
    { ...first.runs[0], id: 'run-bold', text: 'CD', bold: true, fontSize: 14 },
  ]
  const source = modelWith([first, paragraph('second', 'EF')])
  const studio = new apiModule.OfficeStudio({ document: source })
  const document = apiModule.FUniver.newAPI(studio).getActiveDocument()
  assert.ok(document)

  const range = document.getTextRange(1, 6)
  assert.equal(range.getText(), 'BCD\nE')
  assert.deepEqual(range.getTextStyleRuns().map(({ startOffset, endOffset, textStyle }) => ({
    startOffset,
    endOffset,
    textStyle: { bold: textStyle.bold, fontSize: textStyle.fontSize },
  })), [
    { startOffset: 1, endOffset: 2, textStyle: { bold: false, fontSize: 11 } },
    { startOffset: 2, endOffset: 4, textStyle: { bold: true, fontSize: 14 } },
    { startOffset: 5, endOffset: 6, textStyle: { bold: false, fontSize: 11 } },
  ])
  assert.deepEqual(range.getTextStyle(), {})
  assert.throws(() => range.setTextStyle({ fontSize: 0 }), RangeError)

  assert.equal(range.setTextStyle({ color: '#cc0033', italic: true }), true)
  const styled = document.getModel()
  const runAt = (blockId, text) => styled.blocks.find((block) => block.id === blockId).runs.find((run) => run.text.includes(text))
  assert.equal(runAt('first', 'A').color, '#1f2937')
  assert.equal(runAt('first', 'B').color, '#cc0033')
  assert.equal(runAt('first', 'CD').fontSize, 14)
  assert.equal(runAt('second', 'E').italic, true)
  assert.equal(runAt('second', 'F').color, '#1f2937')
  assert.deepEqual(range.getTextStyle(), {})

  const currentTextRange = document.getTextRange(0, 1)
  assert.equal(currentTextRange.setTextStyle({ color: undefined }), false)
  assert.equal(document.getModel().blocks[0].runs[0].color, '#1f2937')

  assert.equal(document.undo(), true)
  assert.equal(document.getModel().blocks[0].runs.find((run) => run.text.includes('B')).color, '#1f2937')
  assert.equal(document.redo(), true)
  assert.equal(document.getModel().blocks[1].runs.find((run) => run.text.includes('E')).color, '#cc0033')

  const paragraphFacade = document.findParagraphs({ paragraphId: 'second' })[0]
  assert.ok(paragraphFacade)
  assert.equal(paragraphFacade.getParagraphStyle().align, 'left')
  assert.equal(paragraphFacade.setParagraphStyle({ align: 'center', indent: 2, lineHeight: 1.75 }), true)
  assert.deepEqual(paragraphFacade.getParagraphStyle(), {
    align: 'center', indent: 2, lineHeight: 1.75, spaceBefore: 0, spaceAfter: 8,
  })
  assert.equal(document.undo(), true)
  assert.equal(paragraphFacade.getParagraphStyle().align, 'left')
  assert.equal(document.redo(), true)
  assert.equal(paragraphFacade.getParagraphStyle().align, 'center')

  studio.dispose()
})

test('paragraph style facade updates paragraphs inside table cells', () => {
  const table = paragraph('table-1', '')
  table.type = 'table'
  table.runs = []
  table.metadata.table = {
    rows: [{
      id: 'row-1',
      cells: [{
        id: 'cell-1',
        blocks: [paragraph('cell-paragraph', 'inside')],
        style: {},
        rowSpan: 1,
        colSpan: 1,
      }],
    }],
  }
  const studio = new apiModule.OfficeStudio({ document: modelWith([table]) })
  const document = apiModule.FUniver.newAPI(studio).getActiveDocument()
  assert.ok(document)

  const paragraphFacade = document.findParagraphs({ paragraphId: 'cell-paragraph' })[0]
  assert.ok(paragraphFacade)
  assert.equal(paragraphFacade.setParagraphStyle({ align: 'right' }), true)
  assert.equal(document.getModel().blocks[0].metadata.table.rows[0].cells[0].blocks[0].style.align, 'right')
  assert.equal(paragraphFacade.setParagraphStyle({ align: undefined }), false)
  assert.equal(document.getModel().blocks[0].metadata.table.rows[0].cells[0].blocks[0].style.align, 'right')
  assert.equal(paragraphFacade.setTextStyle({ bold: true }), true)
  assert.equal(document.getModel().blocks[0].metadata.table.rows[0].cells[0].blocks[0].runs[0].bold, true)

  const tableModel = document.getModel().blocks[0].metadata.table
  const movedParagraph = tableModel.rows[0].cells[0].blocks[0]
  tableModel.rows[0].cells = []
  document.getModel().blocks.push(movedParagraph)
  assert.equal(paragraphFacade.getInfo().row, undefined)
  assert.equal(paragraphFacade.getInfo().blockId, 'cell-paragraph')
  assert.equal(paragraphFacade.setParagraphStyle({ align: 'center' }), true)
  assert.equal(document.getModel().blocks[1].style.align, 'center')

  studio.dispose()
})
