import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

let server
let clipboard
let models

beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  clipboard = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-clipboard.ts')
  models = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
})

afterAll(async () => { await server?.close() })

function paragraph(id, text, patch = {}) {
  const block = models.createParagraph(text)
  block.id = id
  block.runs = [models.createRun(text, patch)]
  return block
}

function tableBlock() {
  const table = paragraph('table-source', '')
  table.type = 'table'
  table.runs = []
  table.metadata.table = {
    rows: [{ id: 'row-source', cells: [
      { id: 'cell-one', blocks: [paragraph('cell-one-paragraph', 'Bold', { bold: true, color: '#aa0000', verticalAlign: 'superscript' })], style: { fill: '#eeeeee' }, rowSpan: 1, colSpan: 1 },
      { id: 'cell-two', blocks: [paragraph('cell-two-paragraph', 'Plain')], style: { fill: '#ffffff' }, rowSpan: 1, colSpan: 1 },
    ] }],
    columnWidths: [80, 120],
    headerRowCount: 1,
  }
  return table
}

const destinationStyle = {
  paragraphStyle: { align: 'center', indent: 12, lineHeight: 1.8, spaceBefore: 4, spaceAfter: 9 },
  textStyle: { bold: false, italic: true, underline: false, strike: true, color: '#003399', fontFamily: 'Aptos', fontSize: 14, verticalAlign: 'subscript' },
}

test('destination mode returns an independent paste model without mutating the clipboard blocks', () => {
  const source = [paragraph('source-paragraph', 'Keep me', { bold: true }), tableBlock()]
  const before = structuredClone(source)
  const pasted = clipboard.applyDocumentPasteMode(source, 'destination', destinationStyle)

  assert.deepEqual(source, before)
  assert.notEqual(pasted, source)
  assert.notEqual(pasted[1].metadata.table, source[1].metadata.table)
})

test('destination mode preserves table and paragraph structure while applying target style and source emphasis', () => {
  const source = [paragraph('source-heading', 'Chapter', { bold: true, underline: true, color: '#aa0000' }), tableBlock()]
  source[0].type = 'heading'
  source[0].level = 2
  source[0].metadata.namedStyleType = 'heading2'
  const before = structuredClone(source)

  const pasted = clipboard.applyDocumentPasteMode(source, 'destination', destinationStyle)
  const headingRun = pasted[0].runs[0]
  const table = pasted[1].metadata.table
  const cellRun = table.rows[0].cells[0].blocks[0].runs[0]

  assert.deepEqual(pasted.map(block => block.type), ['paragraph', 'table'])
  assert.equal(pasted[0].level, 0)
  assert.equal(pasted[0].metadata.namedStyleType, undefined)
  assert.deepEqual(pasted[0].style, destinationStyle.paragraphStyle)
  assert.deepEqual({ bold: headingRun.bold, italic: headingRun.italic, underline: headingRun.underline, strike: headingRun.strike }, { bold: true, italic: false, underline: true, strike: false })
  assert.equal(headingRun.color, '#003399')
  assert.equal(headingRun.fontFamily, 'Aptos')
  assert.equal(headingRun.verticalAlign, 'subscript')
  assert.equal(table.rows[0].cells[0].blocks[0].style.align, 'center')
  assert.deepEqual({ bold: cellRun.bold, color: cellRun.color, verticalAlign: cellRun.verticalAlign }, { bold: true, color: '#003399', verticalAlign: 'superscript' })
  assert.deepEqual(source, before)
})

test('plain-text mode flattens table cells to tabs and rows to paragraphs', () => {
  const source = [paragraph('before', 'Before'), tableBlock(), paragraph('after', 'After')]
  const flattened = clipboard.applyDocumentPasteMode(source, 'text', undefined, 'Before\nBold\tPlain\nAfter')

  assert.deepEqual(flattened.map(block => models.runText(block.runs)), ['Before', 'Bold\tPlain', 'After'])
  assert.ok(flattened.every(block => block.type === 'paragraph'))
})

test('destination hyperlink and field metadata are not character formatting', () => {
  const style = { ...destinationStyle, textStyle: { ...destinationStyle.textStyle, href: 'https://example.org/target', metadata: { field: 'target' } } }
  const source = [paragraph('source', 'label', { href: 'https://example.org/source' })]
  const [matched] = clipboard.applyDocumentPasteMode(source, 'destination', style)
  const [plain] = clipboard.applyDocumentPasteMode(source, 'text', style)
  assert.equal(matched.runs[0].href, source[0].runs[0].href)
  assert.equal(matched.runs[0].metadata, undefined)
  assert.equal(plain.runs[0].href, undefined)
  assert.equal(plain.runs[0].metadata, undefined)
})

test('text paste applies target marks and paragraph spacing while retaining Unicode and empty paragraphs', () => {
  const command = clipboard.documentPasteSpecialCommand({ text: '中文🙂\r\n\r\nLast\r\n' }, 'text', destinationStyle)
  assert.equal(command.id, 'doc.command.paste-blocks')
  assert.deepEqual(command.blocks.map(block => models.runText(block.runs)), ['中文🙂', '', 'Last', ''])
  for (const block of command.blocks) {
    assert.deepEqual(block.style, destinationStyle.paragraphStyle)
    assert.equal(block.runs[0].fontFamily, 'Aptos')
    assert.equal(block.runs[0].fontSize, 14)
    assert.equal(block.runs[0].italic, true)
    assert.equal(block.runs[0].strike, true)
    assert.equal(block.runs[0].verticalAlign, 'subscript')
  }
})

test('empty clipboard data produces no command in any paste mode', () => {
  for (const mode of ['source', 'destination', 'text']) {
    assert.equal(clipboard.documentPasteSpecialCommand({ text: '' }, mode, destinationStyle), undefined)
  }
})

test('source paste remains available for text-only clipboard data', () => {
  assert.deepEqual(clipboard.documentPasteSpecialCommand({ text: 'Source text' }, 'source'), {
    id: 'doc.command.insert-text',
    text: 'Source text',
  })
  assert.equal(clipboard.documentPasteSpecialCommand({ text: '' }, 'source'), undefined)
})

test('selection clipboard writes Unicode text and HTML from the same partial range without exposing the trailing paragraph', () => {
  const text = paragraph('selected', 'A中文🙂Z', { bold: true, italic: true, href: 'https://example.test/link' })
  const trailing = paragraph('trailing', 'outside', { color: '#ff0044' })
  trailing.metadata.codeBlock = { source: 'outside', language: 'plain' }
  const document = { blocks: [text, trailing] }
  const before = structuredClone(document)
  const range = { anchor: { blockId: trailing.id, offset: 0 }, focus: { blockId: text.id, offset: 1 } }
  const data = clipboard.documentSelectionClipboardData(document, range)
  assert.equal(data.text, '中文🙂Z\n')
  assert.match(data.html, /href="https:\/\/example\.test\/link"/)
  assert.match(data.html, /<em><strong>中文🙂Z<\/strong><\/em>/)
  assert.doesNotMatch(data.html, /outside|#ff0044/)
  assert.deepEqual(document, before)
})

test('partial legacy code selection serializes only selected source', () => {
  const block = paragraph('code', 'before SELECTED after')
  block.metadata.codeBlock = { source: 'before SELECTED after', language: 'plain', lineNumbers: false }
  const data = clipboard.documentSelectionClipboardData({ blocks: [block] }, {
    anchor: { blockId: block.id, offset: 7 }, focus: { blockId: block.id, offset: 15 },
  })
  assert.equal(data.text, 'SELECTED')
  assert.match(data.html, /<code>SELECTED<\/code>/)
  assert.doesNotMatch(data.html, /before SELECTED|SELECTED after/)
  assert.equal(block.metadata.codeBlock.source, 'before SELECTED after')
})

test('copy permits visible read-only text and cut additionally requires the selection and host to be editable', () => {
  const block = paragraph('protected', 'author text')
  const range = { anchor: { blockId: block.id, offset: 0 }, focus: { blockId: block.id, offset: 6 } }
  const context = { documentId: 'doc', userId: 'viewer' }
  const denial = { id: 'denial', scope: 'paragraph', targetId: block.id, userIds: ['viewer'], effect: 'deny' }
  assert.deepEqual(clipboard.documentClipboardPermissions({ blocks: [block] }, range, context), { copy: true, cut: true })
  assert.deepEqual(clipboard.documentClipboardPermissions({ blocks: [block] }, range, context, [], false), { copy: true, cut: false })
  assert.deepEqual(clipboard.documentClipboardPermissions({ blocks: [block], permissionRules: [{ ...denial, action: 'edit' }] },
    range, context), { copy: true, cut: false })
  assert.deepEqual(clipboard.documentClipboardPermissions({ blocks: [block], permissionRules: [{ ...denial, action: 'view' }] },
    range, context), { copy: false, cut: false })
  assert.deepEqual(clipboard.documentClipboardPermissions({ blocks: [block] }, models.collapsedSelection(range.anchor), context), { copy: false, cut: false })
})

test('code paste can use plain clipboard data without needing browser HTML parsing', () => {
  assert.deepEqual(clipboard.documentPasteCommand({ text: '  code\n', html: '<p><b>other</b></p>' }, true),
    { id: 'doc.command.insert-text', text: '  code\n' })
})
