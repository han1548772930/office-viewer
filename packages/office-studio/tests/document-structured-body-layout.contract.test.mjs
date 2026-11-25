import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred body/layout contracts. No browser or performance acceptance is implied.
let server, domain, modelUtils, codec, bodyTypes, publicBody, columns, structuralLayout, textLayout, layout, structure, appearance, panelState
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-structured-body-layout', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;[domain, modelUtils, codec, bodyTypes, publicBody, columns, structuralLayout, textLayout, layout, structure, appearance, panelState] = await Promise.all([
    'domain/model', 'features/document/model/document-model', 'api/document-body-codec', 'api/document-body-types', 'api/document-public-body',
    'api/document-columns', 'features/document/layout/document-structural-layout', 'features/document/layout/document-text-layout',
    'features/document/layout/document-layout', 'features/document/commands/document-structure-commands',
    'features/document/model/document-structured-appearance', 'components/document/document-structure-panel-state',
  ].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterAll(async () => { await server?.close() })
const copy = value => structuredClone(value)
const paragraph = (id, text = '') => ({ ...modelUtils.createParagraph(text), id })
const quote = (id, children) => ({ ...paragraph(id), type: 'quote', runs: [], children })
const documentWith = blocks => ({ ...domain.createBlankDocument(), id: 'structured-body', blocks })
const content = blocks => ({ blocks, runs: [], paragraphs: [], metadata: {} })
const measurer = { measure: text => text.length }

test('quote color and real paragraphs survive body encoding without mutating the caller', () => {
  const child = paragraph('inside', 'A😀B'), block = quote('quote', [child])
  block.metadata.quoteStyle = { version: 1, lineColor: '#123456' }; child.runs[0].color = '#aabbcc'
  const input = content([block]), before = copy(input), encoded = codec.encodeDocumentBody(input)
  assert.deepEqual(input, before)
  assert.equal(encoded.body.dataStream, bodyTypes.DOCUMENT_BODY_TOKENS.blockStart + 'A😀B\r' + bodyTypes.DOCUMENT_BODY_TOKENS.blockEnd + '\n')
  assert.equal(encoded.body.paragraphs[0].paragraphStyle.borderLeft.color.rgb, '#123456')
  assert.equal(encoded.body.textRuns.find(run => run.st === 1).ts.cl.rgb, '#aabbcc')
  const decoded = codec.decodeDocumentBody(encoded.body, encoded).content
  assert.equal(decoded.blocks[0].id, 'quote'); assert.equal(decoded.blocks[0].children[0].id, 'inside')
  assert.equal(appearance.documentQuoteLineColor(decoded.blocks[0]), '#123456')
  assert.equal(modelUtils.runText(decoded.blocks[0].children[0].runs), 'A😀B')
})

test('unwrap after body reopen restores original author border and removes a generated-only quote border', () => {
  for (const border of [undefined, { color: { rgb: '#0088aa' }, width: 7, dashStyle: 3, padding: 5 }]) {
    const child = paragraph('inside', 'text')
    child.metadata.univerBodyParagraphSource = { version: 1, paragraph: { paragraphStyle: { ...(border ? { borderLeft: copy(border) } : {}) } } }
    const block = quote('quote', [child]); block.metadata.quoteStyle = { version: 1, lineColor: '#bb0000' }
    let encoded = codec.encodeDocumentBody(content([block]))
    // Repeated save/reopen retains the original pre-quote rule, rather than making the generated rule the new original.
    const first = codec.decodeDocumentBody(encoded.body, encoded).content
    encoded = codec.encodeDocumentBody(first)
    const second = codec.decodeDocumentBody(encoded.body, encoded).content, model = documentWith(second.blocks)
    structure.executeDocumentStructureCommand(model, { anchor: { blockId: 'inside', offset: 0 }, focus: { blockId: 'inside', offset: 0 } },
      { id: 'doc.command.unwrap-structured-block', blockId: 'quote' }, { documentId: model.id, userId: 'editor' })
    const restored = codec.encodeDocumentBody({ ...second, blocks: model.blocks })
    assert.deepEqual(restored.body.paragraphs[0].paragraphStyle.borderLeft, border)
    assert.equal(model.blocks[0].metadata.documentQuoteBorderSource, undefined)
  }
})

test('unverified imported borders and independently edited borders survive quote removal', () => {
  const authorBorder = { color: { rgb: '#112233' }, width: 3 }
  for (const provenance of [undefined, { version: 1, quoteId: 'quote', generated: { color: { rgb: '#bb0000' }, width: 2 } }]) {
    const child = paragraph('inside', 'text')
    child.metadata.univerBodyParagraphSource = { version: 1, paragraph: { paragraphStyle: { borderLeft: copy(authorBorder) } } }
    if (provenance) child.metadata.documentQuoteBorderSource = copy(provenance)
    appearance.clearDocumentQuoteParagraphBorders(quote('quote', [child]))
    assert.deepEqual(child.metadata.univerBodyParagraphSource.paragraph.paragraphStyle.borderLeft, authorBorder)
  }
})

test('column ranges track nested groups and nonprinting SDT markers without swallowing neighboring columns', () => {
  const T = bodyTypes.DOCUMENT_BODY_TOKENS
  const inner = T.groupStart + T.columnStart + 'x\r' + T.columnEnd + T.columnStart + 'y\r' + T.columnEnd + T.groupEnd
  const first = T.columnStart + 'A\r' + inner + T.columnEnd, second = T.columnStart + 'B\r' + T.columnEnd
  const stream = T.groupStart + T.rangeStart + first + T.rangeEnd + second + T.groupEnd
  const group = { columnGroupId: 'outer', startIndex: 0, endIndex: stream.length - 1 }
  assert.deepEqual(columns.documentColumnRanges(stream, group, ['left', 'right']), [
    { column: 0, columnId: 'left', startOffset: 2, endOffset: 2 + first.length - 1 },
    { column: 1, columnId: 'right', startOffset: 3 + first.length, endOffset: stream.length - 2 },
  ])
  assert.throws(() => columns.documentColumnRanges(stream, group, ['missing']), /栏数量/)
})

test('column shrink preserves gap and minima, while stack uses the full host width', () => {
  const group = (minimums, responsive = 'shrink', gap = 16) => ({ version: 1, gap, responsive, layout: 'fixed',
    columns: minimums.map((minWidth, index) => ({ id: String(index), minWidth, widthRatio: 1, blocks: [] })) })
  assert.deepEqual(structuralLayout.documentColumnWidths(group([60, 0]), 100), { widths: [60, 24], gap: 16, stacked: false })
  assert.deepEqual(structuralLayout.documentColumnWidths(group([60, 50]), 100), { widths: [60, 50], gap: 16, stacked: false })
  assert.deepEqual(structuralLayout.documentColumnWidths(group([30, 30], 'stack'), 75), { widths: [75, 75], gap: 16, stacked: true })
  assert.deepEqual(structuralLayout.documentColumnWidths(group([0, 0]), 10), { widths: [0, 0], gap: 16, stacked: false })
})

test('column dimensions retain authored NumberUnit values until edited, then write changed lengths in CSS pixels', () => {
  const block = { ...paragraph('columns'), type: 'columns', runs: [], columns: { version: 1, gap: 16, layout: 'fixed', responsive: 'shrink',
    columns: ['left', 'right'].map(id => ({ id, widthRatio: 1, blocks: [paragraph(id + '-p', id)] })) } }
  const encoded = codec.encodeDocumentBody(content([block])), source = encoded.body.columnGroups[0]
  source.gap = 12; source.horizontalPadding = { v: 6, u: 0, author: 'keep' }; source.minHeight = 54
  const original = copy(encoded), decoded = codec.decodeDocumentBody(encoded.body, encoded).content
  assert.deepEqual(encoded, original)
  assert.deepEqual(structuralLayout.documentColumnMetrics(decoded.blocks[0].columns), { horizontalPadding: 8, minHeight: 72 })
  const unchanged = codec.encodeDocumentBody(decoded).body.columnGroups[0]
  assert.equal(unchanged.gap, 12); assert.deepEqual(unchanged.horizontalPadding, source.horizontalPadding); assert.equal(unchanged.minHeight, 54)
  Object.assign(decoded.blocks[0].columns, { gap: 20, horizontalPadding: 10, minHeight: 96 })
  const changed = codec.encodeDocumentBody(decoded).body.columnGroups[0]
  assert.deepEqual(changed.gap, { v: 20, u: 3 }); assert.deepEqual(changed.horizontalPadding, { v: 10, u: 3 }); assert.deepEqual(changed.minHeight, { v: 96, u: 3 })
  const reopened = codec.decodeDocumentBody(codec.encodeDocumentBody(decoded).body).content.blocks[0].columns
  assert.deepEqual(structuralLayout.documentColumnMetrics(reopened), { horizontalPadding: 10, minHeight: 96 })
})

test('provided malformed column dimensions are rejected without mutating the imported body', () => {
  const block = { ...paragraph('columns'), type: 'columns', runs: [], columns: { version: 1, gap: 16, layout: 'fixed', responsive: 'shrink',
    columns: ['left', 'right'].map(id => ({ id, widthRatio: 1, blocks: [paragraph(id + '-p')] })) } }
  for (const key of ['horizontalPadding', 'minHeight']) for (const value of [null, -1, '8', { v: 1, u: 99 }, { v: -1, u: 3 }, { v: 1_000_001, u: 3 }]) {
    const encoded = codec.encodeDocumentBody(content([block])); encoded.body.columnGroups[0][key] = value
    const before = copy(encoded)
    assert.throws(() => codec.decodeDocumentBody(encoded.body, encoded)); assert.deepEqual(encoded, before)
  }
})

test('nowrap preserves a long logical code line and tab uses measured tab stops with one UTF-16 source character', () => {
  const runs = [modelUtils.createRun('A\tB\n😀C')]
  const lines = textLayout.wrapRuns(runs, 3, measurer, false, 14.67, undefined, undefined, { softWrap: false, tabSize: 4 })
  assert.deepEqual(lines.map(line => [line.start, line.end, line.width]), [[0, 3, 5], [4, 7, 3]])
  const tab = lines[0].glyphs[1]
  assert.deepEqual([tab.text, tab.offset, tab.end, tab.width, tab.tab.leaderText], ['\t', 1, 2, 3, ''])
  assert.ok(textLayout.wrapRuns([modelUtils.createRun('abcdefgh')], 3, measurer).length > 1)
  assert.equal(textLayout.wrapRuns([modelUtils.createRun('abcdefgh')], 3, measurer, false, 14.67, undefined, undefined, { softWrap: false }).length, 1)
})

test('nested code lines retain the inner language and container identity during parent layout annotation', () => {
  const code = (id, language, children) => ({ ...paragraph(id), type: 'code', runs: [], children, code: { language, showLineNumbers: true, tabSize: 4, wrap: true } })
  const model = documentWith([code('outer', 'typescript', [paragraph('outer-p', 'const a = 1'), code('inner', 'rust', [paragraph('inner-p', 'let b = 2')])])])
  const result = layout.layoutDocument(model, measurer)
  assert.equal(result.lines.find(line => line.target.blockId === 'outer-p').codeBlock.containerId, 'outer')
  assert.deepEqual(result.lines.filter(line => line.target.blockId === 'inner-p').map(line => [line.codeBlock.containerId, line.codeBlock.language]), [['inner', 'rust']])
})

test('public offsets use source=stream and target=local text across structural markers and a one-token note label', () => {
  const child = paragraph('inside')
  child.runs = [modelUtils.createRun('A'), modelUtils.createRun('10', { metadata: { note: { type: 'footnote', noteId: 'note' } } }), modelUtils.createRun('😀Z')]
  const model = documentWith([quote('quote', [child]), paragraph('after', 'tail')]), index = publicBody.documentPublicBodyIndex(model)
  const position = index.positions.find(position => position.block.id === 'inside'), start = position.start
  assert.equal(index.data.body.dataStream.slice(start, position.end), 'A\uFFFC😀Z')
  assert.equal(publicBody.documentPublicOffsetAtFlat(index, 3), start + 2)
  assert.equal(publicBody.documentPublicOffsetAtFlat(index, 2, 'start'), start + 1)
  assert.equal(publicBody.documentPublicOffsetAtFlat(index, 2, 'end'), start + 2)
  assert.deepEqual(publicBody.documentPublicRangeToFlat(index, start + 2, start + 4), { startOffset: 3, endOffset: 5 })
  assert.deepEqual(publicBody.documentPublicRangeToFlat(index, start + 4, start + 5), { startOffset: 5, endOffset: 6 })
  const delimiter = publicBody.documentPublicRangeToFlat(index, 0, 1)
  assert.equal(delimiter.startOffset, delimiter.endOffset)
  assert.equal(publicBody.documentPublicOffsetAtFlat(index, 7), index.positions.find(position => position.block.id === 'after').start)
})

test('standalone property sessions accept missing document IDs and reject reopen, parent-source replacement, mutation and epoch changes', () => {
  const model = documentWith([paragraph('body', 'text')]); delete model.id
  const parentSource = copy(model), epoch = 8
  const target = { unitId: '', source: model, standaloneSnapshot: JSON.stringify(model), standaloneEpoch: epoch, standaloneParentSource: parentSource }
  const matches = patch => panelState.documentStandaloneStructureTargetCurrent(target, { model, epoch, parentSource, ...patch })
  assert.equal(matches({}), true)
  assert.equal(matches({ model: copy(model) }), false)
  assert.equal(matches({ parentSource: copy(parentSource) }), false)
  assert.equal(matches({ epoch: epoch + 1 }), false)
  model.blocks[0].runs[0].text = 'changed in place'
  assert.equal(matches({}), false)
})
