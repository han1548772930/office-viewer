import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, runText } from '../src/features/document/model/document-model.ts'
import { createLatexFormulaBlock, getLatexFormulas, updateLatexFormula, DOCUMENT_LATEX_RANGE_KIND } from '../src/features/document/latex/latex-formula.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { copyDocumentBlocksForPaste } from '../src/features/document/services/document-block-paste.ts'
import { decodeDocumentBody, encodeDocumentBody } from '../src/api/document-body-codec.ts'
import { readDocumentBodyLatexRanges } from '../src/api/document-body-latex.ts'
import { publicDocumentResources } from '../src/api/document-input-common.ts'
import { documentNoteInputFromPublic } from '../src/api/document-note-body.ts'
import { readDocumentPluginResources, refreshDocumentLatexPluginWrapper, writeDocumentPluginResources } from '../src/api/document-plugin-resources.ts'
import { wrapRuns } from '../src/features/document/layout/document-text-layout.ts'

// Deferred public-body/plugin contracts; no sample or runtime acceptance is claimed.
const copy = value => structuredClone(value)
const config = latex => ({ latex, kind: 'inline', displayMode: 'inline', authorHint: 'config' })
function sourceFor(latex = '  x^2  ') {
  const stream = 'AB' + latex + 'CD\r\n'
  return { body: { dataStream: stream, paragraphs: [{ paragraphId: 'p', startIndex: stream.length - 2 }],
    sectionBreaks: [{ startIndex: stream.length - 1, sectionId: 'first' }],
    textRuns: [{ st: 0, ed: stream.length - 2, ts: { fs: 14, bl: 1 } }],
    customRanges: [{ rangeType: 5, wholeEntity: true, rangeId: 'formula-id', startIndex: 2, endIndex: 2 + latex.length - 1,
      properties: { kind: DOCUMENT_LATEX_RANGE_KIND, authorHint: 'range' }, officeStudio: { version: 1, runId: 'formula-run', latexMode: 'inline' } }] },
    latexFormulas: { 'formula-id': config(latex) } }
}
const formulaRange = data => data.body.customRanges?.find(range => range.properties?.kind === DOCUMENT_LATEX_RANGE_KIND)
const named = (formulas, extra = {}) => ({ name: 'DOC_LATEX_PLUGIN', ownerHint: 'resource', data: JSON.stringify({ authorHint: 'wrapper', ...extra, formulas }) })

test('public source ranges collapse to one model atom and export the exact original source with inclusive boundaries and stable identities', () => {
  const source = sourceFor(), before = copy(source), decoded = decodeDocumentBody(source.body, { latexFormulas: source.latexFormulas })
  const item = getLatexFormulas({ blocks: decoded.content.blocks })[0]
  assert.equal(runText(decoded.content.blocks[0].runs), 'AB\uFFFCCD')
  assert.deepEqual([item.startOffset, item.endOffset, item.rangeId, item.runId], [2, 3, 'formula-id', 'formula-run'])
  const encoded = encodeDocumentBody(decoded.content), range = formulaRange(encoded)
  assert.equal(encoded.body.dataStream, source.body.dataStream)
  assert.deepEqual([range.startIndex, range.endIndex, range.wholeEntity], [2, 2 + item.formula.latex.length - 1, true])
  assert.equal(range.properties.authorHint, 'range'); assert.equal(encoded.latexFormulas['formula-id'].authorHint, 'config')
  assert.equal(range.officeStudio.runId, 'formula-run'); assert.deepEqual(source, before)
})

test('whitespace and source line breaks inside a formula do not create extra editor paragraphs or sections', () => {
  const source = sourceFor('  x\r\n + y\t '), decoded = decodeDocumentBody(source.body, source)
  assert.equal(decoded.content.blocks.length, 1)
  assert.equal(getLatexFormulas({ blocks: decoded.content.blocks })[0].formula.latex, '  x\r\n + y\t ')
  assert.equal(encodeDocumentBody(decoded.content).body.dataStream, source.body.dataStream)
})

test('unsupported imported syntax retains its raw source and displays a repairable error without a parser-based import rejection', () => {
  const source = sourceFor('  \\begin{matrix}a&b\\end{matrix}'), decoded = decodeDocumentBody(source.body)
  const item = getLatexFormulas({ blocks: decoded.content.blocks })[0]
  assert.equal(item.formula.latex, source.body.dataStream.slice(2, source.body.customRanges[0].endIndex + 1))
  const glyph = wrapRuns(item.run ? [item.run] : [], 1000, { measure: text => text.length * 8 })[0].glyphs[0]
  assert.equal(glyph.latexError.offset, 2)
  assert.equal(encodeDocumentBody(decoded.content).latexFormulas['formula-id'].latex, item.formula.latex)
})

test('malformed identity, wholeEntity, overlap and inline config contracts reject the whole input', () => {
  for (const mutate of [source => source.body.customRanges[0].wholeEntity = false,
    source => source.body.customRanges[0].rangeId = '', source => source.body.customRanges.push(copy(source.body.customRanges[0])),
    source => source.body.customRanges.push({ ...copy(source.body.customRanges[0]), rangeId: 'overlap', startIndex: 3 }),
    source => source.latexFormulas['formula-id'].latex = 'mismatch', source => source.latexFormulas['formula-id'].displayMode = 'block']) {
    const source = sourceFor(); mutate(source); const before = copy(source)
    assert.throws(() => decodeDocumentBody(source.body, source)); assert.deepEqual(source, before)
  }
})

test('UTF-16 surrogate interiors, reversed spans and structural control markers cannot become public formula entities', () => {
  for (const mutate of [source => source.body.customRanges[0].startIndex = 3,
    source => source.body.customRanges[0].endIndex = 2, source => source.body.customRanges[0].endIndex = 0]) {
    const source = sourceFor('😀x'); mutate(source)
    assert.throws(() => readDocumentBodyLatexRanges(source.body, source), RangeError)
  }
  const marker = sourceFor('x\x10y'); assert.throws(() => decodeDocumentBody(marker.body, marker), /结构标记/)
})

test('local block mode survives the body extension while the named public config remains inline-only', () => {
  const content = { blocks: [createLatexFormulaBlock('x^2')], runs: [], paragraphs: [], metadata: {} }, encoded = encodeDocumentBody(content)
  const range = formulaRange(encoded)
  assert.equal(range.officeStudio.latexMode, 'block')
  assert.deepEqual([encoded.latexFormulas[range.rangeId].kind, encoded.latexFormulas[range.rangeId].displayMode], ['inline', 'inline'])
  assert.equal(getLatexFormulas({ blocks: decodeDocumentBody(encoded.body, encoded).content.blocks })[0].formula.mode, 'block')
})

test('deletion removes the source range and modeled plugin config so body reopen cannot resurrect it', () => {
  const source = sourceFor(), decoded = decodeDocumentBody(source.body, source), model = { ...createBlankDocument(), blocks: decoded.content.blocks, metadata: decoded.content.metadata }
  const item = getLatexFormulas(model)[0], selection = { anchor: { ...item.target, offset: 2 }, focus: { ...item.target, offset: 3 } }
  executeDocumentCommand(model, selection, { id: 'doc.command.remove-latex-formula', rangeId: item.rangeId })
  const encoded = encodeDocumentBody({ ...decoded.content, blocks: model.blocks })
  assert.equal(formulaRange(encoded), undefined); assert.equal(encoded.latexFormulas?.['formula-id'], undefined)
  assert.equal(getLatexFormulas({ blocks: decodeDocumentBody(encoded.body, encoded).content.blocks }).length, 0)
  assert.equal(runText(model.blocks[0].runs), 'ABCD')
})

test('a real copied run exports a separate CUSTOM5 range and config while edits retain the imported ID', () => {
  const source = sourceFor(), decoded = decodeDocumentBody(source.body, source), copied = copyDocumentBlocksForPaste(decoded.content.blocks)
  const model = { blocks: decoded.content.blocks }, item = getLatexFormulas(model)[0]
  updateLatexFormula(model, item.runId, ' y_2 ')
  const encoded = encodeDocumentBody({ ...decoded.content, blocks: [...decoded.content.blocks, ...copied] })
  const ranges = encoded.body.customRanges.filter(range => range.properties?.kind === DOCUMENT_LATEX_RANGE_KIND)
  assert.equal(ranges.length, 2); assert.equal(ranges[0].rangeId, 'formula-id'); assert.notEqual(ranges[1].rangeId, ranges[0].rangeId)
  assert.equal(encoded.latexFormulas[ranges[0].rangeId].latex, ' y_2 ')
  assert.equal(encoded.latexFormulas[ranges[1].rangeId].latex, '  x^2  ')
})

test('named formula resources take precedence over compatibility aliases in both normal and note body inputs', () => {
  const source = sourceFor('x'), input = { ...source, latexFormulas: { 'formula-id': config('wrong') }, resources: [named(source.latexFormulas)] }
  const resources = publicDocumentResources(input)
  assert.equal(resources.latexFormulas['formula-id'].latex, 'x')
  const note = documentNoteInputFromPublic(createBlankDocument(), input, { kind: 'footnote' })
  assert.equal(getLatexFormulas({ blocks: note.content.blocks })[0].formula.latex, 'x')
  assert.equal(note.latexFormulas, undefined); assert.equal(note.resources, undefined)
  assert.throws(() => documentNoteInputFromPublic(createBlankDocument(), { resources: input.resources }), /body/)
})

test('every nested wrapper retains unknown fields while deleted configs disappear and conflicting live identities are rejected', () => {
  const source = sourceFor(), decoded = decodeDocumentBody(source.body, source), model = { blocks: decoded.content.blocks }
  updateLatexFormula(model, 'formula-run', 'z')
  const live = encodeDocumentBody(decoded.content), original = { body: source.body, resources: [named(source.latexFormulas)],
    headers: { header: { body: source.body, resources: [{ ...named(source.latexFormulas, { nestedHint: 9 }), ownerHint: 'nested-resource' }] } } }
  const before = copy(original), resources = writeDocumentPluginResources(original, [live]), configs = readDocumentPluginResources(resources).latexFormulas
  const nested = { body: live.body, resources: copy(original.headers.header.resources) }
  refreshDocumentLatexPluginWrapper(nested, configs)
  const wrapper = JSON.parse(nested.resources[0].data)
  assert.equal(wrapper.nestedHint, 9); assert.equal(nested.resources[0].ownerHint, 'nested-resource')
  assert.equal(wrapper.formulas['formula-id'].latex, 'z'); assert.deepEqual(original, before)
  const empty = { body: { dataStream: '\r\n' } }, deletedResources = writeDocumentPluginResources(original, [empty])
  refreshDocumentLatexPluginWrapper(nested, readDocumentPluginResources(deletedResources).latexFormulas ?? {})
  assert.deepEqual(JSON.parse(nested.resources[0].data).formulas, {}); assert.equal(JSON.parse(nested.resources[0].data).nestedHint, 9)
  assert.equal(JSON.parse(deletedResources[0].data).authorHint, 'wrapper')
  const conflict = copy(live); conflict.latexFormulas['formula-id'].latex = 'different'
  assert.throws(() => writeDocumentPluginResources(original, [live, conflict]), /不同配置/)
})
