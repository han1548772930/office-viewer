import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection, createParagraph, createRun, createTable, documentHtml, runText, textTargets } from '../src/features/document/model/document-model.ts'
import { assertLatexFormulaSource, createLatexFormulaBlock, createLatexFormulaRun, getLatexFormula, getLatexFormulas,
  insertLatexFormula, latexFormulaRangeId, layoutLatexFormula, LatexFormulaError, parseLatexFormula, selectedLatexFormula, updateLatexFormula } from '../src/features/document/latex/latex-formula.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { documentLatexSelectionState } from '../src/features/document/commands/document-latex-commands.ts'
import { documentLatexCanvasRequestCurrent } from '../src/features/document/controllers/document-latex-canvas.ts'
import { captureDocumentStructuredCanvasTarget } from '../src/features/document/controllers/document-structured-canvas.ts'
import { documentLatexDragAnchor, hitTestDocumentLatex } from '../src/features/document/layout/document-latex-hit.ts'
import { positionLine, wrapRuns } from '../src/features/document/layout/document-text-layout.ts'
import { copyDocumentBlocksForPaste } from '../src/features/document/services/document-block-paste.ts'
import { createDocumentContentControl } from '../src/features/document/services/content-controls.ts'
import { documentTextFields } from '../src/features/document/model/document-text-fields.ts'
import { documentLatexContextItems } from '../src/components/document/document-latex-menu.ts'

// Deferred model, command and geometry drafts. This batch has not run them.
const copy = value => structuredClone(value)
const point = (blockId, offset) => ({ blockId, offset })
const select = (blockId, start, end = start) => ({ anchor: point(blockId, start), focus: point(blockId, end) })
const paragraph = (id, text = '') => ({ ...createParagraph(text), id })
const withBlocks = blocks => ({ ...createBlankDocument(), id: 'latex-contract', blocks })
const mathParagraph = (id, latex = 'x^2', mode = 'inline') => ({ ...paragraph(id), runs: [createLatexFormulaRun(latex, mode)] })
const measure = (text, font) => Array.from(text).length * Number(/([\d.]+)px/.exec(font)?.[1] ?? 16) * 0.5
const permission = { documentId: 'latex-contract', userId: 'editor' }

test('raw whitespace, text-group spaces and scripts after grouped or spaced bases retain their semantics', () => {
  const tree = parseLatexFormula('  {x} ^ {2} _1 + \\text{a b}  ')
  assert.equal(tree.children[0].type, 'script')
  assert.equal(tree.children[0].sup.value, '2'); assert.equal(tree.children[0].sub.value, '1')
  assert.equal(tree.children.at(-1).type, 'styled'); assert.equal(tree.children.at(-1).style, 'roman')
  assert.deepEqual(tree.children.at(-1).value.children.map(node => node.value), ['a', ' ', 'b'])
  const run = createLatexFormulaRun('  x ^2  ')
  assert.equal(getLatexFormula(run).latex, '  x ^2  ')
})

test('parser error positions use the original source and forbidden structural markers fail at their real offset', () => {
  assert.throws(() => parseLatexFormula('  x + \\bad'), error => error instanceof LatexFormulaError && error.offset === 6)
  assert.throws(() => parseLatexFormula('x^2 ^{3}'), error => error instanceof LatexFormulaError && error.offset === 4)
  assert.throws(() => parseLatexFormula('x_{ }_2'), error => error instanceof LatexFormulaError && error.offset === 5)
  for (const source of ['', ' \t\n', 'a\x10b', 'a\x1fb']) assert.throws(() => assertLatexFormulaSource(source), LatexFormulaError)
  assertLatexFormulaSource('x\r\n +\t y')
  assert.throws(() => updateLatexFormula(withBlocks([mathParagraph('p')]), 'missing', 'x\x10y'), LatexFormulaError)
})

test('astral characters remain complete text atoms with UTF-16 error locations', () => {
  const tree = parseLatexFormula('😀^2')
  assert.equal(tree.base.value, '😀')
  assert.throws(() => parseLatexFormula('😀 + \\bad'), error => error instanceof LatexFormulaError && error.offset === 5)
  assert.equal(layoutLatexFormula('😀^2', { measure }).primitives[0].text, '😀')
})

test('inline insertion inherits selected character marks and explicit pending marks without inheriting formula provenance', () => {
  const block = paragraph('p', 'abc'); Object.assign(block.runs[0], { bold: true, fontSize: 19, color: '#123456' })
  const model = withBlocks([block]), result = insertLatexFormula(model, select('p', 1, 2), '  x^2 ', 'inline', { italic: true, fontSize: 23 })
  const item = getLatexFormulas(model)[0]
  assert.equal(runText(block.runs), 'a\uFFFCc'); assert.equal(result.offset, 2)
  assert.deepEqual([item.run.bold, item.run.italic, item.run.fontSize, item.run.color], [true, true, 23, '#123456'])
  executeDocumentCommand(model, collapsedSelection(result), { id: 'doc.command.insert-text', text: 'tail' }, permission)
  assert.equal(getLatexFormulas(model).length, 1)
  assert.equal(block.runs.find(run => run.text.includes('tail')).metadata?.latexFormula, undefined)
})

test('imported range identity survives updates while the actual paste copier assigns an independent identity', () => {
  const block = mathParagraph('p'), run = block.runs[0]
  run.metadata.univerBodyLatexRangeSource = { rangeId: 'public-formula', runId: run.id, authorHint: 'range' }
  const model = withBlocks([block])
  assert.equal(updateLatexFormula(model, run.id, '  y^3 '), true)
  assert.equal(latexFormulaRangeId(run), 'public-formula')
  const copied = copyDocumentBlocksForPaste([block])[0].runs[0]
  assert.notEqual(copied.id, run.id); assert.equal(latexFormulaRangeId(copied), copied.id)
  assert.equal(getLatexFormula(copied).latex, '  y^3 ')
})

test('only a caret at the atom or its exact selection enters formula editing', () => {
  const block = paragraph('p'); block.runs = [createRun('a'), createLatexFormulaRun('x'), createRun('b')]
  const model = withBlocks([block]), id = getLatexFormulas(model)[0].rangeId
  assert.equal(selectedLatexFormula(model, select('p', 1)).rangeId, id)
  assert.equal(selectedLatexFormula(model, select('p', 2)), undefined)
  assert.equal(selectedLatexFormula(model, select('p', 1, 2)).rangeId, id)
  assert.equal(selectedLatexFormula(model, select('p', 0, 3)), undefined)
  const invalid = copy(block.runs[1]); invalid.text = 'literal\uFFFC'; assert.equal(getLatexFormula(invalid), undefined)
})

test('formula discovery and updates traverse containers, columns, tables and merged-cell continuation content', () => {
  const table = createTable(2, 1), cell = table.rows[0].cells[0]
  cell.rowSpan = 2; cell.blocks = [mathParagraph('cell')]; table.rows[1].cells = []
  cell.style.nativeContinuationContent = { 1: { version: 1, ownerCellId: cell.id, rowId: table.rows[1].id, colSpan: 1, blocks: [mathParagraph('continuation')] } }
  const model = withBlocks([{ ...paragraph('quote'), type: 'quote', runs: [], children: [mathParagraph('quote-p')] },
    { ...paragraph('columns'), type: 'columns', runs: [], columns: { version: 1, gap: 16, horizontalPadding: 0, minHeight: 24,
      layout: 'fixed', responsive: 'shrink', columns: [{ id: 'left', widthRatio: 1, blocks: [mathParagraph('column-p')] }] } },
    { ...paragraph('table'), type: 'table', runs: [], metadata: { table } }])
  assert.deepEqual(getLatexFormulas(model).map(item => item.blockId), ['quote-p', 'column-p', 'cell', 'continuation'])
  const item = getLatexFormulas(model).at(-1), target = textTargets(model).find(target => target.cellBlockId === 'continuation')
  assert.deepEqual(item.target, target)
  executeDocumentCommand(model, { anchor: { ...target, offset: 0 }, focus: { ...target, offset: 1 } },
    { id: 'doc.command.update-latex-formula', rangeId: item.rangeId, latex: 'z_2' }, permission)
  assert.equal(getLatexFormulas(model).at(-1).formula.latex, 'z_2')
})

test('content locks, field locks and target permissions reject formula mutations before source changes', () => {
  for (const lock of ['content', 'field', 'permission']) {
    const model = withBlocks([mathParagraph('p')]), item = getLatexFormulas(model)[0], selection = select('p', 0, 1)
    if (lock === 'content') createDocumentContentControl(model, { id: 'locked', type: 'richText', blockId: 'p', startOffset: 0, endOffset: 1, lock: 'contentLocked' })
    if (lock === 'field') model.blocks[0].metadata.documentTextFields = { version: 1, fields: [{ id: 'locked', field: 'UNKNOWN', startOffset: 0, endOffset: 1, locked: true }] }
    if (lock === 'permission') model.permissionRules = [{ id: 'denied', scope: 'entity', targetId: 'p', action: 'edit', userIds: ['editor'], effect: 'deny' }]
    const before = copy(model)
    assert.equal(documentLatexSelectionState(model, selection, permission).available, false)
    assert.throws(() => executeDocumentCommand(model, selection, { id: 'doc.command.update-latex-formula', rangeId: item.rangeId, latex: 'y' }, permission))
    assert.deepEqual(model, before)
  }
})

test('inline and block insertion remap adjacent control, field and comment anchors to real new paragraphs', () => {
  for (const mode of ['inline', 'block']) {
    const block = paragraph('p', 'abcdef'), model = withBlocks([block])
    createDocumentContentControl(model, { id: 'after', type: 'text', blockId: 'p', startOffset: 4, endOffset: 6 })
    block.metadata.documentTextFields = { version: 1, fields: [{ id: 'after-field', field: 'UNKNOWN', startOffset: 4, endOffset: 6 }] }
    model.comments = { c: { id: 'c', unitId: model.id, anchor: [{ blockId: 'p', startOffset: 4, endOffset: 6 }], messages: [] } }
    insertLatexFormula(model, select('p', 1, 3), 'x', mode)
    const target = mode === 'inline' ? model.blocks[0] : model.blocks.at(-1), expected = mode === 'inline' ? 3 : 1
    assert.equal(model.contentControls[0].blockId, target.id); assert.equal(model.contentControls[0].startOffset, expected)
    assert.equal(documentTextFields(target)[0].startOffset, expected)
    assert.equal(model.comments.c.anchor[0].blockId, target.id); assert.equal(model.comments.c.anchor[0].startOffset, expected)
  }
})

test('remove and source-text replacement affect one formula and ordinary typing clears the stale display-block marker', () => {
  const block = createLatexFormulaBlock('x^2'), model = withBlocks([block]), item = getLatexFormulas(model)[0]
  executeDocumentCommand(model, select(block.id, 1), { id: 'doc.command.insert-text', text: 'suffix' }, permission)
  assert.equal(block.metadata.latexFormulaBlock, undefined)
  assert.ok(documentHtml(model).includes('suffix'))
  executeDocumentCommand(model, select(block.id, 0), { id: 'doc.command.replace-latex-formula-with-text', rangeId: item.rangeId, text: 'x^2' }, permission)
  assert.equal(getLatexFormulas(model).length, 0); assert.equal(runText(block.runs), 'x^2suffix')
  assert.equal(block.runs.some(run => run.metadata?.latexFormula), false)
  assert.throws(() => executeDocumentCommand(model, select(block.id, 0), { id: 'doc.command.remove-latex-formula', rangeId: item.rangeId }, permission), /删除|替换/)
})

test('nested fraction, radical and script metrics enclose every primitive around the same baseline', () => {
  for (const source of ['\\frac{\\frac{a}{b}}{\\sqrt{c_2}}', '\\sqrt{\\frac{x}{y}}^2', 'x_2^3']) {
    const layout = layoutLatexFormula(source, { fontSize: 24, measure }), top = -layout.baseline, bottom = layout.height - layout.baseline
    for (const primitive of layout.primitives) {
      const size = primitive.kind === 'text' ? Number(/([\d.]+)px/.exec(primitive.font)[1]) : 0
      assert.ok((primitive.kind === 'text' ? primitive.y - size * 0.9 : Math.min(primitive.y1, primitive.y2) - primitive.width / 2) >= top - 1e-8)
      assert.ok((primitive.kind === 'text' ? primitive.y + size * 0.3 : Math.max(primitive.y1, primitive.y2) + primitive.width / 2) <= bottom + 1e-8)
    }
  }
})

test('mixed text/formula wrapping uses ascent and descent and an invalid persisted formula has an explicit error layout', () => {
  const run = createLatexFormulaRun('\\frac{a}{b}'), invalid = createLatexFormulaRun('x'); invalid.metadata.latexFormula.latex = '  \\bad'
  const raw = wrapRuns([createRun('left'), run, createRun('right')], 1000, { measure })[0]
  const line = positionLine(raw, { blockId: 'p' }, 20, 40, 1000, raw.fontHeight * 1.3, 0, 'left', false)
  const glyph = line.glyphs.find(glyph => glyph.latex)
  assert.ok(glyph.y >= line.y); assert.ok(glyph.y + glyph.height <= line.y + line.height + 1e-8)
  assert.equal(glyph.y + glyph.latex.baseline, line.baseline)
  const error = wrapRuns([invalid], 1000, { measure })[0].glyphs[0]
  assert.equal(error.latexError.offset, 2); assert.ok(error.latex.primitives[0].text.startsWith('⚠'))
  assert.equal(getLatexFormula(invalid).latex, '  \\bad')
})

test('both formula halves hit one atom and clipping or stale source offsets cannot select the wrong object', () => {
  const block = mathParagraph('p'), model = withBlocks([block]), raw = wrapRuns(block.runs, 1000, { measure })[0]
  const line = positionLine(raw, { blockId: 'p' }, 20, 40, 1000, raw.fontHeight * 1.2, 3, 'left', false), glyph = line.glyphs[0]
  for (const fraction of [0.1, 0.9]) {
    const hit = hitTestDocumentLatex(model, { lines: [line] }, glyph.x + glyph.width * fraction, glyph.y + glyph.height / 2)
    assert.deepEqual([hit.selection.anchor.offset, hit.selection.focus.offset, hit.selection.focus.pageAffinity], [0, 1, 3])
  }
  assert.equal(hitTestDocumentLatex(model, { lines: [{ ...line, clip: { x: 0, y: 0, width: 5, height: 5 } }] }, glyph.x, glyph.y), undefined)
  block.runs.unshift(createRun('prefix'))
  assert.equal(hitTestDocumentLatex(model, { lines: [line] }, glyph.x, glyph.y), undefined)
})

test('formula drag anchors include the complete source atom in both directions', () => {
  const block = paragraph('p'); block.runs = [createRun('a'), createLatexFormulaRun('x'), createRun('b')]
  const model = withBlocks([block]), atom = select('p', 1, 2)
  assert.equal(documentLatexDragAnchor(model, atom, point('p', 0)).offset, 2)
  assert.equal(documentLatexDragAnchor(model, atom, point('p', 3)).offset, 1)
})

test('Canvas edit requests reject changed revision, session, host, selection and standalone in-place content', () => {
  const model = withBlocks([mathParagraph('p')]), selection = select('p', 0, 1), rangeId = getLatexFormulas(model)[0].rangeId
  const session = { valid: true, modelRevision: 1, generation: 3, marks: {}, location: { kind: 'body', selection } }
  const request = { ...captureDocumentStructuredCanvasTarget(model, selection, undefined, session), rangeId }, current = { source: model, selection, session, requireSession: true }
  assert.equal(documentLatexCanvasRequestCurrent(request, current), true)
  for (const next of [{ session: { ...session, generation: 4 } }, { session: { ...session, modelRevision: 2 } }, { host: { kind: 'note', noteKind: 'footnote', noteId: 'other' } },
    { selection: select('p', 1) }, { source: copy(model) }, { disabled: true }]) assert.equal(documentLatexCanvasRequestCurrent(request, { ...current, ...next }), false)
  const standalone = { ...captureDocumentStructuredCanvasTarget(model, selection), rangeId }
  assert.equal(documentLatexCanvasRequestCurrent(standalone, { source: model, selection }), true)
  model.blocks[0].runs[0].metadata.latexFormula.latex = 'y'
  assert.equal(documentLatexCanvasRequestCurrent(standalone, { source: model, selection }), false)
})

test('formula context actions keep the stable ID and exact raw source and disable all mutations under a lock', () => {
  const item = getLatexFormulas(withBlocks([mathParagraph('p', '  x ^2  ')]))[0]
  const actions = documentLatexContextItems(item, { available: true }, key => key)
  assert.deepEqual(actions.map(action => action.commandId), ['docs-latex.operation.open-formula-popup', 'doc.command.remove-latex-formula', 'doc.command.replace-latex-formula-with-text'])
  assert.equal(actions[2].params.text, '  x ^2  ')
  assert.ok(actions.every(action => action.params.rangeId === item.rangeId))
  assert.ok(documentLatexContextItems(item, { available: false, reason: 'locked' }, key => key).every(action => action.disabled && action.disabledReason === 'locked'))
})
