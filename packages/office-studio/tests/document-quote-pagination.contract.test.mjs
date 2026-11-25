import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument, createDocumentDefaultStyle } from '../src/domain/model.ts'
import { createParagraph, createRun, runText } from '../src/features/document/model/document-model.ts'
import { clearDocumentQuotePagination, documentParagraphPagination, documentUsesTraditionalPagination,
  initializeDocumentQuotePagination, reconcileDocumentQuotePagination } from '../src/features/document/model/document-paragraph-pagination.ts'
import { documentDefaultStyle } from '../src/features/document/model/document-default-style.ts'
import { applyDocumentParagraphStylePatch, validateDocumentParagraphStylePatch } from '../src/features/document/services/document-settings.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { documentParagraphStyleFromBody, documentParagraphStyleToBody } from '../src/api/document-body-style.ts'
import { documentParagraphFromBody, documentParagraphToBody } from '../src/api/document-body-paragraph.ts'
import { decodeDocumentBody, encodeDocumentBody } from '../src/api/document-body-codec.ts'
import { layoutDocument, caretRect, hitTestDocument, emptyLayout } from '../src/features/document/layout/document-layout.ts'
import { appendDocumentBlocks } from '../src/features/document/layout/document-block-layout.ts'
import { DocumentSectionFlow } from '../src/features/document/layout/document-section-flow.ts'
import { planDocumentParagraphFragment } from '../src/features/document/layout/document-pagination-rules.ts'
import { DOCUMENT_KEEP_NEXT_SCAN_LIMIT } from '../src/features/document/layout/document-pagination-layout.ts'

// Deferred author/model/body/layout contracts; no checks or browser acceptance
// have been run for this batch. Native DOCX and complex cell compatibility are separate.
const copy = value => structuredClone(value)
const paragraph = (id, lines = 1) => ({ ...createParagraph(Array.from({ length: lines }, (_, index) => 'line ' + index).join('\n')),
  id, style: { align: 'left', indent: 0, spacingRule: 'exact', lineHeight: 20, spaceBefore: 0, spaceAfter: 0 } })
const quote = (id, children) => ({ ...paragraph(id), type: 'quote', runs: [], children })
const content = blocks => ({ blocks, runs: [], paragraphs: [], metadata: {} })
const documentWith = (blocks, metadata = { documentFlavor: 1 }) => ({ ...createBlankDocument(), id: 'pagination-contract', blocks, metadata,
  page: { width: 320, height: 180, marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20 } })
const measurer = { measure: text => Array.from(text).length * 6 }
const pagesOf = (layout, id) => [...new Set(layout.lines.filter(line => line.target.blockId === id).map(line => line.page))]
const selection = (id, offset = 0) => ({ anchor: { blockId: id, offset }, focus: { blockId: id, offset } })

test('four pagination booleans inherit defaults and an explicit false overrides each inherited true', () => {
  const defaults = createDocumentDefaultStyle(), model = { defaultStyle: defaults }
  Object.assign(defaults.paragraph, { keepLines: true, keepNext: true, widowControl: true, pageBreakBefore: true })
  assert.deepEqual(documentParagraphPagination({}, model), { keepLines: true, keepNext: true, widowControl: true, pageBreakBefore: true })
  assert.deepEqual(documentParagraphPagination({ keepLines: false, keepNext: false, widowControl: false, pageBreakBefore: false }, model),
    { keepLines: false, keepNext: false, widowControl: false, pageBreakBefore: false })
})

test('settings and snapshot validation reject malformed pagination switches', () => {
  for (const key of ['keepLines', 'keepNext', 'widowControl', 'pageBreakBefore']) for (const value of [0, 1, 'true', null, [], {}]) {
    assert.throws(() => validateDocumentParagraphStylePatch({ [key]: value }), /boolean/)
    const defaults = createDocumentDefaultStyle(); defaults.paragraph[key] = value
    assert.throws(() => documentDefaultStyle({ defaultStyle: defaults }), /段落格式/)
  }
  for (const value of [false, true]) validateDocumentParagraphStylePatch({ keepLines: value, keepNext: value, widowControl: value, pageBreakBefore: value })
})

test('public BooleanNumber mappings preserve numeric false, true and unknown author properties', () => {
  const defaults = createDocumentDefaultStyle(), source = { keepLines: 1, keepNext: 0, widowControl: true, pageBreakBefore: false, authorHint: 'retained' }
  const style = documentParagraphStyleFromBody(source, defaults)
  assert.deepEqual(documentParagraphPagination(style), { keepLines: true, keepNext: false, widowControl: true, pageBreakBefore: false })
  const exported = documentParagraphStyleToBody(style, source)
  assert.deepEqual([exported.keepLines, exported.keepNext, exported.widowControl, exported.pageBreakBefore], [1, 0, 1, 0])
  assert.equal(exported.authorHint, 'retained')
  for (const value of [2, -1, 'false', {}]) assert.throws(() => documentParagraphStyleFromBody({ keepNext: value }, defaults), /0\/1/)
})

test('named-style pagination remains inherited through body export and same-value writes establish direct formatting', () => {
  const defaults = createDocumentDefaultStyle(), raw = { paragraphId: 'named', styleId: 'child', paragraphStyle: { keepNext: 0, authorHint: 'source' } }
  const options = { defaults, styles: { parent: { paragraphStyle: { keepLines: 1, keepNext: 1 } },
    child: { basedOn: 'parent', paragraphStyle: { widowControl: 1 } } } }
  const block = documentParagraphFromBody(raw, [createRun('named')], options), before = copy(raw)
  assert.deepEqual(documentParagraphPagination(block.style), { keepLines: true, keepNext: false, widowControl: true, pageBreakBefore: false })
  let written = documentParagraphToBody(block, 5)
  assert.equal(written.paragraphStyle.keepLines, undefined); assert.equal(written.paragraphStyle.widowControl, undefined)
  assert.equal(written.paragraphStyle.keepNext, 0); assert.equal(written.paragraphStyle.authorHint, 'source'); assert.deepEqual(raw, before)
  applyDocumentParagraphStylePatch(block, { keepLines: true, widowControl: false })
  written = documentParagraphToBody(block, 5)
  assert.equal(written.paragraphStyle.keepLines, 1); assert.equal(written.paragraphStyle.widowControl, 0)
})

test('quote defaults link adjacent paragraphs and stop at the last paragraph and explicit breaks', () => {
  const first = paragraph('first'), second = paragraph('second'), last = paragraph('last'), stop = { ...paragraph('break'), type: 'pageBreak', runs: [] }
  const block = quote('quote', [first, second, stop, last]); initializeDocumentQuotePagination(block)
  assert.deepEqual([first.style.keepNext, second.style.keepNext, last.style.keepNext], [true, false, false])
  for (const paragraph of [first, second, last]) assert.deepEqual([paragraph.style.keepLines, paragraph.style.widowControl], [true, true])
  assert.equal(stop.style.keepLines, undefined)
})

test('quote removal restores prior authorship while retaining independent author overrides', () => {
  const first = paragraph('first'), last = paragraph('last')
  Object.assign(first.style, { keepLines: false, keepNext: true, widowControl: false, pageBreakBefore: true })
  const block = quote('quote', [first, last]); initializeDocumentQuotePagination(block)
  applyDocumentParagraphStylePatch(last, { keepLines: true, keepNext: true, widowControl: false })
  clearDocumentQuotePagination(block)
  assert.deepEqual(documentParagraphPagination(first.style), { keepLines: false, keepNext: true, widowControl: false, pageBreakBefore: true })
  assert.deepEqual(documentParagraphPagination(last.style), { keepLines: true, keepNext: true, widowControl: false, pageBreakBefore: false })
  assert.equal(first.metadata.documentQuotePaginationSource, undefined)
})

test('reconciliation updates the old and new last paragraph without overriding an author break', () => {
  const first = paragraph('first'), last = paragraph('last'), block = quote('quote', [first, last]), model = documentWith([block, paragraph('body')])
  initializeDocumentQuotePagination(block); applyDocumentParagraphStylePatch(first, { keepNext: false })
  const inserted = paragraph('new-last'); block.children.push(inserted); reconcileDocumentQuotePagination(model)
  assert.deepEqual(block.children.map(paragraph => paragraph.style.keepNext), [false, true, false])
  block.children.pop(); reconcileDocumentQuotePagination(model)
  assert.deepEqual(block.children.map(paragraph => paragraph.style.keepNext), [false, false])
  assert.equal(model.blocks[1].style.keepLines, undefined)
})

test('moving a paragraph out of a managed quote restores only that wrapper provenance', () => {
  const child = paragraph('inside'), block = quote('quote', [child]), model = documentWith([block])
  initializeDocumentQuotePagination(block); block.children = []; model.blocks.push(child); reconcileDocumentQuotePagination(model)
  assert.equal(child.style.keepLines, undefined); assert.equal(child.style.keepNext, undefined)
  assert.equal(child.metadata.documentQuotePaginationSource, undefined)
})

test('public-body save/reopen keeps quote author restoration through multiple cycles', () => {
  const defaults = createDocumentDefaultStyle(); defaults.paragraph.keepLines = true
  const first = documentParagraphFromBody({ paragraphId: 'first', paragraphStyle: { keepNext: 1, widowControl: 0 } }, [createRun('quote')], { defaults })
  const block = quote('quote', [first]); initializeDocumentQuotePagination(block)
  let current = content([block])
  for (let round = 0; round < 2; round++) {
    const encoded = encodeDocumentBody(current, defaults)
    assert.equal(encoded.body.paragraphs[0].paragraphStyle.keepNext, 0)
    current = decodeDocumentBody(encoded.body, encoded, { defaults }).content
  }
  clearDocumentQuotePagination(current.blocks[0])
  const restored = encodeDocumentBody(content(current.blocks[0].children), defaults).body.paragraphs[0].paragraphStyle
  assert.equal(restored.keepLines, undefined); assert.equal(restored.keepNext, 1); assert.equal(restored.widowControl, 0)
})

test('unverified imported quote keep properties survive unwrap', () => {
  const child = paragraph('unverified'); Object.assign(child.style, { keepLines: true, keepNext: false, widowControl: true })
  const block = quote('quote', [child]); clearDocumentQuotePagination(block)
  assert.deepEqual([child.style.keepLines, child.style.keepNext, child.style.widowControl], [true, false, true])
})

test('widow planning moves a one-line opening and avoids a one-line ending', () => {
  assert.deepEqual(planDocumentParagraphFragment([20, 20, 20, 20], 0, 25, 140, true), { lineCount: 0, requiredHeight: 40 })
  assert.equal(planDocumentParagraphFragment([20, 20, 20, 20, 20], 0, 80, 140, true).lineCount, 3)
  assert.equal(planDocumentParagraphFragment([20, 20, 20, 20, 20], 0, 80, 140, false).lineCount, 4)
})

test('widow planning uses variable line heights, host padding and impossible-frame fallbacks', () => {
  assert.deepEqual(planDocumentParagraphFragment([12, 30, 14], 0, 42, 70, true, 8), { lineCount: 0, requiredHeight: 64 })
  assert.equal(planDocumentParagraphFragment([40, 40, 40], 0, 50, 50, true).lineCount, 1)
  assert.deepEqual(planDocumentParagraphFragment([200], 0, 140, 140, true), { lineCount: 0, requiredHeight: 200 })
  assert.equal(planDocumentParagraphFragment([20], 1, 140, 140, true).requiredHeight, 0)
})

test('traditional is public mode 1 and public unspecified/modern/drawing modes retain inert pagination rules', () => {
  assert.equal(documentUsesTraditionalPagination({ metadata: {} }), true)
  assert.equal(documentUsesTraditionalPagination({ metadata: { documentFlavor: 0 } }), true)
  for (const documentFlavor of [0, 1, 2, 3]) {
    const metadata = { documentFlavor, documentFlavorContract: 'univer', documentStyle: { documentFlavor } }
    assert.equal(documentUsesTraditionalPagination({ metadata }), documentFlavor === 1)
    const block = paragraph('kept', 3); block.style.keepLines = true
    const model = documentWith([paragraph('before', 5), block], metadata), before = copy(model)
    assert.deepEqual(pagesOf(layoutDocument(model, measurer), 'kept'), documentFlavor === 1 ? [1] : [0, 1])
    assert.deepEqual(model, before)
  }
})

test('a fitting keepLines paragraph moves wholly to a fresh page', () => {
  const before = paragraph('before', 5), kept = paragraph('kept', 3); kept.style.keepLines = true
  const layout = layoutDocument(documentWith([before, kept]), measurer)
  assert.deepEqual(pagesOf(layout, 'before'), [0]); assert.deepEqual(pagesOf(layout, 'kept'), [1])
  assert.equal(layout.lines.find(line => line.target.blockId === 'kept').y, layout.pages[1].content.y)
})

test('oversized keepLines paragraphs start fresh then split without blank content pages', () => {
  const before = paragraph('before', 3), kept = paragraph('kept', 12); kept.style.keepLines = true; kept.style.widowControl = true
  const layout = layoutDocument(documentWith([before, kept]), measurer)
  assert.deepEqual(pagesOf(layout, 'kept'), [1, 2])
  assert.equal(layout.lines.filter(line => line.target.blockId === 'kept').length, 12)
  assert.equal(layout.pages.length, 3)
})

test('a fitting multi-paragraph quote starts together and does not drag the following body', () => {
  const before = paragraph('before', 5), first = paragraph('first'), last = paragraph('last'), block = quote('quote', [first, last]), tail = paragraph('tail')
  initializeDocumentQuotePagination(block)
  const layout = layoutDocument(documentWith([before, block, tail]), measurer)
  assert.deepEqual(pagesOf(layout, 'first'), [1]); assert.deepEqual(pagesOf(layout, 'last'), [1])
  assert.deepEqual(pagesOf(layout, 'tail'), [1]); assert.equal(last.style.keepNext, false)
  const decoration = layout.containers.filter(item => item.blockId === 'quote')
  assert.equal(decoration.length, 1); assert.equal(decoration[0].page, 1)
  assert.ok(decoration[0].y <= layout.lines.find(line => line.target.blockId === 'first').y - 8)
})

test('a taller-than-page quote has content on every decorated fragment and retains hit offsets', () => {
  const before = paragraph('before', 3), kept = paragraph('inside', 12), block = quote('quote', [kept]); initializeDocumentQuotePagination(block)
  const model = documentWith([before, block]), layout = layoutDocument(model, measurer), lines = layout.lines.filter(line => line.target.blockId === 'inside')
  assert.ok(pagesOf(layout, 'inside').length > 1)
  for (const frame of layout.containers.filter(item => item.blockId === 'quote')) assert.ok(lines.some(line => line.page === frame.page))
  const line = lines.at(-1), point = { blockId: 'inside', offset: line.start }, caret = caretRect(layout, point)
  assert.ok(caret); assert.equal(hitTestDocument(layout, caret.x + 1, caret.y + 2).offset, point.offset)
  assert.equal(runText(kept.runs), runText(model.blocks[1].children[0].runs))
})

test('widow control produces two-line edges while an explicit false permits the greedy split', () => {
  const first = paragraph('before', 6), controlled = paragraph('controlled', 5); controlled.style.widowControl = true
  let layout = layoutDocument(documentWith([first, controlled]), measurer)
  assert.deepEqual(pagesOf(layout, 'controlled'), [1])
  controlled.style.widowControl = false; layout = layoutDocument(documentWith([first, controlled]), measurer)
  assert.deepEqual(pagesOf(layout, 'controlled'), [0, 1])
  const before = paragraph('prefix', 3), split = paragraph('split', 5); split.style.widowControl = true
  layout = layoutDocument(documentWith([before, split]), measurer)
  assert.deepEqual([0, 1].map(page => layout.lines.filter(line => line.target.blockId === 'split' && line.page === page).length), [3, 2])
})

test('pageBreakBefore reuses a blank boundary and overrides keepNext adjacency', () => {
  const first = paragraph('first'), next = paragraph('next'); first.style.keepNext = true; next.style.pageBreakBefore = true
  let layout = layoutDocument(documentWith([first, next]), measurer)
  assert.deepEqual(pagesOf(layout, 'first'), [0]); assert.deepEqual(pagesOf(layout, 'next'), [1])
  const emptyBoundary = { ...paragraph('break'), type: 'pageBreak', runs: [] }
  layout = layoutDocument(documentWith([first, emptyBoundary, next]), measurer)
  assert.equal(layout.pages.length, 2); assert.deepEqual(pagesOf(layout, 'next'), [1])
  layout = layoutDocument(documentWith([next]), measurer); assert.deepEqual(pagesOf(layout, 'next'), [0])
})

test('keepNext uses the next opening when a following paragraph can split and respects the scan budget', () => {
  assert.equal(DOCUMENT_KEEP_NEXT_SCAN_LIMIT, 32)
  const before = paragraph('before', 5), first = paragraph('first'), next = paragraph('next', 8); first.style.keepNext = true
  const layout = layoutDocument(documentWith([before, first, next]), measurer)
  assert.equal(layout.lines.find(line => line.target.blockId === 'first').page, layout.lines.find(line => line.target.blockId === 'next').page)
  assert.ok(pagesOf(layout, 'next').length > 1)
})

test('section columns use the same keep rules and pageBreakBefore starts the next physical page', () => {
  const before = paragraph('before', 5), kept = paragraph('kept', 3), next = paragraph('next'); kept.style.keepLines = true; next.style.pageBreakBefore = true
  const model = documentWith([before, kept, next]); model.sectionSetup = { columnCount: 2, columnGap: 18 }
  const layout = layoutDocument(model, measurer), keptLines = layout.lines.filter(line => line.target.blockId === 'kept')
  assert.ok(keptLines.every(line => line.page === 0 && line.column === 1)); assert.deepEqual(pagesOf(layout, 'next'), [1])
})

test('note reference floors move a fitting kept paragraph together and layout never edits author data', () => {
  const block = paragraph('kept', 2); block.style.keepLines = true
  block.runs = [createRun('line 0\n'), createRun('')]
  const referenceId = block.runs[1].id
  const model = documentWith([block]), before = copy(model)
  // Exercise the same preflight/line path used by the note reservation passes.
  const flow = new DocumentSectionFlow(model, { referencePageFloors: new Map([[referenceId, 2]]) }), result = emptyLayout()
  appendDocumentBlocks(model.blocks, flow, result, measurer, { textDisplay: { references: new Map([[referenceId, { noteKind: 'footnote', noteId: 'note', label: '1' }]]) } })
  assert.equal(result.lines.length, 2); assert.ok(result.lines.every(line => line.page === 2))
  assert.deepEqual(model, before)
})

test('oversized keepNext predecessors keep their last feasible fragment with the successor opening', () => {
  const first = paragraph('first', 13), next = paragraph('next', 3); first.style.keepLines = true; first.style.keepNext = true
  const layout = layoutDocument(documentWith([paragraph('before', 2), first, next]), measurer)
  const last = layout.lines.filter(line => line.target.blockId === 'first').at(-1), following = layout.lines.find(line => line.target.blockId === 'next')
  assert.equal(last.page, following.page); assert.equal(last.column, following.column)
})

test('Enter in a managed quote updates both last-paragraph rules in the same model transaction', () => {
  const child = paragraph('inside'), block = quote('quote', [child]), model = documentWith([block, paragraph('tail')]); initializeDocumentQuotePagination(block)
  executeDocumentCommand(model, selection('inside', runText(child.runs).length), { id: 'doc.command.insert-text', text: '\nnew' })
  assert.equal(model.blocks[0].children.length, 2)
  assert.deepEqual(model.blocks[0].children.map(paragraph => paragraph.style.keepNext), [true, false])
  assert.equal(model.blocks[1].style.keepLines, undefined)
})

test('mixed paragraph formatting is atomic when one target has locked content', () => {
  const first = paragraph('first'), last = paragraph('last'), model = documentWith([first, last])
  model.contentControls = [{ id: 'locked', type: 'text', blockId: 'last', startOffset: 0, endOffset: 6, lock: 'contentLocked', placeholder: '', showingPlaceholder: false }]
  const before = copy(model), range = { anchor: { blockId: 'first', offset: 0 }, focus: { blockId: 'last', offset: 6 } }
  assert.throws(() => executeDocumentCommand(model, range, { id: 'doc.command.format-paragraph', patch: { keepLines: true } }), /locked/)
  assert.deepEqual(model, before)
})
