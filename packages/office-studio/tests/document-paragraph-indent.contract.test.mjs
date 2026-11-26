import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { canChangeDocumentParagraphIndent, executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { applyNamedParagraphStyle, createParagraph, createRun, sliceRuns } from '../src/features/document/model/document-model.ts'
import { numberingDefinitionFromParts } from '../src/features/document/model/document-numbering.ts'
import { officeNumberingMarkerBaseline } from '../src/features/document/model/document-office-numbering-marker.ts'
import { documentStoryParagraphs } from '../src/features/document/model/document-story.ts'
import { createDocumentContentControl } from '../src/features/document/services/content-controls.ts'
import { documentBlockHostSelectionFromLegacyRange, documentModelForBlockHost } from '../src/features/document/services/document-block-host.ts'
import { EMPTY_SECTION_HEADER_FOOTER } from '../src/features/document/services/section-header-footer.ts'

const modelWith = blocks => ({ ...createBlankDocument(), id: 'indent-doc', blocks })
const selection = (first, last = first) => ({ anchor: { blockId: first.id, offset: 0 }, focus: { blockId: last.id, offset: last.runs.map(run => run.text).join('').length } })
const indent = direction => ({ id: 'doc.command.change-paragraph-indent', direction })

test('indent changes each selected paragraph from its own value and respects both bounds', () => {
  const blocks = [0, 40, 100].map(value => { const block = createParagraph(`Indent ${value}`); block.style.indentFirstLine = value; return block })
  const model = modelWith(blocks)
  const ids = blocks.map(block => block.runs[0].id)
  const range = selection(blocks[0], blocks[2])
  executeDocumentCommand(model, range, indent(1))
  assert.deepEqual(blocks.map(block => block.style.indentFirstLine), [20, 60, 100])
  executeDocumentCommand(model, { anchor: range.focus, focus: range.anchor }, indent(-1))
  assert.deepEqual(blocks.map(block => block.style.indentFirstLine), [0, 40, 80])
  assert.deepEqual(blocks.map(block => block.runs[0].id), ids)
  blocks.forEach(block => { block.style.indentFirstLine = 0 })
  const before = JSON.stringify(model)
  assert.equal(canChangeDocumentParagraphIndent(model, range, -1), false)
  executeDocumentCommand(model, range, indent(-1))
  assert.equal(JSON.stringify(model), before)
})

test('capability and execution check only the paragraphs that actually change', () => {
  const first = createParagraph('Editable'), last = createParagraph('Locked at maximum')
  last.style.indentFirstLine = 100
  const model = modelWith([first, last])
  model.permissionRules = [{ id: 'locked', scope: 'paragraph', targetId: last.id, action: 'edit', effect: 'deny', userIds: ['local-user'] }]
  const context = { documentId: model.id, userId: 'local-user' }
  const range = selection(first, last)
  assert.equal(canChangeDocumentParagraphIndent(model, range, 1, context), true)
  executeDocumentCommand(model, range, indent(1), context)
  assert.equal(first.style.indentFirstLine, 20)
  assert.equal(last.style.indentFirstLine, 100)
  const before = JSON.stringify(model)
  assert.equal(canChangeDocumentParagraphIndent(model, range, -1, context), false)
  assert.throws(() => executeDocumentCommand(model, range, indent(-1), context))
  assert.equal(JSON.stringify(model), before)
})

test('indent rejects locked paragraph content outside a caret or partial text selection', () => {
  for (const direction of [-1, 1]) {
    const block = createParagraph('Open locked')
    block.style.indentFirstLine = 20
    const model = modelWith([block])
    createDocumentContentControl(model, { type: 'text', blockId: block.id, startOffset: 5, endOffset: 11, lock: 'contentLocked' })
    const caret = { anchor: { blockId: block.id, offset: 2 }, focus: { blockId: block.id, offset: 2 } }
    const partial = { anchor: { blockId: block.id, offset: 0 }, focus: { blockId: block.id, offset: 4 } }
    const before = JSON.stringify(model)
    for (const range of [caret, partial, { anchor: partial.focus, focus: partial.anchor }]) {
      assert.equal(canChangeDocumentParagraphIndent(model, range, direction), false)
      assert.throws(() => executeDocumentCommand(model, range, indent(direction)), /locked/)
      assert.equal(JSON.stringify(model), before)
    }
  }
})

test('a later paragraph lock rejects indent before changing earlier list depth or imported marker caches', () => {
  const first = { ...createParagraph('01. Editable'), type: 'numbered', level: 0, list: { id: 'imported-list', preset: 'ORDER_LIST' } }
  first.metadata.officeNumberingMarker = {
    version: 1, bookmark: 'office_list_v1_indent', bookmarkId: 99, text: '01. ',
    xml: '<w:bookmarkStart w:id="99" w:name="office_list_v1_indent"/><w:r><w:t>01. </w:t></w:r><w:bookmarkEnd w:id="99"/>',
    baseline: officeNumberingMarkerBaseline(first, '01. '),
  }
  const last = createParagraph('Open locked'), model = modelWith([first, last])
  createDocumentContentControl(model, { type: 'text', blockId: last.id, startOffset: 5, endOffset: 11, lock: 'contentLocked' })
  const forward = { anchor: { blockId: first.id, offset: 1 }, focus: { blockId: last.id, offset: 4 } }
  const before = JSON.stringify(model)
  for (const range of [forward, { anchor: forward.focus, focus: forward.anchor }]) {
    assert.equal(canChangeDocumentParagraphIndent(model, range, 1), false)
    assert.throws(() => executeDocumentCommand(model, range, indent(1)), /locked/)
    assert.equal(JSON.stringify(model), before)
  }
})

test('a content-locked paragraph at the indent limit does not block another paragraph that changes', () => {
  const first = createParagraph('Locked at maximum'), last = createParagraph('Editable')
  first.style.indentFirstLine = 100
  const model = modelWith([first, last])
  createDocumentContentControl(model, { type: 'text', blockId: first.id, startOffset: 0, endOffset: first.runs[0].text.length, lock: 'contentLocked' })
  const range = selection(first, last), untouched = JSON.stringify(first)
  assert.equal(canChangeDocumentParagraphIndent(model, range, 1), true)
  executeDocumentCommand(model, range, indent(1))
  assert.equal(last.style.indentFirstLine, 20)
  assert.equal(JSON.stringify(first), untouched)
  const before = JSON.stringify(model)
  assert.equal(canChangeDocumentParagraphIndent(model, range, -1), false)
  assert.throws(() => executeDocumentCommand(model, range, indent(-1)), /locked/)
  assert.equal(JSON.stringify(model), before)
})

test('half-open paragraph ranges keep a locked final boundary outside indent changes', () => {
  for (const reversed of [false, true]) {
    const first = createParagraph('Selected'), last = createParagraph('Locked'), model = modelWith([first, last])
    createDocumentContentControl(model, { type: 'text', blockId: last.id, startOffset: 0, endOffset: 6, lock: 'contentLocked' })
    const forward = { anchor: { blockId: first.id, offset: 1 }, focus: { blockId: last.id, offset: 0 } }
    const range = reversed ? { anchor: forward.focus, focus: forward.anchor } : forward
    const untouched = JSON.stringify(last)
    assert.equal(canChangeDocumentParagraphIndent(model, range, 1), true)
    executeDocumentCommand(model, range, indent(1))
    assert.equal(first.style.indentFirstLine, 20)
    assert.equal(JSON.stringify(last), untouched)
  }
})

test('story host indentation shares whole-paragraph locks and keeps body and other story paragraphs unchanged', () => {
  const body = createParagraph('Body unchanged'), text = 'Open locked\nTail', model = modelWith([body])
  body.metadata.sectionId = 'first'
  model.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: text, footer: 'Other story', pageFields: [],
    storyContent: { header: { runs: [createRun(text, { italic: true })], paragraphs: documentStoryParagraphs(text, body.style),
      locks: [{ id: 'locked-header-tail', startOffset: 5, endOffset: 11, lock: 'contentLocked' }] } } } }
  const host = { kind: 'story', sectionId: 'first', story: 'header' }
  const range = documentBlockHostSelectionFromLegacyRange(model, 'first', 'header', 1, 4)
  const before = JSON.stringify(model)
  assert.equal(canChangeDocumentParagraphIndent(documentModelForBlockHost(model, host), range, 1), false)
  assert.throws(() => executeDocumentCommand(model, range, { ...indent(1), host }), /locked/)
  assert.equal(JSON.stringify(model), before)
  delete model.sectionHeaderFooters.first.storyContent.header.locks
  const scoped = documentModelForBlockHost(model, host), untouched = JSON.stringify(scoped.blocks[1]), bodyBefore = JSON.stringify(model.blocks)
  assert.equal(canChangeDocumentParagraphIndent(scoped, range, 1), true)
  executeDocumentCommand(model, range, { ...indent(1), host })
  const after = documentModelForBlockHost(model, host)
  assert.equal(after.blocks[0].style.indentFirstLine, 20)
  assert.equal(JSON.stringify(after.blocks[1]), untouched)
  assert.equal(JSON.stringify(model.blocks), bodyBefore)
  assert.equal(model.sectionHeaderFooters.first.footer, 'Other story')
})

test('list level changes retain signed custom offsets relative to their level definitions', () => {
  const numbering = numberingDefinitionFromParts({ abstractNumId: 7, levels: [
    { ilvl: 0, format: 'decimal', text: '%1.', indent: { leftTwips: 420, rightTwips: 45, hangingTwips: 300 } },
    { ilvl: 1, format: 'decimal', text: '%2.', indent: { leftTwips: 840, rightTwips: 60, hangingTwips: 600 } },
  ] }, { numId: 9, abstractNumId: 7 })
  const block = { ...createParagraph('Item'), type: 'numbered', metadata: { numberingDefinition: numbering, numberingListId: 'list' } }
  Object.assign(block.style, { indentStart: 35, indentEnd: -2, indentFirstLine: -5, hanging: 22 })
  const model = modelWith([block]), range = selection(block)
  const before = JSON.stringify(model)
  executeDocumentCommand(model, range, indent(1))
  assert.equal(block.level, 1)
  assert.equal(block.style.indentStart, 63)
  assert.equal(block.style.indentEnd, -1)
  assert.equal(block.style.indentFirstLine, -5)
  assert.equal(block.style.hanging, 42)
  assert.equal(canChangeDocumentParagraphIndent(model, range, 1), false)
  executeDocumentCommand(model, range, indent(-1))
  assert.equal(JSON.stringify(model), before)
  delete block.style.indentStart
  block.style.indent = 1
  executeDocumentCommand(model, range, indent(1))
  assert.equal(block.style.indentStart, 80)
})

test('numbered heading indentation changes list depth while keeping its outline depth', () => {
  const block = { ...createParagraph('Heading'), type: 'heading', level: 2, list: { id: 'headings', preset: 'ORDER_LIST' }, metadata: { numberingLevel: 0 } }
  const model = modelWith([block])
  executeDocumentCommand(model, selection(block), indent(1))
  assert.equal(block.level, 2)
  assert.equal(block.metadata.numberingLevel, 1)
  assert.equal(block.style.indentStart, 56)
  applyNamedParagraphStyle(block, 'heading3')
  assert.equal(block.level, 3)
  assert.equal(block.metadata.numberingLevel, 1)
  block.style.spacingRule = 'exact'
  block.style.lineHeight = 24
  applyNamedParagraphStyle(block, 'normal')
  assert.equal(block.style.spacingRule, 'auto')
  assert.equal(block.style.lineHeight, 1.5)
})

test('paragraph formatting validates every target before committing any style', () => {
  const exact = createParagraph('Exact'), automatic = createParagraph('Automatic')
  Object.assign(exact.style, { spacingRule: 'exact', lineHeight: 24 })
  const model = modelWith([exact, automatic]), before = JSON.stringify(model)
  assert.throws(() => executeDocumentCommand(model, selection(exact, automatic), { id: 'doc.command.format-paragraph', patch: { lineHeight: 6 } }))
  assert.equal(JSON.stringify(model), before)
})

test('empty input and unchanged formatting preserve persistent arrays and run IDs', () => {
  const block = createParagraph('Existing text', { bold: true })
  const model = modelWith([block]), runs = block.runs
  const caret = { anchor: { blockId: block.id, offset: 3 }, focus: { blockId: block.id, offset: 3 } }
  const before = JSON.stringify(model)
  executeDocumentCommand(model, caret, { id: 'doc.command.insert-text', text: '' })
  executeDocumentCommand(model, selection(block), { id: 'doc.command.format-text', patch: { bold: true } })
  assert.equal(block.runs, runs)
  assert.equal(JSON.stringify(model), before)
  assert.equal(sliceRuns(runs, 0, runs[0].text.length)[0].id, runs[0].id)
  assert.notEqual(sliceRuns(runs, 1, 3)[0].id, runs[0].id)
})

test('local text edits keep complete untouched runs stable', () => {
  const block = createParagraph('First')
  block.runs.push({ ...block.runs[0], id: 'untouched-run', text: 'Second', bold: true })
  const model = modelWith([block])
  executeDocumentCommand(model, { anchor: { blockId: block.id, offset: 2 }, focus: { blockId: block.id, offset: 2 } }, { id: 'doc.command.insert-text', text: 'x' })
  assert.equal(block.runs.at(-1).id, 'untouched-run')
  assert.equal(block.runs.at(-1).text, 'Second')
})
