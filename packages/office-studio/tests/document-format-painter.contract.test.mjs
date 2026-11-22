import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection, createParagraph, createRun } from '../src/features/document/model/document-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { captureDocumentFormat, clearDocumentFormat, documentPainterTarget } from '../src/features/document/services/document-format-painter.ts'

const pt = (block, offset) => ({ blockId: block.id, offset })
const apply = (model, selection, format) => executeDocumentCommand(model, selection, { id: 'doc.command.apply-format-painter', format })

test('a protected destination rejects word painting before changing any text or paragraph', () => {
  const model = createBlankDocument(), source = createParagraph('sample', { bold: true }), target = createParagraph('protected')
  model.blocks = [source, target]
  model.permissionRules = [{ id: 'deny', scope: 'paragraph', targetId: target.id, action: 'edit', effect: 'deny', userIds: [] }]
  const before = structuredClone(model)
  const format = captureDocumentFormat(model, collapsedSelection(pt(source, 1)))
  assert.throws(() => apply(model, collapsedSelection(pt(target, 2)), format))
  assert.deepEqual(model, before)
})

test('reversed selection samples its first character, caret samples its left neighbor, target links survive', () => {
  const model = createBlankDocument(), from = createParagraph(), to = createParagraph('next word', { italic: true, href: 'https://example.org' })
  from.runs = [createRun('A', { bold: true }), createRun('BC', { italic: true, verticalAlign: 'superscript', href: 'https://example.com' })]
  model.blocks = [from, to]
  const format = captureDocumentFormat(model, { anchor: pt(from, 3), focus: pt(from, 1) })
  assert.equal(format.text.italic, true)
  assert.equal(format.paragraph, undefined)
  assert.equal(captureDocumentFormat(model, collapsedSelection(pt(from, 1))).text.bold, true)
  const selected = apply(model, collapsedSelection(pt(to, 7)), format)
  assert.deepEqual([selected.anchor.offset, selected.focus.offset], [5, 9])
  assert.deepEqual(to.runs.map(r => [r.text, r.verticalAlign, r.href]), [['next ', undefined, 'https://example.org'], ['word', 'superscript', 'https://example.org']])
  const empty = createParagraph('  ')
  model.blocks.push(empty)
  assert.equal(documentPainterTarget(model, collapsedSelection(pt(empty, 1))), undefined)
})

test('paragraph mark copies list layout to affected paragraphs, excluding the next paragraph at offset zero', () => {
  const model = createBlankDocument(), from = createParagraph('sample'), to = createParagraph('one'), last = createParagraph('two')
  from.type = 'numbered'; from.level = 2; from.style.indent = 3
  from.list = { id: 'sequence', preset: 'ORDER_LIST_2' }
  to.metadata.bookmark = 'destination'
  model.blocks = [from, to, last]
  const format = captureDocumentFormat(model, { anchor: pt(from, 0), focus: pt(to, 0) })
  apply(model, { anchor: pt(to, 0), focus: pt(last, 0) }, format)
  assert.deepEqual([to.type, to.level, to.style.indent, to.metadata.bookmark, last.type], ['numbered', 2, 3, 'destination', 'paragraph'])
  assert.deepEqual(to.list, from.list)
  assert.notEqual(to.list, from.list)
  apply(model, collapsedSelection(pt(to, 3)), clearDocumentFormat())
  assert.equal(to.type, 'paragraph'); assert.equal(to.style.indent, 0); assert.equal(to.runs[0].text, 'one')
  assert.equal(to.list, undefined)
})
