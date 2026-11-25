import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Contract drafts for unified verification. This implementation batch has not run them.
let server, domain, model, fields, bounds, boundaryEdits, affinity, controlBounds, controls, editing, projection, notes
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-note-boundaries', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  model = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  fields = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-fields.ts')
  bounds = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-note-boundary.ts')
  boundaryEdits = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-note-boundary-edit.ts')
  affinity = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-note-insertion-affinity.ts')
  controlBounds = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/content-control-note-boundary.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  editing = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-control-editing.ts')
  projection = await server.ssrLoadModule('/packages/office-studio/src/parser/documentTextProjection.ts')
  notes = await server.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-notes.ts')
})
afterAll(async () => { await server?.close() })

const copy = value => structuredClone(value)
const reference = (id, text = '', kind = 'footnote') => ({ ...model.createRun(text), id,
  metadata: { note: { type: kind, noteId: id }, noteReference: { kind, noteId: id } } })
function fixture() {
  const block = { ...model.createParagraph(), id: 'p', runs: [model.createRun('A'), reference('visible', '1'),
    reference('empty-first'), reference('empty-second'), model.createRun('B')] }
  return { value: { ...domain.createBlankDocument(), blocks: [block], contentControls: [] }, block }
}
const field = (id, startOffset, endOffset, startNoteReferenceCount, endNoteReferenceCount, parent = null) => ({
  id, field: 'REF', instruction: `REF ${id}`, cachedResult: 'retained cache', startOffset, endOffset,
  startNoteReferenceCount, endNoteReferenceCount, metadata: { source: { parentFieldId: parent } },
})
const fieldRange = value => [value.startOffset, value.startNoteReferenceCount ?? 0, value.endOffset, value.endNoteReferenceCount ?? 0]
const add = (value, id, type, startOffset, endOffset, startNoteReferenceCount, endNoteReferenceCount, extra = {}) =>
  controls.createDocumentContentControl(value, { id, type, blockId: 'p', startOffset, endOffset, startNoteReferenceCount, endNoteReferenceCount, ...extra })
function attachNotes(value) {
  const raw = Object.fromEntries(value.blocks[0].runs.filter(run => run.metadata?.note).map(run => [run.id, { id: run.id,
    referenceBlockId: 'p', referenceRunId: run.id, customMark: run.text, text: `Body ${run.id}`,
    content: { version: 1, runs: [], paragraphs: [], blocks: [{ ...model.createParagraph(`Body ${run.id}`), id: `body-${run.id}` }] } }]))
  value.footnotes = notes.normalizeDocumentNotes({ footnotes: raw }, value.blocks).footnotes
}

test('cumulative ranks include visible references and only empty atoms create tied positions', () => {
  const { block } = fixture(), refs = bounds.documentNoteReferenceOffsets(block.runs)
  assert.deepEqual(bounds.documentNoteReferenceBoundaryRanks(refs, 2), { first: 1, last: 3 })
  assert.deepEqual([undefined, 1, 2, 3].map(rank => bounds.assertDocumentNoteReferenceBoundaryRank(refs, 2, rank)), [1, 1, 2, 3])
  for (const rank of [0, 4, -1, 1.5, null, '1', NaN, Number.MAX_SAFE_INTEGER + 1]) {
    assert.throws(() => bounds.assertDocumentNoteReferenceBoundaryRank(refs, 2, rank), /序号/)
  }
  const atomic = [model.createRun('A'), reference('multi', '12'), reference('custom', '😀*', 'endnote')]
  const atomicRefs = bounds.documentNoteReferenceOffsets(atomic)
  for (const offset of [2, 4, 5]) assert.throws(() => bounds.assertDocumentNoteReferenceBoundaryRank(atomicRefs, offset), /字符内部/)
  assert.equal(bounds.assertDocumentNoteReferenceBoundaryRank(atomicRefs, 6), 2)
})

test('native partition selects one empty reference and leaves each unselected run exactly once', () => {
  const { block } = fixture()
  const result = controlBounds.partitionDocumentContentControlNoteRuns(block.runs,
    { startOffset: 2, endOffset: 2, startNoteReferenceCount: 2, endNoteReferenceCount: 3 })
  assert.deepEqual(result.selected.map(run => run.id), ['empty-second'])
  assert.deepEqual(result.before.map(run => run.id), block.runs.slice(0, 3).map(run => run.id))
  assert.deepEqual(result.after.map(run => run.id), block.runs.slice(4).map(run => run.id))
  assert.equal(new Set([...result.before, ...result.selected, ...result.after].map(run => run.id)).size, block.runs.length)
  assert.throws(() => controlBounds.partitionDocumentContentControlNoteRuns([reference('atomic', '12')],
    { startOffset: 0, endOffset: 1 }), /字符内部/)
  const text = model.createRun('ABCD')
  const clipped = controlBounds.partitionDocumentContentControlNoteRuns([text], { startOffset: 1, endOffset: 3 })
  assert.deepEqual([clipped.before, clipped.selected, clipped.after].map(model.runText), ['A', 'BC', 'D'])
  assert.equal(new Set([...clipped.before, ...clipped.selected, ...clipped.after].map(run => run.id)).size, 3)
})

test('plain text projections preserve default and explicit native insertion and replacement sides', () => {
  const { block } = fixture(), refs = bounds.documentNoteReferenceOffsets(block.runs)
  const project = (start, end, length, ranks) => boundaryEdits.documentNoteReferenceOffsetsAfterTextEdit(refs, start, end, length, ranks)
  assert.deepEqual(project(2, 2, 1), [refs[0], { ...refs[1], startOffset: 3, endOffset: 3 }, { ...refs[2], startOffset: 3, endOffset: 3 }])
  for (const rank of [1, 2, 3]) assert.deepEqual(project(2, 2, 1, { start: rank, end: rank }), refs.map((reference, index) =>
    index < rank ? reference : { ...reference, startOffset: reference.startOffset + 1, endOffset: reference.endOffset + 1 }))
  assert.deepEqual(project(2, 2, 0, { start: 1, end: 2 }), [refs[0], refs[2]])
  assert.deepEqual(project(2, 2, 2, { start: 1, end: 2 }), [refs[0], { ...refs[2], startOffset: 4, endOffset: 4 }])
  assert.deepEqual(project(2, 3, 2, { start: 3, end: 3 }), refs)
  assert.deepEqual(project(1, 2, 1, { start: 0, end: 1 }), refs.slice(1))
  for (const ranks of [{ start: 3, end: 1 }, { start: 0, end: 1 }, { start: 1.5, end: 2 }, null]) {
    assert.throws(() => project(2, 2, 0, ranks))
  }
  assert.throws(() => boundaryEdits.documentNoteReferenceOffsetsAfterTextEdit([...refs, refs[2]], 2, 2, 0), /源几何或身份/)
  assert.throws(() => project(2, 2, Number.MAX_SAFE_INTEGER), /参数/)
  const atomic = [{ runId: 'wide', startOffset: 1, endOffset: 3 }]
  assert.throws(() => boundaryEdits.documentNoteReferenceOffsetsAfterTextEdit(atomic, 2, 3, 0), /字符内部/)
})

test('explicit zero-display native deletion works without a separate future-reference projection', () => {
  const { value, block } = fixture()
  add(value, 'remaining', 'richText', 2, 2, 2, 3)
  controls.transformContentControlsForBlockEdit(value, 'p', 2, 2, 0, undefined, true,
    { sourceNoteRanks: { start: 1, end: 2 } })
  block.runs = block.runs.filter(run => run.id !== 'empty-first')
  assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'remaining')), [2, 1, 2, 2])
  controls.validateDocumentContentControlRanges(value)
})

test('plain and empty-note insertion resolve the actual source boundary; renumbering retains ID affinity', () => {
  const { block } = fixture(), refs = bounds.documentNoteReferenceOffsets(block.runs)
  const plain = [...block.runs.slice(0, 2), model.createRun('x'), ...block.runs.slice(2)]
  const point = affinity.documentNoteReferenceInsertionAffinity(refs,
    [{ id: 'p', text: model.runText(plain), noteReferences: bounds.documentNoteReferenceOffsets(plain) }], 2, 1)
  assert.equal(point.sourceRank, 1)
  assert.deepEqual([point.before.offset, point.before.rank, point.after.offset, point.after.rank], [2, 1, 3, 1])
  const empty = [...block.runs.slice(0, 3), reference('new-empty'), ...block.runs.slice(3)]
  const native = affinity.documentNoteReferenceInsertionAffinity(refs,
    [{ id: 'p', text: model.runText(empty), noteReferences: bounds.documentNoteReferenceOffsets(empty) }], 2, 0)
  assert.equal(native.sourceRank, 2)
  assert.deepEqual([native.before.offset, native.before.rank, native.after.offset, native.after.rank], [2, 2, 2, 3])
  const numbered = block.runs.map(run => run.id === 'empty-first' ? { ...run, text: '10' } : run)
  const destination = bounds.documentNoteReferenceOffsets(numbered)
  assert.equal(affinity.documentNoteReferenceInsertionAffinity(refs,
    [{ id: 'p', text: model.runText(numbered), noteReferences: destination }], 2, 2), undefined)
  assert.deepEqual(bounds.remapDocumentNoteReferenceBoundary(refs, destination, 2, 2, 2), { offset: 4, rank: 2 })
  for (const sourceRank of [1, 2, 3]) {
    const index = sourceRank + 1
    const future = [...block.runs.slice(0, index), model.createRun('x'), ...block.runs.slice(index)]
    const mapped = affinity.documentNoteReferenceInsertionAffinity(refs,
      [{ id: 'p', text: model.runText(future), noteReferences: bounds.documentNoteReferenceOffsets(future) }], 2, 1)
    assert.equal(mapped.sourceRank, sourceRank)
    assert.deepEqual([mapped.before.offset, mapped.before.rank, mapped.after.offset, mapped.after.rank], [2, sourceRank, 3, sourceRank])
  }
})

test('empty replacement rejects invalid parameters, source identities and unrelated removed atoms', () => {
  const { block } = fixture(), refs = bounds.documentNoteReferenceOffsets(block.runs)
  const future = block.runs.filter(run => run.id !== 'empty-first')
  const outputs = [{ id: 'p', text: model.runText(future), noteReferences: bounds.documentNoteReferenceOffsets(future) }]
  const valid = affinity.documentNoteReferenceEmptyReplacementAffinity(refs, outputs, 2, 0, { startRank: 1, endRank: 2 })
  assert.deepEqual([valid.before.offset, valid.before.rank, valid.after.offset, valid.after.rank], [2, 1, 2, 1])
  for (const range of [{ startRank: 1, endRank: 1 }, { startRank: 2, endRank: 1 }, { startRank: -1, endRank: 2 },
    { startRank: 0, endRank: 4 }, { startRank: 1.5, endRank: 2 }, { startRank: '1', endRank: 2 }, null]) {
    assert.throws(() => affinity.documentNoteReferenceEmptyReplacementAffinity(refs, outputs, 2, 0, range))
  }
  for (const [offset, length] of [[-1, 0], [2, -1], [2, 0.5], [NaN, 0], [2, Number.MAX_SAFE_INTEGER + 1]]) {
    assert.throws(() => affinity.documentNoteReferenceEmptyReplacementAffinity(refs, outputs, offset, length, { startRank: 1, endRank: 2 }))
  }
  assert.throws(() => affinity.documentNoteReferenceEmptyReplacementAffinity(refs, [], 2, 0, { startRank: 1, endRank: 2 }))
  assert.throws(() => affinity.documentNoteReferenceEmptyReplacementAffinity(refs, outputs, 1, 0, { startRank: 0, endRank: 1 }))
  assert.throws(() => affinity.documentNoteReferenceEmptyReplacementAffinity([...refs, refs[2]], outputs, 2, 0, { startRank: 1, endRank: 2 }))
})

test('field ends exclude inserted text while starts and independent points follow it', () => {
  const { value, block } = fixture()
  fields.setDocumentTextFields(block, [field('left', 0, 2, 0, 1), field('right', 2, 3, 1, 3), field('point', 2, 2, 1, 1)])
  const future = [...block.runs.slice(0, 2), model.createRun('x'), ...block.runs.slice(2)]
  fields.transformDocumentTextFieldsForBlockEdit(block, 2, 2, 'x', undefined, value, bounds.documentNoteReferenceOffsets(future))
  block.runs = future
  const mapped = new Map(fields.documentTextFields(block).map(field => [field.id, field]))
  assert.deepEqual(fieldRange(mapped.get('left')), [0, 0, 2, 1])
  assert.deepEqual(fieldRange(mapped.get('right')), [3, 1, 4, 3])
  assert.deepEqual(fieldRange(mapped.get('point')), [3, 1, 3, 1])
  assert.equal(mapped.get('left').cachedResult, 'retained cache')
  fields.documentTextFieldRanges(value)
})

test('explicit point children follow their parent end, including an explicitly preserved result', () => {
  for (const preserve of [undefined, 'parent']) {
    const { value, block } = fixture()
    fields.setDocumentTextFields(block, [field('parent', 0, 2, 0, 1), field('child', 2, 2, 1, 1, 'parent')])
    const future = [...block.runs.slice(0, 2), model.createRun('x'), ...block.runs.slice(2)]
    fields.transformDocumentTextFieldsForBlockEdit(block, 2, 2, 'x', preserve, value, bounds.documentNoteReferenceOffsets(future))
    block.runs = future
    const mapped = fields.documentTextFields(block)
    assert.equal(mapped.find(field => field.id === 'parent').endOffset, preserve ? 3 : 2)
    assert.deepEqual(fieldRange(mapped.find(field => field.id === 'child')), preserve ? [3, 1, 3, 1] : [2, 1, 2, 1])
    fields.documentTextFieldRanges(value)
  }
})

test('a native deletion preserves the exact field definition and removes points strictly inside deleted atoms', () => {
  const { value, block } = fixture()
  fields.setDocumentTextFields(block, [field('result', 2, 2, 1, 3), field('start', 2, 2, 1, 1, 'result'),
    field('inside', 2, 2, 2, 2, 'result'), field('end', 2, 2, 3, 3, 'result')])
  const future = block.runs.filter(run => !run.id.startsWith('empty-'))
  fields.transformDocumentTextFieldsForBlockEdit(block, 2, 2, '', undefined, value, bounds.documentNoteReferenceOffsets(future))
  block.runs = future
  const mapped = fields.documentTextFields(block)
  assert.deepEqual(mapped.map(field => field.id).sort(), ['end', 'result', 'start'])
  for (const result of mapped) assert.deepEqual(fieldRange(result), [2, 1, 2, 1])
  fields.documentTextFieldRanges(value)
})

test('split and merge recompute local ranks after locating surviving note run IDs across all destination paragraphs', () => {
  const { value, block } = fixture()
  fields.setDocumentTextFields(block, [field('left', 0, 2, 0, 1), field('right', 2, 3, 1, 3), field('point', 2, 2, 1, 1)])
  const first = { ...copy(block), runs: copy(block.runs.slice(0, 2)) }
  const last = { ...model.createParagraph(), id: 'last', runs: copy(block.runs.slice(2)) }
  fields.transformDocumentTextFieldsForBlockRangeEdit([block], 2, 2, '\n', [first, last], undefined, value,
    [bounds.documentNoteReferenceOffsets(first.runs), bounds.documentNoteReferenceOffsets(last.runs)])
  assert.deepEqual(fieldRange(fields.documentTextFields(first)[0]), [0, 0, 2, 1])
  const lastFields = new Map(fields.documentTextFields(last).map(field => [field.id, field]))
  assert.deepEqual(fieldRange(lastFields.get('right')), [0, 0, 1, 2])
  assert.deepEqual(fieldRange(lastFields.get('point')), [0, 0, 0, 0])
  value.blocks = [first, last]
  fields.documentTextFieldRanges(value)
  const merged = { ...copy(first), runs: [...copy(first.runs), ...copy(last.runs)] }
  fields.transformDocumentTextFieldsForBlockRangeEdit([first, last], 2, 0, '', [merged], undefined, value,
    [bounds.documentNoteReferenceOffsets(merged.runs)])
  value.blocks = [merged]
  const restored = new Map(fields.documentTextFields(merged).map(field => [field.id, field]))
  assert.deepEqual(fieldRange(restored.get('right')), [2, 1, 3, 3])
  assert.deepEqual(fieldRange(restored.get('left')), [0, 0, 2, 1])
  fields.documentTextFieldRanges(value)
})

test('SDT ordinary insertion uses half-open native boundaries and preserves point child ownership', () => {
  const { value, block } = fixture()
  add(value, 'left', 'group', 0, 2, 0, 1)
  add(value, 'child', 'richText', 2, 2, 1, 1, { parentId: 'left' })
  add(value, 'right', 'richText', 2, 3, 1, 3)
  const future = [...block.runs.slice(0, 2), model.createRun('x'), ...block.runs.slice(2)]
  controls.transformContentControlsForBlockEdit(value, 'p', 2, 2, 1, undefined, true,
    { projectedNoteReferences: bounds.documentNoteReferenceOffsets(future) })
  block.runs = future
  const mapped = new Map(value.contentControls.map(control => [control.id, control]))
  assert.deepEqual(fieldRange(mapped.get('left')), [0, 0, 2, 1])
  assert.deepEqual(fieldRange(mapped.get('child')), [2, 1, 2, 1])
  assert.deepEqual(fieldRange(mapped.get('right')), [3, 1, 4, 3])
  controls.validateDocumentContentControlRanges(value)
})

test('body insertion maps all three real destination paragraphs with local native ranks', () => {
  const { value, block } = fixture()
  add(value, 'left', 'group', 0, 2, 0, 1)
  add(value, 'child', 'richText', 2, 2, 1, 1, { parentId: 'left' })
  add(value, 'right', 'richText', 2, 3, 1, 3)
  add(value, 'point', 'richText', 2, 2, 1, 1)
  const source = fields.documentTextFieldParagraphs(value)
  const first = { ...copy(block), runs: copy(block.runs.slice(0, 2)) }
  const middle = { ...model.createParagraph('mid'), id: 'middle' }
  const last = { ...model.createParagraph(), id: 'last', runs: copy(block.runs.slice(2)) }
  value.blocks = [first, middle, last]
  controls.transformContentControlsForBodyRangeEdit(value, source, 2, 2, 5, true)
  const mapped = new Map(value.contentControls.map(control => [control.id, control]))
  assert.deepEqual([mapped.get('left').blockId, ...fieldRange(mapped.get('left'))], ['p', 0, 0, 2, 1])
  assert.deepEqual([mapped.get('child').blockId, ...fieldRange(mapped.get('child'))], ['p', 2, 1, 2, 1])
  assert.deepEqual([mapped.get('right').blockId, ...fieldRange(mapped.get('right'))], ['last', 0, 0, 1, 2])
  assert.deepEqual([mapped.get('point').blockId, ...fieldRange(mapped.get('point'))], ['last', 0, 0, 0, 0])
  assert.equal(model.runText(value.blocks[1].runs), 'mid')
  controls.validateDocumentContentControlRanges(value)
})

test('cross-paragraph structural body contexts use the real old model and reject stale snapshots', () => {
  const value = { ...domain.createBlankDocument(), blocks: [{ ...model.createParagraph('AB'), id: 'first' },
    { ...model.createParagraph('CD'), id: 'last' }], contentControls: [] }
  controls.createDocumentContentControl(value, { id: 'outer', type: 'group', blockId: 'first', startOffset: 0, endOffset: 2,
    blockRange: { version: 1, blockIds: ['first', 'last'] }, placement: 'block' })
  const before = copy(value), paragraphs = fields.documentTextFieldParagraphs(before)
  value.blocks = [{ ...copy(value.blocks[0]), runs: [model.createRun('AxxxD')] }]
  const candidate = copy(value)
  assert.throws(() => controls.transformContentControlsForBodyRangeEdit(value, paragraphs, 1, 4, 3, true,
    { activeControlId: 'outer' }), /编辑前模型/)
  assert.deepEqual(value, candidate)
  const stale = copy(before); stale.blocks[0].runs[0].text = 'changed'
  assert.throws(() => controls.transformContentControlsForBodyRangeEdit(value, paragraphs, 1, 4, 3, true,
    { activeControlId: 'outer', sourceModel: stale }), /源模型与捕获段落/)
  assert.deepEqual(value, candidate)
  controls.transformContentControlsForBodyRangeEdit(value, paragraphs, 1, 4, 3, true,
    { activeControlId: 'outer', sourceModel: before })
  assert.equal(value.contentControls[0].id, 'outer')
  const locked = copy(before); locked.contentControls[0].lock = 'sdtContentLocked'
  const unchanged = copy(candidate)
  assert.throws(() => controls.transformContentControlsForBodyRangeEdit(unchanged, fields.documentTextFieldParagraphs(locked), 1, 4, 3, true,
    { activeControlId: 'outer', sourceModel: locked }), /content is locked/)
  assert.deepEqual(unchanged, candidate)
  // The enclosing block planner owns replacement of the structural blockIds.
})

test('explicit nonempty scalar replacement keeps tied notes and locked fields outside the native selection', () => {
  const { value, block } = fixture()
  attachNotes(value)
  add(value, 'input', 'richText', 2, 3, 3, 3)
  fields.setDocumentTextFields(block, [field('first-note-field', 2, 2, 1, 2), field('second-note-field', 2, 2, 2, 3)])
  block.metadata.documentTextFields.fields[0].locked = true
  controls.setDocumentContentControlValue(value, 'input', 'XX')
  const live = value.blocks[0]
  assert.equal(model.runText(live.runs), 'A1XX')
  assert.deepEqual(bounds.documentNoteReferenceOffsets(live.runs).map(reference => reference.runId), ['visible', 'empty-first', 'empty-second'])
  const mapped = new Map(fields.documentTextFields(live).map(field => [field.id, field]))
  assert.deepEqual(fieldRange(mapped.get('first-note-field')), [2, 1, 2, 2])
  assert.deepEqual(fieldRange(mapped.get('second-note-field')), [2, 2, 2, 3])
  assert.equal(mapped.get('first-note-field').locked, true)
  assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'input')), [2, 3, 4, 3])
  assert.equal(Object.keys(value.footnotes).length, 3)
  fields.documentTextFieldRanges(value); controls.validateDocumentContentControlRanges(value)
})

test('setting equal empty or visible text replaces the selected note entity and its scope atomically', () => {
  for (const visible of [false, true]) {
    const { value, block } = fixture()
    attachNotes(value)
    const id = visible ? 'visible' : 'empty-second'
    const start = visible ? 1 : 2, startRank = visible ? 0 : 2, endRank = visible ? 1 : 3
    add(value, 'input', 'richText', start, 2, startRank, endRank, { value: visible ? '1' : '' })
    fields.setDocumentTextFields(block, [field('replacement-result', start, 2, startRank, endRank),
      field('unselected', 2, 2, 1, 2)])
    const before = copy(value)
    controls.setDocumentContentControlValue(value, 'input', visible ? '1' : '')
    const live = value.blocks[0], remaining = bounds.documentNoteReferenceOffsets(live.runs).map(reference => reference.runId)
    assert.equal(model.runText(live.runs), model.runText(before.blocks[0].runs))
    assert.equal(remaining.includes(id), false)
    assert.equal(Object.hasOwn(value.footnotes, id), false)
    assert.equal(Object.keys(value.footnotes).length, 2)
    assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'input')), visible ? [1, 0, 2, 0] : [2, 2, 2, 2])
    const mapped = new Map(fields.documentTextFields(live).map(field => [field.id, field]))
    assert.deepEqual(fieldRange(mapped.get('unselected')), visible ? [2, 0, 2, 1] : [2, 1, 2, 2])
    assert.equal(mapped.get('replacement-result').cachedResult, 'retained cache')
    fields.documentTextFieldRanges(value); controls.validateDocumentContentControlRanges(value)
    const plain = copy(value)
    controls.setDocumentContentControlValue(value, 'input', visible ? '1' : '')
    assert.deepEqual(value, plain)
  }
})

test('equal scalar text does not bypass locks on the selected note story or a selected native field', () => {
  for (const lockStory of [false, true]) {
    const { value, block } = fixture()
    attachNotes(value)
    add(value, 'input', 'richText', 2, 2, 2, 3, { value: '' })
    if (lockStory) value.footnotes['empty-second'].content.locks = [{ id: 'note-lock', startOffset: 0, endOffset: 1, lock: 'sdtContentLocked' }]
    else fields.setDocumentTextFields(block, [{ ...field('locked-note-field', 2, 2, 2, 3), locked: true }])
    const before = copy(value)
    assert.throws(() => controls.setDocumentContentControlValue(value, 'input', ''), /锁定/)
    assert.deepEqual(value, before)
  }
})

test('explicit source native ranks disagreeing with future insertion or deletion reject both range plans', () => {
  const { value, block } = fixture()
  add(value, 'input', 'richText', 2, 2, 1, 3)
  fields.setDocumentTextFields(block, [field('notes', 2, 2, 1, 3)])
  const before = copy(value), future = block.runs.filter(run => run.id !== 'empty-first')
  const refs = bounds.documentNoteReferenceOffsets(future), wrong = { start: 2, end: 3 }
  assert.throws(() => controls.transformContentControlsForBlockEdit(value, 'p', 2, 2, 0, undefined, true,
    { projectedNoteReferences: refs, sourceNoteRanks: wrong }), /源注释引用端点/)
  assert.throws(() => fields.transformDocumentTextFieldsForBlockEdit(block, 2, 2, '', undefined, value, refs, wrong), /源注释引用端点/)
  assert.deepEqual(value, before)
})

test('inserting an empty note inside a point SDT changes content, binding dirtiness and temporary-wrapper history', () => {
  const { value, block } = fixture()
  add(value, 'group', 'group', 2, 2, 1, 1)
  add(value, 'input', 'richText', 2, 2, 1, 1, { parentId: 'group', temporary: true,
    dataBinding: { xpath: '/value', storeItemID: '{test}' } })
  const before = copy(value), future = [...block.runs.slice(0, 2), reference('new-empty'), ...block.runs.slice(2)]
  controls.transformContentControlsForBlockEdit(value, 'p', 2, 2, 0, undefined, true,
    { activeControlId: 'input', projectedNoteReferences: bounds.documentNoteReferenceOffsets(future) })
  block.runs = future
  assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'input')), [2, 1, 2, 2])
  const result = editing.synchronizeDocumentContentControlsAfterEdit(before, value, { origin: 'text-input', activeControlId: 'input',
    selection: { anchor: { blockId: 'p', offset: 2 }, focus: { blockId: 'p', offset: 2 } } })
  assert.ok(result.changedControlIds.includes('input'))
  assert.deepEqual(result.removedTemporaryControlIds, ['input'])
  assert.equal(result.removedTemporaryControls[0].dataBinding.dirty, true)
  assert.equal(value.contentControls.some(control => control.id === 'input'), false)
  assert.equal(model.runText(block.runs), model.runText(before.blocks[0].runs))
})

test('derived empty-to-visible renumbering retains field and locked SDT identity rather than inserting text inside them', () => {
  const { value, block } = fixture()
  fields.setDocumentTextFields(block, [field('first', 2, 2, 1, 2), field('second', 2, 2, 2, 3)])
  add(value, 'first-control', 'richText', 2, 2, 1, 2, { lock: 'sdtContentLocked' })
  add(value, 'second-control', 'richText', 2, 2, 2, 3, { lock: 'sdtContentLocked' })
  const future = block.runs.map(run => run.id === 'empty-first' ? { ...run, text: '10' } : run)
  const refs = bounds.documentNoteReferenceOffsets(future)
  controls.transformContentControlsForBlockEdit(value, 'p', 2, 2, 2, undefined, true,
    { projectedNoteReferences: refs, noteReferenceLabelRunId: 'empty-first' })
  fields.transformDocumentTextFieldsForBlockEdit(block, 2, 2, '10', undefined, value, refs)
  block.runs = future
  assert.deepEqual(fieldRange(fields.documentTextFields(block).find(field => field.id === 'first')), [2, 1, 4, 2])
  assert.deepEqual(fieldRange(fields.documentTextFields(block).find(field => field.id === 'second')), [4, 2, 4, 3])
  assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'first-control')), [2, 1, 4, 2])
  assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'second-control')), [4, 2, 4, 3])
  controls.validateDocumentContentControlRanges(value)
})

test('noncontiguous empty deletions and duplicate destination references fail before publishing field or SDT plans', () => {
  const { value, block } = fixture()
  block.runs.splice(4, 0, reference('empty-third'))
  fields.setDocumentTextFields(block, [field('whole', 0, 3, 0, 4)])
  add(value, 'whole-control', 'richText', 0, 3, 0, 4)
  const before = copy(value), future = block.runs.filter(run => !['empty-first', 'empty-third'].includes(run.id))
  const refs = bounds.documentNoteReferenceOffsets(future)
  assert.throws(() => controls.transformContentControlsForBlockEdit(value, 'p', 2, 2, 0, undefined, true, { projectedNoteReferences: refs }), /连续/)
  assert.throws(() => fields.transformDocumentTextFieldsForBlockEdit(block, 2, 2, '', undefined, value, refs), /连续/)
  assert.deepEqual(value, before)
  const duplicated = [{ ...copy(block), id: 'a' }, { ...copy(block), id: 'b' }]
  assert.throws(() => affinity.assertDocumentNoteReferenceParagraphIds(duplicated.map(block => ({ noteReferences: bounds.documentNoteReferenceOffsets(block.runs) }))), /重复/)
})

test('repeating an item containing one empty note copies only that native atom and reanchors adjacent fields and items', () => {
  const { value, block } = fixture()
  add(value, 'section', 'repeatingSection', 2, 2, 1, 3)
  add(value, 'item-first', 'repeatingSectionItem', 2, 2, 1, 2, { parentId: 'section' })
  add(value, 'item-second', 'repeatingSectionItem', 2, 2, 2, 3, { parentId: 'section' })
  fields.setDocumentTextFields(block, [field('first-field', 2, 2, 1, 2), field('second-field', 2, 2, 2, 3)])
  const raw = Object.fromEntries(block.runs.filter(run => run.metadata?.note).map(run => [run.id, { id: run.id,
    referenceBlockId: 'p', referenceRunId: run.id, customMark: run.id === 'visible' ? '1' : '', text: `Body ${run.id}`,
    content: { version: 1, runs: [], paragraphs: [], blocks: [{ ...model.createParagraph(`Body ${run.id}`), id: `body-${run.id}` }] } }]))
  value.footnotes = notes.normalizeDocumentNotes({ footnotes: raw }, value.blocks).footnotes
  const originalBodies = copy(value.footnotes), copied = controls.duplicateRepeatingSectionItem(value, 'item-first')
  const live = value.blocks[0], refs = bounds.documentNoteReferenceOffsets(live.runs)
  const selected = controlBounds.partitionDocumentContentControlNoteRuns(live.runs, copied).selected
  assert.equal(selected.filter(run => run.metadata?.note).length, 1)
  assert.notEqual(selected[0].id, 'empty-first')
  assert.deepEqual(fieldRange(copied), [2, 2, 2, 3])
  assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'item-first')), [2, 1, 2, 2])
  assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'item-second')), [2, 3, 2, 4])
  assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'section')), [2, 1, 2, 4])
  assert.deepEqual(refs.map(reference => reference.runId).filter(id => ['visible', 'empty-first', 'empty-second'].includes(id)), ['visible', 'empty-first', 'empty-second'])
  assert.equal(Object.keys(value.footnotes).length, 4)
  for (const id of ['visible', 'empty-first', 'empty-second']) assert.equal(value.footnotes[id].text, originalBodies[id].text)
  const fieldCopies = fields.documentTextFields(live)
  assert.equal(fieldCopies.length, 3)
  assert.deepEqual(fieldRange(fieldCopies.find(field => field.id === 'first-field')), [2, 1, 2, 2])
  assert.deepEqual(fieldRange(fieldCopies.find(field => field.id === 'second-field')), [2, 3, 2, 4])
  assert.deepEqual(fieldRange(fieldCopies.find(field => !['first-field', 'second-field'].includes(field.id))), [2, 2, 2, 3])
  controls.validateDocumentContentControlRanges(value); fields.documentTextFieldRanges(value)
  controls.deleteRepeatingSectionItem(value, copied.id)
  assert.deepEqual(bounds.documentNoteReferenceOffsets(value.blocks[0].runs).map(reference => reference.runId), ['visible', 'empty-first', 'empty-second'])
  assert.equal(Object.keys(value.footnotes).length, 3)
  assert.deepEqual(fieldRange(value.contentControls.find(control => control.id === 'item-second')), [2, 2, 2, 3])
  assert.deepEqual(fields.documentTextFields(value.blocks[0]).map(field => field.id).sort(), ['first-field', 'second-field'])
  controls.validateDocumentContentControlRanges(value); fields.documentTextFieldRanges(value)
})

test('parser projection validates source cumulative geometry before mapping changed empty labels', () => {
  const { block } = fixture()
  const literal = block.runs.map(run => ({ type: 'text', text: run.text,
    ...(run.metadata?.noteReference ? { noteReference: { kind: run.metadata.noteReference.kind } } : {}) }))
  block.metadata.source = { details: { inlines: literal } }
  block.runs = block.runs.map((run, index) => ({ ...run, text: run.id === 'empty-first' ? 'X' : run.text,
    metadata: { ...run.metadata, officeParsedRunSource: { version: 1, index } } }))
  const mapped = projection.parsedDocumentTextProjection(block)
  assert.equal(mapped.sourceText, 'A1B'); assert.equal(mapped.targetText, 'A1XB')
  assert.deepEqual([1, 2, 3].map(rank => mapped.mapBoundary(2, rank)), [2, 3, 3])
  for (const rank of [0, 4, null, '1', -1]) assert.equal(mapped.mapBoundary(2, rank), undefined)
  assert.equal(mapped.mapBoundary(3, 0), undefined)
  assert.equal(mapped.mapBoundary(3, 3), 4)
  const plain = projection.parsedDocumentTextProjection(model.createParagraph('AB'))
  assert.equal(plain.mapBoundary(1, 0), 1); assert.equal(plain.mapBoundary(1, 1), undefined)
})
