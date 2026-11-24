import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred DOC-11 input contracts. These drafts have not been executed.
let server, domain, modelUtils, commands, controls, editing, presentation, fields
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-editing', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  commands = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  editing = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-control-editing.ts')
  presentation = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-control-presentation.ts')
  fields = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-fields.ts')
})
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const blockText = block => block.runs.map(run => run.text).join('')
const bodyText = model => model.blocks.map(blockText).join('\n')
const info = (model, id = 'control') => model.contentControls.find(control => control.id === id)
const selection = (start, end = start, target = { blockId: 'p' }) => ({ anchor: { ...target, offset: start }, focus: { ...target, offset: end } })
function fixture(text = 'abcd') {
  const model = domain.createBlankDocument(); model.id = 'content-control-input'; model.blocks[0].id = 'p'; model.blocks[0].runs[0].text = text
  model.contentControls = []; return model
}
function add(model, properties = {}) {
  return controls.createDocumentContentControl(model, { id: 'control', type: 'text', blockId: 'p', startOffset: 0,
    endOffset: blockText(model.blocks[0]).length, lock: 'unlocked', rawSdtPr: '<w:sdtPr><w:tag w:val="retained"/></w:sdtPr>',
    metadata: { originalSource: { version: 1, part: 'word/document.xml' } }, ...properties })
}
function input(model, sourceSelection, text, options = {}) {
  const before = copy(model), draft = copy(model)
  const prepared = editing.prepareDocumentContentControlTextInputSelection(before, sourceSelection,
    options.activeControlId ? { controlId: options.activeControlId } : {})
  const caret = commands.executeDocumentCommand(draft, prepared, { id: 'doc.command.insert-text', text,
    ...(options.activeControlId ? { activeContentControlId: options.activeControlId } : {}),
    ...(options.host ? { host: options.host } : {}) })
  const result = editing.synchronizeDocumentContentControlsAfterEdit(before, draft, {
    origin: 'text-input', selection: prepared, ...(options.activeControlId ? { activeControlId: options.activeControlId } : {}),
  })
  return { draft, prepared, caret, result }
}

test('ordinary replacement removes a temporary shell and retains UTF-16 text in the caller-owned candidate', () => {
  const model = fixture('a😀b'); add(model, { value: 'a😀b', temporary: true })
  const before = copy(model), { draft, result } = input(model, selection(1, 3), '中')
  assert.equal(bodyText(draft), 'a中b'); assert.deepEqual(draft.contentControls, [])
  assert.deepEqual(result.removedTemporaryControlIds, ['control']); assert.deepEqual(result.changedControlIds, ['control'])
  assert.deepEqual(model, before)
})

test('no-op and formatting-only candidates keep temporary, placeholder, binding and unknown metadata intact', () => {
  const model = fixture(); add(model, { value: 'abcd', temporary: true, dataBinding: { xpath: '/name', dirty: false } })
  const before = copy(model), draft = copy(model)
  assert.deepEqual(editing.synchronizeDocumentContentControlsAfterEdit(model, draft, { origin: 'text-input', selection: selection(2) }), {
    changedControlIds: [], removedTemporaryControlIds: [], retainedLockedTemporaryControlIds: [],
  })
  assert.deepEqual(draft, before)
  draft.blocks[0].runs[0].bold = true
  assert.deepEqual(editing.synchronizeDocumentContentControlsAfterEdit(model, draft, { origin: 'text-input' }).changedControlIds, [])
  assert.deepEqual(draft.contentControls, before.contentControls); assert.equal(draft.blocks[0].runs[0].bold, true)
  for (const id of ['doc.command.format-text', 'doc.command.clear-formatting', 'doc.command.update-content-control',
    'doc.command.update-selected-fields', 'doc.command.refresh-header-footer-fields']) assert.equal(editing.documentContentControlCommandIsTextInput(id), false)
  for (const id of ['doc.command.insert-text', 'doc.command.delete', 'doc.command.paste-blocks', 'doc.command.table-set-cell-text',
    'doc.command.table-set-range-text', 'doc.command.replace-range']) assert.equal(editing.documentContentControlCommandIsTextInput({ id }), true)
})

test('explicit same values are no-ops while equal list labels with a different stored value still remove temporary', () => {
  const text = fixture(); add(text, { value: 'abcd', temporary: true, dataBinding: { dirty: false } })
  const unchanged = copy(text); controls.setDocumentContentControlValue(text, 'control', 'abcd'); assert.deepEqual(text, unchanged)
  const checkbox = fixture('_'); add(checkbox, { type: 'checkbox', temporary: true, dataBinding: { dirty: false } })
  const unchecked = copy(checkbox); controls.setDocumentContentControlChecked(checkbox, 'control', false); assert.deepEqual(checkbox, unchecked)
  const list = fixture('Alpha'); add(list, { type: 'dropDownList', temporary: true, value: 'A',
    items: [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Alpha' }], dataBinding: { dirty: false } })
  const runs = copy(list.blocks[0].runs)
  controls.setDocumentContentControlValue(list, 'control', 'B')
  assert.deepEqual(list.contentControls, []); assert.equal(bodyText(list), 'Alpha'); assert.deepEqual(list.blocks[0].runs, runs)
})

test('wrapper locks retain temporary after content edits, and content locks reject the candidate', () => {
  const model = fixture('ab'); add(model, { value: 'ab', temporary: true, lock: 'sdtLocked', dataBinding: { dirty: false } })
  const { draft, result } = input(model, selection(1), 'X')
  assert.equal(info(draft).value, 'aXb'); assert.equal(info(draft).dataBinding.dirty, true); assert.equal(info(draft).temporary, true)
  assert.deepEqual(result.retainedLockedTemporaryControlIds, ['control']); assert.deepEqual(result.removedTemporaryControlIds, [])
  for (const lock of ['contentLocked', 'sdtContentLocked']) {
    const locked = copy(model); info(locked).lock = lock; const before = copy(locked)
    assert.throws(() => input(locked, selection(1), 'X'), /locked|锁定/); assert.deepEqual(locked, before)
  }
})

test('a structural candidate cannot bypass content or ancestor locks by removing every temporary shell first', () => {
  for (const ancestor of [false, true]) {
    const model = fixture('AB')
    if (ancestor) add(model, { id: 'parent', type: 'group', value: 'AB' })
    add(model, { temporary: true, value: 'AB', ...(ancestor ? { parentId: 'parent' } : {}) })
    info(model, ancestor ? 'parent' : 'control').lock = 'contentLocked'
    const source = copy(model), draft = copy(model); draft.contentControls = []; draft.blocks[0].runs[0].text = ''
    const staged = copy(draft)
    assert.throws(() => editing.synchronizeDocumentContentControlsAfterEdit(model, draft, { origin: 'text-input', selection: selection(0, 2) }), /locked|锁定/)
    assert.deepEqual(draft, staged); assert.deepEqual(model, source)
  }
})

test('temporary ancestors reparent surviving descendants to the nearest retained wrapper', () => {
  const model = fixture('abcdef')
  add(model, { id: 'root', type: 'group', startOffset: 0, endOffset: 6, dataBinding: { dirty: false } })
  add(model, { id: 'middle', type: 'group', startOffset: 1, endOffset: 5, parentId: 'root', temporary: true })
  add(model, { id: 'inner', type: 'group', startOffset: 2, endOffset: 4, parentId: 'middle', temporary: true })
  add(model, { id: 'leaf', type: 'richText', startOffset: 2, endOffset: 4, parentId: 'inner', value: 'cd',
    metadata: { unknown: { retained: ['x', 7] } }, dataBinding: { xpath: '/child', dirty: false } })
  const { draft, result } = input(model, selection(3), 'X')
  assert.equal(bodyText(draft), 'abcXdef'); assert.deepEqual(result.removedTemporaryControlIds.sort(), ['inner', 'middle'])
  assert.equal(info(draft, 'leaf').parentId, 'root'); assert.equal(info(draft, 'leaf').value, 'cXd')
  assert.deepEqual(info(draft, 'leaf').metadata, { unknown: { retained: ['x', 7] } })
  assert.equal(info(draft, 'leaf').dataBinding.dirty, true); assert.equal(info(draft, 'root').dataBinding.dirty, true)
  assert.deepEqual([info(draft, 'leaf').startOffset, info(draft, 'leaf').endOffset, info(draft, 'root').endOffset], [2, 5, 7])
})

test('a locked ancestor rejects explicit child values before either text or wrapper state is published', () => {
  const model = fixture('abcd'); add(model, { id: 'parent', type: 'group' })
  add(model, { value: 'bc', startOffset: 1, endOffset: 3, parentId: 'parent', temporary: true })
  info(model, 'parent').lock = 'contentLocked'; const before = copy(model)
  assert.throws(() => controls.setDocumentContentControlValue(model, 'control', 'XX'), /locked|锁定/)
  assert.deepEqual(model, before)
})

test('a later invalid dropdown edit cannot publish an earlier checkbox state or font plan', () => {
  const model = fixture('☐Alpha'); add(model, { id: 'check', type: 'checkbox', startOffset: 0, endOffset: 1 })
  add(model, { id: 'list', type: 'dropDownList', startOffset: 1, endOffset: 6, value: 'A', items: [{ value: 'A', displayText: 'Alpha' }] })
  model.blocks[0].runs[0].fontFamily = 'Original Font'
  const source = copy(model), draft = copy(model); draft.blocks[0].runs = [modelUtils.createRun('☒Gamma', { fontFamily: 'Original Font' })]
  const staged = copy(draft)
  assert.throws(() => editing.synchronizeDocumentContentControlsAfterEdit(model, draft, { origin: 'text-input' }), /下拉列表/)
  assert.deepEqual(draft, staged); assert.deepEqual(model, source)
})

test('placeholder preparation selects the complete deepest prompt and preserves reversed selection and cell coordinates', () => {
  const model = fixture('0123456789'); add(model, { id: 'parent', type: 'group', showingPlaceholder: true })
  add(model, { startOffset: 2, endOffset: 8, parentId: 'parent', showingPlaceholder: true })
  const source = selection(6, 4), before = copy(model), prepared = editing.prepareDocumentContentControlTextInputSelection(model, source)
  assert.deepEqual(prepared, selection(8, 2)); assert.deepEqual(source, selection(6, 4)); assert.deepEqual(model, before)
  const paragraph = model.blocks[0]; model.blocks = [{ ...modelUtils.createParagraph(), id: 'table', type: 'table', runs: [],
    metadata: { table: { rows: [{ id: 'row', cells: [{ id: 'cell', blocks: [paragraph], style: {}, rowSpan: 1, colSpan: 1 }] }] } } }]
  const target = { blockId: 'table', row: 0, col: 0, cellBlockId: 'p' }
  assert.deepEqual(editing.prepareDocumentContentControlTextInputSelection(model, selection(4, 6, target)), selection(2, 8, target))
})

test('the first ordinary placeholder input replaces its prompt and synchronizes value, placeholder and dirty', () => {
  const model = fixture('Name'); add(model, { value: 'Name', placeholder: 'Name', showingPlaceholder: true, dataBinding: { dirty: false } })
  const { draft, prepared } = input(model, selection(2), '李😀')
  assert.deepEqual(prepared, selection(0, 4)); assert.equal(bodyText(draft), '李😀')
  assert.equal(info(draft).value, '李😀'); assert.equal(info(draft).endOffset, 3)
  assert.equal(info(draft).showingPlaceholder, false); assert.equal(info(draft).dataBinding.dirty, true)
})

test('empty insertion grows the selected control and ancestors while ordinary boundaries remain half-open', () => {
  const model = fixture('AB'); add(model, { id: 'parent', type: 'group', startOffset: 1, endOffset: 1 })
  add(model, { startOffset: 1, endOffset: 1, parentId: 'parent', value: '' })
  const { draft } = input(model, selection(1), '😀')
  assert.equal(bodyText(draft), 'A😀B'); assert.equal(info(draft).value, '😀')
  assert.deepEqual([info(draft).startOffset, info(draft).endOffset, info(draft, 'parent').startOffset, info(draft, 'parent').endOffset], [1, 3, 1, 3])
  const bounded = fixture('abcd'); add(bounded, { startOffset: 1, endOffset: 3, value: 'bc' })
  assert.deepEqual([info(input(bounded, selection(1), 'X').draft).startOffset, info(input(bounded, selection(3), 'X').draft).endOffset], [2, 3])
  const insideStart = input(bounded, selection(1), 'X', { activeControlId: 'control' }).draft
  const insideEnd = input(bounded, selection(3), 'X', { activeControlId: 'control' }).draft
  assert.equal(info(insideStart).value, 'Xbc'); assert.equal(info(insideEnd).value, 'bcX')
  assert.deepEqual([info(insideStart).startOffset, info(insideStart).endOffset, info(insideEnd).startOffset, info(insideEnd).endOffset], [1, 4, 1, 4])
})

test('ordinary checkbox input honors configured Unicode glyphs, boolean state and font without guessing arbitrary characters', () => {
  const model = fixture('○'); add(model, { type: 'checkbox', checkbox: { checked: false,
    checkedState: { value: '1F7E2', font: 'Configured Font' }, uncheckedState: { value: '25CB', font: 'Configured Font' } } })
  const { draft } = input(model, selection(0, 1), '🟢')
  assert.equal(info(draft).checked, true); assert.equal(info(draft).checkbox.checked, true); assert.equal(info(draft).value, 'true')
  assert.equal(info(draft).endOffset, 2); assert.equal(draft.blocks[0].runs[0].fontFamily, 'Configured Font')
  const before = copy(model); assert.throws(() => input(model, selection(0, 1), 'yes'), /复选框/); assert.deepEqual(model, before)
})

test('filling an empty nested value grows every ancestor before temporary parents unwrap', () => {
  const model = fixture('AB'); add(model, { id: 'root', type: 'group', startOffset: 1, endOffset: 1 })
  add(model, { id: 'middle', type: 'group', parentId: 'root', startOffset: 1, endOffset: 1, temporary: true })
  add(model, { parentId: 'middle', startOffset: 1, endOffset: 1, value: '' })
  controls.setDocumentContentControlValue(model, 'control', '😀')
  assert.equal(bodyText(model), 'A😀B'); assert.equal(info(model, 'middle'), undefined); assert.equal(info(model).parentId, 'root')
  assert.deepEqual([info(model).startOffset, info(model).endOffset, info(model, 'root').startOffset, info(model, 'root').endOffset], [1, 3, 1, 3])
})

test('ordinary date text updates ISO storage when unambiguous and preserves the last known date for invalid text', () => {
  const model = fixture('03 Sep 2026'); add(model, { type: 'date', value: '2026-09-03', dateValue: '2026-09-03', dateFormat: 'dd MMM yyyy', dataBinding: { dirty: false } })
  const valid = input(model, selection(0, 11), '04 Sep 2026').draft
  assert.equal(info(valid).dateValue, '2026-09-04'); assert.equal(info(valid).value, '2026-09-04')
  assert.equal(info(valid).metadata.contentControlLastKnownDate, '2026-09-04'); assert.equal(info(valid).dataBinding.dirty, true)
  const raw = input(model, selection(0, 11), 'TBD').draft
  assert.equal(info(raw).value, 'TBD'); assert.equal(info(raw).dateValue, undefined)
  assert.equal(info(raw).metadata.contentControlLastKnownDate, '2026-09-03'); assert.equal(info(raw).rawSdtPr, info(model).rawSdtPr)
  controls.setDocumentContentControlValue(model, 'control', '')
  assert.equal(info(model).dateValue, ''); assert.equal(info(model).value, ''); assert.equal(info(model).metadata.contentControlLastKnownDate, '2026-09-03')
})

test('combo input stores its actual free text while dropdown input remains a selection command', () => {
  const model = fixture('Alpha'); add(model, { type: 'comboBox', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Alpha' }] })
  const free = input(model, selection(0, 5), 'Beta').draft
  assert.equal(info(free).value, 'Beta'); assert.equal(bodyText(free), 'Beta')
  const list = fixture('Alpha'); add(list, { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Beta' }] })
  const before = copy(list); assert.throws(() => input(list, selection(0, 5), 'Beta'), /下拉列表/); assert.deepEqual(list, before)
  controls.setDocumentContentControlValue(list, 'control', 'B'); assert.equal(info(list).value, 'B'); assert.equal(bodyText(list), 'Beta')
})

test('a real paragraph break unwraps an unlocked temporary group and preserves disjoint prefix and suffix children', () => {
  const model = fixture('abcdef'); add(model, { id: 'group', type: 'group', startOffset: 1, endOffset: 5, temporary: true })
  add(model, { id: 'before', parentId: 'group', startOffset: 1, endOffset: 2, value: 'b' })
  add(model, { id: 'after', parentId: 'group', startOffset: 4, endOffset: 5, value: 'e', metadata: { unknown: 'retain' } })
  const { draft, result } = input(model, selection(3), '\n')
  assert.equal(bodyText(draft), 'abc\ndef'); assert.deepEqual(result.removedTemporaryControlIds, ['group'])
  assert.equal(info(draft, 'before').parentId, undefined); assert.equal(info(draft, 'before').blockId, 'p')
  assert.equal(info(draft, 'after').parentId, undefined); assert.equal(info(draft, 'after').blockId, draft.blocks[1].id)
  assert.deepEqual([info(draft, 'after').startOffset, info(draft, 'after').endOffset], [1, 2]); assert.equal(info(draft, 'after').value, 'e')
  assert.deepEqual(info(draft, 'after').metadata, { unknown: 'retain' })
})

test('a temporary empty rich-text control unwraps on an actual newline, while single-line types reject breaks atomically', () => {
  const model = fixture(''); add(model, { type: 'richText', value: '', temporary: true })
  const { draft, result } = input(model, selection(0), '\n')
  assert.equal(bodyText(draft), '\n'); assert.deepEqual(draft.contentControls, []); assert.deepEqual(result.removedTemporaryControlIds, ['control'])
  for (const properties of [{ type: 'text' }, { type: 'checkbox' }, { type: 'date', dateValue: '2026-11-24' }]) {
    const single = fixture('abcd'); add(single, { ...properties, temporary: true }); const before = copy(single)
    assert.throws(() => input(single, selection(0, info(single).endOffset), '\n'), /换行|range|range splits|位置|范围/); assert.deepEqual(single, before)
  }
})

test('persistent or wrapper-locked cross-paragraph ranges and preserved derived-cache paths remain atomic', () => {
  for (const properties of [{ type: 'richText' }, { type: 'richText', temporary: true, lock: 'sdtLocked' }]) {
    const model = fixture(); add(model, properties); const before = copy(model)
    assert.throws(() => input(model, selection(2), '\n'), /cannot split/); assert.deepEqual(model, before)
  }
  const model = fixture(); add(model, { type: 'richText', temporary: true })
  controls.transformContentControlsForBlockEdit(model, 'p', 1, 2, 1, undefined, true)
  assert.equal(info(model).temporary, true)
  const before = copy(model)
  assert.throws(() => controls.transformContentControlsForBlockEdit(model, 'p', 2, 2, 1,
    { blockId: 'suffix', firstBreakOffset: 2, lastParagraphOffset: 3 }, true), /cannot split/)
  assert.deepEqual(model, before)
})

test('explicit inside affinity unwraps temporary at either nonempty paragraph boundary in both range planners', () => {
  for (const offset of [1, 2]) for (const planner of ['block', 'body']) {
    const before = fixture('AB'); add(before, { type: 'richText', startOffset: 1, endOffset: 2, value: 'B', temporary: true })
    const draft = copy(before), firstText = offset === 1 ? 'A' : 'AB', suffixText = offset === 1 ? 'B' : ''
    if (planner === 'block') controls.transformContentControlsForBlockEdit(draft, 'p', offset, offset, 1,
      { blockId: 'suffix', firstBreakOffset: offset, lastParagraphOffset: offset + 1 }, false, { activeControlId: 'control' })
    draft.blocks[0].runs = [modelUtils.createRun(firstText)]
    const suffix = modelUtils.createParagraph(suffixText); suffix.id = 'suffix'; draft.blocks.push(suffix)
    if (planner === 'body') controls.transformContentControlsForBodyRangeEdit(draft, fields.documentTextFieldParagraphs(before), offset, offset, 1,
      false, { activeControlId: 'control' })
    const result = editing.synchronizeDocumentContentControlsAfterEdit(before, draft, { origin: 'text-input', selection: selection(offset), activeControlId: 'control' })
    assert.equal(bodyText(draft), firstText + '\n' + suffixText); assert.deepEqual(draft.contentControls, [])
    assert.deepEqual(result.removedTemporaryControlIds, ['control']); assert.equal(info(before).temporary, true)
  }
})

test('ordinary newline insertion at a nonempty boundary remains outside without an explicit control identity', () => {
  for (const offset of [1, 2]) for (const planner of ['block', 'body']) {
    const before = fixture('AB'); add(before, { type: 'richText', startOffset: 1, endOffset: 2, value: 'B', temporary: true })
    const draft = copy(before)
    if (planner === 'block') controls.transformContentControlsForBlockEdit(draft, 'p', offset, offset, 1,
      { blockId: 'suffix', firstBreakOffset: offset, lastParagraphOffset: offset + 1 })
    draft.blocks[0].runs = [modelUtils.createRun(offset === 1 ? 'A' : 'AB')]
    const suffix = modelUtils.createParagraph(offset === 1 ? 'B' : ''); suffix.id = 'suffix'; draft.blocks.push(suffix)
    if (planner === 'body') controls.transformContentControlsForBodyRangeEdit(draft, fields.documentTextFieldParagraphs(before), offset, offset, 1, false)
    const result = editing.synchronizeDocumentContentControlsAfterEdit(before, draft, { origin: 'text-input', selection: selection(offset) })
    assert.equal(info(draft).temporary, true); assert.equal(info(draft).value, 'B'); assert.deepEqual(result.changedControlIds, [])
    assert.equal(info(draft).blockId, offset === 1 ? 'suffix' : 'p')
  }
})

test('inside complete deletion preserves the active empty shell for persistent and wrapper-locked controls, and unwraps temporary', () => {
  for (const planner of ['block', 'body']) for (const properties of [{ temporary: false }, { temporary: true }, { temporary: true, lock: 'sdtLocked' }]) {
    const before = fixture('AB'); add(before, { startOffset: 1, endOffset: 2, value: 'B', dataBinding: { dirty: false }, ...properties })
    const draft = copy(before)
    if (planner === 'block') controls.transformContentControlsForBlockEdit(draft, 'p', 1, 2, 0, undefined, false, { activeControlId: 'control' })
    draft.blocks[0].runs = [modelUtils.createRun('A')]
    if (planner === 'body') controls.transformContentControlsForBodyRangeEdit(draft, fields.documentTextFieldParagraphs(before), 1, 2, 0,
      false, { activeControlId: 'control' })
    assert.deepEqual([info(draft).startOffset, info(draft).endOffset], [1, 1])
    const result = editing.synchronizeDocumentContentControlsAfterEdit(before, draft, { origin: 'text-input', selection: selection(1, 2), activeControlId: 'control' })
    if (properties.temporary && properties.lock !== 'sdtLocked') {
      assert.equal(info(draft), undefined); assert.deepEqual(result.removedTemporaryControlIds, ['control'])
    } else {
      assert.equal(info(draft).value, ''); assert.equal(info(draft).showingPlaceholder, false); assert.equal(info(draft).dataBinding.dirty, true)
      assert.deepEqual(result.removedTemporaryControlIds, [])
      assert.deepEqual(result.retainedLockedTemporaryControlIds, properties.temporary ? ['control'] : [])
    }
    assert.equal(bodyText(draft), 'A'); assert.equal(info(before).value, 'B')
  }
})

test('inside affinity includes coincident ancestors and rejects stale identities or locked contents before anchor publication', () => {
  const source = fixture('AB')
  add(source, { id: 'parent', type: 'group', startOffset: 1, endOffset: 2, temporary: true })
  add(source, { parentId: 'parent', type: 'richText', startOffset: 1, endOffset: 2, value: 'B', temporary: true })
  const draft = copy(source)
  controls.transformContentControlsForBlockEdit(draft, 'p', 1, 1, 1,
    { blockId: 'suffix', firstBreakOffset: 1, lastParagraphOffset: 2 }, false, { activeControlId: 'control' })
  assert.deepEqual(draft.contentControls, [])
  for (const [id, offset] of [['missing', 1], ['control', 0]]) {
    const candidate = copy(source), before = copy(candidate)
    assert.throws(() => controls.transformContentControlsForBlockEdit(candidate, 'p', offset, offset, 1, undefined, false, { activeControlId: id }), /上下文/)
    assert.deepEqual(candidate, before)
  }
  for (const planner of ['block', 'body']) {
    const locked = copy(source); info(locked, 'parent').lock = 'contentLocked'; const before = copy(locked)
    assert.throws(() => planner === 'block'
      ? controls.transformContentControlsForBlockEdit(locked, 'p', 1, 1, 1, undefined, false, { activeControlId: 'control' })
      : controls.transformContentControlsForBodyRangeEdit(locked, fields.documentTextFieldParagraphs(before), 1, 1, 1, false, { activeControlId: 'control' }), /locked|锁定/)
    assert.deepEqual(locked, before)
  }
})

function storyContent(id, text) {
  const paragraph = modelUtils.createParagraph(text); paragraph.id = id
  return { blocks: [paragraph], runs: copy(paragraph.runs), paragraphs: [{ startOffset: 0, endOffset: text.length, style: copy(paragraph.style) }] }
}
function hostedFixture() {
  const model = fixture('BODY'), cell = modelUtils.createParagraph('CELL'); cell.id = 'cell-p'
  model.blocks.push({ ...modelUtils.createParagraph(), id: 'table', type: 'table', runs: [], metadata: {
    table: { rows: [{ id: 'row', cells: [{ id: 'cell', blocks: [cell], style: {}, rowSpan: 1, colSpan: 1 }] }] },
  } })
  model.sectionHeaderFooters = { first: { header: 'HEAD', footer: 'FOOT', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '',
    differentFirstPage: false, differentOddEven: false, linkToPrevious: false,
    storyContent: { header: storyContent('header-p', 'HEAD'), footer: storyContent('footer-p', 'FOOT') } } }
  model.footnotes = { note: { id: 'note', text: 'NOTE', referenceBlockId: 'p', referenceRunId: model.blocks[0].runs[0].id, content: storyContent('note-p', 'NOTE') } }
  for (const [id, blockId, value] of [['body', 'p', 'BODY'], ['cell', 'cell-p', 'CELL'], ['header', 'header-p', 'HEAD'], ['footer', 'footer-p', 'FOOT'], ['note', 'note-p', 'NOTE']]) {
    add(model, { id, blockId, startOffset: 0, endOffset: 4, value, dataBinding: { dirty: false } })
  }
  return model
}

test('body, table-cell, header, footer and note text use the same state synchronizer without editing another host', () => {
  const model = hostedFixture()
  const cases = [
    { id: 'body', target: { blockId: 'p' } }, { id: 'cell', target: { blockId: 'table', row: 0, col: 0, cellBlockId: 'cell-p' } },
    { id: 'header', target: { blockId: 'header-p' }, host: { kind: 'story', sectionId: 'first', story: 'header' } },
    { id: 'footer', target: { blockId: 'footer-p' }, host: { kind: 'story', sectionId: 'first', story: 'footer' } },
    { id: 'note', target: { blockId: 'note-p' }, host: { kind: 'note', noteKind: 'footnote', noteId: 'note' } },
  ]
  for (const sample of cases) {
    const { draft, result } = input(model, selection(1, 2, sample.target), '中', { host: sample.host })
    assert.equal(info(draft, sample.id).value, info(model, sample.id).value.slice(0, 1) + '中' + info(model, sample.id).value.slice(2))
    assert.equal(info(draft, sample.id).dataBinding.dirty, true); assert.deepEqual(result.changedControlIds, [sample.id])
    for (const other of cases.filter(candidate => candidate.id !== sample.id)) assert.deepEqual(info(draft, other.id), info(model, other.id))
    if (sample.host) assert.deepEqual(draft.blocks, model.blocks)
    assert.equal(draft.metadata.documentEditingBlockHost, undefined)
  }
})

test('cell and range text setters synchronize scalar values, unwrap multiline temporary and preserve no-ops', () => {
  for (const sample of [
    { id: 'doc.command.table-set-cell-text', text: 'VALUE', expected: 'VALUE' },
    { id: 'doc.command.table-set-range-text', mode: 'matrix', values: [['MATRIX']], expected: 'MATRIX' },
    { id: 'doc.command.table-set-cell-text', text: 'First\nSecond', expected: 'First\nSecond', temporary: true },
    { id: 'doc.command.table-set-cell-text', text: 'one\ntwo', expected: 'one\ntwo', temporary: true, empty: true },
    { id: 'doc.command.table-set-cell-text', text: 'CELL', expected: 'CELL', temporary: true, noop: true },
  ]) {
    const model = hostedFixture(), owner = model.blocks[1].metadata.table.rows[0].cells[0]
    if (sample.temporary) { info(model, 'cell').type = 'richText'; info(model, 'cell').temporary = true }
    if (sample.empty) { owner.blocks[0].runs[0].text = ''; info(model, 'cell').endOffset = 0; info(model, 'cell').value = '' }
    const before = copy(model), draft = copy(model), sourceSelection = selection(0, 0, { blockId: 'table', row: 0, col: 0, cellBlockId: 'cell-p' })
    const command = sample.id === 'doc.command.table-set-cell-text'
      ? { id: sample.id, tableId: 'table', row: 0, column: 0, text: sample.text }
      : { id: sample.id, tableId: 'table', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, mode: sample.mode, values: sample.values }
    commands.executeDocumentCommand(draft, sourceSelection, command)
    const result = editing.synchronizeDocumentContentControlsAfterEdit(model, draft, { origin: 'text-input', selection: sourceSelection })
    assert.equal(draft.blocks[1].metadata.table.rows[0].cells[0].blocks.map(blockText).join('\n'), sample.expected)
    if (sample.noop) { assert.deepEqual(draft, before); assert.deepEqual(result.changedControlIds, []) }
    else if (sample.temporary) { assert.equal(info(draft, 'cell'), undefined); assert.deepEqual(result.removedTemporaryControlIds, ['cell']) }
    else assert.equal(info(draft, 'cell').value, sample.expected)
    for (const id of ['body', 'header', 'footer', 'note']) assert.deepEqual(info(draft, id), info(model, id))
    assert.deepEqual(model, before)
  }
})

test('a full-model checkbox font plan updates the host projection together with its authoritative block', () => {
  const model = hostedFixture()
  controls.updateDocumentContentControl(model, 'header', { type: 'checkbox', checkbox: { checked: false,
    checkedState: { value: '2612', font: 'Checked Font' }, uncheckedState: { value: '2610', font: 'Unchecked Font' } } })
  const before = copy(model), draft = copy(model), sourceSelection = selection(0, 1, { blockId: 'header-p' })
  commands.executeDocumentCommand(draft, sourceSelection, { id: 'doc.command.insert-text', text: '☒', marks: { fontFamily: 'Typing Font' },
    host: { kind: 'story', sectionId: 'first', story: 'header' } })
  editing.synchronizeDocumentContentControlsAfterEdit(model, draft, { origin: 'text-input', selection: sourceSelection })
  const content = draft.sectionHeaderFooters.first.storyContent.header
  assert.equal(content.blocks[0].runs[0].fontFamily, 'Checked Font'); assert.equal(content.runs[0].fontFamily, 'Checked Font')
  assert.equal(draft.sectionHeaderFooters.first.header, '☒'); assert.equal(info(draft, 'header').checked, true)
  assert.deepEqual(draft.blocks, before.blocks); assert.deepEqual(model, before)
})

test('cross-host moves, duplicate paragraph ownership and surrogate-splitting anchors are rejected before publication', () => {
  const model = hostedFixture(), draft = copy(model); info(draft, 'body').blockId = 'header-p'; const staged = copy(draft)
  assert.throws(() => editing.synchronizeDocumentContentControlsAfterEdit(model, draft, { origin: 'text-input' }), /其他内容区/)
  assert.deepEqual(draft, staged)
  const duplicate = copy(model); duplicate.sectionHeaderFooters.first.storyContent.header.blocks[0].id = 'p'
  assert.throws(() => editing.synchronizeDocumentContentControlsAfterEdit(duplicate, copy(duplicate), { origin: 'text-input' }), /重复/)
  const emoji = fixture('😀'); add(emoji, { value: '😀' }); const invalid = copy(emoji); info(invalid).endOffset = 1
  assert.throws(() => editing.synchronizeDocumentContentControlsAfterEdit(emoji, invalid, { origin: 'text-input' }), /UTF-16/)
})

test('value replacement preserves boundary note entities exactly once and does not inherit their identity into new text', () => {
  const model = fixture('ABCD'); add(model, { startOffset: 1, endOffset: 3, value: 'BC' })
  const original = model.blocks[0].runs[0]
  model.blocks[0].runs = [{ ...copy(original), id: 'prefix', text: 'A' }, { ...copy(original), id: 'note', text: '', metadata: { noteReference: { id: 'note-1' } } },
    { ...copy(original), id: 'selected', text: 'BC' }, { ...copy(original), id: 'end-field', text: '', metadata: { fieldEntity: 'retain' } },
    { ...copy(original), id: 'suffix', text: 'D' }]
  controls.setDocumentContentControlValue(model, 'control', 'XYZ')
  assert.equal(bodyText(model), 'AXYZD'); assert.equal(model.blocks[0].runs.filter(run => run.id === 'note').length, 1)
  assert.equal(model.blocks[0].runs.filter(run => run.id === 'end-field').length, 1)
  assert.equal(model.blocks[0].runs.find(run => run.text === 'XYZ').metadata?.noteReference, undefined)
  assert.equal(info(model).value, 'XYZ'); assert.deepEqual([info(model).startOffset, info(model).endOffset], [1, 4])
})

test('Gregorian reverse parsing is strict about format, weekday, literals, locale and actual dates', () => {
  for (const [text, format, locale, expected] of [
    ['2026-11-24', 'yyyy-MM-dd', 'en-US', '2026-11-24'], ['2024/02/29', 'yyyy/MM/dd', 'en-US', '2024-02-29'],
    ['24 November 2026', 'dd MMMM yyyy', 'en-US', '2026-11-24'], ['Tuesday, 24 November 2026', 'dddd, dd MMMM yyyy', 'en-US', '2026-11-24'],
    ['date: 2026-11-24', "'date: 'yyyy-MM-dd", 'en-US', '2026-11-24'], ['2026-11-24T', 'yyyy-MM-dd\\T', 'en-US', '2026-11-24'],
    ['24 KASIM 2026', 'dd MMMM yyyy', 'tr-TR', '2026-11-24'], ['24 November 2026', 'dd MMMM yyyy', 'en-US-u-ca-islamic', '2026-11-24'],
    ['2026/02/29', 'yyyy/MM/dd', 'en-US', undefined], ['Wednesday, 24 November 2026', 'dddd, dd MMMM yyyy', 'en-US', undefined],
    ['11/24/26', 'MM/dd/yy', 'en-US', undefined], ['11/24', 'MM/dd', 'en-US', undefined],
    ['2026/13/01', 'yyyy/MM/dd', 'en-US', undefined], ['prefix2026/11/24', "'prefixyyyy/MM/dd", 'en-US', undefined],
  ]) assert.equal(presentation.parseContentControlDateText(text, format, locale), expected, text)
  assert.equal(presentation.parseContentControlDateText(''), '')
  assert.equal(presentation.parseContentControlDateText('1448/03/04', 'yyyy/MM/dd', 'en-US', 'hijri'), undefined)
  assert.equal(presentation.parseContentControlDateText('2026-11-24', 'yyyy/MM/dd', 'en-US', 'hijri'), '2026-11-24')
})
