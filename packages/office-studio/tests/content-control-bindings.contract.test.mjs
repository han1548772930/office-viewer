import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-11 scalar binding drafts. Execution is deferred to the unified validation phase.
let server, api, models, controls, bindings, editor, code, search, fields, comments, tableOrigins, tableMutations
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-bindings', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  models = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  bindings = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-control-bindings.ts')
  editor = await server.ssrLoadModule('/packages/office-studio/src/api/document-editor-commands.ts')
  code = await server.ssrLoadModule('/packages/office-studio/src/features/document/code-blocks/code-blocks.ts')
  search = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-search.ts')
  fields = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-fields.ts')
  comments = await server.ssrLoadModule('/packages/office-studio/src/features/document/comments/thread-comments.ts')
  tableOrigins = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-table-text-origin.ts')
  tableMutations = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/table-mutations.ts')
})
afterAll(async () => { await server?.close() })

// Facade models can be Vue proxies; snapshots use the project's JSON contract.
const copy = value => JSON.parse(JSON.stringify(value))
const permission = { documentId: 'sdt-bindings', userId: 'editor' }
const binding = (xpath = '/person/name') => ({ storeItemID: '{STORE}', xpath, prefixMappings: 'xmlns:p="urn:person"', cachedResult: 'old', dirty: false })
const info = (model, id) => model.contentControls?.find(control => control.id === id)
const controlText = (model, id) => controls.contentControlText(model, info(model, id))
const selection = (blockId, start, end = start, extra = {}) => ({ anchor: { blockId, offset: start, ...extra }, focus: { blockId, offset: end, ...extra } })
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })
function paragraph(id, value) { const block = models.createParagraph(value); block.id = id; return block }
function fixture(value = 'old') {
  const model = api.createBlankDocument(); model.id = permission.documentId
  model.blocks = [paragraph('p', value)]; model.contentControls = []
  return model
}
function add(model, id, blockId, value, patch = {}) {
  return controls.createDocumentContentControl(model, { id, type: 'richText', blockId, startOffset: 0, endOffset: value.length,
    value, dataBinding: binding(), metadata: { source: { part: 'retained', id } }, ...patch })
}
function pair(value = 'old', sourcePatch = {}, peerPatch = {}) {
  const model = fixture(value); model.blocks.push(paragraph('peer-p', value))
  add(model, 'origin', 'p', value, sourcePatch); add(model, 'peer', 'peer-p', value, peerPatch)
  return model
}
function content(id, value) {
  const block = paragraph(id, value)
  return { blocks: [block], runs: copy(block.runs), paragraphs: [{ startOffset: 0, endOffset: value.length, style: copy(block.style) }] }
}
function hostedFixture() {
  const model = fixture(); model.blocks[0].id = 'body-p'
  const cell = paragraph('cell-p', 'old')
  model.blocks.push({ ...paragraph('table', ''), type: 'table', runs: [], metadata: {
    table: { rows: [{ id: 'row', cells: [{ id: 'cell', blocks: [cell], style: {}, rowSpan: 1, colSpan: 1 }] }] },
  } })
  model.sectionHeaderFooters = { first: { header: 'old', footer: 'old', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '',
    differentFirstPage: false, differentOddEven: false, linkToPrevious: false,
    storyContent: { header: content('header-p', 'old'), footer: content('footer-p', 'old') } } }
  for (const [kind, id, blockId] of [['footnote', 'note', 'note-p'], ['endnote', 'end', 'zz-endnote-p']]) {
    const reference = paragraph(kind + '-reference', '1')
    reference.runs[0].metadata = { note: { type: kind, noteId: id } }; reference.runs[0].verticalAlign = 'superscript'
    model.blocks.push(reference)
    model[kind === 'footnote' ? 'footnotes' : 'endnotes'] = { [id]: { id, text: 'old', referenceBlockId: reference.id,
      referenceRunId: reference.runs[0].id, content: content(blockId, 'old') } }
  }
  for (const [id, blockId] of [['body', 'body-p'], ['cell', 'cell-p'], ['header', 'header-p'], ['footer', 'footer-p'],
    ['note', 'note-p'], ['end', 'zz-endnote-p']]) add(model, id, blockId, 'old')
  return model
}
function assertHosted(model, value, omitted = []) {
  for (const id of ['body', 'cell', 'header', 'footer', 'note', 'end'].filter(id => !omitted.includes(id))) {
    assert.equal(controlText(model, id), value); assert.equal(info(model, id).value, value)
    assert.equal(info(model, id).dataBinding.cachedResult, value); assert.equal(info(model, id).dataBinding.dirty, true)
  }
  assert.equal(model.sectionHeaderFooters.first.header, value); assert.equal(model.sectionHeaderFooters.first.footer, value)
  assert.equal(model.footnotes.note.text, value); assert.equal(model.endnotes.end.text, value)
  assert.equal(model.metadata.documentEditingBlockHost, undefined)
}
function open(model) {
  model.wordCount = models.textTargets(model).reduce((count, target) => count + models.runText(models.runsForTarget(model, target)).replace(/\s/g, '').length, 0)
  const studio = new api.OfficeStudio({ document: model, currentUserId: permission.userId }); studio.activateDocument(model.id)
  return { studio, document: new api.FDocument(studio, model.id) }
}
function setAndBind(model, id, value, options = {}) {
  const before = copy(model); controls.setDocumentContentControlValue(model, id, value)
  return bindings.synchronizeDocumentContentControlBindings(before, model, { sourceControlId: id, sourceValue: value, permissionContext: permission, ...options })
}

test('binding identity folds only store IDs and keeps XPath and namespace mappings distinct', () => {
  const key = value => bindings.documentContentControlBindingKey({ dataBinding: value })
  const base = binding()
  assert.equal(key(base), key({ ...base, storeItemID: '{store}' }))
  assert.notEqual(key(base), key({ ...base, xpath: '/person/Name' }))
  assert.notEqual(key(base), key({ ...base, prefixMappings: 'xmlns:p="urn:other"' }))
  assert.equal(key({ xpath: base.xpath }), undefined); assert.equal(key({ storeItemID: base.storeItemID }), undefined)
  assert.equal(key({ ...base, prefixMappings: undefined }), key({ ...base, prefixMappings: '' }))
})

test('binding values use saved list values, ISO dates and checkbox state while placeholders supply an empty scalar', () => {
  const model = fixture('Label')
  const text = { id: 'text', type: 'richText', blockId: 'p', startOffset: 0, endOffset: 5 }
  assert.equal(bindings.documentContentControlBindingValue(model, text), 'Label')
  assert.equal(bindings.documentContentControlBindingValue(model, { ...text, showingPlaceholder: true }), '')
  assert.equal(bindings.documentContentControlBindingValue(model, { ...text, type: 'dropDownList', value: 'code' }), 'code')
  assert.equal(bindings.documentContentControlBindingValue(model, { ...text, type: 'comboBox', value: 'free' }), 'free')
  assert.equal(bindings.documentContentControlBindingValue(model, { ...text, type: 'date', dateValue: '2026-11-24' }), '2026-11-24')
  assert.equal(bindings.documentContentControlBindingValue(model, { ...text, type: 'checkbox', checked: true }), 'true')
  assert.equal(bindings.documentContentControlBindingValue(model, { ...text, type: 'picture' }), undefined)
})

test('source collection returns detached scalar snapshots without publishing caches or linked text', () => {
  const before = pair(), draft = copy(before)
  controls.setDocumentContentControlValue(draft, 'origin', 'new😀')
  const staged = copy(draft), sources = bindings.documentContentControlBindingSources(before, draft)
  assert.deepEqual(sources.map(source => [source.control.id, source.key, source.value]),
    [['origin', bindings.documentContentControlBindingKey(info(draft, 'origin')), 'new😀']])
  sources[0].control.dataBinding.cachedResult = 'detached'; sources[0].control.metadata.source.part = 'detached'
  assert.deepEqual(draft, staged); assert.equal(controlText(draft, 'peer'), 'old')
  assert.equal(info(before, 'origin').dataBinding.cachedResult, 'old')
})

test('ordinary input from body, cells, header, footer and both note kinds synchronizes every same-binding scalar', () => {
  const cases = [
    { id: 'body', blockId: 'body-p' }, { id: 'cell', blockId: 'table', point: { row: 0, col: 0, cellBlockId: 'cell-p' } },
    { id: 'header', blockId: 'header-p', host: { kind: 'story', sectionId: 'first', story: 'header' } },
    { id: 'footer', blockId: 'footer-p', host: { kind: 'story', sectionId: 'first', story: 'footer' } },
    { id: 'note', blockId: 'note-p', host: { kind: 'note', noteKind: 'footnote', noteId: 'note' } },
    { id: 'end', blockId: 'zz-endnote-p', host: { kind: 'note', noteKind: 'endnote', noteId: 'end' } },
  ]
  for (const sample of cases) {
    const model = hostedFixture(), value = '李😀'
    const result = editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection(sample.blockId, 0, 3, sample.point),
      command: { id: 'doc.command.insert-text', text: value, host: sample.host, activeContentControlId: sample.id } }, permission)
    assertHosted(model, value); assert.equal(result.focus.offset, value.length)
    assert.equal(result.focus.cellBlockId ?? result.focus.blockId, info(model, sample.id).blockId)
  }
})

test('list peers display their own labels and keep a changed saved value even when the source label stays equal', () => {
  const model = fixture('Same'); model.blocks.push(paragraph('peer-p', 'First'))
  add(model, 'origin', 'p', 'Same', { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'Same' }, { value: 'B', displayText: 'Same' }] })
  add(model, 'peer', 'peer-p', 'First', { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'First' }, { value: 'B', displayText: 'Second' }] })
  setAndBind(model, 'origin', 'B')
  assert.equal(controlText(model, 'origin'), 'Same'); assert.equal(controlText(model, 'peer'), 'Second')
  assert.equal(info(model, 'origin').value, 'B'); assert.equal(info(model, 'peer').value, 'B')
  assert.equal(info(model, 'origin').dataBinding.cachedResult, 'B'); assert.equal(info(model, 'peer').dataBinding.cachedResult, 'B')
})

test('date peers share the ISO scalar while rendering their own Gregorian format', () => {
  const model = fixture('2026-11-24'); model.blocks.push(paragraph('peer-p', '24/11/2026'))
  add(model, 'origin', 'p', '2026-11-24', { type: 'date', dateValue: '2026-11-24', dateFormat: 'yyyy-MM-dd' })
  add(model, 'peer', 'peer-p', '24/11/2026', { type: 'date', dateValue: '2026-11-24', dateFormat: 'dd/MM/yyyy' })
  setAndBind(model, 'origin', '2026-11-25')
  assert.equal(controlText(model, 'peer'), '25/11/2026'); assert.equal(info(model, 'peer').dateValue, '2026-11-25')
  assert.equal(info(model, 'peer').dataBinding.cachedResult, '2026-11-25')
})

test('checkbox peers accept canonical booleans and numeric booleans and reject other source strings atomically', () => {
  for (const [value, expected] of [['true', true], ['1', true], ['false', false], ['0', false], ['', false]]) {
    const model = pair('old', {}, { type: 'checkbox' }); setAndBind(model, 'origin', value)
    assert.equal(info(model, 'peer').checked, expected); assert.equal(controlText(model, 'peer'), expected ? '☒' : '☐')
  }
  const before = pair('old', {}, { type: 'checkbox' }), draft = copy(before)
  controls.setDocumentContentControlValue(draft, 'origin', 'maybe'); const staged = copy(draft)
  assert.throws(() => bindings.synchronizeDocumentContentControlBindings(before, draft, { permissionContext: permission }), /布尔标量/)
  assert.deepEqual(draft, staged); assert.equal(controlText(before, 'origin'), 'old')
})

test('empty bindings restore doc-part prompt text or the public zero-width fallback without treating a part name as text', () => {
  const model = pair('old', { placeholder: 'Origin prompt' }, { placeholder: 'PartName', placeholderDocPart: { name: 'PartName', text: 'Your value' } })
  model.blocks.push(paragraph('last-p', 'old')); add(model, 'last', 'last-p', 'old')
  const before = copy(model); controls.resetDocumentContentControlPlaceholder(model, 'origin')
  bindings.synchronizeDocumentContentControlBindings(before, model, { sourceControlId: 'origin', sourceValue: '', permissionContext: permission })
  assert.equal(controlText(model, 'peer'), 'Your value'); assert.equal(controlText(model, 'last'), '\u200b')
  for (const id of ['origin', 'peer', 'last']) { assert.equal(info(model, id).showingPlaceholder, true); assert.equal(info(model, id).dataBinding.cachedResult, '') }
})

test('the nullish binding fallback preserves an explicit empty prompt instead of replacing or rejecting it', () => {
  const model = pair('old', { placeholder: 'Origin prompt' }, { placeholder: '', placeholderDocPart: { name: 'Part', text: 'Unused prompt' } })
  const before = copy(model); controls.resetDocumentContentControlPlaceholder(model, 'origin')
  bindings.synchronizeDocumentContentControlBindings(before, model, { sourceControlId: 'origin', sourceValue: '', permissionContext: permission })
  assert.equal(controlText(model, 'peer'), ''); assert.equal(info(model, 'peer').startOffset, info(model, 'peer').endOffset)
  assert.equal(info(model, 'peer').showingPlaceholder, true); assert.equal(info(model, 'peer').dataBinding.cachedResult, '')
})

test('placeholder reset unwraps actual temporary sources and peers while retaining their prompt text and empty scalar', () => {
  const model = pair('old', { temporary: true, placeholder: 'Origin prompt' }, { temporary: true, placeholder: 'Peer prompt' })
  model.blocks.push(paragraph('last-p', 'old')); add(model, 'last', 'last-p', 'old', { placeholder: 'Last prompt' })
  const before = copy(model); controls.resetDocumentContentControlPlaceholder(model, 'origin')
  bindings.synchronizeDocumentContentControlBindings(before, model, { sourceControlId: 'origin', sourceValue: '', permissionContext: permission })
  assert.equal(info(model, 'origin'), undefined); assert.equal(info(model, 'peer'), undefined)
  assert.equal(models.runText(model.blocks[0].runs), 'Origin prompt'); assert.equal(models.runText(model.blocks[1].runs), 'Peer prompt')
  assert.equal(controlText(model, 'last'), 'Last prompt'); assert.equal(info(model, 'last').showingPlaceholder, true)
  assert.equal(info(model, 'last').dataBinding.cachedResult, ''); assert.equal(info(model, 'last').dataBinding.dirty, true)
})

test('date and dropdown resets keep placeholder text separate from semantic date/list values', () => {
  for (const patch of [
    { type: 'date', value: '2026-11-24', dateValue: '2026-11-24', placeholder: 'Choose date' },
    { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }], placeholder: 'Choose item' },
  ]) {
    const model = pair(patch.type === 'date' ? '2026-11-24' : 'Alpha', patch, { placeholder: 'Peer prompt' })
    const before = copy(model); controls.resetDocumentContentControlPlaceholder(model, 'origin')
    bindings.synchronizeDocumentContentControlBindings(before, model, { sourceControlId: 'origin', sourceValue: '', permissionContext: permission })
    assert.equal(controlText(model, 'origin'), patch.placeholder); assert.equal(info(model, 'origin').showingPlaceholder, true)
    assert.equal(info(model, 'origin').dataBinding.cachedResult, ''); assert.equal(controlText(model, 'peer'), 'Peer prompt')
    if (patch.type === 'date') {
      assert.equal(info(model, 'origin').dateValue, undefined)
      assert.equal(info(model, 'origin').metadata.contentControlLastKnownDate, '2026-11-24')
    }
  }
})

test('resetting an existing prompt leaves revision, caches and a pending redo unchanged', () => {
  const { studio, document } = open(pair('old', { placeholder: 'Prompt' }))
  try {
    assert.equal(document.getContentControl('origin').resetPlaceholder(), true)
    document.getContentControl('origin').setValue('new'); document.undo()
    const before = copy(document.getModel()), revision = studio.getDocumentRevision(document.id), stack = history(studio, document.id)
    assert.equal(document.getContentControl('origin').resetPlaceholder(), false)
    assert.deepEqual(document.getModel(), before); assert.equal(studio.getDocumentRevision(document.id), revision)
    assert.deepEqual(history(studio, document.id), stack); assert.equal(document.redo(), true)
    assert.equal(controlText(document.getModel(), 'origin'), 'new')
  } finally { studio.dispose() }
})

test('a failed picture reset with an invalid imported descendant does not publish structural, field or metadata restoration', () => {
  const model = fixture('Before Choose picture After')
  controls.createDocumentContentControl(model, { id: 'picture', type: 'picture', blockId: 'p', startOffset: 7, endOffset: 21, placeholder: 'Choose picture' })
  controls.replacePictureContentControlImage(model, 'picture', { src: 'data:image/png;base64,AA==', width: 40, height: 20 })
  const picture = info(model, 'picture')
  // Malformed imported nesting must be rejected without deleting otherwise retained picture/tail data.
  model.contentControls.push({ id: 'invalid-child', type: 'richText', blockId: picture.blockId, startOffset: picture.startOffset,
    endOffset: picture.endOffset, parentId: picture.id, lock: 'unlocked' })
  const before = copy(model)
  assert.throws(() => controls.resetDocumentContentControlPlaceholder(model, 'picture'), /nested control/)
  assert.deepEqual(model, before); assert.equal(model.blocks.some(block => block.type === 'image'), true)
})

test('temporary source and temporary first peer do not lose the scalar needed by later peers', () => {
  const model = pair('old', { temporary: true }, { temporary: true })
  model.blocks.push(paragraph('z-p', 'old')); add(model, 'last', 'z-p', 'old')
  setAndBind(model, 'origin', 'changed😀')
  assert.equal(info(model, 'origin'), undefined); assert.equal(info(model, 'peer'), undefined)
  assert.equal(models.runText(model.blocks[1].runs), 'changed😀'); assert.equal(controlText(model, 'last'), 'changed😀')
  assert.equal(info(model, 'last').dataBinding.cachedResult, 'changed😀'); assert.equal(info(model, 'last').dataBinding.dirty, true)
})

test('a later peer content lock or invalid dropdown leaves the caller-owned originating draft unchanged', () => {
  for (const patch of [{ lock: 'contentLocked' }, { type: 'dropDownList', value: 'old', items: [{ value: 'old' }] }]) {
    const before = pair('old', { temporary: true })
    before.blocks.push(paragraph('z-p', 'old')); add(before, 'last', 'z-p', 'old', patch)
    const draft = copy(before); controls.setDocumentContentControlValue(draft, 'origin', 'changed')
    const staged = copy(draft)
    assert.throws(() => bindings.synchronizeDocumentContentControlBindings(before, draft, { sourceControlId: 'origin', sourceValue: 'changed',
      permissionContext: permission }), /locked|dropdown/)
    assert.deepEqual(draft, staged); assert.equal(controlText(draft, 'peer'), 'old'); assert.equal(info(draft, 'peer').dataBinding.dirty, false)
  }
})

test('later cross-host permission and ancestor locks reject the whole facade transaction including source unwrapping and caches', () => {
  for (const scenario of ['permission', 'ancestor']) {
    const model = hostedFixture(); info(model, 'body').temporary = true
    if (scenario === 'permission') model.permissionRules = [{ id: 'deny-end', scope: 'entity', targetId: 'endnote:end', action: 'edit', effect: 'deny', userIds: ['editor'] }]
    else {
      controls.createDocumentContentControl(model, { id: 'group', type: 'group', blockId: 'zz-endnote-p', startOffset: 0, endOffset: 3 })
      controls.updateDocumentContentControl(model, 'end', { parentId: 'group' }); controls.updateDocumentContentControl(model, 'group', { lock: 'contentLocked' })
    }
    const { studio, document } = open(model), before = copy(document.getModel()), events = []
    const subscription = studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') events.push(event.id) })
    try {
      assert.throws(() => document.getContentControl('body').setValue('changed'), /permission denied|locked/)
      assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); assert.deepEqual(events, [])
    } finally { subscription.dispose(); studio.dispose() }
  }
})

test('a source paragraph denial blocks metadata-only list value changes with an unchanged display label', () => {
  const model = pair('Same', { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'Same' }, { value: 'B', displayText: 'Same' }] })
  model.permissionRules = [{ id: 'deny-source', scope: 'paragraph', targetId: 'p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  const { studio, document } = open(model), before = copy(document.getModel())
  try {
    assert.throws(() => document.getContentControl('origin').setValue('B'), /permission denied/)
    assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 })
  } finally { studio.dispose() }
})

test('same-value commands preserve revision, binding dirtiness and the pending redo entry', () => {
  const { studio, document } = open(pair())
  try {
    document.getContentControl('origin').setValue('new'); document.undo()
    const before = copy(document.getModel()), revision = studio.getDocumentRevision(document.id), stack = history(studio, document.id)
    assert.equal(document.getContentControl('origin').setValue('old'), false)
    assert.equal(studio.getDocumentRevision(document.id), revision); assert.deepEqual(history(studio, document.id), stack)
    assert.deepEqual(document.getModel(), before); assert.equal(info(document.getModel(), 'origin').dataBinding.dirty, false)
    assert.equal(document.redo(), true); assert.equal(controlText(document.getModel(), 'peer'), 'new')
  } finally { studio.dispose() }
})

test('same-paragraph peers before a persistent or temporary source map the actual UTF-16 caret', () => {
  for (const temporary of [false, true]) {
    const model = fixture('A / A / A')
    for (const [id, offset] of [['first', 0], ['second', 4], ['origin', 8]]) add(model, id, 'p', 'A', { startOffset: offset,
      endOffset: offset + 1, temporary: id === 'origin' && temporary })
    const before = copy(model), value = '中😀'
    const result = editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection('p', 8, 9),
      command: { id: 'doc.command.insert-text', text: value, activeContentControlId: 'origin' } }, permission)
    assert.equal(models.runText(model.blocks[0].runs), value + ' / ' + value + ' / ' + value)
    assert.equal(result.focus.offset, 15); assert.equal(result.anchor.offset, 15)
    if (temporary) {
      assert.equal(info(model, 'origin'), undefined)
      assert.deepEqual(bindings.documentContentControlTemporaryBindingCaret(before, model, 'origin', value), selection('p', 15))
    } else assert.equal(info(model, 'origin').endOffset, 15)
  }
})

test('temporary caret reconstruction reads actual surviving peer displays across different binding keys', () => {
  const before = fixture('A / A / A')
  for (const [id, offset, xpath] of [['first', 0, '/first'], ['second', 4, '/second'], ['origin', 8, '/origin']]) {
    add(before, id, 'p', 'A', { startOffset: offset, endOffset: offset + 1, temporary: id === 'origin', dataBinding: binding(xpath) })
  }
  const snapshot = copy(before), after = copy(before), value = '中😀'
  controls.setDocumentContentControlValue(after, 'origin', value)
  controls.setDocumentContentControlValue(after, 'first', 'P😀')
  controls.setDocumentContentControlValue(after, 'second', 'Longer')
  const expected = 'P😀 / Longer / 中😀'
  assert.equal(models.runText(after.blocks[0].runs), expected)
  assert.deepEqual(bindings.documentContentControlTemporaryBindingCaret(before, after, 'origin', value), selection('p', expected.length))
  assert.deepEqual(before, snapshot)
  const unrelated = copy(after); unrelated.blocks[0].runs[0].text = 'Changed ' + unrelated.blocks[0].runs[0].text
  assert.equal(bindings.documentContentControlTemporaryBindingCaret(before, unrelated, 'origin', value), undefined)
})

test('conflicting same-key values from one replace-all batch roll back every originating edit', () => {
  const model = fixture('A1'); model.blocks.push(paragraph('peer-p', 'A2'))
  add(model, 'origin', 'p', 'A1'); add(model, 'peer', 'peer-p', 'A2')
  const before = copy(model), matches = search.findDocumentMatches(model, 'A')
  assert.throws(() => search.replaceDocumentMatches(model, matches, 'B', permission), /冲突的绑定标量/)
  assert.deepEqual(model, before)
})

test('whole-cell input obtains a temporary multiline scalar from the actual target while the UI caret is in another paragraph', () => {
  const model = hostedFixture(); info(model, 'cell').temporary = true
  const value = 'First\nSecond'
  const result = editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection('body-p', 1),
    command: { id: 'doc.command.table-set-cell-text', tableId: 'table', row: 0, column: 0, text: value } }, permission)
  assert.equal(info(model, 'cell'), undefined); assertHosted(model, value, ['cell'])
  assert.equal(model.blocks[1].metadata.table.rows[0].cells[0].blocks.map(block => models.runText(block.runs)).join('\n'), value)
  assert.equal(result.focus.blockId, 'table'); assert.equal(result.focus.offset, 'Second'.length)
})

test('an explicit text-control identity on a whole-cell command rejects before changing either target', () => {
  const model = hostedFixture(), before = copy(model)
  assert.throws(() => editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection('body-p', 1),
    command: { id: 'doc.command.table-set-cell-text', tableId: 'table', row: 0, column: 0, text: 'new', activeContentControlId: 'body' } }, permission), /整格输入/)
  assert.deepEqual(model, before)
})

test('an empty temporary cell retains its actual multiline source value after the structural planner unwraps it', () => {
  const model = hostedFixture()
  model.blocks[1].metadata.table.rows[0].cells[0].blocks[0].runs[0].text = ''
  Object.assign(info(model, 'cell'), { startOffset: 0, endOffset: 0, value: '', temporary: true })
  const value = 'First\nSecond'
  editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection('body-p', 1),
    command: { id: 'doc.command.table-set-cell-text', tableId: 'table', row: 0, column: 0, text: value } }, permission)
  assert.equal(info(model, 'cell'), undefined); assertHosted(model, value, ['cell'])
})

test('matrix text records multiple disjoint cell sources and synchronizes independent bindings after early temporary unwrapping', () => {
  const model = fixture('outside')
  model.blocks.push(paragraph('a-peer', 'A'), paragraph('b-peer', 'B'))
  model.blocks.push({ ...paragraph('table', ''), type: 'table', runs: [], metadata: { table: { rows: [{ id: 'row', cells:
    [['a-cell', 'A'], ['b-cell', 'B']].map(([id, value]) => ({ id, blocks: [paragraph(id + '-p', value)], style: {}, rowSpan: 1, colSpan: 1 })) }] } } })
  for (const [id, blockId, value, xpath, temporary] of [['a-origin', 'a-cell-p', 'A', '/a', true], ['b-origin', 'b-cell-p', 'B', '/b', true],
    ['a-peer', 'a-peer', 'A', '/a', false], ['b-peer', 'b-peer', 'B', '/b', false]]) add(model, id, blockId, value, { temporary, dataBinding: binding(xpath) })
  editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection('p', 2), command: { id: 'doc.command.table-set-range-text',
    tableId: 'table', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }, mode: 'matrix', values: [['AA\nA', 'BB\nB']] } }, permission)
  assert.equal(info(model, 'a-origin'), undefined); assert.equal(info(model, 'b-origin'), undefined)
  assert.equal(controlText(model, 'a-peer'), 'AA\nA'); assert.equal(controlText(model, 'b-peer'), 'BB\nB')
  assert.equal(info(model, 'a-peer').dataBinding.cachedResult, 'AA\nA'); assert.equal(info(model, 'b-peer').dataBinding.cachedResult, 'BB\nB')
})

test('mixed object-only, empty and bound text cells preserve structural insertion order and independent temporary source values', () => {
  const model = fixture('outside'); model.blocks.push(paragraph('a-peer-p', 'A'), paragraph('b-peer-p', 'B'))
  const image = id => ({ ...paragraph(id, ''), type: 'image', runs: [], metadata: { image: { src: 'data:image/png;base64,AA==', width: 10, height: 10 } } })
  const definitions = [['left', [image('left-image')]], ['a', [paragraph('a-p', 'A')]], ['middle', []],
    ['b', [paragraph('b-p', 'B')]], ['right', [image('right-image')]]]
  model.blocks.push({ ...paragraph('table', ''), type: 'table', runs: [], metadata: { table: { rows: [{ id: 'row',
    cells: definitions.map(([id, blocks]) => ({ id, blocks, style: {}, rowSpan: 1, colSpan: 1 })) }] } } })
  for (const [id, blockId, value, xpath, temporary] of [['a-origin', 'a-p', 'A', '/a', true], ['b-origin', 'b-p', 'B', '/b', true],
    ['a-peer', 'a-peer-p', 'A', '/a', false], ['b-peer', 'b-peer-p', 'B', '/b', false]]) add(model, id, blockId, value, { temporary, dataBinding: binding(xpath) })
  const values = [['Left\n😀', 'A😀\nA', 'Middle', 'BB\nB', 'Right']]
  editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection('p', 2), command: { id: 'doc.command.table-set-range-text',
    tableId: 'table', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 }, mode: 'matrix', values } }, permission)
  assert.equal(info(model, 'a-origin'), undefined); assert.equal(info(model, 'b-origin'), undefined)
  assert.equal(controlText(model, 'a-peer'), 'A😀\nA'); assert.equal(controlText(model, 'b-peer'), 'BB\nB')
  const current = model.blocks.find(block => block.id === 'table').metadata.table
  assert.deepEqual(current.rows[0].cells.map(cell => cell.blocks.map(block => models.runText(block.runs)).join('\n')), values[0])
  assert.equal(info(model, 'a-peer').dataBinding.cachedResult, values[0][1]); assert.equal(info(model, 'b-peer').dataBinding.cachedResult, values[0][3])
})

test('source cells at a shared offset remain distinguishable when a neighboring bound paragraph is empty', () => {
  const model = fixture('outside'); model.blocks.push(paragraph('peer-p', 'old'))
  model.blocks.push({ ...paragraph('table', ''), type: 'table', runs: [], metadata: { table: { rows: [{ id: 'row', cells: [
    { id: 'before', blocks: [], style: {}, rowSpan: 1, colSpan: 1 },
    { id: 'origin-cell', blocks: [paragraph('origin-p', '')], style: {}, rowSpan: 1, colSpan: 1 },
    { id: 'after', blocks: [], style: {}, rowSpan: 1, colSpan: 1 },
  ] }] } } })
  add(model, 'origin', 'origin-p', '', { temporary: true }); add(model, 'peer', 'peer-p', 'old')
  editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection('p', 0), command: { id: 'doc.command.table-set-range-text',
    tableId: 'table', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 }, mode: 'matrix', values: [['Left', 'Actual\n😀', 'Right']] } }, permission)
  assert.equal(info(model, 'origin'), undefined); assert.equal(controlText(model, 'peer'), 'Actual\n😀')
  assert.equal(info(model, 'peer').dataBinding.cachedResult, 'Actual\n😀')
})

test('a source protocol for cells with no text paragraphs retains cell identity, delimiters and validation without publishing', () => {
  const before = fixture(); before.blocks = [{ ...paragraph('table', ''), type: 'table', runs: [], metadata: { table: { rows: [{ id: 'row', cells: [
    { id: 'a', blocks: [], style: {}, rowSpan: 1, colSpan: 1 }, { id: 'b', blocks: [], style: {}, rowSpan: 1, colSpan: 1 },
  ] }] } } }]
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }, command = { id: 'doc.command.table-set-range-text', tableId: 'table', range,
    mode: 'matrix', values: [['A😀', 'B\nB']] }, snapshot = copy(before)
  const edits = tableMutations.documentTableTextInputEdits(before, command, { blockId: 'table', range })
  assert.deepEqual(tableOrigins.documentTableTextOriginCells(before).map(origin => [origin.cell.id, origin.order, origin.paragraphIndex]), [['a', 0, 0], ['b', 1, 0]])
  assert.deepEqual(edits.map(edit => [edit.startOffset, edit.endOffset, edit.text, edit.tableCell.cellId, edit.tableCell.value]),
    [[0, 0, 'A😀\n', 'a', 'A😀'], [0, 0, 'B\nB', 'b', 'B\nB']])
  assert.deepEqual(before, snapshot)
  const after = copy(before)
  after.blocks[0].metadata.table.rows[0].cells[0].blocks = [paragraph('a-p', 'A😀')]
  after.blocks[0].metadata.table.rows[0].cells[1].blocks = [paragraph('b-p1', 'B'), paragraph('b-p2', 'B')]
  assert.deepEqual(bindings.documentContentControlBindingSources(before, after, { originTextEdits: edits }), [])
  for (const patch of [{ tableCell: { ...edits[0].tableCell, cellId: 'wrong' } }, { tableCell: { ...edits[0].tableCell, order: 1 } },
    { tableCell: { ...edits[0].tableCell, value: 'wrong' } }, { text: 'Wrong delimiter' }]) {
    const invalid = [{ ...edits[0], ...patch }, edits[1]]
    assert.throws(() => bindings.documentContentControlBindingSources(before, after, { originTextEdits: invalid }), /身份|实际值|分隔符/)
    assert.deepEqual(before, snapshot)
  }
})

test('materializing a non-text cell uses the shared cross-paragraph field cache and lock contract', () => {
  for (const locked of [false, true]) {
    const model = fixture('Start')
    const image = { ...paragraph('image', ''), type: 'image', runs: [], metadata: { image: { src: 'data:image/png;base64,AA==', width: 10, height: 10 } } }
    model.blocks.push({ ...paragraph('table', ''), type: 'table', runs: [], metadata: { table: { rows: [{ id: 'row', cells: [
      { id: 'object', blocks: [image], style: {}, rowSpan: 1, colSpan: 1 },
      { id: 'source', blocks: [paragraph('source-p', 'A')], style: {}, rowSpan: 1, colSpan: 1 },
    ] }] } } }, paragraph('end-p', 'End'), paragraph('peer-p', 'A'))
    add(model, 'origin', 'source-p', 'A', { temporary: true }); add(model, 'peer', 'peer-p', 'A')
    fields.setDocumentTextFields(model.blocks[0], [{ id: 'cross', field: 'MERGEFIELD', instruction: 'MERGEFIELD value', startOffset: 0,
      endOffset: 3, endBlockId: 'end-p', cachedResult: 'Start\nA\nEnd', locked }])
    const before = copy(model), mutate = () => editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection('p', 0),
      command: { id: 'doc.command.table-set-cell-text', tableId: 'table', row: 0, column: 0, text: 'Inserted' } }, permission)
    if (locked) {
      assert.throws(mutate, /锁定字段/); assert.deepEqual(model, before)
    } else {
      mutate(); assert.equal(fields.documentTextFields(model.blocks[0])[0].cachedResult, 'Start\nInserted\nA\nEnd')
      assert.equal(controlText(model, 'origin'), 'A'); assert.equal(info(model, 'origin').temporary, true); assert.equal(controlText(model, 'peer'), 'A')
    }
  }
})

test('clearing a temporary cell source propagates an empty scalar while preserving the table selection', () => {
  const model = hostedFixture(); info(model, 'cell').temporary = true
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
  const source = { ...selection('table', 0, 3, { row: 0, col: 0, cellBlockId: 'cell-p' }),
    tableSelection: { tableId: 'table', kind: 'cell', range } }
  const result = editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: source,
    command: { id: 'doc.command.table-set-range-text', tableId: 'table', range, mode: 'clear' } }, permission)
  assert.equal(info(model, 'cell'), undefined); assert.equal(controlText(model, 'body'), '\u200b')
  assert.equal(info(model, 'body').showingPlaceholder, true); assert.equal(info(model, 'body').dataBinding.cachedResult, '')
  assert.equal(result.focus.blockId, 'table'); assert.equal(result.focus.cellBlockId, 'cell-p'); assert.equal(result.focus.offset, 0)
  assert.deepEqual(result.tableSelection, source.tableSelection)
})

test('invalid UTF-16, overlapping or mismatched source edit plans fail before any peer is published', () => {
  const cases = [
    { value: '😀', edits: [{ startOffset: 1, endOffset: 2, text: 'x' }] },
    { value: 'AB', edits: [{ startOffset: 0, endOffset: 2, text: 'x' }, { startOffset: 1, endOffset: 2, text: 'y' }] },
    { value: 'AB', edits: [{ startOffset: 0, endOffset: 2, text: 'wrong' }] },
  ]
  for (const sample of cases) {
    const before = pair(sample.value, { temporary: true }), draft = copy(before)
    draft.blocks[0].runs[0].text = 'new\nvalue'; draft.contentControls = draft.contentControls.filter(control => control.id !== 'origin')
    const staged = copy(draft)
    assert.throws(() => bindings.synchronizeDocumentContentControlBindings(before, draft, { originTextEdits: sample.edits, permissionContext: permission }), /UTF-16|重叠|不一致/)
    assert.deepEqual(draft, staged); assert.equal(controlText(draft, 'peer'), sample.value)
  }
})

test('multiple original-coordinate edits recover independent temporary scalars in one paragraph without shifting the later source twice', () => {
  const before = fixture('AB--CD'); before.blocks.push(paragraph('a-peer-p', 'AB'), paragraph('b-peer-p', 'CD'))
  for (const [id, blockId, start, end, value, xpath, temporary] of [
    ['a-origin', 'p', 0, 2, 'AB', '/a', true], ['b-origin', 'p', 4, 6, 'CD', '/b', true],
    ['a-peer', 'a-peer-p', 0, 2, 'AB', '/a', false], ['b-peer', 'b-peer-p', 0, 2, 'CD', '/b', false],
  ]) add(before, id, blockId, value, { startOffset: start, endOffset: end, temporary, dataBinding: binding(xpath) })
  const draft = copy(before)
  draft.blocks[0].runs[0].text = 'A😀\nA--B\nB'
  draft.contentControls = draft.contentControls.filter(control => !control.temporary)
  const edits = [{ startOffset: 0, endOffset: 2, text: 'A😀\nA' }, { startOffset: 4, endOffset: 6, text: 'B\nB' }]
  bindings.synchronizeDocumentContentControlBindings(before, draft, { originTextEdits: edits, permissionContext: permission })
  assert.equal(controlText(draft, 'a-peer'), 'A😀\nA'); assert.equal(controlText(draft, 'b-peer'), 'B\nB')
  assert.equal(info(draft, 'a-peer').dataBinding.cachedResult, 'A😀\nA'); assert.equal(info(draft, 'b-peer').dataBinding.cachedResult, 'B\nB')
})

test('replace-all retains a temporary source scalar across inserted paragraphs and updates a footer peer', () => {
  const model = hostedFixture(); info(model, 'body').temporary = true
  const matches = search.findDocumentMatches(model, 'old', { selection: selection('body-p', 0, 3) })
  search.replaceDocumentMatches(model, matches, 'First\nSecond', permission)
  assert.equal(info(model, 'body'), undefined); assertHosted(model, 'First\nSecond', ['body'])
})

test('search and linked value edits keep authoritative code source text in step with its formatted runs', () => {
  const model = pair('AA')
  for (const block of model.blocks) block.metadata.codeBlock = { source: 'AA', language: 'typescript', lineNumbers: false }
  search.replaceDocumentMatches(model, search.findDocumentMatches(model, 'AA', { selection: selection('p', 0, 2) }), 'BB', permission)
  assert.equal(code.codeBlockData(model.blocks[0]).source, 'BB'); assert.equal(code.codeBlockData(model.blocks[1]).source, 'BB')
  assert.equal(controlText(model, 'peer'), 'BB')
})

test('full-model code updates map the source selection and synchronize peers outside its header host', () => {
  const model = hostedFixture(), header = model.sectionHeaderFooters.first.storyContent.header.blocks[0], value = 'header😀'
  header.metadata.codeBlock = { source: 'old', language: 'plain', lineNumbers: false }
  const result = code.updateCodeBlockWithSelection(model, 'header-p', { source: value }, selection('header-p', 1), permission)
  assert.equal(result.changed, true); assert.equal(result.selection.focus.offset, value.length)
  assertHosted(model, value); assert.equal(code.codeBlockData(model.sectionHeaderFooters.first.storyContent.header.blocks[0]).source, value)
})

test('an unbound code update still rejects a denied source paragraph before publishing', () => {
  const model = fixture('old'); model.blocks[0].metadata.codeBlock = { source: 'old', language: 'plain', lineNumbers: false }
  model.permissionRules = [{ id: 'deny-code', scope: 'paragraph', targetId: 'p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  const before = copy(model)
  assert.throws(() => code.updateCodeBlockWithSelection(model, 'p', { source: 'new' }, undefined, permission), /permission denied/)
  assert.deepEqual(model, before)
})

test('peer replacement maps comment and native-field anchors and keeps code metadata aligned', () => {
  const model = fixture('A'); model.blocks.push(paragraph('peer-p', 'LAR'))
  add(model, 'origin', 'p', 'A'); add(model, 'peer', 'peer-p', 'A', { startOffset: 1, endOffset: 2 })
  const peer = model.blocks.find(block => block.id === 'peer-p')
  peer.metadata.codeBlock = { source: 'LAR', language: 'plain', lineNumbers: true }
  fields.setDocumentTextFields(peer, [{ id: 'field', field: 'DATE', instruction: 'DATE', cachedResult: 'A', startOffset: 1, endOffset: 2 }])
  const thread = comments.createDocumentThreadComment(model, selection('peer-p', 1, 2), 'Review',
    { documentId: model.id, userId: 'editor', displayName: 'Editor' }, '2026-11-24T00:00:00.000Z')
  setAndBind(model, 'origin', 'Beta')
  const current = model.blocks.find(block => block.id === 'peer-p')
  assert.equal(models.runText(current.runs), 'LBetaR'); assert.equal(code.codeBlockData(current).source, 'LBetaR')
  assert.deepEqual(fields.documentTextFields(current).map(field => [field.id, field.startOffset, field.endOffset]), [['field', 1, 5]])
  assert.deepEqual(model.comments[thread.id].anchor.map(anchor => [anchor.startOffset, anchor.endOffset, anchor.text]), [[1, 5, 'Beta']])
  assert.equal(model.comments[thread.id].referenceText, 'Beta')
})

test('one facade transaction includes source unwrapping and every host peer through undo, redo and internal save/reopen', () => {
  const model = hostedFixture(); info(model, 'body').temporary = true
  const { studio, document } = open(model), before = copy(document.getModel()); let reopened
  try {
    assert.equal(document.getContentControl('body').setValue('new😀'), true)
    assert.equal(info(document.getModel(), 'body'), undefined); assertHosted(document.getModel(), 'new😀', ['body'])
    assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0 })
    const changed = copy(document.getModel().contentControls)
    document.undo(); assert.deepEqual(document.getModel().contentControls, before.contentControls)
    document.redo(); assert.deepEqual(document.getModel().contentControls, changed); assertHosted(document.getModel(), 'new😀', ['body'])
    reopened = open(copy(document.saveModel()))
    assert.deepEqual(reopened.document.getModel().contentControls, changed); assertHosted(reopened.document.getModel(), 'new😀', ['body'])
    assert.deepEqual(history(reopened.studio, document.id), { undo: 0, redo: 0 })
  } finally { studio.dispose(); reopened?.studio.dispose() }
})
