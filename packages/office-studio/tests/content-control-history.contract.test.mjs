import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred DOC-11 API/history integration drafts. Internal snapshots are not DOCX roundtrip evidence.
let server, api, modelUtils, controls, commandState
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-history', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  commandState = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/content-control-command-state.ts')
})
afterAll(async () => { await server?.close() })

// Active facade models may contain Vue proxies; use the project's JSON snapshot contract.
const copy = value => JSON.parse(JSON.stringify(value))
function fixture(text = '_', control = { type: 'checkbox' }, id = 'sdt-history') {
  const model = api.createBlankDocument(); model.id = id; model.blocks[0].id = 'p'; model.blocks[0].runs[0].text = text; model.contentControls = []
  controls.createDocumentContentControl(model, { id: 'control', blockId: 'p', startOffset: 0, endOffset: text.length,
    rawSdtPr: '<w:sdtPr><w:tag w:val="retain"/></w:sdtPr>', metadata: { originalSource: { version: 1 } }, ...control })
  return model
}
function open(model) {
  const studio = new api.OfficeStudio({ document: model, currentUserId: 'editor' })
  studio.activateDocument(model.id)
  const document = new api.FDocument(studio, model.id)
  return { studio, document }
}
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })

test('checkbox, date and list commands each publish one transaction and survive undo, redo and snapshot reopen', () => {
  for (const sample of [
    { text: '_', info: { type: 'checkbox' }, id: 'doc.command.set-content-control-checked', params: { checked: true }, expected: '☒' },
    { text: '2026-09-03', info: { type: 'date', dateValue: '2026-09-03', value: '2026-09-03', dateFormat: 'dd MMM yyyy' }, id: 'doc.command.set-content-control-value', params: { value: '2026-09-04' }, expected: '04 Sep 2026' },
    { text: 'Alpha', info: { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Beta' }] }, id: 'doc.command.select-content-control-item', params: { value: 'B' }, expected: 'Beta' },
  ]) {
    const { studio, document } = open(fixture(sample.text, sample.info)), before = copy(document.getModel()), events = []
    const subscription = studio.events.addEvent('CommandExecuted', event => {
      if (event.unitId === document.id && event.kind === 'document' && event.id.startsWith('doc.command.')) events.push(event.id)
    })
    let reopened
    try {
      assert.equal(document.executeContentControlCommand(sample.id, { controlId: 'control', ...sample.params }), true)
      assert.equal(document.getText(), sample.expected)
      assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0 }); assert.deepEqual(events, [sample.id])
      const changed = copy(document.getContentControl('control').getInfo())
      assert.equal(document.undo(), true); assert.deepEqual(document.getModel().contentControls, before.contentControls)
      assert.equal(document.redo(), true); assert.deepEqual(document.getContentControl('control').getInfo(), changed)
      const saved = document.saveModel(); reopened = open(copy(saved))
      assert.equal(reopened.document.getText(), sample.expected)
      assert.deepEqual(reopened.document.getContentControl('control').getInfo(), changed)
      assert.deepEqual(history(reopened.studio, document.id), { undo: 0, redo: 0 })
    } finally { subscription.dispose(); studio.dispose(); reopened?.studio.dispose() }
  }
})

test('command cancellation and invalid values retain text, model, history and after-command events', () => {
  const { studio, document } = open(fixture()), before = copy(document.getModel()), events = []
  const after = studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  const cancellation = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.set-content-control-checked') event.cancel = true })
  try {
    assert.throws(() => document.executeContentControlCommand('doc.command.set-content-control-checked', { controlId: 'control', checked: true }), /canceled/)
    assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); assert.deepEqual(events, [])
    cancellation.dispose()
    assert.throws(() => document.executeContentControlCommand('doc.command.update-content-control', { controlId: 'control', patch: { checkbox: { checked: false, uncheckedState: { value: 'D800' } } } }), /Unicode/)
    assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); assert.deepEqual(events, [])
  } finally { cancellation.dispose(); after.dispose(); studio.dispose() }
})

test('paragraph permission denial applies to the facade and Canvas capability without recording a transaction', () => {
  const model = fixture()
  model.permissionRules = [{ id: 'deny-control', scope: 'paragraph', targetId: 'p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  const { studio, document } = open(model), before = copy(document.getModel())
  try {
    const command = { id: 'doc.command.set-content-control-checked', params: { controlId: 'control', checked: true } }
    assert.equal(commandState.contentControlValueCommandState(document.getModel(), command, { permission: { documentId: document.id, userId: 'editor' } }).available, false)
    assert.throws(() => document.executeContentControlCommand(command.id, command.params), /permission denied/)
    assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 })
  } finally { studio.dispose() }
})

test('two documents with the same content-control ID retain independent content histories', () => {
  const { studio, document: first } = open(fixture('_', { type: 'checkbox' }, 'first-doc'))
  const office = api.FOffice.newAPI(studio), second = office.createDocument(fixture('_', { type: 'checkbox' }, 'second-doc'))
  try {
    first.getContentControl('control').setChecked(true); second.getContentControl('control').setChecked(true)
    assert.deepEqual(history(studio, first.id), { undo: 1, redo: 0 }); assert.deepEqual(history(studio, second.id), { undo: 1, redo: 0 })
    first.undo()
    assert.equal(first.getText(), '☐'); assert.equal(second.getText(), '☒')
    assert.deepEqual(history(studio, first.id), { undo: 0, redo: 1 }); assert.deepEqual(history(studio, second.id), { undo: 1, redo: 0 })
  } finally { studio.dispose() }
})

function content(text, id) {
  const paragraph = modelUtils.createParagraph(text); paragraph.id = id
  return { blocks: [paragraph], runs: copy(paragraph.runs), paragraphs: [{ startOffset: 0, endOffset: text.length, style: {} }] }
}
function hostedFixture() {
  const model = api.createBlankDocument(); model.id = 'hosted-sdt'; model.blocks[0].id = 'body'; model.blocks[0].runs[0].text = 'BODY'; model.contentControls = []
  model.sectionHeaderFooters = { first: { header: '_', footer: '', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '',
    differentFirstPage: false, differentOddEven: false, linkToPrevious: false, storyContent: { header: content('_', 'header-p') } } }
  model.footnotes = { note: { id: 'note', text: 'Alpha', referenceBlockId: 'body', referenceRunId: model.blocks[0].runs[0].id, content: content('Alpha', 'note-p') } }
  controls.createDocumentContentControl(model, { id: 'header-control', type: 'checkbox', blockId: 'header-p', startOffset: 0, endOffset: 1, rawSdtPr: '<w:sdtPr/>' })
  controls.createDocumentContentControl(model, { id: 'note-control', type: 'dropDownList', blockId: 'note-p', startOffset: 0, endOffset: 5, value: 'A',
    items: [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Beta' }] })
  return model
}

test('header and footnote controls reuse commands and document history while leaving the body intact', () => {
  const { studio, document } = open(hostedFixture()), body = copy(document.getModel().blocks)
  let reopened
  try {
    document.getContentControl('header-control').setChecked(true)
    document.getContentControl('note-control').selectItem('B')
    assert.deepEqual(document.getModel().blocks, body)
    assert.equal(document.getModel().sectionHeaderFooters.first.header, '☒')
    assert.equal(document.getModel().footnotes.note.text, 'Beta')
    assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0 })
    document.undo(); assert.equal(document.getModel().footnotes.note.text, 'Alpha'); assert.equal(document.getModel().sectionHeaderFooters.first.header, '☒')
    document.undo(); assert.equal(document.getModel().sectionHeaderFooters.first.header, '☐')
    document.redo(); document.redo()
    reopened = open(copy(document.saveModel()))
    assert.equal(reopened.document.getContentControl('header-control').getText(), '☒')
    assert.equal(reopened.document.getContentControl('note-control').getText(), 'Beta')
    assert.deepEqual(reopened.document.getModel().blocks, body)
    assert.equal(reopened.document.getModel().metadata?.documentEditingBlockHost, undefined)
  } finally { studio.dispose(); reopened?.studio.dispose() }
})

test('a note-host entity denial blocks its Canvas capability and facade edit atomically', () => {
  const model = hostedFixture()
  model.permissionRules = [{ id: 'deny-note', scope: 'entity', targetId: 'footnote:note', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  const { studio, document } = open(model), before = copy(document.getModel())
  try {
    const command = { id: 'doc.command.select-content-control-item', params: { controlId: 'note-control', value: 'B' } }
    assert.equal(commandState.contentControlValueCommandState(document.getModel(), command, { host: { kind: 'note', noteKind: 'footnote', noteId: 'note' },
      permission: { documentId: document.id, userId: 'editor' } }).available, false)
    assert.throws(() => document.executeContentControlCommand(command.id, command.params), /permission denied/)
    assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 })
  } finally { studio.dispose() }
})

const textSelection = (start, end = start, target = { blockId: 'p' }) => ({ anchor: { ...target, offset: start }, focus: { ...target, offset: end } })

test('ordinary input and temporary unwrapping share one undo record and persist the remaining text on snapshot reopen', () => {
  const { studio, document } = open(fixture('a😀b', { type: 'richText', temporary: true, value: 'a😀b' }))
  const before = copy(document.getModel()), events = []
  const subscription = studio.events.addEvent('CommandExecuted', event => { if (event.unitId === document.id && event.id === 'doc.command.insert-text') events.push(event.id) })
  let reopened
  try {
    assert.equal(document.executeCommand({ id: 'doc.command.insert-text', text: '中' }, textSelection(1, 3)).changed, true)
    assert.equal(document.getText(), 'a中b'); assert.deepEqual(document.getModel().contentControls, [])
    assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0 }); assert.deepEqual(events, ['doc.command.insert-text'])
    assert.equal(document.undo(), true); assert.deepEqual(document.getModel().contentControls, before.contentControls); assert.equal(document.getText(), 'a😀b')
    assert.equal(document.redo(), true); assert.deepEqual(document.getModel().contentControls, [])
    reopened = open(copy(document.saveModel())); assert.equal(reopened.document.getText(), 'a中b')
    assert.deepEqual(reopened.document.getModel().contentControls, []); assert.deepEqual(history(reopened.studio, document.id), { undo: 0, redo: 0 })
  } finally { subscription.dispose(); studio.dispose(); reopened?.studio.dispose() }
})

test('ordinary caret input replaces a placeholder once and restores placeholder state on undo', () => {
  const { studio, document } = open(fixture('Name', { type: 'text', value: 'Name', placeholder: 'Name', showingPlaceholder: true, dataBinding: { xpath: '/name', dirty: false } }))
  const before = copy(document.getModel())
  try {
    const result = document.executeCommand({ id: 'doc.command.insert-text', text: '李😀' }, textSelection(2))
    assert.equal(result.changed, true); assert.equal(document.getText(), '李😀')
    const changed = document.getContentControl('control').getInfo()
    assert.equal(changed.value, '李😀'); assert.equal(changed.showingPlaceholder, false); assert.equal(changed.dataBinding.dirty, true)
    assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0 })
    assert.equal(document.undo(), true); assert.deepEqual(document.getModel().contentControls, before.contentControls)
    assert.equal(document.redo(), true); assert.equal(document.getContentControl('control').getInfo().value, '李😀')
  } finally { studio.dispose() }
})

test('same explicit values keep revision, savepoint, redo history and temporary unchanged', () => {
  const { studio, document } = open(fixture('Name', { type: 'text', value: 'Name', temporary: true, dataBinding: { dirty: false } }))
  try {
    document.getContentControl('control').update({ alias: 'changed title' }); document.undo()
    const before = copy(document.getModel()), revision = document.getModelRevision(), generation = document.getHistoryGeneration(), stack = history(studio, document.id)
    assert.equal(document.executeContentControlCommand('doc.command.set-content-control-value', { controlId: 'control', value: 'Name' }), false)
    assert.deepEqual(document.getModel(), before); assert.equal(document.getModelRevision(), revision); assert.equal(document.getHistoryGeneration(), generation)
    assert.deepEqual(history(studio, document.id), stack); assert.equal(document.getContentControl('control').getInfo().temporary, true)
    assert.equal(document.redo(), true); assert.equal(document.getContentControl('control').getInfo().alias, 'changed title')
  } finally { studio.dispose() }
})

test('temporary paragraph splitting, typed dates and combo values use the existing editing command history', () => {
  for (const sample of [
    { text: 'abcd', properties: { type: 'richText', temporary: true, value: 'abcd' }, selection: textSelection(2), input: '\n', expected: 'ab\ncd', removed: true },
    { text: '03 Sep 2026', properties: { type: 'date', value: '2026-09-03', dateValue: '2026-09-03', dateFormat: 'dd MMM yyyy' }, selection: textSelection(0, 11), input: '04 Sep 2026', expected: '04 Sep 2026', value: '2026-09-04' },
    { text: 'Free', properties: { type: 'comboBox', value: 'Free', items: [{ value: 'A', displayText: 'Alpha' }] }, selection: textSelection(0, 4), input: 'Alpha', expected: 'Alpha', value: 'Alpha' },
  ]) {
    const { studio, document } = open(fixture(sample.text, sample.properties)), before = copy(document.getModel())
    try {
      assert.equal(document.executeCommand({ id: 'doc.command.insert-text', text: sample.input }, sample.selection).changed, true)
      assert.equal(document.getText(), sample.expected); assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0 })
      if (sample.removed) assert.deepEqual(document.getModel().contentControls, [])
      else assert.equal(document.getContentControl('control').getInfo().value, sample.value)
      document.undo(); assert.deepEqual(document.getModel().contentControls, before.contentControls); assert.equal(document.getText(), sample.text)
      document.redo(); assert.equal(document.getText(), sample.expected)
    } finally { studio.dispose() }
  }
})

test('ordinary header and note input synchronizes controls and compatibility projections in document history', () => {
  const model = hostedFixture(), note = model.contentControls.find(control => control.id === 'note-control')
  note.type = 'comboBox'; note.value = 'Alpha'
  const { studio, document } = open(model), body = copy(document.getModel().blocks)
  try {
    document.executeCommand({ id: 'doc.command.insert-text', text: '☒', marks: { fontFamily: 'Typing Font' }, host: { kind: 'story', sectionId: 'first', story: 'header' } },
      textSelection(0, 1, { blockId: 'header-p' }))
    document.executeCommand({ id: 'doc.command.insert-text', text: 'typed note', host: { kind: 'note', noteKind: 'footnote', noteId: 'note' } },
      textSelection(0, 5, { blockId: 'note-p' }))
    assert.equal(document.getContentControl('header-control').getInfo().checked, true)
    assert.equal(document.getModel().sectionHeaderFooters.first.header, '☒')
    assert.equal(document.getModel().sectionHeaderFooters.first.storyContent.header.runs[0].fontFamily, 'Segoe UI Symbol')
    assert.equal(document.getContentControl('note-control').getInfo().value, 'typed note'); assert.equal(document.getModel().footnotes.note.text, 'typed note')
    assert.deepEqual(document.getModel().blocks, body); assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0 })
    document.undo(); assert.equal(document.getModel().footnotes.note.text, 'Alpha')
    document.undo(); assert.equal(document.getContentControl('header-control').getInfo().checked, false)
    document.redo(); document.redo(); assert.equal(document.getModel().footnotes.note.text, 'typed note')
  } finally { studio.dispose() }
})

test('cancellation, invalid dropdown input and paragraph permission denial cannot remove a temporary shell', () => {
  for (const scenario of ['cancel', 'dropdown', 'permission']) {
    const model = scenario === 'dropdown'
      ? fixture('Alpha', { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }], temporary: true })
      : fixture('Name', { type: 'text', value: 'Name', temporary: true })
    if (scenario === 'permission') model.permissionRules = [{ id: 'deny-input', scope: 'paragraph', targetId: 'p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
    const { studio, document } = open(model), before = copy(document.getModel()), events = []
    const after = studio.events.addEvent('CommandExecuted', event => { if (event.id === 'doc.command.insert-text') events.push(event.id) })
    const cancellation = scenario === 'cancel' ? studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.insert-text') event.cancel = true }) : undefined
    try {
      assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: 'typed' }, textSelection(0, scenario === 'dropdown' ? 5 : 4)), /canceled|下拉列表|permission denied/)
      assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); assert.deepEqual(events, [])
    } finally { after.dispose(); cancellation?.dispose(); studio.dispose() }
  }
})
