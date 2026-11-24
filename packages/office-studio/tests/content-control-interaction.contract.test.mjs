import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred DOC-11 interaction contracts; these do not prove native focus or IME behavior.
let server, domain, modelUtils, controls, interaction
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-interaction', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  interaction = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/content-control-interaction.ts')
})
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const selection = (start, end = start, target = { blockId: 'p' }) => ({ anchor: { ...target, offset: start }, focus: { ...target, offset: end } })
const session = (value, patch = {}) => ({ valid: true, modelRevision: 4, generation: 7, marks: { bold: true },
  location: { kind: 'body', selection: copy(value) }, ...patch })
function fixture(text = 'a😀b') {
  const model = domain.createBlankDocument(); model.id = 'control-interaction'; model.blocks[0].id = 'p'; model.blocks[0].runs[0].text = text
  model.contentControls = []
  controls.createDocumentContentControl(model, { id: 'control', type: 'text', blockId: 'p', startOffset: 0, endOffset: text.length,
    value: text, temporary: true, lock: 'unlocked' })
  return model
}
function request(model, value = selection(0, 4), host) {
  return { source: model, controlId: 'control', selection: copy(value), ...(host ? { host: { ...host } } : {}),
    inputSession: session(value, host ? { location: { kind: 'body', selection: copy(value), host: { ...host } } } : {}) }
}
function context(model, value, input, host) {
  return { source: model, selection: value, session: input, ...(host ? { host } : {}) }
}

test('a removed temporary shell can focus only its checked current caret, without restoring the old selection', () => {
  const source = fixture(), pending = request(source), edited = copy(source)
  controls.setDocumentContentControlValue(edited, 'control', '李😀')
  assert.deepEqual(edited.contentControls, [])
  const caret = selection(3), current = session(caret, { modelRevision: 5, generation: 9 })
  const token = interaction.captureDocumentContentControlFocusToken(pending, current)
  assert.ok(token); assert.equal(interaction.documentContentControlFocusCurrent(token, context(edited, caret, current)), true)
  assert.equal(interaction.documentContentControlFocusCurrent(token, context(edited, pending.selection, current)), false)
  assert.equal(interaction.documentContentControlFocusTokenMatchesRequest(token, pending), true)
  assert.deepEqual(pending.selection, selection(0, 4)); assert.deepEqual(current.location.selection, caret)
})

test('value edit selection resolves surviving ranges and the UTF-16 end of a removed temporary within surrounding text', () => {
  const source = fixture('Before a😀b After'), start = 'Before '.length, end = start + 'a😀b'.length
  source.contentControls[0].startOffset = start; source.contentControls[0].endOffset = end; source.contentControls[0].value = 'a😀b'
  const pending = request(source, selection(start, end)), edited = copy(source)
  controls.setDocumentContentControlValue(edited, 'control', '李😀')
  assert.equal(modelUtils.runText(edited.blocks[0].runs), 'Before 李😀 After')
  assert.deepEqual(interaction.documentContentControlSelectionAfterValueEdit(pending, edited), selection(start + 3))
  assert.deepEqual(pending.selection, selection(start, end))
  const persistent = copy(source); persistent.contentControls[0].temporary = false
  const persistentRequest = request(persistent, selection(start, end)), persistentEdit = copy(persistent)
  controls.setDocumentContentControlValue(persistentEdit, 'control', '李😀')
  assert.deepEqual(interaction.documentContentControlSelectionAfterValueEdit(persistentRequest, persistentEdit), selection(start, start + 3))
  assert.deepEqual(interaction.documentContentControlSelectionAfterValueEdit(persistentRequest, persistent), selection(start, end))
  const missing = copy(persistentEdit); missing.contentControls = []
  assert.equal(interaction.documentContentControlSelectionAfterValueEdit(persistentRequest, missing), undefined)
  assert.equal(interaction.documentContentControlSelectionAfterValueEdit(pending, { ...edited, id: 'replacement' }), undefined)
  const alteredPrefix = copy(edited); alteredPrefix.blocks[0].runs[0].text = 'Changed ' + alteredPrefix.blocks[0].runs[0].text
  assert.equal(interaction.documentContentControlSelectionAfterValueEdit(pending, alteredPrefix), undefined)
})

test('focus capture copies selections and input marks instead of retaining a mutable callback result', () => {
  const source = fixture(), pending = request(source), current = session(selection(3), { modelRevision: 5, generation: 9 })
  const token = interaction.captureDocumentContentControlFocusToken(pending, current)
  assert.ok(token)
  current.location.selection.focus.offset = 2; current.marks.bold = false
  pending.selection.anchor.offset = 1; pending.inputSession.location.selection.focus.offset = 3
  assert.equal(token.session.location.selection.focus.offset, 3); assert.equal(token.session.marks.bold, true)
  assert.equal(token.request.selection.anchor.offset, 0); assert.equal(token.request.inputSession.location.selection.focus.offset, 4)
  assert.equal(interaction.documentContentControlFocusTokenMatchesRequest(token, pending), false)
})

test('same-model no-op and cancelled dispatch can restore focus with the completed selection generation', () => {
  const source = fixture(), pending = request(source), current = session(pending.selection, { generation: 9 })
  const token = interaction.captureDocumentContentControlFocusToken(pending, current)
  assert.ok(token); assert.equal(interaction.documentContentControlFocusCurrent(token, context(source, pending.selection, current)), true)
  assert.equal(interaction.documentContentControlFocusCurrent(token, context(source, pending.selection, { ...current, generation: 10 })), false)
  assert.equal(interaction.documentContentControlFocusCurrent(token, context(source, pending.selection, { ...current, valid: false })), false)
  assert.equal(interaction.documentContentControlFocusCurrent(token, context(source, pending.selection, { ...current, modelRevision: 5 })), false)
})

test('an observer-selected current caret remains authoritative and later asynchronous selection changes reject focus', () => {
  const source = fixture(), pending = request(source), edited = copy(source)
  controls.setDocumentContentControlValue(edited, 'control', '李😀')
  const outside = modelUtils.createParagraph('Outside'); outside.id = 'outside'; edited.blocks.push(outside)
  const caret = selection(0, 0, { blockId: 'outside' }), current = session(caret, { modelRevision: 5, generation: 9 })
  const token = interaction.captureDocumentContentControlFocusToken(pending, current)
  assert.ok(token); assert.equal(interaction.documentContentControlFocusCurrent(token, context(edited, caret, current)), true)
  const later = session(selection(2, 2, { blockId: 'outside' }), { modelRevision: 5, generation: 10 })
  assert.equal(interaction.documentContentControlFocusCurrent(token, context(edited, later.location.selection, later)), false)
  assert.deepEqual(current.location.selection, caret)
})

test('absent, invalid, older and multiply advanced sessions cannot authorize a delayed focus', () => {
  const source = fixture(), pending = request(source)
  for (const current of [undefined, { ...session(selection(1)), valid: false }, session(selection(1), { generation: 6 }),
    session(selection(1), { modelRevision: 3 }), session(selection(1), { modelRevision: 6 }), session(selection(1), { generation: 7.5 }),
    { ...session(selection(1)), location: { kind: 'story', range: { sectionId: 'first', story: 'header', start: 0, end: 0 } } }]) {
    assert.equal(interaction.captureDocumentContentControlFocusToken(pending, current), undefined)
  }
  assert.equal(interaction.captureDocumentContentControlFocusToken({ ...pending, inputSession: undefined }, session(selection(1))), undefined)
})

test('logical UTF-16 boundaries, exact table coordinates and surviving targets are required at focus time', () => {
  const source = fixture(), pending = request(source), current = session(selection(2), { modelRevision: 5, generation: 9 })
  const invalid = interaction.captureDocumentContentControlFocusToken(pending, current)
  assert.ok(invalid); assert.equal(interaction.documentContentControlFocusCurrent(invalid, context(source, selection(2), current)), false)
  const caret = selection(3), valid = session(caret, { modelRevision: 5, generation: 9 })
  const token = interaction.captureDocumentContentControlFocusToken(pending, valid)
  assert.equal(interaction.documentContentControlFocusCurrent(token, context({ ...source, blocks: [] }, caret, valid)), false)
  const table = modelUtils.createParagraph(); table.id = 'table'; table.type = 'table'; table.runs = []
  table.metadata = { table: { rows: [{ id: 'row', cells: [{ id: 'cell', blocks: [source.blocks[0]], rowSpan: 1, colSpan: 1, style: {} }] }] } }
  const tabular = { ...source, blocks: [table] }, target = { blockId: 'table', row: 0, col: 0, cellBlockId: 'p' }
  const tableRequest = request(tabular, selection(0, 4, target)), tableCaret = selection(3, 3, target)
  const tableInput = session(tableCaret, { modelRevision: 5, generation: 9 }), tableToken = interaction.captureDocumentContentControlFocusToken(tableRequest, tableInput)
  assert.equal(interaction.documentContentControlFocusCurrent(tableToken, context(tabular, tableCaret, tableInput)), true)
  const wrong = selection(3, 3, { ...target, col: 1 }), wrongInput = session(wrong, { modelRevision: 5, generation: 9 })
  const wrongToken = interaction.captureDocumentContentControlFocusToken(tableRequest, wrongInput)
  assert.equal(interaction.documentContentControlFocusCurrent(wrongToken, context(tabular, wrong, wrongInput)), false)
})

test('document identity, host identity, disabled state and the current view selection constrain focus', () => {
  const source = fixture(), pending = request(source), caret = selection(3), current = session(caret, { modelRevision: 5, generation: 9 })
  const token = interaction.captureDocumentContentControlFocusToken(pending, current)
  assert.equal(interaction.documentContentControlFocusCurrent(token, { ...context(source, caret, current), disabled: true }), false)
  assert.equal(interaction.documentContentControlFocusCurrent(token, context({ ...source, id: 'replacement' }, caret, current)), false)
  const host = { kind: 'note', noteKind: 'footnote', noteId: 'note' }
  assert.equal(interaction.documentContentControlFocusCurrent(token, context(source, caret, current, host)), false)
  assert.equal(interaction.documentContentControlFocusCurrent(token, context(source, selection(1), current)), false)
  assert.equal(interaction.documentContentControlFocusTokenMatchesRequest(token, { ...pending, controlId: 'other' }), false)
  assert.equal(interaction.documentContentControlFocusTokenMatchesRequest(token, { ...pending, source: copy(source) }), false)
})

test('header and note focus tokens validate their stored block host after temporary removal', () => {
  for (const host of [{ kind: 'story', sectionId: 'first', story: 'header' }, { kind: 'note', noteKind: 'footnote', noteId: 'note' }]) {
    const source = fixture(), content = { blocks: [source.blocks[0]], runs: copy(source.blocks[0].runs),
      paragraphs: [{ startOffset: 0, endOffset: 4, style: {} }] }
    source.blocks = [modelUtils.createParagraph('Body')]
    if (host.kind === 'story') source.sectionHeaderFooters = { first: { header: 'a😀b', footer: '', firstPageHeader: '', firstPageFooter: '',
      evenPageHeader: '', evenPageFooter: '', differentFirstPage: false, differentOddEven: false, linkToPrevious: false, storyContent: { header: content } } }
    else source.footnotes = { note: { id: 'note', text: 'a😀b', content } }
    const pending = request(source, selection(0, 4), host), edited = copy(source)
    controls.setDocumentContentControlValue(edited, 'control', '李😀')
    const caret = selection(3), current = session(caret, { modelRevision: 5, generation: 9, location: { kind: 'body', selection: caret, host } })
    const token = interaction.captureDocumentContentControlFocusToken(pending, current)
    assert.ok(token); assert.equal(interaction.documentContentControlFocusCurrent(token, context(edited, caret, current, host)), true)
    assert.deepEqual(interaction.documentContentControlSelectionAfterValueEdit(pending, edited), caret)
    assert.equal(interaction.documentContentControlFocusCurrent(token, context(edited, caret, current)), false)
    const removed = copy(edited)
    if (host.kind === 'story') delete removed.sectionHeaderFooters.first
    else delete removed.footnotes.note
    assert.equal(interaction.documentContentControlFocusCurrent(token, context(removed, caret, current, host)), false)
  }
})
