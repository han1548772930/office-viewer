import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred DOC-11 interaction contracts; these do not prove native focus or IME behavior.
let server, domain, modelUtils, controls, interaction, commandState, noteReferences
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-interaction', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  interaction = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/content-control-interaction.ts')
  commandState = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/content-control-command-state.ts')
  noteReferences = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-note-reference.ts')
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

function parentFixture() {
  const source = fixture('L😀R')
  Object.assign(source.contentControls[0], { startOffset: 1, endOffset: 3, value: '😀', temporary: false, parentId: 'parent' })
  source.contentControls.unshift({ id: 'parent', type: 'group', blockId: 'p', startOffset: 0, endOffset: 4, lock: 'unlocked' })
  return source
}
function putParentFixtureInTable(source) {
  const table = modelUtils.createParagraph(); table.id = 'parent-table'; table.type = 'table'; table.runs = []
  table.metadata = { table: { rows: [{ id: 'parent-row', cells: [{ id: 'parent-cell', blocks: [source.blocks[0]], rowSpan: 1, colSpan: 1, style: {} }] }] } }
  source.blocks = [table]
  return { blockId: table.id, row: 0, col: 0, cellBlockId: 'p' }
}

function repeatingFixture(text = 'L😀R', startOffset = 1, endOffset = 3) {
  const source = fixture(text)
  Object.assign(source.contentControls[0], { type: 'repeatingSectionItem', parentId: 'section', startOffset, endOffset,
    value: undefined, temporary: false })
  source.contentControls.unshift({ id: 'section', type: 'repeatingSection', blockId: 'p', startOffset: 0,
    endOffset: text.length, repeatingSection: { doNotAllowInsertDelete: false }, lock: 'unlocked' })
  return source
}
const repeatingRequest = (source, id = 'doc.command.duplicate-repeating-section-item', target = { blockId: 'p' }, host) => ({
  ...request(source, selection(source.contentControls.find(control => control.id === 'control').startOffset,
    source.contentControls.find(control => control.id === 'control').endOffset, target), host),
  command: { id, params: { controlId: 'control' } },
})

function noteRepeatingFixture() {
  const source = repeatingFixture('L9R10Z', 1, 2), paragraph = source.blocks[0]
  const run = (id, text, kind, noteId) => ({ ...modelUtils.createRun(text), id,
    ...(kind ? { metadata: { note: { type: kind, noteId }, noteReference: { kind, id: noteId, retained: 'reference extension' } }, verticalAlign: 'superscript' } : {}) })
  paragraph.runs = [run('left', 'L'), run('n9-run', '9', 'footnote', 'n9'), run('middle', 'R'), run('n10-run', '10', 'footnote', 'n10'), run('right', 'Z')]
  source.footnotes = { n9: { id: 'n9', referenceBlockId: 'p', referenceRunId: 'n9-run', text: 'Ninth note' },
    n10: { id: 'n10', referenceBlockId: 'p', referenceRunId: 'n10-run', text: 'Tenth note' } }
  const prefix = Array.from({ length: 8 }, (_, index) => {
    const block = modelUtils.createParagraph(); block.id = 'prefix-' + index
    block.runs = [run('prefix-run-' + index, String(index + 1), 'footnote', 'prefix-note-' + index)]
    source.footnotes['prefix-note-' + index] = { id: 'prefix-note-' + index, referenceBlockId: block.id, referenceRunId: block.runs[0].id, text: 'Earlier note ' + index }
    return block
  })
  source.blocks.unshift(...prefix)
  return source
}

test('actual inline note duplicate maps a 9/10/11 renumber to the new item range and preserves both reference markers', () => {
  const source = noteRepeatingFixture(), next = copy(source), before = copy(source)
  const copied = controls.duplicateRepeatingSectionItem(next, 'control')
  assert.deepEqual(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source), next), {
    controlId: copied.id, selection: selection(2, 4),
  })
  assert.equal(modelUtils.runText(modelUtils.findDocumentBlock(next, 'p').runs), 'L910R11Z')
  const reference = modelUtils.findDocumentBlock(next, 'p').runs.find(run => run.metadata?.note?.noteId !== 'n9' && run.text === '10')
  const note = next.footnotes[reference.metadata.note.noteId]
  assert.equal(reference.metadata.noteReference.id, note.id); assert.equal(reference.metadata.noteReference.retained, 'reference extension')
  assert.notEqual(note.id, 'n9'); assert.notEqual(reference.id, 'n9-run'); assert.equal(note.text, 'Ninth note')
  assert.equal(note.referenceBlockId, 'p'); assert.equal(note.referenceRunId, reference.id)
  assert.equal(next.footnotes.n10.referenceRunId, 'n10-run'); assert.deepEqual(source, before)
})

test('actual inline note deletion restores a verified caret despite shorter derived suffix labels', () => {
  const source = noteRepeatingFixture(), next = copy(source)
  controls.deleteRepeatingSectionItem(next, 'control')
  assert.equal(modelUtils.runText(modelUtils.findDocumentBlock(next, 'p').runs), 'LR9Z')
  assert.equal(next.footnotes.n9, undefined); assert.equal(next.footnotes.n10.referenceRunId, 'n10-run')
  assert.deepEqual(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, 'doc.command.delete-repeating-section-item'), next), {
    selection: selection(1),
  })
})

test('invalid note resources, copied IDs and suffix changes cannot fall back to matching plain text', () => {
  const source = noteRepeatingFixture(), next = copy(source), copied = controls.duplicateRepeatingSectionItem(next, 'control')
  const request = repeatingRequest(source), image = copy(next), paragraph = modelUtils.findDocumentBlock(image, 'p')
  paragraph.runs.at(-1).text = '!'
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(request, image), undefined)
  const missing = copy(next); missing.footnotes = {}
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(request, missing), undefined)
  const reused = copy(next), newRun = modelUtils.findDocumentBlock(reused, 'p').runs.find(run => run.text === '10')
  newRun.id = 'n9-run'; reused.footnotes[newRun.metadata.note.noteId].referenceRunId = 'n9-run'
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(request, reused), undefined)
  const ambiguous = copy(next); ambiguous.contentControls.push({ ...copy(copied), id: 'second-copy' })
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(request, ambiguous), undefined)
  const noWidthChange = repeatingFixture('L1R', 1, 2), reference = modelUtils.findDocumentBlock(noWidthChange, 'p').runs[0]
  reference.text = 'L'; modelUtils.findDocumentBlock(noWidthChange, 'p').runs.push({ ...modelUtils.createRun('1'), id: 'small-note-run',
    metadata: { note: { type: 'footnote', noteId: 'small-note' } } }, modelUtils.createRun('R'))
  noWidthChange.footnotes = { 'small-note': { id: 'small-note', text: 'Small note', referenceBlockId: 'p', referenceRunId: 'small-note-run' } }
  const committed = copy(noWidthChange); controls.duplicateRepeatingSectionItem(committed, 'control')
  const copiedNote = Object.values(committed.footnotes).find(note => note.id !== 'small-note'); delete committed.footnotes[copiedNote.id]
  // The text would satisfy the original 1-character insertion algorithm if the missing resource were ignored.
  modelUtils.findDocumentBlock(committed, 'p').runs.find(run => run.metadata?.note?.noteId === copiedNote.id).text = '1'
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(noWidthChange), committed), undefined)
})

test('note identity rejects conflicting fields and disambiguates equal numeric footnote/endnote IDs', () => {
  const foot = { id: '7', referenceBlockId: 'p', referenceRunId: 'foot', text: '' }, end = { ...foot, referenceRunId: 'end' }
  const reference = { id: 'migrated', metadata: { officeStoryRunSourceId: 'foot', noteReference: { footnoteId: 7 } } }
  assert.equal(noteReferences.documentNoteReferenceRunMatches(reference, foot, 'footnote'), true)
  assert.equal(noteReferences.documentNoteReferenceRunMatches(reference, end, 'endnote'), false)
  assert.equal(noteReferences.documentNoteReferenceRunMatches({ id: 'foot', metadata: { note: { noteId: '7', id: '8', type: 'footnote' } } }, foot, 'footnote'), false)
  assert.equal(noteReferences.documentNoteReferenceRunMatches({ id: 'foot', metadata: { note: { noteId: '7', type: 'endnote' } } }, foot, 'footnote'), false)
  const kindless = { id: 'foot', metadata: { noteReference: 7 } }
  assert.equal(noteReferences.documentNoteReferenceRunMatches(kindless, foot, 'footnote'), true)
  assert.equal(noteReferences.documentNoteReferenceRunMatches(kindless, end, 'endnote'), false)
})

test('repeating item copy chooses the new direct item ID and exact UTF-16 table range despite an existing item at the same boundary', () => {
  const source = repeatingFixture(), target = putParentFixtureInTable(source)
  source.contentControls.push({ id: 'sibling', type: 'repeatingSectionItem', parentId: 'section', blockId: 'p', startOffset: 3, endOffset: 4, lock: 'unlocked' },
    { id: 'nested', type: 'richText', parentId: 'control', blockId: 'p', startOffset: 1, endOffset: 3, lock: 'sdtLocked' })
  const before = copy(source), next = copy(source), copied = controls.duplicateRepeatingSectionItem(next, 'control')
  const result = interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, undefined, target), next)
  assert.deepEqual(result, { controlId: copied.id, selection: selection(3, 5, target) })
  assert.equal(next.contentControls.find(control => control.id === 'sibling').startOffset, 5)
  assert.equal(next.contentControls.find(control => control.parentId === copied.id).type, 'richText')
  result.selection.anchor.row = 9
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, undefined, target), next).selection.anchor.row, 0)
  assert.deepEqual(source, before)
})

test('a zero-length repeating item copy still identifies its new shell and never selects an old equal range or nested item', () => {
  const source = repeatingFixture('LR', 1, 1)
  source.contentControls.push({ id: 'equal-sibling', type: 'repeatingSectionItem', parentId: 'section', blockId: 'p', startOffset: 1, endOffset: 1, lock: 'unlocked' },
    { id: 'a-nested-item', type: 'repeatingSectionItem', parentId: 'control', blockId: 'p', startOffset: 1, endOffset: 1, lock: 'unlocked' })
  source.contentControls.sort((left, right) => left.id.localeCompare(right.id))
  const next = copy(source), copied = controls.duplicateRepeatingSectionItem(next, 'control')
  assert.equal(copied.parentId, 'section')
  assert.deepEqual(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source), next), {
    controlId: copied.id, selection: selection(1),
  })
  assert.equal(modelUtils.runText(next.blocks[0].runs), 'LR')
})

test('repeating item delete resolves a current table caret and rejects a canceled command or changed surrounding text', () => {
  const source = repeatingFixture(), target = putParentFixtureInTable(source), command = 'doc.command.delete-repeating-section-item'
  const next = copy(source); controls.deleteRepeatingSectionItem(next, 'control')
  assert.deepEqual(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, command, target), next), { selection: selection(1, 1, target) })
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, command, target), copy(source)), undefined)
  const paragraph = modelUtils.findDocumentBlock(next, 'p'); paragraph.runs[0].text = 'XLR'
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, command, target), next), undefined)
})

test('repeating copy restoration rejects absent, ambiguous, reordered-table, invalid-host and unrelated-document results', () => {
  const source = repeatingFixture(), target = putParentFixtureInTable(source), next = copy(source)
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, undefined, target), next), undefined)
  const copied = controls.duplicateRepeatingSectionItem(next, 'control')
  const ambiguous = copy(next); ambiguous.contentControls.push({ ...copy(copied), id: 'unexpected-copy' })
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, undefined, target), ambiguous), undefined)
  const moved = copy(next); moved.blocks[0].metadata.table.rows.unshift({ id: 'inserted-row', cells: [{ id: 'new-cell', blocks: [modelUtils.createParagraph('Outside')], rowSpan: 1, colSpan: 1, style: {} }] })
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, undefined, target), moved), undefined)
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, undefined, target), { ...next, id: 'other' }), undefined)
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(repeatingRequest(source, undefined, target, { kind: 'note', noteKind: 'footnote', noteId: 'missing' }), next), undefined)
  const mismatched = repeatingRequest(source, undefined, target); mismatched.command.params.controlId = 'sibling'
  assert.equal(interaction.documentContentControlSelectionAfterRepeatingItemEdit(mismatched, next), undefined)
})

test('repeating capabilities share section flags, own and ancestor locks, nested deletion locks, permission and UTF-16 validation with the service', () => {
  const source = repeatingFixture(), before = copy(source)
  for (const id of ['doc.command.duplicate-repeating-section-item', 'doc.command.delete-repeating-section-item']) {
    const command = { id, params: { controlId: 'control' } }, action = id.includes('duplicate') ? 'duplicate' : 'delete'
    assert.equal(commandState.contentControlRepeatingItemCommandState(source, command).available, true)
    for (const lock of ['sdtLocked', 'controlLocked', 'contentLocked', 'sdtContentLocked']) {
      const invalid = copy(source); invalid.contentControls.find(control => control.id === 'control').lock = lock
      assert.equal(commandState.contentControlRepeatingItemCommandState(invalid, command).available, false)
      const unchanged = copy(invalid); assert.throws(() => controls.assertDocumentRepeatingSectionItemEditable(invalid, 'control', action)); assert.deepEqual(invalid, unchanged)
    }
    const denied = copy(source); denied.permissionRules = [{ id: 'deny-repeat', scope: 'paragraph', targetId: 'p', action: 'edit', effect: 'deny', userIds: ['local-user'] }]
    assert.equal(commandState.contentControlRepeatingItemCommandState(denied, command).available, false)
    const disabled = commandState.contentControlRepeatingItemCommandState(source, command, { disabled: true }); assert.equal(disabled.available, false)
    const flagged = copy(source); flagged.contentControls.find(control => control.id === 'section').repeatingSection.doNotAllowInsertDelete = true
    assert.equal(commandState.contentControlRepeatingItemCommandState(flagged, command).available, false)
    const parentLocked = copy(source); parentLocked.contentControls.find(control => control.id === 'section').lock = 'contentLocked'
    assert.equal(commandState.contentControlRepeatingItemCommandState(parentLocked, command).available, false)
    const malformed = copy(source); malformed.contentControls.find(control => control.id === 'control').startOffset = 2
    assert.equal(commandState.contentControlRepeatingItemCommandState(malformed, command).available, false)
  }
  const nested = copy(source); nested.contentControls.push({ id: 'locked-child', type: 'text', parentId: 'control', blockId: 'p', startOffset: 1, endOffset: 3, lock: 'sdtLocked' })
  assert.equal(commandState.contentControlRepeatingItemCommandState(nested, { id: 'doc.command.delete-repeating-section-item', params: { controlId: 'control' } }).available, false)
  assert.equal(commandState.contentControlRepeatingItemCommandState(nested, { id: 'doc.command.duplicate-repeating-section-item', params: { controlId: 'control' } }).available, true)
  assert.deepEqual(source, before)
})

test('delayed requests cannot forge table coordinates for an otherwise valid control paragraph', () => {
  const source = repeatingFixture(), target = putParentFixtureInTable(source), captured = request(source, selection(1, 3, target))
  assert.equal(interaction.documentContentControlInteractionCurrent(captured, { source, session: captured.inputSession, requireSession: true }), true)
  for (const point of [{ ...target, row: 1 }, { ...target, col: 1 }, { ...target, blockId: 'another-table' }, { blockId: 'p' }]) {
    const forged = request(source, selection(1, 3, point))
    assert.equal(interaction.documentContentControlInteractionCurrent(forged, { source, session: forged.inputSession, requireSession: true }), false)
  }
})

test('parent selection resolves the exact table paragraph and returns detached UTF-16 points without changing the model', () => {
  const source = parentFixture(), target = putParentFixtureInTable(source), before = copy(source)
  assert.deepEqual(interaction.documentContentControlSelectionForId(source, 'control'), selection(1, 3, target))
  const parent = interaction.documentContentControlParentSelection(source, 'control')
  assert.deepEqual(parent, { controlId: 'parent', selection: selection(0, 4, target) })
  parent.selection.anchor.col = 9; parent.selection.focus.offset = 1
  assert.deepEqual(interaction.documentContentControlParentSelection(source, 'control'), { controlId: 'parent', selection: selection(0, 4, target) })
  assert.deepEqual(source, before)
  assert.equal(interaction.documentContentControlParentSelection(source, 'parent'), undefined)
})

test('legacy strict containment chooses the smallest range while an explicit parent remains authoritative', () => {
  const source = parentFixture(), child = source.contentControls.find(control => control.id === 'control')
  delete child.parentId
  source.contentControls.push({ id: 'nearer', type: 'group', blockId: 'p', startOffset: 1, endOffset: 4, lock: 'unlocked' })
  assert.deepEqual(interaction.documentContentControlParentSelection(source, 'control'), { controlId: 'nearer', selection: selection(1, 4) })
  child.parentId = 'parent'
  assert.deepEqual(interaction.documentContentControlParentSelection(source, 'control'), { controlId: 'parent', selection: selection(0, 4) })
  child.parentId = 'removed'
  assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
  child.parentId = 'control'
  assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
})

test('equal ranges need an explicit relationship because local control sorting does not preserve upstream custom-range order', () => {
  const source = parentFixture(), child = source.contentControls.find(control => control.id === 'control')
  const parent = source.contentControls.find(control => control.id === 'parent')
  Object.assign(parent, { startOffset: 1, endOffset: 3 })
  assert.deepEqual(interaction.documentContentControlParentSelection(source, 'control'), { controlId: 'parent', selection: selection(1, 3) })
  delete child.parentId
  assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
})

test('parent navigation rejects cross-paragraph relationships, missing hosts, synthetic story locks and malformed UTF-16 ranges', () => {
  const source = parentFixture(), child = source.contentControls.find(control => control.id === 'control')
  const parent = source.contentControls.find(control => control.id === 'parent'), outside = modelUtils.createParagraph('L😀R'); outside.id = 'outside'
  source.blocks.push(outside)
  parent.blockId = 'outside'; assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
  parent.blockId = 'p'
  for (const [startOffset, endOffset] of [[-1, 4], [0, 5], [0.5, 4], [4, 0]]) {
    Object.assign(parent, { startOffset, endOffset })
    assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
  }
  Object.assign(parent, { startOffset: 0, endOffset: 4 }); child.startOffset = 2
  assert.equal(interaction.documentContentControlSelectionForId(source, 'control'), undefined)
  assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
  child.startOffset = 1; child.endOffset = 1; parent.endOffset = 2
  assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
  child.endOffset = 3; parent.endOffset = 4; parent.metadata = { officeStoryLockSource: { version: 1 } }
  assert.equal(interaction.documentContentControlSelectionForId(source, 'parent'), undefined)
  assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
  delete child.parentId
  assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
  assert.equal(interaction.documentContentControlSelectionForId(source, 'removed'), undefined)
  assert.equal(interaction.documentContentControlParentSelection(source, 'control', { kind: 'note', noteKind: 'footnote', noteId: 'removed' }), undefined)
})

test('header and note parent navigation keep table coordinates and never resolve a parent from another content host', () => {
  for (const host of [{ kind: 'story', sectionId: 'first', story: 'header' }, { kind: 'note', noteKind: 'footnote', noteId: 'note' }]) {
    const source = parentFixture(), paragraph = source.blocks[0], target = putParentFixtureInTable(source)
    const content = { blocks: source.blocks, runs: copy(paragraph.runs), paragraphs: [{ startOffset: 0, endOffset: 4, style: {} }] }
    source.blocks = [modelUtils.createParagraph('Body')]
    if (host.kind === 'story') source.sectionHeaderFooters = { first: { header: 'L😀R', footer: '', firstPageHeader: '', firstPageFooter: '',
      evenPageHeader: '', evenPageFooter: '', differentFirstPage: false, differentOddEven: false, linkToPrevious: false, storyContent: { header: content } } }
    else source.footnotes = { note: { id: 'note', text: 'L😀R', content } }
    const before = copy(source)
    assert.deepEqual(interaction.documentContentControlSelectionForId(source, 'control', host), selection(1, 3, target))
    assert.deepEqual(interaction.documentContentControlParentSelection(source, 'control', host), { controlId: 'parent', selection: selection(0, 4, target) })
    assert.equal(interaction.documentContentControlParentSelection(source, 'control'), undefined)
    const elsewhere = modelUtils.createParagraph('L😀R'); elsewhere.id = 'elsewhere'; source.blocks.push(elsewhere)
    source.contentControls.push({ id: 'body-parent', type: 'group', blockId: 'elsewhere', startOffset: 0, endOffset: 4 })
    source.contentControls.find(control => control.id === 'control').parentId = 'body-parent'
    assert.equal(interaction.documentContentControlParentSelection(source, 'control', host), undefined)
    source.blocks.pop(); source.contentControls.pop(); source.contentControls.find(control => control.id === 'control').parentId = 'parent'
    assert.deepEqual(source, before)
  }
})

test('the shared reset capability requires a supported control and a nonempty prompt that is not already showing', () => {
  const command = { id: 'doc.command.reset-content-control-placeholder', params: { controlId: 'control' } }
  for (const type of ['text', 'richText', 'date', 'dropDownList', 'comboBox', 'picture']) {
    const source = fixture('A'); Object.assign(source.contentControls[0], { type, placeholder: 'Choose😀' })
    const before = copy(source)
    assert.equal(commandState.contentControlValueCommandState(source, command).available, true, type)
    source.contentControls[0].showingPlaceholder = true
    assert.equal(commandState.contentControlValueCommandState(source, command).available, false, type)
    source.contentControls[0].showingPlaceholder = false
    assert.deepEqual(source, before)
    for (const placeholder of ['', undefined]) {
      source.contentControls[0].placeholder = placeholder
      assert.equal(commandState.contentControlValueCommandState(source, command).available, false, type)
    }
  }
  for (const type of ['checkbox', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj']) {
    const source = fixture('A'); Object.assign(source.contentControls[0], { type, placeholder: 'Choose😀' })
    assert.equal(commandState.contentControlValueCommandState(source, command).available, false, type)
  }
})

test('reset capability honors content and ancestor locks, permission and host while allowing a wrapper-locked temporary', () => {
  const command = { id: 'doc.command.reset-content-control-placeholder', params: { controlId: 'control' } }
  const source = fixture('A'); source.contentControls[0].placeholder = 'Choose😀'
  for (const lock of ['contentLocked', 'sdtContentLocked']) {
    source.contentControls[0].lock = lock
    assert.equal(commandState.contentControlValueCommandState(source, command).available, false, lock)
  }
  source.contentControls[0].lock = 'sdtLocked'
  assert.equal(commandState.contentControlValueCommandState(source, command).available, true)
  source.contentControls[0].lock = 'unlocked'
  source.contentControls.push({ id: 'parent', type: 'group', blockId: 'p', startOffset: 0, endOffset: 1, lock: 'contentLocked' })
  source.contentControls[0].parentId = 'parent'
  assert.equal(commandState.contentControlValueCommandState(source, command).available, false)
  source.contentControls = source.contentControls.filter(control => control.id !== 'parent'); delete source.contentControls[0].parentId
  source.permissionRules = [{ id: 'denied', scope: 'paragraph', targetId: 'p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  assert.equal(commandState.contentControlValueCommandState(source, command, { permission: { documentId: source.id, userId: 'editor' } }).available, false)
  assert.equal(commandState.contentControlValueCommandState(source, command, { disabled: true }).available, false)
  assert.equal(commandState.contentControlValueCommandState(source, command, { host: { kind: 'note', noteKind: 'footnote', noteId: 'removed' } }).available, false)
  assert.equal(commandState.contentControlValueCommandState(source, { ...command, params: { controlId: 'removed' } }).available, false)
})

test('picture replacement capability shares host, permission and content-lock checks while allowing wrapper locks', () => {
  const source = fixture('P'), before = copy(source)
  assert.equal(commandState.contentControlPictureCommandState(source, 'control').available, false)
  source.contentControls[0].type = 'picture'
  for (const lock of ['unlocked', 'sdtLocked', 'controlLocked']) {
    source.contentControls[0].lock = lock
    assert.equal(commandState.contentControlPictureCommandState(source, 'control').available, true, lock)
  }
  for (const lock of ['contentLocked', 'sdtContentLocked']) {
    source.contentControls[0].lock = lock
    assert.equal(commandState.contentControlPictureCommandState(source, 'control').available, false, lock)
  }
  source.contentControls[0].lock = 'unlocked'
  source.contentControls.push({ id: 'parent', type: 'group', blockId: 'p', startOffset: 0, endOffset: 1, lock: 'contentLocked' })
  source.contentControls[0].parentId = 'parent'
  assert.equal(commandState.contentControlPictureCommandState(source, 'control').available, false)
  source.contentControls.pop(); delete source.contentControls[0].parentId
  source.permissionRules = [{ id: 'deny-picture', scope: 'paragraph', targetId: 'p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  assert.equal(commandState.contentControlPictureCommandState(source, 'control', { permission: { documentId: source.id, userId: 'editor' } }).available, false)
  assert.equal(commandState.contentControlPictureCommandState(source, 'control', { disabled: true }).available, false)
  assert.equal(commandState.contentControlPictureCommandState(source, 'control', { host: { kind: 'note', noteKind: 'footnote', noteId: 'removed' } }).available, false)
  assert.equal(commandState.contentControlPictureCommandState(source, 'removed').available, false)
  if (Object.hasOwn(before, 'permissionRules')) source.permissionRules = copy(before.permissionRules)
  else delete source.permissionRules
  source.contentControls[0].type = 'text'
  assert.deepEqual(source, before)
})

test('reset selection resolves actual prompt text after temporary removal and selects retained wrapper-locked shells', () => {
  const source = fixture('Before Value After'), start = 'Before '.length, end = start + 'Value'.length, placeholder = 'Choose😀'
  Object.assign(source.contentControls[0], { startOffset: start, endOffset: end, value: 'Value', placeholder })
  const pending = { ...request(source, selection(start, end)), command: { id: 'doc.command.reset-content-control-placeholder', params: { controlId: 'control' } } }
  const edited = copy(source), before = copy(source)
  controls.resetDocumentContentControlPlaceholder(edited, 'control')
  assert.deepEqual(edited.contentControls, []); assert.equal(modelUtils.runText(edited.blocks[0].runs), 'Before Choose😀 After')
  const caret = selection(start + placeholder.length)
  assert.deepEqual(interaction.documentContentControlSelectionAfterValueEdit(pending, edited), caret)
  const current = session(caret, { modelRevision: 5, generation: 9 }), token = interaction.captureDocumentContentControlFocusToken(pending, current)
  assert.ok(token); assert.equal(interaction.documentContentControlFocusCurrent(token, context(edited, caret, current)), true)
  assert.deepEqual(source, before)
  for (const properties of [{ temporary: false }, { lock: 'sdtLocked' }]) {
    const retained = copy(source); Object.assign(retained.contentControls[0], properties)
    const retainedRequest = { ...pending, source: retained }, after = copy(retained)
    controls.resetDocumentContentControlPlaceholder(after, 'control')
    assert.equal(after.contentControls[0].showingPlaceholder, true)
    assert.deepEqual(interaction.documentContentControlSelectionAfterValueEdit(retainedRequest, after), selection(start, start + placeholder.length))
  }
})

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
