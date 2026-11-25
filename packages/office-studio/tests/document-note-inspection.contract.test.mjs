import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-12 inspection/session drafts. No tests, browser, Office file reopen or performance acceptance were run.
let server, domain, model, stories, hosts, inspection, layout, state, draft, canvas, panel, decoration, settings, api
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-note-inspection', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  model = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  stories = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-story-blocks.ts')
  hosts = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-block-host.ts')
  inspection = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-block-host-inspection.ts')
  layout = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
  state = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/document-note-settings-state.ts')
  draft = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/document-note-settings-draft.ts')
  canvas = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/document-note-canvas.ts')
  panel = await server.ssrLoadModule('/packages/office-studio/src/features/document/controllers/document-note-panel.ts')
  decoration = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-note-decoration-host.ts')
  settings = await server.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-note-settings.ts')
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value)), measurer = { measure: text => [...text].length * 7 }
const permission = { documentId: 'note-inspection', userId: 'editor' }
function paragraph(id, text = '') { const block = model.createParagraph(text); block.id = id; block.runs[0].id = id + '-run'; return block }
const story = blocks => stories.projectDocumentStoryBlocks({ blocks, runs: [], paragraphs: [] }, 'display').content
function fixture() {
  const value = domain.createBlankDocument(); value.id = 'note-inspection'; value.blocks = [paragraph('body', 'Body text')]
  const cell = paragraph('note-cell', 'Nested note text'), image = { ...paragraph('note-image'), type: 'image', runs: [],
    metadata: { image: { src: 'data:image/png;base64,AA==', width: 30, height: 18, alt: 'Note image' }, anchorBlockId: cell.id, anchorOffset: 0 } }
  const table = { ...paragraph('note-table'), type: 'table', runs: [], metadata: { table: { columnWidths: [180], headerRowCount: 0,
    rows: [{ id: 'note-row', cells: [{ id: 'note-cell-id', blocks: [cell, image], style: {}, rowSpan: 1, colSpan: 1 }] }] } } }
  value.footnotes = { orphan: { id: 'orphan', text: 'Nested note text', referenceBlockId: '', referenceRunId: '', content: story([table]) } }
  const dormant = story([paragraph('dormant', 'Dormant separator')]); dormant.metadata = {
    decorationSource: { version: 1, kind: 'footnote', type: 'separator', noteId: 99, partName: 'word/footnotes.xml', sourceXml: '<w:footnote w:id="99"/>' }, rawXml: '<w:footnote w:id="99"/>',
  }
  value.noteSettings = { footnote: { separator: story([paragraph('active', 'Active separator')]),
    metadata: { officeDormantNoteDecorations: { version: 1, contents: [dormant] } } } }
  return value
}

test('inspection lays out complete orphan table/image hosts without replacing body pagination or creating references', () => {
  const value = fixture(), before = copy(value), physical = layout.layoutDocument(value, measurer)
  const host = { kind: 'note', noteKind: 'footnote', noteId: 'orphan' }, view = inspection.layoutDocumentBlockHostInspection(value, host, measurer)
  assert.deepEqual(value, before); assert.deepEqual(view.inspectionHost, host)
  assert.equal(view.lines.length, 0); assert.equal(view.cells.length, 0); assert.equal(view.images.length, 0); assert.equal(view.noteDecorations.length, 0)
  const lines = view.auxiliaryContent.flatMap(content => content.lines), images = view.auxiliaryContent.flatMap(content => content.images)
  assert.ok(lines.some(line => line.target.cellBlockId === 'note-cell' && line.host.noteId === 'orphan'))
  assert.ok(images.some(image => image.blockId === 'note-image' && image.host.noteId === 'orphan'))
  assert.ok(view.auxiliaryContent.every(content => content.host.kind === 'note' && content.host.noteId === 'orphan'))
  assert.equal(lines.flatMap(line => line.glyphs).some(glyph => glyph.note), false)
  assert.deepEqual(layout.layoutDocument(value, measurer).pages, physical.pages)
  assert.equal(physical.auxiliaryContent.some(content => content.host.kind === 'note'), false)
})

test('active, default and every dormant decoration have separate real addresses and native provenance', () => {
  const value = fixture(), before = copy(value), entries = state.documentNoteDecorationEntries(value, 'footnote')
  assert.equal(entries.length, 4); assert.deepEqual(entries.slice(0, 3).map(entry => entry.key), ['separator', 'continuationSeparator', 'continuationNotice'])
  const dormant = entries.find(entry => entry.dormant); assert.equal(dormant.nativeId, 99); assert.equal(dormant.host.dormantIndex, 0)
  const active = inspection.layoutDocumentBlockHostInspection(value, entries[0].host, measurer)
  const inactive = inspection.layoutDocumentBlockHostInspection(value, dormant.host, measurer)
  assert.ok(active.auxiliaryContent.flatMap(content => content.lines).every(line => line.host.dormantIndex === undefined))
  assert.ok(inactive.auxiliaryContent.flatMap(content => content.lines).every(line => line.host.dormantIndex === 0 && line.target.blockId === 'dormant'))
  assert.deepEqual(value, before); assert.equal(Object.keys(value.footnotes).length, 1)
})

test('public special definitions retain string IDs without fabricated DOCX provenance and native validation cannot be bypassed', () => {
  const value = fixture(), dormant = value.noteSettings.footnote.metadata.officeDormantNoteDecorations.contents[0]
  delete dormant.metadata.decorationSource
  dormant.metadata.univerNoteDecorationSource = { version: 1, kind: 'footnote', type: 'separator', noteId: 'extra-public-separator' }
  const before = copy(value), entry = state.documentNoteDecorationEntries(value, 'footnote').find(entry => entry.dormant)
  assert.equal(entry.publicId, 'extra-public-separator'); assert.equal(entry.nativeId, undefined)
  assert.deepEqual(decoration.documentNoteDecorationSourceIdentity(dormant, 'footnote'), {
    format: 'univer', kind: 'footnote', type: 'separator', noteId: 'extra-public-separator',
  })
  inspection.layoutDocumentBlockHostInspection(value, entry.host, measurer); assert.deepEqual(value, before)
  settings.validateDocumentNoteSettings(value.noteSettings.footnote, 'footnote')
  for (const invalid of [null, {}, { version: 2, kind: 'footnote', type: 'separator', noteId: 'extra' },
    { version: 1, kind: 'endnote', type: 'separator', noteId: 'extra' }, { version: 1, kind: 'footnote', type: 'separator', noteId: 99 },
    { version: 1, kind: 'footnote', type: 'separator', noteId: 'bad\nID' }]) {
    const bad = copy(dormant); bad.metadata.univerNoteDecorationSource = invalid
    assert.throws(() => decoration.documentNoteDecorationSourceIdentity(bad, 'footnote'))
  }
  const invalidNative = copy(dormant); invalidNative.metadata.decorationSource = { version: 1, kind: 'footnote', type: 'separator', noteId: 99 }
  assert.throws(() => decoration.documentNoteDecorationSourceIdentity(invalidNative, 'footnote'), /原生/)
  const active = value.noteSettings.footnote.separator
  active.metadata = { univerNoteDecorationSource: { version: 1, kind: 'footnote', type: 'separator', noteId: 'active-public' } }
  active.metadata.decorationSource = { version: 1, kind: 'footnote', type: 'separator', noteId: -1,
    partName: 'word/footnotes.xml', sourceXml: '<w:footnote w:id="-1" w:type="separator"/>' }
  const activeEntry = state.documentNoteDecorationEntries(value, 'footnote').find(entry => entry.key === 'separator')
  assert.equal(activeEntry.publicId, 'active-public'); assert.equal(activeEntry.nativeId, -1)
  assert.equal(settings.setDocumentNoteSettings(value, 'footnote', { ...value.noteSettings.footnote, separator: story([paragraph('replacement', 'Changed separator')]) }), true)
  assert.equal(value.noteSettings.footnote.separator.metadata.univerNoteDecorationSource.noteId, 'active-public')
  assert.equal(value.noteSettings.footnote.separator.metadata.decorationSource.noteId, -1)
  assert.deepEqual(value.noteSettings.footnote.metadata.officeDormantNoteDecorations, before.noteSettings.footnote.metadata.officeDormantNoteDecorations)
})

test('default empty continuation notice is a stable detached draft and only its first registered edit persists content', () => {
  const value = fixture(), studio = new api.OfficeStudio({ document: value, currentUserId: 'editor' }); studio.activateDocument(value.id)
  const document = new api.FDocument(studio, value.id), host = { kind: 'noteDecoration', noteKind: 'footnote', decoration: 'continuationNotice' }
  try {
    const before = copy(document.getModel()), first = hosts.documentModelForBlockHost(document.getModel(), host), second = hosts.documentModelForBlockHost(document.getModel(), host)
    assert.equal(first.blocks[0].type, 'paragraph'); assert.equal(first.blocks[0].id, second.blocks[0].id)
    inspection.layoutDocumentBlockHostInspection(document.getModel(), host, measurer)
    assert.deepEqual(document.getModel(), before); assert.equal(studio.editor.history.units[document.id]?.undo.length ?? 0, 0)
    const point = { ...model.textTargets(first)[0], offset: 0 }, selection = { anchor: point, focus: point }
    document.setInputSelection(selection, undefined, host)
    document.executeCommand({ id: 'doc.command.insert-text', text: 'Continue😀', host }, selection)
    assert.equal(model.runText(document.getModel().noteSettings.footnote.continuationNotice.runs), 'Continue😀')
    assert.equal(studio.editor.history.units[document.id].undo.length, 1)
    document.undo(); assert.equal(document.getModel().noteSettings.footnote.continuationNotice, undefined)
    document.redo(); assert.equal(model.runText(document.getModel().noteSettings.footnote.continuationNotice.runs), 'Continue😀')
    const saved = copy(document.saveModel()), reopened = new api.OfficeStudio({ document: saved, currentUserId: 'editor' })
    try { assert.deepEqual(new api.FDocument(reopened, saved.id).getModel().noteSettings, saved.noteSettings) } finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('settings/mark drafts and scoped view denial guard inspection while inherited form values preserve decoration sources', () => {
  const value = fixture(), host = { kind: 'noteDecoration', noteKind: 'footnote', decoration: 'separator' }
  assert.equal(state.documentNoteDecorationInspectState(value, host, { permission, pendingSettings: true }).available, false)
  assert.equal(state.documentNoteDecorationInspectState(value, host, { permission, pendingMark: true }).available, false)
  value.permissionRules = [{ id: 'hide-decoration', scope: 'paragraph', targetId: 'active', action: 'view', effect: 'deny', userIds: ['editor'] }]
  assert.equal(state.documentNoteDecorationInspectState(value, host, { permission }).available, false)
  value.permissionRules[0].action = 'edit'
  assert.equal(state.documentNoteDecorationInspectState(value, host, { permission }).available, true)
  assert.equal(state.documentNoteDecorationResetState(value, host, { permission }).available, false)
  const base = copy(value.noteSettings.footnote), converted = draft.documentNoteSettingsFromForm(base,
    { position: 'beneathText', restart: 'continuous', numberFormat: 'decimal', startNumber: 2, columnCount: 1 }, 'footnote')
  assert.deepEqual(converted.separator, base.separator); assert.deepEqual(converted.metadata, base.metadata); assert.deepEqual(value.noteSettings.footnote, base)
})

test('note preview, panel and Canvas requests honor scoped view rules and reject stale revisions, generations and locations', () => {
  const value = fixture(), body = { blockId: 'body', offset: 0 }, selection = { anchor: body, focus: body }
  const session = { valid: true, marks: {}, modelRevision: 4, generation: 8, location: { kind: 'body', selection } }
  const request = { type: 'note', source: value, inputSession: copy(session), kind: 'footnote', noteId: 'orphan', toReference: false }
  const context = { source: value, session, requireSession: true, permission }
  assert.equal(canvas.documentNoteCanvasRequestCurrent(request, context), true)
  for (const current of [{ ...session, generation: 9 }, { ...session, modelRevision: 5 }, { ...session, valid: false },
    { ...session, location: { kind: 'body', selection: { anchor: { ...body, offset: 1 }, focus: { ...body, offset: 1 } } } }]) {
    assert.equal(canvas.documentNoteCanvasRequestCurrent(request, { ...context, session: current }), false)
  }
  assert.equal(canvas.documentNoteCanvasRequestCurrent(request, { ...context, source: copy(value) }), false)
  value.permissionRules = [{ id: 'hide-note-cell', scope: 'paragraph', targetId: 'note-cell', action: 'view', effect: 'deny', userIds: ['editor'] }]
  assert.equal(canvas.documentNotePreview(value, 'footnote', 'orphan', undefined, permission), undefined)
  assert.equal(canvas.documentNoteCanvasRequestCurrent(request, context), false)
  const entry = panel.documentNotePanelEntries(value, { 'footnote:orphan': 'SECRET' }, permission)[0]
  assert.equal(entry.viewable, false); assert.equal(entry.text, ''); assert.equal(entry.label, ''); assert.equal(entry.editable, false)
  value.permissionRules[0].action = 'edit'
  assert.ok(canvas.documentNotePreview(value, 'footnote', 'orphan', undefined, permission))
  assert.equal(canvas.documentNoteCanvasRequestCurrent(request, context), true)
})
