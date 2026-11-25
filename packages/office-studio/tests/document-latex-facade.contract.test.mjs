import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred contracts use the real command bus, content history and snapshot codec.
let server, OfficeStudio, FDocument, domain, model, latex, codec, controls, publicCommands
const studios = new Set(), copy = value => structuredClone(value)
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-latex-facade', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, model, latex, codec, controls, publicCommands] = await Promise.all(['domain/model', 'features/document/model/document-model',
    'features/document/latex/latex-formula', 'api/document-body-codec', 'features/document/services/content-controls', 'api/document-latex-commands']
    .map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const paragraph = (id, text = '') => ({ ...model.createParagraph(text), id })
const selection = (id, offset = 0, end = offset) => ({ anchor: { blockId: id, offset }, focus: { blockId: id, offset: end } })
function open(input) {
  const source = input ?? { ...domain.createBlankDocument(), id: 'latex-api', blocks: [paragraph('p', 'ABCD'), paragraph('tail', 'Tail')], modified: false }
  const studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(source.id)
  return { studio, document: new FDocument(studio, source.id) }
}
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })
const named = formulas => ({ name: 'DOC_LATEX_PLUGIN', ownerHint: 'resource', data: JSON.stringify({ authorHint: 'wrapper', formulas }) })

test('formula Facade create/query/update/remove share canonical command events and one content undo per edit', () => {
  const { studio, document } = open(), events = []
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') events.push(event.id) })
  const formula = document.insertLatexAtOffset(2, '  x ^2  ')
  assert.ok(formula); assert.equal(formula.getLatex(), '  x ^2  ')
  assert.deepEqual(formula.getRange(), { rangeId: formula.getId(), segmentId: '', startOffset: 2, endOffset: 10 })
  assert.equal(document.getLatexFormulaAt(9).getId(), formula.getId()); assert.equal(document.getLatexFormulaAt(10), null)
  assert.equal(document.findLatexFormulaByText('x ^').getId(), formula.getId())
  assert.deepEqual(formula.describe().context, { before: 'AB', after: 'CD' })
  assert.equal(formula.update({ latex: 'y_2' }), true); assert.equal(history(studio, document.id).undo, 2)
  assert.equal(document.undo(), true); assert.equal(formula.getLatex(), '  x ^2  ')
  assert.equal(document.redo(), true); assert.equal(formula.getLatex(), 'y_2')
  assert.equal(formula.remove(), true); assert.equal(formula.getRange(), null); assert.equal(document.getLatexFormulas().length, 0)
  assert.equal(document.undo(), true); assert.equal(formula.getLatex(), 'y_2')
  assert.deepEqual(events.filter(id => id.includes('latex-formula')), ['doc.command.insert-latex-formula', 'doc.command.update-latex-formula', 'doc.command.remove-latex-formula'])
})

test('paragraph append and range replacement use public source offsets and reject formula-interior replacement boundaries', () => {
  const { document } = open(), appended = document.getParagraph('p').appendLatex('a^2')
  assert.equal(appended.getRange().startOffset, 4)
  const range = document.getTextRange(0, 2), replacement = range.replaceWithLatex('\\frac{x}{y}')
  assert.ok(replacement); assert.equal(document.getParagraph('p').getText(), '\\frac{x}{y}CDa^2')
  const before = copy(document.getModel()), formulaRange = replacement.getRange()
  assert.throws(() => document.insertLatexAtOffset(formulaRange.startOffset + 1, 'z'), /实体内部/)
  assert.throws(() => document.getTextRange(formulaRange.startOffset + 1, formulaRange.endOffset).replaceWithLatex('z'), /实体内部/)
  assert.deepEqual(document.getModel(), before)
})

test('public aliases validate explicit range, segment and revision guards while keeping a single canonical history entry', () => {
  const { studio, document } = open(), commands = publicCommands.DOCUMENT_LATEX_PUBLIC_COMMANDS
  const created = studio.executeCommandSync(commands.insert, { unitId: document.id, startOffset: 1, latex: 'x', expectedRevision: document.getModelRevision(), expectedHistoryGeneration: document.getHistoryGeneration() })
  assert.ok(created); assert.equal(history(studio, document.id).undo, 1)
  assert.equal(studio.executeCommandSync(commands.update, { unitId: document.id, rangeId: created.rangeId, latex: 'y' }), true)
  assert.equal(studio.executeCommandSync(commands.replaceWithText, { unitId: document.id, rangeId: created.rangeId, text: 'y' }), true)
  assert.equal(history(studio, document.id).undo, 3)
  const before = copy(document.getModel()), previous = history(studio, document.id)
  for (const input of [{ startOffset: 0, selection: { startOffset: 1, endOffset: 1 } }, { endOffset: 2 },
    { startOffset: 0, expectedRevision: 0 }, { startOffset: 0, expectedHistoryGeneration: 0 }, { segmentId: 'other' }, { selection: false }]) {
    assert.throws(() => studio.executeCommandSync(commands.insert, { unitId: document.id, latex: 'z', ...input }))
  }
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous)
})

test('save/reopen retains stable run and range identity, exact raw source and public inline config, then removes deleted configs', () => {
  const { document } = open(), formula = document.insertLatexAtOffset(1, '  x^2  '), id = formula.getId()
  const saved = document.save(), resource = JSON.parse(saved.resources.find(resource => resource.name === 'DOC_LATEX_PLUGIN').data)
  assert.deepEqual(resource.formulas[id], { latex: '  x^2  ', kind: 'inline', displayMode: 'inline' })
  const reopened = open(copy(saved)).document, current = reopened.getLatexFormula(id)
  assert.equal(current.getLatex(), '  x^2  '); assert.equal(current.getRange().startOffset, 1)
  assert.equal(current.replaceWithText('  x^2  '), true)
  const deleted = reopened.save(), last = open(copy(deleted)).document
  assert.equal(last.getLatexFormulas().length, 0)
  assert.ok(!deleted.resources?.some(resource => resource.name === 'DOC_LATEX_PLUGIN' && JSON.parse(resource.data).formulas[id]))
  assert.ok(last.getBody().dataStream.includes('  x^2  '))
})

test('selection insertion applies pending format overrides without creating a separate content history for marks', () => {
  const { studio, document } = open(), selected = selection('p', 2)
  document.setInputSelection(selected)
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true, italic: true, color: '#123456', fontSize: 21 } }, selected)
  assert.equal(history(studio, document.id).undo, 0)
  const formula = document.insertLatexAtSelection('x^2'), run = latex.getLatexFormulas(document.getModel())[0].run
  assert.ok(formula); assert.deepEqual([run.bold, run.italic, run.color, run.fontSize], [true, true, '#123456', 21])
  assert.equal(history(studio, document.id).undo, 1)
})

test('canceled commands and locked formulas preserve model, pending marks and history atomically', () => {
  const unlocked = open(), formula = unlocked.document.insertLatexAtOffset(1, 'x'), before = copy(unlocked.document.getModel()), previous = history(unlocked.studio, unlocked.document.id)
  const stop = unlocked.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.update-latex-formula') event.cancel = true })
  try { assert.throws(() => formula.update({ latex: 'y' }), /canceled/) } finally { stop.dispose() }
  assert.deepEqual(unlocked.document.getModel(), before); assert.deepEqual(history(unlocked.studio, unlocked.document.id), previous)
  const source = { ...domain.createBlankDocument(), id: 'locked-api', blocks: [{ ...paragraph('p'), runs: [latex.createLatexFormulaRun('x')] }] }
  controls.createDocumentContentControl(source, { id: 'locked', type: 'richText', blockId: 'p', startOffset: 0, endOffset: 1, lock: 'contentLocked' })
  const locked = open(source), original = copy(locked.document.getModel()), lockedHistory = history(locked.studio, locked.document.id)
  assert.throws(() => locked.document.getLatexFormulas()[0].remove(), /locked/)
  assert.deepEqual(locked.document.getModel(), original); assert.deepEqual(history(locked.studio, locked.document.id), lockedHistory)
})

test('observer navigation after an insertion keeps the new location and withholds an obsolete formula handle', () => {
  const { studio, document } = open()
  document.setInputSelection(selection('p', 1))
  const observer = studio.events.addEvent('CommandExecuted', event => { if (event.id === 'doc.command.insert-latex-formula') document.setInputSelection(selection('tail', 1)) })
  try { assert.equal(document.insertLatexAtSelection('x^2'), null) } finally { observer.dispose() }
  assert.equal(document.getInputState().location.selection.focus.blockId, 'tail')
  assert.equal(document.getLatexFormulas().length, 1); assert.equal(history(studio, document.id).undo, 1)
})

test('shared public header IDs resolve independently copied hosts through the same generated aliases used on save', () => {
  const source = open().document.save(), content = { blocks: [{ ...paragraph('header-p'), runs: [latex.createLatexFormulaRun('h^2')] }], runs: [], paragraphs: [], metadata: {} }
  const data = codec.encodeDocumentBody(content)
  source.documentStyle.defaultHeaderId = 'shared'; source.documentStyle.firstPageHeaderId = 'shared'; source.documentStyle.useFirstPageHeaderFooter = 1
  source.headers = { shared: { ...data, headerId: 'shared', resources: [named(data.latexFormulas)] } }
  const { document } = open(source), snapshot = document.getSnapshot(), alias = snapshot.documentStyle.firstPageHeaderId
  assert.equal(snapshot.documentStyle.defaultHeaderId, 'shared'); assert.notEqual(alias, 'shared')
  assert.equal(document.getSegmentHost('shared').story, 'header'); assert.equal(document.getSegmentHost(alias).story, 'firstPageHeader')
  const first = document.getLatexFormulas(alias)[0], defaultFormula = document.getLatexFormulas('shared')[0]
  assert.notEqual(first.getId(), defaultFormula.getId()); assert.equal(first.update({ latex: 'f^2' }), true)
  assert.equal(defaultFormula.getLatex(), 'h^2')
  const reopened = open(document.save()).document
  assert.equal(reopened.getLatexFormulas(alias)[0].getLatex(), 'f^2'); assert.equal(reopened.getLatexFormulas('shared')[0].getLatex(), 'h^2')
})

test('note body formulas use the same Facade mutations and named resources and their source-text replacement reopens without a formula', () => {
  const { studio, document } = open(), content = { blocks: [{ ...paragraph('note-p'), runs: [latex.createLatexFormulaRun('n_1')] }], runs: [], paragraphs: [], metadata: {} }
  const data = codec.encodeDocumentBody(content), noteId = document.insertFootnote(1, { ...data, resources: [named(data.latexFormulas)] })
  assert.ok(noteId); const formula = document.getLatexFormulas(noteId)[0], before = history(studio, document.id)
  assert.equal(formula.getSegmentId(), noteId); assert.equal(formula.update({ latex: '  n_2  ' }), true)
  assert.equal(history(studio, document.id).undo, before.undo + 1)
  assert.equal(document.undo(), true); assert.equal(formula.getLatex(), 'n_1'); assert.equal(document.redo(), true)
  assert.equal(formula.replaceWithText(formula.getLatex()), true)
  const reopened = open(document.save()).document
  assert.equal(reopened.getLatexFormulas(noteId).length, 0); assert.ok(reopened.getBody(noteId).dataStream.includes('  n_2  '))
})
