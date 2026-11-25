import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred API/runtime ownership drafts. No server or runner was started here.
let server, OfficeStudio, FDocument, domain, model, latex, templates, editors, codec
const studios = new Set(), copy = value => structuredClone(value)
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-latex-template-facade', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, model, latex, templates, editors, codec] = await Promise.all(['domain/model', 'features/document/model/document-model',
    'features/document/latex/latex-formula', 'features/document/latex/latex-templates', 'api/document-latex-template-editors', 'api/document-body-codec']
    .map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const paragraph = (id, text = '') => ({ ...model.createParagraph(text), id })
const selection = (id, start = 0, end = start) => ({ anchor: { blockId: id, offset: start }, focus: { blockId: id, offset: end } })
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })
function open(input) {
  const source = input ?? { ...domain.createBlankDocument(), id: 'template-api', blocks: [paragraph('p', 'ABCD'), paragraph('tail', 'Tail')], modified: false }
  const studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(source.id)
  return { studio, document: new FDocument(studio, source.id) }
}
const named = formulas => ({ name: 'DOC_LATEX_PLUGIN', data: JSON.stringify({ formulas }) })

test('unbound and unknown-template operations return false without changing document state or history', () => {
  const { studio, document } = open(), before = copy(document.getModel()), previous = history(studio, document.id)
  assert.equal(document.insertLatexTemplate('fraction-basic'), false)
  const handle = document.bindLatexTemplateEditor(() => { throw new Error('Unknown IDs must not reach an editor') })
  assert.equal(document.insertLatexTemplate('missing'), false)
  assert.equal(studio.executeCommandSync(templates.DOCUMENT_LATEX_TEMPLATE_OPERATION, { unitId: document.id, templateId: 123 }), false)
  handle.dispose(); assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous)
})

test('Facade, public and local operations share a draft owner and keep snapshot, revision and content history unchanged', () => {
  const { studio, document } = open(), calls = [], before = copy(document.getModel()), revision = document.getModelRevision(), generation = document.getHistoryGeneration()
  const handle = document.bindLatexTemplateEditor(id => { calls.push(id); return true })
  assert.equal(document.insertLatexTemplate('fraction-basic'), true)
  assert.equal(studio.executeCommandSync(templates.DOCUMENT_LATEX_TEMPLATE_OPERATION, { unitId: document.id, templateId: 'matrix-2x2' }), true)
  assert.equal(studio.executeCommandSync(templates.DOCUMENT_LATEX_TEMPLATE_LOCAL_OPERATION, { unitId: document.id, templateId: 'large-sum' }), true)
  assert.deepEqual(calls, ['fraction-basic', 'matrix-2x2', 'large-sum'])
  assert.deepEqual(document.getModel(), before); assert.equal(document.getModelRevision(), revision)
  assert.equal(document.getHistoryGeneration(), generation); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 })
  handle.dispose()
})

test('draft ownership is isolated between units in one studio and between studios with the same unit ID', () => {
  const first = open(), second = open(), otherModel = first.studio.createDocument({ ...domain.createBlankDocument(), id: 'other-template-unit' }, false)
  const other = new FDocument(first.studio, otherModel.id), calls = []
  const handles = [first.document.bindLatexTemplateEditor(id => { calls.push('first:' + id); return true }),
    other.bindLatexTemplateEditor(id => { calls.push('other:' + id); return true }), second.document.bindLatexTemplateEditor(id => { calls.push('second:' + id); return true })]
  assert.equal(other.insertLatexTemplate('radical-square'), true); assert.equal(second.document.insertLatexTemplate('fraction-basic'), true)
  assert.deepEqual(calls, ['other:radical-square', 'second:fraction-basic']); assert.equal(handles[0].isCurrent(), true)
  assert.deepEqual(history(first.studio, first.document.id), { undo: 0, redo: 0 }); assert.deepEqual(history(first.studio, other.id), { undo: 0, redo: 0 })
  handles.forEach(handle => handle.dispose())
})

test('replacement invalidates an old owner and disposing the old handle cannot remove its replacement', () => {
  const { document } = open(), calls = []
  const old = document.bindLatexTemplateEditor(() => { calls.push('old'); return true })
  const current = document.bindLatexTemplateEditor(() => { calls.push('current'); return true })
  assert.equal(old.isCurrent(), false); assert.equal(current.isCurrent(), true)
  old.dispose(); old.dispose(); assert.equal(document.insertLatexTemplate('fraction-basic'), true); assert.deepEqual(calls, ['current'])
  current.dispose(); current.dispose(); assert.equal(current.isCurrent(), false); assert.equal(document.insertLatexTemplate('fraction-basic'), false)
})

test('disposing a unit clears its draft owner before an identical unit ID is created again', () => {
  const { studio, document } = open(), handle = document.bindLatexTemplateEditor(() => { throw new Error('Disposed owner must not run') })
  assert.equal(studio.disposeUnit(document.id), true); assert.equal(handle.isCurrent(), false)
  assert.equal(editors.requestDocumentLatexTemplate(studio, document.id, 'fraction-basic'), false)
  studio.createDocument({ ...domain.createBlankDocument(), id: document.id }, false)
  assert.equal(new FDocument(studio, document.id).insertLatexTemplate('fraction-basic'), false)
  handle.dispose()
})

test('studio disposal removes all runtime owners and later handle disposal is safe', () => {
  const { studio, document } = open(), handle = document.bindLatexTemplateEditor(() => true)
  studio.dispose(); assert.equal(handle.isCurrent(), false)
  assert.equal(editors.requestDocumentLatexTemplate(studio, document.id, 'fraction-basic'), false)
  handle.dispose()
})

test('explicit unit, revision and history guards reject before requesting source insertion', () => {
  const { studio, document } = open(), calls = [], handle = document.bindLatexTemplateEditor(id => { calls.push(id); return true })
  const id = templates.DOCUMENT_LATEX_TEMPLATE_OPERATION
  for (const guards of [{ unitId: 'missing' }, { expectedRevision: document.getModelRevision() + 1 },
    { expectedHistoryGeneration: document.getHistoryGeneration() + 1 }, { expectedRevision: NaN }, { unitId: 123 }]) {
    assert.throws(() => studio.executeCommandSync(id, { unitId: document.id, templateId: 'fraction-basic', ...guards }))
  }
  assert.deepEqual(calls, []); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 })
  assert.equal(studio.executeCommandSync(id, { unitId: document.id, templateId: 'fraction-basic', expectedRevision: document.getModelRevision(),
    expectedHistoryGeneration: document.getHistoryGeneration() }), true)
  handle.dispose()
})

test('a stale draft rejection and a canceled operation leave both draft source and content history untouched', () => {
  const { studio, document } = open(), calls = [], before = copy(document.getModel())
  const handle = document.bindLatexTemplateEditor(id => { calls.push(id); return false })
  assert.equal(document.insertLatexTemplate('fraction-basic'), false); assert.deepEqual(calls, ['fraction-basic'])
  const stop = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === templates.DOCUMENT_LATEX_TEMPLATE_OPERATION) event.cancel = true })
  try { assert.throws(() => document.insertLatexTemplate('radical-square'), /canceled/) } finally { stop.dispose() }
  assert.deepEqual(calls, ['fraction-basic']); assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); handle.dispose()
})

test('a template source update becomes content only on canonical apply and preserves the formula ID through undo and reopen', () => {
  const { studio, document } = open(), formula = document.insertLatexAtOffset(1, 'x_2'), formulaId = formula.getId(), before = history(studio, document.id)
  let draft = formula.getLatex()
  const handle = document.bindLatexTemplateEditor(id => { draft = templates.insertDocumentLatexTemplate(draft, id, 0, draft.length).value; return true })
  assert.equal(document.insertLatexTemplate('radical-nth'), true); assert.equal(draft, '\\sqrt[n]{x_2}')
  assert.equal(formula.getLatex(), 'x_2'); assert.deepEqual(history(studio, document.id), before)
  assert.equal(formula.update({ latex: draft }), true); assert.equal(formula.getId(), formulaId)
  assert.equal(history(studio, document.id).undo, before.undo + 1)
  assert.equal(document.undo(), true); assert.equal(formula.getLatex(), 'x_2')
  assert.equal(document.redo(), true); assert.equal(formula.getLatex(), draft)
  assert.equal(open(copy(document.save())).document.getLatexFormula(formulaId).getLatex(), draft); handle.dispose()
})

test('an empty formula draft can insert a complete matrix template and apply it as a new range once', () => {
  const { studio, document } = open()
  document.setInputSelection(selection('p', 2))
  let draft = '', placeholder
  const handle = document.bindLatexTemplateEditor(id => { const result = templates.insertDocumentLatexTemplate(draft, id, 0, 0); draft = result.value; placeholder = [result.selectionStart, result.selectionEnd]; return true })
  assert.equal(document.insertLatexTemplate('matrix-2x2'), true)
  assert.equal(draft.slice(...placeholder), 'a'); assert.equal(document.getLatexFormulas().length, 0)
  const formula = document.insertLatexAtSelection(draft)
  assert.ok(formula); assert.equal(formula.getLatex(), draft); assert.equal(history(studio, document.id).undo, 1)
  assert.equal(open(copy(document.save())).document.getLatexFormula(formula.getId()).getLatex(), draft); handle.dispose()
})

test('a template draft applies to a nested table formula while preserving its target and stable identity', () => {
  const table = model.createTable(1, 1); table.rows[0].cells[0].blocks = [{ ...paragraph('cell-p'), runs: [latex.createLatexFormulaRun('z_2')] }]
  const source = { ...domain.createBlankDocument(), id: 'nested-template-api', blocks: [{ ...paragraph('table'), type: 'table', runs: [], metadata: { table } }] }
  const { studio, document } = open(source), formula = document.getLatexFormulas()[0], formulaId = formula.getId()
  const target = model.textTargets(document.getModel()).find(target => target.cellBlockId === 'cell-p'), selected = { anchor: { ...target, offset: 0 }, focus: { ...target, offset: 1 } }
  document.setInputSelection(selected)
  let draft = formula.getLatex()
  const handle = document.bindLatexTemplateEditor(id => { draft = templates.insertDocumentLatexTemplate(draft, id, 0, draft.length).value; return true })
  assert.equal(document.insertLatexTemplate('accent-hat'), true)
  document.executeCommand({ id: 'doc.command.update-latex-formula', rangeId: formulaId, latex: draft }, selected)
  assert.equal(history(studio, document.id).undo, 1)
  const reopened = open(copy(document.save())).document, item = latex.getLatexFormulas(reopened.getModel())[0]
  assert.equal(item.rangeId, formulaId); assert.equal(item.blockId, 'cell-p'); assert.equal(item.formula.latex, '\\hat{z_2}'); handle.dispose()
})

test('header and note template applies preserve segment IDs, independent formulas and named resources on save/reopen', () => {
  const saved = open().document.save(), headerBody = codec.encodeDocumentBody({ blocks: [{ ...paragraph('header-p'), runs: [latex.createLatexFormulaRun('h_1')] }], runs: [], paragraphs: [], metadata: {} })
  saved.documentStyle.defaultHeaderId = 'math-header'; saved.headers = { 'math-header': { ...headerBody, headerId: 'math-header', resources: [named(headerBody.latexFormulas)] } }
  const { studio, document } = open(saved), noteBody = codec.encodeDocumentBody({ blocks: [{ ...paragraph('note-p'), runs: [latex.createLatexFormulaRun('n_1')] }], runs: [], paragraphs: [], metadata: {} })
  const noteId = document.insertFootnote(1, { ...noteBody, resources: [named(noteBody.latexFormulas)] }), values = []
  for (const segmentId of ['math-header', noteId]) {
    const formula = document.getLatexFormulas(segmentId)[0], original = formula.getLatex(), previous = history(studio, document.id)
    let draft = original
    const handle = document.bindLatexTemplateEditor(id => { draft = templates.insertDocumentLatexTemplate(draft, id, 0, draft.length).value; return true })
    assert.equal(document.insertLatexTemplate('radical-square'), true); assert.equal(formula.getLatex(), original)
    assert.equal(formula.update({ latex: draft }), true); assert.equal(history(studio, document.id).undo, previous.undo + 1)
    values.push({ segmentId, id: formula.getId(), source: draft }); handle.dispose()
  }
  const reopened = open(copy(document.save())).document
  for (const item of values) {
    const formula = reopened.getLatexFormula(item.id, item.segmentId)
    assert.equal(formula.getSegmentId(), item.segmentId); assert.equal(formula.getLatex(), item.source)
    assert.ok(reopened.getBody(item.segmentId).dataStream.includes(item.source))
  }
})
