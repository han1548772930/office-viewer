import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { SHEET_RANGE_THEME_SECTIONS, createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { initializeWorkbookCellStyles } from '../src/features/sheet-workbench/sheet-style-inheritance.ts'
import { normalizeSheetRangeThemeStyle } from '../src/features/sheet-workbench/sheet-range-theme-style.ts'
import { sheetRangeThemeStyleAt } from '../src/features/sheet-workbench/sheet-range-theme-resolver.ts'
import { sheetStyleToPublicData } from '../src/features/sheet-workbench/sheet-public-style.ts'
import { prepareSheetRangeThemeEditorDraft, resetSheetRangeThemeEditorSection, setSheetRangeThemeEditorBorder,
  setSheetRangeThemeEditorValue, uniqueSheetRangeThemeDraftName } from '../src/features/sheet-workbench/sheet-range-theme-editor-state.ts'
import { applySheetRangeThemeWorkbookPlan, planWorkbookSheetRangeThemeCommand, sheetRangeThemeCommandInput,
  SHEET_RANGE_THEME_COMMANDS } from '../src/features/sheet-workbench/sheet-range-theme-command.ts'

// Implementation-stage drafts. No runner, server, browser, type/build check or benchmark was executed.
const range = (startRow = 0, endRow = startRow, startCol = 0, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
const definition = (name = 'custom') => normalizeSheetRangeThemeStyle({ name, wholeStyle: { fs: 18, ol: { vendor: 'keep' } },
  headerRowStyle: { bl: 1, bg: { rgb: '#123456' } }, firstRowStyle: { bg: { rgb: '#ddeeff' } }, metadata: { vendor: 'theme' } })
function sample() {
  const model = createBlankWorkbook(); model.id = 'range-theme-editor'; model.calculationMode = 'manual'; model.rangeThemes = { custom: definition() }
  const sheet = model.sheets[0]; sheet.id = 'data'; sheet.name = 'Data'; sheet.rowCount = 50; sheet.columnCount = 16; sheet.defaultRowAutoHeight = false
  model.activeSheetId = sheet.id; sheet.cells['0:0'] = { ...createCell(0, 0, 'source'), authoredStyle: {} }
  sheet.rangeThemeRules = [{ id: 'first', themeName: 'custom', range: range(0, 4, 0, 3), metadata: { keep: true } }]
  const other = createBlankSheet('Other'); other.id = 'other'; other.defaultRowAutoHeight = false
  other.rangeThemeRules = [{ id: 'second', themeName: 'custom', range: range(0, 2, 0, 1) }]; model.sheets.push(other)
  initializeWorkbookCellStyles(model); return model
}

test('all nine editor sections stay independent and preserve the imported author extensions', () => {
  let theme = definition(); const original = cloneModel(theme)
  SHEET_RANGE_THEME_SECTIONS.forEach((section, index) => { theme = setSheetRangeThemeEditorValue(theme, section, 'font.size', index + 11) })
  SHEET_RANGE_THEME_SECTIONS.forEach((section, index) => assert.equal(theme[section].font.size, index + 11))
  assert.deepEqual(theme.wholeStyle.metadata.sourceStyle.ol, { vendor: 'keep' }); assert.deepEqual(theme.metadata, original.metadata)
  assert.equal(original.wholeStyle.font.size, 18); assert.equal(original.lastRowStyle, undefined)
})

test('inheritance, explicit false and no color retain distinct effects through section composition', () => {
  let theme = normalizeSheetRangeThemeStyle({ name: 'modes', wholeStyle: { bl: 1, bg: { rgb: '#ff0000' } }, headerRowStyle: {} })
  theme = setSheetRangeThemeEditorValue(theme, 'headerRowStyle', 'font.bold', false)
  theme = setSheetRangeThemeEditorValue(theme, 'headerRowStyle', 'fill', null)
  assert.equal(sheetRangeThemeStyleAt(theme, 0, 0, false, false).font.bold, false); assert.equal(sheetRangeThemeStyleAt(theme, 0, 0, false, false).fill, null)
  theme = setSheetRangeThemeEditorValue(theme, 'headerRowStyle', 'font.bold', undefined)
  theme = setSheetRangeThemeEditorValue(theme, 'headerRowStyle', 'fill', undefined)
  assert.equal(sheetRangeThemeStyleAt(theme, 0, 0, false, false).font.bold, true); assert.equal(sheetRangeThemeStyleAt(theme, 0, 0, false, false).fill, '#ff0000')
  const cleared = resetSheetRangeThemeEditorSection(theme, 'headerRowStyle')
  assert.equal(cleared.headerRowStyle, null); assert.equal(sheetRangeThemeStyleAt(cleared, 0, 0, false, false).fill, '#ff0000')
})

test('changing a font field preserves unrelated font, outline and source metadata', () => {
  const source = normalizeSheetRangeThemeStyle({ name: 'extensions', wholeStyle: { font: { family: 'Georgia', vendorFont: 7 },
    ul: { s: 1, vendorLine: 8 }, ol: { vendorOutline: 9 } } })
  const changed = setSheetRangeThemeEditorValue(source, 'wholeStyle', 'font.bold', false)
  assert.equal(changed.wholeStyle.font.family, 'Georgia'); assert.equal(changed.wholeStyle.font.underline, true)
  assert.deepEqual(changed.wholeStyle.metadata.sourceStyle, source.wholeStyle.metadata.sourceStyle)
  assert.equal(sheetStyleToPublicData(changed.wholeStyle).ol.vendorOutline, 9)
})

test('clearing a theme-only color prevents its old theme and tint from returning in a public read', () => {
  let theme = normalizeSheetRangeThemeStyle({ name: 'theme-colors', wholeStyle: { bg: { th: 4, tint: 0.4 }, cl: { th: 2, tint: -0.2 }, ol: { keep: true } } })
  theme = setSheetRangeThemeEditorValue(theme, 'wholeStyle', 'fill', undefined)
  theme = setSheetRangeThemeEditorValue(theme, 'wholeStyle', 'font.color', undefined)
  const output = sheetStyleToPublicData(theme.wholeStyle)
  assert.equal(Object.hasOwn(output, 'bg'), false); assert.equal(Object.hasOwn(output, 'cl'), false); assert.deepEqual(output.ol, { keep: true })
})

test('invalid field edits reject a detached candidate and never mutate the editor source', () => {
  const theme = definition(), before = cloneModel(theme)
  for (const [property, value] of [['font.size', 0], ['font.bold', 'yes'], ['rotation', Infinity], ['metadata', {}]]) {
    assert.throws(() => setSheetRangeThemeEditorValue(theme, 'wholeStyle', property, value)); assert.deepEqual(theme, before)
  }
  assert.throws(() => setSheetRangeThemeEditorValue(theme, 'missing', 'fill', null)); assert.deepEqual(theme, before)
})

test('wrap edits keep boolean and strategy coherent and inherited reset clears both', () => {
  let theme = definition(); theme = setSheetRangeThemeEditorValue(theme, 'wholeStyle', 'wrapStrategy', 'wrap')
  assert.equal(theme.wholeStyle.wrap, true); assert.equal(sheetStyleToPublicData(theme.wholeStyle).tb, 3)
  assert.equal(theme.wholeStyle.shrinkToFit, false)
  theme = setSheetRangeThemeEditorValue(theme, 'wholeStyle', 'shrinkToFit', true)
  assert.equal(theme.wholeStyle.wrap, false); assert.equal(theme.wholeStyle.wrapStrategy, 'overflow')
  theme = setSheetRangeThemeEditorValue(theme, 'wholeStyle', 'wrapStrategy', 'clip'); assert.equal(theme.wholeStyle.wrap, false)
  theme = setSheetRangeThemeEditorValue(theme, 'wholeStyle', 'wrapStrategy', undefined)
  assert.equal(theme.wholeStyle.wrap, undefined); assert.equal(theme.wholeStyle.wrapStrategy, undefined)
  theme = setSheetRangeThemeEditorValue(theme, 'wholeStyle', 'rotationVertical', true)
  theme = setSheetRangeThemeEditorValue(theme, 'wholeStyle', 'rotation', 45)
  assert.equal(theme.wholeStyle.rotationVertical, false); assert.equal(theme.wholeStyle.rotation, 45)
})

test('border line edits preserve unrelated edges, colors and imported edge extensions', () => {
  const source = normalizeSheetRangeThemeStyle({ name: 'borders', wholeStyle: { bd: {
    t: { s: 1, cl: { rgb: '#123456' }, vendor: 7 }, l: { s: 7, cl: { rgb: '#abcdef' } },
  } } })
  const theme = setSheetRangeThemeEditorBorder(source, 'wholeStyle', 'top', 'style', 'thick'), output = sheetStyleToPublicData(theme.wholeStyle)
  assert.equal(output.bd.t.s, 13); assert.equal(output.bd.t.cl.rgb, '#123456'); assert.equal(output.bd.t.vendor, 7); assert.deepEqual(output.bd.l, sheetStyleToPublicData(source.wholeStyle).bd.l)
  assert.equal(source.wholeStyle.border.top.style, 'thin')
})

test('explicit border removal blocks lower sections while inheritance restores their edges', () => {
  let theme = normalizeSheetRangeThemeStyle({ name: 'edge-modes', wholeStyle: { bd: { t: { s: 1 } } }, headerRowStyle: {} })
  theme = setSheetRangeThemeEditorBorder(theme, 'headerRowStyle', 'top', 'style', 'none')
  assert.equal(sheetRangeThemeStyleAt(theme, 0, 0, false, false).border.top, null)
  theme = setSheetRangeThemeEditorBorder(theme, 'headerRowStyle', 'top', 'style', undefined)
  assert.equal(sheetRangeThemeStyleAt(theme, 0, 0, false, false).border.top.style, 'thin')
})

test('inherited border color removes both native and raw aliases without dropping the line', () => {
  const source = normalizeSheetRangeThemeStyle({ name: 'edge-color', wholeStyle: { bd: { t: { s: 7, cl: { rgb: '#111111', th: 3 }, vendor: 4 } } } })
  const theme = setSheetRangeThemeEditorBorder(source, 'wholeStyle', 'top', 'color', undefined), output = sheetStyleToPublicData(theme.wholeStyle)
  assert.equal(output.bd.t.s, 7); assert.equal(output.bd.t.cl, undefined); assert.equal(output.bd.t.vendor, 4)
  assert.equal(source.wholeStyle.border.top.color, '#111111')
})

test('draft names reserve built-ins, reject accidental catalog replacement and generate unique copies', () => {
  const theme = definition(); assert.throws(() => prepareSheetRangeThemeEditorDraft(theme, 'default', []), /内置/)
  assert.throws(() => prepareSheetRangeThemeEditorDraft(theme, 'other', ['custom', 'other'], 'custom'), /已存在/)
  assert.equal(prepareSheetRangeThemeEditorDraft(theme, ' custom ', ['custom'], 'custom').name, 'custom')
  assert.equal(uniqueSheetRangeThemeDraftName('副本', ['副本', '副本 2']), '副本 3')
  assert.equal(prepareSheetRangeThemeEditorDraft(theme, '__proto__', []).name, '__proto__')
})

test('unknown top-level theme extensions normalize without requiring a metadata object', () => {
  const source = { name: 'extended', vendorTheme: { keep: true }, headerRowStyle: { bl: 0 } }
  const theme = normalizeSheetRangeThemeStyle(source)
  assert.deepEqual(theme.metadata.vendorTheme, { keep: true }); assert.equal(theme.headerRowStyle.font.bold, false)
  assert.equal(source.metadata, undefined)
})

test('one workbook plan captures a new catalog definition and target rule without changing any source cells', () => {
  const model = sample(), before = cloneModel(model), next = definition('new')
  const plan = planWorkbookSheetRangeThemeCommand(model, 'data', [range(0, 2, 5, 6)], 'set', { themeName: 'new', rangeThemeStyleJson: next })
  assert.equal(plan.catalogChanged, true); assert.equal(plan.rulesChanged, true); assert.deepEqual(model, before)
  const candidate = cloneModel(model); applySheetRangeThemeWorkbookPlan(candidate, plan)
  assert.deepEqual(candidate.rangeThemes.new, next); assert.equal(candidate.sheets[0].rangeThemeRules.at(-1).themeName, 'new')
  assert.deepEqual(candidate.sheets.map(sheet => sheet.cells), model.sheets.map(sheet => sheet.cells))
})

test('same-name catalog updates report every referencing worksheet even when the selected rule is unchanged', () => {
  const model = sample(), next = setSheetRangeThemeEditorValue(model.rangeThemes.custom, 'wholeStyle', 'font.size', 22)
  const plan = planWorkbookSheetRangeThemeCommand(model, 'data', [range(0, 4, 0, 3)], 'set', { themeName: 'custom', rangeThemeStyleJson: next })
  assert.equal(plan.catalogChanged, true); assert.equal(plan.rulesChanged, false); assert.deepEqual(Object.keys(plan.affected).sort(), ['data', 'other'])
  assert.deepEqual(plan.affected.other, [range(0, 2, 0, 1)]); assert.equal(plan.rangeThemeRules[0].id, 'first')
})

test('identical registration and application are no-ops and catalog ordering stays stable', () => {
  const model = sample(); model.rangeThemes.tail = definition('tail')
  for (const kind of ['register', 'set']) {
    const plan = planWorkbookSheetRangeThemeCommand(model, 'data', [range(0, 4, 0, 3)], kind, { themeName: 'custom', rangeThemeStyleJson: definition() })
    assert.equal(plan.changed, false); assert.deepEqual(Object.keys(plan.rangeThemes), ['custom', 'tail'])
  }
})

test('empty selections cannot register through apply and incompatible definition actions fail atomically', () => {
  const model = sample(), before = cloneModel(model)
  assert.equal(planWorkbookSheetRangeThemeCommand(model, 'data', [], 'set', { themeName: 'new', rangeThemeStyleJson: definition('new') }).changed, false)
  for (const kind of ['remove', 'unregister']) assert.throws(() => planWorkbookSheetRangeThemeCommand(model, 'data', [range()], kind,
    { themeName: 'custom', rangeThemeStyleJson: definition() }))
  assert.throws(() => planWorkbookSheetRangeThemeCommand(model, 'data', [range()], 'set', { themeName: 'custom', rangeThemeStyleJson: definition('different') }))
  assert.deepEqual(model, before)
})

test('publishing a detached catalog refreshes all existing view projections and keeps domain sheets free of view copies', () => {
  const model = sample(), view = toExcelWorkbookView(model), next = definition('new')
  const plan = planWorkbookSheetRangeThemeCommand(view, 'data', [range(5, 6)], 'set', { themeName: 'new', rangeThemeStyleJson: next })
  applySheetRangeThemeWorkbookPlan(view, plan); assert.deepEqual(view.sheets[0].rangeThemes.new, next); assert.deepEqual(view.sheets[1].rangeThemes.new, next)
  view.sheets[0].rangeThemes.new.wholeStyle.font.size = 1; assert.equal(view.sheets[1].rangeThemes.new.wholeStyle.font.size, 18); assert.equal(view.rangeThemes.new.wholeStyle.font.size, 18)
  const domainPlan = planWorkbookSheetRangeThemeCommand(model, 'data', [range(5, 6)], 'set', { themeName: 'new', rangeThemeStyleJson: next })
  applySheetRangeThemeWorkbookPlan(model, domainPlan); assert.ok(model.sheets.every(sheet => !Object.hasOwn(sheet, 'rangeThemes')))
})

test('workbook plans validate bounds, names, builders and confirmation versions before publication', () => {
  const model = sample(), before = cloneModel(model)
  assert.throws(() => planWorkbookSheetRangeThemeCommand(model, 'data', [range(-1)], 'set', { themeName: 'new', rangeThemeStyleJson: definition('new') }))
  assert.throws(() => planWorkbookSheetRangeThemeCommand(model, 'data', [range()], 'set', { themeName: 'default', rangeThemeStyleJson: definition('default') }), /内置/)
  assert.throws(() => planWorkbookSheetRangeThemeCommand(model, 'data', [range()], 'register', { rangeThemeStyle: {} }))
  for (const input of [{ expectedRevision: -1 }, { expectedHistoryGeneration: 0 }, { host: { unitId: 'x' } }]) assert.throws(() => sheetRangeThemeCommandInput(input))
  assert.deepEqual(model, before)
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
function withStudio(run) {
  const studio = new OfficeStudio({ workbook: sample() }), api = FOffice.newAPI(studio), book = api.getActiveWorkbook()
  try { return run(studio, api, book) } finally { studio.dispose() }
}
const apply = (studio, book, theme, selected = range(5, 7, 5, 6), extra = {}) => studio.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.set,
  { unitId: book.getId(), sheetId: 'data', ranges: [selected], themeName: theme.name, rangeThemeStyleJson: theme, ...extra })

test('save-and-apply has one history step and undo, redo and internal reopen preserve every author section', () => withStudio((studio, api, book) => {
  let theme = definition('edited'); for (const section of SHEET_RANGE_THEME_SECTIONS) theme = setSheetRangeThemeEditorValue(theme, section, 'font.size', 24)
  theme = setSheetRangeThemeEditorValue(theme, 'wholeStyle', 'indent', 2)
  const before = cloneModel(book.save()), history = studio.editor.history.units[book.getId()]?.undo.length ?? 0
  assert.equal(apply(studio, book, theme), true); assert.equal(studio.editor.history.units[book.getId()].undo.length, history + 1)
  assert.equal(book.getActiveSheet().getRange('F6').getCellStyleData().fs, 24); assert.equal(book.save().rangeThemes.edited.wholeStyle.indent, 2)
  assert.equal(Object.keys(book.save().sheets[0].cells).length, 1)
  book.undo(); assert.deepEqual(book.save().rangeThemes, before.rangeThemes); assert.deepEqual(book.save().sheets[0].rangeThemeRules, before.sheets[0].rangeThemeRules)
  book.redo(); const reopened = api.createWorkbook({ ...book.save(), id: 'editor-reopened' })
  assert.deepEqual(reopened.save().rangeThemes.edited, theme); assert.equal(reopened.getActiveSheet().getRange('F6').getUsedThemeStyle(), 'edited')
}))

test('editing a used custom theme updates blank ranges on every sheet and preserves author rule identities', () => withStudio((studio, _api, book) => {
  const rules = book.save().sheets.map(sheet => cloneModel(sheet.rangeThemeRules)), next = setSheetRangeThemeEditorValue(definition(), 'wholeStyle', 'font.size', 27)
  assert.equal(apply(studio, book, next, range(0, 4, 0, 3)), true)
  assert.equal(book.getActiveSheet().getRange('A2').getCellStyleData().fs, 27); assert.equal(book.getSheetBySheetId('other').getRange('A2').getCellStyleData().fs, 27)
  assert.deepEqual(book.save().sheets.map(sheet => sheet.rangeThemeRules), rules)
}))

test('a protected reference on another sheet rejects the whole catalog update and leaves target history unchanged', () => withStudio((studio, _api, book) => {
  studio.getWorkbookModel(book.getId()).sheets[1].rangeProtectionRules = [{ id: 'locked', ranges: [range(2, 2, 1, 1)], allowedUserIds: ['owner'] }]
  const before = cloneModel(book.save()), revision = studio.getWorkbookRevision(book.getId())
  assert.throws(() => apply(studio, book, setSheetRangeThemeEditorValue(definition(), 'wholeStyle', 'font.size', 28)), /保护/)
  assert.deepEqual(book.save(), before); assert.equal(studio.getWorkbookRevision(book.getId()), revision); assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
}))

test('protected target ranges reject a new definition without leaving an unused catalog registration', () => withStudio((studio, _api, book) => {
  studio.getWorkbookModel(book.getId()).sheets[0].rangeProtectionRules = [{ id: 'locked', ranges: [range(5, 5, 5, 5)], allowedUserIds: ['owner'] }]
  const before = cloneModel(book.save()); assert.throws(() => apply(studio, book, definition('blocked')), /保护/)
  assert.deepEqual(book.save(), before); assert.equal(Object.hasOwn(book.save().rangeThemes, 'blocked'), false)
}))

test('stale model and history versions reject save-and-apply before any catalog or rule is published', () => withStudio((studio, _api, book) => {
  const before = cloneModel(book.save()), revision = studio.getWorkbookRevision(book.getId()), generation = studio.editor.history.captureSavepoint(book.getId()).generation
  for (const extra of [{ expectedRevision: revision + 1 }, { expectedHistoryGeneration: generation + 1 }]) {
    assert.throws(() => apply(studio, book, definition('stale'), undefined, extra), /已更改/); assert.deepEqual(book.save(), before)
  }
}))

test('builder serialization that resets history invalidates the atomic save-and-apply candidate', () => withStudio((studio, _api, book) => {
  const builder = book.createRangeThemeStyle('reentrant', { headerRowStyle: { bl: 0 } }), serialize = builder.toJson.bind(builder)
  builder.toJson = () => { studio.editor.history.clear(book.getId()); return serialize() }
  const before = cloneModel(book.save())
  assert.throws(() => studio.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.set, { unitId: book.getId(), sheetId: 'data', ranges: [range(5)], themeName: 'reentrant', rangeThemeStyle: builder }), /已更改/)
  assert.deepEqual(book.save(), before); assert.equal(Object.hasOwn(book.save().rangeThemes, 'reentrant'), false)
}))

test('identical atomic reuse adds no history while a changed definition commits one step', () => withStudio((studio, _api, book) => {
  assert.equal(apply(studio, book, definition(), range(0, 4, 0, 3)), false)
  assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
  assert.equal(apply(studio, book, setSheetRangeThemeEditorValue(definition(), 'headerRowStyle', 'font.bold', false), range(0, 4, 0, 3)), true)
  assert.equal(studio.editor.history.units[book.getId()].undo.length, 1)
}))

test('an explicit background-sheet command works while a stale captured UI host is rejected', () => withStudio((studio, _api, book) => {
  const input = { unitId: book.getId(), sheetId: 'other', ranges: [range(5)], themeName: 'background', rangeThemeStyleJson: definition('background') }
  assert.throws(() => studio.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.set, { ...input, host: { unitId: book.getId(), sheetId: 'other' } }), /宿主/)
  assert.equal(Object.hasOwn(book.save().rangeThemes, 'background'), false)
  assert.equal(studio.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.set, input), true); assert.equal(book.getSheetBySheetId('other').getRange('A6').getUsedThemeStyle(), 'background')
}))

test('invalid atomic style DTOs preserve the source model, permissions and history', () => withStudio((studio, _api, book) => {
  const before = cloneModel(book.save())
  for (const theme of [{ name: 'bad', wholeStyle: { font: { size: 0 } } }, { name: 'default', wholeStyle: {} }, { name: 'bad', headerRowStyle: 'catalog' }]) {
    assert.throws(() => apply(studio, book, theme)); assert.deepEqual(book.save(), before)
  }
  assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
}))
