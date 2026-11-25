import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { normalizeSheetRangeThemeStyle } from '../src/features/sheet-workbench/sheet-range-theme-style.ts'
import { assertSheetRangeThemeRegionSnapshot, setSheetRangeThemeRules } from '../src/features/sheet-workbench/sheet-range-theme-model.ts'
import { captureSheetRangeThemeRegion, resolveSheetRangeThemeStyle, sheetRangeThemeRegionStyle, sheetUsedRangeThemeRule } from '../src/features/sheet-workbench/sheet-range-theme-resolver.ts'
import { assertSheetRangeThemeStyleRange, planSheetRangeThemeAxisEdit, planSheetRangeThemeCellShift, planSheetRangeThemeCut,
  sheetRangeThemeChangedRanges } from '../src/features/sheet-workbench/sheet-range-theme-structure.ts'
import { moveSheetCellRange } from '../src/features/sheet-workbench/sheet-cell-range-move.ts'
import { captureSheetClipboard, sheetClipboardSourceFingerprint } from '../src/features/sheet-workbench/sheet-clipboard-data.ts'
import { applySheetClipboardPastePlan, createSheetClipboardPastePlan } from '../src/features/sheet-workbench/sheet-clipboard.ts'
import { refreshWorkbookCellStyles } from '../src/features/sheet-workbench/sheet-style-inheritance.ts'
import { sheetRangeContainsCell } from '../src/features/sheet-workbench/sheet-range-command.ts'
import { SHEET_CLIPBOARD_COMMANDS } from '../src/features/sheet-workbench/sheet-clipboard-command.ts'

// Drafts only. No tests, type/build checks, browser/dev or Rust/WASM/benchmarks were run.
const range = (startRow, endRow, startCol = 0, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
const rule = () => ({ id: 'pattern-rule', themeName: 'pattern', range: range(0, 5, 0, 3), metadata: { vendor: { keep: true } } })
function sample() {
  const book = createBlankWorkbook(), sheet = book.sheets[0], other = createBlankSheet('Other')
  book.id = 'theme-structure'; book.calculationMode = 'manual'; sheet.id = 'data'; book.activeSheetId = sheet.id; other.id = 'other'
  for (const item of [sheet, other]) { item.rowCount = 100; item.columnCount = 20; item.defaultRowAutoHeight = false }
  book.sheets.push(other); sheet.rangeThemeRules = [rule()]
  book.rangeThemes = { pattern: normalizeSheetRangeThemeStyle({ name: 'pattern',
    wholeStyle: { ff: 'Georgia', cl: { rgb: '#101010' } }, headerRowStyle: { bg: { rgb: '#123456' } },
    firstRowStyle: { bg: { rgb: '#abcdef' } }, secondRowStyle: { bg: { rgb: '#fedcba' } }, lastRowStyle: { bg: { rgb: '#444444' } },
  }) }
  refreshWorkbookCellStyles(book, undefined, false); return book
}
const style = (rules, row, col, book = sample()) => resolveSheetRangeThemeStyle({ ...book.sheets[0], rangeThemeRules: rules }, row, col, book)
function pasted(book, from, to, sheetId = 'data', mode = 'all', copyType = 'cut') {
  const source = { unitId: book.id, sheetId: 'data', range: from }, snapshot = captureSheetClipboard(book, source, { copyType })
  const plan = createSheetClipboardPastePlan(book, { unitId: book.id, sheetId, range: to }, snapshot, mode, { userId: 'local-user' })
  const candidate = cloneModel(book); if (plan.result.accepted) applySheetClipboardPastePlan(candidate, plan)
  return { plan, candidate, snapshot }
}

test('partial down insertion preserves the untouched column geometry and keeps one logical header', () => {
  const source = [rule()], before = cloneModel(source)
  const next = planSheetRangeThemeCellShift(source, { axis: 'row', kind: 'insert', start: 2, count: 2, bandStart: 1, bandEnd: 2 })
  assert.equal(next.length, 3); assert.equal(new Set(next.map(item => item.id)).size, 3); assert.deepEqual(source, before)
  assert.equal(style(next, 0, 1).fill, '#123456'); assert.equal(style(next, 7, 1).fill, '#444444')
  assert.equal(style(next, 5, 0).fill, '#444444'); assert.equal(style(next, 6, 0), undefined)
  assert.ok(next.filter(item => item.styleRange).every(item => item.styleRange.startCol === 0))
})

test('partial right insertion retains original header columns outside the edited rows', () => {
  const source = [rule()], next = planSheetRangeThemeCellShift(source, { axis: 'column', kind: 'insert', start: 1, count: 1, bandStart: 2, bandEnd: 3 })
  assert.equal(next.length, 3); assert.ok(next.some(item => item.range.startRow === 2 && item.range.endCol === 4))
  assert.equal(style(next, 0, 3).fill, '#123456'); assert.equal(style(next, 2, 4).fill, '#fedcba')
  assert.equal(style(next, 4, 4), undefined)
})

test('delete-left and delete-up contract only membership inside their perpendicular bands', () => {
  for (const edit of [{ axis: 'column', kind: 'delete', start: 1, count: 2, bandStart: 2, bandEnd: 3 },
    { axis: 'row', kind: 'delete', start: 1, count: 2, bandStart: 1, bandEnd: 2 }]) {
    const next = planSheetRangeThemeCellShift([rule()], edit)
    assert.equal(next.length, 3); assert.equal(style(next, 5, 0).fill, '#444444')
    assert.equal(style(next, 0, 0).fill, '#123456')
  }
})

test('deleting an exhausted themed band removes it while leaving other columns sparse', () => {
  const next = planSheetRangeThemeCellShift([rule()], { axis: 'row', kind: 'delete', start: 0, count: 6, bandStart: 1, bandEnd: 2 })
  assert.equal(next.length, 2); assert.equal(style(next, 0, 1), undefined); assert.equal(style(next, 0, 0).fill, '#123456')
  assert.deepEqual(planSheetRangeThemeCellShift([rule()], { axis: 'row', kind: 'delete', start: 0, count: 6, bandStart: 0, bandEnd: 3 }), [])
})

test('invalid bands, duplicate IDs and logical geometry reject without mutating source rules', () => {
  const source = [rule()], before = cloneModel(source)
  assert.throws(() => planSheetRangeThemeCellShift(source, { axis: 'row', kind: 'insert', start: 1, count: 1, bandStart: -1, bandEnd: 2 }))
  assert.throws(() => planSheetRangeThemeAxisEdit([rule(), rule()], { axis: 'row', kind: 'insert', start: 1, count: 1 }))
  for (const value of [null, [], range(-1_048_577, 0), range(2, 1), { ...range(0, 1), endCol: 0.5 }]) assert.throws(() => assertSheetRangeThemeStyleRange(value))
  assert.throws(() => assertSheetRangeThemeStyleRange(range(1, 4, 1, 2), source[0].range), /包含/)
  assert.throws(() => workbookFromInput({ ...sample(), sheets: [{ ...sample().sheets[0], rangeThemeRules: [{ ...rule(), styleRange: range(1, 4, 1, 2) }] }] }), /包含/)
  assert.deepEqual(source, before)
})

test('insertion rejects styled membership pushed beyond Office bounds and leaves untouched bands intact', () => {
  for (const [axis, edge, bandEnd] of [['row', range(1_048_574, 1_048_575, 1, 2), 2], ['column', range(1, 2, 16_382, 16_383), 2]]) {
    const source = [{ ...rule(), range: edge }], before = cloneModel(source)
    const edit = { axis, kind: 'insert', start: 1, count: 1, bandStart: 1, bandEnd }
    assert.throws(() => planSheetRangeThemeCellShift(source, edit), /移出/)
    assert.deepEqual(source, before)
    assert.deepEqual(planSheetRangeThemeCellShift(source, { ...edit, bandStart: 3, bandEnd: 3 }), before)
  }
})

test('the existing real cell move path carries theme fragment geometry through all four directions', () => {
  const book = sample(), view = toExcelWorkbookView(book), before = cloneModel(view.sheets[0])
  for (const direction of ['right', 'down', 'left', 'up']) {
    const result = moveSheetCellRange(view.sheets[0], { startRow: 2, endRow: 3, startColumn: 1, endColumn: 2 }, direction)
    assert.equal(result.status, 'applied'); assert.equal(result.sheet.rangeThemeRules.length, 3)
    assert.equal(resolveSheetRangeThemeStyle(result.sheet, 0, 0, book).fill, '#123456')
    assert.equal(resolveSheetRangeThemeStyle(result.sheet, 5, 0, book).fill, '#444444')
  }
  assert.deepEqual(view.sheets[0], before)
})

test('a partial cut transfers physical membership without generating new headers in the source or destination', () => {
  const source = [rule()], before = cloneModel(source), from = range(2, 3, 1, 2), to = range(10, 11, 5, 6)
  const next = planSheetRangeThemeCut(source, [], from, to, false)
  assert.equal(next.sourceRules.length, 4); assert.equal(next.targetRules.length, 1); assert.deepEqual(source, before)
  assert.equal(style(next.sourceRules, 2, 1), undefined); assert.equal(style(next.sourceRules, 2, 0).fill, '#fedcba')
  assert.equal(style(next.targetRules, 10, 5).fill, '#fedcba'); assert.equal(style(next.targetRules, 11, 5).fill, '#abcdef')
  assert.deepEqual(next.targetRules[0].styleRange, range(8, 13, 4, 7))
})

test('clipped logical origins below zero survive region snapshots and model reopening', () => {
  const book = sample(), next = planSheetRangeThemeCut([rule()], [], range(4, 5, 1, 2), range(0, 1, 0, 1), false)
  const sheet = book.sheets[1]; sheet.rangeThemeRules = next.targetRules
  assert.deepEqual(sheet.rangeThemeRules[0].styleRange, range(-4, 1, -1, 2))
  const snapshot = captureSheetRangeThemeRegion(sheet, [0, 1], [0, 1], book); assertSheetRangeThemeRegionSnapshot(snapshot, 2, 2)
  assert.equal(sheetRangeThemeRegionStyle(snapshot, 0, 0).fill, '#fedcba'); assert.equal(sheetRangeThemeRegionStyle(snapshot, 1, 0).fill, '#444444')
  const reopened = workbookFromInput(book); assert.deepEqual(reopened.sheets[1].rangeThemeRules, sheet.rangeThemeRules)
})

test('a whole-rule cross-sheet cut preserves its ID and detaches metadata while replacing only the target rectangle', () => {
  const source = [rule()], target = [{ id: 'target', themeName: 'default', range: range(8, 20, 4, 10) }]
  const next = planSheetRangeThemeCut(source, target, source[0].range, range(10, 15, 5, 8), false)
  assert.deepEqual(next.sourceRules, []); assert.equal(next.targetRules.at(-1).id, 'pattern-rule'); assert.equal(next.targetRules.at(-1).styleRange, undefined)
  next.targetRules.at(-1).metadata.vendor.keep = false; assert.equal(source[0].metadata.vendor.keep, true)
  assert.deepEqual(style(next.targetRules, 9, 4), style(target, 9, 4))
  assert.equal(style(next.targetRules, 10, 5).fill, '#123456')
})

test('same-sheet overlapping cuts snapshot the source first and keep later unresolved rules as blockers', () => {
  const source = [rule(), { id: 'missing', themeName: 'unregistered', range: range(2, 3, 1, 2) }]
  const next = planSheetRangeThemeCut(source, source, range(2, 3, 1, 2), range(3, 4, 1, 2), true)
  assert.deepEqual(next.sourceRules, next.targetRules); assert.equal(style(next.targetRules, 2, 1), undefined)
  assert.equal(sheetUsedRangeThemeRule({ rangeThemeRules: next.targetRules }, range(3, 3, 1, 1)).themeName, 'unregistered')
  assert.equal(style(next.targetRules, 3, 1), undefined); assert.equal(style(next.targetRules, 0, 0).fill, '#123456')
})

test('an unthemed cut clears destination theme membership and retains the untouched target fragment style geometry', () => {
  const next = planSheetRangeThemeCut([], [rule()], range(10, 11, 1, 2), range(2, 3, 1, 2), false)
  assert.equal(style(next.targetRules, 2, 1), undefined); assert.equal(style(next.targetRules, 5, 1).fill, '#444444')
  assert.ok(next.targetRules.every(item => item.styleRange)); assert.deepEqual(next.sourceRules, [])
})

test('axis insertion and movement after a partial cut transform both physical and logical geometry', () => {
  const moved = planSheetRangeThemeCut([rule()], [], range(2, 3, 1, 2), range(10, 11, 5, 6), false).targetRules
  const inserted = planSheetRangeThemeAxisEdit(moved, { axis: 'row', kind: 'insert', start: 0, count: 2 })
  assert.deepEqual(inserted[0].range, range(12, 13, 5, 6)); assert.deepEqual(inserted[0].styleRange, range(10, 15, 4, 7))
  const shifted = planSheetRangeThemeAxisEdit(inserted, { axis: 'row', kind: 'move', start: 12, count: 2, destination: 20 })
  assert.deepEqual(shifted[0].range, range(18, 19, 5, 6)); assert.equal(style(shifted, 18, 5).fill, '#fedcba')
})

test('axis copies keep clipped source stripe and footer offsets while assigning fresh rule identities', () => {
  const next = planSheetRangeThemeAxisEdit([rule()], { axis: 'row', kind: 'copy', start: 4, count: 2, destination: 10 })
  assert.equal(next.length, 2); assert.notEqual(next[0].id, next[1].id)
  assert.equal(style(next, 10, 1).fill, '#fedcba'); assert.equal(style(next, 11, 1).fill, '#444444')
})

test('an explicit same-name reapplication replaces clipped geometry with a fresh header while keeping identity', () => {
  const book = sample(), rules = planSheetRangeThemeCut([rule()], [], range(2, 3, 1, 2), range(10, 11, 5, 6), false).targetRules
  const sheet = book.sheets[1]; sheet.rangeThemeRules = rules; const id = rules[0].id
  setSheetRangeThemeRules(sheet, [rules[0].range], 'pattern', book)
  assert.equal(sheet.rangeThemeRules[0].id, id); assert.equal(sheet.rangeThemeRules[0].styleRange, undefined)
  assert.equal(resolveSheetRangeThemeStyle(sheet, 10, 5, book).fill, '#123456')
})

test('changed ranges exclude protected retained fragments and include only actual cut or pasted membership', () => {
  const source = [rule()], from = range(2, 3, 1, 2), to = range(10, 11, 5, 6), next = planSheetRangeThemeCut(source, [], from, to, false)
  const sourceChanges = sheetRangeThemeChangedRanges(source, next.sourceRules)
  assert.ok(sourceChanges.some(item => sheetRangeContainsCell(item, 2, 1))); assert.ok(sourceChanges.every(item => !sheetRangeContainsCell(item, 5, 3)))
  assert.deepEqual(sheetRangeThemeChangedRanges([], next.targetRules), [to])
  const clones = source.map(item => ({ ...cloneModel(item), id: 'other-id' })); assert.deepEqual(sheetRangeThemeChangedRanges(source, clones), [])
})

test('the real cut planner transfers sparse theme rules and raw direct sources without allocating themed blank cells', () => {
  const book = sample(), sheet = book.sheets[0]; sheet.cells['2:1'] = createCell(2, 1, 'move'); sheet.cells['2:1'].authoredStyle = { font: { italic: true } }
  refreshWorkbookCellStyles(book, undefined, false)
  const before = cloneModel(book), { plan, candidate } = pasted(book, range(2, 3, 1, 2), range(10, 11, 5, 6))
  assert.equal(plan.result.accepted, true); assert.equal(plan.result.changed, true); assert.deepEqual(book, before)
  assert.deepEqual(Object.keys(candidate.sheets[0].cells), ['10:5']); assert.deepEqual(candidate.sheets[0].cells['10:5'].authoredStyle, { font: { italic: true } })
  assert.equal(resolveSheetRangeThemeStyle(candidate.sheets[0], 2, 1, candidate), undefined)
  assert.equal(resolveSheetRangeThemeStyle(candidate.sheets[0], 10, 5, candidate).fill, '#fedcba')
})

test('re-registering a moved theme still updates its cells and retains their direct overrides', () => {
  const book = sample(); book.sheets[0].cells['2:1'] = createCell(2, 1, 'move'); book.sheets[0].cells['2:1'].authoredStyle = { font: { bold: false } }
  refreshWorkbookCellStyles(book, undefined, false)
  const { candidate } = pasted(book, range(2, 3, 1, 2), range(10, 11, 5, 6))
  candidate.rangeThemes.pattern.secondRowStyle = { fill: '#778899', font: { bold: true } }; refreshWorkbookCellStyles(candidate, undefined, false)
  assert.equal(candidate.sheets[0].cells['10:5'].style.fill, '#778899'); assert.equal(candidate.sheets[0].cells['10:5'].style.font.bold, false)
})

test('raw theme author changes invalidate a pending cut even when their effective color is unchanged', () => {
  const book = sample(), from = range(2, 3, 1, 2), snapshot = captureSheetClipboard(book, { unitId: book.id, sheetId: 'data', range: from }, { copyType: 'cut' })
  const before = sheetClipboardSourceFingerprint(book.sheets[0], from, book); book.sheets[0].rangeThemeRules[0].metadata.vendor.keep = false
  assert.notEqual(sheetClipboardSourceFingerprint(book.sheets[0], from, book), before)
  const plan = createSheetClipboardPastePlan(book, { unitId: book.id, sheetId: 'data', range: range(10, 11, 5, 6) }, snapshot, 'all', { userId: 'local-user' })
  assert.equal(plan.result.reason, 'stale-source'); assert.deepEqual(plan.sheetChanges, [])
})

test('protected cells outside a retained fragment allow a cut while protected destinations reject it atomically', () => {
  const book = sample(); book.sheets[0].rangeProtectionRules = [{ id: 'tail', ranges: [range(5, 5, 3, 3)], allowedUserIds: ['owner'] }]
  const from = range(2, 3, 1, 2), to = range(10, 11, 5, 6)
  assert.equal(pasted(book, from, to).plan.result.accepted, true)
  book.sheets[0].rangeProtectionRules.push({ id: 'target', ranges: [to], allowedUserIds: ['owner'] }); const before = cloneModel(book)
  assert.throws(() => pasted(book, from, to), /保护/); assert.deepEqual(book, before)
})

test('same-range cuts are no-ops and ordinary copy or values-only paste keeps the existing format semantics', () => {
  const book = sample(), from = range(2, 3, 1, 2), to = range(10, 11, 5, 6)
  assert.equal(pasted(book, from, from).plan.result.changed, false)
  for (const [mode, copyType] of [['all', 'copy'], ['values', 'cut']]) {
    const { candidate } = pasted(book, from, to, 'data', mode, copyType)
    assert.deepEqual(candidate.sheets[0].rangeThemeRules, book.sheets[0].rangeThemeRules)
  }
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
function apiPaste(studio, book, from, to, sheetId = 'data', copying = false) {
  const target = { unitId: book.getId(), sheetId: 'data', range: from }
  assert.equal(studio.executeCommandSync(copying ? SHEET_CLIPBOARD_COMMANDS.copy : SHEET_CLIPBOARD_COMMANDS.cut, target).accepted, true)
  return studio.executeCommandSync(SHEET_CLIPBOARD_COMMANDS.paste, { unitId: book.getId(), sheetId, range: to })
}

test('Facade cut and paste move sparse rules in one history record and re-registration stays dynamic', () => withStudio((studio, _api, book) => {
  book.getActiveSheet().getRange('B3').setValue('move'); const before = cloneModel(book.save()), count = studio.editor.history.units[book.getId()].undo.length
  assert.equal(apiPaste(studio, book, range(2, 3, 1, 2), range(10, 11, 5, 6)).changed, true)
  assert.equal(studio.editor.history.units[book.getId()].undo.length, count + 1); assert.equal(book.getActiveSheet().getRange('F11').getValue(), 'move')
  assert.equal(book.getActiveSheet().getRange('F11').getCellStyleData().bg.rgb, '#fedcba')
  book.undo(); assert.deepEqual(book.save().sheets, before.sheets); book.redo()
  book.registerRangeTheme(book.createRangeThemeStyle('pattern', { secondRowStyle: { bg: { rgb: '#778899' } } }))
  assert.equal(book.getActiveSheet().getRange('F11').getCellStyleData().bg.rgb, '#778899')
}))

test('cross-sheet partial cuts keep logical origins in JSON save and reopened public style reads', () => withStudio((studio, api, book) => {
  assert.equal(apiPaste(studio, book, range(4, 5, 1, 2), range(0, 1, 0, 1), 'other').changed, true)
  const target = book.getSheetBySheetId('other'); assert.equal(target.getRange('A1').getCellStyleData().bg.rgb, '#fedcba')
  assert.equal(target.getRange('A2').getCellStyleData().bg.rgb, '#444444'); assert.equal(Object.keys(book.save().sheets[1].cells).length, 0)
  const reopened = api.createWorkbook({ ...book.save(), id: 'theme-structure-reopened' })
  assert.deepEqual(reopened.save().sheets[1].rangeThemeRules, book.save().sheets[1].rangeThemeRules)
  assert.equal(reopened.getSheetBySheetId('other').getRange('A2').getCellStyleData().bg.rgb, '#444444')
}))

test('blocked cut or paste commands leave the source rules and history intact', () => withStudio((studio, _api, book) => {
  const from = range(2, 3, 1, 2), to = range(10, 11, 5, 6)
  studio.editor.workbook.model.sheets[0].rangeProtectionRules = [{ id: 'destination', ranges: [to], allowedUserIds: ['owner'] }]
  const before = cloneModel(book.save()), revision = studio.getWorkbookRevision(book.getId()), count = studio.editor.history.units[book.getId()]?.undo.length ?? 0
  assert.throws(() => apiPaste(studio, book, from, to), /保护/)
  assert.deepEqual(book.save(), before); assert.equal(studio.getWorkbookRevision(book.getId()), revision)
  assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, count)
}))

test('stale theme rule metadata rejects the public paste command without creating a content step', () => withStudio((studio, _api, book) => {
  studio.executeCommandSync(SHEET_CLIPBOARD_COMMANDS.cut, { unitId: book.getId(), sheetId: 'data', range: range(2, 3, 1, 2) })
  studio.editor.workbook.model.sheets[0].rangeThemeRules[0].metadata.vendor.keep = false
  const before = cloneModel(book.save()), result = studio.executeCommandSync(SHEET_CLIPBOARD_COMMANDS.paste, { unitId: book.getId(), sheetId: 'data', range: range(10, 11, 5, 6) })
  assert.equal(result.reason, 'stale-source'); assert.deepEqual(book.save(), before); assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
}))

test('normal copy continues to paste effective formatting while leaving theme author rules in their source ranges', () => withStudio((studio, _api, book) => {
  const rules = cloneModel(book.save().sheets[0].rangeThemeRules)
  assert.equal(apiPaste(studio, book, range(2, 3, 1, 2), range(10, 11, 5, 6), 'data', true).changed, true)
  assert.deepEqual(book.save().sheets[0].rangeThemeRules, rules)
  assert.equal(book.getActiveSheet().getRange('F11').getCellStyleData().bg.rgb, '#fedcba')
}))
