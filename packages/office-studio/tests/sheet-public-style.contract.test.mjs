import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { composeSheetCellStyles, normalizeSheetStyleInput, resolveSheetStyleSource } from '../src/features/sheet-workbench/sheet-style-inheritance.ts'
import { sheetStyleDataWithPublicFields, sheetStyleToPublicData } from '../src/features/sheet-workbench/sheet-public-style.ts'
import { normalizeSheetBorderStyle, sheetBorderStyleNumber } from '../src/features/sheet-workbench/sheet-cell-border.ts'

// Draft only: no tests, type/build, browser or WASM checks were executed for this batch.
test('sparse direct style reads emit only present fixed font fields and retain explicit false and null', () => {
  const style = { font: { family: 'Inter', bold: false, underline: true, color: null }, fill: null }
  assert.deepEqual(sheetStyleToPublicData(style), { ff: 'Inter', bl: 0, ul: { s: 1 }, cl: { rgb: null }, bg: { rgb: null } })
  assert.deepEqual(sheetStyleToPublicData({}), {})
})

test('resolved default formatting uses fixed middle, overflow and numeric decoration enums', () => {
  const publicData = sheetStyleToPublicData(composeSheetCellStyles())
  assert.equal(publicData.ff, 'Aptos'); assert.equal(publicData.fs, 11); assert.equal(publicData.vt, 2); assert.equal(publicData.ht, 0)
  assert.equal(publicData.bl, 0); assert.equal(publicData.it, 0); assert.equal(publicData.tb, 1)
  assert.deepEqual(publicData.ul, { s: 0 }); assert.deepEqual(publicData.st, { s: 0 }); assert.deepEqual(publicData.n, { pattern: 'General' })
})

test('fixed decoration, color, rotation, format and unknown source extensions survive sparse readback', () => {
  const input = { ff: 'Arial', fs: 14, bl: 0, it: 1, ul: { s: 1, t: 10, c: 0, cl: { rgb: '#111111' } },
    st: { s: 0, vendor: 'keep' }, cl: { rgb: '#123456', vendor: 'font-color' }, bg: { rgb: null, vendor: 'fill' },
    ht: 6, vt: 0, tb: 0, stf: 1, tr: { a: 45, v: 0, vendor: 'rotation' }, n: { pattern: '0.00%', vendor: 'format' },
    eastAsiaFontFamily: 'SimSun', va: 3, pd: { l: 8 }, vendor: { keep: true } }
  const style = normalizeSheetStyleInput(input)
  assert.deepEqual(sheetStyleToPublicData(style), input)
})

test('source enum aliases remain readable while changed local alignment and wrapping emit their new values', () => {
  const style = normalizeSheetStyleInput({ ht: 5, vt: 0, tb: 0 })
  assert.deepEqual(sheetStyleToPublicData(style), { ht: 5, vt: 0, tb: 0 })
  style.horizontal = 'right'; style.vertical = 'bottom'; style.wrap = true; style.wrapStrategy = 'wrap'
  assert.deepEqual(sheetStyleToPublicData(style), { ht: 3, vt: 3, tb: 3 })
})

test('all fourteen fixed border line values round-trip through local pen styles', () => {
  for (let value = 0; value <= 13; value++) assert.equal(sheetBorderStyleNumber(normalizeSheetBorderStyle(value)), value)
  assert.throws(() => sheetBorderStyleNumber('invalid'), /线型/)
})

test('physical and six diagonal border fields preserve their fixed IDs, nulls, colors and source extensions', () => {
  const edge = { s: 12, cl: { rgb: '#123456', th: 2 }, vendor: { keep: true } }
  const input = { bd: { t: edge, r: null, b: edge, l: edge, tl_br: edge, tl_bc: edge, tl_mr: edge,
    bl_tr: edge, ml_tr: edge, bc_tr: edge, vendorEdge: { keep: 1 } } }
  assert.deepEqual(sheetStyleToPublicData(normalizeSheetStyleInput(input)), input)
})

test('an edited local border replaces its modeled style and color while retaining unrelated edge data', () => {
  const style = normalizeSheetStyleInput({ bd: { t: { s: 7, cl: { rgb: '#123456' }, vendor: 'keep' }, l: null } })
  style.border.top.style = 'medium'; style.border.top.color = '#654321'
  assert.deepEqual(sheetStyleToPublicData(style).bd, { t: { s: 8, cl: { rgb: '#654321' }, vendor: 'keep' }, l: null })
  style.border = {}; assert.deepEqual(sheetStyleToPublicData(style).bd, {})
})

test('removed modeled properties do not return from saved decoration, rotation, format or alignment metadata', () => {
  const style = normalizeSheetStyleInput({ ul: { s: 1, t: 10 }, n: { pattern: '0.00' }, tr: { a: 45, v: 0 }, ht: 5, tb: 0 })
  delete style.font.underline; delete style.numberFormat; delete style.rotation; delete style.rotationVertical
  delete style.horizontal; delete style.wrap; delete style.wrapStrategy
  const publicData = sheetStyleToPublicData(style)
  for (const field of ['ul', 'n', 'tr', 'ht', 'tb']) assert.equal(Object.hasOwn(publicData, field), false)
})

test('public and nested compatibility results are detached from authoritative patches and extension metadata', () => {
  const style = normalizeSheetStyleInput({ font: { family: 'Arial', vendor: { keep: true } }, ul: { s: 1, t: 10 }, vendor: { value: 1 } })
  const before = cloneModel(style), result = sheetStyleDataWithPublicFields(style)
  assert.equal(result.ff, 'Arial'); assert.equal(result.font.family, 'Arial'); assert.equal(result.font.vendor.keep, true)
  result.font.family = 'Changed'; result.font.vendor.keep = false; result.ul.t = 99; result.vendor.value = 2
  assert.deepEqual(style, before)
})

test('only owned catalog IDs resolve while legitimate own prototype-like IDs remain valid', () => {
  for (const id of ['constructor', 'toString', '__proto__']) assert.throws(() => resolveSheetStyleSource(id, { styles: {} }), /找不到/)
  const styles = Object.fromEntries([['constructor', { font: { family: 'Inter' } }]])
  assert.equal(resolveSheetStyleSource('constructor', { styles }).font.family, 'Inter')
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const withStudio = run => {
  const model = createBlankWorkbook(); model.id = 'public-style-reads'; model.calculationMode = 'manual'
  model.styles = { referenced: { font: { size: 18, underline: true }, metadata: { sourceStyle: { ul: { s: 1, t: 10 } } } } }
  const sheet = model.sheets[0]; sheet.rowStyles = { 0: { font: { family: 'Georgia' } } }; sheet.columnStyles = { 0: { font: { family: 'Courier New' } } }
  sheet.cells['0:0'] = createCell(0, 0, 'abc'); sheet.cells['0:0'].authoredStyleReference = 'referenced'
  const studio = new OfficeStudio({ workbook: model })
  try { return run(studio, FOffice.newAPI(studio)) } finally { studio.dispose() }
}

test('Facade row, col and direct reads provide fixed ff/fs fields together with local nested compatibility', () => withStudio((_studio, api) => {
  const sheet = api.getActiveWorkbook().getActiveSheet(), range = sheet.getRange('A1')
  assert.equal(range.getCellStyleData().ff, 'Georgia'); assert.equal(range.getCellStyleData('col').ff, 'Courier New')
  const direct = range.getCellStyleData('cell')
  assert.equal(direct.ff, undefined); assert.equal(direct.fs, 18); assert.equal(direct.font.size, 18); assert.equal(direct.ul.t, 10)
  assert.equal(range.getFontFamily('cell'), null); assert.equal(sheet.getRange('B2').getCellStyleData('cell'), null)
}))

test('Facade reset masks remove fixed aliases and readback remains usable as public cell.s with history and reopen', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet(), range = sheet.getRange('A1')
  range.setFontSize(null); assert.equal(range.getCellStyleData('cell').fs, undefined)
  book.undo(); const style = range.getCellStyleData('cell')
  sheet.getRange('B2').setValue({ v: 'copied format', s: style })
  assert.equal(sheet.getRange('B2').getCellStyleData('cell').fs, 18); assert.equal(sheet.getRange('B2').getCellStyleData('cell').ul.t, 10)
  book.undo(); assert.equal(sheet.getRange('B2').getValue(), null); book.redo()
  const reopened = api.createWorkbook({ ...book.save(), id: 'public-style-read-reopened' }).getActiveSheet()
  assert.equal(reopened.getRange('B2').getCellStyleData('cell').fs, 18)
  const copied = studio.getWorkbookModel('public-style-read-reopened').sheets[0].cells['1:1']
  assert.equal(copied.authoredStyle.font.size, 18); assert.equal(copied.authoredStyle.metadata.sourceStyle.ul.t, 10)
  assert.equal(copied.authoredStyleReference, undefined)
}))

test('public style reads and caller mutation have no content revision or undo side effects', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), range = book.getActiveSheet().getRange('A1'), before = cloneModel(book.save())
  const revision = studio.getWorkbookRevision(book.getId()), result = range.getCellStyleData()
  result.ff = 'Changed'; result.font.family = 'Changed'; result.ul.t = 99
  assert.deepEqual(book.save(), before); assert.equal(studio.getWorkbookRevision(book.getId()), revision)
  assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
}))
