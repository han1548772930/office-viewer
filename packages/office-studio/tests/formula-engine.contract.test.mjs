import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let formula
let model

before(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  formula = await server.ssrLoadModule('/packages/office-studio/src/features/formula/index.ts')
  const domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  model = domain.createBlankWorkbook()
  const sheet = model.sheets[0]
  sheet.id = 'sheet-main'
  sheet.name = 'Data'
  const put = (row, col, value, formulaText) => {
    const cell = domain.createCell(row, col, value)
    cell.formula = formulaText ?? null
    cell.display = formulaText ? '' : value == null ? '' : String(value)
    sheet.cells[`${row}:${col}`] = cell
  }
  put(0, 0, 2)
  put(0, 1, 3)
  put(0, 2, null, '=SUM(A1:B1)')
  put(0, 3, null, '=C1*2')
  put(0, 4, null, '=IF(D1>=10,"ok","bad")')
})

after(async () => { await server?.close() })

test('parser preserves references, ranges, functions and absolute markers', () => {
  const node = formula.parseFormula('=SUM($A1:B$2)+Data!C3')
  assert.equal(node.kind, 'binary')
  assert.deepEqual(node.left.kind, 'function')
  assert.deepEqual(node.right.kind, 'reference')
  assert.equal(node.right.reference.sheetName, 'Data')
  assert.equal(node.left.arguments[0].kind, 'range')
  assert.equal(node.left.arguments[0].reference.start.absoluteColumn, true)
  assert.equal(node.left.arguments[0].reference.end.absoluteRow, true)
})

test('recalculation updates dependent cached values and builds reverse dependency graph', () => {
  const result = formula.recalculateWorkbookFormulas(model)
  const cells = result.workbook.sheets[0].cells
  assert.equal(cells['0:2'].value, 5)
  assert.equal(cells['0:2'].display, '5')
  assert.equal(cells['0:3'].value, 10)
  assert.equal(cells['0:4'].value, 'ok')
  const sourceKey = 'sheet-main!0:0'
  const dependentKey = 'sheet-main!0:2'
  assert.ok(result.graph.dependents.get(sourceKey)?.has(dependentKey))
  assert.deepEqual(result.diagnostics, [])
})

test('errors and cycles are represented in cached values without throwing or losing formulas', () => {
  const cyclic = structuredClone(model)
  cyclic.sheets[0].cells['0:2'].formula = '=D1'
  cyclic.sheets[0].cells['0:3'].formula = '=C1'
  const result = formula.recalculateWorkbookFormulas(cyclic)
  assert.equal(result.workbook.sheets[0].cells['0:2'].formula, '=D1')
  assert.equal(result.workbook.sheets[0].cells['0:2'].value, '#CIRCULAR!')
  assert.equal(result.workbook.sheets[0].cells['0:3'].value, '#CIRCULAR!')
  const bad = formula.evaluateFormula('=1/0', model, 'sheet-main')
  assert.deepEqual(bad, { kind: 'error', code: '#DIV/0!' })
})

test('reference translation shifts only relative axes and emits ref for boundary underflow', () => {
  assert.equal(formula.translateFormulaReferences('=A1+$B$2+C$3+$D4', 2, 1), '=B3+$B$2+D$3+$D6')
  assert.equal(formula.translateFormulaReferences('=A1', -1, 0), '=#REF!')
  assert.equal(formula.translateFormulaReferences('="A1"&A1', 1, 0), '="A1"&A2')
})

test('cross-sheet dependencies resolve by sheet name and unsupported imports preserve cached values', () => {
  const imported = structuredClone(model)
  const other = { ...structuredClone(imported.sheets[0]), id: 'sheet-other', name: 'Other', cells: {} }
  const domainCell = imported.sheets[0].cells['0:2']
  domainCell.formula = '=Other!A1+1'
  other.cells['0:0'] = { ...imported.sheets[0].cells['0:0'], id: '0:0', row: 0, col: 0, value: 7, display: '7', formula: null }
  imported.sheets.push(other)
  const first = formula.recalculateWorkbookFormulas(imported)
  assert.equal(first.workbook.sheets[0].cells['0:2'].value, 8)
  const cached = structuredClone(imported)
  cached.sheets[0].cells['0:2'].formula = '=UNSUPPORTED_CUSTOM(A1)'
  cached.sheets[0].cells['0:2'].value = 42
  cached.sheets[0].cells['0:2'].display = '42'
  const preserved = formula.recalculateWorkbookFormulas(cached, cached)
  assert.equal(preserved.workbook.sheets[0].cells['0:2'].value, 42)
  assert.ok(preserved.diagnostics.some((entry) => entry.code === 'FORMULA_UNSUPPORTED_CACHE_PRESERVED'))
})

