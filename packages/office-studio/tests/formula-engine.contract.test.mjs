import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let formula
let model
let functionCatalog

beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  formula = await server.ssrLoadModule('/packages/office-studio/src/features/formula/index.ts')
  functionCatalog = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-workbench/sheet-function-catalog.ts')
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

afterAll(async () => { await server?.close() })

test('statistical aggregates retain fixed scalar conversions, empty data, arity and canonical function entries', () => {
  const examples = [
    ['AVEDEV', '2,4', 1], ['DEVSQ', '2,4', 2], ['GEOMEAN', '4,9', 6],
    ['HARMEAN', '2,6', 3], ['MODE.SNGL', '2,4,4', 4],
  ]
  for (const [name, argumentsText, expected] of examples) {
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '统计')
    assert.equal(entry.example, argumentsText)
    assert.ok(Math.abs(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, argumentsText), model, 'sheet-main') - expected) < 1e-12)
    for (const invalid of ['', Array(256).fill('1').join(',')]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${invalid})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    }
    assert.equal(formula.evaluateFormula(`=${name}(${Array(255).fill('1').join(',')})`, model, 'sheet-main'), name === 'AVEDEV' || name === 'DEVSQ' ? 0 : 1)
    assert.deepEqual(formula.evaluateFormula(`=${name}(#REF!)`, model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  }
  for (const [expression, expected] of [
    ['AVEDEV(1,"2")', 0.5], ['AVEDEV(2,TRUE)', 0.5], ['AVEDEV(2,FALSE)', 1], ['AVEDEV("50%",1)', 0.25],
    ['DEVSQ("2","4")', 2], ['DEVSQ(2,TRUE,FALSE,"text"," ",A2)', 0], ['DEVSQ(2,"0x10")', 98],
    ['GEOMEAN("4","9")', 6], ['GEOMEAN(4,"25%","2012-2-2",TRUE)', 4],
    ['HARMEAN("2","6")', 3], ['HARMEAN(4,"25%","2012-2-2",FALSE)', 4],
    ['MODE.SNGL(2,"2",2)', 2], ['MODE.SNGL(5,2,2,5)', 5], ['MODE.SNGL(2,4,4,2)', 2],
  ]) assert.ok(Math.abs(formula.evaluateFormula(`=${expression}`, model, 'sheet-main') - expected) < 1e-12, expression)
  for (const [expression, code] of [
    ['AVEDEV("text")', '#VALUE!'], ['AVEDEV(A2)', '#NUM!'], ['DEVSQ("text",TRUE,A2)', '#NUM!'],
    ['GEOMEAN("text",2)', '#VALUE!'], ['HARMEAN("text",2)', '#VALUE!'], ['MODE.SNGL("text",2,2)', '#VALUE!'],
    ['GEOMEAN(TRUE,A2)', '#NUM!'], ['HARMEAN(TRUE,A2)', '#N/A'], ['MODE.SNGL("2","2")', '#N/A'],
    ['GEOMEAN(0,2)', '#NUM!'], ['HARMEAN(-1,2)', '#NUM!'], ['MODE.SNGL(1,2)', '#N/A'],
    ['GEOMEAN(-1,#REF!)', '#REF!'], ['HARMEAN(-1,#REF!)', '#REF!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=${expression}`, model, 'sheet-main'), { kind: 'error', code }, expression)
})

test('statistical aggregates preserve row-order errors, typed reference text, INDEX and first-occurrence ties', () => {
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  const template = sheet.cells['0:0']
  sheet.cells = {}
  for (const [row, col, value, dataType] of [
    [1, 2, 4, 'number'], [0, 2, '4', 'text'], [1, 1, 'skip', 'text'],
    [0, 1, 4, 'number'], [1, 0, true, 'boolean'], [0, 0, 2, 'number'],
    [3, 1, '#REF!', 'error'], [3, 0, '#DIV/0!', 'error'],
  ]) sheet.cells[`${row}:${col}`] = { ...template, id: `${row}:${col}`, row, col, value, dataType, formula: null, display: String(value) }
  const cases = [['AVEDEV', 8 / 9], ['DEVSQ', 3], ['GEOMEAN', 128 ** 0.25], ['HARMEAN', 3.2], ['MODE.SNGL', 4]]
  for (const [name, expected] of cases) {
    assert.ok(Math.abs(formula.evaluateFormula(`=${name}(A1:C2)`, source, 'sheet-main') - expected) < 1e-12, name)
    assert.deepEqual(formula.evaluateFormula(`=${name}(A4:B4)`, source, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  }
  for (const reference of ['C1', 'C1:C1', 'INDEX(A1:C2,1,3)']) {
    for (const [name, expected] of [['DEVSQ', 0], ['GEOMEAN', 4], ['HARMEAN', 4]]) {
      assert.equal(formula.evaluateFormula(`=${name}(${reference})`, source, 'sheet-main'), expected)
    }
    for (const [name, code] of [['AVEDEV', '#NUM!'], ['MODE.SNGL', '#N/A']]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${reference})`, source, 'sheet-main'), { kind: 'error', code })
    }
  }
  sheet.cells['1:0'].formula = '=TRUE()'
  assert.ok(Math.abs(formula.evaluateFormula('=AVEDEV(A1:C2)', source, 'sheet-main') - 8 / 9) < 1e-12)
  for (const pattern of ['@', '@@@']) {
    sheet.cells['0:1'] = { ...sheet.cells['0:1'], style: { ...template.style, numberFormat: pattern } }
    assert.equal(formula.evaluateFormula('=AVEDEV(A1:C2)', source, 'sheet-main'), 1)
    assert.deepEqual(formula.evaluateFormula('=MODE.SNGL(A1:C2)', source, 'sheet-main'), { kind: 'error', code: '#N/A' })
    assert.equal(formula.evaluateFormula('=DEVSQ(A1:C2)', source, 'sheet-main'), 3)
  }
  for (const [row, col, value] of [[7, 2, 6], [7, 1, 5], [7, 0, 4], [6, 2, 4], [6, 1, 3], [6, 0, 2], [5, 2, 3], [5, 1, 2], [5, 0, 1]]) {
    sheet.cells[`${row}:${col}`] = { ...template, id: `${row}:${col}`, row, col, value, dataType: 'number', formula: null, display: String(value) }
  }
  assert.equal(formula.evaluateFormula('=MODE.SNGL(A6:C8)', source, 'sheet-main'), 2)
  for (const name of ['GEOMEAN', 'HARMEAN']) {
    assert.deepEqual(formula.evaluateFormula(`=${name}(-1,A4:B4)`, source, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  }
})

test('statistical aggregates retain AVEDEV phase errors and unguarded raw numeric results', () => {
  const source = structuredClone(model)
  source.sheets[0].cells['0:0'].value = 1e308
  source.sheets[0].cells['0:1'].value = 1e308
  for (const [expression, code] of [
    ['AVEDEV(1e308,1e308)', '#VALUE!'], ['AVEDEV(1e308,1e308,#REF!)', '#REF!'],
    ['AVEDEV(A1:B1,#REF!)', '#VALUE!'], ['AVEDEV(A1:A1,B1:B1,#REF!)', '#NUM!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=${expression}`, source, 'sheet-main'), { kind: 'error', code }, expression)
  assert.equal(formula.evaluateFormula('=DEVSQ(1e308,-1e308)', source, 'sheet-main'), Infinity)
  assert.equal(formula.evaluateFormula('=GEOMEAN(1e308,1e308)', source, 'sheet-main'), Infinity)
  assert.equal(formula.evaluateFormula('=HARMEAN(QUOTIENT("1e999",1),QUOTIENT("1e999",1))', source, 'sheet-main'), Infinity)
  for (const [name, suffix, expected] of [['DEVSQ', '2', 0], ['GEOMEAN', '2', 2], ['HARMEAN', '2', 2], ['MODE.SNGL', '2,2', 2]]) {
    assert.equal(formula.evaluateFormula(`=${name}(QUOTIENT("1e999","1e999"),${suffix})`, source, 'sheet-main'), expected)
  }
  assert.deepEqual(formula.evaluateFormula('=AVEDEV(QUOTIENT("1e999","1e999"),2)', source, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
})

test('ISDATE reads date and time text with the fixed date system and preserves scalar reference formats', () => {
  for (const [argument, expected] of [
    ['"2012-2-2"', true], ['"10:30:00"', true], ['"31/01/2024"', false],
    ['"test"', false], ['"50%"', false], ['40826.4375', false], ['TRUE', false], ['A2', false],
  ]) {
    assert.equal(formula.evaluateFormula(`=ISDATE(${argument})`, model, 'sheet-main', { numberFormatLocale: 'fr-FR' }), expected, argument)
  }
  assert.equal(formula.evaluateFormula('=ISDATE("01/01/1900")', model, 'sheet-main'), true)
  assert.equal(formula.evaluateFormula('=ISDATE("01/01/1900")', { ...model, date1904: true }, 'sheet-main'), false)
  for (const argument of ['', '1,2']) assert.deepEqual(formula.evaluateFormula(`=ISDATE(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.deepEqual(formula.evaluateFormula('=ISDATE(#REF!)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.deepEqual(formula.evaluateFormula('=ISDATE(A1:B1)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })

  const source = structuredClone(model)
  source.sheets[0].cells['0:0'].value = 40826.4375
  for (const [pattern, expected] of [
    ['General', false], ['yyyy-mm-dd hh:mm:ss', true], ['hh:mm:ss', true], ['[h]:mm:ss', true],
    ['0"yyyy"', false], ['0.00;0.00;yyyy-mm-dd', true], ['@', false], ['@@@', false],
  ]) {
    source.sheets[0].cells['0:0'].style.numberFormat = pattern
    for (const reference of ['A1', 'A1:A1', 'INDEX(A1:B1,1,1)']) {
      assert.equal(formula.evaluateFormula(`=ISDATE(${reference})`, source, 'sheet-main'), expected, `${pattern}: ${reference}`)
    }
  }
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'ISDATE')
  assert.equal(entry.category, '信息')
  assert.equal(entry.example, 'A1')
  assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula('ISDATE', '"2012-2-2"'), model, 'sheet-main'), true)
})

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

test('the fixed common projection preserves canonical categories and builds formulas evaluated by the existing engine', () => {
  const cases = [
    ['SUMIF', '数学', 'A1:B1,">0"', 5],
    ['SUM', '数学', 'A1:B1', 5],
    ['AVERAGE', '统计', 'A1:B1', 2.5],
    ['IF', '逻辑', 'A1>0,2,3', 2],
    ['COUNT', '统计', 'A1:B1', 2],
    ['SIN', '数学', '0', 0],
    ['MAX', '统计', 'A1:B1', 3],
  ]
  assert.deepEqual(functionCatalog.SHEET_COMMON_FUNCTION_CATALOG.map(entry => entry.name), cases.map(([name]) => name))
  for (const [name, category, argumentsText, value] of cases) {
    const entry = functionCatalog.SHEET_COMMON_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry, functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name))
    assert.equal(entry.category, category)
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, argumentsText), model, 'sheet-main'), value)
  }
  for (const name of ['COUNTA', 'MIN']) {
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name).category, '统计')
    assert.ok(!functionCatalog.SHEET_COMMON_FUNCTION_CATALOG.some(entry => entry.name === name))
  }
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

test('structural edits update only references on the edited sheet and preserve reference qualifiers', () => {
  const right = { kind: 'insert', axis: 'column', start: 1, count: 2, bandStart: 0, bandEnd: 0 }
  assert.equal(formula.transformFormulaReferencesForStructureEdit('=SUM(A1:C1)+Other!B1+"B1"', right, 'Data', 'Data'), '=SUM(A1:E1)+Other!B1+"B1"')
  assert.equal(formula.transformFormulaReferencesForStructureEdit("='Data'!$B$1", right, 'Other', 'Data'), "='Data'!$D$1")

  const deleteColumn = { kind: 'delete', axis: 'column', start: 1, count: 1, bandStart: 0, bandEnd: 0 }
  assert.equal(formula.transformFormulaReferencesForStructureEdit('=B1', deleteColumn, 'Data', 'Data'), '=#REF!')
  assert.equal(formula.transformFormulaReferencesForStructureEdit('=A1:B1', deleteColumn, 'Data', 'Data'), '=A1:A1')
  assert.equal(formula.transformFormulaReferencesForStructureEdit('=A1:C1', deleteColumn, 'Data', 'Data'), '=A1:B1')
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

test('lookup functions resolve exact positions, table columns, blanks and invalid indexes', () => {
  const data = structuredClone(model)
  const sheet = data.sheets[0]
  const put = (row, col, value) => {
    const cell = { ...sheet.cells['0:0'], id: `${row}:${col}`, row, col, value, formula: null, display: value == null ? '' : String(value) }
    sheet.cells[`${row}:${col}`] = cell
  }
  put(0, 0, 'East'); put(0, 1, 100); put(0, 2, 'West'); put(0, 3, 200)
  put(1, 0, 'North'); put(1, 1, 300); put(1, 2, 'South'); put(1, 3, 400)
  assert.equal(formula.evaluateFormula('=MATCH("West",A1:C1,0)', data, 'sheet-main'), 3)
  assert.equal(formula.evaluateFormula('=INDEX(A1:D2,2,4)', data, 'sheet-main'), 400)
  assert.equal(formula.evaluateFormula('=VLOOKUP("North",A1:B2,2,FALSE)', data, 'sheet-main'), 300)
  assert.deepEqual(formula.evaluateFormula('=INDEX(A1:B2,3,1)', data, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.deepEqual(formula.evaluateFormula('=MATCH("Missing",A1:A2,0)', data, 'sheet-main'), { kind: 'error', code: '#N/A' })
})

test('reference functions format addresses and resolve scalar row, column and area counts', () => {
  assert.equal(formula.evaluateFormula('=ADDRESS(2,3)', model, 'sheet-main'), '$C$2')
  assert.equal(formula.evaluateFormula('=ADDRESS(2,3,2)', model, 'sheet-main'), 'C$2')
  assert.equal(formula.evaluateFormula('=ADDRESS(2,3,3)', model, 'sheet-main'), '$C2')
  assert.equal(formula.evaluateFormula('=ADDRESS(2,3,4)', model, 'sheet-main'), 'C2')
  assert.equal(formula.evaluateFormula('=ADDRESS(2,3,2,FALSE)', model, 'sheet-main'), 'R2C[3]')
  assert.equal(formula.evaluateFormula('=ADDRESS(1,1,1,TRUE,"Sheet\'1")', model, 'sheet-main'), "'Sheet''1'!$A$1")
  assert.equal(formula.evaluateFormula('=ADDRESS(1,1,1,TRUE,"A1")', model, 'sheet-main'), "'A1'!$A$1")
  assert.deepEqual(formula.evaluateFormula('=ADDRESS(0,1)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.deepEqual(formula.evaluateFormula('=ADDRESS(2,3,5)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })

  assert.equal(formula.evaluateFormula('=ROW()', model, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=COLUMN()', model, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=ROW(B7)', model, 'sheet-main'), 7)
  assert.equal(formula.evaluateFormula('=COLUMN(B7)', model, 'sheet-main'), 2)
  assert.deepEqual(formula.evaluateFormula('=ROW(B7:B8)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=COLUMN(B7:C7)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.equal(formula.evaluateFormula('=AREAS(B7:C8)', model, 'sheet-main'), 1)
  assert.deepEqual(formula.evaluateFormula('=AREAS("B7")', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })

  const positioned = structuredClone(model)
  const sheet = positioned.sheets[0]
  for (const [row, col, formulaText] of [[4, 3, '=ROW()'], [4, 4, '=COLUMN()'], [5, 3, '=IF(TRUE,ROW(),0)']]) {
    const cell = { ...sheet.cells['0:0'], id: `${row}:${col}`, row, col, value: null, formula: formulaText, display: '' }
    sheet.cells[`${row}:${col}`] = cell
  }
  const recalculated = formula.recalculateWorkbookFormulas(positioned).workbook.sheets[0]
  assert.equal(recalculated.cells['4:3'].value, 5)
  assert.equal(recalculated.cells['4:4'].value, 5)
  assert.equal(recalculated.cells['5:3'].value, 6)

  for (const name of ['ADDRESS', 'AREAS', 'COLUMN', 'ROW']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
})

test('financial and statistical functions calculate finite results and reject invalid domains', () => {
  const data = structuredClone(model)
  const sheet = data.sheets[0]
  const put = (row, col, value) => {
    const cell = { ...sheet.cells['0:0'], id: `${row}:${col}`, row, col, value, formula: null, display: value == null ? '' : String(value) }
    sheet.cells[`${row}:${col}`] = cell
  }
  put(0, 0, 1); put(1, 0, 2); put(2, 0, 3); put(3, 0, 10)
  assert.equal(formula.evaluateFormula('=MEDIAN(A1:A4)', data, 'sheet-main'), 2.5)
  assert.equal(formula.evaluateFormula('=LARGE(A1:A4,2)', data, 'sheet-main'), 3)
  assert.equal(formula.evaluateFormula('=SMALL(A1:A4,2)', data, 'sheet-main'), 2)
  assert.equal(formula.evaluateFormula('=VAR.S(A1:A4)', data, 'sheet-main'), 50 / 3)
  assert.equal(formula.evaluateFormula('=STDEV.S(A1:A4)', data, 'sheet-main'), Math.sqrt(50 / 3))
  assert.equal(formula.evaluateFormula('=PMT(0,4,300)', data, 'sheet-main'), -75)
  assert.equal(formula.evaluateFormula('=PV(0,4,-75)', data, 'sheet-main'), 300)
  assert.equal(formula.evaluateFormula('=FV(0,4,-75)', data, 'sheet-main'), 300)
  assert.equal(formula.evaluateFormula('=NPV(10%,100,100)', data, 'sheet-main'), 100 / 1.1 + 100 / 1.21)
  assert.deepEqual(formula.evaluateFormula('=STDEV.S(A1:A1)', data, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.deepEqual(formula.evaluateFormula('=LARGE(A1:A4,0)', data, 'sheet-main'), { kind: 'error', code: '#NUM!' })
})

test('product, population statistics, percentiles and quartiles are executable catalog functions', () => {
  const data = structuredClone(model)
  const sheet = data.sheets[0]
  const put = (row, col, value) => {
    const cell = { ...sheet.cells['0:0'], id: `${row}:${col}`, row, col, value, formula: null, display: value == null ? '' : String(value) }
    sheet.cells[`${row}:${col}`] = cell
  }
  put(0, 0, 1); put(1, 0, 2); put(2, 0, 3); put(3, 0, 4); put(4, 0, 'text')

  assert.equal(formula.evaluateFormula('=PRODUCT(A1:A5)', data, 'sheet-main'), 24)
  assert.equal(formula.evaluateFormula('=PRODUCT(2,3,4)', data, 'sheet-main'), 24)
  assert.equal(formula.evaluateFormula('=VAR.P(A1:A4)', data, 'sheet-main'), 1.25)
  assert.equal(formula.evaluateFormula('=STDEV.P(A1:A4)', data, 'sheet-main'), Math.sqrt(1.25))
  assert.equal(formula.evaluateFormula('=PERCENTILE.INC(A1:A4,0.25)', data, 'sheet-main'), 1.75)
  assert.equal(formula.evaluateFormula('=PERCENTILE.EXC(A1:A4,0.4)', data, 'sheet-main'), 2)
  assert.equal(formula.evaluateFormula('=QUARTILE.INC(A1:A4,1)', data, 'sheet-main'), 1.75)
  assert.equal(formula.evaluateFormula('=QUARTILE.INC(A1:A4,1.5)', data, 'sheet-main'), 2.125)
  assert.equal(formula.evaluateFormula('=QUARTILE.EXC(A1:A4,1)', data, 'sheet-main'), 1.25)
  assert.deepEqual(formula.evaluateFormula('=PERCENTILE.EXC(A1:A4,0.1)', data, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(formula.evaluateFormula('=QUARTILE.INC(A1:A4,5)', data, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(formula.evaluateFormula('=VAR.S(A1:A1)', data, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })

  for (const name of ['PRODUCT', 'VAR.P', 'STDEV.P', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.INC', 'QUARTILE.EXC']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
})

test('logical, text, date and lookup additions preserve argument, range and error semantics', () => {
  const data = structuredClone(model)
  const sheet = data.sheets[0]
  const put = (row, col, value) => {
    const cell = { ...sheet.cells['0:0'], id: `${row}:${col}`, row, col, value, formula: null, display: value == null ? '' : String(value) }
    sheet.cells[`${row}:${col}`] = cell
  }
  put(0, 0, 'Q1'); put(0, 1, 'Q2'); put(0, 2, 'Q3')
  put(1, 0, 10); put(1, 1, 20); put(1, 2, 30)
  put(2, 0, 'a'); put(2, 1, ''); put(2, 2, 'c')

  assert.equal(formula.evaluateFormula('=IFNA(MATCH("missing",A1:C1,0),"none")', data, 'sheet-main'), 'none')
  assert.deepEqual(formula.evaluateFormula('=IFNA(1/0,5)', data, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.equal(formula.evaluateFormula('=IFS(A1="Q2",1,B1="Q2",2)', data, 'sheet-main'), 2)
  assert.deepEqual(formula.evaluateFormula('=IFS(FALSE,1,FALSE,2)', data, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(formula.evaluateFormula('=XOR(TRUE,FALSE,TRUE)', data, 'sheet-main'), false)
  assert.equal(formula.evaluateFormula('=HLOOKUP("Q2",A1:C2,2,FALSE)', data, 'sheet-main'), 20)
  assert.equal(formula.evaluateFormula('=CHOOSE(2,"one","two","three")', data, 'sheet-main'), 'two')
  assert.equal(formula.evaluateFormula('=ROWS(A1:C3)', data, 'sheet-main'), 3)
  assert.equal(formula.evaluateFormula('=COLUMNS(A1:C3)', data, 'sheet-main'), 3)

  assert.equal(formula.evaluateFormula('=TEXTJOIN("-",TRUE,A3:C3)', data, 'sheet-main'), 'a-c')
  assert.equal(formula.evaluateFormula('=FIND("b","abc")', data, 'sheet-main'), 2)
  assert.equal(formula.evaluateFormula('=SEARCH("B","abc")', data, 'sheet-main'), 2)
  assert.equal(formula.evaluateFormula('=REPLACE("abcdef",2,3,"XY")', data, 'sheet-main'), 'aXYef')
  assert.equal(formula.evaluateFormula('=SUBSTITUTE("one one","one","two",2)', data, 'sheet-main'), 'one two')
  assert.equal(formula.evaluateFormula('=REPT("x",3)', data, 'sheet-main'), 'xxx')
  assert.deepEqual(formula.evaluateFormula('=FIND("z","abc")', data, 'sheet-main'), { kind: 'error', code: '#VALUE!' })

  assert.equal(formula.evaluateFormula('=EDATE(DATE(2024,1,31),1)', data, 'sheet-main'), 45351)
  assert.equal(formula.evaluateFormula('=EOMONTH(DATE(2024,2,1),0)', data, 'sheet-main'), 45351)
  assert.equal(formula.evaluateFormula('=ISOWEEKNUM(DATE(2021,1,1))', data, 'sheet-main'), 53)
  assert.equal(formula.evaluateFormula('=DAYS360(DATE(2024,1,1),DATE(2024,2,1))', data, 'sheet-main'), 30)
  assert.equal(formula.evaluateFormula('=DAYS360(DATE(2024,2,29),DATE(2024,3,31))', data, 'sheet-main'), 30)
  assert.equal(formula.evaluateFormula('=DAYS360(DATE(2024,2,29),DATE(2024,3,31),TRUE)', data, 'sheet-main'), 31)
  assert.equal(formula.evaluateFormula('=DAYS360(DATE(2024,3,31),DATE(2024,2,29))', data, 'sheet-main'), -30)
  assert.equal(formula.evaluateFormula('=DAYS360(DATE(2024,3,31),DATE(2024,2,29),TRUE)', data, 'sheet-main'), -31)

  for (const name of ['IFNA', 'IFS', 'XOR', 'HLOOKUP', 'CHOOSE', 'ROWS', 'COLUMNS', 'TEXTJOIN', 'FIND', 'SEARCH', 'REPLACE', 'SUBSTITUTE', 'REPT', 'EDATE', 'EOMONTH', 'ISOWEEKNUM', 'DAYS360']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
})

test('business-day and year/date functions honor weekend codes, holidays and workbook date system', () => {
  const data = structuredClone(model)
  const sheet = data.sheets[0]
  const put = (row, col, value) => {
    const cell = { ...sheet.cells['0:0'], id: `${row}:${col}`, row, col, value, formula: null, display: value == null ? '' : String(value) }
    sheet.cells[`${row}:${col}`] = cell
  }
  put(0, 0, 45292)
  put(1, 0, 45293)

  assert.equal(formula.evaluateFormula('=WEEKNUM(DATE(2021,1,1),21)', data, 'sheet-main'), 53)
  assert.equal(formula.evaluateFormula('=WEEKNUM(DATE(2021,1,1),2)', data, 'sheet-main'), 1)
  assert.deepEqual(formula.evaluateFormula('=WEEKNUM(DATE(2021,1,1),3)', data, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.equal(formula.evaluateFormula('=YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),1)', data, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),3)', data, 'sheet-main'), 366 / 365)
  assert.deepEqual(formula.evaluateFormula('=YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),5)', data, 'sheet-main'), { kind: 'error', code: '#NUM!' })

  assert.equal(formula.evaluateFormula('=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7))', data, 'sheet-main'), 5)
  assert.equal(formula.evaluateFormula('=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),11)', data, 'sheet-main'), 6)
  assert.equal(formula.evaluateFormula('=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),"0000011")', data, 'sheet-main'), 5)
  assert.equal(formula.evaluateFormula('=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7),A1:A2)', data, 'sheet-main'), 3)
  assert.equal(formula.evaluateFormula('=NETWORKDAYS(DATE(2024,1,7),DATE(2024,1,1))', data, 'sheet-main'), -5)
  assert.equal(formula.evaluateFormula('=WORKDAY(DATE(2024,1,5),1)', data, 'sheet-main'), 45299)
  assert.equal(formula.evaluateFormula('=WORKDAY(DATE(2024,1,5),-1)', data, 'sheet-main'), 45295)
  assert.equal(formula.evaluateFormula('=WORKDAY.INTL(DATE(2024,1,5),1,11)', data, 'sheet-main'), 45297)
  assert.equal(formula.evaluateFormula('=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),"1111111")', data, 'sheet-main'), 0)
  assert.deepEqual(formula.evaluateFormula('=WORKDAY.INTL(DATE(2024,1,1),1,"1111111")', data, 'sheet-main'), { kind: 'error', code: '#NUM!' })

  const date1904 = structuredClone(data)
  date1904.date1904 = true
  assert.equal(formula.evaluateFormula('=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7))', date1904, 'sheet-main'), 5)
  assert.equal(formula.evaluateFormula('=WEEKNUM(DATE(2021,1,1),21)', date1904, 'sheet-main'), 53)

  for (const name of ['WEEKNUM', 'YEARFRAC', 'NETWORKDAYS', 'NETWORKDAYS.INTL', 'WORKDAY', 'WORKDAY.INTL']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
})

test('DATEDIF, NOW and TODAY follow date-system and recalculation clock semantics', () => {
  const data = structuredClone(model)
  const evaluate = (formulaText, workbook = data, now) => formula.evaluateFormula(formulaText, workbook, 'sheet-main', { now })
  const start = 'DATE(2011,1,29)'
  const end = 'DATE(2021,3,31)'
  assert.equal(evaluate(`=DATEDIF(${start},${end},"Y")`), 10)
  assert.equal(evaluate(`=DATEDIF(${start},${end},"M")`), 122)
  assert.equal(evaluate(`=DATEDIF(${start},${end},"D")`), 3714)
  assert.equal(evaluate(`=DATEDIF(${start},${end},"MD")`), 2)
  assert.equal(evaluate(`=DATEDIF(${start},${end},"YM")`), 2)
  assert.equal(evaluate(`=DATEDIF(${start},${end},"YD")`), 61)
  assert.deepEqual(evaluate('=DATEDIF(DATE(2021,1,1),DATE(2020,1,1),"Y")'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(evaluate('=DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"YY")'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(evaluate('=DATEDIF(DATE(2020,1,1),DATE(2021,1,1),1)'), { kind: 'error', code: '#NUM!' })
  assert.equal(evaluate('=DATEDIF(60,90,"MD")'), 1)

  const date1904 = structuredClone(data)
  date1904.date1904 = true
  assert.equal(evaluate('=DATEDIF(0,60,"M")', date1904), 2)

  const clock = new Date(2024, 0, 2, 3, 4, 5, 987)
  const daySerial = evaluate('=DATE(2024,1,2)')
  const expectedNow = daySerial + (3 * 3600 + 4 * 60 + 5) / 86400
  assert.equal(evaluate('=TODAY()', data, clock), daySerial)
  assert.ok(Math.abs(evaluate('=NOW()', data, clock) - expectedNow) < 1e-12)
  assert.deepEqual(evaluate('=TODAY(1)', data, clock), { kind: 'error', code: '#N/A' })

  const recalculation = structuredClone(data)
  const sheet = recalculation.sheets[0]
  for (const [row, formulaText] of [[2, '=NOW()'], [3, '=TODAY()'], [4, '=NOW()-TODAY()']]) {
    const cell = { ...sheet.cells['0:0'], id: `${row}:0`, row, col: 0, value: null, formula: formulaText, display: '' }
    sheet.cells[`${row}:0`] = cell
  }
  const calculated = formula.recalculateWorkbookFormulas(recalculation, undefined, { now: clock }).workbook.sheets[0]
  assert.ok(Math.abs(calculated.cells['2:0'].value - expectedNow) < 1e-12)
  assert.equal(calculated.cells['3:0'].value, daySerial)
  assert.ok(Math.abs(calculated.cells['4:0'].value - (expectedNow - daySerial)) < 1e-12)

  for (const name of ['DATEDIF', 'NOW', 'TODAY']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
})

test('text additions match Unicode, exact comparison, cleanup, parsing and delimiter boundaries', () => {
  assert.equal(formula.evaluateFormula('=CHAR(65)', model, 'sheet-main'), 'A')
  assert.equal(formula.evaluateFormula('=UNICHAR(128512)', model, 'sheet-main'), '😀')
  assert.equal(formula.evaluateFormula('=UNICODE("😀")', model, 'sheet-main'), 128512)
  assert.equal(formula.evaluateFormula('=CODE("ab")', model, 'sheet-main'), 97)
  assert.deepEqual(formula.evaluateFormula('=UNICHAR(55296)', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.deepEqual(formula.evaluateFormula('=CHAR(1114112)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.equal(formula.evaluateFormula('=EXACT("Text","text")', model, 'sheet-main'), false)
  assert.equal(formula.evaluateFormula('=CLEAN("A"&CHAR(9)&"B")', model, 'sheet-main'), 'AB')
  assert.equal(formula.evaluateFormula('=PROPER("hello WORLD")', model, 'sheet-main'), 'Hello World')
  assert.equal(formula.evaluateFormula('=NUMBERVALUE("1.234,5",",",".")', model, 'sheet-main'), 1234.5)
  assert.equal(formula.evaluateFormula('=NUMBERVALUE("25%")', model, 'sheet-main'), 0.25)
  assert.deepEqual(formula.evaluateFormula('=NUMBERVALUE("12,34,5")', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.equal(formula.evaluateFormula('=TEXTAFTER("a-b-c","-",-1)', model, 'sheet-main'), 'c')
  assert.equal(formula.evaluateFormula('=TEXTBEFORE("a-b-c","-",-1)', model, 'sheet-main'), 'a-b')
  assert.equal(formula.evaluateFormula('=TEXTAFTER("Alpha-Beta","beta",1,1)', model, 'sheet-main'), '')
  assert.equal(formula.evaluateFormula('=TEXTBEFORE("a-b","-",3,0,1)', model, 'sheet-main'), 'a-b')
  assert.equal(formula.evaluateFormula('=TEXTAFTER("a-b","-",3,0,0,"missing")', model, 'sheet-main'), 'missing')
  assert.equal(formula.evaluateFormula('=TEXTAFTER("abc","",1)', model, 'sheet-main'), 'abc')

  for (const name of ['CHAR', 'CODE', 'CLEAN', 'EXACT', 'NUMBERVALUE', 'PROPER', 'TEXTAFTER', 'TEXTBEFORE', 'UNICHAR', 'UNICODE']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
  assert.ok(!functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === 'TEXTSPLIT'))
})

test('TRUE/FALSE literal and function forms plus SWITCH preserve upstream branch semantics', () => {
  assert.equal(formula.evaluateFormula('=TRUE', model, 'sheet-main'), true)
  assert.equal(formula.evaluateFormula('=FALSE', model, 'sheet-main'), false)
  assert.equal(formula.evaluateFormula('=TRUE()', model, 'sheet-main'), true)
  assert.equal(formula.evaluateFormula('=FALSE()', model, 'sheet-main'), false)
  assert.deepEqual(formula.evaluateFormula('=TRUE(1)', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(formula.evaluateFormula('=SWITCH("B","a","first","b","second","fallback")', model, 'sheet-main'), 'second')
  assert.equal(formula.evaluateFormula('=SWITCH("B","a","first","fallback")', model, 'sheet-main'), 'fallback')
  assert.deepEqual(formula.evaluateFormula('=SWITCH("B","a","first")', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(formula.evaluateFormula('=SWITCH(1,1,"selected",1/0)', model, 'sheet-main'), 'selected')
  assert.deepEqual(formula.evaluateFormula('=SWITCH(1/0,1,"selected")', model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.deepEqual(formula.evaluateFormula('=SWITCH(2,1,"first",1/0,"fallback")', model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.deepEqual(formula.evaluateFormula('=SWITCH()', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.deepEqual(formula.evaluateFormula('=SWITCH(1)', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.deepEqual(formula.evaluateFormula('=SWITCH(#REF!)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.equal(formula.evaluateFormula('=SWITCH(1,"short-default")', model, 'sheet-main'), 'short-default')
  assert.deepEqual(formula.evaluateFormula('=SWITCH(#REF!,"short-default")', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.equal(formula.evaluateFormula(`=SWITCH(1,${Array(128).fill('1,"selected"').join(',')})`, model, 'sheet-main'), 'selected')
  assert.equal(functionCatalog.buildSheetFunctionFormula('TRUE', ''), '=TRUE()')
  assert.equal(functionCatalog.buildSheetFunctionFormula('FALSE', ''), '=FALSE()')

  for (const name of ['TRUE', 'FALSE', 'SWITCH']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
})

test('zero-argument logical and date functions report fixed arity errors before evaluating extra arguments', () => {
  for (const name of ['TRUE', 'FALSE', 'NOW', 'TODAY']) {
    for (const argumentsText of ['1', '1/0', '#REF!', 'A1:B1', '1,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, `${name}(${argumentsText})`)
    }
    assert.equal(formula.evaluateFormula(`=IFNA(${name}(#REF!),"arity")`, model, 'sheet-main'), 'arity', name)
    assert.equal(functionCatalog.buildSheetFunctionFormula(name, ''), `=${name}()`)
  }
})

test('PERCENTOF computes scalar and sparse-reference ratios while preserving subtotal errors and zero denominators', () => {
  for (const [argumentsText, expected] of [
    ['A1,A1:B1', 0.4], ['A1:B1,A1:B1', 1], ['A9,A1:B1', 0],
    ['TRUE,4', 0.25], ['"2","5"', 0.4], ['-2,5', -0.4],
  ]) assert.equal(formula.evaluateFormula(`=PERCENTOF(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  for (const [argumentsText, code] of [
    ['1,0', '#DIV/0!'], ['0,0', '#DIV/0!'], ['NA(),#REF!', '#N/A'],
    ['#REF!,NA()', '#REF!'], ['A1,NA()', '#N/A'], ['"text",#REF!', '#VALUE!'],
    ['', '#N/A'], ['#REF!', '#N/A'], ['#REF!,1,2', '#N/A'],
  ]) assert.deepEqual(formula.evaluateFormula(`=PERCENTOF(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'PERCENTOF')
  assert.equal(entry.category, '逻辑')
  assert.equal(functionCatalog.buildSheetFunctionFormula(entry.name, 'A1,A1:B1'), '=PERCENTOF(A1,A1:B1)')
})

test('logical functions report fixed arity errors before input errors and preserve the 255-argument boundary', () => {
  for (const [name, argumentsText] of [
    ['AND', ''], ['OR', ''], ['XOR', ''], ['NOT', ''], ['NOT', '#REF!,1'],
    ['IF', ''], ['IF', '#REF!'], ['IF', '#REF!,1,2,3'],
    ['IFERROR', ''], ['IFERROR', '#REF!'], ['IFERROR', '#REF!,1,2'],
    ['IFNA', ''], ['IFNA', '#REF!'], ['IFNA', '#REF!,1,2'],
    ['IFS', ''], ['IFS', '#REF!'], ['IFS', '#REF!,1,TRUE'],
  ]) {
    assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, name)
    assert.equal(formula.evaluateFormula(`=IFNA(${name}(${argumentsText}),"arity")`, model, 'sheet-main'), 'arity', name)
  }
  for (const name of ['AND', 'OR', 'XOR']) {
    assert.equal(formula.evaluateFormula(`=${name}(${Array(255).fill('TRUE').join(',')})`, model, 'sheet-main'), true, name)
    assert.deepEqual(formula.evaluateFormula(`=${name}(#REF!,${Array(255).fill('TRUE').join(',')})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, name)
  }
  assert.equal(formula.evaluateFormula(`=IFS(TRUE,7,${Array(126).fill('FALSE,0').join(',')})`, model, 'sheet-main'), 7)
  assert.deepEqual(formula.evaluateFormula(`=IFS(#REF!,7,${Array(127).fill('FALSE,0').join(',')})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(formula.evaluateFormula('=IF(TRUE,7,1/0)', model, 'sheet-main'), 7)
  assert.deepEqual(formula.evaluateFormula('=NOT(#REF!)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
})

test('NA picker builds no-argument formulas and participates in error-handling expressions', () => {
  assert.equal(functionCatalog.buildSheetFunctionFormula('NA', ''), '=NA()')
  assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'NA').category, '信息')
  assert.deepEqual(formula.evaluateFormula('=NA()', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  for (const argument of ['1', '1,2', '1/0', '#REF!', 'A1:B1', '"text"']) {
    assert.deepEqual(formula.evaluateFormula(`=NA(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  assert.equal(formula.evaluateFormula('=IFNA(NA(),"missing")', model, 'sheet-main'), 'missing')
  assert.equal(formula.evaluateFormula('=IFNA(NA(1/0),"missing")', model, 'sheet-main'), 'missing')
  assert.equal(formula.evaluateFormula('=IFERROR(NA(),0)', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=IFERROR(NA(1,2),0)', model, 'sheet-main'), 0)
  assert.deepEqual(formula.evaluateFormula('=NA()+1', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
})

test('information predicates consume error values and preserve scalar types without coercion', () => {
  for (const [expression, expected] of [
    ['ISERROR(1/0)', true], ['ISERR(1/0)', true], ['ISNA(1/0)', false],
    ['ISERROR(NA())', true], ['ISERR(NA())', false], ['ISNA(NA())', true],
    ['ISBLANK(1/0)', false], ['ISLOGICAL(1/0)', false], ['ISNONTEXT(1/0)', true],
    ['ISNUMBER(1/0)', false], ['ISTEXT(1/0)', false], ['ISERROR("#DIV/0!")', false],
    ['ISNA("#N/A")', true], ['ISTEXT("#N/A")', true], ['ISNUMBER("12")', false],
    ['ISNUMBER(12)', true], ['ISLOGICAL(TRUE)', true], ['ISLOGICAL(1)', false],
    ['ISLOGICAL("TRUE")', false], ['ISTEXT("")', true], ['ISNONTEXT("")', false],
    ['ISNONTEXT(TRUE)', true], ['ISBLANK("")', false], ['ISNUMBER(A1:A1)', true],
    ['IF(ISERROR(1/0),"handled","missed")', 'handled'],
  ]) {
    assert.equal(formula.evaluateFormula(`=${expression}`, model, 'sheet-main'), expected, expression)
  }
  for (const name of ['ISBLANK', 'ISERR', 'ISERROR', 'ISLOGICAL', 'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISTEXT']) {
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.filter(entry => entry.name === name && entry.category === '信息').length, 1)
    assert.equal(functionCatalog.buildSheetFunctionFormula(name, 'A1'), `=${name}(A1)`)
    assert.deepEqual(formula.evaluateFormula(`=${name}()`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    assert.deepEqual(formula.evaluateFormula(`=${name}(1/0,1)`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    assert.equal(formula.evaluateFormula(`=IFNA(${name}(#REF!,1),"arity")`, model, 'sheet-main'), 'arity')
    assert.deepEqual(formula.evaluateFormula(`=${name}(A1:B1)`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  }
})

test('information predicates read numeric text formats and forced error text through scalar references and INDEX', () => {
  const data = structuredClone(model)
  const sheet = data.sheets[0]
  for (const [column, value, patch] of [
    [0, 12, { dataType: 'number', style: { ...sheet.cells['0:0'].style, numberFormat: '@' } }],
    [1, 12, { dataType: 'formula', cachedDataType: 'number', formula: '=A1', style: { ...sheet.cells['0:0'].style, numberFormat: '@@@' } }],
    [2, '#N/A', { dataType: 'forcedString' }],
    [3, '#N/A', { dataType: 'error' }],
  ]) {
    sheet.cells[`4:${column}`] = { ...sheet.cells['0:0'], id: `4:${column}`, row: 4, col: column, value, formula: null, ...patch }
  }
  for (const reference of ['A5', 'A5:A5', 'B5', 'INDEX(A5:B5,1,1)', 'Data!A5']) {
    assert.equal(formula.evaluateFormula(`=ISNUMBER(${reference})`, data, 'sheet-main'), false, reference)
    assert.equal(formula.evaluateFormula(`=ISTEXT(${reference})`, data, 'sheet-main'), true, reference)
    assert.equal(formula.evaluateFormula(`=ISNONTEXT(${reference})`, data, 'sheet-main'), false, reference)
    assert.equal(formula.evaluateFormula(`=ISBLANK(${reference})`, data, 'sheet-main'), false, reference)
  }
  assert.equal(formula.evaluateFormula('=ISNA(C5)', data, 'sheet-main'), true)
  assert.equal(formula.evaluateFormula('=ISERROR(C5)', data, 'sheet-main'), false)
  assert.equal(formula.evaluateFormula('=ISERR(C5)', data, 'sheet-main'), false)
  assert.equal(formula.evaluateFormula('=ISTEXT(C5)', data, 'sheet-main'), true)
  assert.equal(formula.evaluateFormula('=ISNA(D5)', data, 'sheet-main'), true)
  assert.equal(formula.evaluateFormula('=ISERROR(D5)', data, 'sheet-main'), true)
  assert.equal(formula.evaluateFormula('=ISERR(D5)', data, 'sheet-main'), false)
  assert.equal(formula.evaluateFormula('=ISNUMBER(A1)', data, 'sheet-main'), true)
  assert.equal(data.sheets[0].cells['4:0'].value, 12)
})

test('ERROR.TYPE maps fixed error values and literal error text without changing unsupported error identities', () => {
  for (const [code, expected] of [
    ['#NULL!', 1], ['#DIV/0!', 2], ['#VALUE!', 3], ['#REF!', 4], ['#NAME?', 5],
    ['#NUM!', 6], ['#N/A', 7], ['#GETTING_DATA', 8], ['#CALC!', 14],
  ]) {
    assert.equal(formula.evaluateFormula(`=ERROR.TYPE("${code}")`, model, 'sheet-main'), expected)
    if (expected < 8) assert.equal(formula.evaluateFormula(`=ERROR.TYPE(${code})`, model, 'sheet-main'), expected)
  }
  for (const argument of ['7', 'TRUE', 'FALSE', 'A9', '""', '"#SPILL!"', '"#CYCLE!"', '"#ERROR!"', '#CIRCULAR!']) {
    assert.deepEqual(formula.evaluateFormula(`=ERROR.TYPE(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, argument)
  }
  assert.equal(formula.evaluateFormula('=ERROR.TYPE(INDEX(A1:C1,1,3)/0)', model, 'sheet-main'), 2)
  for (const [name, example, expected] of [['ERROR.TYPE', '"#NULL!"', 1], ['N', '7', 7], ['TYPE', 'A2', 1], ['ISREF', 'A1', true]]) {
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '信息')
    assert.equal(entry.example, example)
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, example), model, 'sheet-main'), expected)
    for (const argumentsText of ['', '#REF!,2', '1,2,3']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, `${name}(${argumentsText})`)
    }
  }
})

test('N, TYPE and ISREF preserve first-cell conversion, reference shape and typed errors through INDEX', () => {
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  for (const [column, value, dataType] of [[0, 7, 'number'], [1, '12', 'string'], [2, true, 'boolean'], [3, '#DIV/0!', 'error'], [4, '#DIV/0!', 'string']]) {
    sheet.cells[`2:${column}`] = { ...sheet.cells['0:0'], id: `2:${column}`, row: 2, col: column,
      value, dataType, formula: null, display: String(value) }
  }
  sheet.cells['2:5'] = { ...sheet.cells['2:0'], id: '2:5', col: 5, style: { ...sheet.cells['2:0'].style, numberFormat: '@' } }
  for (const [expression, expected] of [
    ['N(A3:E3)', 7], ['N(B3:D3)', 0], ['N(C3:E3)', 1], ['N(A9:B10)', 0],
    ['N("12")', 0], ['N("#DIV/0!")', 0], ['N(FALSE)', 0], ['N(F3)', 0],
    ['N(INDEX(A3:E3,1,1))', 7], ['N(INDEX(A3:E3,1,2))', 0],
    ['TYPE(A3)', 1], ['TYPE(B3)', 2], ['TYPE(C3)', 4], ['TYPE(D3)', 16], ['TYPE(E3)', 2],
    ['TYPE(F3)', 2], ['TYPE(A9)', 1], ['TYPE(A3:E3)', 64], ['TYPE(D3:E3)', 64],
    ['TYPE(INDEX(A3:E3,1,4))', 16], ['TYPE(INDEX(A3:E3,0,1))', 1], ['TYPE(INDEX(A3:E3,1,0))', 64],
    ['TYPE(NA())', 16], ['TYPE("7")', 2], ['TYPE(FALSE)', 4], ['TYPE(Missing!A1)', 16],
    ['ISREF(D3)', true], ['ISREF(A3:E3)', true], ['ISREF(A:A)', true], ['ISREF(1:1)', true],
    ['ISREF(Data!A3)', true], ['ISREF(INDEX(A3:E3,1,4))', true], ['ISREF(INDEX(A3:E3,1,0))', true],
    ['ISREF("A1")', false], ['ISREF(TRUE)', false], ['ISREF(7)', false], ['ISREF(NA())', false],
    ['ISREF(Missing!A1)', false], ['ISREF(INDEX(A3:E3,2,1))', false], ['ISREF(A3+0)', false],
    ['ERROR.TYPE(D3)', 2], ['ERROR.TYPE(E3)', 2],
  ]) assert.equal(formula.evaluateFormula(`=${expression}`, source, 'sheet-main'), expected, expression)
  assert.deepEqual(formula.evaluateFormula('=N(D3:E3)', source, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.deepEqual(formula.evaluateFormula('=N(NA())', source, 'sheet-main'), { kind: 'error', code: '#N/A' })
})

test('ISBETWEEN retains scalar type ordering, inclusivity, blank direction and numeric comparison precision', () => {
  for (const [argumentsText, expected] of [
    ['7.9,1.2,12.45', true], ['7.9,1.2,-12.45', '#NUM!'], ['"7.9",1.2,12.45', false],
    ['"test",1.2,12.45', false], ['A9,1.2,12.45', false], ['TRUE,1.2,12.45', false],
    ['0,"AA",1,1,FALSE', '#NUM!'], ['0,"AA","BB",1,FALSE', false], ['0,1,FALSE', false],
    ['FALSE,1,TRUE', true], ['"10","1","2"', true], ['"10","2","9"', false],
    ['1,1,2', true], ['1,1,2,FALSE', false], ['2,1,2,TRUE,FALSE', false],
    ['1,1,2,2,-1', true], ['1,1,2,A9', false], ['2,1,2,TRUE,A9', false],
    ['A9,0,0', true], ['0,A9,A9', true], ['FALSE,A9,A9', true], ['A9,FALSE,FALSE', true],
    ['A9,"","a"', true], ['"",A9,A9', false], ['A1:A1,1,3', true],
    ['0.1+0.2,0.3,0.3', true], ['0.1+0.2,0.3,0.3,FALSE', false],
    ['1.00000000000005,1,1', true], ['1000000000000.0001,1000000000000,1000000000000', false],
  ]) {
    const value = formula.evaluateFormula(`=ISBETWEEN(${argumentsText})`, model, 'sheet-main')
    if (typeof expected === 'string') assert.deepEqual(value, { kind: 'error', code: expected }, argumentsText)
    if (typeof expected !== 'string') assert.equal(value, expected, argumentsText)
  }
  assert.equal(functionCatalog.buildSheetFunctionFormula('ISBETWEEN', '7.9,1.2,12.45'), '=ISBETWEEN(7.9,1.2,12.45)')
  assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'ISBETWEEN').category, '信息')
})

test('ISBETWEEN retains literal and wildcard comparisons and ordered shape, error, flag and arity failures', () => {
  for (const [argumentsText, expected] of [
    ['"A","a","A"', true], ['"é","é","é"', true], ['"é","e","e"', false],
    ['"straße","STRASSE","strasse"', true], ['"oﬃce","office","office"', true],
    ['"ab","a*","z",FALSE', true], ['"ab","a","a*"', false],
    ['"a\nb","a*b","z"', true], ['"~*","~~*","z"', true], ['"İ","?","z"', true],
  ]) assert.equal(formula.evaluateFormula(`=ISBETWEEN(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  for (const [argumentsText, code] of [
    ['#REF!,A1:B1,5', '#REF!'], ['A1:B1,#REF!,5', '#VALUE!'], ['1,0,2,A1:B1', '#VALUE!'],
    ['1,0,2,TRUE,A1:B1', '#VALUE!'], ['1,2,0,"TRUE",#REF!', '#REF!'], ['1,2,0,"TRUE"', '#VALUE!'],
    ['1,0,2,"0"', '#VALUE!'], ['1,0,2,TRUE,""', '#VALUE!'], ['1,2,0,TRUE,1/0', '#DIV/0!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=ISBETWEEN(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  for (const argumentsText of ['', '#REF!,1', '#REF!,1,2,TRUE,TRUE,TRUE']) {
    assert.deepEqual(formula.evaluateFormula(`=ISBETWEEN(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
})

test('ISEVEN and ISODD truncate raw scalar numbers and retain Boolean, shape, error and arity semantics', () => {
  for (const [argument, even, odd] of [
    ['-1', false, true], ['2.5', true, false], ['-3.9', false, true], ['0', true, false],
    ['"122"', true, false], ['"123"', false, true], ['"0x2"', true, false],
    ['""', true, false], ['" "', true, false], ['A9', true, false], ['A1:A1', true, false],
    ['"1e999"', false, true],
  ]) {
    assert.equal(formula.evaluateFormula(`=ISEVEN(${argument})`, model, 'sheet-main'), even, argument)
    assert.equal(formula.evaluateFormula(`=ISODD(${argument})`, model, 'sheet-main'), odd, argument)
  }
  for (const name of ['ISEVEN', 'ISODD']) {
    for (const argument of ['TRUE', 'FALSE', '"25%"', '"$2"', '"invalid"', 'A1:B1', 'QUOTIENT("1e999","1e999")']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, `${name}(${argument})`)
    }
    assert.deepEqual(formula.evaluateFormula(`=${name}(#REF!)`, model, 'sheet-main'), { kind: 'error', code: '#REF!' })
    for (const argumentsText of ['', '#REF!,1']) assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name).category, '信息')
  }
})

test('ISFORMULA reads stored formulas without resolving their errors or self references', () => {
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  sheet.cells['2:0'] = { ...sheet.cells['0:0'], id: '2:0', row: 2, col: 0, formula: '=1/0' }
  sheet.cells['2:1'] = { ...sheet.cells['0:0'], id: '2:1', row: 2, col: 1, value: '=SUM(A1)', dataType: 'string', formula: null }
  sheet.cells['3:0'] = { ...sheet.cells['0:0'], id: '3:0', row: 3, col: 0, formula: '=ISFORMULA(A4)' }
  for (const [expression, expected] of [
    ['ISFORMULA(A1)', false], ['ISFORMULA(C1)', true], ['ISFORMULA(A3)', true],
    ['ISFORMULA(B3)', false], ['ISFORMULA(C3)', false], ['ISFORMULA(Data!A3)', true],
    ['ISFORMULA(INDEX(A3:B3,1,1))', true], ['ISFORMULA(INDEX(A3:B3,1,2))', false], ['ISFORMULA(A4)', true],
  ]) assert.equal(formula.evaluateFormula(`=${expression}`, source, 'sheet-main'), expected, expression)
  assert.equal(formula.recalculateWorkbookFormulas(source).workbook.sheets[0].cells['3:0'].value, true)
  for (const [argument, code] of [['7', '#N/A'], ['TRUE', '#N/A'], ['"A1"', '#N/A'], ['#REF!', '#REF!'], ['Missing!A1', '#REF!'], ['A3:B3', '#VALUE!']]) {
    assert.deepEqual(formula.evaluateFormula(`=ISFORMULA(${argument})`, source, 'sheet-main'), { kind: 'error', code }, argument)
  }
  for (const argumentsText of ['', '#REF!,1']) assert.deepEqual(formula.evaluateFormula(`=ISFORMULA(${argumentsText})`, source, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(functionCatalog.buildSheetFunctionFormula('ISFORMULA', 'A1'), '=ISFORMULA(A1)')
})

test('SHEET and SHEETS use workbook order, hidden sheets and reference identity with fixed arity', () => {
  const source = structuredClone(model)
  source.sheets.push({ ...source.sheets[0], id: 'sheet-other', name: 'Other', cells: {} })
  source.sheets.push({ ...source.sheets[0], id: 'sheet-hidden', name: '3', visible: false, cells: {} })
  for (const [expression, expected] of [
    ['SHEET()', 1], ['SHEET(A1)', 1], ['SHEET(Data!C1)', 1], ['SHEET(Other!A1:B2)', 2],
    ['SHEET("oThEr")', 2], ['SHEET(3)', 3], ['SHEET(INDEX(Other!A1:B2,1,1))', 2], ['SHEETS()', 3],
  ]) assert.equal(formula.evaluateFormula(`=${expression}`, source, 'sheet-main'), expected, expression)
  assert.equal(formula.evaluateFormula('=SHEET()', source, 'sheet-other'), 2)
  source.sheets.reverse()
  assert.equal(formula.evaluateFormula('=SHEET()', source, 'sheet-main'), 3)
  assert.equal(formula.evaluateFormula('=SHEET(3)', source, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=SHEETS()', source, 'sheet-main'), 3)
  for (const [expression, code] of [
    ['SHEET("unknown")', '#N/A'], ['SHEET(TRUE)', '#N/A'], ['SHEET(#REF!)', '#REF!'],
    ['SHEET(Missing!A1)', '#REF!'], ['SHEET(NA(),1)', '#N/A'], ['SHEETS(#REF!)', '#N/A'],
  ]) assert.deepEqual(formula.evaluateFormula(`=${expression}`, source, 'sheet-main'), { kind: 'error', code }, expression)
  assert.equal(functionCatalog.buildSheetFunctionFormula('SHEET', ''), '=SHEET()')
  assert.equal(functionCatalog.buildSheetFunctionFormula('SHEETS', ''), '=SHEETS()')
  assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'SHEET').example, 'A1')
  assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'SHEETS').example, '')
})

test('ISEMAIL retains the fixed domain list, dot-separated local parts, whitespace and length boundaries', () => {
  for (const [text, expected] of [
    ['developer@univer.ai', true], ['A.B+c-d@sub.example.co.uk', true], ['name@sample.onion', true],
    ['a_b@under_score.example.com', true], ['developer@univer.ai ', false], [' developer@univer.ai', false],
    ['a..b@univer.ai', false], ['a@univer', false], ['a@sample.app', false], ['a@sample.invalid', false],
    ['用户@univer.ai', false], ['#DIV/0!', false], ['a'.repeat(242) + '@example.com', true], ['a'.repeat(243) + '@example.com', false],
  ]) assert.equal(formula.evaluateFormula(`=ISEMAIL("${text}")`, model, 'sheet-main'), expected, text)
  for (const argument of ['7', 'TRUE', 'FALSE', 'A9', '""']) assert.equal(formula.evaluateFormula(`=ISEMAIL(${argument})`, model, 'sheet-main'), false)
  assert.deepEqual(formula.evaluateFormula('=ISEMAIL(#REF!)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.deepEqual(formula.evaluateFormula('=ISEMAIL(A1:B1)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  for (const argumentsText of ['', '#REF!,1']) assert.deepEqual(formula.evaluateFormula(`=ISEMAIL(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'ISEMAIL').example, '"developer@univer.ai"')
})

test('ISURL retains fixed schemes, schemeless domains, encoded paths, mailto and trimmed length boundaries', () => {
  const prefix = 'https://univer.ai/'
  for (const [text, expected] of [
    ['univer.ai', true], [' \nuniver.ai\t ', true], ['HTTPS://host.invalid:443/a%20b?q=中文#页', true],
    ['http:///local', true], ['SSH://local_1:22', true], ['sftp://host/a', true], ['ftps://host/a', true],
    ['nfs://host/a', true], ['mailto:developer@univer.ai', true], ['news:comp.lang', true], ['aim:goim?screenname=test', true],
    ['sample.invalid', false], ['sample.app', false], ['//univer.ai', false], ['gopher://univer.ai', false],
    ['https://user@univer.ai', false], ['https://[::1]', false], ['https://univer.ai/%2G', false],
    ['https://univer.ai/中文', false], ['https://univer.ai/?a b', false], ['mailto:a@sample.app', false],
    [prefix + 'a'.repeat(1000 - prefix.length), true], [prefix + 'a'.repeat(1001 - prefix.length), false],
  ]) assert.equal(formula.evaluateFormula(`=ISURL("${text}")`, model, 'sheet-main'), expected, text)
  for (const argument of ['7', 'TRUE', 'FALSE', 'A9', '""']) assert.equal(formula.evaluateFormula(`=ISURL(${argument})`, model, 'sheet-main'), false)
  assert.deepEqual(formula.evaluateFormula('=ISURL(#REF!)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.deepEqual(formula.evaluateFormula('=ISURL(A1:B1)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  for (const argumentsText of ['', '#REF!,1']) assert.deepEqual(formula.evaluateFormula(`=ISURL(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'ISURL').example, '"univer.ai"')
})

test('ISBLANK retains cell-reference identity for empty cells, formulas, rich text and INDEX', () => {
  const data = structuredClone(model)
  const sheet = data.sheets[0]
  for (const [column, value, patch] of [
    [1, null, {}], [2, '', { dataType: 'string' }], [3, '', { formula: '=""' }],
    [4, null, { formula: '=IF(TRUE,A3)' }], [5, ' ', {}], [6, null, { richText: [{ text: '' }] }],
    [7, '#N/A', { dataType: 'error' }], [8, '#N/A', { dataType: 'string' }],
  ]) {
    sheet.cells[`2:${column}`] = { ...sheet.cells['0:0'], id: `2:${column}`, row: 2, col: column,
      value, formula: null, display: value === null ? '' : String(value), ...patch }
  }
  sheet.cells['2:1'].style = { ...sheet.cells['2:1'].style, font: { ...sheet.cells['2:1'].style.font, bold: true } }
  for (const [expression, expected] of [
    ['ISBLANK(A3)', true], ['ISBLANK(B3)', true], ['ISBLANK(C3)', true],
    ['ISBLANK(D3)', false], ['ISBLANK(E3)', false], ['ISBLANK(F3)', false],
    ['ISBLANK(G3)', false], ['ISBLANK(INDEX(A3:C3,1,3))', true],
    ['ISBLANK(INDEX(D3:E3,1,2))', false], ['ISBLANK(Data!C3)', true],
    ['ISERROR(H3)', true], ['ISNA(H3:H3)', true], ['ISTEXT(H3)', false],
    ['ISERROR(I3)', false], ['ISNA(I3)', true], ['ISTEXT(I3)', true],
  ]) {
    assert.equal(formula.evaluateFormula(`=${expression}`, data, 'sheet-main'), expected, expression)
  }
  const target = { ...sheet.cells['0:2'], id: '3:0', row: 3, col: 0, formula: '=ISBLANK(E3)' }
  sheet.cells['3:0'] = target
  const result = formula.recalculateWorkbookFormulas(data)
  assert.equal(result.workbook.sheets[0].cells['3:0'].value, false)
  assert.equal(result.workbook.sheets[0].cells['3:0'].cachedDataType, 'boolean')
  assert.equal(result.workbook.sheets[0].cells['3:0'].formula, '=ISBLANK(E3)')
})

test('RIGHT returns empty text for zero length and preserves existing suffix and parameter behavior', () => {
  assert.equal(formula.evaluateFormula('=RIGHT("abc",0)', model, 'sheet-main'), '')
  assert.equal(formula.evaluateFormula('=RIGHT("abc")', model, 'sheet-main'), 'c')
  assert.equal(formula.evaluateFormula('=RIGHT("abc",2)', model, 'sheet-main'), 'bc')
  assert.equal(formula.evaluateFormula('=RIGHT("abc",8)', model, 'sheet-main'), 'abc')
  assert.deepEqual(formula.evaluateFormula('=RIGHT("abc",-1)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
})

test('LEFT, RIGHT and MID floor numeric positions and lengths before checking their bounds', () => {
  for (const [expression, expected] of [
    ['LEFT("abcdef",2.9)', 'ab'], ['LEFT("abcdef",0.9)', ''],
    ['RIGHT("abcdef",2.9)', 'ef'], ['RIGHT("abcdef",0.9)', ''],
    ['MID("abcdef",2.9,3.9)', 'bcd'], ['MID("abcdef",1.9,0.9)', ''],
    ['MID("abcdef",9.9,2.9)', ''],
  ]) assert.equal(formula.evaluateFormula(`=${expression}`, model, 'sheet-main'), expected, expression)
  for (const expression of ['LEFT("abc",-0.1)', 'RIGHT("abc",-0.1)', 'MID("abc",0.9,1)', 'MID("abc",1,-0.1)']) {
    assert.deepEqual(formula.evaluateFormula(`=${expression}`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, expression)
  }
  assert.deepEqual(formula.evaluateFormula('=MID("abc",1,NA())', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
})

test('WEEKDAY uses every upstream week start, floors return types and retains both serial-date epochs', () => {
  for (const [returnType, expectedSunday] of [[1, 1], [2, 7], [3, 6], [11, 7], [12, 6], [13, 5], [14, 4], [15, 3], [16, 2], [17, 1]]) {
    assert.equal(formula.evaluateFormula(`=WEEKDAY(DATE(2024,1,7),${returnType})`, model, 'sheet-main'), expectedSunday, String(returnType))
  }
  assert.equal(formula.evaluateFormula('=WEEKDAY(DATE(2024,1,8),12)', model, 'sheet-main'), 7)
  assert.equal(formula.evaluateFormula('=WEEKDAY(DATE(2024,1,9),12)', model, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=WEEKDAY(DATE(2024,1,7),2.9)', model, 'sheet-main'), 7)
  assert.equal(formula.evaluateFormula('=WEEKDAY(DATE(2024,1,7),11.9)', model, 'sheet-main'), 7)
  assert.equal(formula.evaluateFormula('=WEEKDAY(DATE(2024,1,7))', model, 'sheet-main'), 1)
  for (const invalid of [0, 4, 10, 18, -1]) {
    assert.deepEqual(formula.evaluateFormula(`=WEEKDAY(45298,${invalid})`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  }
  assert.deepEqual(formula.evaluateFormula('=WEEKDAY(45298,"invalid")', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=WEEKDAY(45298,NA())', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  for (const [serial, expected1900, expected1904] of [[0, 7, 6], [59, 3, 2], [60, 4, 3], [61, 5, 4]]) {
    assert.equal(formula.evaluateFormula(`=WEEKDAY(${serial})`, model, 'sheet-main'), expected1900)
    assert.equal(formula.evaluateFormula(`=WEEKDAY(${serial})`, { ...model, date1904: true }, 'sheet-main'), expected1904)
  }
})

test('text slices use formatted numeric arguments and read single-cell ranges without losing errors', () => {
  for (const [expression, expected] of [
    ['LEFT("abcdef","2.9")', 'ab'], ['RIGHT("abcdef","2.9")', 'ef'],
    ['MID("abcdef","2.9","3.9")', 'bcd'], ['MID("abcdef"," 2 ","3e0")', 'bcd'],
    ['LEFT("abcdef","200%")', 'ab'], ['RIGHT("abcdef","$2")', 'ef'],
    ['LEFT("abcdef","0x2")', 'ab'], ['LEFT("abcdef",TRUE)', 'a'],
    ['LEFT("abcdef","1e999")', 'abcdef'], ['MID("abcdef",1,"1e999")', 'abcdef'],
  ]) assert.equal(formula.evaluateFormula(`=${expression}`, model, 'sheet-main'), expected, expression)
  for (const expression of ['LEFT("abc","")', 'RIGHT("abc"," ")', 'MID("abc",1,"invalid")', 'LEFT("abc","1,00")', 'RIGHT("abc","(0.1)")']) {
    assert.deepEqual(formula.evaluateFormula(`=${expression}`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, expression)
  }
  const source = structuredClone(model), sheet = source.sheets[0]
  for (const [column, value] of [[0, 'abcdef'], [1, '2.9'], [2, '3.9'], [3, '2024-01-07'], [4, '12'], [5, '#N/A']]) {
    sheet.cells[`2:${column}`] = { ...sheet.cells['0:0'], id: `2:${column}`, row: 2, col: column,
      value, display: value, formula: null, dataType: column === 5 ? 'error' : 'text' }
  }
  for (const [expression, expected] of [
    ['LEFT(A3:A3,B3:B3)', 'ab'], ['RIGHT(A3,B3:B3)', 'ef'], ['MID(A3:A3,B3,C3:C3)', 'bcd'],
    ['WEEKDAY(D3:D3,E3:E3)', 6], ['LEFT(A3,A4:A4)', ''],
  ]) assert.equal(formula.evaluateFormula(`=${expression}`, source, 'sheet-main'), expected, expression)
  for (const expression of ['LEFT(A3:A3,F3:F3)', 'WEEKDAY(F3:F3)']) {
    assert.deepEqual(formula.evaluateFormula(`=${expression}`, source, 'sheet-main'), { kind: 'error', code: '#N/A' }, expression)
  }
  assert.deepEqual(formula.evaluateFormula('=LEFT(A3:B3,1)', source, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
})

test('text slices retain the fixed source UTF-16 boundaries for surrogate pairs and combining text', () => {
  assert.equal(formula.evaluateFormula('=LEFT("A😀B",2)', model, 'sheet-main'), 'A\uD83D')
  assert.equal(formula.evaluateFormula('=RIGHT("A😀B",2)', model, 'sheet-main'), '\uDE00B')
  assert.equal(formula.evaluateFormula('=MID("A😀B",2,1)', model, 'sheet-main'), '\uD83D')
  assert.equal(formula.evaluateFormula('=MID("A😀B",2,2)', model, 'sheet-main'), '😀')
  assert.equal(formula.evaluateFormula('=LEFT("éx",1)', model, 'sheet-main'), 'e')
})

test('text slices, VALUE and WEEKDAY build from the insert-function catalog and evaluate their supplied arguments', () => {
  for (const [name, argumentsText, expected] of [
    ['LEFT', '"Univer",3', 'Uni'], ['MID', '"Univer",2,3', 'niv'],
    ['RIGHT', '"Univer",3', 'ver'], ['WEEKDAY', '"2008-2-14",2', 4],
    ['VALUE', '"123"', 123],
  ]) {
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.filter(entry => entry.name === name).length, 1)
    const expression = functionCatalog.buildSheetFunctionFormula(name, argumentsText)
    assert.equal(formula.evaluateFormula(expression, model, 'sheet-main'), expected, name)
    assert.throws(() => functionCatalog.buildSheetFunctionFormula(name, ''), /请输入函数参数/)
  }
})

test('VALUE converts formatted numeric text, dates and times while retaining scalar reference types and errors', () => {
  for (const [argument, expected] of [
    ['"1,234.5"', 1234.5], ['"$2"', 2], ['"25%"', 0.25], ['"(2.5)"', -2.5],
    ['"2024-01-07"', 45298], ['"12:30"', 12.5 / 24], ['"1e999"', Infinity], ['2.5', 2.5],
  ]) assert.equal(formula.evaluateFormula(`=VALUE(${argument})`, model, 'sheet-main'), expected, argument)
  for (const argument of ['""', '" "', '"invalid"', '"1,00"', 'TRUE', 'FALSE']) {
    assert.deepEqual(formula.evaluateFormula(`=VALUE(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, argument)
  }
  assert.equal(formula.evaluateFormula('=VALUE(A1:A1)', model, 'sheet-main'), 2)
  assert.equal(formula.evaluateFormula('=VALUE(INDEX(A1:A1,1))', model, 'sheet-main'), 2)
  assert.equal(formula.evaluateFormula('=VALUE(A9:A9)', model, 'sheet-main'), 0)
  assert.deepEqual(formula.evaluateFormula('=VALUE(A1:B1)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=VALUE(NA())', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(formula.evaluateFormula('=VALUE("2024-01-07")', { ...model, date1904: true }, 'sheet-main'), 43836)
})

test('TRIM collapses ordinary and ideographic spaces while retaining tabs, newlines and non-breaking spaces', () => {
  for (const [text, expected] of [
    ['  a   b  ', 'a b'], ['　 a　　b 　', 'a　b'], ['a 　b', 'a b'], ['a　 b', 'a　b'],
    ['\ta\tb\n', '\ta\tb\n'], ['\u00A0a  b\u00A0', '\u00A0a b\u00A0'], ['a \t b', 'a \t b'], [' 　 ', ''],
  ]) assert.equal(formula.evaluateFormula(`=TRIM("${text}")`, model, 'sheet-main'), expected, JSON.stringify(text))
})

test('LEN, UPPER, LOWER and TRIM read scalar ranges and INDEX while preserving reference errors', () => {
  const source = structuredClone(model), sheet = source.sheets[0]
  sheet.cells['2:0'] = { ...sheet.cells['0:0'], id: '2:0', row: 2, col: 0, value: ' A😀  b ', display: ' A😀  b ', formula: null, dataType: 'text' }
  sheet.cells['2:1'] = { ...sheet.cells['0:1'], id: '2:1', row: 2, col: 1, value: '#N/A', display: '#N/A', formula: null, dataType: 'error' }
  for (const [name, expected] of [['LEN', 8], ['UPPER', ' A😀  B '], ['LOWER', ' a😀  b '], ['TRIM', 'A😀 b']]) {
    assert.equal(formula.evaluateFormula(`=${name}(A3:A3)`, source, 'sheet-main'), expected, name)
    assert.equal(formula.evaluateFormula(`=${name}(INDEX(A3:A3,1))`, source, 'sheet-main'), expected, name)
    assert.deepEqual(formula.evaluateFormula(`=${name}(B3:B3)`, source, 'sheet-main'), { kind: 'error', code: '#N/A' }, name)
    assert.deepEqual(formula.evaluateFormula(`=${name}(A3:B3)`, source, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, name)
  }
  assert.equal(formula.evaluateFormula('=LEN(A9:A9)', source, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=TRIM(A9:A9)', source, 'sheet-main'), '')
})

test('WEEKDAY accepts numeric, date and time text with the fixed serial limits for each epoch', () => {
  for (const expression of ['WEEKDAY("45298","12")', 'WEEKDAY("2024-01-07",12)', 'WEEKDAY("2024-01-07T12:30:00","$12")']) {
    assert.equal(formula.evaluateFormula(`=${expression}`, model, 'sheet-main'), 6, expression)
  }
  assert.equal(formula.evaluateFormula('=WEEKDAY("2024-01-07","2.9")', model, 'sheet-main'), 7)
  assert.equal(formula.evaluateFormula('=WEEKDAY("12:30")', model, 'sheet-main'), 7)
  for (const value of ['""', '" "', '"invalid"', '"$45298"', '"TRUE"']) {
    assert.deepEqual(formula.evaluateFormula(`=WEEKDAY(${value})`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, value)
  }
  for (const [date1904, maxSerial] of [[false, 2958465], [true, 2957003]]) {
    const source = { ...model, date1904 }
    assert.equal(formula.evaluateFormula('=WEEKDAY("2024-01-07",12)', source, 'sheet-main'), 6)
    assert.equal(formula.evaluateFormula(`=WEEKDAY(${maxSerial})`, source, 'sheet-main'), 6)
    for (const serial of [-1, maxSerial + 0.1, maxSerial + 1]) {
      assert.deepEqual(formula.evaluateFormula(`=WEEKDAY(${serial})`, source, 'sheet-main'), { kind: 'error', code: '#NUM!' }, String(serial))
    }
  }
  assert.deepEqual(formula.evaluateFormula('=WEEKDAY("1e999")', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
})

test('date conversion functions honor workbook epochs, accepted time text and scalar passthrough', () => {
  assert.equal(formula.evaluateFormula('=EPOCHTODATE(0)', model, 'sheet-main'), 25569)
  assert.ok(Math.abs(formula.evaluateFormula('=EPOCHTODATE(1655906710)', model, 'sheet-main') - 44734.58692129629) < 1e-10)
  assert.ok(Math.abs(formula.evaluateFormula('=EPOCHTODATE(1655906710,2)', model, 'sheet-main') - 25588.165586921295) < 1e-10)
  assert.ok(Math.abs(formula.evaluateFormula('=EPOCHTODATE(1655906710,3)', model, 'sheet-main') - 25569.019165578702) < 1e-10)
  assert.deepEqual(formula.evaluateFormula('=EPOCHTODATE(-1)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(formula.evaluateFormula('=EPOCHTODATE(0,4)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(formula.evaluateFormula('=EPOCHTODATE("0")', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.ok(Math.abs(formula.evaluateFormula('=EPOCHTODATE(A1:A1)', model, 'sheet-main') - (25569 + 2000 / 86400000)) < 1e-12)
  assert.deepEqual(formula.evaluateFormula('=EPOCHTODATE(A1:B1)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })

  const date1904 = structuredClone(model)
  date1904.date1904 = true
  assert.equal(formula.evaluateFormula('=EPOCHTODATE(0)', date1904, 'sheet-main'), 24107)

  assert.equal(formula.evaluateFormula('=TIMEVALUE("10:11:12")', model, 'sheet-main'), 36672 / 86400)
  assert.equal(formula.evaluateFormula('=TIMEVALUE("2020-01-02")', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=TIMEVALUE("2020-01-02 13:14:15")', model, 'sheet-main'), 47655 / 86400)
  assert.equal(formula.evaluateFormula('=TIMEVALUE("3:20 PM")', model, 'sheet-main'), 55200 / 86400)
  assert.deepEqual(formula.evaluateFormula('=TIMEVALUE("25:00")', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=TIMEVALUE(0.5)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=TIMEVALUE(A1:A1)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })

  assert.equal(formula.evaluateFormula('=TO_DATE(40826.4375)', model, 'sheet-main'), 40826.4375)
  assert.equal(formula.evaluateFormula('=TO_DATE(A1:A1)', model, 'sheet-main'), 2)
  assert.deepEqual(formula.evaluateFormula('=TO_DATE(A1:B1)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.equal(formula.evaluateFormula('=TO_DATE("40826.4375")', model, 'sheet-main'), '40826.4375')
  assert.equal(formula.evaluateFormula('=TO_DATE(TRUE)', model, 'sheet-main'), true)
  assert.deepEqual(formula.evaluateFormula('=TO_DATE(1/0)', model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })

  for (const name of ['EPOCHTODATE', 'TIMEVALUE', 'TO_DATE']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
})

test('PI keeps the fixed zero-argument result and argument-count error precedence', () => {
  assert.equal(functionCatalog.buildSheetFunctionFormula('PI', ''), '=PI()')
  assert.equal(formula.evaluateFormula('=PI()', model, 'sheet-main'), Math.PI)
  assert.equal(formula.evaluateFormula('=DEGREES(PI())', model, 'sheet-main'), 180)
  assert.equal(formula.evaluateFormula('=RADIANS(180)', model, 'sheet-main'), Math.PI)
  for (const argument of ['1', '1/0', '#REF!', 'NA()', 'A1:B1', '1,2']) {
    assert.deepEqual(formula.evaluateFormula(`=PI(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'PI')
  assert.equal(entry.category, '数学')
  assert.equal(entry.example, '')
  assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.filter(entry => entry.name === 'PI').length, 1)
})

test('angle conversions preserve scalar coercion, source errors, non-finite input and fixed arity', () => {
  const source = structuredClone(model)
  for (const [row, value, dataType] of [[1, '#N/A', 'error'], [2, Infinity, 'number'], [3, -Infinity, 'number'], [4, NaN, 'number']]) {
    source.sheets[0].cells[`${row}:0`] = { ...source.sheets[0].cells['0:0'], id: `${row}:0`, row, col: 0, value, display: String(value), formula: null, dataType }
  }
  for (const [name, atOne] of [['DEGREES', 57.29577951308232], ['RADIANS', 0.017453292519943295]]) {
    for (const argument of ['1', 'TRUE', '"1"', '"$1"', '"100%"']) {
      assert.equal(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), atOne, `${name}(${argument})`)
    }
    for (const argument of ['A1', 'A1:A1', 'INDEX(A1:A1,1)']) {
      assert.equal(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), atOne * 2, `${name}(${argument})`)
    }
    for (const argument of ['0', 'FALSE', 'A9:A9']) {
      assert.equal(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), 0, `${name}(${argument})`)
    }
    for (const argumentsText of ['', '1,2', '1/0,2', '#REF!,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, `${name}(${argumentsText})`)
    }
    for (const argument of ['""', '" "', '"invalid"', '"1e999"', 'A3', 'A4:A4', 'A5:A5']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, source, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, `${name}(${argument})`)
    }
    for (const [argument, code] of [['NA()', '#N/A'], ['A2:A2', '#N/A'], ['1/0', '#DIV/0!'], ['#REF!', '#REF!']]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, source, 'sheet-main'), { kind: 'error', code }, `${name}(${argument})`)
    }
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学', name)
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.filter(entry => entry.name === name).length, 1, name)
    assert.throws(() => functionCatalog.buildSheetFunctionFormula(name, ''), /请输入函数参数/)
  }
  assert.equal(formula.evaluateFormula('=DEGREES(-2)', model, 'sheet-main'), -114.59155902616465)
  assert.equal(formula.evaluateFormula('=RADIANS(270)', model, 'sheet-main'), 4.71238898038469)
  assert.equal(formula.evaluateFormula('=DEGREES("1e308")', model, 'sheet-main'), Infinity)
  assert.equal(formula.evaluateFormula('=DEGREES("-1e308")', model, 'sheet-main'), -Infinity)
  assert.deepEqual(formula.evaluateFormula('=RADIANS(DEGREES("1e308"))', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
})

test('inverse trigonometric scalar functions share fixed coercion and arity while retaining individual domains', () => {
  const source = structuredClone(model)
  for (const [name, number, expected] of [
    ['ACOS', 0.5, 1.0471975511965979], ['ACOSH', 2, 1.3169578969248166],
    ['ACOT', -2, 2.677945044588987], ['ACOTH', 2, 0.5493061443340549],
    ['ASIN', 0.5, 0.5235987755982989], ['ASINH', 10, 2.99822295029797],
    ['ATAN', 1, 0.7853981633974483], ['ATANH', 0.5, 0.5493061443340548],
  ]) {
    source.sheets[0].cells['0:0'].value = number
    for (const argument of [String(number), `"${number}"`, 'A1', 'A1:A1', 'INDEX(A1:A1,1)']) {
      assert.ok(Math.abs(formula.evaluateFormula(`=${name}(${argument})`, source, 'sheet-main') - expected) < 1e-12, `${name}(${argument})`)
    }
    for (const argumentsText of ['', '1,2', '1/0,2', '#REF!,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, source, 'sheet-main'), { kind: 'error', code: '#N/A' }, `${name}(${argumentsText})`)
    }
    for (const argument of ['""', '" "', '"invalid"']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, source, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, `${name}(${argument})`)
    }
    for (const [argument, code] of [['NA()', '#N/A'], ['1/0', '#DIV/0!'], ['#REF!', '#REF!']]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, source, 'sheet-main'), { kind: 'error', code }, `${name}(${argument})`)
    }
    assert.deepEqual(formula.evaluateFormula(`=${name}("1e999")`, source, 'sheet-main'), { kind: 'error', code: name === 'ACOT' || name === 'ACOTH' ? '#VALUE!' : '#NUM!' }, name)
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name).category, '数学', name)
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.filter(entry => entry.name === name).length, 1, name)
    assert.throws(() => functionCatalog.buildSheetFunctionFormula(name, ''), /请输入函数参数/)
  }
  for (const name of ['ACOS', 'ASIN']) {
    for (const number of [-1.01, 1.01]) assert.deepEqual(formula.evaluateFormula(`=${name}(${number})`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  }
  for (const [name, argument] of [['ACOSH', '0.99'], ['ATANH', '-1'], ['ATANH', '1'], ['ACOTH', '-1'], ['ACOTH', '1'], ['ACOTH', '0']]) {
    assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' }, `${name}(${argument})`)
  }
  assert.equal(formula.evaluateFormula('=ACOS(TRUE)', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=ACOSH(TRUE)', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=ACOT(FALSE)', model, 'sheet-main'), Math.PI / 2)
  assert.equal(formula.evaluateFormula('=ACOT(A9:A9)', model, 'sheet-main'), Math.PI / 2)
  assert.equal(formula.evaluateFormula('=ACOT(-0)', model, 'sheet-main'), -Math.PI / 2)
  assert.equal(formula.evaluateFormula('=ASIN(FALSE)', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=ASINH(A9:A9)', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=ATAN(A9:A9)', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=ATANH(FALSE)', model, 'sheet-main'), 0)
  assert.ok(formula.evaluateFormula('=ACOTH(-2)', model, 'sheet-main') < 0)
})

test('hyperbolic scalar functions preserve coercion, reciprocal zero errors and fixed overflow ordering', () => {
  for (const [name, atOne, atTwo, atZero] of [
    ['SINH', 1.1752011936438014, 3.626860407847019, 0],
    ['COSH', 1.5430806348152437, 3.7621956910836314, 1],
    ['TANH', 0.7615941559557649, 0.9640275800758169, 0],
    ['COTH', 1.3130352854993315, 1.0373147207275482, '#DIV/0!'],
    ['SECH', 0.6480542736638855, 0.2658022288340797, 1],
    ['CSCH', 0.8509181282393216, 0.2757205647717832, '#DIV/0!'],
  ]) {
    for (const argument of ['1', 'TRUE', '"1"', '"$1"', '"100%"']) {
      assert.ok(Math.abs(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main') - atOne) < 1e-12, `${name}(${argument})`)
    }
    for (const argument of ['A1', 'A1:A1', 'INDEX(A1:A1,1)']) {
      assert.ok(Math.abs(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main') - atTwo) < 1e-12, `${name}(${argument})`)
    }
    for (const argument of ['0', 'FALSE', 'A9:A9']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), typeof atZero === 'number' ? atZero : { kind: 'error', code: atZero }, `${name}(${argument})`)
    }
    for (const argumentsText of ['', '1,2', '1/0,2', '#REF!,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, `${name}(${argumentsText})`)
    }
    for (const [argument, code] of [['""', '#VALUE!'], ['"invalid"', '#VALUE!'], ['1/0', '#DIV/0!'], ['NA()', '#N/A'], ['#REF!', '#REF!']]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), { kind: 'error', code }, `${name}(${argument})`)
    }
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name).category, '数学', name)
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.filter(entry => entry.name === name).length, 1, name)
    assert.throws(() => functionCatalog.buildSheetFunctionFormula(name, ''), /请输入函数参数/)
  }
  for (const name of ['SINH', 'COSH']) assert.deepEqual(formula.evaluateFormula(`=${name}(1000)`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  for (const name of ['SECH', 'CSCH']) {
    for (const argument of ['1000', '-1000', '134217728', '"1e999"']) assert.equal(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), 0)
  }
  for (const name of ['TANH', 'COTH']) {
    assert.equal(formula.evaluateFormula(`=${name}(1000)`, model, 'sheet-main'), 1)
    assert.equal(formula.evaluateFormula(`=${name}(-1000)`, model, 'sheet-main'), -1)
  }
  for (const name of ['COTH', 'CSCH']) {
    assert.deepEqual(formula.evaluateFormula(`=${name}("1e-320")`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
    assert.deepEqual(formula.evaluateFormula(`=${name}(LOG("1e999","1e999"))`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  }
  assert.equal(formula.evaluateFormula('=SECH(LOG("1e999","1e999"))', model, 'sheet-main'), 0)
  assert.deepEqual(formula.evaluateFormula('=COTH("1e999")', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
})

test('exponential, logarithmic and sign functions retain fixed domains, defaults and error precedence', () => {
  for (const [name, argumentsText, expected] of [['EXP', '1', 2.718281828459045], ['LN', 'EXP(3)', 3], ['LOG10', '100000', 5], ['SIGN', '-10', -1]]) {
    assert.ok(Math.abs(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main') - expected) < 1e-12, name)
    for (const invalidArity of ['', '1,2', '#REF!,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${invalidArity})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, name)
    }
    for (const [argument, code] of [['""', '#VALUE!'], ['"invalid"', '#VALUE!'], ['1/0', '#DIV/0!'], ['NA()', '#N/A']]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), { kind: 'error', code }, name)
    }
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name).category, '数学', name)
  }
  assert.equal(formula.evaluateFormula('=EXP(FALSE)', model, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=EXP(-1000)', model, 'sheet-main'), 0)
  assert.deepEqual(formula.evaluateFormula('=EXP(1000)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.equal(formula.evaluateFormula('=LN("100%")', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=LOG10("$100")', model, 'sheet-main'), 2)
  for (const name of ['LN', 'LOG10']) {
    for (const argument of ['0', '-1', 'FALSE', 'A9:A9', '"1e999"']) assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' }, `${name}(${argument})`)
  }
  assert.equal(formula.evaluateFormula('=SIGN("-25%")', model, 'sheet-main'), -1)
  assert.equal(formula.evaluateFormula('=SIGN(-0)', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=SIGN("1e999")', model, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=SIGN("-1e999")', model, 'sheet-main'), -1)
  assert.equal(formula.evaluateFormula('=SIGN(LOG("1e999","1e999"))', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=LOG(100)', model, 'sheet-main'), 2)
  assert.equal(formula.evaluateFormula('=LOG("8","2")', model, 'sheet-main'), 3)
  assert.equal(formula.evaluateFormula('=LOG(0.25,0.5)', model, 'sheet-main'), 2)
  assert.equal(formula.evaluateFormula('=LOG(TRUE)', model, 'sheet-main'), 0)
  for (const argumentsText of ['0', '-1', '8,0', '8,-2', 'A9:A9']) assert.deepEqual(formula.evaluateFormula(`=LOG(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' }, argumentsText)
  assert.deepEqual(formula.evaluateFormula('=LOG(8,TRUE)', model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  for (const argumentsText of ['', '8,2,1', '#REF!,2,1']) assert.deepEqual(formula.evaluateFormula(`=LOG(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.deepEqual(formula.evaluateFormula('=LOG("invalid",#REF!)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.deepEqual(formula.evaluateFormula('=LOG(NA(),1/0)', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(formula.evaluateFormula('=LOG("1e999")', model, 'sheet-main'), Infinity)
  assert.equal(formula.evaluateFormula('=LOG(8,"1e999")', model, 'sheet-main'), 0)
  assert.ok(Number.isNaN(formula.evaluateFormula('=LOG("1e999","1e999")', model, 'sheet-main')))
  assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'LOG').category, '数学')
})

test('ATAN2 preserves coordinate order, quadrants, scalar references and fixed boolean-origin handling', () => {
  for (const [argumentsText, expected] of [
    ['1,2', 1.1071487177940904], ['4,3', 0.6435011087932844], ['-2,-3', -2.158798930342464],
    ['"0.5",2', 1.3258176636680326], ['0,1', Math.PI / 2], ['-1,0', Math.PI],
    ['0,-1', -Math.PI / 2], ['1,0', 0], ['A1:A1,B1:B1', 0.982793723247329],
    ['FALSE,1', Math.PI / 2], ['FALSE,FALSE', 0], ['FALSE,0', 0],
  ]) assert.ok(Math.abs(formula.evaluateFormula(`=ATAN2(${argumentsText})`, model, 'sheet-main') - expected) < 1e-12, argumentsText)
  for (const argumentsText of ['', '1', '1,2,3', '#REF!,1,2']) assert.deepEqual(formula.evaluateFormula(`=ATAN2(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  for (const argumentsText of ['0,0', '0,FALSE', 'A9:A9,0']) assert.deepEqual(formula.evaluateFormula(`=ATAN2(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  for (const [argumentsText, code] of [['"invalid",#REF!', '#VALUE!'], ['1,#REF!', '#REF!'], ['NA(),1/0', '#N/A'], ['"1e999",1', '#NUM!']]) {
    assert.deepEqual(formula.evaluateFormula(`=ATAN2(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  }
  assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'ATAN2').category, '数学')
  assert.throws(() => functionCatalog.buildSheetFunctionFormula('ATAN2', ''), /请输入函数参数/)
})

test('EVEN and ODD preserve outward rounding, fixed epsilon and scalar conversion contracts', () => {
  for (const [name, argumentsText, expected] of [
    ['EVEN', '1.5', 2], ['EVEN', '-3.5', -4], ['EVEN', '0', 0], ['EVEN', '-0', 0],
    ['EVEN', 'TRUE', 2], ['EVEN', 'FALSE', 0], ['EVEN', '"100%"', 2], ['EVEN', '"$1.5"', 2],
    ['EVEN', 'A1:A1', 2], ['EVEN', 'INDEX(A1:A1,1)', 2], ['EVEN', 'A9:A9', 0],
    ['EVEN', '2.0000000000000004', 2], ['EVEN', '2.000000000000002', 4], ['EVEN', '"1e-320"', 0],
    ['ODD', '1.5', 3], ['ODD', '-3.5', -5], ['ODD', '0', 1], ['ODD', '-0', 1],
    ['ODD', 'TRUE', 1], ['ODD', 'FALSE', 1], ['ODD', '"100%"', 1], ['ODD', '"$1.5"', 3],
    ['ODD', 'A1:A1', 3], ['ODD', 'INDEX(A1:A1,1)', 3], ['ODD', 'A9:A9', 1],
    ['ODD', '3.0000000000000004', 3], ['ODD', '3.000000000000002', 5], ['ODD', '"-1e-320"', -1],
  ]) assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), expected, `${name}(${argumentsText})`)
  for (const name of ['EVEN', 'ODD']) {
    for (const argumentsText of ['', '1,2', '#REF!,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, name)
    }
    for (const argument of ['""', '" "', '"invalid"', '"1e999"', 'LOG("1e999","1e999")', 'A1:B1']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, `${name}(${argument})`)
    }
    for (const [argument, code] of [['NA()', '#N/A'], ['1/0', '#DIV/0!'], ['#REF!', '#REF!']]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), { kind: 'error', code }, name)
    }
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学')
    assert.equal(entry.example, '1.5')
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, entry.example), model, 'sheet-main'), name === 'EVEN' ? 2 : 3)
  }
})

test('QUOTIENT rejects booleans and preserves direct numeric text, error order and non-finite results', () => {
  for (const [argumentsText, expected] of [
    ['5,2', 2], ['-5,2', -2], ['5,-2', -2], ['-5,-2', 2], ['1,-10', -0],
    ['"5.9","2"', 2], ['"0x10",2', 8], ['"",2', 0], ['" ",2', 0],
    ['A1:A1,B1:B1', 0], ['INDEX(A1:A1,1),1', 2], ['A9:A9,2', 0],
    ['"1e999",2', Infinity], ['5,"1e999"', 0],
  ]) assert.equal(formula.evaluateFormula(`=QUOTIENT(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  for (const [argumentsText, code] of [
    ['TRUE,2', '#VALUE!'], ['5,FALSE', '#VALUE!'], ['TRUE,NA()', '#VALUE!'],
    ['"invalid",NA()', '#N/A'], ['#REF!,TRUE', '#REF!'], ['"100%",2', '#VALUE!'],
    ['"$5",2', '#VALUE!'], ['"2024-01-07",2', '#VALUE!'], ['A1:B1,2', '#VALUE!'],
    ['5,0', '#DIV/0!'], ['5,A9:A9', '#DIV/0!'], ['5,""', '#DIV/0!'],
    ['LOG("1e999","1e999"),0', '#VALUE!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=QUOTIENT(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  for (const argumentsText of ['', '5', '5,2,1', '#REF!,TRUE,1']) {
    assert.deepEqual(formula.evaluateFormula(`=QUOTIENT(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  assert.ok(Number.isNaN(formula.evaluateFormula('=QUOTIENT("1e999","1e999")', model, 'sheet-main')))
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'QUOTIENT')
  assert.equal(entry.category, '数学')
  assert.equal(entry.example, '5,2')
  assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(entry.name, entry.example), model, 'sheet-main'), 2)
})

test('SQRTPI retains the fixed negative domain and unfiltered non-finite result semantics', () => {
  for (const argument of ['1', 'TRUE', '"100%"', '"$1"']) {
    assert.ok(Math.abs(formula.evaluateFormula(`=SQRTPI(${argument})`, model, 'sheet-main') - 1.7724538509055159) < 1e-12, argument)
  }
  for (const argument of ['A1', 'A1:A1', 'INDEX(A1:A1,1)']) {
    assert.ok(Math.abs(formula.evaluateFormula(`=SQRTPI(${argument})`, model, 'sheet-main') - 2.5066282746310002) < 1e-12, argument)
  }
  assert.equal(formula.evaluateFormula('=SQRTPI(FALSE)', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=SQRTPI(A9:A9)', model, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=SQRTPI(-0)', model, 'sheet-main'), -0)
  for (const argument of ['-1', '"-1e999"']) assert.deepEqual(formula.evaluateFormula(`=SQRTPI(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  for (const argument of ['""', '"invalid"', 'A1:B1']) assert.deepEqual(formula.evaluateFormula(`=SQRTPI(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  for (const argumentsText of ['', '1,2', '1/0,2']) assert.deepEqual(formula.evaluateFormula(`=SQRTPI(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.deepEqual(formula.evaluateFormula('=SQRTPI(NA())', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(formula.evaluateFormula('=SQRTPI("1e308")', model, 'sheet-main'), Infinity)
  assert.equal(formula.evaluateFormula('=SQRTPI("1e999")', model, 'sheet-main'), Infinity)
  assert.ok(Number.isNaN(formula.evaluateFormula('=SQRTPI(LOG("1e999","1e999"))', model, 'sheet-main')))
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'SQRTPI')
  assert.equal(entry.category, '数学')
  assert.equal(entry.example, '1')
  assert.ok(Math.abs(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(entry.name, entry.example), model, 'sheet-main') - 1.7724538509055159) < 1e-12)
})

test('factorials retain floor conversion, single or double step, fixed NaN result and overflow errors', () => {
  for (const [name, argument, expected] of [
    ['FACT', '5', 120], ['FACT', '5.9', 120], ['FACT', '"$5"', 120], ['FACT', '"50%"', 1],
    ['FACT', 'A1:A1', 2], ['FACT', 'INDEX(A1:A1,1)', 2],
    ['FACTDOUBLE', '6', 48], ['FACTDOUBLE', '7', 105], ['FACTDOUBLE', '6.9', 48],
  ]) assert.equal(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), expected)
  for (const name of ['FACT', 'FACTDOUBLE']) {
    for (const argument of ['0', '-0', 'TRUE', 'FALSE', 'A9:A9', 'LOG("1e999","1e999")']) {
      assert.equal(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), 1, `${name}(${argument})`)
    }
    for (const [argument, code] of [
      ['-0.1', '#NUM!'], ['1000', '#NUM!'], ['"1e999"', '#NUM!'], ['"-1e999"', '#NUM!'],
      ['""', '#VALUE!'], ['"invalid"', '#VALUE!'], ['A1:B1', '#VALUE!'], ['#REF!', '#REF!'], ['1/0', '#DIV/0!'],
    ]) assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), { kind: 'error', code })
    for (const argumentsText of ['', '1,2', '1/0,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    }
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学')
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, entry.example), model, 'sheet-main'), name === 'FACT' ? 120 : 48)
  }
  assert.ok(Number.isFinite(formula.evaluateFormula('=FACT(170)', model, 'sheet-main')))
  assert.deepEqual(formula.evaluateFormula('=FACT(171)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
})

test('combinations retain chosen limits, repetition, floor conversion and fixed error priority', () => {
  for (const [name, argumentsText, expected] of [
    ['COMBIN', '8,2', 28], ['COMBIN', '3.9,2.9', 3], ['COMBIN', '"800%","$2"', 28],
    ['COMBIN', 'A1:A1,1', 2], ['COMBIN', 'INDEX(A1:A1,1),1', 2], ['COMBIN', 'TRUE,FALSE', 1],
    ['COMBIN', '0,0', 1], ['COMBIN', '"1e999",0', 1], ['COMBIN', '"1e999","1e999"', 1],
    ['COMBIN', 'LOG("1e999","1e999"),2', 1],
    ['COMBINA', '8,2', 36], ['COMBINA', '2,3', 4], ['COMBINA', '3.9,2.9', 6],
    ['COMBINA', '0,0', 1], ['COMBINA', 'A9:A9,0', 1], ['COMBINA', '1,"1e999"', 1],
  ]) assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), expected, `${name}(${argumentsText})`)
  for (const name of ['COMBIN', 'COMBINA']) {
    for (const [argumentsText, code] of [
      ['-0.1,0', '#NUM!'], ['2,-0.1', '#NUM!'], ['0,1', '#NUM!'], ['10000,5000', '#NUM!'],
      ['"invalid",#REF!', '#VALUE!'], ['#REF!,"invalid"', '#REF!'], ['-1,NA()', '#N/A'],
      ['"",1', '#VALUE!'], ['2,1/0', '#DIV/0!'], ['A1:B1,1', '#VALUE!'],
    ]) assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, `${name}(${argumentsText})`)
    for (const argumentsText of ['', '2', '2,1,0', '#REF!,1,0']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    }
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学')
    assert.equal(entry.example, '8,2')
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, entry.example), model, 'sheet-main'), name === 'COMBIN' ? 28 : 36)
  }
  assert.deepEqual(formula.evaluateFormula('=COMBIN(2,3)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
})

test('rounding functions truncate precision, apply fixed epsilon and retain boolean precision semantics', () => {
  for (const [name, argumentsText, expected] of [
    ['ROUND', '2.15,1', 2.2], ['ROUND', '2.149,1', 2.1], ['ROUND', '-1.475,2', -1.48],
    ['ROUND', '1.005,2', 1.01], ['ROUND', '-1.005,2', -1.01], ['ROUND', '2.15,1.9', 2.2],
    ['ROUND', '149,-1.9', 150], ['ROUND', '"$21.5","-1"', 20], ['ROUND', '-1.25,TRUE', -1.2],
    ['ROUNDDOWN', '2.15,1', 2.1], ['ROUNDDOWN', '-1.475,2', -1.47], ['ROUNDDOWN', '"21.5","-1"', 20],
    ['ROUNDDOWN', '-1.25,TRUE', -1.3], ['ROUNDDOWN', '1.2,1', 1.2], ['ROUNDDOWN', '2.15,1.9', 2.1],
    ['ROUNDUP', '2.15,1', 2.2], ['ROUNDUP', '-1.475,2', -1.48], ['ROUNDUP', '"21.5","-1"', 30],
    ['ROUNDUP', '-1.25,TRUE', -1.2], ['ROUNDUP', '1.2,1', 1.2], ['ROUNDUP', '2.15,1.9', 2.2],
  ]) assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), expected, `${name}(${argumentsText})`)
  for (const name of ['ROUND', 'ROUNDDOWN', 'ROUNDUP']) {
    for (const argument of ['A1', 'A1:A1', 'INDEX(A1:A1,1)']) {
      assert.equal(formula.evaluateFormula(`=${name}(${argument},1)`, model, 'sheet-main'), 2)
    }
    assert.equal(formula.evaluateFormula(`=${name}(2.15,A9:A9)`, model, 'sheet-main'), name === 'ROUNDUP' ? 3 : 2)
    for (const [argumentsText, code] of [
      ['"invalid",#REF!', '#VALUE!'], ['#REF!,"invalid"', '#REF!'], ['1,1/0', '#DIV/0!'],
      ['"",1', '#VALUE!'], ['A1:B1,1', '#VALUE!'], ['"1e999",0', '#NUM!'],
      ['1,"1e999"', '#NUM!'], ['1,309', '#NUM!'], ['1,-324', '#NUM!'],
    ]) assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code })
    for (const argumentsText of ['', '1', '1,2,3', '1/0,2,3']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    }
    const nonfiniteBoolean = formula.evaluateFormula(`=${name}("1e999",TRUE)`, model, 'sheet-main')
    if (name === 'ROUNDUP') assert.ok(Number.isNaN(nonfiniteBoolean))
    if (name !== 'ROUNDUP') assert.equal(nonfiniteBoolean, Infinity)
  }
  assert.ok(Object.is(formula.evaluateFormula('=ROUND(-0.1,0)', model, 'sheet-main'), -0))
  assert.ok(Object.is(formula.evaluateFormula('=ROUNDDOWN(-0.1,0)', model, 'sheet-main'), -0))
})

test('TRUNC retains optional precision, signed epsilon, nonfinite results and fixed conversion order', () => {
  for (const [argumentsText, expected] of [
    ['2.15', 2], ['2.149,2', 2.14], ['-1.475,2', -1.47], ['"$21.5","-1"', 20],
    ['2.15,TRUE', 2.1], ['2.15,A9:A9', 2], ['0.45,1.9', 0.4], ['A1:A1', 2], ['TRUE', 1],
    ['A9:A9', 0], ['-8', -7], ['-1.2,1', -1.1], ['"1e999"', Infinity],
  ]) assert.equal(formula.evaluateFormula(`=TRUNC(${argumentsText})`, model, 'sheet-main'), expected, `TRUNC(${argumentsText})`)
  for (const argumentsText of ['"-1e999"', '1,309', '1,-324', 'LOG("1e999","1e999")']) {
    assert.ok(Number.isNaN(formula.evaluateFormula(`=TRUNC(${argumentsText})`, model, 'sheet-main')), argumentsText)
  }
  for (const [argumentsText, code] of [
    ['"invalid",#REF!', '#VALUE!'], ['#REF!,"invalid"', '#REF!'], ['1,1/0', '#DIV/0!'],
    ['"",1', '#VALUE!'], ['A1:B1', '#VALUE!'], ['1,"invalid"', '#VALUE!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=TRUNC(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code })
  for (const argumentsText of ['', '1,2,3', '1/0,2,3']) {
    assert.deepEqual(formula.evaluateFormula(`=TRUNC(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  assert.ok(Object.is(formula.evaluateFormula('=TRUNC(-0.1)', model, 'sheet-main'), -0))
})

test('ROUNDBANK rounds decimal halves to even and retains clamping and original error priority', () => {
  for (const [argumentsText, expected] of [
    ['2.344,2', 2.34], ['2.346,2', 2.35], ['2.345,2', 2.34], ['2.3451,2', 2.35],
    ['2.355,2', 2.36], ['-2.345,2', -2.34], ['-2.355,2', -2.36], ['-222.346,-2', -200],
    ['"$2.345",2.9', 2.34], ['2.15,TRUE', 2.2], ['2.15,A9:A9', 2], ['A1:A1,0', 2],
    ['2.344,2000', 2.344], ['2.344,-2000', 0], ['"1e999",17', Infinity], ['5,"1e999"', 5],
    ['5,"-1e999"', 0], ['A9:A9,0', 0],
  ]) assert.equal(formula.evaluateFormula(`=ROUNDBANK(${argumentsText})`, model, 'sheet-main'), expected, `ROUNDBANK(${argumentsText})`)
  for (const [argumentsText, code] of [
    ['"invalid",#REF!', '#REF!'], ['#REF!,"invalid"', '#REF!'], ['1,1/0', '#DIV/0!'],
    ['"invalid",0', '#VALUE!'], ['"",1', '#VALUE!'], ['A1:B1,1', '#VALUE!'], ['1,"invalid"', '#VALUE!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=ROUNDBANK(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code })
  for (const argumentsText of ['', '1', '1,2,3', '1/0,2,3']) {
    assert.deepEqual(formula.evaluateFormula(`=ROUNDBANK(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  assert.ok(Number.isNaN(formula.evaluateFormula('=ROUNDBANK(LOG("1e999","1e999"),2)', model, 'sheet-main')))
  for (const [name, example, expected] of [
    ['ROUNDBANK', '2.345,2', 2.34], ['ROUNDDOWN', '3.2,0', 3], ['ROUNDUP', '3.2,0', 4], ['TRUNC', '0.45,1', 0.4],
  ]) {
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学')
    assert.equal(entry.example, example)
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, entry.example), model, 'sheet-main'), expected)
  }
})

test('GCD, LCM and MULTINOMIAL retain integer aggregation, blanks, numeric text and fixed limits', () => {
  for (const [name, argumentsText, expected] of [
    ['GCD', '5,2', 1], ['GCD', '1000000,3222', 2], ['GCD', '"1.5",8', 1],
    ['GCD', '0,0', 0], ['GCD', 'A9:A10', 0], ['GCD', 'A9:A9,8', 8],
    ['GCD', '9007199254740991', 9007199254740991],
    ['LCM', '5,2', 10], ['LCM', '1000000,3222', 1611000000], ['LCM', '"1.5",8', 8],
    ['LCM', '0,0', 0], ['LCM', 'A9:A10', 0], ['LCM', 'A9:A9,8', 8],
    ['LCM', '9007199254740992,0', 0],
    ['MULTINOMIAL', '5,2', 21], ['MULTINOMIAL', '"1.5",8', 9],
    ['MULTINOMIAL', '0,0', 1], ['MULTINOMIAL', 'A9:A10', 1], ['MULTINOMIAL', '170', 1],
  ]) assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), expected, `${name}(${argumentsText})`)
  for (const [name, expected] of [['GCD', 0], ['LCM', 0], ['MULTINOMIAL', 1]]) {
    assert.equal(formula.evaluateFormula(`=${name}(${Array(255).fill('0').join(',')})`, model, 'sheet-main'), expected)
    for (const argumentsText of ['', Array(256).fill('0').join(',')]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    }
    for (const [argumentsText, code] of [
      ['TRUE,8', '#VALUE!'], ['FALSE', '#VALUE!'], ['"",8', '#VALUE!'], ['" ",8', '#VALUE!'],
      ['"invalid",#REF!', '#VALUE!'], ['-0.1,8', '#NUM!'], ['#REF!,8', '#REF!'], ['0,#REF!', '#REF!'],
    ]) assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code })
  }
  assert.deepEqual(formula.evaluateFormula('=GCD(9007199254740992)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(formula.evaluateFormula('=LCM(9007199254740992)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=LCM(9007199254740992,#REF!)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.deepEqual(formula.evaluateFormula('=MULTINOMIAL(170,1)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(formula.evaluateFormula('=MULTINOMIAL(171,#REF!)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  for (const [name, example, expected] of [['GCD', '5,2', 1], ['LCM', '5,2', 10], ['MULTINOMIAL', '5,2', 21], ['MROUND', '10,3', 9]]) {
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学')
    assert.equal(entry.example, example)
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, entry.example), model, 'sheet-main'), expected)
  }
})

test('integer aggregates traverse sparse ranges in row order and propagate range errors before later values', () => {
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  for (const [row, col, value] of [[1, 1, 3.9], [0, 1, '2.9'], [1, 0, null], [0, 0, 4.9], [5, 3, 171], [4, 3, true], [4, 2, '#REF!']]) {
    sheet.cells[`${row}:${col}`] = { ...sheet.cells['0:0'], id: `${row}:${col}`, row, col, value,
      formula: null, display: value === null ? '' : String(value), dataType: value === '#REF!' ? 'error' : undefined }
  }
  for (const [name, expected] of [['GCD', 1], ['LCM', 12], ['MULTINOMIAL', 1260]]) {
    assert.equal(formula.evaluateFormula(`=${name}(A1:B2)`, source, 'sheet-main'), expected)
    assert.deepEqual(formula.evaluateFormula(`=${name}(C5:D6)`, source, 'sheet-main'), { kind: 'error', code: '#REF!' })
  }
  for (const [value, code] of [[-1, '#NUM!'], [1, '#VALUE!']]) {
    sheet.cells['4:2'] = { ...sheet.cells['4:2'], value, display: String(value), dataType: 'number' }
    for (const name of ['GCD', 'LCM', 'MULTINOMIAL']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(C5:D6)`, source, 'sheet-main'), { kind: 'error', code })
    }
  }
  for (const name of ['GCD', 'LCM', 'MULTINOMIAL']) {
    assert.deepEqual(formula.evaluateFormula(`=${name}(D5)`, source, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  }
})

test('MROUND retains raw numeric coercion, signed multiples, scalar range constraints and original error priority', () => {
  for (const [argumentsText, expected] of [
    ['2.15,1', 2], ['"21.5","1"', 22], ['10,3', 9], ['-10,-3', -9], ['2.5,1', 3], ['-2.5,-1', -3],
    ['"",1', 0], ['" ",1', 0], ['"0x10",3', 15], ['A9:A9,1', 0], ['2.15,A9:A9', 0],
    ['A1:A1,B1:B1', 3], ['INDEX(A1:A1,1),3', 3], ['"1e999",0', 0],
    ['9007199254740992,1', 9007199254740994],
  ]) assert.equal(formula.evaluateFormula(`=MROUND(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  assert.ok(Math.abs(formula.evaluateFormula('=MROUND(6.05,0.1)', model, 'sheet-main') - 6.1) < 1e-14)
  assert.ok(Object.is(formula.evaluateFormula('=MROUND(0,-1)', model, 'sheet-main'), -0))
  assert.equal(formula.evaluateFormula('=MROUND("1e999",1)', model, 'sheet-main'), Infinity)
  assert.ok(Number.isNaN(formula.evaluateFormula('=MROUND("1e999","1e999")', model, 'sheet-main')))
  for (const [argumentsText, code] of [
    ['10,-3', '#NUM!'], ['-10,3', '#NUM!'], ['TRUE,1', '#VALUE!'], ['1,FALSE', '#VALUE!'],
    ['"25%",1', '#VALUE!'], ['"$2",1', '#VALUE!'], ['"invalid",0', '#VALUE!'],
    ['"invalid",#REF!', '#REF!'], ['TRUE,#REF!', '#REF!'], ['#REF!,TRUE', '#REF!'],
    ['A1:B1,1', '#VALUE!'], ['1,A1:B1', '#VALUE!'], ['LOG("1e999","1e999"),1', '#VALUE!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=MROUND(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  for (const argumentsText of ['', '1', '#REF!,1,2']) {
    assert.deepEqual(formula.evaluateFormula(`=MROUND(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
})

test('ABS, SQRT and INT retain scalar ranges, numeric text, date systems and fixed integer precision', () => {
  for (const [name, argumentsText, expected] of [
    ['ABS', '"-0.5"', 0.5], ['ABS', 'A1:A1', 2], ['SQRT', '"16"', 4], ['SQRT', 'INDEX(A1:A1,1)', Math.sqrt(2)],
    ['INT', '"-0.5"', -1], ['INT', '5.999999999999999', 6], ['INT', '999999999999.9', 999999999999],
    ['INT', '1212121212121212', 1212121212121212], ['INT', 'A1:A1', 2],
  ]) assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), expected, `${name}(${argumentsText})`)
  for (const name of ['ABS', 'SQRT', 'INT']) {
    assert.equal(formula.evaluateFormula(`=${name}(TRUE)`, model, 'sheet-main'), 1)
    for (const argumentsText of ['FALSE', 'A9:A9']) assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), 0)
    for (const argumentsText of ['""', '" "', '"invalid"', 'A1:B1']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
    }
    assert.deepEqual(formula.evaluateFormula(`=${name}(#REF!)`, model, 'sheet-main'), { kind: 'error', code: '#REF!' })
    for (const argumentsText of ['', '1,2', '#REF!,1']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    }
  }
  assert.ok(Object.is(formula.evaluateFormula('=SQRT(-0)', model, 'sheet-main'), -0))
  assert.deepEqual(formula.evaluateFormula('=SQRT(-1)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  for (const name of ['ABS', 'SQRT']) for (const argumentsText of ['"1e999"', 'LOG("1e999","1e999")']) {
    assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  }
  assert.equal(formula.evaluateFormula('=INT("1e999")', model, 'sheet-main'), Infinity)
  assert.ok(Number.isNaN(formula.evaluateFormula('=INT(LOG("1e999","1e999"))', model, 'sheet-main')))
  assert.equal(formula.evaluateFormula('=ABS("1904-1-1")', model, 'sheet-main'), 1462)
  assert.equal(formula.evaluateFormula('=ABS("1904-1-1")', { ...model, date1904: true }, 'sheet-main'), 0)
})

test('MOD and POWER retain ordered coercion, Boolean divisors, magnitude limits and zero-base errors', () => {
  for (const [name, argumentsText, expected] of [
    ['MOD', '"5","2"', 1], ['MOD', 'A1:A1,B1:B1', 2], ['MOD', '-5,2', 1], ['MOD', '5,-2', -1],
    ['MOD', '1125899999999,1', 0], ['MOD', '1125900000000,TRUE', 0], ['MOD', 'A9:A9,2', 0],
    ['POWER', '"5","2"', 25], ['POWER', 'A1:A1,B1:B1', 8], ['POWER', '2,TRUE', 2], ['POWER', '2,FALSE', 1],
    ['POWER', '0,2', 0], ['POWER', 'A9:A9,2', 0], ['POWER', '-0,3', 0],
  ]) assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), expected, `${name}(${argumentsText})`)
  assert.ok(Math.abs(formula.evaluateFormula('=MOD(5,1.23)', model, 'sheet-main') - 0.08) < 1e-14)
  assert.ok(Number.isNaN(formula.evaluateFormula('=MOD("1e999",TRUE)', model, 'sheet-main')))
  assert.ok(Number.isNaN(formula.evaluateFormula('=MOD(LOG("1e999","1e999"),TRUE)', model, 'sheet-main')))
  for (const [name, argumentsText, code] of [
    ['MOD', '1125900000000,1', '#NUM!'], ['MOD', '1,"1e999"', '#NUM!'], ['MOD', '"1e999",1', '#NUM!'],
    ['MOD', '1,0', '#DIV/0!'], ['MOD', '1,FALSE', '#DIV/0!'], ['MOD', '1,A9:A9', '#DIV/0!'], ['MOD', '"1e999",0', '#DIV/0!'],
    ['POWER', '0,0', '#NUM!'], ['POWER', '0,FALSE', '#NUM!'], ['POWER', '0,-1', '#DIV/0!'],
    ['POWER', '-1,0.5', '#NUM!'], ['POWER', '2,"1e999"', '#NUM!'], ['POWER', '"1e999",0', '#NUM!'],
    ['POWER', '2,LOG("1e999","1e999")', '#VALUE!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, `${name}(${argumentsText})`)
  for (const name of ['MOD', 'POWER']) {
    for (const [argumentsText, code] of [['"invalid",#REF!', '#VALUE!'], ['#REF!,"invalid"', '#REF!'], ['1,#REF!', '#REF!'], ['1,A1:B1', '#VALUE!']]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code })
    }
    for (const argumentsText of ['', '1', '#REF!,1,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    }
  }
})

test('SUMSQ retains scalar conversion, fixed arithmetic errors and its canonical math entry', () => {
  for (const [argumentsText, expected] of [['1,2', 5], ['2,"2"', 8], ['2,TRUE', 5], ['2,FALSE', 4], ['A9:A9', 0], ['2,A9:A10', 4]]) {
    assert.equal(formula.evaluateFormula(`=SUMSQ(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  }
  for (const [argumentsText, code] of [
    ['"invalid",#REF!', '#VALUE!'], ['#REF!,"invalid"', '#REF!'], ['1,#REF!', '#REF!'],
    ['"1e999"', '#VALUE!'], ['1e200', '#VALUE!'], ['1e154,1e154,1', '#NUM!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=SUMSQ(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  assert.equal(formula.evaluateFormula(`=SUMSQ(${Array(255).fill('1').join(',')})`, model, 'sheet-main'), 255)
  for (const argumentsText of ['', Array(256).fill('1').join(',')]) {
    assert.deepEqual(formula.evaluateFormula(`=SUMSQ(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'SUMSQ')
  assert.equal(entry.category, '数学')
  assert.equal(entry.example, '3,4')
  assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(entry.name, entry.example), model, 'sheet-main'), 25)
})

test('SUMSQ ignores typed text and Boolean references and preserves row-order errors and INDEX references', () => {
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  const template = sheet.cells['0:0']
  sheet.cells = {}
  for (const [row, col, value, dataType] of [
    [1, 4, -3, 'number'], [0, 3, true, 'boolean'], [0, 2, 1.23, 'number'], [0, 0, 1, 'number'],
    [1, 1, '100', 'text'], [1, 2, '2.34', 'text'], [0, 1, ' ', 'text'], [0, 4, false, 'boolean'], [0, 5, null, 'blank'],
  ]) sheet.cells[`${row}:${col}`] = { ...template, id: `${row}:${col}`, row, col, value, dataType, formula: null, display: value === null ? '' : String(value) }
  assert.ok(Math.abs(formula.evaluateFormula('=SUMSQ(2,A1:F2)', source, 'sheet-main') - 15.5129) < 1e-12)
  for (const reference of ['B2', 'B2:B2', 'INDEX(A1:F2,2,2)']) {
    assert.equal(formula.evaluateFormula(`=SUMSQ(${reference})`, source, 'sheet-main'), 0)
  }
  sheet.cells['1:1'] = { ...sheet.cells['1:1'], value: 100, dataType: 'number' }
  sheet.cells['1:2'] = { ...sheet.cells['1:2'], value: 2.34, dataType: 'number' }
  assert.ok(Math.abs(formula.evaluateFormula('=SUMSQ(2,A1:F2)', source, 'sheet-main') - 10020.9885) < 1e-9)
  sheet.cells['1:1'] = { ...sheet.cells['1:1'], style: { ...sheet.cells['1:1'].style, numberFormat: '@' } }
  sheet.cells['1:2'] = { ...sheet.cells['1:2'], style: { ...sheet.cells['1:2'].style, numberFormat: '@@@' } }
  assert.ok(Math.abs(formula.evaluateFormula('=SUMSQ(2,A1:F2)', source, 'sheet-main') - 15.5129) < 1e-12)
  assert.equal(formula.evaluateFormula('=SUMSQ(B2)', source, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=ABS(B2)', source, 'sheet-main'), 100)
  for (const [row, col, value, dataType] of [[3, 0, '#REF!', 'error'], [2, 2, 1, 'number'], [2, 1, 1e154, 'number'], [2, 0, 1e154, 'number']]) {
    sheet.cells[`${row}:${col}`] = { ...template, id: `${row}:${col}`, row, col, value, dataType, formula: null, display: String(value) }
  }
  assert.deepEqual(formula.evaluateFormula('=SUMSQ(A3:B3)', source, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(formula.evaluateFormula('=SUMSQ(A3:C3)', source, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=SUMSQ(A3:C4)', source, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.deepEqual(formula.evaluateFormula('=SUMSQ(A3:B3,#REF!)', source, 'sheet-main'), { kind: 'error', code: '#NUM!' })
})

test('SUMX functions retain single-cell conversion, raw Number text rules, ordered errors and canonical entries', () => {
  const names = ['SUMX2MY2', 'SUMX2PY2', 'SUMXMY2']
  for (const [argumentsText, expected] of [
    ['3,4', [-7, 25, 1]], ['"3","4"', [-7, 25, 1]], ['A1:A1,B1:B1', [-5, 13, 1]],
    ['" ",4', [-16, 16, 16]], ['"",4', [-16, 16, 16]], ['"0x10",2', [252, 260, 196]],
  ]) for (const [index, name] of names.entries()) {
    assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), expected[index], `${name}(${argumentsText})`)
  }
  for (const name of names) {
    for (const [argumentsText, code] of [
      ['TRUE,2', '#DIV/0!'], ['2,FALSE', '#DIV/0!'], ['"invalid",2', '#DIV/0!'], ['"25%",2', '#DIV/0!'],
      ['A9,2', '#VALUE!'], ['A9:A9,2', '#VALUE!'], ['"invalid",#REF!', '#REF!'], ['A1:B1,#REF!', '#REF!'],
      ['A1:B1,A1', '#N/A'], ['A1,A1:B1', '#N/A'],
    ]) assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
    for (const argumentsText of ['', '1', '#REF!,1,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    }
    assert.equal(formula.evaluateFormula(`=${name}("1e999",1)`, model, 'sheet-main'), Infinity)
    assert.ok(Number.isNaN(formula.evaluateFormula(`=${name}(LOG("1e999","1e999"),1)`, model, 'sheet-main')))
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学')
    assert.equal(entry.example, '3,4')
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, entry.example), model, 'sheet-main'), [-7, 25, 1][names.indexOf(name)])
  }
})

test('SUMX pairs equal-sized references in row order across different shapes and retains skipped-left error masking', () => {
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  const template = sheet.cells['0:0']
  sheet.cells = {}
  for (const [row, col, value, dataType] of [
    [1, 2, 6, 'number'], [0, 0, 1, 'number'], [0, 1, 2, 'number'], [0, 2, 3, 'number'], [1, 0, 4, 'number'], [1, 1, 5, 'number'],
    [0, 6, 1, 'number'], [0, 7, null, 'blank'], [1, 6, 'test', 'text'], [1, 7, 2, 'number'], [2, 6, 3, 'number'], [2, 7, true, 'boolean'],
    [5, 0, '9', 'text'], [5, 1, 2, 'number'], [5, 6, '#REF!', 'error'], [5, 7, 3, 'number'],
  ]) sheet.cells[`${row}:${col}`] = { ...template, id: `${row}:${col}`, row, col, value, dataType, formula: null, display: value === null ? '' : String(value) }
  const names = ['SUMX2MY2', 'SUMX2PY2', 'SUMXMY2']
  for (const [index, name] of names.entries()) {
    for (const argumentsText of ['A1:C2,G1:H3', 'C2:A1,H3:G1', 'INDEX(A1:C2,0,0),INDEX(G1:H3,0,0)', 'A1:INDEX(A1:C2,2,3),G1:H3']) {
      assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, source, 'sheet-main'), [28, 56, 8][index], argumentsText)
    }
  }
  for (const [value, dataType, numberFormat] of [['9', 'text', undefined], [true, 'boolean', undefined], [null, 'blank', undefined], [9, 'number', '@'], [9, 'number', '@@@']]) {
    sheet.cells['5:0'] = { ...sheet.cells['5:0'], value, dataType, style: { ...template.style, numberFormat } }
    for (const [index, name] of names.entries()) {
      assert.equal(formula.evaluateFormula(`=${name}(A6:B6,G6:H6)`, source, 'sheet-main'), [-5, 13, 1][index])
    }
  }
  sheet.cells['5:0'] = { ...sheet.cells['5:0'], value: 9, dataType: 'number', style: template.style }
  for (const name of names) assert.deepEqual(formula.evaluateFormula(`=${name}(A6:B6,G6:H6)`, source, 'sheet-main'), { kind: 'error', code: '#REF!' })
  sheet.cells['5:0'] = { ...sheet.cells['5:0'], value: '#DIV/0!', dataType: 'error' }
  sheet.cells['5:6'] = { ...sheet.cells['5:6'], value: 'ignored', dataType: 'text' }
  for (const name of names) {
    assert.deepEqual(formula.evaluateFormula(`=${name}(A6:B6,G6:H6)`, source, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
    assert.deepEqual(formula.evaluateFormula(`=${name}(A6:B6,G6)`, source, 'sheet-main'), { kind: 'error', code: '#N/A' })
    assert.deepEqual(formula.evaluateFormula(`=${name}(A9:B9,G6:H6)`, source, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  }
})

test('SUMPRODUCT matches rows and columns, broadcasts later scalars and retains first-array error precedence', () => {
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  const template = sheet.cells['0:0']
  sheet.cells = {}
  for (const [row, col, value, dataType] of [
    [0, 0, 1, 'number'], [0, 1, '2', 'text'], [0, 2, 3, 'number'], [1, 0, null, 'blank'], [1, 1, 'test', 'text'], [1, 2, 2, 'number'],
    [0, 6, 1, 'number'], [0, 7, 1, 'number'], [0, 8, 1, 'number'], [1, 6, null, 'blank'], [1, 7, 'test', 'text'], [1, 8, 2, 'number'],
  ]) sheet.cells[`${row}:${col}`] = { ...template, id: `${row}:${col}`, row, col, value, dataType, formula: null, display: value === null ? '' : String(value) }
  for (const [argumentsText, expected] of [['A1:C2', 6], ['A1:C2,G1:I2', 8], ['A1:C2,3', 18], ['A1:C2,G1:I2,3', 24], ['INDEX(A1:C2,0,0),G1:I2', 8]]) {
    assert.equal(formula.evaluateFormula(`=SUMPRODUCT(${argumentsText})`, source, 'sheet-main'), expected, argumentsText)
  }
  for (const argumentsText of ['A1:C2,G1:H3', 'A1:C2,G1:G1', '3,A1:C2', 'A1:C2,G1:H3,#REF!']) {
    assert.deepEqual(formula.evaluateFormula(`=SUMPRODUCT(${argumentsText})`, source, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  }
  sheet.cells['0:1'] = { ...sheet.cells['0:1'], value: 2, dataType: 'number', style: { ...template.style, numberFormat: '@' } }
  sheet.cells['0:2'] = { ...sheet.cells['0:2'], style: { ...template.style, numberFormat: '@@@' } }
  assert.equal(formula.evaluateFormula('=SUMPRODUCT(A1:C2,G1:I2)', source, 'sheet-main'), 5)
  sheet.cells['0:7'] = { ...sheet.cells['0:7'], value: '#REF!', dataType: 'error' }
  assert.deepEqual(formula.evaluateFormula('=SUMPRODUCT(A1:C2,G1:I2)', source, 'sheet-main'), { kind: 'error', code: '#REF!' })
  sheet.cells['1:2'] = { ...sheet.cells['1:2'], value: '#DIV/0!', dataType: 'error' }
  for (const argumentsText of ['A1:C2,#REF!', 'A1:C2,G1:H3', 'A1:C2,G1:I2']) {
    assert.deepEqual(formula.evaluateFormula(`=SUMPRODUCT(${argumentsText})`, source, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  }
})

test('SUMPRODUCT distinguishes one-argument addition errors from multi-argument arithmetic and parameter limits', () => {
  for (const [argumentsText, expected] of [['3,4', 12], ['"4"', 0], ['TRUE', 0], ['A9', 0], ['A1:B1,3', 15], ['1e308,2,"ignored"', 0]]) {
    assert.equal(formula.evaluateFormula(`=SUMPRODUCT(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  }
  assert.equal(formula.evaluateFormula('=SUMPRODUCT(1e308,2)', model, 'sheet-main'), Infinity)
  assert.ok(Number.isNaN(formula.evaluateFormula('=SUMPRODUCT(0,INT("1e999"))', model, 'sheet-main')))
  assert.deepEqual(formula.evaluateFormula('=SUMPRODUCT(LOG("1e999","1e999"))', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=SUMPRODUCT(0,#REF!)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.equal(formula.evaluateFormula(`=SUMPRODUCT(${Array(255).fill('1').join(',')})`, model, 'sheet-main'), 1)
  for (const argumentsText of ['', Array(256).fill('1').join(',')]) {
    assert.deepEqual(formula.evaluateFormula(`=SUMPRODUCT(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  for (const [col, value, dataType] of [[0, 1e308, 'number'], [1, 1e308, 'number'], [2, 1, 'number'], [3, '#REF!', 'error']]) {
    sheet.cells[`3:${col}`] = { ...sheet.cells['0:0'], id: `3:${col}`, row: 3, col, value, dataType, formula: null, display: String(value) }
  }
  for (const [reference, code] of [['A4:B4', '#NUM!'], ['A4:C4', '#VALUE!'], ['A4:D4', '#REF!']]) {
    assert.deepEqual(formula.evaluateFormula(`=SUMPRODUCT(${reference})`, source, 'sheet-main'), { kind: 'error', code })
  }
  sheet.cells['4:0'] = { ...sheet.cells['0:0'], id: '4:0', row: 4, col: 0, value: Infinity, dataType: 'number', formula: null }
  assert.deepEqual(formula.evaluateFormula('=SUMPRODUCT(A5)', source, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(formula.evaluateFormula('=SUMX2MY2(A5,1)', source, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  sheet.cells['4:0'] = { ...sheet.cells['4:0'], style: { ...sheet.cells['4:0'].style, numberFormat: '@' } }
  assert.equal(formula.evaluateFormula('=SUMPRODUCT(A5)', source, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=SUMX2MY2(A5,1)', source, 'sheet-main'), Infinity)
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'SUMPRODUCT')
  assert.equal(entry.category, '数学')
  assert.equal(entry.example, 'C2:C5,D2:D5')
  assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(entry.name, entry.example), model, 'sheet-main'), 0)
})

test('SERIESSUM retains phase-ordered errors, fractional exponents, typed coefficients and sparse coefficient positions', () => {
  for (const [argumentsText, expected] of [['2,1,1,3', 6], ['4,"0.5","1.5","3"', 6], ['2,A9:A9,1,3', 3], ['2,1,1,A9', 0], ['2,0,1,"25%"', 0.25]]) {
    assert.equal(formula.evaluateFormula(`=SERIESSUM(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  }
  for (const [argumentsText, code] of [
    ['"invalid",#REF!,0,1', '#REF!'], ['TRUE,#REF!,0,1', '#VALUE!'], ['"invalid",0,0,#REF!', '#VALUE!'],
    ['A1:B1,0,0,#REF!', '#VALUE!'], ['1,#REF!,A1:B1,1', '#REF!'], ['1,0,0,TRUE', '#VALUE!'],
    ['1,0,0," "', '#VALUE!'], ['#REF!,0,0,IF(FALSE,1,A9)', '#N/A'], ['2,IF(FALSE,1,A9),1,3', '#N/A'],
    ['0,-1,1,0', '#NUM!'], ['INT("1e999"),1,1,1', '#NUM!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=SERIESSUM(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  for (const argumentsText of ['', '1,2,3', '#REF!,1,2,3,4']) {
    assert.deepEqual(formula.evaluateFormula(`=SERIESSUM(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  const template = sheet.cells['0:0']
  sheet.cells = {}
  for (const [row, col, value, dataType] of [
    [0, 0, 1, 'number'], [0, 1, -0.5, 'number'], [1, 0, 0.041666667, 'number'], [1, 1, -0.001388889, 'number'],
    [4, 0, 1, 'number'], [4, 2, '3', 'text'],
  ]) sheet.cells[`${row}:${col}`] = { ...template, id: `${row}:${col}`, row, col, value, dataType, formula: null, display: String(value) }
  assert.ok(Math.abs(formula.evaluateFormula('=SERIESSUM(0.785398163,0,2,A1:B2)', source, 'sheet-main') - 0.707103215205) < 1e-12)
  for (const reference of ['A5:C5', 'INDEX(A5:C5,0,0)', 'A5:INDEX(A5:C5,1,3)']) {
    assert.equal(formula.evaluateFormula(`=SERIESSUM(2,0,1,${reference})`, source, 'sheet-main'), 13)
  }
  sheet.cells['4:1'] = { ...template, id: '4:1', row: 4, col: 1, value: null, dataType: 'blank', formula: null, display: '' }
  assert.equal(formula.evaluateFormula('=SERIESSUM(2,0,1,A5:C5)', source, 'sheet-main'), 13)
  sheet.cells['4:2'] = { ...sheet.cells['4:2'], value: '#REF!', dataType: 'error' }
  assert.deepEqual(formula.evaluateFormula('=SERIESSUM(9000,0,90000,A5:C5)', source, 'sheet-main'), { kind: 'error', code: '#REF!' })
  sheet.cells['4:1'] = { ...sheet.cells['4:1'], value: true, dataType: 'boolean' }
  assert.deepEqual(formula.evaluateFormula('=SERIESSUM(2,0,1,A5:C5)', source, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'SERIESSUM')
  assert.equal(entry.category, '数学')
  assert.equal(entry.example, '0.785398163,0,2,A1:A4')
  assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(entry.name, entry.example), model, 'sheet-main'), 2)
})

test('MDETERM retains square matrix values, row swaps, singular pivots and the fixed size-dependent threshold', () => {
  for (const [values, expected] of [
    [[[2]], 2], [[[0, 1], [1, 0]], -1], [[[1, 0], [0, 1e-12]], 1e-12],
    [[[1, 2, 3], [4, 5, 6], [7, 8, 9]], 0],
    [[[1.5, 2, 3], [2, 5, 6], [7, 8, 9]], -13.5],
    [[[0, 1, 0], [1, 0, 0], [0, 0, 1]], -1],
    [[[1, 0, 0], [0, 1, 0], [0, 0, 1e-11]], 0],
    [[[1, 0, 0], [0, 1, 0], [0, 0, 1.01e-11]], 1.01e-11],
    [[['1.5', 2, 3], [2, '5', 6], [7, 8, 9]], -13.5],
    [Array.from({ length: 30 }, (_, row) => Array.from({ length: 30 }, (_, column) => row === column ? 1 : 0)), 1],
  ]) {
    const source = structuredClone(model)
    const sheet = source.sheets[0]
    const template = sheet.cells['0:0']
    sheet.cells = {}
    for (const [row, cells] of values.entries()) {
      for (const [col, value] of cells.entries()) {
        sheet.cells[`${row}:${col}`] = { ...template, id: `${row}:${col}`, row, col, value, dataType: typeof value === 'string' ? 'text' : 'number', formula: null, display: String(value) }
      }
    }
    const lastColumn = values[0].length === 30 ? 'AD' : String.fromCharCode(64 + values[0].length)
    const reference = `A1:${lastColumn}${values.length}`
    const before = structuredClone(source)
    const result = formula.evaluateFormula(`=MDETERM(${reference})`, source, 'sheet-main')
    assert.equal(typeof result, 'number')
    assert.ok(Math.abs(result - expected) <= 1e-12 * Math.max(Math.abs(expected), 1e-11), reference)
    assert.deepEqual(source, before)
  }
})

test('MDETERM preserves scalar text conversion, signed zero, row-order errors and reference shape checks', () => {
  for (const [argumentsText, expected] of [['5', 5], ['"1,234"', 1234], ['"10%"', 0.1], ['A1:A1', 2], ['INDEX(A1:B1,1,1)', 2]]) {
    assert.equal(formula.evaluateFormula(`=MDETERM(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  }
  assert.ok(Object.is(formula.evaluateFormula('=MDETERM(-0)', model, 'sheet-main'), -0))
  assert.equal(formula.evaluateFormula('=MDETERM("1e999")', model, 'sheet-main'), Infinity)
  assert.ok(Number.isNaN(formula.evaluateFormula('=MDETERM(QUOTIENT("1e999","1e999"))', model, 'sheet-main')))
  for (const [argumentsText, code] of [['TRUE', '#VALUE!'], ['A9:A9', '#VALUE!'], ['" "', '#VALUE!'], ['"invalid"', '#VALUE!'], ['A1:B1', '#VALUE!'], ['#REF!', '#REF!']]) {
    assert.deepEqual(formula.evaluateFormula(`=MDETERM(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  }
  for (const argumentsText of ['', '1,#REF!']) {
    assert.deepEqual(formula.evaluateFormula(`=MDETERM(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  const source = structuredClone(model)
  const sheet = source.sheets[0]
  const template = sheet.cells['0:0']
  sheet.cells = {
    '0:2': { ...template, id: '0:2', col: 2, value: '#DIV/0!', dataType: 'error', formula: null },
    '0:1': { ...template, id: '0:1', col: 1, value: true, dataType: 'boolean', formula: null },
    '0:0': { ...template, value: '#REF!', dataType: 'error', formula: null },
  }
  assert.deepEqual(formula.evaluateFormula('=MDETERM(A1:C1)', source, 'sheet-main'), { kind: 'error', code: '#REF!' })
  sheet.cells['0:0'] = { ...sheet.cells['0:0'], value: null, dataType: 'blank' }
  assert.deepEqual(formula.evaluateFormula('=MDETERM(A1:C1)', source, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  sheet.cells['0:0'] = { ...sheet.cells['0:0'], value: 1, dataType: 'number' }
  assert.deepEqual(formula.evaluateFormula('=MDETERM(A1:C1)', source, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  sheet.cells['0:1'] = { ...sheet.cells['0:1'], value: 2, dataType: 'number' }
  assert.deepEqual(formula.evaluateFormula('=MDETERM(A1:C1)', source, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'MDETERM')
  assert.equal(entry.category, '数学')
  assert.equal(entry.example, 'A1:C3')
})

test('RAND returns finite unit-interval values and retains strict zero-argument arity', () => {
  for (let index = 0; index < 32; index++) {
    const value = formula.evaluateFormula('=RAND()', model, 'sheet-main')
    assert.equal(typeof value, 'number')
    assert.ok(value >= 0 && value < 1)
  }
  for (const argumentsText of ['0', '#REF!']) {
    assert.deepEqual(formula.evaluateFormula(`=RAND(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'RAND')
  assert.equal(entry.category, '数学')
  assert.equal(entry.example, '')
  assert.equal(functionCatalog.buildSheetFunctionFormula(entry.name, entry.example), '=RAND()')
})

test('RANDBETWEEN retains inclusive integer bounds, raw text conversion, scalar ranges and error priority', () => {
  for (const [argumentsText, lower, upper] of [['-10,10', -10, 10], ['"1.1","3.9"', 2, 3], ['A9,3', 0, 3]]) {
    for (let index = 0; index < 16; index++) {
      const value = formula.evaluateFormula(`=RANDBETWEEN(${argumentsText})`, model, 'sheet-main')
      assert.ok(Number.isInteger(value))
      assert.ok(value >= lower && value <= upper, argumentsText)
    }
  }
  for (const [argumentsText, expected] of [['7,7', 7], ['" ",""', 0], ['A1:A1,A1:A1', 2], ['0.1,0.2', 1], ['-1.9,-1.1', -1]]) {
    assert.equal(formula.evaluateFormula(`=RANDBETWEEN(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  }
  for (const [argumentsText, code] of [
    ['3,1', '#NUM!'], ['TRUE,3', '#VALUE!'], ['1,FALSE', '#VALUE!'], ['"10%",1', '#VALUE!'], ['"invalid",1', '#VALUE!'],
    ['TRUE,#REF!', '#REF!'], ['"invalid",#DIV/0!', '#DIV/0!'], ['#REF!,TRUE', '#REF!'],
    ['A1:B1,#REF!', '#VALUE!'], ['TRUE,A1:B1', '#VALUE!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=RANDBETWEEN(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  assert.ok(Number.isNaN(formula.evaluateFormula('=RANDBETWEEN("1e999","1e999")', model, 'sheet-main')))
  for (const argumentsText of ['', '1', '#REF!,1,2']) {
    assert.deepEqual(formula.evaluateFormula(`=RANDBETWEEN(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === 'RANDBETWEEN')
  assert.equal(entry.category, '数学')
  assert.equal(entry.example, '1,100')
})

test('ARABIC retains the fixed Roman character rules, signed zero, length limit and text-only inputs', () => {
  for (const [argumentsText, expected] of [
    ['"MXMVII"', 1997], ['"-mcmxii"', -1912], ['"IL"', 49], ['"ID"', 499], ['""', 0], ['A9:A9', 0],
    [`"${'IX'.repeat(127)}I"`, 1144],
  ]) assert.equal(formula.evaluateFormula(`=ARABIC(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  assert.ok(Object.is(formula.evaluateFormula('=ARABIC("-")', model, 'sheet-main'), -0))
  for (const argumentsText of ['1', 'TRUE', 'A1:A1', '" I"', '"I "', '"+I"', '"IIII"', '"VX"', '"LXM"', `"${'IX'.repeat(128)}"`, 'A1:B1']) {
    assert.deepEqual(formula.evaluateFormula(`=ARABIC(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, argumentsText)
  }
  assert.deepEqual(formula.evaluateFormula('=ARABIC(#REF!)', model, 'sheet-main'), { kind: 'error', code: '#REF!' })
  for (const argumentsText of ['', '"I",#REF!']) {
    assert.deepEqual(formula.evaluateFormula(`=ARABIC(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
})

test('ROMAN retains all five forms, Boolean forms, numeric text and original error priority', () => {
  for (const [form, expected] of [[0, 'CDXCIX'], [1, 'LDVLIV'], [2, 'XDIX'], [3, 'VDIV'], [4, 'ID']]) {
    assert.equal(formula.evaluateFormula(`=ROMAN(499,${form})`, model, 'sheet-main'), expected)
  }
  for (const [argumentsText, expected] of [
    ['499', 'CDXCIX'], ['499,TRUE', 'CDXCIX'], ['499,FALSE', 'ID'], ['"499.9","4.9"', 'ID'],
    ['"1,999"', 'MCMXCIX'], ['"49900%",4', 'ID'], ['TRUE', 'I'], ['FALSE', ''], ['A9', ''],
    ['0', ''], ['3999', 'MMMCMXCIX'], ['QUOTIENT("1e999","1e999"),0', ''],
  ]) assert.equal(formula.evaluateFormula(`=ROMAN(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  for (const [argumentsText, code] of [
    ['-0.1', '#VALUE!'], ['4000', '#VALUE!'], ['1,-0.1', '#VALUE!'], ['1,5', '#VALUE!'], ['" "', '#VALUE!'],
    ['"invalid",#REF!', '#REF!'], ['#DIV/0!,"invalid"', '#DIV/0!'], ['1,"invalid"', '#VALUE!'],
    ['A1:B1,0', '#VALUE!'], ['1,A1:B1', '#VALUE!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=ROMAN(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  for (const argumentsText of ['', '1,0,#REF!']) {
    assert.deepEqual(formula.evaluateFormula(`=ROMAN(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
})

test('BASE retains radix and padding bounds, ordered conversion, numeric text and uncapped minimum length', () => {
  for (const [argumentsText, expected] of [
    ['7,2', '111'], ['100,16', '64'], ['15,2,10', '0000001111'], ['"15.9","2.9","4.9"', '1111'],
    ['"1,024",2', '10000000000'], ['TRUE,2', '1'], ['A9,2', '0'], ['35,36', 'Z'],
    ['9007199254740991,16', '1FFFFFFFFFFFFF'], ['QUOTIENT("1e999","1e999"),16', 'NAN'],
  ]) assert.equal(formula.evaluateFormula(`=BASE(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  assert.equal(formula.evaluateFormula('=BASE(1,2,256)', model, 'sheet-main'), '0'.repeat(255) + '1')
  for (const [argumentsText, code] of [
    ['-0.1,2', '#NUM!'], ['9007199254740992,16', '#NUM!'], ['1,1', '#NUM!'], ['1,37', '#NUM!'], ['1,2,-0.1', '#NUM!'],
    ['" ",2', '#VALUE!'], ['1,"invalid"', '#VALUE!'], ['1,2,"invalid"', '#VALUE!'],
    ['"invalid",#REF!', '#REF!'], ['"invalid",2,#DIV/0!', '#DIV/0!'], ['#REF!,#DIV/0!', '#REF!'],
    ['A1:B1,2', '#VALUE!'], ['1,A1:B1', '#VALUE!'], ['1,2,A1:B1', '#VALUE!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=BASE(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  for (const argumentsText of ['', '1', '1,2,3,#REF!']) {
    assert.deepEqual(formula.evaluateFormula(`=BASE(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
})

test('DECIMAL retains raw digit text, numeric input bounds, whitespace and the fixed NaN radix path', () => {
  for (const [argumentsText, expected] of [
    ['"FF",16', 255], ['111,2', 7], ['"z",36', 35], ['"1e3",16', 483], ['1e3,16', 4096],
    ['"0xFF",36', 43323], ['A9,16', 0], ['"",2', 0], ['" \t",2', 0], ['"ff","16.9"', 255],
    ['"12x",QUOTIENT("1e999","1e999")', 12],
  ]) assert.equal(formula.evaluateFormula(`=DECIMAL(${argumentsText})`, model, 'sheet-main'), expected, argumentsText)
  for (const [argumentsText, code] of [
    ['TRUE,36', '#NUM!'], ['"FaLsE",36', '#NUM!'], ['"-1",16', '#NUM!'], ['"1.5",16', '#NUM!'],
    ['"9007199254740992",16', '#NUM!'], ['"0xFF",16', '#NUM!'], ['" 111",2', '#NUM!'], ['"1e3",10', '#NUM!'],
    ['"FF",1', '#NUM!'], ['"FF",37', '#NUM!'], ['"FF","invalid"', '#VALUE!'],
    ['"invalid",#REF!', '#REF!'], ['#DIV/0!,"invalid"', '#DIV/0!'], ['A1:B1,16', '#VALUE!'], ['"FF",A1:B1', '#VALUE!'],
  ]) assert.deepEqual(formula.evaluateFormula(`=DECIMAL(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, argumentsText)
  for (const argumentsText of ['', '"FF"', '"FF",16,#REF!']) {
    assert.deepEqual(formula.evaluateFormula(`=DECIMAL(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  for (const [name, example, expected] of [['ARABIC', '"MCMXII"', 1912], ['ROMAN', '499,0', 'CDXCIX'], ['BASE', '15,2,10', '0000001111'], ['DECIMAL', '"FF",16', 255]]) {
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学')
    assert.equal(entry.example, example)
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, example), model, 'sheet-main'), expected)
  }
})

test('CEILING and FLOOR variants retain signed steps, optional modes, fixed epsilon and canonical dotted names', () => {
  for (const [name, argumentsText, expected] of [
    ['CEILING', '2.5,1', 3], ['CEILING', '-2.5,2', -2], ['CEILING', '-2.5,-2', -4],
    ['FLOOR', '3.7,2', 2], ['FLOOR', '-2.5,2', -4], ['FLOOR', '-2.5,-2', -2],
    ['CEILING.MATH', '-5.5,2,-1', -6], ['CEILING.MATH', '-5.5,-2', -4],
    ['CEILING.MATH', '-5.5,-2,TRUE', -6], ['CEILING.MATH', '-5.5,2,0.0001', -6],
    ['FLOOR.MATH', '-5.5,2,-1', -4], ['FLOOR.MATH', '-5.5,-2', -6],
    ['FLOOR.MATH', '-5.5,-2,TRUE', -4], ['FLOOR.MATH', '-5.5,2,0.0001', -4],
    ['CEILING.PRECISE', '4.3,-2', 6], ['CEILING.PRECISE', '-4.3,-2', -4],
    ['FLOOR.PRECISE', '-3.2,-1', -4], ['FLOOR.PRECISE', '4.3,-2', 4],
    ['ISO.CEILING', '4.3', 5], ['ISO.CEILING', '-4.3', -4], ['ISO.CEILING', '-4.3,-2', -4],
  ]) assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), expected, `${name}(${argumentsText})`)
  for (const name of ['CEILING', 'CEILING.MATH', 'CEILING.PRECISE', 'ISO.CEILING', 'FLOOR', 'FLOOR.MATH', 'FLOOR.PRECISE']) {
    assert.ok(Math.abs(formula.evaluateFormula(`=${name}(0.3,0.1)`, model, 'sheet-main') - 0.3) < 1e-14, name)
  }
  for (const [name, example, expected] of [
    ['CEILING', '2.5,1', 3], ['CEILING.MATH', '-5.5,2,-1', -6], ['CEILING.PRECISE', '4.3,2', 6],
    ['FLOOR', '3.7,2', 2], ['FLOOR.MATH', '-5.5,2,-1', -4], ['FLOOR.PRECISE', '-3.2,-1', -4],
    ['ISO.CEILING', '-4.3,2', -4],
  ]) {
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学')
    assert.equal(entry.example, example)
    assert.equal(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, entry.example), model, 'sheet-main'), expected)
  }
})

test('CEILING and FLOOR retain scalar conversion, zero steps, original errors and variant argument counts', () => {
  for (const name of ['CEILING', 'CEILING.MATH', 'CEILING.PRECISE', 'ISO.CEILING', 'FLOOR', 'FLOOR.MATH', 'FLOOR.PRECISE']) {
    const legacy = name === 'CEILING' || name === 'FLOOR'
    for (const argument of ['A1', 'A1:A1', 'INDEX(A1:A1,1)']) {
      assert.equal(formula.evaluateFormula(`=${name}(${argument},TRUE)`, model, 'sheet-main'), 2)
    }
    assert.equal(formula.evaluateFormula(`=${name}("250%","$1")`, model, 'sheet-main'), name.startsWith('FLOOR') ? 2 : 3)
    for (const argumentsText of ['FALSE,TRUE', 'A9:A9,2', '0,0']) {
      assert.equal(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), 0)
    }
    const zeroStep = name === 'FLOOR' ? { kind: 'error', code: '#DIV/0!' } : 0
    assert.deepEqual(formula.evaluateFormula(`=${name}(2.15,A9:A9)`, model, 'sheet-main'), zeroStep)
    assert.deepEqual(formula.evaluateFormula(`=${name}("1e999",0)`, model, 'sheet-main'), zeroStep)
    for (const [argumentsText, code] of [
      ['"invalid",#REF!', name === 'FLOOR.MATH' ? '#VALUE!' : '#REF!'],
      ['#REF!,"invalid"', '#REF!'], ['"invalid",1', '#VALUE!'], ['"",1', '#VALUE!'],
      ['1,"invalid"', '#VALUE!'], ['1,1/0', '#DIV/0!'], ['A1:B1,1', '#VALUE!'],
    ]) assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code }, `${name}(${argumentsText})`)
    for (const argumentsText of ['', '#REF!,1,2,3', ...(legacy ? ['1', '#REF!'] : [])]) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
    }
    if (legacy) assert.deepEqual(formula.evaluateFormula(`=${name}(2.5,-1)`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
    if (!legacy) assert.equal(formula.evaluateFormula(`=${name}(-5.5)`, model, 'sheet-main'), name.startsWith('FLOOR') ? -6 : -5)
    if (!name.endsWith('.MATH')) assert.deepEqual(formula.evaluateFormula(`=${name}(#REF!,1,2)`, model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  }
  for (const name of ['CEILING.MATH', 'FLOOR.MATH']) {
    assert.deepEqual(formula.evaluateFormula(`=${name}(0,0,"invalid")`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
    assert.deepEqual(formula.evaluateFormula(`=${name}("invalid",1,#REF!)`, model, 'sheet-main'), { kind: 'error', code: name === 'FLOOR.MATH' ? '#VALUE!' : '#REF!' })
    assert.deepEqual(formula.evaluateFormula(`=${name}(1,"invalid",#REF!)`, model, 'sheet-main'), { kind: 'error', code: name === 'FLOOR.MATH' ? '#VALUE!' : '#REF!' })
  }
})

test('CEILING and FLOOR keep fixed negative zero, large-integer epsilon and nonfinite numeric results', () => {
  for (const name of ['CEILING', 'CEILING.MATH', 'CEILING.PRECISE', 'ISO.CEILING']) {
    assert.ok(Object.is(formula.evaluateFormula(`=${name}(-0.01,1)`, model, 'sheet-main'), -0), name)
  }
  assert.ok(Object.is(formula.evaluateFormula('=FLOOR.MATH(-0.01,1,1)', model, 'sheet-main'), -0))
  assert.equal(formula.evaluateFormula('=CEILING(9007199254740992,1)', model, 'sheet-main'), 9007199254740990)
  assert.equal(formula.evaluateFormula('=FLOOR(9007199254740992,1)', model, 'sheet-main'), 9007199254740994)
  const source = structuredClone(model)
  for (const [value, ceilingResult, floorResult] of [[NaN, NaN, NaN], [Infinity, NaN, Infinity], [-Infinity, -Infinity, NaN], [-0, 0, 0]]) {
    source.sheets[0].cells['0:0'].value = value
    for (const name of ['CEILING', 'CEILING.MATH', 'CEILING.PRECISE', 'ISO.CEILING', 'FLOOR', 'FLOOR.MATH', 'FLOOR.PRECISE']) {
      const result = formula.evaluateFormula(`=${name}(A1,1)`, source, 'sheet-main')
      assert.equal(result, name.startsWith('FLOOR') ? floorResult : ceilingResult, name)
    }
  }
})

test('trigonometric functions coerce scalar text and references while preserving errors and parameter contracts', () => {
  const source = structuredClone(model)
  source.sheets[0].cells['1:0'] = { ...source.sheets[0].cells['0:0'], id: '1:0', row: 1, col: 0, value: '#N/A', display: '#N/A', formula: null, dataType: 'error' }
  for (const [name, atOne, atTwo, atZero] of [
    ['SIN', 0.8414709848078965, 0.9092974268256817, 0],
    ['COS', 0.5403023058681398, -0.4161468365471424, 1],
    ['TAN', 1.5574077246549023, -2.185039863261519, 0],
    ['COT', 0.642092615934331, -0.45765755436028577, '#DIV/0!'],
    ['SEC', 1.8508157176809255, -2.402997961722381, 1],
    ['CSC', 1.188395105778121, 1.0997501702946164, '#DIV/0!'],
  ]) {
    for (const argument of ['1', 'TRUE', '"1"', '"$1"', '"100%"']) {
      assert.ok(Math.abs(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main') - atOne) < 1e-14, `${name}(${argument})`)
    }
    for (const argument of ['A1', 'A1:A1', 'INDEX(A1:A1,1)']) {
      assert.ok(Math.abs(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main') - atTwo) < 1e-14, `${name}(${argument})`)
    }
    for (const argument of ['0', 'FALSE', 'A9:A9']) {
      const expected = typeof atZero === 'number' ? atZero : { kind: 'error', code: atZero }
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), expected, `${name}(${argument})`)
    }
    for (const argumentsText of ['', '1,2', '1/0,2', '#REF!,2']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, `${name}(${argumentsText})`)
    }
    for (const argumentsText of ['""', '" "', '"invalid"', 'A1:B1']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argumentsText})`, model, 'sheet-main'), { kind: 'error', code: '#VALUE!' }, `${name}(${argumentsText})`)
    }
    assert.deepEqual(formula.evaluateFormula(`=${name}(NA())`, model, 'sheet-main'), { kind: 'error', code: '#N/A' }, name)
    assert.deepEqual(formula.evaluateFormula(`=${name}(A2:A2)`, source, 'sheet-main'), { kind: 'error', code: '#N/A' }, name)
    assert.deepEqual(formula.evaluateFormula(`=${name}(#REF!)`, model, 'sheet-main'), { kind: 'error', code: '#REF!' }, name)
    assert.deepEqual(formula.evaluateFormula(`=${name}("1e999")`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' }, name)
    const entry = functionCatalog.SHEET_FUNCTION_CATALOG.find(entry => entry.name === name)
    assert.equal(entry.category, '数学', name)
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.filter(entry => entry.name === name).length, 1, name)
    assert.ok(Math.abs(formula.evaluateFormula(functionCatalog.buildSheetFunctionFormula(name, entry.example), model, 'sheet-main') - atOne) < 1e-14, name)
  }
})

test('reciprocal trigonometric functions enforce the fixed magnitude limit without limiting SIN, COS or TAN', () => {
  for (const name of ['COT', 'SEC', 'CSC']) {
    for (const argument of ['134217728', '-134217728', '134217729', '"134217728"']) {
      assert.deepEqual(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main'), { kind: 'error', code: '#NUM!' }, `${name}(${argument})`)
    }
    for (const argument of ['134217727', '-134217727']) {
      assert.ok(Number.isFinite(formula.evaluateFormula(`=${name}(${argument})`, model, 'sheet-main')), `${name}(${argument})`)
    }
  }
  for (const name of ['SIN', 'COS', 'TAN']) {
    assert.ok(Number.isFinite(formula.evaluateFormula(`=${name}(134217728)`, model, 'sheet-main')), name)
  }
})

test('conditional count, sum and average functions handle blanks, criteria and range contracts', () => {
  const data = structuredClone(model)
  const sheet = data.sheets[0]
  const put = (row, col, value, formulaText) => {
    const cell = { ...sheet.cells['0:0'], id: `${row}:${col}`, row, col, value, formula: formulaText ?? null, display: formulaText ? '' : value == null ? '' : String(value) }
    sheet.cells[`${row}:${col}`] = cell
  }
  put(0, 0, 'East'); put(0, 1, 10); put(0, 2, 1)
  put(1, 0, 'West'); put(1, 1, 20); put(1, 2, 2)
  put(2, 0, null); put(2, 1, 30); put(2, 2, 3)
  put(3, 0, '', '=""'); put(3, 1, 40); put(3, 2, 4)
  put(4, 0, 'Alpha'); put(4, 1, 50); put(4, 2, 5)
  put(5, 0, 'alpha'); put(5, 1, 60); put(5, 2, 6)
  put(6, 0, 'a*b'); put(6, 1, 70); put(6, 2, 7)
  put(7, 0, '#N/A'); put(7, 1, 80); put(7, 2, 8)
  put(8, 0, 45292); put(8, 1, 90); put(8, 2, 9)

  assert.equal(formula.evaluateFormula('=COUNTBLANK(A1:A9)', data, 'sheet-main'), 2)
  assert.equal(formula.evaluateFormula('=COUNTBLANK(A3)', data, 'sheet-main'), 1)
  assert.deepEqual(formula.evaluateFormula('=COUNTBLANK(#REF!)', data, 'sheet-main'), { kind: 'error', code: '#REF!' })
  assert.equal(formula.evaluateFormula('=COUNTA(A1:A9)', data, 'sheet-main'), 8)
  assert.equal(formula.evaluateFormula('=COUNTA(A8:A8)', data, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=COUNTBLANK(A8:A8)', data, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=COUNTIF(B1:B9,">50")', data, 'sheet-main'), 4)
  assert.equal(formula.evaluateFormula('=COUNTIF(A1:A9,"a*")', data, 'sheet-main'), 3)
  assert.equal(formula.evaluateFormula('=COUNTIF(A1:A9,"a~*b")', data, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=COUNTIF(A1:A9,#N/A)', data, 'sheet-main'), 0)
  assert.equal(formula.evaluateFormula('=COUNTIF(A1:A9,"<>")', data, 'sheet-main'), 8)
  assert.equal(formula.evaluateFormula('=COUNTIFS(A1:A9,"E*",B1:B9,">5")', data, 'sheet-main'), 1)
  assert.equal(formula.evaluateFormula('=COUNTIF(A9:A9,">="&DATE(2024,1,1))', data, 'sheet-main'), 1)

  assert.equal(formula.evaluateFormula('=SUMIF(A1:A9,"E*",B1:B9)', data, 'sheet-main'), 10)
  assert.equal(formula.evaluateFormula('=SUMIF(B1:B9,">50",C1:C2)', data, 'sheet-main'), 30)
  assert.equal(formula.evaluateFormula('=SUMIFS(B1:B9,A1:A9,"W*",C1:C9,">1")', data, 'sheet-main'), 20)
  assert.equal(formula.evaluateFormula('=AVERAGEIF(A1:A9,"a*",B1:B9)', data, 'sheet-main'), 60)
  assert.equal(formula.evaluateFormula('=AVERAGEIFS(B1:B9,A1:A9,"E*")', data, 'sheet-main'), 10)
  assert.deepEqual(formula.evaluateFormula('=AVERAGEIF(A1:A9,"Missing",B1:B9)', data, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.deepEqual(formula.evaluateFormula('=COUNTIFS(A1:A9,"E*",B1:B8,"x")', data, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=SUMIFS(B1:B9,A1:A9,"E*",C1:C8,">0")', data, 'sheet-main'), { kind: 'error', code: '#VALUE!' })

  for (const name of ['COUNTBLANK', 'COUNTA', 'COUNTIF', 'COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
    assert.equal(functionCatalog.SHEET_FUNCTION_CATALOG.filter(entry => entry.name === name).length, 1, `${name} has a single catalog entry`)
  }
})

test('financial period, interest, principal, schedule and internal-return functions honor cash-flow conventions', () => {
  assert.ok(Math.abs(formula.evaluateFormula('=NPER(0.12/12,-100,-1000,10000,0)', model, 'sheet-main') - 60.0821228538) < 1e-9)
  assert.ok(Math.abs(formula.evaluateFormula('=NPER(0.12/12,-100,-1000,10000,1)', model, 'sheet-main') - 59.6738656743) < 1e-9)
  assert.deepEqual(formula.evaluateFormula('=NPER(0,0,100)', model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })

  assert.ok(Math.abs(formula.evaluateFormula('=RATE(48,-200,100000,0,0,10%)', model, 'sheet-main') - (-0.0726317700766)) < 1e-10)
  assert.deepEqual(formula.evaluateFormula('=RATE(4,0,100,0)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })

  assert.ok(Math.abs(formula.evaluateFormula('=IPMT(0.1/12,1,36,80000)', model, 'sheet-main') - (-666.666666666667)) < 1e-9)
  assert.equal(formula.evaluateFormula('=IPMT(0.1/12,1,36,80000,0,1)', model, 'sheet-main'), 0)
  assert.ok(Math.abs(formula.evaluateFormula('=PPMT(0.1/12,1,36,80000)', model, 'sheet-main') - (-1914.70830884034)) < 1e-9)
  assert.deepEqual(formula.evaluateFormula('=PPMT(0.1/12,0,36,80000)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })

  const scheduleData = structuredClone(model)
  for (const [column, value] of [0.09, 0.11, 0.1].entries()) {
    const cell = { ...scheduleData.sheets[0].cells['0:0'], id: `0:${column}`, row: 0, col: column, value, formula: null, display: String(value) }
    scheduleData.sheets[0].cells[`0:${column}`] = cell
  }
  assert.ok(Math.abs(formula.evaluateFormula('=FVSCHEDULE(1,A1:C1)', scheduleData, 'sheet-main') - 1.33089) < 1e-12)
  assert.ok(Math.abs(formula.evaluateFormula('=NPV(10%,-100,40,40,40,40)', model, 'sheet-main') - 24.35874350361063) < 1e-10)

  const cashflows = structuredClone(model)
  const sheet = cashflows.sheets[0]
  for (const [row, value] of [-700000, 120000, 150000, 180000, 210000, 260000].entries()) {
    const cell = { ...sheet.cells['0:0'], id: `${row}:0`, row, col: 0, value, formula: null, display: String(value) }
    sheet.cells[`${row}:0`] = cell
  }
  assert.ok(Math.abs(formula.evaluateFormula('=IRR(A1:A6)', cashflows, 'sheet-main') - 0.0866309480365) < 1e-10)
  assert.deepEqual(formula.evaluateFormula('=IRR(B1:B2)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.deepEqual(formula.evaluateFormula('=IRR(A1:A6,-1)', cashflows, 'sheet-main'), { kind: 'error', code: '#NUM!' })

  for (const name of ['NPER', 'RATE', 'IPMT', 'PPMT', 'IRR', 'FVSCHEDULE']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
})

test('depreciation and fractional-dollar functions follow fixed-reference financial contracts', () => {
  assert.equal(formula.evaluateFormula('=DOLLARDE(1.02,16)', model, 'sheet-main'), 1.125)
  assert.equal(formula.evaluateFormula('=DOLLARDE("1.02","16")', model, 'sheet-main'), 1.125)
  assert.equal(formula.evaluateFormula('=DOLLARFR(1.125,16)', model, 'sheet-main'), 1.02)
  assert.deepEqual(formula.evaluateFormula('=DOLLARFR(TRUE,16)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.deepEqual(formula.evaluateFormula('=DOLLARDE(1.02,0.5)', model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.deepEqual(formula.evaluateFormula('=DOLLARFR(1.02,-1)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })

  assert.ok(Math.abs(formula.evaluateFormula('=DB(10000000,1000000,6,1,7)', model, 'sheet-main') - 1860833.33333333) < 1e-8)
  assert.ok(Math.abs(formula.evaluateFormula('=DB(10000000,1000000,6,6,7)', model, 'sheet-main') - 558417.567360284) < 1e-8)
  assert.ok(Math.abs(formula.evaluateFormula('=DDB(24000,3000,6,1)', model, 'sheet-main') - 8000) < 1e-9)
  assert.ok(Math.abs(formula.evaluateFormula('=DDB(24000,3000,6,1,1)', model, 'sheet-main') - 4000) < 1e-9)
  assert.deepEqual(formula.evaluateFormula('=DDB(100,0,5,1,0)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })

  assert.equal(formula.evaluateFormula('=SLN(300000,75000,10)', model, 'sheet-main'), 22500)
  assert.deepEqual(formula.evaluateFormula('=SLN(300000,75000,0)', model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.equal(formula.evaluateFormula('=SYD(300000,75000,10,10)', model, 'sheet-main'), 4090.909090909091)
  assert.deepEqual(formula.evaluateFormula('=SYD(300000,75000,10,11)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })
  assert.ok(Math.abs(formula.evaluateFormula('=RRI(96,10000,11000)', model, 'sheet-main') - 0.000993307376291) < 1e-14)
  assert.ok(Math.abs(formula.evaluateFormula('=RRI(96,10000,11000,"ignored")', model, 'sheet-main') - 0.000993307376291) < 1e-14)
  assert.deepEqual(formula.evaluateFormula('=RRI(96,0,11000)', model, 'sheet-main'), { kind: 'error', code: '#NUM!' })

  for (const name of ['DB', 'DDB', 'DOLLARDE', 'DOLLARFR', 'RRI', 'SLN', 'SYD']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
})

