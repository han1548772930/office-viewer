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
  assert.deepEqual(evaluate('=TODAY(1)', data, clock), { kind: 'error', code: '#VALUE!' })

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
  assert.deepEqual(formula.evaluateFormula('=TRUE(1)', model, 'sheet-main'), { kind: 'error', code: '#VALUE!' })
  assert.equal(formula.evaluateFormula('=SWITCH("B","a","first","b","second","fallback")', model, 'sheet-main'), 'second')
  assert.equal(formula.evaluateFormula('=SWITCH("B","a","first","fallback")', model, 'sheet-main'), 'fallback')
  assert.deepEqual(formula.evaluateFormula('=SWITCH("B","a","first")', model, 'sheet-main'), { kind: 'error', code: '#N/A' })
  assert.equal(formula.evaluateFormula('=SWITCH(1,1,"selected",1/0)', model, 'sheet-main'), 'selected')
  assert.deepEqual(formula.evaluateFormula('=SWITCH(1/0,1,"selected")', model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.deepEqual(formula.evaluateFormula('=SWITCH(2,1,"first",1/0,"fallback")', model, 'sheet-main'), { kind: 'error', code: '#DIV/0!' })
  assert.equal(functionCatalog.buildSheetFunctionFormula('TRUE', ''), '=TRUE()')
  assert.equal(functionCatalog.buildSheetFunctionFormula('FALSE', ''), '=FALSE()')

  for (const name of ['TRUE', 'FALSE', 'SWITCH']) {
    assert.ok(functionCatalog.SHEET_FUNCTION_CATALOG.some(entry => entry.name === name), `${name} is in the insert-function catalog`)
  }
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

