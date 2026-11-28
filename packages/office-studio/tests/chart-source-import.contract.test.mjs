import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createBlankSheet, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { createChartDefinition } from '../src/features/charts/chart-engine.ts'
import { captureChartObjectClipboard, chartObjectPasteDefinition } from '../src/features/charts/chart-object-clipboard.ts'
import { planChartSourceImport, planChartSourceImportAsync, applyChartSourceImport, assertChartSourceImportSource } from '../src/features/charts/chart-source-import.ts'
import { evaluateFormula, recalculateWorkbookFormulas } from '../src/features/formula/formula-engine.ts'
import { readFormulaDefinedNames } from '../src/features/formula/formula-defined-names.ts'
import { readFormulaImportSource } from '../src/features/formula/formula-import-context.ts'
import { planWorksheetCopy } from '../src/features/sheet-workbench/sheet-worksheet-copy.ts'
import { getPivotChartReport } from '../src/features/pivot/services/pivot-chart-engine.ts'
import { calculatePivotReport, refreshPivotTable } from '../src/features/pivot/services/pivot-engine.ts'
import { pivotChartEngineWorkbook } from './fixtures/pivot-chart-engine-fixture.mjs'
import { chartResourceFixture } from './fixtures/chart-resource-platform-fixture.mjs'

// Drafts only. No tests, servers, browsers, Office applications or WASM were run.
const range = (startRow, endRow, startCol = 0, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
const value = (workbook, sheet, formula) => {
  const cell = Object.values(sheet.cells).find(cell => cell.formula === formula)
  return evaluateFormula(formula, workbook, sheet.id, { ...(cell ? { formulaCell: { row: cell.row, column: cell.col } } : {}) })
}
const protection = () => ({ id: 'deny', permissionId: 'deny-permission', unitType: 'worksheet', ownerUserId: 'other',
  allowedUserIds: [], editState: 'onlyMe', viewState: 'noOneElseCanView' })
function targetWorkbook(id = 'target') {
  const target = createBlankWorkbook(); target.id = id; target.sheets[0].id = `${id}-sheet`; target.sheets[0].name = 'Data'; target.activeSheetId = target.sheets[0].id
  target.sheets[0].cells['0:3'] = createCell(0, 3, 'Target content')
  return target
}
function fixture() {
  const source = createBlankWorkbook(), first = source.sheets[0], other = createBlankSheet('Other')
  source.id = 'source'; first.id = 'data'; first.name = 'Data'; source.activeSheetId = first.id; other.id = 'other'; source.sheets.push(other)
  first.cells = { '0:0': createCell(0, 0, 'Label'), '0:1': createCell(0, 1, 'Value'),
    '1:0': createCell(1, 0, 'One'), '1:1': createCell(1, 1, 7), '2:0': createCell(2, 0, 'Two'), '2:1': createCell(2, 1, 4),
    '900000:8': createCell(900000, 8, 'Unrelated sparse content') }
  const chart = createChartDefinition(source, first.id, { type: 'line', source: { sheetId: first.id, orientation: 'columns', range: range(0, 2, 0, 1) },
    categoryField: 0, valueFields: [1], position: { x: 40, y: 50 }, size: { width: 400, height: 280 }, title: 'Source chart' }, 'source-chart')
  source.charts = [chart]
  const target = targetWorkbook(), capture = () => captureChartObjectClipboard(source, { unitId: source.id, sheetId: chart.sheetId, chartId: chart.id, snapshot: JSON.stringify(chart) })
  const options = { userId: 'local-user', targetSheetId: target.activeSheetId }
  return { source, target, first, other, chart, capture, options }
}
const importedSheet = (plan, originalId) => plan.sheets.find(sheet => sheet.metadata.chartDependencyImport.sourceSheetId === originalId)
function publish(source, target, chart, plan, clipboard) {
  const created = chartObjectPasteDefinition(target, clipboard, { unitId: target.id, sheetId: target.activeSheetId, newChartId: `${chart.id}-copy`, clipboardId: clipboard.id },
    { sourceWorkbook: source, userId: 'local-user', sourceImport: plan })
  target.charts ??= []; target.charts.push(created); return created
}

test('a chart imports only its sparse dependency closure, preserving target content and detached author/model ownership', () => {
  const { source, target, first, other, chart, capture, options } = fixture(), before = JSON.stringify(source), targetBefore = cloneModel(target.sheets[0])
  const copied = capture(), plan = planChartSourceImport(source, target, copied, options)
  assert.equal(JSON.stringify(source), before); assert.equal(plan.copiedCells, 6); assert.equal(plan.sheets.length, 1)
  assert.equal(importedSheet(plan, other.id), undefined); assert.equal(importedSheet(plan, first.id).cells['900000:8'], undefined)
  assert.equal(importedSheet(plan, first.id).visible, false); assert.equal(importedSheet(plan, first.id).metadata.hiddenState, 'hidden')
  const created = publish(source, target, chart, plan, copied)
  assert.deepEqual(target.sheets[0], targetBefore); assert.equal(created.source.sheetId, importedSheet(plan, first.id).id)
  first.cells['1:1'].value = 999; assert.equal(target.sheets.find(sheet => sheet.id === created.source.sheetId).cells['1:1'].value, 7)
  created.title = 'Detached'; assert.equal(chart.title, 'Source chart')
})

test('global source names become hidden local names and cannot bind or repair unrelated target names', () => {
  const { source, target, first, other, chart, capture, options } = fixture()
  source.metadata = { definedNames: [{ name: 'Rate', refersTo: '=Other!A1' }, { name: 'OnlySource', formulaOrRefString: '=2' }] }
  other.cells['0:0'] = createCell(0, 0, 6); first.cells['1:1'].formula = '=Rate+OnlySource'
  target.metadata = { definedNames: [{ name: 'Rate', refersTo: '=99' }] }; target.sheets[0].cells['1:3'] = { ...createCell(1, 3, null), formula: '=OnlySource' }
  const copied = capture(), plan = planChartSourceImport(source, target, copied, options), copiedSource = importedSheet(plan, first.id)
  publish(source, target, chart, plan, copied)
  assert.equal(value(target, copiedSource, copiedSource.cells['1:1'].formula), 8)
  assert.equal(value(target, target.sheets[0], '=Rate'), 99); assert.equal(value(target, target.sheets[0], '=OnlySource').code, '#NAME?')
  assert.equal(readFormulaDefinedNames(target).filter(name => name.owner === undefined).length, 1)
  assert.ok(readFormulaDefinedNames(target).filter(name => name.owner !== undefined).every(name => plan.sheets.some(sheet => sheet.id === name.owner)))
})

test('identical captions in distinct local scopes retain their own expressions through cross-workbook import', () => {
  const { source, target, first, other, chart, capture, options } = fixture()
  source.metadata = { definedNames: [{ name: 'Rate', formulaOrRefString: '=2', localSheetId: first.id },
    { name: 'Rate', refersTo: '=7', localSheetId: 1 }] }
  first.cells['1:1'].formula = '=Rate+Other!Rate'
  const copied = capture(), plan = planChartSourceImport(source, target, copied, options)
  publish(source, target, chart, plan, copied)
  const copiedSource = importedSheet(plan, first.id), names = readFormulaDefinedNames(target).filter(name => name.name === 'Rate')
  assert.equal(names.length, 2); assert.notEqual(names[0].owner, names[1].owner)
  assert.equal(value(target, copiedSource, copiedSource.cells['1:1'].formula), 9)
})

test('structured tables and escaped column selectors import their full data with fresh table identities', () => {
  const { source, target, first, chart, capture, options } = fixture()
  source.tables = [{ id: 'sales', name: 'Sales', sheetId: first.id, range: range(0, 2, 0, 1), showHeaderRow: true, showTotalsRow: false,
    styleName: 'TableStyleMedium2', columns: [{ id: 'label', name: 'Label' }, { id: 'value', name: 'X]Value' }] }]
  first.cells['0:1'].value = 'X]Value'; first.cells['1:2'] = { ...createCell(1, 2, 11), formula: "=SUM(Sales[X']Value])" }
  first.cells['0:2'] = createCell(0, 2, 'Sum'); chart.source.range.endCol = 2; chart.valueFields = [2]
  target.tables = [{ ...cloneModel(source.tables[0]), id: 'existing-table', sheetId: target.activeSheetId }]
  const copied = capture(), plan = planChartSourceImport(source, target, copied, options), imported = importedSheet(plan, first.id), table = plan.tables[0]
  assert.notEqual(table.id, 'sales'); assert.notEqual(table.name, 'Sales'); assert.equal(table.columns[1].name, 'X]Value')
  publish(source, target, chart, plan, copied)
  assert.equal(value(target, imported, imported.cells['1:2'].formula), 11); assert.equal(target.tables[0].name, 'Sales')
})

test('static 3D references retain all intermediate worksheet descriptors and their source ordering', () => {
  const { source, target, first, other, capture, options } = fixture(), last = createBlankSheet('Last')
  last.id = 'last'; source.sheets.push(last); first.name = 'First'
  first.cells['0:0'] = createCell(0, 0, 1); other.cells['0:0'] = createCell(0, 0, 2); last.cells['0:0'] = createCell(0, 0, 3)
  first.cells['1:1'].formula = '=SUM(First:Last!A1)'
  const plan = planChartSourceImport(source, target, capture(), options)
  assert.deepEqual(plan.sheets.map(sheet => sheet.metadata.chartDependencyImport.sourceSheetId), [first.id, other.id, last.id])
  assert.equal(importedSheet(plan, other.id).cells['0:0'].value, 2); assert.equal(importedSheet(plan, last.id).cells['0:0'].value, 3)
  // This proves the import closure, not full 3D formula evaluation or Office acceptance.
})

test.each([false, true])('long name chains and cycle=%s are collected without recursive stack growth and preserve evaluation identity', cycle => {
  const { source, target, first, chart, capture, options } = fixture()
  source.metadata = { definedNames: Array.from({ length: 2048 }, (_, index) => ({ name: `Chain_${index}`,
    formulaOrRefString: index === 2047 ? cycle ? '=Chain_0' : '=4' : `=Chain_${index + 1}` })) }
  first.cells['1:1'].formula = '=Chain_0'
  const copied = capture(), plan = planChartSourceImport(source, target, copied, options), imported = importedSheet(plan, first.id)
  publish(source, target, chart, plan, copied); assert.equal(readFormulaDefinedNames(target).length, 2048)
  const actual = value(target, imported, imported.cells['1:1'].formula)
  assert.equal(cycle ? actual.code : actual, cycle ? '#CIRCULAR!' : 4)
})

test('whole-axis dependencies stay sparse and include newly addressed far cells with complete empty-range permissions', () => {
  const { source, target, first, chart, capture, options } = fixture()
  first.cells['1:1'].formula = '=SUM(I:I)'; first.cells['900000:8'] = createCell(900000, 8, 13)
  const copied = capture(), plan = planChartSourceImport(source, target, copied, options), imported = importedSheet(plan, first.id)
  assert.equal(Object.keys(imported.cells).length, 7); publish(source, target, chart, plan, copied)
  assert.equal(value(target, imported, imported.cells['1:1'].formula), 13)
  first.rangeProtectionRules = [{ ...protection(), unitType: 'range', ranges: [range(100, 200, 8)] }]
  assert.throws(() => planChartSourceImport(source, targetWorkbook('blocked-target'), copied, { ...options, targetSheetId: 'blocked-target-sheet' }), /查看|权限/)
})

test('formula and name dependency permissions cover off-chart sources and intersecting merged geometry', () => {
  const { source, target, first, other, capture, options } = fixture()
  first.cells['1:1'].formula = '=Other!A1'; other.cells['0:0'] = createCell(0, 0, 3); other.mergedRanges = [range(0, 0, 0, 1)]
  other.rangeProtectionRules = [{ ...protection(), unitType: 'range', ranges: [range(0, 0, 1)] }]
  const before = JSON.stringify(target)
  assert.throws(() => planChartSourceImport(source, target, capture(), options), /查看|权限/); assert.equal(JSON.stringify(target), before)
  other.rangeProtectionRules = []; other.worksheetProtection = protection()
  assert.throws(() => planChartSourceImport(source, target, capture(), options), /查看|权限/)
})

test('INDIRECT imports potential sparse local dependencies and remaps source aliases without rebinding target names', () => {
  const { source, target, first, other, chart, capture, options } = fixture()
  first.cells['1:1'].formula = '=INDIRECT("Other!A1")'; other.cells['0:0'] = createCell(0, 0, 15)
  const copied = capture(), plan = planChartSourceImport(source, target, copied, options), imported = importedSheet(plan, first.id)
  assert.ok(plan.diagnostics.includes('dynamic-formula-reference')); publish(source, target, chart, plan, copied)
  assert.equal(value(target, imported, imported.cells['1:1'].formula), 15)
  assert.equal(imported.cells['1:1'].formula, '=INDIRECT("Other!A1")')
})

test('formula length, names and memory budgets reject candidates before target mutation', () => {
  const { source, target, first, capture, options } = fixture(), before = JSON.stringify(target)
  assert.throws(() => planChartSourceImport(source, target, capture(), { ...options, memoryBudgetBytes: 128 }), /预算|内存/)
  first.cells['1:1'].formula = `=SUM(${"A1+".repeat(11000)}0)`
  assert.throws(() => planChartSourceImport(source, target, capture(), options), /公式表达式超出预算/)
  first.cells['1:1'].formula = '=Many_0'; source.metadata = { definedNames: Array.from({ length: 8193 }, (_, index) => ({ name: `Many_${index}`, refersTo: '=1' })) }
  assert.throws(() => planChartSourceImport(source, target, capture(), options), /定义名称数量|预算/)
  assert.equal(JSON.stringify(target), before)
})

test('cooperative preparation accepts ordinary timer cancellation and releases candidates without partial publication', async () => {
  const { source, target, first, chart, capture, options } = fixture(), controller = new AbortController(), before = JSON.stringify(target)
  chart.source.range.endRow = 600
  for (let row = 3; row <= 600; row++) first.cells[`${row}:1`] = createCell(row, 1, row)
  const pending = planChartSourceImportAsync(source, target, capture(), { ...options, signal: controller.signal })
  setTimeout(() => controller.abort(), 0)
  await assert.rejects(pending, error => error.code === 'COMPUTE_CANCELLED'); assert.equal(JSON.stringify(target), before)
})

test.each(['cell', 'new-cell', 'name', 'style'])('direct %s mutations during a cooperative import invalidate its exact source state', async kind => {
  const { source, target, first, chart, capture, options } = fixture(), before = JSON.stringify(target)
  chart.source.range.endRow = 600
  for (let row = 3; row <= 600; row++) first.cells[`${row}:1`] = createCell(row, 1, row)
  const pending = planChartSourceImportAsync(source, target, capture(), options)
  setTimeout(() => {
    if (kind === 'cell') first.cells['1:1'].value = 100
    if (kind === 'new-cell') first.cells['550:0'] = createCell(550, 0, 'New sparse cell')
    if (kind === 'name') source.metadata = { definedNames: [{ name: 'Added', refersTo: '=1' }] }
    if (kind === 'style') first.defaultStyle = { bold: true }
  }, 0)
  await assert.rejects(pending, /更改|重新粘贴/); assert.equal(JSON.stringify(target), before)
})

test('final source checks detect additions, deletions, name resources and raw image replacement without a revision increment', () => {
  const { source, target, first, chart, capture, options } = fixture(), image = chartResourceFixture()
  chart.type = 'wordCloud'; chart.options = { wordCloud: { maskImageId: image.id } }; source.metadata = { chartImageResources: { [image.id]: image } }
  chart.source.range.endRow = 3
  const copied = capture(), plan = planChartSourceImport(source, target, copied, options), original = cloneModel(source)
  for (const mutate of [() => { first.cells['1:1'].value = 100 }, () => { first.cells['3:0'] = createCell(3, 0, 'New cell') },
    () => { delete first.cells['1:0'] }, () => { source.metadata.definedNames = [{ name: 'New', refersTo: '=1' }] },
    () => { source.metadata.chartImageResources[image.id].name = 'Replaced mask' }]) {
    source.sheets = cloneModel(original.sheets); source.metadata = cloneModel(original.metadata)
    Object.assign(first, source.sheets[0]); source.sheets[0] = first
    mutate(); assert.throws(() => assertChartSourceImportSource(source, plan, options.userId), /更改|删除|图片|重新粘贴/)
  }
})

test.each(['sheet-id', 'sheet-name', 'table-id', 'style', 'binding'])('an apply-time %s conflict rejects the whole staged candidate', kind => {
  const { source, target, first, capture, options } = fixture()
  source.styles = { retained: { bold: true } }; first.cells['1:1'].authoredStyleReference = 'retained'
  source.tables = [{ id: 'sales', name: 'Sales', sheetId: first.id, range: range(0, 2, 0, 1), showHeaderRow: true, showTotalsRow: false,
    styleName: 'TableStyleMedium2', columns: [{ id: 'label', name: 'Label' }, { id: 'value', name: 'Value' }] }]
  const plan = planChartSourceImport(source, target, capture(), options)
  if (kind === 'sheet-id') target.sheets.push({ ...createBlankSheet('Occupied'), id: plan.sheets[0].id })
  if (kind === 'sheet-name') target.sheets.push({ ...createBlankSheet(plan.sheets[0].name), id: 'occupied' })
  if (kind === 'table-id') target.tables = [{ ...cloneModel(plan.tables[0]), sheetId: target.activeSheetId }]
  if (kind === 'style') target.styles = { retained: { italic: true } }
  const candidate = kind === 'binding' ? { ...plan, chart: { ...plan.chart, source: { ...plan.chart.source, sheetId: 'missing' } } } : plan
  const before = JSON.stringify(target)
  assert.throws(() => applyChartSourceImport(target, candidate), /占用|更改|丢失|绑定|目标/); assert.equal(JSON.stringify(target), before)
})

test.each(['owned', 'linked'])('a %s pivot chart receives detached cache, canonical fields and editable report ownership', ownership => {
  const source = pivotChartEngineWorkbook({ ownership }), target = targetWorkbook(), chart = source.charts[0]
  const copied = captureChartObjectClipboard(source, { unitId: source.id, sheetId: chart.sheetId, chartId: chart.id, snapshot: JSON.stringify(chart) })
  const before = calculatePivotReport(source, getPivotChartReport(source, chart)).values
  const plan = planChartSourceImport(source, target, copied, { userId: 'local-user', targetSheetId: target.activeSheetId }), created = publish(source, target, chart, plan, copied)
  const report = getPivotChartReport(target, created)
  assert.notEqual(report.id, getPivotChartReport(source, chart).id); assert.notEqual(report.pivotCacheId, 'engine-cache')
  assert.equal(report.source.unitId, target.id); assert.equal(report.anchor.unitId, target.id)
  assert.deepEqual(calculatePivotReport(target, report).values, before)
  if (ownership === 'owned') { assert.equal(target.pivotTables?.length ?? 0, 0); assert.equal(report.anchor.subUnitId, target.activeSheetId) }
  else { assert.equal(target.pivotTables.length, 1); refreshPivotTable(target, report); assert.deepEqual(calculatePivotReport(target, report).values, before) }
  source.sheets[0].cells['3:4'].value = 999; assert.deepEqual(calculatePivotReport(target, report).values, before)
})

test('calendar-typed cells, dependent formula results and numeric quantities survive import and return to the source calendar', () => {
  const { source, target, first, chart, capture, options } = fixture()
  target.date1904 = true; first.cells['1:0'] = { ...createCell(1, 0, 60.5), style: { ...createCell(1, 0, 0).style, numberFormat: 'yyyy-mm-dd hh:mm' }, dataType: 'date' }
  first.cells['1:1'].formula = '=DAY(A2)'; first.cells['2:1'].formula = '=DATE(1900,2,29)'
  first.cells['2:1'].style.numberFormat = 'yyyy-mm-dd'; first.cells['2:1'].cachedDataType = 'date'
  const copied = capture(), plan = planChartSourceImport(source, target, copied, options), imported = importedSheet(plan, first.id)
  publish(source, target, chart, plan, copied)
  assert.equal(imported.cells['1:0'].value, -1401.5); assert.equal(imported.cells['1:0'].metadata.chartDependencyDateSource.value, 60.5)
  assert.equal(value(target, imported, imported.cells['1:1'].formula), 29)
  const calculated = recalculateWorkbookFormulas(target).workbook
  assert.equal(calculated.sheets.find(sheet => sheet.id === imported.id).cells['2:1'].value, -1402)
  assert.equal(calculated.sheets.find(sheet => sheet.id === imported.id).cells['2:1'].metadata.chartDependencyDateSource.value, 60)
  const restored = targetWorkbook('restored'); restored.date1904 = false
  const currentChart = calculated.charts[0], ticket = captureChartObjectClipboard(calculated,
    { unitId: calculated.id, sheetId: currentChart.sheetId, chartId: currentChart.id, snapshot: JSON.stringify(currentChart) })
  const returnPlan = planChartSourceImport(calculated, restored, ticket, { userId: 'local-user', targetSheetId: restored.activeSheetId })
  assert.equal(importedSheet(returnPlan, imported.id).cells['1:0'].value, 60.5)
})

test('copying an imported formula to a different sheet keeps its alias owner in a subsequent cross-workbook closure', () => {
  const { source, target, first, other, chart, capture, options } = fixture()
  first.cells['1:1'].formula = '=INDIRECT("Other!A1")'; other.cells['0:0'] = createCell(0, 0, 17)
  const ticket = capture(), initial = planChartSourceImport(source, target, ticket, options), imported = importedSheet(initial, first.id)
  publish(source, target, chart, initial, ticket)
  const copied = planWorksheetCopy(target, imported.id, { userId: 'local-user', targetSheetId: 'formula-copy', targetSheetName: 'Copied formulas' }).workbook
  const copiedSheet = copied.sheets.find(sheet => sheet.id === 'formula-copy'), cell = copiedSheet.cells['1:1']
  // Model/API cell-copy paths preserve the source owner when the receiving sheet
  // does not duplicate the source alias catalog.
  cell.metadata.chartDependencyFormulaSource.aliasSheetId = imported.id
  const nextChart = createChartDefinition(copied, copiedSheet.id, { ...cloneModel(chart), source: { ...chart.source, sheetId: copiedSheet.id } }, 'copied-formula-chart')
  copied.charts.push(nextChart)
  const nextTicket = captureChartObjectClipboard(copied, { unitId: copied.id, sheetId: copiedSheet.id, chartId: nextChart.id, snapshot: JSON.stringify(nextChart) })
  const third = targetWorkbook('third'), plan = planChartSourceImport(copied, third, nextTicket, { userId: 'local-user', targetSheetId: third.activeSheetId })
  const formulaSheet = importedSheet(plan, copiedSheet.id), owner = importedSheet(plan, imported.id)
  assert.ok(owner); assert.equal(readFormulaImportSource(formulaSheet.cells['1:1']).aliasSheetId, owner.id)
  publish(copied, third, nextChart, plan, nextTicket)
  assert.equal(value(third, formulaSheet, formulaSheet.cells['1:1'].formula), 17)
})
