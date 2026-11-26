import { pivotShowDataAsWorkbook } from './pivot-show-data-as-fixture.mjs'
import { calculatePivotReport } from '../../src/features/pivot/services/pivot-engine.ts'
import { applyPivotChartEdit, getPivotChartReport } from '../../src/features/pivot/services/pivot-chart-engine.ts'

export const showAsUIIds = { unit: 'show-data-as-unit', data: 'show-data', host: 'show-host', other: 'show-dashboard',
  pivot: 'show-table', chart: 'show-chart-a', amount: 'amount-value', product: 'product-label', region: 'region-label', year: 'year-label', quarter: 'quarter-label' }
export const showAsUIRange = { startRow: 1, endRow: 9, startCol: 2, endCol: 6 }
export const showAsUIRule = (type, baseFieldId = '', baseItem = '', baseItemType) => ({ type, baseFieldId, baseItem, ...(baseItemType ? { baseItemType } : {}) })
export const showAsUICases = [
  [showAsUIRule('normal'), '无计算', 10], [showAsUIRule('percentOfGrandTotal'), '总计的百分比', .05],
  [showAsUIRule('percentOfRowTotal'), '行汇总的百分比', 1 / 3], [showAsUIRule('percentOfColumnTotal'), '列汇总的百分比', 1 / 7],
  [showAsUIRule('percentOfParentRowTotal'), '父行汇总的百分比', 1 / 4], [showAsUIRule('percentOfParentColumnTotal'), '父列汇总的百分比', 1 / 3],
  [showAsUIRule('index'), '指数', 20 / 21], [showAsUIRule('percentOfParent', 'region-label'), '指定父项汇总的百分比', 1 / 4],
  [showAsUIRule('percentOf', 'product-label', 'B'), '指定项的百分比', 1 / 3],
  [showAsUIRule('differenceFrom', 'product-label', 'B'), '与指定项的差值', -20],
  [showAsUIRule('percentDifferenceFrom', 'product-label', 'B'), '与指定项差值的百分比', -2 / 3],
  [showAsUIRule('runningTotal', 'product-label'), '按字段累计', 10],
  [showAsUIRule('percentRunningTotal', 'product-label'), '按字段累计百分比', 1 / 4],
  [showAsUIRule('rankAscending', 'product-label'), '升序排名', 1], [showAsUIRule('rankDescending', 'product-label'), '降序排名', 2],
]

export function showAsUIWorkbook({ ownership = 'linked', grouping = false, repeated = false, blank = false, compatibility = false,
  initialRule, hiddenSource = false, protectedOutput = false } = {}) {
  const workbook = pivotShowDataAsWorkbook({ ownership })
  const report = getPivotChartReport(workbook, workbook.charts[0])
  if (grouping) report.fields.find(field => field.id === showAsUIIds.product).manualGroups = [{ name: 'Mix', members: ['A', 'B'] }]
  if (repeated) report.fields.push({ ...structuredClone(report.fields.find(field => field.id === showAsUIIds.amount)), id: 'amount-second', position: 1, displayName: 'Other amount' })
  if (blank) workbook.sheets.find(sheet => sheet.id === showAsUIIds.data).cells['2:3'].value = null
  if (compatibility) report.fields.find(field => field.id === showAsUIIds.amount).showDataAs = showAsUIRule('percentOfParentGrandTotal')
  if (initialRule) applyPivotChartEdit(workbook, workbook.charts[0], { type: 'setShowDataAs', fieldId: showAsUIIds.amount, showDataAs: initialRule })
  const protect = (sheet, id, range, hidden) => { sheet.rangeProtectionRules = [{ id, permissionId: id, description: '', ownerUserId: 'owner',
    allowedUserIds: [], editState: 'onlyMe', viewState: hidden ? 'noOneElseCanView' : 'othersCanView', ranges: [structuredClone(range)] }] }
  if (hiddenSource) protect(workbook.sheets.find(sheet => sheet.id === showAsUIIds.data), 'show-as-hidden-source', showAsUIRange, true)
  if (protectedOutput && ownership !== 'owned') protect(workbook.sheets.find(sheet => sheet.id === showAsUIIds.host), 'show-as-output', report.outputRange, false)
  return workbook
}

export function showAsUIValue(workbook, chartId = showAsUIIds.chart, labels = ['North', 'A'], columns = [2024, 'Q1']) {
  const chart = workbook.charts.find(chart => chart.id === chartId), report = getPivotChartReport(workbook, chart)
  const matrix = calculatePivotReport(workbook, report).values
  const valueName = report.fields.find(field => field.id === showAsUIIds.amount)?.displayName ?? 'Amount'
  const row = matrix.slice(3).find(row => labels.every((label, index) => Object.is(row[index], label)))
  const column = matrix[0].findIndex((value, index) => index >= 2 && Object.is(value, columns[0])
    && Object.is(matrix[1][index], columns[1]) && matrix[2][index] === valueName)
  if (!row || column < 0) throw new Error('Expected real hierarchy/value coordinate is missing')
  return row[column]
}
