import { pivotChartFieldButtonsIds as ids, pivotChartFieldButtonsWorkbook } from './pivot-chart-field-buttons-fixture.mjs'
import { applyPivotChartEdit, getPivotChartReport } from '../../src/features/pivot/services/pivot-chart-engine.ts'

// Deferred actual sparse workbook: two measures and up to two dimension fields per axis.
// ΣValues stays in options; no fake source field is registered.
export function pivotValuesUiWorkbook({ multipleValues = true, dimensions = 2, valuePosition = 'column', valueIndex = -1, ...options } = {}) {
  const workbook = pivotChartFieldButtonsWorkbook({ ...options, overflow: multipleValues })
  for (const chartId of [ids.owned, ids.linked]) {
    const chart = workbook.charts.find(chart => chart.id === chartId), report = getPivotChartReport(workbook, chart)
    const fields = structuredClone(report.fields).filter(field => field.aggregation !== 'count')
    for (const field of fields.filter(field => field.area === 'value')) field.displayName = field.position === 0 ? 'Amount sum' : 'Amount average'
    if (dimensions > 1) {
      const date = fields.find(field => field.name === 'Date')
      date.area = 'row'; date.position = 1
      fields.push({ id: `${report.id}:score-axis`, sourceFieldId: '4', sourceColumn: 7, name: 'Score', area: 'column', position: 1 })
    }
    applyPivotChartEdit(workbook, chart, { type: 'setFieldsConfig', config: { fields,
      options: { ...report.options, valuePosition, valueIndex } } })
  }
  return workbook
}
