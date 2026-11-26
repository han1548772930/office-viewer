import type { ChartModel, PivotChartControls } from '../../../domain/model.ts'

/** The two fixed v1.6.0 presentation controls are author state, not report fields. */
export function normalizePivotChartControls(input: unknown): PivotChartControls {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('透视图按钮设置无效。')
  const value = input as Record<string, unknown>
  if (Object.keys(value).some(key => key !== 'showFieldButtons' && key !== 'showExpandCollapseButtons')
    || Object.values(value).some(value => typeof value !== 'boolean')
    || value.showFieldButtons !== undefined && typeof value.showFieldButtons !== 'boolean'
    || value.showExpandCollapseButtons !== undefined && typeof value.showExpandCollapseButtons !== 'boolean') throw new TypeError('透视图按钮设置需要布尔值。')
  return { showFieldButtons: value.showFieldButtons !== false, showExpandCollapseButtons: value.showExpandCollapseButtons !== false }
}

export function pivotChartControls(chart: Pick<ChartModel, 'controls'>): PivotChartControls {
  return normalizePivotChartControls(chart.controls ?? {})
}
