import type { ChartModel, ChartType } from '../../../domain/model.ts'
import type { ChartConfig } from '../../charts/chart-engine.ts'
import { chartColumnName, chartConfigFromDraft, createChartSettingsDraft, type ChartSettingsDraft } from '../../charts/chart-settings-draft.ts'
import { chartCellPosition, chartPositionCell, type ChartPlacementGeometry } from '../../charts/chart-placement'

export interface PivotChartDraft {
  chartId: string
  pivotTableId: string
  type: ChartType
  title: string
  position: string
  legendVisible: boolean
  legendPosition: 'top' | 'right' | 'bottom' | 'left'
  settings: ChartSettingsDraft
  initialPositionText: string
}

export function createPivotChartDraft(chart?: ChartModel, pivotTableId = '', geometry?: ChartPlacementGeometry): PivotChartDraft {
  const position = chart ? chartPositionCell(chart.position, geometry) : geometry
    ? `${chartColumnName(Math.min(5, geometry.columns.count - 1))}${Math.min(1, geometry.rows.count - 1) + 1}` : 'F2'
  const settings = createChartSettingsDraft(chart ?? {
    type: 'bar', source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, orientation: 'columns' },
    pivotTableId, position: chartCellPosition(position, geometry), size: { width: 480, height: 300 },
  })
  return {
    chartId: chart?.id ?? '',
    pivotTableId: chart?.pivotTableId ?? pivotTableId,
    settings,
    initialPositionText: position,
    position,
    get type() { return this.settings.config.type },
    set type(value: ChartType) { this.settings.config.type = value },
    get title() { return this.settings.config.title ?? '' },
    set title(value: string) { this.settings.config.title = value },
    get legendVisible() { return this.settings.config.legend?.visible ?? true },
    set legendVisible(value: boolean) { this.settings.config.legend = { ...this.settings.config.legend, visible: value } },
    get legendPosition() { return this.settings.config.legend?.position ?? 'bottom' },
    set legendPosition(value: 'top' | 'right' | 'bottom' | 'left') { this.settings.config.legend = { ...this.settings.config.legend, position: value } },
  }
}

export function pivotChartConfig(draft: PivotChartDraft, geometry?: ChartPlacementGeometry): ChartConfig & { pivotTableId: string } {
  if (!draft.pivotTableId.trim()) throw new TypeError('Select a pivot table for this chart')
  const config = chartConfigFromDraft(draft.settings)
  // The address is a placement shortcut. Leaving it unchanged preserves the
  // precise dragged position, source, size and full authored chart settings.
  return { ...config, pivotTableId: draft.pivotTableId,
    position: draft.position.trim().toUpperCase() === draft.initialPositionText ? { ...config.position } : chartCellPosition(draft.position, geometry) }
}

export function pivotChartPatch(draft: PivotChartDraft, geometry?: ChartPlacementGeometry): Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>> {
  return pivotChartConfig(draft, geometry)
}

export function addressToPosition(address: string, geometry?: ChartPlacementGeometry): { x: number; y: number } { return chartCellPosition(address, geometry) }
