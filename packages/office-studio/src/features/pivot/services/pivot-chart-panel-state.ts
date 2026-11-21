import type { ChartModel, ChartType } from '../../../domain/model.ts'
import type { ChartConfig } from '../../charts/chart-engine.ts'

export interface PivotChartDraft {
  chartId: string
  pivotTableId: string
  type: ChartType
  title: string
  position: string
  legendVisible: boolean
  legendPosition: 'top' | 'right' | 'bottom' | 'left'
}

export function createPivotChartDraft(chart?: ChartModel, pivotTableId = ''): PivotChartDraft {
  return {
    chartId: chart?.id ?? '',
    pivotTableId: chart?.pivotTableId ?? pivotTableId,
    type: chart?.type ?? 'bar',
    title: chart?.title ?? '',
    position: chart ? positionToAddress(chart.position.x, chart.position.y) : 'F2',
    legendVisible: chart?.legend.visible ?? true,
    legendPosition: chart?.legend.position ?? 'bottom',
  }
}

export function pivotChartConfig(draft: PivotChartDraft): ChartConfig & { pivotTableId: string } {
  if (!draft.pivotTableId.trim()) throw new TypeError('Select a pivot table for this chart')
  const position = addressToPosition(draft.position)
  return {
    type: draft.type,
    source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, orientation: 'columns' },
    pivotTableId: draft.pivotTableId,
    position,
    size: { width: 480, height: 300 },
    title: draft.title.trim() || undefined,
    legend: { visible: draft.legendVisible, position: draft.legendPosition },
  }
}

export function pivotChartPatch(draft: PivotChartDraft): Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>> {
  const config = pivotChartConfig(draft)
  return {
    type: config.type,
    source: config.source,
    pivotTableId: config.pivotTableId,
    position: config.position,
    size: config.size,
    title: config.title,
    legend: { visible: draft.legendVisible, position: draft.legendPosition },
  }
}

export function addressToPosition(address: string): { x: number; y: number } {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(address.trim())
  if (!match) throw new TypeError('Enter a valid cell address such as F2')
  let column = 0
  for (const char of match[1]!.toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64
  return { x: (column - 1) * 96, y: (Number(match[2]) - 1) * 24 }
}

function positionToAddress(x: number, y: number): string {
  let column = Math.floor(Math.max(0, x) / 96) + 1
  let name = ''
  while (column > 0) {
    const remainder = (column - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    column = Math.floor((column - 1) / 26)
  }
  return `${name}${Math.floor(Math.max(0, y) / 24) + 1}`
}
