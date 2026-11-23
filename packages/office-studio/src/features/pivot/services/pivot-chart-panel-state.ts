import type { ChartModel, ChartType } from '../../../domain/model.ts'
import type { ChartConfig } from '../../charts/chart-engine.ts'
import { chartConfigFromDraft, createChartSettingsDraft, parseChartSourceRange, type ChartSettingsDraft } from '../../charts/chart-settings-draft.ts'

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

export function createPivotChartDraft(chart?: ChartModel, pivotTableId = ''): PivotChartDraft {
  const position = chart ? positionToAddress(chart.position.x, chart.position.y) : 'F2'
  const settings = createChartSettingsDraft(chart ?? {
    type: 'bar', source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, orientation: 'columns' },
    pivotTableId, position: addressToPosition(position), size: { width: 480, height: 300 },
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

export function pivotChartConfig(draft: PivotChartDraft): ChartConfig & { pivotTableId: string } {
  if (!draft.pivotTableId.trim()) throw new TypeError('Select a pivot table for this chart')
  const config = chartConfigFromDraft(draft.settings)
  // The address is a placement shortcut. Leaving it unchanged preserves the
  // precise dragged position, source, size and full authored chart settings.
  return { ...config, pivotTableId: draft.pivotTableId,
    position: draft.position.trim().toUpperCase() === draft.initialPositionText ? { ...config.position } : addressToPosition(draft.position) }
}

export function pivotChartPatch(draft: PivotChartDraft): Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>> {
  return pivotChartConfig(draft)
}

export function addressToPosition(address: string): { x: number; y: number } {
  const range = parseChartSourceRange(address)
  if (!range || address.includes(':')) throw new TypeError('Enter a valid cell address such as F2')
  return { x: range.startCol * 96, y: range.startRow * 24 }
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
