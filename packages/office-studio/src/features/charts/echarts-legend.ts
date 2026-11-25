import type { EChartsOption } from 'echarts'
import type { ChartLegendSelectMode } from '../../domain/model'
import type { ChartLegendSelection } from './chart-legend'
import type { ChartEChartsOption, ChartEChartsPointerRuntime } from './echarts-option'

export interface ChartCanvasLegendFrameOption {
  option: ChartEChartsOption
  pointer?: ChartEChartsPointerRuntime
}

/** A renderer-owned operation. Functions and selected/scroll state are stripped
 * from author signatures, snapshots, compute messages and ECharts.setOption. */
export interface ChartCanvasLegendRuntime {
  names: string[]
  selectMode: ChartLegendSelectMode
  rebuild?: (selected: ChartLegendSelection, retainedPointer?: ChartEChartsPointerRuntime) => ChartCanvasLegendFrameOption
  preparePointer?: (selected: ChartLegendSelection) => ChartEChartsPointerRuntime | undefined
}

export type ChartCanvasLegendHit = { kind: 'item'; name: string } | { kind: 'page'; direction: 'previous' | 'next' }

export function chartNativeLegendNames(option: EChartsOption): string[] {
  const legend = Array.isArray(option.legend) ? option.legend[0] : option.legend
  if (!legend || legend.show === false) return []
  return (legend.data ?? []).flatMap(item => {
    const name = typeof item === 'string' ? item : item.name
    return [typeof name === 'string' ? name : '']
  })
}

export function chartNativeLegendSelection(option: EChartsOption, selected: ChartLegendSelection, scrollDataIndex = 0): EChartsOption {
  if (!option.legend) return option
  return { ...option, legend: Array.isArray(option.legend)
    ? option.legend.map(legend => ({ ...legend, selected: { ...selected }, scrollDataIndex })) : { ...option.legend, selected: { ...selected }, scrollDataIndex } }
}

export function chartNativeLegendScrollIndex(option: EChartsOption): number {
  const legend = Array.isArray(option.legend) ? option.legend[0] : option.legend
  return legend && 'scrollDataIndex' in legend && typeof legend.scrollDataIndex === 'number' && Number.isSafeInteger(legend.scrollDataIndex) ? Math.max(0, legend.scrollDataIndex) : 0
}

/** ECharts 6.1's isolated Canvas hit adapter: LegendView tags item groups with
 * __legendDataIndex and ScrollableLegendView names its native page icons.
 * Only the real painted/clipped hit path enters this walk. */
export function chartNativeLegendHit(element: { parent?: unknown; name?: string } | undefined, names: readonly string[]): ChartCanvasLegendHit | undefined {
  let current: unknown = element
  while (current && typeof current === 'object') {
    if ('__legendDataIndex' in current && typeof current.__legendDataIndex === 'number' && Number.isSafeInteger(current.__legendDataIndex)) {
      const name = names[current.__legendDataIndex]
      if (name !== undefined && name !== '' && name !== '\n') return { kind: 'item', name }
    }
    if ('name' in current && (current.name === 'pagePrev' || current.name === 'pageNext')) return { kind: 'page', direction: current.name === 'pagePrev' ? 'previous' : 'next' }
    current = 'parent' in current ? current.parent : undefined
  }
  return undefined
}
