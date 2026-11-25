import type { ChartColor, ChartLabelPattern } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartAutomaticGradient } from './chart-auto-gradient'
import { chartColorChannels, chartSolidColorOpacity, isChartLinearGradient } from './chart-color'

/** Fixed wp reduces patterns to their mean paint and retains authored gradients. */
export function chartLabelBackground(background: ChartColor | ChartLabelPattern | undefined): ChartColor | undefined {
  if (background === undefined || typeof background === 'string' || isChartLinearGradient(background)) return background
  const foreground = chartColorChannels(background.foregroundColor), base = chartColorChannels(background.backgroundColor)
  const percent = /^pct(\d+)$/.exec(background.preset), ratio = percent ? Math.min(100, Number(percent[1])) / 100 : 0.5
  const channels = base.map((channel, index) => channel + (foreground[index]! - channel) * ratio)
  return `rgba(${Math.round(channels[0]! * 255)}, ${Math.round(channels[1]! * 255)}, ${Math.round(channels[2]! * 255)}, ${channels[3]})`
}

export function chartLabelBackgroundPaint(background: ChartColor | undefined, bounds: RenderRect): ChartAutomaticGradient | undefined {
  if (!isChartLinearGradient(background)) return undefined
  return { authored: true, start: { ...background.start }, end: { ...background.end }, bounds: { ...bounds },
    stops: background.stops.map(stop => ({ offset: stop.offset, color: chartSolidColorOpacity(stop.color, stop.opacity ?? 1) })).sort((a, b) => a.offset - b.offset) }
}
