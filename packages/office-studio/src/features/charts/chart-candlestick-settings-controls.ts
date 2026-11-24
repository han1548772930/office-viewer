import type { ChartCandlestickStateStyleOptions } from '../../domain/model'
import type { ChartConfig, ChartSeriesData } from './chart-engine'
import { chartCandlestickOverlayOptions, chartCandlestickStyle, type ChartCandlestickMode, type ChartCandlestickState } from './chart-candlestick'
import { CHART_POINT_SHAPE_CHOICES } from './chart-markers'
import { chartPalette } from './chart-options'
import { colorControl, numberControl, optionGroup, selectControl, switchControl, type ChartSettingControl } from './chart-settings-controls'

export type ChartCandlestickResetTarget = ChartCandlestickState | 'candleWidth' | 'highLowLine' | 'closeMarker' | 'all'

/** Clearing a state removes that authored layer. Earlier snapshot fields and
 * generic series styles remain inherited; clearing all removes legacy fields. */
export function clearChartCandlestickSettings(config: ChartConfig, target: ChartCandlestickResetTarget): void {
  if (!config.options?.candlestick) return
  // The settings command deep-merges patches: explicit undefined clears the
  // previous author property, and serialized snapshots omit that property.
  if (target === 'all') config.options.candlestick = undefined
  else config.options.candlestick[target] = undefined
}

function resettable(control: ChartSettingControl, canReset: (config: ChartConfig) => boolean,
  reset: (config: ChartConfig) => void): ChartSettingControl {
  return { ...control, canReset, reset }
}
function stateGroup(config: ChartConfig, state: ChartCandlestickState): ChartCandlestickStateStyleOptions {
  return optionGroup(config, 'candlestick')[state] ??= {}
}
function effectiveState(config: ChartConfig, state: ChartCandlestickState, source?: ChartSeriesData) {
  return chartCandlestickStyle(config, source?.field ?? config.valueFields?.[0] ?? 0, undefined, state,
    chartPalette(config), source?.sourceName ?? source?.name)
}

export function chartCandlestickWidthControls(): ChartSettingControl[] {
  return [{ ...resettable(numberControl('candlestick.candleWidth', '蜡烛宽度', config => config.options?.candlestick?.candleWidth,
    (config, value) => { optionGroup(config, 'candlestick').candleWidth = value }, 1),
  config => config.options?.candlestick?.candleWidth !== undefined, config => clearChartCandlestickSettings(config, 'candleWidth')),
  extra: '留空使用自动宽度。' }]
}

export function chartCandlestickBodyControls(state: ChartCandlestickState, source?: ChartSeriesData): ChartSettingControl[] {
  if (state === 'doji') return []
  const read = (config: ChartConfig) => config.options?.candlestick?.[state], key = `candlestick.${state}`
  return [
    resettable(colorControl(`${key}.color`, '蜡烛填充颜色', config => effectiveState(config, state, source).color,
      (config, value) => { stateGroup(config, state).color = value }), config => read(config)?.color !== undefined,
    config => { const current = read(config); if (current) current.color = undefined }),
    resettable(numberControl(`${key}.fillOpacity`, '蜡烛填充不透明度', config => effectiveState(config, state, source).fillOpacity,
      (config, value) => { stateGroup(config, state).fillOpacity = value }, 0, 1, 1), config => read(config)?.fillOpacity !== undefined,
    config => { const current = read(config); if (current) current.fillOpacity = undefined }),
    resettable(switchControl(`${key}.hollow`, '空心蜡烛', config => effectiveState(config, state, source).hollow,
      (config, value) => { stateGroup(config, state).hollow = value }), config => read(config)?.hollow !== undefined,
    config => { const current = read(config); if (current) current.hollow = undefined }),
  ]
}

export function chartCandlestickBorderControls(state: ChartCandlestickState, source?: ChartSeriesData): ChartSettingControl[] {
  const read = (config: ChartConfig) => config.options?.candlestick?.[state]?.border, key = `candlestick.${state}.border`
  const write = (config: ChartConfig) => stateGroup(config, state).border ??= {}
  return [
    resettable(colorControl(`${key}.color`, '蜡烛边框颜色', config => effectiveState(config, state, source).border?.color,
      (config, value) => { write(config).color = value }), config => read(config)?.color !== undefined,
    config => { const current = read(config); if (current) current.color = undefined }),
    resettable(numberControl(`${key}.width`, '蜡烛边框宽度', config => effectiveState(config, state, source).border?.width,
      (config, value) => { write(config).width = value }, 0), config => read(config)?.width !== undefined,
    config => { const current = read(config); if (current) current.width = undefined }),
    resettable(numberControl(`${key}.opacity`, '蜡烛边框不透明度', config => effectiveState(config, state, source).border?.opacity,
      (config, value) => { write(config).opacity = value }, 0, 1, 1), config => read(config)?.opacity !== undefined,
    config => { const current = read(config); if (current) current.opacity = undefined }),
  ]
}

export function chartCandlestickHighLowControls(mode: ChartCandlestickMode): ChartSettingControl[] {
  const read = (config: ChartConfig) => config.options?.candlestick?.highLowLine
  const effective = (config: ChartConfig) => chartCandlestickOverlayOptions(config.options?.candlestick, mode).highLowLine
  const group = (config: ChartConfig) => optionGroup(config, 'candlestick').highLowLine ??= {}
  const border = (config: ChartConfig) => group(config).border ??= {}
  const visible = (config: ChartConfig) => effective(config).visible
  return [
    resettable(switchControl('candlestick.highLowLine.visible', '显示高低线', config => effective(config).visible,
      (config, value) => { group(config).visible = value }), config => read(config)?.visible !== undefined,
    config => { const current = read(config); if (current) current.visible = undefined }),
    { ...resettable(colorControl('candlestick.highLowLine.border.color', '高低线颜色', config => effective(config).border.color,
      (config, value) => { border(config).color = value }), config => read(config)?.border?.color !== undefined,
    config => { const current = read(config)?.border; if (current) current.color = undefined }), visible },
    { ...resettable(numberControl('candlestick.highLowLine.border.width', '高低线宽度', config => effective(config).border.width,
      (config, value) => { border(config).width = value }, 0), config => read(config)?.border?.width !== undefined,
    config => { const current = read(config)?.border; if (current) current.width = undefined }), visible },
    { ...resettable(selectControl('candlestick.highLowLine.border.dashType', '高低线线型', [['solid', '实线'], ['dashed', '虚线'], ['dotted', '点线']],
      config => effective(config).border.dashType, (config, value) => { border(config).dashType = value }, 'solid'), config => read(config)?.border?.dashType !== undefined,
    config => { const current = read(config)?.border; if (current) current.dashType = undefined }), visible },
    { ...resettable(numberControl('candlestick.highLowLine.border.opacity', '高低线不透明度', config => effective(config).border.opacity,
      (config, value) => { border(config).opacity = value }, 0, 1, 1), config => read(config)?.border?.opacity !== undefined,
    config => { const current = read(config)?.border; if (current) current.opacity = undefined }), visible },
  ]
}

export function chartCandlestickCloseMarkerControls(mode: ChartCandlestickMode): ChartSettingControl[] {
  const read = (config: ChartConfig) => config.options?.candlestick?.closeMarker
  const effective = (config: ChartConfig) => chartCandlestickOverlayOptions(config.options?.candlestick, mode).closeMarker
  const group = (config: ChartConfig) => optionGroup(config, 'candlestick').closeMarker ??= {}
  const point = (config: ChartConfig) => group(config).point ??= {}
  const visible = (config: ChartConfig) => effective(config).visible
  return [
    resettable(switchControl('candlestick.closeMarker.visible', '显示收盘标记', config => effective(config).visible,
      (config, value) => { group(config).visible = value }), config => read(config)?.visible !== undefined,
    config => { const current = read(config); if (current) current.visible = undefined }),
    { ...resettable(colorControl('candlestick.closeMarker.point.color', '收盘标记颜色', config => effective(config).point.color,
      (config, value) => { point(config).color = value }), config => read(config)?.point?.color !== undefined,
    config => { const current = read(config)?.point; if (current) current.color = undefined }), visible },
    { ...resettable(selectControl('candlestick.closeMarker.point.shape', '收盘标记形状', CHART_POINT_SHAPE_CHOICES,
      config => effective(config).point.shape, (config, value) => { point(config).shape = value }, 'dot'), config => read(config)?.point?.shape !== undefined,
    config => { const current = read(config)?.point; if (current) current.shape = undefined }), visible },
    { ...resettable(numberControl('candlestick.closeMarker.point.size', '收盘标记大小', config => effective(config).point.size,
      (config, value) => { point(config).size = value }, 0, 256, 8), config => read(config)?.point?.size !== undefined,
    config => { const current = read(config)?.point; if (current) current.size = undefined }), visible },
  ]
}
