import type { ChartSemanticOptions } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { numberControl, optionGroup, selectControl, type ChartSettingControl } from './chart-settings-controls'
import { DEFAULT_CHART_HISTOGRAM_BIN_COUNT, DEFAULT_CHART_HISTOGRAM_BIN_GAP, DEFAULT_CHART_HISTOGRAM_BIN_WIDTH,
  MAX_CHART_HISTOGRAM_BINS, MAX_CHART_HISTOGRAM_BIN_GAP } from './chart-histogram'

type HistogramKey = keyof NonNullable<ChartSemanticOptions['histogram']>
function resetField(control: ChartSettingControl, key: HistogramKey): ChartSettingControl {
  return { ...control, canReset: config => config.options?.histogram?.[key] !== undefined,
    reset: config => { if (config.options?.histogram) config.options.histogram[key] = undefined } }
}

/** Data controls precede paint controls in the fixed contextual panel. Binning
 * reset removes authorship and lets the statistics service use Scott Auto. */
export function chartHistogramDataControls(config: ChartConfig): ChartSettingControl[] {
  const controls: ChartSettingControl[] = [resetField(selectControl('histogram.binning.mode', '分箱方式',
    [['auto', 'Scott 自动'], ['count', '按箱数'], ['width', '按箱宽']], item => item.options?.histogram?.binning?.mode ?? 'auto', (item, value) => {
      const histogram = optionGroup(item, 'histogram'), current = histogram.binning
      histogram.binning = value === 'count' ? { mode: 'count', count: current?.mode === 'count' ? current.count : DEFAULT_CHART_HISTOGRAM_BIN_COUNT }
        : value === 'width' ? { mode: 'width', width: current?.mode === 'width' ? current.width : DEFAULT_CHART_HISTOGRAM_BIN_WIDTH } : undefined
    }, 'auto'), 'binning')]
  if (config.options?.histogram?.binning?.mode === 'count') controls.push({ ...resetField(numberControl('histogram.binning.count', '箱数',
    item => item.options?.histogram?.binning?.mode === 'count' ? item.options.histogram.binning.count : undefined,
    (item, value) => { optionGroup(item, 'histogram').binning = value === undefined ? undefined : { mode: 'count', count: value } },
    1, MAX_CHART_HISTOGRAM_BINS, DEFAULT_CHART_HISTOGRAM_BIN_COUNT), 'binning'), precision: 0, step: 1, extra: '箱数包含启用的下溢和上溢箱。' })
  if (config.options?.histogram?.binning?.mode === 'width') controls.push(resetField(numberControl('histogram.binning.width', '箱宽',
    item => item.options?.histogram?.binning?.mode === 'width' ? item.options.histogram.binning.width : undefined,
    (item, value) => { optionGroup(item, 'histogram').binning = value === undefined ? undefined : { mode: 'width', width: value } },
    Number.MIN_VALUE, undefined, DEFAULT_CHART_HISTOGRAM_BIN_WIDTH), 'binning'))
  controls.push(
    resetField(numberControl('histogram.underflowThreshold', '下溢箱阈值（≤）', item => item.options?.histogram?.underflowThreshold,
      (item, value) => { optionGroup(item, 'histogram').underflowThreshold = value }), 'underflowThreshold'),
    resetField(numberControl('histogram.overflowThreshold', '上溢箱阈值（>）', item => item.options?.histogram?.overflowThreshold,
      (item, value) => { optionGroup(item, 'histogram').overflowThreshold = value }), 'overflowThreshold'),
  )
  return controls
}

export function chartHistogramStyleControls(): ChartSettingControl[] {
  return [{ ...resetField(numberControl('histogram.binGap', '箱间隙比例', item => item.options?.histogram?.binGap,
    (item, value) => { optionGroup(item, 'histogram').binGap = value }, 0, MAX_CHART_HISTOGRAM_BIN_GAP, DEFAULT_CHART_HISTOGRAM_BIN_GAP), 'binGap'), step: 0.05 }]
}
