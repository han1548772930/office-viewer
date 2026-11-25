import type { ChartAxisOptions, ChartDataPointOptions, ChartModel, ChartSeriesOptions } from '../../domain/model'
import type { ChartConfig, ChartUpdatePatch } from './chart-engine'
import type { ChartCanvasElement } from './chart-element-interaction'
import { chartAxisControls, chartCommonControls, chartFontControls, chartLabelControls, chartLegendFont, chartLineControls,
  chartSeriesControls, chartColorControl, colorControl, numberControl, selectControl, seriesStyleGroup, optionGroup, type ChartSettingControl } from './chart-settings-controls'
import { chartSeriesOverride } from './chart-series-selector'
import { chartPalette, chartSeriesStyle } from './chart-options'
import { chartPointStyle } from './chart-derived'
import { chartNullableValues, patchChartSeriesOverride } from './chart-cartesian-api'
import { cloneChartDraftConfig } from './chart-settings-draft'
import { CHART_POINT_SHAPE_CHOICES } from './chart-markers'
import { supportsChartDateAxis } from './chart-date-axis'
import { resolveChartType } from './chart-catalog'
import { chartSpecialMarkAuthorSignature, chartSpecialMarkCompatible, chartSpecialMarkControls, chartSpecialMarkHidden, chartSpecialMarkLabel,
  chartSpecialMarkPatch, deleteChartSpecialMarkDraft, resetChartSpecialMarkDraft } from './chart-special-mark-edit'

export const CHART_ELEMENT_LABEL_ROTATIONS = [0, 30, 45, 60, 90] as const
const axisLabels = { xAxis: '分类／水平轴', yAxis: '数值／垂直轴', rightYAxis: '次要数值轴' } as const
export function chartCanvasElementCompatible(config: ChartConfig, element: ChartCanvasElement): boolean {
  return element.kind !== 'mark' || chartSpecialMarkCompatible(config, element.directEdit)
}
export function chartCanvasElementLabel(element: ChartCanvasElement): string {
  if (element.kind === 'mark') return chartSpecialMarkLabel(element.directEdit)
  if (element.kind === 'axisTitle') return `${axisLabels[element.axisId]}标题`
  if (element.kind === 'axis') return axisLabels[element.axisId]
  if (element.kind === 'gridLine') return `${axisLabels[element.axisId]}网格线`
  if (element.kind === 'tick') return `${axisLabels[element.axisId]}刻度线`
  if (element.kind === 'legend') return '图例'
  return element.kind === 'series' ? `系列 · ${element.binding.sourceName}`
    : element.kind === 'dataPoint' ? `数据点 ${element.dataIndex + 1} · ${element.binding.sourceName}` : `数据标签 · ${element.binding.sourceName}`
}
function axis(config: ChartConfig, element: Extract<ChartCanvasElement, { axisId: unknown }>): Partial<ChartAxisOptions> | undefined {
  return element.axisId === 'rightYAxis' ? config.options?.rightYAxis : config[element.axisId]
}
function writeAxis(config: ChartConfig, element: Extract<ChartCanvasElement, { axisId: unknown }>): ChartAxisOptions {
  if (element.axisId === 'rightYAxis') return (config.options ??= {}).rightYAxis ??= { scaleType: 'linear' }
  return config[element.axisId] ??= { scaleType: 'linear' }
}
function pointControls(element: Extract<ChartCanvasElement, { kind: 'dataPoint' }>): ChartSettingControl[] {
  const { binding, dataIndex } = element, selector = binding.selector
  const read = (config: ChartConfig): ChartDataPointOptions | undefined => chartSeriesOverride(config.options, selector)?.dataPoints?.[dataIndex]
  const write = (config: ChartConfig): ChartDataPointOptions => (seriesStyleGroup(config, selector).dataPoints ??= {})[dataIndex] ??= {}
  const effective = (config: ChartConfig): ChartSeriesOptions => chartPointStyle(chartSeriesStyle(config, binding.field,
    ['pie', 'funnel', 'gauge'].includes(resolveChartType(config.type).family) ? dataIndex : binding.sourcePosition,
    chartPalette(config), binding.sourceName), dataIndex, config.type)
  return [
    chartColorControl('point.color', '数据点颜色', config => read(config)?.color, (config, value) => { const point = write(config); point.color = value; point.fillPattern = null }, 'dataPoint', config => effective(config).color),
    numberControl('point.fillOpacity', '填充不透明度', config => read(config)?.fillOpacity, (config, value) => { write(config).fillOpacity = value }, 0, 1),
    numberControl('point.size', '标记大小', config => read(config)?.size, (config, value) => { write(config).size = value }, 0, 256),
    selectControl('point.shape', '标记形状', CHART_POINT_SHAPE_CHOICES, config => read(config)?.shape, (config, value) => { write(config).shape = value }),
    colorControl('point.border.color', '边框颜色', config => read(config)?.border?.color, (config, value) => { (write(config).border ??= {}).color = value }),
    numberControl('point.border.width', '边框宽度', config => read(config)?.border?.width, (config, value) => { (write(config).border ??= {}).width = value }, 0, 30),
    numberControl('point.border.opacity', '边框不透明度', config => read(config)?.border?.opacity, (config, value) => { (write(config).border ??= {}).opacity = value }, 0, 1),
  ].map(control => ({ ...control, visible: (config: ChartConfig) => !['point.size', 'point.shape'].includes(control.key)
    || ['line', 'area', 'scatter', 'bubble', 'radar', 'combination', 'boxplot'].includes(resolveChartType(config.type).family)
      && (config.type !== 'boxplot' || config.options?.boxplot?.showOutliers !== false) }))
}

/** The float form reuses the tested model controls used by the full editor.
 * Its apply planner below limits writes to the selected author scope. */
export function chartCanvasElementControls(element: ChartCanvasElement, config: ChartConfig): ChartSettingControl[] {
  if (element.kind === 'mark') return chartSpecialMarkControls(element.directEdit, config)
  if (element.kind === 'axisTitle' || element.kind === 'axis' || element.kind === 'gridLine' || element.kind === 'tick') {
    const key = element.axisId, dates = key === 'xAxis' && supportsChartDateAxis(config) && !!config.options?.mapping?.useDateAxis
    if (element.kind === 'gridLine' || element.kind === 'tick') {
      const group = element.kind
      const controls = chartLineControls(`${key}.${group}`, item => axis(item, element)?.[group], item => writeAxis(item, element)[group] ??= {}, element.kind === 'gridLine')
      if (group === 'tick') controls.push(
        numberControl(`${key}.tick.length`, '刻度线长度', item => axis(item, element)?.tick?.length, (item, value) => { (writeAxis(item, element).tick ??= {}).length = value }, 0, 50, 4),
        selectControl(`${key}.tick.position`, '刻度线位置', [['inside', '内部'], ['outside', '外部']], item => axis(item, element)?.tick?.position, (item, value) => { (writeAxis(item, element).tick ??= {}).position = value }, 'outside'))
      return controls
    }
    const numeric = key !== 'xAxis' || dates || config.type === 'scatter' || config.type === 'bubble'
    const controls = chartAxisControls(key, numeric, dates).filter(control => element.kind === 'axisTitle' ? control.key.startsWith(`${key}.title`)
      : !control.key.startsWith(`${key}.title`) && !control.key.startsWith(`${key}.gridLine`) && !control.key.startsWith(`${key}.tick`))
    if (element.kind === 'axis') {
      const rotation = controls.findIndex(control => control.key === `${key}.label.rotate`)
      if (rotation >= 0) controls[rotation] = selectControl(`${key}.label.rotate`, '标签旋转', CHART_ELEMENT_LABEL_ROTATIONS.map(value => [String(value), `${value}°`] as const),
        item => axis(item, element)?.label?.rotate === undefined ? undefined : String(axis(item, element)!.label!.rotate),
        (item, value) => { (writeAxis(item, element).label ??= {}).rotate = value === undefined ? undefined : Number(value) }, '0')
      if (key === 'xAxis' && supportsChartDateAxis(config)) controls.push(...chartDateAxisToggle(element))
    }
    return controls
  }
  if (element.kind === 'legend') return [...chartCommonControls(config).filter(control => control.key.startsWith('legend.')),
    ...chartFontControls('legend.font', item => item.legend?.font, chartLegendFont)]
  if (element.kind === 'dataPoint') return pointControls(element)
  if (element.kind === 'dataLabel' && resolveChartType(config.type).family === 'pie') return chartLabelControls('pie.label', item => item.options?.pie?.label, item => optionGroup(item, 'pie').label ??= {}, true)
  const { binding } = element
  const controls = chartSeriesControls(binding.selector, false, false, false,
    { field: binding.field, position: binding.sourcePosition, name: binding.sourceName }, config.type === 'bubble')
  const family = resolveChartType(config.type).family
  return element.kind === 'dataLabel' ? controls.filter(control => control.key.includes('.label.'))
    : controls.filter(control => ['.name', '.color', '.fillOpacity', '.point.shape', '.point.size', '.point.color', '.border.color', '.border.width', '.border.opacity', '.border.dashType', '.label.visible', '.label.position', '.label.contentType'].some(suffix => control.key.endsWith(suffix)))
      .filter(control => !control.key.endsWith('.name') || family !== 'pie' && family !== 'funnel')
      .map(control => !control.key.includes('.point.') ? control : { ...control,
        visible: (item: ChartConfig) => control.visible?.(item) !== false
          && ['line', 'area', 'scatter', 'bubble', 'radar', 'combination', 'boxplot'].includes(resolveChartType(item.type).family)
          && (item.type !== 'boxplot' || item.options?.boxplot?.showOutliers !== false) })
}
function chartDateAxisToggle(element: Extract<ChartCanvasElement, { kind: 'axis' }>): ChartSettingControl[] {
  return [{ key: 'mapping.useDateAxis', label: '日期轴', kind: 'switch', fallback: false,
    read: config => config.options?.mapping?.useDateAxis, write: (config, value) => {
      optionGroup(config, 'mapping').useDateAxis = value
      if (value) { const selected = writeAxis(config, element); selected.scaleType = 'linear'; selected.logBase = undefined }
    } }]
}
const axisFormatFields = ['scaleType', 'logBase', 'min', 'max', 'baseTimeUnit', 'majorTimeUnit', 'majorUnit', 'reverse', 'line', 'label', 'format'] as const
function axisValues(config: ChartConfig, element: Extract<ChartCanvasElement, { axisId: unknown }>): Partial<ChartAxisOptions> {
  const source = axis(config, element), fields = element.kind === 'axisTitle' ? ['title', 'titleStyle', 'titleAlignment'] as const
    : element.kind === 'gridLine' ? ['gridLine'] as const : element.kind === 'tick' ? ['tick'] as const : axisFormatFields
  return Object.fromEntries(fields.map(key => [key, source?.[key]]))
}
export function chartCanvasElementAuthorSignature(config: ChartConfig, element: ChartCanvasElement): string {
  if (element.kind === 'mark') return chartSpecialMarkAuthorSignature(config, element.directEdit)
  if (element.kind === 'axisTitle' || element.kind === 'axis' || element.kind === 'gridLine' || element.kind === 'tick') return JSON.stringify([axisValues(config, element),
    element.kind === 'axis' && element.axisId === 'xAxis' ? config.options?.mapping?.useDateAxis : undefined])
  if (element.kind === 'legend') return JSON.stringify(config.legend) ?? ''
  if (element.kind === 'dataLabel' && resolveChartType(config.type).family === 'pie') return JSON.stringify(config.options?.pie?.label) ?? ''
  const style = chartSeriesOverride(config.options, element.binding.selector)
  return JSON.stringify(element.kind === 'dataPoint' ? style?.dataPoints?.[element.dataIndex] : element.kind === 'dataLabel' ? style?.label : style) ?? ''
}
export function chartCanvasElementPatch(original: ChartModel, draft: ChartConfig, element: ChartCanvasElement): ChartUpdatePatch {
  if (element.kind === 'mark') return chartSpecialMarkPatch(original, draft, element.directEdit)
  if (chartCanvasElementAuthorSignature(original, element) === chartCanvasElementAuthorSignature(draft, element)) return {}
  if (element.kind === 'axisTitle' || element.kind === 'axis' || element.kind === 'gridLine' || element.kind === 'tick') {
    const values = axisValues(draft, element), patch: ChartUpdatePatch = element.axisId === 'rightYAxis' ? { options: { rightYAxis: { scaleType: 'linear', ...axis(original, element), ...values } } }
      : { [element.axisId]: { scaleType: 'linear', ...axis(original, element), ...values } }
    if (element.kind === 'axis' && element.axisId === 'xAxis' && original.options?.mapping?.useDateAxis !== draft.options?.mapping?.useDateAxis)
      patch.options = { ...patch.options, mapping: { useDateAxis: draft.options?.mapping?.useDateAxis } }
    return patch
  }
  if (element.kind === 'legend') return { legend: { visible: true, position: 'bottom', ...draft.legend } }
  if (element.kind === 'dataLabel' && resolveChartType(original.type).family === 'pie') return { options: { pie: { label: draft.options?.pie?.label } } }
  const selector = element.binding.selector, authored = chartSeriesOverride(draft.options, selector)
  const style = element.kind === 'dataPoint' ? chartNullableValues<ChartSeriesOptions>({ dataPoints: { [element.dataIndex]: authored?.dataPoints?.[element.dataIndex] ?? null } })
    : element.kind === 'dataLabel' ? { label: authored?.label } : authored
  return { options: patchChartSeriesOverride(original.options, selector, style) }
}
export function resetChartCanvasElementDraft(config: ChartConfig, element: ChartCanvasElement): ChartConfig {
  if (element.kind === 'mark') return resetChartSpecialMarkDraft(config, element.directEdit)
  const draft = cloneChartDraftConfig(config)
  if (element.kind === 'axisTitle') { const selected = writeAxis(draft, element); selected.titleStyle = undefined; selected.titleAlignment = undefined }
  else if (element.kind === 'axis') {
    const selected = writeAxis(draft, element)
    for (const key of axisFormatFields) if (key !== 'scaleType') Object.assign(selected, { [key]: undefined })
    selected.scaleType = 'linear'
  } else if (element.kind === 'gridLine' || element.kind === 'tick') writeAxis(draft, element)[element.kind] = undefined
  else if (element.kind === 'legend') draft.legend = { visible: true, position: 'bottom', font: undefined, wrap: undefined, selectMode: undefined }
  else if (element.kind === 'dataLabel' && resolveChartType(config.type).family === 'pie') optionGroup(draft, 'pie').label = undefined
  else {
    const style = seriesStyleGroup(draft, element.binding.selector)
    if (element.kind === 'dataPoint') { if (style.dataPoints) delete style.dataPoints[element.dataIndex] }
    else if (element.kind === 'dataLabel') style.label = undefined
    else for (const key of ['color', 'fillOpacity', 'fillPattern', 'border', 'point', 'label'] as const) Object.assign(style, { [key]: undefined })
  }
  return draft
}

/** Delete hides a painted chrome element. Data values and series are edited
 * through their source controls; a selected point never deletes its chart. */
export function chartCanvasElementDeleteDraft(config: ChartConfig, element: ChartCanvasElement): ChartConfig | undefined {
  if (element.kind === 'mark') return deleteChartSpecialMarkDraft(config, element.directEdit)
  const draft = cloneChartDraftConfig(config)
  if (element.kind === 'axisTitle') { const selected = writeAxis(draft, element); selected.title = undefined; selected.titleStyle = undefined; selected.titleAlignment = undefined }
  else if (element.kind === 'axis') { const selected = writeAxis(draft, element); (selected[element.part] ??= {}).visible = false }
  else if (element.kind === 'gridLine' || element.kind === 'tick') (writeAxis(draft, element)[element.kind] ??= {}).visible = false
  else if (element.kind === 'legend') (draft.legend ??= { visible: true, position: 'bottom' }).visible = false
  else if (element.kind === 'dataLabel') {
    if (resolveChartType(config.type).family === 'pie') (optionGroup(draft, 'pie').label ??= {}).visible = false
    else (seriesStyleGroup(draft, element.binding.selector).label ??= {}).visible = false
  } else return undefined
  return draft
}

export function chartCanvasElementHidden(config: ChartConfig, element: ChartCanvasElement): boolean {
  if (element.kind === 'mark') return chartSpecialMarkHidden(config, element.directEdit)
  if (element.kind === 'axisTitle') return !axis(config, element)?.title
  if (element.kind === 'axis') return axis(config, element)?.[element.part]?.visible === false
  if (element.kind === 'gridLine' || element.kind === 'tick') return axis(config, element)?.[element.kind]?.visible === false
  if (element.kind === 'legend') return config.legend?.visible === false
  if (element.kind === 'dataLabel') return (resolveChartType(config.type).family === 'pie' ? config.options?.pie?.label
    : chartSeriesOverride(config.options, element.binding.selector)?.label)?.visible === false
  return false
}
