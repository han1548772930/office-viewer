<script setup lang="ts">
import { computed, nextTick, onScopeDispose, ref, watch } from 'vue'
import { Alert, Collapse, CollapsePanel, FormItem, Input, Select } from 'antdv-next'
import type { ChartConfig, ChartData } from '../chart-engine'
import { chartDataHint, chartTypeSelectOptions, isChartType, resolveChartType, supportsChartErrorDirection, supportsChartStatistics } from '../chart-catalog'
import { chartAxisControls, chartAxisPointerControls, chartCommonControls, chartDateAxisControls, chartFontControls, chartLegendFont, chartLineControls, chartSeriesControls, chartTitleFontControls, optionGroup } from '../chart-settings-controls'
import { chartAggregationControls, chartDataMappingControls, chartTypeSettingsControls } from '../chart-type-settings-controls'
import { chartDraftFieldCount, chartDraftFields, cloneChartDraftConfig, type ChartSettingsDraft, type ChartSettingsField, type ChartSettingsSourceSheet } from '../chart-settings-draft'
import { chartSettingsDateAxisError, projectChartSettings, type ChartSettingsProjection } from '../chart-settings-data'
import { chartRelationSourceFormat } from '../chart-relation-data'
import { supportsChartDateAxis } from '../chart-date-axis'
import { chartHasLegend } from '../chart-legend'
import { chartHierarchyItems, chartNumber } from '../chart-data'
import { chartGaugeSettingsMapping } from '../chart-gauge-settings-controls'
import { chartRadarCategoryAxisControls, chartRadarValueAxisControls } from '../chart-radar-settings-controls'
import { convertChartType } from '../chart-type-conversion'
import { validateChartNamedSeriesBindings } from '../chart-series-selector'
import { chartParetoSettingsCategories } from '../chart-pareto'
import { chartHistogramCategories } from '../chart-histogram'
import { chartHistogramDataControls } from '../chart-histogram-settings-controls'
import { chartAutoGradientControls, chartPlotAreaControls } from '../chart-plot-area-controls'
import { supportsChartPlotArea, validateChartPlotArea } from '../chart-plot-area'
import ChartOptionFields from './ChartOptionFields.vue'
import ChartSeriesSettings from './ChartSeriesSettings.vue'
import ChartParetoSettings from './ChartParetoSettings.vue'
import ChartWaterfallSettings from './ChartWaterfallSettings.vue'
import ChartPieSliceSettings from './ChartPieSliceSettings.vue'
import ChartCandlestickSettings from './ChartCandlestickSettings.vue'
import ChartTreemapSettings from './ChartTreemapSettings.vue'
import ChartSunburstSettings from './ChartSunburstSettings.vue'
import ChartChordSettings from './ChartChordSettings.vue'
import ChartWordCloudSettings from './ChartWordCloudSettings.vue'
import ChartChordMappingFields from './ChartChordMappingFields.vue'
import { chartChordSettingsMapping } from '../chart-chord-settings-controls'
import ChartHierarchyMappingFields from './ChartHierarchyMappingFields.vue'
import ChartStatisticsSettings from './ChartStatisticsSettings.vue'
import ChartCollectionSettings from './ChartCollectionSettings.vue'
import { chartEditPanelCapabilities, type ChartEditPanelNavigation, type ChartEditPanelTab } from '../chart-edit-panel-route'

const props = defineProps<{ draft: ChartSettingsDraft; sheets?: ChartSettingsSourceSheet[]; hostSheetId?: string; data?: ChartData; projection?: ChartSettingsProjection; date1904?: boolean; lockedSource?: boolean; active?: boolean; sessionKey?: number | string; panelTab?: ChartEditPanelTab; navigation?: ChartEditPanelNavigation; navigationKey?: number; navigationReady?: boolean; navigationCurrent?: () => boolean; disabled?: boolean }>()
const emit = defineEmits<{ 'update:draft': [value: ChartSettingsDraft]; 'resource-loading': [value: boolean] }>()
const config = computed(() => props.draft.config), descriptor = computed(() => resolveChartType(config.value.type))
const sourceSheetId = computed(() => config.value.source.sheetId ?? props.hostSheetId)
const sourceSheet = computed(() => props.sheets?.find(sheet => sheet.id === sourceSheetId.value))
const sourceLocked = computed(() => props.lockedSource || Boolean(config.value.pivotTableId))
const projection = computed(() => props.projection ?? projectChartSettings(props.draft, sourceSheet.value, props.data, props.date1904 ?? props.data?.date1904,
  { hostSheetId: props.hostSheetId }))
const dateAxisError = computed(() => chartSettingsDateAxisError(config.value, projection.value.data,
  props.date1904 ?? projection.value.data?.date1904 ?? props.data?.date1904, projection.value.prepared ? 'options' : 'source'))
const namedBindingError = computed(() => {
  try { validateChartNamedSeriesBindings(config.value.options, projection.value.data?.series ?? []); return undefined }
  catch (reason) { return reason instanceof Error ? reason.message : String(reason) }
})
const plotAreaError = computed(() => {
  try { validateChartPlotArea(config.value.options?.plotArea); return undefined }
  catch (reason) { return reason instanceof Error ? reason.message : String(reason) }
})
const plotAreaControls = computed(() => chartPlotAreaControls(projection.value.data?.series))
const fieldCount = computed(() => projection.value.data?.fields?.length ?? projection.value.source?.fields.length ?? (sourceLocked.value || !sourceSheet.value ? props.data?.fields?.length : undefined) ?? chartDraftFieldCount(props.draft))
const fields = computed<ChartSettingsField[]>(() => {
  const sourceFields = projection.value.data?.fields ?? projection.value.source?.fields ?? (sourceLocked.value || !sourceSheet.value ? props.data?.fields : undefined)
  return sourceFields ? sourceFields.map((label, value) => ({ label: label || `字段 ${value + 1}`, value })) : chartDraftFields(props.draft, sourceSheet.value)
})
const categories = computed(() => projection.value.data?.categories ?? [])
const relationFormat = computed(() => chartRelationSourceFormat(config.value, projection.value.source ?? projection.value.data))
const pointCategories = computed(() => {
  if (descriptor.value.family === 'histogram') return chartHistogramCategories(projection.value.statistics?.histogram)
  if (descriptor.value.family === 'pareto') {
    return chartParetoSettingsCategories(projection.value.statistics)
  }
  return descriptor.value.family === 'boxplot' ? projection.value.data?.series.map(series => series.name) ?? [] : projection.value.data?.hierarchy
    ? chartHierarchyItems(projection.value.data.hierarchy).map(item => item.path.join(' / ')) : categories.value
})
const numericFields = computed(() => {
  const numeric = projection.value.data?.sourceValueFields ?? projection.value.data?.sourceNumericFields
    ?? projection.value.source?.sourceValueFields ?? projection.value.source?.sourceNumericFields
  if (numeric) { const values = new Set(numeric); return fields.value.filter(field => values.has(field.value)) }
  if (projection.value.prepared) {
    const values = new Set(projection.value.data?.series.map(series => series.field) ?? [])
    return fields.value.filter(field => values.has(field.value))
  }
  return fields.value.filter(field => projection.value.source?.records.some(record => chartNumber(record[field.value]) !== null))
})
const seriesFields = computed<ChartSettingsField[]>(() => {
  if (config.value.type === 'chord') {
    const mapping = chartChordSettingsMapping(config.value, fieldCount.value), value = mapping.valueIndex ?? mapping.targetIndex
    return [{ value, label: mapping.valueIndex === undefined ? '计数' : fields.value.find(field => field.value === value)?.label ?? '和弦权重' }]
  }
  if (projection.value.data?.series.length) {
    const values = new Set(projection.value.data.series.map(series => series.field))
    return [...values].map(value => fields.value.find(field => field.value === value) ?? { value, label: `字段 ${value + 1}` })
  }
  if (descriptor.value.dataShape === 'gauge') { const value = chartGaugeSettingsMapping(config.value, fieldCount.value).valueIndex; return [{ value, label: fields.value.find(field => field.value === value)?.label ?? '仪表数值' }] }
  if (config.value.valueFields?.length) return config.value.valueFields.map(value => fields.value.find(field => field.value === value) ?? { value, label: `字段 ${value + 1}` })
  if (descriptor.value.dataShape === 'edge') { const value = config.value.options?.mapping?.edge?.valueIndex ?? 2; return [{ value, label: fields.value.find(field => field.value === value)?.label ?? '关系权重' }] }
  if (descriptor.value.dataShape === 'bubble') { const value = config.value.options?.mapping?.bubble?.yIndex ?? 1; return [{ value, label: fields.value.find(field => field.value === value)?.label ?? 'Y 数值' }] }
  return numericFields.value
})
const panelCapabilities = computed(() => chartEditPanelCapabilities(config.value, projection.value.data))
const hasAxes = computed(() => panelCapabilities.value.hasAxes)
const dateAxis = computed(() => supportsChartDateAxis(config.value))
const usesDateAxis = computed(() => dateAxis.value && Boolean(config.value.options?.mapping?.useDateAxis))
const xAxisControls = computed(() => [...(dateAxis.value ? chartDateAxisControls() : []),
  ...chartAxisControls('xAxis', descriptor.value.dataShape === 'xy' || descriptor.value.dataShape === 'bubble' || usesDateAxis.value, usesDateAxis.value)])
const indicator = computed(() => panelCapabilities.value.indicator)
const secondaryAxis = computed(() => panelCapabilities.value.secondaryAxis)
const hasRightAxis = computed(() => panelCapabilities.value.rightAxis)
const secondaryAxisTitle = computed(() => descriptor.value.family === 'bar' ? '顶部数值轴' : '右侧数值轴')
const statistics = computed(() => panelCapabilities.value.statistics)
const statisticsFields = computed(() => seriesFields.value.filter((field, index) => projection.value.data?.series.length
  ? projection.value.data.series.some((series, position) => series.field === field.value && supportsChartStatistics(config.value, field.value, position, series.sourceName ?? series.name))
  : supportsChartStatistics(config.value, field.value, index)))
const typeOptions = chartTypeSelectOptions()
const textFontControls = chartFontControls('appearance.font', item => item.options?.appearance?.font, item => optionGroup(item, 'appearance').font ??= {})
const titleControls = chartTitleFontControls('title')
const subtitleControls = chartTitleFontControls('subtitle')
const legendControls = chartFontControls('legend.font', item => item.legend?.font, chartLegendFont)
const borderControls = chartLineControls('appearance.border', item => item.options?.appearance?.border, item => optionGroup(item, 'appearance').border ??= {})

const root = ref<HTMLElement>()
const openPanels = ref<string[]>(props.panelTab === 'style' ? ['common'] : props.panelTab === 'data' ? ['data'] : ['data', 'common'])
const openFonts = ref<string[]>([]), openAxes = ref<string[]>([]), openSeries = ref<string[]>([])
const composing = ref(false)
let pendingFocus = false, focusGeneration = 0, focusFrame: number | undefined, disposed = false
const navigation = computed(() => props.navigation?.route.tab === props.panelTab ? props.navigation : undefined)
const genericTypeSection = computed(() => descriptor.value.family === 'pie' ? 'pieStyle'
  : ['line', 'area', 'combination'].includes(descriptor.value.family) ? 'lineAndArea' : descriptor.value.family)
function visiblePanel(key: string): boolean { return !props.panelTab || (key === 'data' || key === 'aggregation') === (props.panelTab === 'data') }
function collapseKeys(value: unknown): string[] { return (Array.isArray(value) ? value : [value]).filter((key): key is string => typeof key === 'string') }
function cancelFocus(): void { ++focusGeneration; if (focusFrame !== undefined) cancelAnimationFrame(focusFrame); focusFrame = undefined }
function focusCurrent(): boolean { return !disposed && props.active !== false && !props.disabled && !composing.value && !!props.navigationReady && !!navigation.value && props.navigationCurrent?.() !== false }
function availableControl(region: Element | undefined): HTMLElement | undefined {
  return region && Array.from(region.querySelectorAll<HTMLElement>('input, textarea, button, [role="combobox"], [tabindex="0"]'))
    .find(node => !node.matches(':disabled, [aria-disabled="true"]') && !node.closest('[aria-disabled="true"]') && node.getClientRects().length > 0)
}
function navigationControl(): HTMLElement | undefined {
  const element = root.value, plan = navigation.value
  if (!element || !plan) return undefined
  if (plan.waterfallRole) {
    const region = Array.from(element.querySelectorAll<HTMLElement>('[data-chart-waterfall-role]'))
      .find(node => node.dataset.chartWaterfallRole === plan.waterfallRole)
    const source = region?.querySelector<HTMLElement>('[data-chart-route-series]')
    if (!source || source.dataset.chartSelectedSeries !== (plan.seriesTarget ?? 'all')) return undefined
  }
  if (plan.pieSourcePointIndex !== undefined && !Array.from(element.querySelectorAll<HTMLElement>('[data-chart-pie-source-point]'))
    .some(node => node.dataset.chartPieSourcePoint === String(plan.pieSourcePointIndex))) return undefined
  if (plan.seriesTarget) {
    for (const region of element.querySelectorAll<HTMLElement>('[data-chart-route-series]')) {
      if (region.dataset.chartSelectedSeries !== plan.seriesTarget) continue
      if (plan.optionKey) {
        const option = Array.from(element.querySelectorAll<HTMLElement>('[data-chart-option-key]')).find(node => node.dataset.chartOptionKey === plan.optionKey)
        const focused = availableControl(option)
        if (focused) return focused
        if (plan.route.focus) return undefined
      }
      const control = availableControl(region)
      if (control) return control
    }
    return undefined
  }
  if (plan.statisticsKind) return Array.from(element.querySelectorAll<HTMLElement>('[data-chart-statistics-add]')).find(node => node.dataset.chartStatisticsAdd === plan.statisticsKind)
  if (plan.paretoGroup) return availableControl(Array.from(element.querySelectorAll<HTMLElement>('[data-chart-pareto-group]')).find(node => node.dataset.chartParetoGroup === plan.paretoGroup))
  if (plan.route.tab === 'data') {
    const control = availableControl(element.querySelector('[data-chart-source-range]') ?? undefined)
    if (control) return control
  }
  if (plan.optionKey) {
    for (const option of element.querySelectorAll<HTMLElement>('[data-chart-option-key]')) {
      if (option.dataset.chartOptionKey !== plan.optionKey && !option.dataset.chartOptionKey?.startsWith(`${plan.optionKey}.`)) continue
      const control = availableControl(option)
      if (control) return control
    }
  }
  if (plan.route.focus) return undefined
  for (const section of element.querySelectorAll<HTMLElement>('[data-chart-edit-section]')) if (section.dataset.chartEditSection === plan.route.section) {
    const control = availableControl(section)
    if (control) return control
  }
  for (const group of element.querySelectorAll<HTMLElement>('[data-chart-settings-group]')) if (plan.panels.includes(group.dataset.chartSettingsGroup ?? '')) {
    const control = availableControl(group)
    if (control) return control
  }
}
function scheduleFocus(): void {
  if (!pendingFocus) return
  if (composing.value) { pendingFocus = false; return }
  if (!focusCurrent()) return
  const ticket = ++focusGeneration
  void nextTick(() => {
    if (ticket !== focusGeneration || !pendingFocus || !focusCurrent()) return
    focusFrame = requestAnimationFrame(() => {
      focusFrame = undefined
      if (ticket !== focusGeneration || !pendingFocus || !focusCurrent()) return
      const ownDialog = root.value?.closest('[role="dialog"]'), activeDialog = document.activeElement?.closest('[role="dialog"]')
      if (activeDialog && activeDialog !== ownDialog) { pendingFocus = false; return }
      const control = navigationControl()
      if (!control || control.matches(':disabled, [aria-disabled="true"]') || !control.getClientRects().length) return
      pendingFocus = false; control.focus({ preventScroll: true }); control.closest('.ant-form-item')?.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'auto' })
    })
  })
}
watch(navigation, plan => {
  cancelFocus(); pendingFocus = !!plan
  if (!plan) return
  openPanels.value = [...plan.panels]; openFonts.value = [...plan.fonts]; openAxes.value = [...plan.axes]; openSeries.value = [...plan.series]
}, { immediate: true, flush: 'sync' })
watch([navigation, () => props.navigationReady, () => props.active, () => props.disabled, projection], scheduleFocus, { immediate: true, flush: 'post' })
function compositionStart(): void { composing.value = true; pendingFocus = false; cancelFocus() }
onScopeDispose(() => { disposed = true; cancelFocus() })

function change(value: ChartConfig): void { emit('update:draft', { ...props.draft, config: value }) }
function sourceText(value: unknown): void { if (typeof value === 'string') emit('update:draft', { ...props.draft, sourceText: value }) }
function sourceSheetChange(value: unknown): void { if (typeof value === 'string') { const next = cloneChartDraftConfig(config.value); next.source.sheetId = value; change(next) } }
function orientation(value: unknown): void { if (value === 'columns' || value === 'rows') { const next = cloneChartDraftConfig(config.value); next.source.orientation = value; change(next) } }
function chartType(value: unknown): void { if (isChartType(value)) change(convertChartType(config.value, value, projection.value.source)) }
</script>

<template>
  <div ref="root" class="chart-settings-fields" :data-chart-route-tab="panelTab" :data-chart-route-section="navigation?.route.section" @compositionstart.capture="compositionStart" @compositionend.capture="composing = false">
    <FormItem label="图表类型"><Select :value="descriptor.value" :options="typeOptions" show-search option-filter-prop="label" aria-label="图表类型" :disabled="disabled" @update:value="chartType" /></FormItem>
    <Alert type="info" :message="chartDataHint(config.type)" show-icon />
    <Alert v-if="projection.error" type="warning" :message="projection.error" show-icon />
    <Alert v-if="namedBindingError" type="warning" :message="namedBindingError" show-icon />
    <Alert v-if="dateAxisError" type="warning" :message="dateAxisError" show-icon />
    <Alert v-if="plotAreaError" type="warning" :message="plotAreaError" show-icon />
    <Collapse :active-key="openPanels" @update:active-key="openPanels = collapseKeys($event)">
      <CollapsePanel v-if="visiblePanel('data')" key="data" header="数据源与字段">
        <div data-chart-settings-group="data">
        <FormItem v-if="sheets?.length" label="数据所在工作表"><Select :value="sourceSheetId" :options="sheets.map(sheet => ({ value: sheet.id, label: sheet.name }))" aria-label="图表数据工作表" :disabled="sourceLocked || disabled" @update:value="sourceSheetChange" /></FormItem>
        <FormItem label="数据区域" data-chart-source-range><Input :value="draft.sourceText" placeholder="A1:C4" aria-label="图表数据区域" :disabled="sourceLocked || disabled" @update:value="sourceText" /></FormItem>
        <FormItem label="系列方向"><Select :value="config.source.orientation" :options="[{ value: 'columns', label: '按列，首行为字段名' }, { value: 'rows', label: '按行，首列为字段名' }]" aria-label="图表系列方向" :disabled="sourceLocked || disabled" @update:value="orientation" /></FormItem>
        <ChartHierarchyMappingFields v-if="descriptor.dataShape === 'hierarchy'" :config="config" :fields="fields" :numeric-fields="numericFields" :field-count="fieldCount" :disabled="disabled" @update:config="change" />
        <ChartChordMappingFields v-else-if="config.type === 'chord'" :config="config" :fields="fields" :numeric-fields="numericFields" :field-count="fieldCount" :disabled="disabled" @update:config="change" />
        <ChartOptionFields v-else :config="config" :controls="chartDataMappingControls(config, fieldCount, projection.data, projection.source)" :fields="fields" :field-count="fieldCount" :disabled="disabled" @update:config="change" />
        <Alert v-if="config.type === 'relation' && !config.options?.mapping?.relationFormat" type="info" :message="relationFormat === 'coWordMatrix' ? '当前识别为共词矩阵，行节点与矩阵列标题对应。' : '当前使用来源、目标和权重数据。'" show-icon />
        <ChartOptionFields v-if="config.type === 'histogram'" :config="config" :controls="chartHistogramDataControls(config)" :disabled="disabled" @update:config="change" />
        <Alert v-if="sourceLocked" type="info" message="数据区域由关联的数据透视表管理，字段与样式在此编辑。" />
        </div>
      </CollapsePanel>
      <CollapsePanel v-if="visiblePanel('common')" key="common" header="标题、图例与外观"><div data-chart-settings-group="common"><ChartOptionFields :config="config" :controls="chartCommonControls(config)" :disabled="disabled" @update:config="change" /></div></CollapsePanel>
      <CollapsePanel v-if="visiblePanel('fonts')" key="fonts" header="字体与边框">
        <Collapse :active-key="openFonts" @update:active-key="openFonts = collapseKeys($event)">
          <CollapsePanel key="text" header="图表字体"><ChartOptionFields :config="config" :controls="textFontControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="title" header="标题字体"><ChartOptionFields :config="config" :controls="titleControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="subtitle" header="副标题字体"><ChartOptionFields :config="config" :controls="subtitleControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel v-if="chartHasLegend(config.type)" key="legend" header="图例字体"><ChartOptionFields :config="config" :controls="legendControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="border" header="图表边框"><ChartOptionFields :config="config" :controls="borderControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
        </Collapse>
      </CollapsePanel>
      <CollapsePanel v-if="visiblePanel('type')" key="type" header="图型专属设置">
        <div data-chart-settings-group="type">
        <ChartOptionFields :config="config" :controls="chartAutoGradientControls()" :disabled="disabled" @update:config="change" />
        <div :data-chart-edit-section="genericTypeSection"><ChartOptionFields :config="config" :controls="chartTypeSettingsControls(config, fieldCount, projection.data?.series[0]?.sourceName ?? projection.data?.series[0]?.name, projection.data)" :categories="categories" :disabled="disabled" @update:config="change" /></div>
        <ChartPieSliceSettings v-if="descriptor.family === 'pie' && !descriptor.threeD" :config="config" :data="projection.data" :route-point-index="navigation?.pieSourcePointIndex" :route-key="navigationKey" :disabled="disabled" @update:config="change" />
        <ChartParetoSettings v-if="config.type === 'pareto'" :config="config" :data="projection.data" :statistics="projection.statistics" :route-group="navigation?.paretoGroup" :route-key="navigationKey" :disabled="disabled" @update:config="change" />
        <div v-if="config.type === 'waterfall'" data-chart-edit-section="waterfallSeries"><ChartWaterfallSettings :config="config" :fields="seriesFields" :categories="categories" :data="projection.data" :route-series-target="navigation?.seriesTarget" :route-role="navigation?.waterfallRole" :route-key="navigationKey" :disabled="disabled" @update:config="change" /></div>
        <div v-if="config.type === 'candlestick'" data-chart-edit-section="candlestick"><ChartCandlestickSettings :config="config" :data="projection.data" :disabled="disabled" @update:config="change" /></div>
        <div v-if="config.type === 'treemap'" data-chart-edit-section="treemap"><ChartTreemapSettings :config="config" :disabled="disabled" @update:config="change" /></div>
        <div v-if="config.type === 'sunburst'" data-chart-edit-section="sunburst"><ChartSunburstSettings :config="config" :disabled="disabled" @update:config="change" /></div>
        <div v-if="config.type === 'chord'" data-chart-edit-section="chord"><ChartChordSettings :config="config" :disabled="disabled" @update:config="change" /></div>
        <div v-if="config.type === 'wordCloud'" data-chart-edit-section="wordCloud"><ChartWordCloudSettings :config="config" :active="active" :session-key="sessionKey" :disabled="disabled" @update:config="change" @resource-loading="emit('resource-loading', $event)" /></div>
        </div>
      </CollapsePanel>
      <CollapsePanel v-if="visiblePanel('plot-area') && supportsChartPlotArea(config.type)" key="plot-area" header="绘图区"><ChartOptionFields :config="config" :controls="plotAreaControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
      <CollapsePanel v-if="visiblePanel('series')" key="series" header="系列、标记与数据标签">
        <div data-chart-settings-group="series">
        <Collapse :active-key="openSeries" @update:active-key="openSeries = collapseKeys($event)">
          <CollapsePanel key="all" header="所有系列"><ChartOptionFields :config="config" :controls="chartSeriesControls(undefined, descriptor.family === 'combination', secondaryAxis, descriptor.family === 'bar', undefined, config.type === 'bubble')" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="one" header="单独设置系列"><ChartSeriesSettings :config="config" :fields="seriesFields" :source-fields="fields" :categories="pointCategories" :data="projection.data" :route-series-target="navigation?.seriesTarget" :route-key="navigationKey" :combination="descriptor.family === 'combination'" :secondary-axis="secondaryAxis" :horizontal="descriptor.family === 'bar'" :scatter="descriptor.family === 'scatter'" :disabled="disabled" @update:config="change" /></CollapsePanel>
        </Collapse>
        </div>
      </CollapsePanel>
      <CollapsePanel v-if="visiblePanel('axes') && hasAxes" key="axes" header="坐标轴">
        <div data-chart-settings-group="axes">
        <Collapse :active-key="openAxes" @update:active-key="openAxes = collapseKeys($event)">
          <CollapsePanel key="x" :header="descriptor.family === 'radar' ? '指标轴 / 分类标签' : descriptor.family === 'heatmap' ? 'X 轴 / 数值字段名' : 'X 轴 / 分类轴'"><ChartOptionFields :config="config" :controls="descriptor.family === 'radar' ? chartRadarCategoryAxisControls() : xAxisControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="y" :header="descriptor.family === 'radar' ? '径向数值轴' : descriptor.family === 'heatmap' ? 'Y 轴 / 分类标签' : 'Y 轴 / 数值轴'"><ChartOptionFields :config="config" :controls="descriptor.family === 'radar' ? chartRadarValueAxisControls() : chartAxisControls('yAxis', descriptor.family !== 'heatmap')" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel v-if="hasRightAxis" key="right" :header="secondaryAxisTitle"><ChartOptionFields :config="config" :controls="chartAxisControls('rightYAxis')" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel v-if="descriptor.threeD" key="depth" header="系列深度轴"><ChartOptionFields :config="config" :controls="chartAxisControls('depthAxis', false)" :disabled="disabled" @update:config="change" /></CollapsePanel>
        </Collapse>
        </div>
      </CollapsePanel>
      <CollapsePanel v-if="visiblePanel('indicator') && indicator" key="indicator" header="指示线与坐标标签"><div data-chart-settings-group="indicator"><ChartOptionFields :config="config" :controls="chartAxisPointerControls()" :disabled="disabled" @update:config="change" /></div></CollapsePanel>
      <CollapsePanel v-if="visiblePanel('statistics') && statistics" key="statistics" header="趋势线与误差线"><div data-chart-settings-group="statistics"><ChartStatisticsSettings :config="config" :series-fields="statisticsFields" :source-fields="fields" :numeric-x="supportsChartErrorDirection(config, 'x')" :route-kind="navigation?.statisticsKind" :route-key="navigationKey" :disabled="disabled" @update:config="change" /></div></CollapsePanel>
      <CollapsePanel v-if="visiblePanel('aggregation') && descriptor.dataShape === 'category'" key="aggregation" header="分类聚合与前 N 项"><ChartOptionFields :config="config" :controls="chartAggregationControls(config)" :disabled="disabled" @update:config="change" /></CollapsePanel>
      <CollapsePanel v-if="visiblePanel('collections')" key="collections" header="配色与项目设置"><ChartCollectionSettings :config="config" :series-fields="seriesFields" :categories="pointCategories" :disabled="disabled" @update:config="change" /></CollapsePanel>
    </Collapse>
  </div>
</template>

<style scoped>
.chart-settings-fields { display: grid; gap: 12px; min-width: 0; }
.chart-settings-fields :deep(.ant-form-item) { margin-bottom: 10px; }
.chart-settings-fields :deep(.ant-collapse-content-box) { min-width: 0; }
</style>
