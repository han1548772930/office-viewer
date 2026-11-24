<script setup lang="ts">
import { computed } from 'vue'
import { Alert, Collapse, CollapsePanel, FormItem, Input, Select } from 'antdv-next'
import type { ChartConfig, ChartData } from '../chart-engine'
import { chartDataHint, chartTypeSelectOptions, isChartType, resolveChartType, supportsChartErrorDirection, supportsChartStatistics } from '../chart-catalog'
import { chartAxisControls, chartAxisPointerControls, chartCommonControls, chartDateAxisControls, chartFontControls, chartLegendFont, chartLineControls, chartSeriesControls, optionGroup } from '../chart-settings-controls'
import { chartAggregationControls, chartDataMappingControls, chartTypeSettingsControls } from '../chart-type-settings-controls'
import { chartDraftFieldCount, chartDraftFields, cloneChartDraftConfig, type ChartSettingsDraft, type ChartSettingsField, type ChartSettingsSourceSheet } from '../chart-settings-draft'
import { chartSettingsDateAxisError, projectChartSettings } from '../chart-settings-data'
import { supportsChartDateAxis } from '../chart-date-axis'
import { supportsChartAxisPointer } from '../chart-axis-pointer'
import { chartHierarchyItems, chartNumber } from '../chart-data'
import { chartGaugeSettingsMapping } from '../chart-gauge-settings-controls'
import { chartRadarCategoryAxisControls, chartRadarValueAxisControls } from '../chart-radar-settings-controls'
import { convertChartType } from '../chart-type-conversion'
import { chartSeriesStyle } from '../chart-options'
import { validateChartNamedSeriesBindings } from '../chart-series-selector'
import { chartParetoSettingsCategories } from '../chart-pareto'
import { chartWaterfallUsesSecondaryAxis } from '../chart-waterfall'
import { chartAutoGradientControls, chartPlotAreaControls } from '../chart-plot-area-controls'
import { supportsChartPlotArea, validateChartPlotArea } from '../chart-plot-area'
import ChartOptionFields from './ChartOptionFields.vue'
import ChartSeriesSettings from './ChartSeriesSettings.vue'
import ChartParetoSettings from './ChartParetoSettings.vue'
import ChartWaterfallSettings from './ChartWaterfallSettings.vue'
import ChartStatisticsSettings from './ChartStatisticsSettings.vue'
import ChartCollectionSettings from './ChartCollectionSettings.vue'

const props = defineProps<{ draft: ChartSettingsDraft; sheets?: ChartSettingsSourceSheet[]; hostSheetId?: string; data?: ChartData; date1904?: boolean; lockedSource?: boolean; disabled?: boolean }>()
const emit = defineEmits<{ 'update:draft': [value: ChartSettingsDraft] }>()
const config = computed(() => props.draft.config), descriptor = computed(() => resolveChartType(config.value.type))
const sourceSheetId = computed(() => config.value.source.sheetId ?? props.hostSheetId)
const sourceSheet = computed(() => props.sheets?.find(sheet => sheet.id === sourceSheetId.value))
const sourceLocked = computed(() => props.lockedSource || Boolean(config.value.pivotTableId))
const projection = computed(() => projectChartSettings(props.draft, sourceSheet.value, props.data, props.date1904 ?? props.data?.date1904))
const dateAxisError = computed(() => chartSettingsDateAxisError(config.value, projection.value.data, props.date1904 ?? props.data?.date1904))
const namedBindingError = computed(() => {
  try { validateChartNamedSeriesBindings(config.value.options, projection.value.data?.series ?? []); return undefined }
  catch (reason) { return reason instanceof Error ? reason.message : String(reason) }
})
const plotAreaError = computed(() => {
  try { validateChartPlotArea(config.value.options?.plotArea); return undefined }
  catch (reason) { return reason instanceof Error ? reason.message : String(reason) }
})
const plotAreaControls = computed(() => chartPlotAreaControls(projection.value.data?.series))
const fieldCount = computed(() => projection.value.source?.fields.length ?? (sourceLocked.value || !sourceSheet.value ? props.data?.fields?.length : undefined) ?? chartDraftFieldCount(props.draft))
const fields = computed<ChartSettingsField[]>(() => {
  const sourceFields = projection.value.source?.fields ?? (sourceLocked.value || !sourceSheet.value ? props.data?.fields : undefined)
  return sourceFields ? sourceFields.map((label, value) => ({ label: label || `字段 ${value + 1}`, value })) : chartDraftFields(props.draft, sourceSheet.value)
})
const categories = computed(() => projection.value.data?.categories ?? [])
const pointCategories = computed(() => {
  if (descriptor.value.family === 'pareto') {
    try { return chartParetoSettingsCategories(projection.value.data, config.value.options?.pareto?.includeZeroValues, config.value) }
    catch { return [] }
  }
  return descriptor.value.family === 'boxplot' ? projection.value.data?.series.map(series => series.name) ?? [] : projection.value.data?.hierarchy
    ? chartHierarchyItems(projection.value.data.hierarchy).map(item => item.path.join(' / ')) : categories.value
})
const numericFields = computed(() => fields.value.filter(field => projection.value.source?.records.some(record => chartNumber(record[field.value]) !== null)))
const seriesFields = computed<ChartSettingsField[]>(() => {
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
const hasAxes = computed(() => ['line', 'column', 'bar', 'area', 'scatter', 'bubble', 'combination', 'waterfall', 'pareto', 'heatmap', 'boxplot', 'candlestick', 'histogram', 'radar'].includes(descriptor.value.family))
const dateAxis = computed(() => supportsChartDateAxis(config.value))
const usesDateAxis = computed(() => dateAxis.value && Boolean(config.value.options?.mapping?.useDateAxis))
const xAxisControls = computed(() => [...(dateAxis.value ? chartDateAxisControls() : []),
  ...chartAxisControls('xAxis', descriptor.value.dataShape === 'xy' || descriptor.value.dataShape === 'bubble' || usesDateAxis.value, usesDateAxis.value)])
const indicator = computed(() => supportsChartAxisPointer(config.value))
const secondaryAxis = computed(() => !descriptor.value.threeD && ['line', 'column', 'bar', 'area', 'scatter', 'bubble', 'combination', 'waterfall', 'boxplot'].includes(descriptor.value.family))
const hasRightAxis = computed(() => descriptor.value.family === 'pareto' || chartWaterfallUsesSecondaryAxis(config.value, projection.value.data?.series) || secondaryAxis.value && (projection.value.data
  ? projection.value.data.series.some((series, index) => chartSeriesStyle(config.value, series.field, index, undefined, series.sourceName ?? series.name).rightAxis)
  : config.value.options?.allSeriesStyle?.rightAxis || Object.values(config.value.options?.series ?? {}).some(style => style.rightAxis) || config.value.options?.namedSeries?.some(entry => entry.style.rightAxis)))
const secondaryAxisTitle = computed(() => descriptor.value.family === 'bar' ? '顶部数值轴' : '右侧数值轴')
const statistics = computed(() => supportsChartStatistics(config.value))
const statisticsFields = computed(() => seriesFields.value.filter((field, index) => projection.value.data?.series.length
  ? projection.value.data.series.some((series, position) => series.field === field.value && supportsChartStatistics(config.value, field.value, position, series.sourceName ?? series.name))
  : supportsChartStatistics(config.value, field.value, index)))
const typeOptions = chartTypeSelectOptions()
const textFontControls = chartFontControls('appearance.font', item => item.options?.appearance?.font, item => optionGroup(item, 'appearance').font ??= {})
const titleControls = chartFontControls('titleStyle', item => item.options?.titleStyle, item => optionGroup(item, 'titleStyle'))
const subtitleControls = chartFontControls('subtitleStyle', item => item.options?.subtitleStyle, item => optionGroup(item, 'subtitleStyle'))
const legendControls = chartFontControls('legend.font', item => item.legend?.font, chartLegendFont)
const borderControls = chartLineControls('appearance.border', item => item.options?.appearance?.border, item => optionGroup(item, 'appearance').border ??= {})

function change(value: ChartConfig): void { emit('update:draft', { ...props.draft, config: value }) }
function sourceText(value: unknown): void { if (typeof value === 'string') emit('update:draft', { ...props.draft, sourceText: value }) }
function sourceSheetChange(value: unknown): void { if (typeof value === 'string') { const next = cloneChartDraftConfig(config.value); next.source.sheetId = value; change(next) } }
function orientation(value: unknown): void { if (value === 'columns' || value === 'rows') { const next = cloneChartDraftConfig(config.value); next.source.orientation = value; change(next) } }
function chartType(value: unknown): void { if (isChartType(value)) change(convertChartType(config.value, value, projection.value.source)) }
</script>

<template>
  <div class="chart-settings-fields">
    <FormItem label="图表类型"><Select :value="descriptor.value" :options="typeOptions" show-search option-filter-prop="label" aria-label="图表类型" :disabled="disabled" @update:value="chartType" /></FormItem>
    <Alert type="info" :message="chartDataHint(config.type)" show-icon />
    <Alert v-if="projection.error" type="warning" :message="projection.error" show-icon />
    <Alert v-if="namedBindingError" type="warning" :message="namedBindingError" show-icon />
    <Alert v-if="dateAxisError" type="warning" :message="dateAxisError" show-icon />
    <Alert v-if="plotAreaError" type="warning" :message="plotAreaError" show-icon />
    <Collapse :default-active-key="['data', 'common']">
      <CollapsePanel key="data" header="数据源与字段">
        <FormItem v-if="sheets?.length" label="数据所在工作表"><Select :value="sourceSheetId" :options="sheets.map(sheet => ({ value: sheet.id, label: sheet.name }))" aria-label="图表数据工作表" :disabled="sourceLocked || disabled" @update:value="sourceSheetChange" /></FormItem>
        <FormItem label="数据区域"><Input :value="draft.sourceText" placeholder="A1:C4" aria-label="图表数据区域" :disabled="sourceLocked || disabled" @update:value="sourceText" /></FormItem>
        <FormItem label="系列方向"><Select :value="config.source.orientation" :options="[{ value: 'columns', label: '按列，首行为字段名' }, { value: 'rows', label: '按行，首列为字段名' }]" aria-label="图表系列方向" :disabled="sourceLocked || disabled" @update:value="orientation" /></FormItem>
        <ChartOptionFields :config="config" :controls="chartDataMappingControls(config, fieldCount)" :fields="fields" :field-count="fieldCount" :disabled="disabled" @update:config="change" />
        <Alert v-if="sourceLocked" type="info" message="数据区域由关联的数据透视表管理，字段与样式在此编辑。" />
      </CollapsePanel>
      <CollapsePanel key="common" header="标题、图例与外观"><ChartOptionFields :config="config" :controls="chartCommonControls()" :disabled="disabled" @update:config="change" /></CollapsePanel>
      <CollapsePanel key="fonts" header="字体与边框">
        <Collapse>
          <CollapsePanel key="text" header="图表字体"><ChartOptionFields :config="config" :controls="textFontControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="title" header="标题字体"><ChartOptionFields :config="config" :controls="titleControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="subtitle" header="副标题字体"><ChartOptionFields :config="config" :controls="subtitleControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="legend" header="图例字体"><ChartOptionFields :config="config" :controls="legendControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="border" header="图表边框"><ChartOptionFields :config="config" :controls="borderControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
        </Collapse>
      </CollapsePanel>
      <CollapsePanel key="type" header="图型专属设置">
        <ChartOptionFields :config="config" :controls="chartAutoGradientControls()" :disabled="disabled" @update:config="change" />
        <ChartOptionFields :config="config" :controls="chartTypeSettingsControls(config, fieldCount, projection.data?.series[0]?.sourceName ?? projection.data?.series[0]?.name)" :categories="categories" :disabled="disabled" @update:config="change" />
        <ChartParetoSettings v-if="config.type === 'pareto'" :config="config" :data="projection.data" :disabled="disabled" @update:config="change" />
        <ChartWaterfallSettings v-if="config.type === 'waterfall'" :config="config" :fields="seriesFields" :categories="categories" :data="projection.data" :disabled="disabled" @update:config="change" />
      </CollapsePanel>
      <CollapsePanel v-if="supportsChartPlotArea(config.type)" key="plot-area" header="绘图区"><ChartOptionFields :config="config" :controls="plotAreaControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
      <CollapsePanel key="series" header="系列、标记与数据标签">
        <Collapse>
          <CollapsePanel key="all" header="所有系列"><ChartOptionFields :config="config" :controls="chartSeriesControls(undefined, descriptor.family === 'combination', secondaryAxis, descriptor.family === 'bar')" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="one" header="单独设置系列"><ChartSeriesSettings :config="config" :fields="seriesFields" :source-fields="fields" :categories="pointCategories" :data="projection.data" :combination="descriptor.family === 'combination'" :secondary-axis="secondaryAxis" :horizontal="descriptor.family === 'bar'" :scatter="descriptor.family === 'scatter'" :disabled="disabled" @update:config="change" /></CollapsePanel>
        </Collapse>
      </CollapsePanel>
      <CollapsePanel v-if="hasAxes" key="axes" header="坐标轴">
        <Collapse>
          <CollapsePanel key="x" :header="descriptor.family === 'radar' ? '指标轴 / 分类标签' : 'X 轴 / 分类轴'"><ChartOptionFields :config="config" :controls="descriptor.family === 'radar' ? chartRadarCategoryAxisControls() : xAxisControls" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel key="y" :header="descriptor.family === 'radar' ? '径向数值轴' : 'Y 轴 / 数值轴'"><ChartOptionFields :config="config" :controls="descriptor.family === 'radar' ? chartRadarValueAxisControls() : chartAxisControls('yAxis')" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel v-if="hasRightAxis" key="right" :header="secondaryAxisTitle"><ChartOptionFields :config="config" :controls="chartAxisControls('rightYAxis')" :disabled="disabled" @update:config="change" /></CollapsePanel>
          <CollapsePanel v-if="descriptor.threeD" key="depth" header="系列深度轴"><ChartOptionFields :config="config" :controls="chartAxisControls('depthAxis', false)" :disabled="disabled" @update:config="change" /></CollapsePanel>
        </Collapse>
      </CollapsePanel>
      <CollapsePanel v-if="indicator" key="indicator" header="指示线与坐标标签"><ChartOptionFields :config="config" :controls="chartAxisPointerControls()" :disabled="disabled" @update:config="change" /></CollapsePanel>
      <CollapsePanel v-if="statistics" key="statistics" header="趋势线与误差线"><ChartStatisticsSettings :config="config" :series-fields="statisticsFields" :source-fields="fields" :numeric-x="supportsChartErrorDirection(config, 'x')" :disabled="disabled" @update:config="change" /></CollapsePanel>
      <CollapsePanel v-if="descriptor.dataShape === 'category'" key="aggregation" header="分类聚合与前 N 项"><ChartOptionFields :config="config" :controls="chartAggregationControls(config)" :disabled="disabled" @update:config="change" /></CollapsePanel>
      <CollapsePanel key="collections" header="配色与项目设置"><ChartCollectionSettings :config="config" :series-fields="seriesFields" :categories="pointCategories" :disabled="disabled" @update:config="change" /></CollapsePanel>
    </Collapse>
  </div>
</template>

<style scoped>
.chart-settings-fields { display: grid; gap: 12px; min-width: 0; }
.chart-settings-fields :deep(.ant-form-item) { margin-bottom: 10px; }
.chart-settings-fields :deep(.ant-collapse-content-box) { min-width: 0; }
</style>
