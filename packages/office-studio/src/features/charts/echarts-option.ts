import type {
  BarSeriesOption,
  LineSeriesOption,
  PieSeriesOption,
  RadarSeriesOption,
  ScatterSeriesOption,
} from 'echarts/charts'
import type {
  GridComponentOption,
  LegendComponentOption,
  RadarComponentOption,
  TitleComponentOption,
  TooltipComponentOption,
} from 'echarts/components'
import type { ComposeOption } from 'echarts/core'
import type { XAxisComponentOption, YAxisComponentOption } from 'echarts'
import type { ChartAxisOptions, ChartModel, WorkbookModel } from '../../domain/model'
import { radarValueExtent, readChartData } from './chart-engine'

export type ChartEChartsOption = ComposeOption<
  | BarSeriesOption
  | LineSeriesOption
  | PieSeriesOption
  | RadarSeriesOption
  | ScatterSeriesOption
  | GridComponentOption
  | LegendComponentOption
  | RadarComponentOption
  | TitleComponentOption
  | TooltipComponentOption
  | XAxisComponentOption
  | YAxisComponentOption
>

export interface ChartEChartsTheme {
  backgroundColor?: string
  textColor?: string
  axisColor?: string
  splitLineColor?: string
  palette?: string[]
}

const DEFAULT_PALETTE = ['#2563eb', '#dc2626', '#059669', '#d97706', '#7c3aed', '#0891b2']

/** Converts the editable workbook chart model to an ECharts rendering option. */
export function chartEChartsOption(
  workbook: WorkbookModel,
  chart: ChartModel,
  theme: ChartEChartsTheme = {},
): ChartEChartsOption {
  const data = readChartData(workbook, chart)
  const color = theme.textColor ?? '#334155'
  const axisColor = theme.axisColor ?? '#94a3b8'
  const splitLineColor = theme.splitLineColor ?? '#e2e8f0'
  const palette = [...(theme.palette?.length ? theme.palette : DEFAULT_PALETTE)]
  const isPie = chart.type === 'pie' || chart.type === 'doughnut'
  const isScatter = chart.type === 'scatter'
  const isRadar = chart.type === 'radar'
  const seriesNames = data.series.map(series => series.name)
  const titleHeight = chart.title ? 36 : 0
  const legendHeight = chart.legend.visible && (chart.legend.position === 'top' || chart.legend.position === 'bottom') ? 28 : 0
  const sideLegendWidth = chart.legend.visible && (chart.legend.position === 'left' || chart.legend.position === 'right') ? Math.min(96, chart.size.width * 0.24) : 0
  const grid = {
    left: 56 + (chart.legend.position === 'left' ? sideLegendWidth : 0),
    right: 20 + (chart.legend.position === 'right' ? sideLegendWidth : 0),
    top: 16 + titleHeight + (chart.legend.position === 'top' ? legendHeight : 0),
    bottom: 42 + (chart.legend.position === 'bottom' ? legendHeight : 0),
    containLabel: true,
  }
  const title: TitleComponentOption[] = []
  if (chart.title) title.push({ text: chart.title, left: 'center', top: 8, textStyle: { color } })
  if (isRadar && chart.xAxis.title) title.push({
    text: chart.xAxis.title, left: 'center', bottom: grid.bottom - 30, textStyle: { color, fontSize: 11, fontWeight: 'normal' },
  })
  if (isRadar && chart.yAxis.title) title.push({
    text: chart.yAxis.title, left: grid.left - 44, top: grid.top, textStyle: { color, fontSize: 11, fontWeight: 'normal' },
  })

  const option: ChartEChartsOption = {
    animation: false,
    backgroundColor: theme.backgroundColor ?? 'transparent',
    color: palette,
    textStyle: { color },
    ...(title.length ? { title: title.length === 1 ? title[0] : title } : {}),
    legend: {
      show: chart.legend.visible,
      data: isPie ? data.categories : seriesNames,
      ...legendPosition(chart.legend.position, titleHeight),
      textStyle: { color },
    },
    tooltip: { trigger: isPie || isScatter || isRadar ? 'item' : 'axis' },
    ...(isPie || isRadar ? {} : {
      grid,
      xAxis: isScatter
        ? numericAxis(chart.xAxis, axisColor, splitLineColor, 'x')
        : categoryAxis(data.categories, chart.xAxis, axisColor, splitLineColor),
      yAxis: numericAxis(chart.yAxis, axisColor, splitLineColor, 'y'),
    }),
    ...(isRadar ? {
      radar: radarComponent(data, chart, grid, color, axisColor, splitLineColor),
    } : {}),
    series: isPie
      ? [pieSeries(data, chart, palette, color, grid)]
      : isScatter
        ? scatterSeries(data, chart)
        : isRadar
          ? radarSeries(data)
          : categorySeries(data, chart),
  }

  return option
}

function radarComponent(
  data: ReturnType<typeof readChartData>,
  chart: ChartModel,
  grid: { left: number; right: number; top: number; bottom: number },
  color: string,
  axisColor: string,
  splitLineColor: string,
): RadarComponentOption {
  const extent = radarValueExtent(data)
  const availableWidth = chart.size.width - grid.left - grid.right
  const availableHeight = chart.size.height - grid.top - grid.bottom
  return {
    center: [(chart.size.width + grid.left - grid.right) / 2, (chart.size.height + grid.top - grid.bottom) / 2],
    radius: Math.max(1, Math.min(availableWidth, availableHeight) * 0.37),
    clockwise: true,
    indicator: data.categories.map(name => ({ name, ...extent })),
    axisName: { color },
    axisLine: { lineStyle: { color: axisColor } },
    axisLabel: { show: true, color: axisColor },
    splitLine: { lineStyle: { color: splitLineColor } },
    splitArea: { show: false },
  }
}

function radarSeries(data: ReturnType<typeof readChartData>): RadarSeriesOption[] {
  return data.series.map(series => ({
    type: 'radar',
    name: series.name,
    symbol: 'none',
    data: [{ name: series.name, value: series.values }],
  }))
}

function legendPosition(position: ChartModel['legend']['position'], titleHeight: number): Pick<LegendComponentOption, 'top' | 'right' | 'bottom' | 'left' | 'orient'> {
  switch (position) {
    case 'top': return { top: 8 + titleHeight, left: 'center', orient: 'horizontal' }
    case 'right': return { right: 8, top: 'middle', orient: 'vertical' }
    case 'left': return { left: 8, top: 'middle', orient: 'vertical' }
    case 'bottom': return { bottom: 8, left: 'center', orient: 'horizontal' }
  }
}

function categoryAxis(
  categories: string[],
  axis: ChartAxisOptions,
  axisColor: string,
  splitLineColor: string,
): XAxisComponentOption {
  return {
    type: 'category',
    data: categories,
    name: axis.title,
    nameLocation: 'middle',
    nameGap: 28,
    axisLine: { lineStyle: { color: axisColor } },
    axisLabel: { color: axisColor },
    splitLine: { show: false, lineStyle: { color: splitLineColor } },
  }
}

function numericAxis(
  axis: ChartAxisOptions,
  axisColor: string,
  splitLineColor: string,
  dimension: 'x',
): XAxisComponentOption
function numericAxis(
  axis: ChartAxisOptions,
  axisColor: string,
  splitLineColor: string,
  dimension: 'y',
): YAxisComponentOption
function numericAxis(
  axis: ChartAxisOptions,
  axisColor: string,
  splitLineColor: string,
  dimension: 'x' | 'y',
): XAxisComponentOption | YAxisComponentOption {
  return {
    type: axis.scaleType === 'log' ? 'log' : 'value',
    ...(axis.scaleType === 'log' ? { logBase: axis.logBase ?? 10 } : {}),
    name: axis.title,
    nameLocation: 'middle',
    nameGap: dimension === 'x' ? 28 : 42,
    axisLine: { lineStyle: { color: axisColor } },
    axisLabel: { color: axisColor },
    splitLine: { show: dimension === 'y', lineStyle: { color: splitLineColor } },
  }
}

function pieSeries(
  data: ReturnType<typeof readChartData>,
  chart: ChartModel,
  palette: string[],
  color: string,
  grid: { left: number; right: number; top: number; bottom: number },
): PieSeriesOption {
  const source = data.series[0]
  return {
    type: 'pie',
    name: source?.name ?? '',
    radius: chart.type === 'doughnut' ? ['52%', '74%'] : '74%',
    center: ['50%', '50%'],
    left: grid.left - 40,
    right: grid.right,
    top: grid.top,
    bottom: grid.bottom - 26,
    avoidLabelOverlap: true,
    stillShowZeroSum: false,
    showEmptyCircle: false,
    label: { color },
    data: data.categories.map((name, index) => ({
      name,
      value: Math.max(0, source?.values[index] ?? 0),
      itemStyle: { color: palette[index % palette.length] },
    })),
  }
}

function scatterSeries(data: ReturnType<typeof readChartData>, chart: ChartModel): ScatterSeriesOption[] {
  return data.series.map(series => ({
    type: 'scatter',
    name: series.name,
    data: series.values.flatMap((value, index) => {
      const x = series.xValues?.[index]
      if (value === null || x == null || !Number.isFinite(value) || !Number.isFinite(x)) return []
      if (chart.xAxis.scaleType === 'log' && x <= 0) return []
      if (chart.yAxis.scaleType === 'log' && value <= 0) return []
      return [[x, value]]
    }),
    xAxisIndex: 0,
    yAxisIndex: 0,
  }))
}

function categorySeries(
  data: ReturnType<typeof readChartData>,
  chart: ChartModel,
): Array<BarSeriesOption | LineSeriesOption> {
  return data.series.map(series => {
    const values = series.values.map(value => value === null || (chart.yAxis.scaleType === 'log' && value <= 0) ? null : value)
    if (chart.type === 'bar' || chart.type === 'stacked-column') {
      return {
        type: 'bar',
        name: series.name,
        data: values,
        ...(chart.type === 'stacked-column' ? { stack: 'total' } : {}),
      }
    }
    return {
      type: 'line',
      name: series.name,
      data: values,
      showSymbol: true,
      connectNulls: false,
      ...(chart.type === 'area' ? { areaStyle: { opacity: 0.28 } } : {}),
    }
  })
}
