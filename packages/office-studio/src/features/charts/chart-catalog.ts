import type { ChartModel, ChartType } from '../../domain/model'
import { chartNamedSeriesStyle } from './chart-series-selector'

export type ChartDataShape = 'category' | 'xy' | 'bubble' | 'edge' | 'hierarchy' | 'distribution' | 'ohlc' | 'heatmap' | 'gauge'
export type ChartFamily = 'line' | 'column' | 'bar' | 'pie' | 'area' | 'radar' | 'scatter' | 'combination'
  | 'wordCloud' | 'funnel' | 'bubble' | 'relation' | 'waterfall' | 'pareto' | 'sankey' | 'heatmap'
  | 'boxplot' | 'candlestick' | 'histogram' | 'treemap' | 'sunburst' | 'gauge' | 'chord'

export interface ChartCatalogEntry {
  /** Fixed v1.6.0 CHART_TYPE_CATALOG label identity; distinct from legacy local types. */
  id: string
  label: string
  value: ChartType
  family: ChartFamily
  dataShape: ChartDataShape
  stacked?: boolean
  percent?: boolean
  threeD?: boolean
}

/** Source: @univerjs-pro/chart-ui@1.6.0, lib/types/components/ChartTypeOptions.d.ts. */
export const CHART_TYPE_CATALOG: readonly ChartCatalogEntry[] = [
  { id: 'line', label: '折线图', value: 'line', family: 'line', dataShape: 'category' },
  { id: 'lineStacked', label: '堆积折线图', value: 'lineStacked', family: 'line', dataShape: 'category', stacked: true },
  { id: 'linePercentStacked', label: '百分比堆积折线图', value: 'linePercentStacked', family: 'line', dataShape: 'category', stacked: true, percent: true },
  { id: 'line3D', label: '三维折线图', value: 'line3D', family: 'line', dataShape: 'category', threeD: true },
  { id: 'column', label: '簇状柱形图', value: 'column', family: 'column', dataShape: 'category' },
  { id: 'columnStacked', label: '堆积柱形图', value: 'stacked-column', family: 'column', dataShape: 'category', stacked: true },
  { id: 'columnPercentStacked', label: '百分比堆积柱形图', value: 'columnPercentStacked', family: 'column', dataShape: 'category', stacked: true, percent: true },
  { id: 'column3D', label: '三维柱形图', value: 'column3D', family: 'column', dataShape: 'category', threeD: true },
  { id: 'columnStacked3D', label: '三维堆积柱形图', value: 'columnStacked3D', family: 'column', dataShape: 'category', stacked: true, threeD: true },
  { id: 'columnPercentStacked3D', label: '三维百分比堆积柱形图', value: 'columnPercentStacked3D', family: 'column', dataShape: 'category', stacked: true, percent: true, threeD: true },
  { id: 'bar', label: '簇状条形图', value: 'bar-horizontal', family: 'bar', dataShape: 'category' },
  { id: 'barStacked', label: '堆积条形图', value: 'barStacked', family: 'bar', dataShape: 'category', stacked: true },
  { id: 'barPercentStacked', label: '百分比堆积条形图', value: 'barPercentStacked', family: 'bar', dataShape: 'category', stacked: true, percent: true },
  { id: 'bar3D', label: '三维条形图', value: 'bar3D', family: 'bar', dataShape: 'category', threeD: true },
  { id: 'barStacked3D', label: '三维堆积条形图', value: 'barStacked3D', family: 'bar', dataShape: 'category', stacked: true, threeD: true },
  { id: 'barPercentStacked3D', label: '三维百分比堆积条形图', value: 'barPercentStacked3D', family: 'bar', dataShape: 'category', stacked: true, percent: true, threeD: true },
  { id: 'pie', label: '饼图', value: 'pie', family: 'pie', dataShape: 'category' },
  { id: 'donut', label: '圆环图', value: 'doughnut', family: 'pie', dataShape: 'category' },
  { id: 'pieOfPie', label: '复合饼图', value: 'pieOfPie', family: 'pie', dataShape: 'category' },
  { id: 'barOfPie', label: '复合条形饼图', value: 'barOfPie', family: 'pie', dataShape: 'category' },
  { id: 'pie3D', label: '三维饼图', value: 'pie3D', family: 'pie', dataShape: 'category', threeD: true },
  { id: 'area', label: '面积图', value: 'area', family: 'area', dataShape: 'category' },
  { id: 'areaStacked', label: '堆积面积图', value: 'areaStacked', family: 'area', dataShape: 'category', stacked: true },
  { id: 'areaPercentStacked', label: '百分比堆积面积图', value: 'areaPercentStacked', family: 'area', dataShape: 'category', stacked: true, percent: true },
  { id: 'area3D', label: '三维面积图', value: 'area3D', family: 'area', dataShape: 'category', threeD: true },
  { id: 'areaStacked3D', label: '三维堆积面积图', value: 'areaStacked3D', family: 'area', dataShape: 'category', stacked: true, threeD: true },
  { id: 'areaPercentStacked3D', label: '三维百分比堆积面积图', value: 'areaPercentStacked3D', family: 'area', dataShape: 'category', stacked: true, percent: true, threeD: true },
  { id: 'radar', label: '雷达图', value: 'radar', family: 'radar', dataShape: 'category' },
  { id: 'scatter', label: '散点图', value: 'scatter', family: 'scatter', dataShape: 'xy' },
  { id: 'combination', label: '组合图', value: 'combination', family: 'combination', dataShape: 'category' },
  { id: 'wordCloud', label: '词云图', value: 'wordCloud', family: 'wordCloud', dataShape: 'category' },
  { id: 'funnel', label: '漏斗图', value: 'funnel', family: 'funnel', dataShape: 'category' },
  { id: 'bubble', label: '气泡图', value: 'bubble', family: 'bubble', dataShape: 'bubble' },
  { id: 'relation', label: '关系图', value: 'relation', family: 'relation', dataShape: 'edge' },
  { id: 'waterfall', label: '瀑布图', value: 'waterfall', family: 'waterfall', dataShape: 'category' },
  { id: 'pareto', label: '帕累托图', value: 'pareto', family: 'pareto', dataShape: 'category' },
  { id: 'sankey', label: '桑基图', value: 'sankey', family: 'sankey', dataShape: 'edge' },
  { id: 'heatmap', label: '热力图', value: 'heatmap', family: 'heatmap', dataShape: 'heatmap' },
  { id: 'boxplot', label: '箱线图', value: 'boxplot', family: 'boxplot', dataShape: 'distribution' },
  { id: 'candlestick', label: 'K 线图', value: 'candlestick', family: 'candlestick', dataShape: 'ohlc' },
  { id: 'histogram', label: '直方图', value: 'histogram', family: 'histogram', dataShape: 'distribution' },
  { id: 'treemap', label: '树图', value: 'treemap', family: 'treemap', dataShape: 'hierarchy' },
  { id: 'sunburst', label: '旭日图', value: 'sunburst', family: 'sunburst', dataShape: 'hierarchy' },
  { id: 'gauge', label: '仪表盘', value: 'gauge', family: 'gauge', dataShape: 'gauge' },
  { id: 'chord', label: '和弦图', value: 'chord', family: 'chord', dataShape: 'edge' },
]

export function resolveChartType(type: ChartType): ChartCatalogEntry {
  const canonical = type === 'bar' ? 'column' : type === 'columnStacked' ? 'stacked-column' : type
  const entry = CHART_TYPE_CATALOG.find(item => item.value === canonical)
  if (!entry) throw new TypeError(`不支持的图表类型：${type}`)
  return entry
}

export function isChartType(value: unknown): value is ChartType {
  return typeof value === 'string' && (value === 'bar' || value === 'columnStacked' || CHART_TYPE_CATALOG.some(item => item.value === value))
}

export function chartTypeSelectOptions(): Array<{ label: string; value: ChartType }> {
  return CHART_TYPE_CATALOG.map(({ label, value }) => ({ label, value }))
}

/** Shared by controls, the small-data path and Worker scheduling. Combination
 * charts apply this to each series, because only some may be stacked. */
export function supportsChartStatistics(chart: Pick<ChartModel, 'type' | 'options'>, field?: number, position = 0, sourceName?: string): boolean {
  const descriptor = resolveChartType(chart.type)
  if (descriptor.threeD || descriptor.stacked || !['line', 'column', 'bar', 'area', 'scatter', 'bubble', 'combination'].includes(descriptor.family)) return false
  if (descriptor.family !== 'combination' || field === undefined) return true
  const seriesType = chartNamedSeriesStyle(chart.options, sourceName)?.type ?? chart.options?.series?.[field]?.type
    ?? chart.options?.allSeriesStyle?.type ?? (position === 0 ? 'column' : 'line')
  return !seriesType.endsWith('Stacked')
}

export function supportsChartErrorDirection(chart: Pick<ChartModel, 'type' | 'options'>, direction: 'x' | 'y'): boolean {
  return supportsChartStatistics(chart) && (direction === 'y' || ['xy', 'bubble'].includes(resolveChartType(chart.type).dataShape))
}

export function chartDataHint(type: ChartType): string {
  switch (resolveChartType(type).dataShape) {
    case 'xy': return '选择数值 X 与 Y 字段；每个系列可以使用独立 X 字段。'
    case 'bubble': return '选择数值 X、Y 和气泡大小字段，可另选标签及系列分组字段。'
    case 'edge': return '选择来源与目标字段；权重字段可选，缺省按关系出现次数计数。'
    case 'hierarchy': return '按从根到叶顺序选择层级字段，再选择一个数值权重字段；同一路径合并后仅显示正的层级。'
    case 'ohlc': return '选择最高、最低和收盘字段；开盘字段可选，缺省保留三字段 HLC 语义。'
    case 'distribution': return type === 'histogram' ? '选择一个数值样本字段；分箱使用 Scott 自动规则、箱数或箱宽。' : '每个数值字段作为一组样本，显示四分位、须线和异常点。'
    case 'gauge': return '选择仪表数值和可选标签字段；刻度和进度在图表设置中调整。'
    case 'heatmap': return '分类字段定义纵轴行，所选数值字段名定义横轴列，各字段数值决定单元格色阶。'
    default: return type === 'pareto' ? '选择一个分类字段和一个数值权重字段；同名分类先合并，再按合计降序排列并生成累计百分比。负合计排除，零合计可选保留。'
      : resolveChartType(type).family === 'pie' ? '选择一个非负数值系列，合计须大于零。' : '首行或首列为字段名，分类字段与数值系列独立配置。'
  }
}
