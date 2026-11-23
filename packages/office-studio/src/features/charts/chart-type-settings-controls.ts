import type { ChartSemanticOptions } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { chartLabelControls, colorControl, fieldControl, fieldsControl, numberControl, optionGroup, selectControl, switchControl, textControl, type ChartSettingControl } from './chart-settings-controls'
import { chartGaugeSettingsControls, chartGaugeSettingsMapping } from './chart-gauge-settings-controls'
import { chartRadarSettingsControls } from './chart-radar-settings-controls'

type Mapping = NonNullable<ChartSemanticOptions['mapping']>
function mapping(config: ChartConfig): Mapping { return (config.options ??= {}).mapping ??= {} }
function bubble(config: ChartConfig): NonNullable<Mapping['bubble']> { return mapping(config).bubble ??= { xIndex: config.categoryField ?? 0, yIndex: config.valueFields?.[0] ?? 1, sizeIndex: config.valueFields?.[1] } }
function edge(config: ChartConfig, fieldCount: number): NonNullable<Mapping['edge']> { return mapping(config).edge ??= { sourceIndex: config.categoryField ?? 0, targetIndex: config.valueFields?.[0] ?? 1, valueIndex: config.valueFields?.[1] ?? (fieldCount > 2 ? 2 : undefined) } }
function candle(config: ChartConfig): NonNullable<Mapping['candlestick']> {
  const fields = config.valueFields?.length ? config.valueFields : [1, 2, 3]
  return mapping(config).candlestick ??= fields.length >= 4 ? { openIndex: fields[0], highIndex: fields[1]!, lowIndex: fields[2]!, closeIndex: fields[3]! } : { highIndex: fields[0]!, lowIndex: fields[1]!, closeIndex: fields[2]! }
}
function gaugeMapping(config: ChartConfig, fieldCount: number): NonNullable<Mapping['gauge']> { return mapping(config).gauge ??= chartGaugeSettingsMapping(config, fieldCount) }
function secondary(config: ChartConfig): NonNullable<NonNullable<ChartSemanticOptions['pie']>['secondaryPlot']> { return optionGroup(config, 'pie').secondaryPlot ??= {} }
function split(config: ChartConfig): NonNullable<NonNullable<NonNullable<ChartSemanticOptions['pie']>['secondaryPlot']>['split']> { return secondary(config).split ??= { type: 'auto' } }

export function chartDataMappingControls(config: ChartConfig, fieldCount: number): ChartSettingControl[] {
  const descriptor = resolveChartType(config.type), category = fieldControl('categoryField', descriptor.dataShape === 'xy' ? '共享 X 字段' : '分类字段', item => item.categoryField, (item, value) => { item.categoryField = value ?? 0 })
  const values = fieldsControl('valueFields', descriptor.dataShape === 'distribution' ? '样本字段' : descriptor.dataShape === 'xy' ? 'Y 系列字段' : '数值系列字段', item => item.valueFields, (item, value) => { item.valueFields = value })
  if (descriptor.dataShape === 'bubble') return [
    fieldControl('mapping.bubble.xIndex', 'X 数值字段', item => item.options?.mapping?.bubble?.xIndex ?? item.categoryField ?? 0, (item, value) => { bubble(item).xIndex = value ?? 0 }),
    fieldControl('mapping.bubble.yIndex', 'Y 数值字段', item => item.options?.mapping?.bubble?.yIndex ?? item.valueFields?.[0] ?? 1, (item, value) => { bubble(item).yIndex = value ?? 1 }),
    fieldControl('mapping.bubble.sizeIndex', '气泡大小字段', item => item.options?.mapping?.bubble?.sizeIndex ?? item.valueFields?.[1], (item, value) => { bubble(item).sizeIndex = value }),
    fieldControl('mapping.bubble.labelIndex', '气泡标签字段', item => item.options?.mapping?.bubble?.labelIndex, (item, value) => { bubble(item).labelIndex = value }),
    fieldControl('mapping.bubble.seriesIndex', '系列分组字段', item => item.options?.mapping?.bubble?.seriesIndex, (item, value) => { bubble(item).seriesIndex = value }),
  ]
  if (descriptor.dataShape === 'edge') return [
    fieldControl('mapping.edge.sourceIndex', '来源字段', item => item.options?.mapping?.edge?.sourceIndex ?? item.categoryField ?? 0, (item, value) => { edge(item, fieldCount).sourceIndex = value ?? 0 }),
    fieldControl('mapping.edge.targetIndex', '目标字段', item => item.options?.mapping?.edge?.targetIndex ?? item.valueFields?.[0] ?? 1, (item, value) => { edge(item, fieldCount).targetIndex = value ?? 1 }),
    fieldControl('mapping.edge.valueIndex', '权重字段', item => item.options?.mapping?.edge?.valueIndex, (item, value) => { edge(item, fieldCount).valueIndex = value }),
  ]
  if (descriptor.dataShape === 'ohlc') return [category,
    fieldControl('mapping.candlestick.openIndex', '开盘字段（可选）', item => item.options?.mapping?.candlestick?.openIndex ?? (item.valueFields && item.valueFields.length >= 4 ? item.valueFields[0] : undefined), (item, value) => { candle(item).openIndex = value }),
    fieldControl('mapping.candlestick.highIndex', '最高字段', item => item.options?.mapping?.candlestick?.highIndex ?? candle({ ...item, options: undefined }).highIndex, (item, value) => { if (value !== undefined) candle(item).highIndex = value }),
    fieldControl('mapping.candlestick.lowIndex', '最低字段', item => item.options?.mapping?.candlestick?.lowIndex ?? candle({ ...item, options: undefined }).lowIndex, (item, value) => { if (value !== undefined) candle(item).lowIndex = value }),
    fieldControl('mapping.candlestick.closeIndex', '收盘字段', item => item.options?.mapping?.candlestick?.closeIndex ?? candle({ ...item, options: undefined }).closeIndex, (item, value) => { if (value !== undefined) candle(item).closeIndex = value }),
  ]
  if (descriptor.dataShape === 'gauge') return [
    fieldControl('mapping.gauge.valueIndex', '仪表数值字段', item => chartGaugeSettingsMapping(item, fieldCount).valueIndex, (item, value) => {
      const field = value ?? (fieldCount > 1 ? 1 : 0)
      gaugeMapping(item, fieldCount).valueIndex = field; item.valueFields = [field]
    }),
    fieldControl('mapping.gauge.labelIndex', '仪表标签字段（可选）', item => chartGaugeSettingsMapping(item, fieldCount).labelIndex, (item, value) => {
      gaugeMapping(item, fieldCount).labelIndex = value
      if (value !== undefined) item.categoryField = value
    }),
  ]
  if (descriptor.dataShape === 'hierarchy') return [
    fieldsControl('mapping.hierarchyIndexes', '层级字段（从根到叶）', item => item.options?.mapping?.hierarchyIndexes ?? [item.categoryField ?? 0], (item, value) => { mapping(item).hierarchyIndexes = value.length ? value : undefined }),
    fieldControl('valueFields.0', '权重字段', item => item.valueFields?.[0], (item, value) => { item.valueFields = value === undefined ? [] : [value] }),
  ]
  if (descriptor.dataShape === 'distribution') return [values]
  return [category, values, ...(descriptor.dataShape === 'xy' ? [] : [
    fieldsControl('mapping.categoryIndexes', '多级分类字段（按顺序）', item => item.options?.mapping?.categoryIndexes, (item, value) => { mapping(item).categoryIndexes = value.length ? value : undefined }),
  ])]
}

export function chartTypeSettingsControls(config: ChartConfig, fieldCount?: number, sourceName?: string): ChartSettingControl[] {
  const descriptor = resolveChartType(config.type), controls: ChartSettingControl[] = []
  if (descriptor.threeD) controls.push(
    numberControl('threeD.view.rotX', '三维俯仰角', item => item.options?.threeD?.view?.rotX, (item, value) => { (optionGroup(item, 'threeD').view ??= {}).rotX = value }, -90, 90, 28),
    numberControl('threeD.view.rotY', '三维旋转角', item => item.options?.threeD?.view?.rotY, (item, value) => { (optionGroup(item, 'threeD').view ??= {}).rotY = value }, 0, 360, 28),
    numberControl('threeD.view.perspective', '透视强度', item => item.options?.threeD?.view?.perspective, (item, value) => { (optionGroup(item, 'threeD').view ??= {}).perspective = value }, 0, 120, 30),
    switchControl('threeD.view.rightAngleAxes', '直角投影', item => item.options?.threeD?.view?.rightAngleAxes, (item, value) => { (optionGroup(item, 'threeD').view ??= {}).rightAngleAxes = value }),
    numberControl('threeD.view.heightPercent', '高度比例（%）', item => item.options?.threeD?.view?.heightPercent, (item, value) => { (optionGroup(item, 'threeD').view ??= {}).heightPercent = value }, 1, 500, 100),
    numberControl('threeD.view.depthPercent', '深度比例（%）', item => item.options?.threeD?.view?.depthPercent, (item, value) => { (optionGroup(item, 'threeD').view ??= {}).depthPercent = value }, 1, 500, 100),
    selectControl('threeD.grouping', '三维系列布局', [['standard', '沿深度排列'], ['clustered', '按分类聚集']], item => item.options?.threeD?.grouping, (item, value) => { optionGroup(item, 'threeD').grouping = value }, 'standard'),
    numberControl('threeD.gapDepth', '系列深度间距（%）', item => item.options?.threeD?.gapDepth, (item, value) => { optionGroup(item, 'threeD').gapDepth = value }, 0, 500, 150),
  )
  if (descriptor.family === 'column' || descriptor.family === 'bar' || descriptor.family === 'combination' || descriptor.family === 'waterfall') controls.push(
    numberControl('bar.gapWidth', '柱条间距（%）', item => item.options?.bar?.gapWidth, (item, value) => { optionGroup(item, 'bar').gapWidth = value }, 0, 500, 150),
    numberControl('bar.overlap', '系列重叠（%）', item => item.options?.bar?.overlap, (item, value) => { optionGroup(item, 'bar').overlap = value }, -100, 100, 0),
  )
  if (descriptor.family === 'area' && !descriptor.threeD) controls.push(switchControl('area.smooth', '平滑面积边界', item => item.options?.area?.smooth, (item, value) => { optionGroup(item, 'area').smooth = value }))
  if (descriptor.family === 'pie') {
    controls.push(
      numberControl('pie.startAngle', '起始角', item => item.options?.pie?.startAngle, (item, value) => { optionGroup(item, 'pie').startAngle = value }, -360, 360, -90),
      numberControl('pie.explosion', '扇区分离（%）', item => item.options?.pie?.explosion, (item, value) => { optionGroup(item, 'pie').explosion = value }, 0, 100, 0),
      ...(config.type === 'doughnut' ? [numberControl('pie.doughnutHole', '圆环内径比例', item => item.options?.pie?.doughnutHole, (item, value) => { optionGroup(item, 'pie').doughnutHole = value }, 0, 0.95)] : []),
      ...(!descriptor.threeD ? [switchControl('pie.rosePie', '玫瑰图半径', item => item.options?.pie?.rosePie, (item, value) => { optionGroup(item, 'pie').rosePie = value }),
        switchControl('pie.isHalfPie', '半圆饼图', item => item.options?.pie?.isHalfPie, (item, value) => { optionGroup(item, 'pie').isHalfPie = value }),
        switchControl('pie.showLabelLine', '显示标签引导线', item => item.options?.pie?.showLabelLine, (item, value) => { optionGroup(item, 'pie').showLabelLine = value }, true),
        ...chartLabelControls('pie.label', item => item.options?.pie?.label, item => optionGroup(item, 'pie').label ??= {}, true)] : []),
    )
    if (config.type === 'pieOfPie' || config.type === 'barOfPie') {
      controls.push(
        numberControl('pie.secondaryPlot.size', '副图大小（%）', item => item.options?.pie?.secondaryPlot?.size, (item, value) => { secondary(item).size = value }, 15, 100, 65),
        numberControl('pie.secondaryPlot.gapWidth', '主副图间距（%）', item => item.options?.pie?.secondaryPlot?.gapWidth, (item, value) => { secondary(item).gapWidth = value }, 0, 500, 50),
        textControl('pie.aggregateLabel', '主图合计标签', item => item.options?.pie?.aggregateLabel, (item, value) => { optionGroup(item, 'pie').aggregateLabel = value }),
        selectControl('pie.secondaryPlot.split.type', '拆分规则', [['auto', '自动'], ['position', '按位置'], ['value', '按数值'], ['percentage', '按百分比'], ['custom', '指定项目']], item => item.options?.pie?.secondaryPlot?.split?.type, (item, value) => { split(item).type = value ?? 'auto' }, 'auto'),
      )
      const splitType = config.options?.pie?.secondaryPlot?.split?.type
      if (splitType === 'position' || splitType === 'auto') controls.push(numberControl('pie.secondaryPlot.split.count', '副图项目数', item => item.options?.pie?.secondaryPlot?.split?.count, (item, value) => { split(item).count = value }, 1, 100_000))
      if (splitType === 'value' || splitType === 'percentage') controls.push(numberControl('pie.secondaryPlot.split.lessThan', splitType === 'percentage' ? '低于百分比' : '低于数值', item => item.options?.pie?.secondaryPlot?.split?.lessThan, (item, value) => { split(item).lessThan = value }, 0, splitType === 'percentage' ? 100 : undefined))
      if (splitType === 'custom') controls.push(fieldsControl('pie.secondaryPlot.split.pointIndexes', '副图项目', item => item.options?.pie?.secondaryPlot?.split?.pointIndexes, (item, value) => { split(item).pointIndexes = value }, true))
    }
  }
  if (descriptor.family === 'radar') controls.push(...chartRadarSettingsControls(config))
  if (descriptor.family === 'bubble') controls.push(
    numberControl('bubble.scale', '气泡比例', item => item.options?.bubble?.scale, (item, value) => { optionGroup(item, 'bubble').scale = value }, 0.01, 10_000, 1),
    selectControl('bubble.sizeRepresents', '数值对应', [['area', '面积'], ['width', '直径']], item => item.options?.bubble?.sizeRepresents, (item, value) => { optionGroup(item, 'bubble').sizeRepresents = value }, 'area'),
  )
  if (descriptor.family === 'funnel') controls.push(
    numberControl('funnel.gap', '漏斗段间距', item => item.options?.funnel?.gap, (item, value) => { optionGroup(item, 'funnel').gap = value }, 0, 100, 2),
    switchControl('funnel.useAbsoluteValue', '使用权重绝对值', item => item.options?.funnel?.useAbsoluteValue, (item, value) => { optionGroup(item, 'funnel').useAbsoluteValue = value }),
    selectControl('funnel.sort', '漏斗排序', [['descending', '降序'], ['ascending', '升序'], ['none', '原始顺序']], item => item.options?.funnel?.sort, (item, value) => { optionGroup(item, 'funnel').sort = value }, 'descending'),
  )
  if (descriptor.family === 'relation') controls.push(
    selectControl('relation.layout', '关系布局', [['none', '固定布局'], ['circular', '圆形'], ['force', '力导向']], item => item.options?.relation?.layout, (item, value) => { optionGroup(item, 'relation').layout = value }, 'force'),
    switchControl('relation.useValueAsSymbolSize', '权重决定节点大小', item => item.options?.relation?.useValueAsSymbolSize, (item, value) => { optionGroup(item, 'relation').useValueAsSymbolSize = value }),
    numberControl('relation.force.repulsion', '排斥强度', item => item.options?.relation?.force?.repulsion, (item, value) => { (optionGroup(item, 'relation').force ??= {}).repulsion = value }, 0, 100_000, 180),
    numberControl('relation.force.gravity', '中心引力', item => item.options?.relation?.force?.gravity, (item, value) => { (optionGroup(item, 'relation').force ??= {}).gravity = value }, 0, 10, 0.1),
    numberControl('relation.force.edgeLength', '连线长度', item => item.options?.relation?.force?.edgeLength, (item, value) => { (optionGroup(item, 'relation').force ??= {}).edgeLength = value }, 0, 10_000, 70),
  )
  if (descriptor.family === 'waterfall') controls.push(
    switchControl('waterfall.connector', '显示累计连接线', item => item.options?.waterfall?.connector, (item, value) => { optionGroup(item, 'waterfall').connector = value }, true),
    colorControl('waterfall.positiveColor', '增加颜色', item => item.options?.waterfall?.positiveColor, (item, value) => { optionGroup(item, 'waterfall').positiveColor = value }),
    colorControl('waterfall.negativeColor', '减少颜色', item => item.options?.waterfall?.negativeColor, (item, value) => { optionGroup(item, 'waterfall').negativeColor = value }),
    colorControl('waterfall.totalColor', '合计颜色', item => item.options?.waterfall?.totalColor, (item, value) => { optionGroup(item, 'waterfall').totalColor = value }),
  )
  if (descriptor.family === 'pareto') controls.push(switchControl('pareto.includeZeroValues', '保留零权重分类', item => item.options?.pareto?.includeZeroValues, (item, value) => { optionGroup(item, 'pareto').includeZeroValues = value }))
  if (descriptor.family === 'histogram') {
    controls.push(selectControl('histogram.binning.mode', '分箱方式', [['auto', 'Scott 自动'], ['count', '按箱数'], ['width', '按箱宽']], item => item.options?.histogram?.binning?.mode ?? 'auto', (item, value) => { optionGroup(item, 'histogram').binning = value === 'count' ? { mode: 'count', count: 10 } : value === 'width' ? { mode: 'width', width: 1 } : undefined }, 'auto'))
    if (config.options?.histogram?.binning?.mode === 'count') controls.push(numberControl('histogram.binning.count', '箱数', item => item.options?.histogram?.binning?.mode === 'count' ? item.options.histogram.binning.count : undefined, (item, value) => { if (value !== undefined) optionGroup(item, 'histogram').binning = { mode: 'count', count: value } }, 1, 10_000, 10))
    if (config.options?.histogram?.binning?.mode === 'width') controls.push(numberControl('histogram.binning.width', '箱宽', item => item.options?.histogram?.binning?.mode === 'width' ? item.options.histogram.binning.width : undefined, (item, value) => { if (value !== undefined) optionGroup(item, 'histogram').binning = { mode: 'width', width: value } }, Number.MIN_VALUE))
    controls.push(
      numberControl('histogram.underflowThreshold', '下溢箱阈值（≤）', item => item.options?.histogram?.underflowThreshold, (item, value) => { optionGroup(item, 'histogram').underflowThreshold = value }),
      numberControl('histogram.overflowThreshold', '上溢箱阈值（>）', item => item.options?.histogram?.overflowThreshold, (item, value) => { optionGroup(item, 'histogram').overflowThreshold = value }),
      numberControl('histogram.binGap', '箱间隙比例', item => item.options?.histogram?.binGap, (item, value) => { optionGroup(item, 'histogram').binGap = value }, 0, 0.9, 0.05),
    )
  }
  if (descriptor.family === 'boxplot') controls.push(
    selectControl('boxplot.quartileMethod', '四分位方法', [['inclusive', '包含中位数'], ['exclusive', '排除中位数']], item => item.options?.boxplot?.quartileMethod, (item, value) => { optionGroup(item, 'boxplot').quartileMethod = value }, 'inclusive'),
    switchControl('boxplot.showOutliers', '显示异常点', item => item.options?.boxplot?.showOutliers, (item, value) => { optionGroup(item, 'boxplot').showOutliers = value }, true),
    numberControl('boxplot.whiskerMultiplier', '须线 IQR 倍数', item => item.options?.boxplot?.whiskerMultiplier, (item, value) => { optionGroup(item, 'boxplot').whiskerMultiplier = value }, 0, 100, 1.5),
  )
  if (descriptor.family === 'candlestick') controls.push(
    numberControl('candlestick.candleWidth', '蜡烛宽度', item => item.options?.candlestick?.candleWidth, (item, value) => { optionGroup(item, 'candlestick').candleWidth = value }, 1, 100),
    colorControl('candlestick.risingColor', '上涨颜色', item => item.options?.candlestick?.risingColor, (item, value) => { optionGroup(item, 'candlestick').risingColor = value }),
    colorControl('candlestick.fallingColor', '下跌颜色', item => item.options?.candlestick?.fallingColor, (item, value) => { optionGroup(item, 'candlestick').fallingColor = value }),
    switchControl('candlestick.hollow', '上涨蜡烛空心', item => item.options?.candlestick?.hollow, (item, value) => { optionGroup(item, 'candlestick').hollow = value }),
  )
  if (descriptor.family === 'heatmap') controls.push(
    selectControl('heatmap.visualMapType', '色阶类型', [['continuous', '连续'], ['piecewise', '分段']], item => item.options?.heatmap?.visualMapType, (item, value) => { optionGroup(item, 'heatmap').visualMapType = value }, 'continuous'),
    numberControl('heatmap.min', '色阶最小值', item => item.options?.heatmap?.min, (item, value) => { optionGroup(item, 'heatmap').min = value }),
    numberControl('heatmap.max', '色阶最大值', item => item.options?.heatmap?.max, (item, value) => { optionGroup(item, 'heatmap').max = value }),
    textControl('heatmap.unit', '数值单位', item => item.options?.heatmap?.unit, (item, value) => { optionGroup(item, 'heatmap').unit = value }),
    ...chartLabelControls('heatmap.label', item => item.options?.heatmap?.label, item => optionGroup(item, 'heatmap').label ??= {}),
  )
  if (descriptor.family === 'treemap') controls.push(
    selectControl('treemap.parentLabelLayout', '父级标签布局', [['banner', '独立标题带'], ['overlapping', '叠放'], ['none', '隐藏']], item => item.options?.treemap?.parentLabelLayout, (item, value) => { optionGroup(item, 'treemap').parentLabelLayout = value }, 'banner'),
    ...chartLabelControls('treemap.label', item => item.options?.treemap?.label, item => optionGroup(item, 'treemap').label ??= {}, true),
  )
  if (descriptor.family === 'sunburst') controls.push(
    switchControl('sunburst.centerLabel.visible', '显示中心标题', item => item.options?.sunburst?.centerLabel?.visible, (item, value) => { (optionGroup(item, 'sunburst').centerLabel ??= {}).visible = value }),
    ...chartLabelControls('sunburst.label', item => item.options?.sunburst?.label, item => optionGroup(item, 'sunburst').label ??= {}, true),
  )
  if (descriptor.family === 'gauge') controls.push(...chartGaugeSettingsControls(config, fieldCount, sourceName))
  if (descriptor.family === 'chord') controls.push(
    numberControl('chord.startAngle', '和弦起始角', item => item.options?.chord?.startAngle, (item, value) => { optionGroup(item, 'chord').startAngle = value }, -360, 360, -90),
    switchControl('chord.clockwise', '顺时针', item => item.options?.chord?.clockwise, (item, value) => { optionGroup(item, 'chord').clockwise = value }, true),
    numberControl('chord.padAngle', '节点间隔角', item => item.options?.chord?.padAngle, (item, value) => { optionGroup(item, 'chord').padAngle = value }, 0, 45, 3),
    numberControl('chord.innerRadiusRatio', '内半径比例', item => item.options?.chord?.innerRadiusRatio, (item, value) => { optionGroup(item, 'chord').innerRadiusRatio = value }, 0.05, 0.95, 0.66),
    numberControl('chord.outerRadiusRatio', '外半径比例', item => item.options?.chord?.outerRadiusRatio, (item, value) => { optionGroup(item, 'chord').outerRadiusRatio = value }, 0.05, 1, 0.73),
    numberControl('chord.ribbonOpacity', '弦带不透明度', item => item.options?.chord?.ribbonOpacity, (item, value) => { optionGroup(item, 'chord').ribbonOpacity = value }, 0, 1, 0.55),
    ...chartLabelControls('chord.label', item => item.options?.chord?.label, item => optionGroup(item, 'chord').label ??= {}, true),
  )
  if (descriptor.family === 'wordCloud') controls.push(
    selectControl('wordCloud.shape', '词云形状', [['cloud', '云形'], ['rect', '矩形'], ['circle', '圆形'], ['cardioid', '心迹线'], ['heart', '心形'], ['diamond', '菱形'], ['triangle-forward', '右三角'], ['triangle', '三角形'], ['pentagon', '五边形'], ['star', '星形']], item => item.options?.wordCloud?.shape, (item, value) => { optionGroup(item, 'wordCloud').shape = value }, 'cloud'),
    switchControl('wordCloud.repeat', '重复填充词语', item => item.options?.wordCloud?.repeat, (item, value) => { optionGroup(item, 'wordCloud').repeat = value }),
    numberControl('wordCloud.minFontSize', '最小字号', item => item.options?.wordCloud?.minFontSize, (item, value) => { optionGroup(item, 'wordCloud').minFontSize = value }, 1, 256, 10),
    numberControl('wordCloud.maxFontSize', '最大字号', item => item.options?.wordCloud?.maxFontSize, (item, value) => { optionGroup(item, 'wordCloud').maxFontSize = value }, 1, 256, 54),
    numberControl('wordCloud.rotation', '词语旋转角', item => item.options?.wordCloud?.rotation, (item, value) => { optionGroup(item, 'wordCloud').rotation = value }, -90, 90, 0),
  )
  return controls
}

export function chartAggregationControls(config: ChartConfig): ChartSettingControl[] {
  const read = (item: ChartConfig) => item.options?.aggregation
  const top = (item: ChartConfig) => optionGroup(item, 'aggregation').topN ??= { count: 10 }
  return [
    switchControl('aggregation.aggregate', '合并相同分类', item => read(item)?.aggregate, (item, value) => { optionGroup(item, 'aggregation').aggregate = value }),
    switchControl('aggregation.topN', '显示前 N 项', item => Boolean(read(item)?.topN), (item, value) => { optionGroup(item, 'aggregation').topN = value ? { count: read(item)?.topN?.count ?? 10 } : undefined }),
    ...(config.options?.aggregation?.topN ? [
      numberControl('aggregation.topN.count', '项目数', item => read(item)?.topN?.count, (item, value) => { top(item).count = value ?? 10 }, 1, 100_000, 10),
      switchControl('aggregation.topN.includeOthers', '合并其余为其他', item => read(item)?.topN?.includeOthers, (item, value) => { top(item).includeOthers = value }),
      textControl('aggregation.topN.othersLabel', '其他分类标签', item => read(item)?.topN?.othersLabel, (item, value) => { top(item).othersLabel = value }),
    ] : []),
  ]
}
