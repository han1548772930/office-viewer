import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartLineOptions, ChartSemanticOptions, ChartSeriesOptions, ChartSeriesFillPattern } from '../../domain/model'
import { chartColorChannels, validateChartColor } from './chart-color'
import { isChartFillPatternPreset } from './chart-fill-pattern'
import { validateChartPlotArea } from './chart-plot-area'
import { CHART_POINT_SHAPES } from './chart-markers'
import { CHART_LABEL_POSITION_CHOICES, chartLabelContentMask } from './chart-labels'
import { validateChartChordColor } from './chart-chord-color'

export function validateChartPalette(palette: readonly string[] | undefined): void {
  if (palette === undefined) return
  if (!Array.isArray(palette) || palette.length > 256) throw new RangeError('图表配色须为不超过 256 项的颜色数组。')
  palette.forEach((entry, index) => color(entry, `图表配色 ${index + 1}`))
}

/** Validate authored values before cloning, reading data or committing a
 * candidate. Unknown serializable fields are retained for future snapshots. */
export function validateChartOptions(options: ChartSemanticOptions, { partial = false }: { partial?: boolean } = {}): void {
  plain(options, '图表配置')
  visitValues(options)
  enumeration(options.invalidValueStrategy, ['zero', 'break', 'link'], '空值处理')
  boolean(options.plotVisibleOnly, '仅绘制可见单元格')
  boolean(options.autoGradientFill, '自动渐变填充')
  validateChartPlotArea(options.plotArea)
  for (const key of ['appearance', 'titleStyle', 'subtitleStyle', 'allSeriesStyle', 'axisPointer', 'rightYAxis', 'depthAxis', 'mapping', 'threeD', 'bar', 'area', 'pie', 'radar', 'bubble', 'funnel', 'relation', 'waterfall', 'pareto', 'histogram', 'boxplot', 'candlestick', 'heatmap', 'treemap', 'sunburst', 'gauge', 'chord', 'wordCloud', 'aggregation'] as const) optionalObject(options[key], `图表.${key}`)
  if (options.appearance) {
    color(options.appearance.background, '图表背景'); line(options.appearance.border, '图表边框'); font(options.appearance.font, '图表字体')
  }
  font(options.titleStyle, '标题字体'); font(options.subtitleStyle, '副标题字体')
  if (options.axisPointer) {
    color(options.axisPointer.indicatorLineColor, '轴指示线颜色'); color(options.axisPointer.indicatorLabelColor, '轴指示标签背景')
    color(options.axisPointer.indicatorLabelTextColor, '轴指示标签文字'); enumeration(options.axisPointer.indicatorLineType, ['solid', 'dashed', 'dotted'], '轴指示线型')
  }
  series(options.allSeriesStyle, '所有系列样式'); axis(options.rightYAxis, '右侧数值轴')
  if (options.depthAxis) {
    for (const key of ['scaleType', 'logBase', 'min', 'max', 'format', 'baseTimeUnit', 'majorTimeUnit', 'majorUnit']) if (key in options.depthAxis) throw new TypeError(`分类深度轴不接受 ${key}。`)
    axis({ scaleType: 'linear', ...options.depthAxis }, '分类深度轴')
  }
  if (options.series !== undefined) {
    plain(options.series, '系列样式')
    for (const [key, style] of Object.entries(options.series)) { recordIndex(key, '系列字段'); series(style, `系列 ${key}`) }
  }
  if (options.namedSeries !== undefined) {
    optionalArray(options.namedSeries, '名称系列样式')
    const names = new Set<string>()
    for (let index = 0; index < options.namedSeries.length; index++) {
      if (!Object.hasOwn(options.namedSeries, index)) throw new TypeError('名称系列样式不接受稀疏数组。')
      const entry = options.namedSeries[index]!
      plain(entry, `名称系列 ${index + 1}`)
      if (typeof entry.name !== 'string') throw new TypeError('名称系列选择器须为文本。')
      if (names.has(entry.name)) throw new RangeError(`名称系列 ${JSON.stringify(entry.name)} 不能重复配置。`)
      names.add(entry.name)
      plain(entry.style, `名称系列 ${index + 1} 样式`)
      series(entry.style, `名称系列 ${index + 1}`)
    }
  }

  const mapping = options.mapping
  if (mapping) {
    boolean(mapping.useDateAxis, '使用日期轴')
    indexes(mapping.categoryIndexes, '分类字段', true); indexes(mapping.hierarchyIndexes, '层级字段', true)
    for (const key of ['bubble', 'edge', 'chord', 'candlestick', 'gauge'] as const) {
      const role = mapping[key]
      optionalObject(role, `${key} 字段映射`)
      if (role) for (const [name, value] of Object.entries(role)) if (name.endsWith('Index')) optionalInteger(value, 0, 1_048_575, `${key}.${name}`)
    }
    if (mapping.bubble) { integer(mapping.bubble.xIndex, 0, 1_048_575, 'bubble.xIndex'); integer(mapping.bubble.yIndex, 0, 1_048_575, 'bubble.yIndex') }
    if (mapping.chord) {
      integer(mapping.chord.sourceIndex, 0, 1_048_575, 'chord.sourceIndex')
      integer(mapping.chord.targetIndex, 0, 1_048_575, 'chord.targetIndex')
      if (mapping.chord.sourceIndex === mapping.chord.targetIndex || mapping.chord.valueIndex === mapping.chord.sourceIndex
        || mapping.chord.valueIndex === mapping.chord.targetIndex) throw new RangeError('和弦来源、目标和权重须使用互不重复的字段。')
    }
  }
  const threeD = options.threeD
  if (threeD) {
    enumeration(threeD.grouping, ['standard', 'clustered'], '三维系列布局'); optionalNumber(threeD.gapDepth, 0, 500, '三维深度间距')
    optionalObject(threeD.view, '三维视角')
    if (threeD.view) {
      optionalNumber(threeD.view.rotX, -90, 90, '三维俯仰角'); optionalNumber(threeD.view.rotY, 0, 360, '三维旋转角')
      optionalNumber(threeD.view.perspective, 0, 120, '透视强度'); boolean(threeD.view.rightAngleAxes, '直角投影')
      optionalNumber(threeD.view.heightPercent, 1, 500, '三维高度比例'); optionalNumber(threeD.view.depthPercent, 1, 500, '三维深度比例')
    }
  }
  optionalNumber(options.bar?.gapWidth, 0, 500, '柱条间距'); optionalNumber(options.bar?.overlap, -100, 100, '系列重叠')
  boolean(options.area?.smooth, '平滑面积边界')
  const pie = options.pie
  if (pie) {
    optionalNumber(pie.startAngle, -360, 360, '饼图起始角'); optionalNumber(pie.doughnutHole, 0, 0.95, '圆环内径比例'); optionalNumber(pie.explosion, 0, 100, '扇区分离')
    boolean(pie.rosePie, '玫瑰图半径'); boolean(pie.isHalfPie, '半圆饼图'); boolean(pie.showLabelLine, '饼图标签引导线')
    label(pie.label, '饼图标签'); string(pie.aggregateLabel, '主图合计标签')
    optionalObject(pie.secondaryPlot, '饼图副图')
    if (pie.secondaryPlot) {
      const secondary = pie.secondaryPlot
      enumeration(secondary.type, ['pie', 'bar'], '饼图副图类型'); optionalNumber(secondary.size, 15, 100, '副图大小'); optionalNumber(secondary.gapWidth, 0, 500, '主副图间距')
      optionalObject(secondary.split, '复合饼图拆分规则')
      if (secondary.split) {
        const split = secondary.split
        requiredEnum(split.type, ['auto', 'position', 'value', 'percentage', 'custom'], '复合饼图拆分类型')
        optionalInteger(split.count, 1, 100_000, '副图项目数')
        optionalNumber(split.lessThan, 0, split.type === 'percentage' ? 100 : Number.MAX_VALUE, '副图拆分阈值')
        indexes(split.pointIndexes, '副图项目索引')
      }
    }
  }
  enumeration(options.radar?.shape, ['polygon', 'circle'], '雷达网格'); boolean(options.radar?.fill, '雷达区域填充')
  optionalObject(options.radar?.axisLabel, '雷达数值刻度标签')
  if (options.radar?.axisLabel) {
    font(options.radar.axisLabel, '雷达数值刻度标签'); boolean(options.radar.axisLabel.visible, '雷达数值刻度标签显隐'); string(options.radar.axisLabel.format, '雷达刻度数字格式')
  }
  optionalNumber(options.bubble?.scale, Number.MIN_VALUE, 10_000, '气泡比例'); enumeration(options.bubble?.sizeRepresents, ['area', 'width'], '气泡大小含义')
  optionalNumber(options.funnel?.gap, 0, 100, '漏斗段间距'); boolean(options.funnel?.useAbsoluteValue, '漏斗绝对值'); enumeration(options.funnel?.sort, ['ascending', 'descending', 'none'], '漏斗排序')
  const relation = options.relation
  if (relation) {
    enumeration(relation.layout, ['none', 'circular', 'force'], '关系图布局'); boolean(relation.useValueAsSymbolSize, '节点权重大小')
    optionalObject(relation.force, '关系图力导向配置')
    optionalNumber(relation.force?.repulsion, 0, 100_000, '关系图排斥强度'); optionalNumber(relation.force?.gravity, 0, 10, '关系图中心引力'); optionalNumber(relation.force?.edgeLength, 0, 10_000, '关系图边长')
  }
  const waterfall = options.waterfall
  if (waterfall) {
    enumeration(waterfall.stackType, ['arrangement', 'stacked'], '瀑布图排列方式'); boolean(waterfall.useSubtotal, '瀑布图自动合计')
    if (waterfall.connector !== undefined && typeof waterfall.connector !== 'boolean') line(waterfall.connector, '瀑布图连接线')
    color(waterfall.positiveColor, '瀑布图增加颜色'); color(waterfall.negativeColor, '瀑布图减少颜色'); color(waterfall.totalColor, '瀑布图合计颜色')
    waterfallStyles(waterfall.styles, '瀑布图共享角色样式')
    optionalArray(waterfall.pointRoles, '瀑布图项目含义')
    if (waterfall.stackType === 'stacked' && waterfall.pointRoles?.length) throw new TypeError('堆积瀑布图不能指定源项目含义。请先切换为排列瀑布图。')
    waterfall.pointRoles?.forEach(role => {
      plain(role, '瀑布图项目含义'); integer(role.seriesIndex, 0, 1_048_575, '瀑布图系列字段'); integer(role.dataPointIndex, 0, 1_048_575, '瀑布图项目索引')
      requiredEnum(role.role, ['change', 'total'], '瀑布图项目含义')
    })
  }
  const pareto = options.pareto
  boolean(pareto?.includeZeroValues, '保留零权重分类')
  if (pareto) {
    optionalObject(pareto.barStyle, '帕累托柱形样式'); optionalObject(pareto.cumulativeLineStyle, '帕累托累计线样式')
    const bar = pareto.barStyle, cumulative = pareto.cumulativeLineStyle
    if (bar) {
      series({ color: bar.color, fillOpacity: bar.opacity, border: bar.border, label: bar.label, dataPoints: bar.dataPoints }, '帕累托柱形')
      chartLabelContentMask(bar.label?.contentType, false, '帕累托柱形标签内容')
    }
    if (cumulative) {
      optionalNumber(cumulative.width, -Number.MAX_VALUE, 100, '帕累托累计线宽'); enumeration(cumulative.dashType, ['solid', 'dashed', 'dotted'], '帕累托累计线型')
      series({ color: cumulative.color, fillOpacity: cumulative.opacity, point: cumulative.point, label: cumulative.label, dataPoints: cumulative.dataPoints }, '帕累托累计线')
      chartLabelContentMask(cumulative.label?.contentType, false, '帕累托累计线标签内容')
      if (cumulative.label?.position === 'outside') throw new TypeError('帕累托累计线标签不接受 outside 位置。')
    }
  }
  const histogram = options.histogram
  if (histogram) {
    optionalObject(histogram.binning, '直方图分箱')
    if (histogram.binning) {
      requiredEnum(histogram.binning.mode, ['count', 'width'], '直方图分箱方式')
      // Ineffective finite author values remain visible and persisted. Only
      // the runtime binning resolver chooses Scott Auto for those values.
      if (histogram.binning.mode === 'count') number(histogram.binning.count, -Number.MAX_VALUE, Number.MAX_VALUE, '直方图箱数')
      else number(histogram.binning.width, -Number.MAX_VALUE, Number.MAX_VALUE, '直方图箱宽')
    }
    optionalNumber(histogram.binGap, -Number.MAX_VALUE, Number.MAX_VALUE, '直方图箱间距')
    optionalNumber(histogram.underflowThreshold, -Number.MAX_VALUE, Number.MAX_VALUE, '下溢阈值'); optionalNumber(histogram.overflowThreshold, -Number.MAX_VALUE, Number.MAX_VALUE, '上溢阈值')
  }
  enumeration(options.boxplot?.quartileMethod, ['medianOfHalves', 'inclusive', 'exclusive'], '四分位方法'); enumeration(options.boxplot?.whiskerMethod, ['range', 'iqr'], '箱线图须线方法'); boolean(options.boxplot?.showOutliers, '箱线图异常点'); optionalNumber(options.boxplot?.whiskerMultiplier, 0, 100, '箱线图须线倍数')
  const candlestick = options.candlestick
  if (candlestick) {
    optionalNumber(candlestick.candleWidth, Number.MIN_VALUE, Number.MAX_VALUE, '蜡烛宽度')
    color(candlestick.risingColor, '上涨颜色'); color(candlestick.fallingColor, '下跌颜色'); boolean(candlestick.hollow, '上涨蜡烛空心')
    for (const [state, title] of [['rising', '上涨'], ['falling', '下跌'], ['doji', '十字线']] as const) {
      const style = candlestick[state]
      optionalObject(style, `${title}蜡烛样式`)
      if (!style) continue
      color(style.color, `${title}蜡烛颜色`); optionalNumber(style.fillOpacity, 0, 1, `${title}蜡烛填充不透明度`); boolean(style.hollow, `${title}蜡烛空心`)
      optionalObject(style.border, `${title}蜡烛边框`)
      if (style.border) {
        color(style.border.color, `${title}蜡烛边框颜色`); optionalNumber(style.border.width, 0, Number.MAX_VALUE, `${title}蜡烛边框线宽`)
        optionalNumber(style.border.opacity, 0, 1, `${title}蜡烛边框不透明度`)
      }
    }
    optionalObject(candlestick.highLowLine, 'K 线高低线'); boolean(candlestick.highLowLine?.visible, 'K 线高低线显隐')
    const highLowBorder = candlestick.highLowLine?.border
    optionalObject(highLowBorder, 'K 线高低线样式')
    if (highLowBorder) {
      boolean(highLowBorder.visible, 'K 线高低线线条显隐'); color(highLowBorder.color, 'K 线高低线颜色')
      optionalNumber(highLowBorder.width, 0, Number.MAX_VALUE, 'K 线高低线线宽'); optionalNumber(highLowBorder.opacity, 0, 1, 'K 线高低线不透明度')
      enumeration(highLowBorder.dashType, ['solid', 'dashed', 'dotted'], 'K 线高低线线型')
    }
    optionalObject(candlestick.closeMarker, 'K 线收盘标记'); boolean(candlestick.closeMarker?.visible, 'K 线收盘标记显隐')
    const closePoint = candlestick.closeMarker?.point
    optionalObject(closePoint, 'K 线收盘标记样式')
    enumeration(closePoint?.shape, CHART_POINT_SHAPES, 'K 线收盘标记形状'); optionalNumber(closePoint?.size, 0, 256, 'K 线收盘标记大小'); color(closePoint?.color, 'K 线收盘标记颜色')
  }
  enumeration(options.heatmap?.visualMapType, ['continuous', 'piecewise'], '热力图色阶'); string(options.heatmap?.unit, '热力图单位'); label(options.heatmap?.label, '热力图标签')
  optionalNumber(options.heatmap?.min, -Number.MAX_VALUE, Number.MAX_VALUE, '热力图最小值'); optionalNumber(options.heatmap?.max, -Number.MAX_VALUE, Number.MAX_VALUE, '热力图最大值')
  if (options.heatmap?.min !== undefined && options.heatmap.max !== undefined && options.heatmap.min > options.heatmap.max) throw new RangeError('热力图最小值不能大于最大值。')
  enumeration(options.treemap?.parentLabelLayout, ['banner', 'overlapping', 'none'], '树图父级标签'); label(options.treemap?.label, '树图标签')
  validateChartTreemapBackground(options.treemap?.label?.backgroundColor)
  label(options.sunburst?.label, '旭日图标签')
  enumeration(options.sunburst?.label?.position, ['inside', 'outside'], '旭日图标签位置')
  validateChartTreemapBackground(options.sunburst?.label?.backgroundColor, '旭日图标签背景')
  font(options.sunburst?.centerLabel, '旭日图中心标签'); boolean(options.sunburst?.centerLabel?.visible, '旭日图中心标签显隐')
  string(options.sunburst?.centerLabel?.format, '旭日图中心数字格式'); font(options.sunburst?.centerLabel?.font, '旭日图中心旧版字体')
  const gauge = options.gauge
  if (gauge) {
    enumeration(gauge.mode, ['modern', 'classic', 'pointer', 'progress'], '仪表模式'); optionalNumber(gauge.min, -Number.MAX_VALUE, Number.MAX_VALUE, '仪表最小刻度'); optionalNumber(gauge.max, -Number.MAX_VALUE, Number.MAX_VALUE, '仪表最大刻度')
    const min = gauge.min ?? 0, max = gauge.max ?? 100
    if (min >= max) throw new RangeError('仪表最大刻度须大于最小刻度，包括未设置端使用的默认值。')
    optionalNumber(gauge.startAngle, -360, 360, '仪表起始角'); optionalNumber(gauge.endAngle, -360, 360, '仪表终止角'); optionalInteger(gauge.splitNumber, 1, 100, '仪表主刻度分段数')
    optionalArray(gauge.ranges, '仪表颜色区间')
    let previousEnd = -Infinity
    for (const range of gauge.ranges ?? []) {
      plain(range, '仪表颜色区间'); number(range.from, min, max, '仪表区间起点'); number(range.to, min, max, '仪表区间终点'); color(range.color, '仪表区间颜色', true)
      if (range.from >= range.to || range.from < previousEnd) throw new RangeError('仪表颜色区间须按起点升序排列，且不能重叠或为空。')
      previousEnd = range.to
    }
    optionalObject(gauge.axisLine, '仪表轴线'); boolean(gauge.axisLine?.visible, '仪表轴线显隐'); optionalNumber(gauge.axisLine?.width, 0, 100, '仪表轴线宽度'); color(gauge.axisLine?.color, '仪表轴线颜色')
    optionalObject(gauge.axisTick, '仪表次刻度'); boolean(gauge.axisTick?.visible, '仪表次刻度显隐'); optionalInteger(gauge.axisTick?.splitNumber, 1, 100, '仪表每段次刻度数')
    optionalNumber(gauge.axisTick?.length, 0, 100, '仪表次刻度长度'); optionalNumber(gauge.axisTick?.width, 0, 100, '仪表次刻度线宽'); color(gauge.axisTick?.color, '仪表次刻度颜色')
    optionalObject(gauge.splitLine, '仪表主刻度'); boolean(gauge.splitLine?.visible, '仪表主刻度显隐'); optionalNumber(gauge.splitLine?.length, 0, 100, '仪表主刻度长度'); optionalNumber(gauge.splitLine?.width, 0, 100, '仪表主刻度线宽'); color(gauge.splitLine?.color, '仪表主刻度颜色')
    font(gauge.axisLabel, '仪表刻度标签'); boolean(gauge.axisLabel?.visible, '仪表刻度标签显隐'); optionalNumber(gauge.axisLabel?.distance, -100, 100, '仪表刻度标签距离'); string(gauge.axisLabel?.format, '仪表刻度数字格式')
    optionalObject(gauge.pointer, '仪表指针'); boolean(gauge.pointer?.visible, '仪表指针显隐'); optionalNumber(gauge.pointer?.length, 0, 1, '指针长度'); optionalNumber(gauge.pointer?.width, 0, 100, '指针宽度'); color(gauge.pointer?.color, '指针颜色'); enumeration(gauge.pointer?.shape, ['needle', 'slim', 'arrow'], '仪表指针形状')
    optionalObject(gauge.anchor, '仪表中心锚点'); boolean(gauge.anchor?.visible, '仪表中心锚点显隐'); optionalNumber(gauge.anchor?.size, 0, 256, '仪表中心锚点大小'); color(gauge.anchor?.color, '仪表中心锚点颜色'); color(gauge.anchor?.borderColor, '仪表中心锚点边框颜色'); optionalNumber(gauge.anchor?.borderWidth, 0, 100, '仪表中心锚点边框宽度')
    optionalObject(gauge.progress, '仪表进度弧'); boolean(gauge.progress?.visible, '仪表进度弧显隐'); boolean(gauge.progress?.rounded, '进度弧圆端'); optionalNumber(gauge.progress?.width, 0, 100, '进度弧宽度'); color(gauge.progress?.color, '进度弧颜色'); label(gauge.label, '仪表标签')
    font(gauge.title, '仪表标题'); boolean(gauge.title?.visible, '仪表标题显隐'); optionalNumber(gauge.title?.offsetY, 0, 1, '仪表标题纵向比例')
    font(gauge.detail, '仪表数值文本'); boolean(gauge.detail?.visible, '仪表数值文本显隐'); optionalNumber(gauge.detail?.offsetY, 0, 1, '仪表数值文本纵向比例'); string(gauge.detail?.format, '仪表数值数字格式')
  }
  const chord = options.chord
  if (chord) {
    optionalObject(chord.layout, '和弦图布局'); optionalObject(chord.node, '和弦图节点')
    optionalObject(chord.ribbon, '和弦图连接带'); optionalObject(chord.emphasis, '和弦图悬浮强调')
    for (const layout of [chord, chord.layout]) {
      if (!layout) continue
      optionalNumber(layout.startAngle, -Number.MAX_VALUE, Number.MAX_VALUE, '和弦图起始角'); boolean(layout.clockwise, '和弦图方向')
      optionalNumber(layout.padAngle, 0, Number.MAX_VALUE, '和弦图节点间距'); optionalNumber(layout.minAngle, 0, Number.MAX_VALUE, '和弦图最小角度')
      optionalNumber(layout.innerRadiusRatio, 0, 1, '和弦图内半径'); optionalNumber(layout.outerRadiusRatio, 0, 1, '和弦图外半径')
    }
    // Partial setters can change one radius while retaining the other authored
    // radius. Validate their relationship only after the candidate is merged.
    const inner = chord.layout?.innerRadiusRatio ?? chord.innerRadiusRatio ?? 0.58
    const outer = chord.layout?.outerRadiusRatio ?? chord.outerRadiusRatio ?? 0.78
    if (!partial && inner >= outer) throw new RangeError('和弦图内半径须小于外半径，包括未设置端使用的默认值。')
    validateChartChordColor(chord.node?.color, '和弦图节点颜色'); optionalNumber(chord.node?.opacity, 0, 1, '和弦图节点不透明度')
    validateChartChordColor(chord.node?.borderColor, '和弦图节点边框颜色'); optionalNumber(chord.node?.borderWidth, 0, Number.MAX_VALUE, '和弦图节点边框宽度')
    enumeration(chord.node?.borderType, ['solid', 'dashed', 'dotted'], '和弦图节点边框线型')
    font(chord.label, '和弦图标签', 0, Number.MAX_VALUE, validateChartChordColor); boolean(chord.label?.visible, '和弦图标签显隐')
    if (chord.label?.family !== undefined && !chord.label.family.trim()) throw new TypeError('和弦图标签字体须为非空字符串。')
    enumeration(chord.label?.position, ['inside', 'outside'], '和弦图标签位置')
    optionalNumber(chord.label?.rotation, -Number.MAX_VALUE, Number.MAX_VALUE, '和弦图标签旋转')
    optionalNumber(chord.label?.distance, 0, Number.MAX_VALUE, '和弦图标签距离')
    optionalNumber(chord.ribbonOpacity, 0, 1, '和弦带旧版不透明度'); optionalNumber(chord.ribbon?.opacity, 0, 1, '和弦带不透明度')
    enumeration(chord.ribbon?.color, ['source', 'target', 'gradient'], '和弦带颜色方式')
    optionalNumber(chord.ribbon?.width, 0, Number.MAX_VALUE, '和弦带边框宽度'); enumeration(chord.ribbon?.type, ['solid', 'dashed', 'dotted'], '和弦带边框线型')
    enumeration(chord.emphasis?.focus, ['adjacency', 'self', 'none'], '和弦图悬浮焦点')
  }
  const wordCloud = options.wordCloud
  if (wordCloud) {
    enumeration(wordCloud.shape, ['cloud', 'rect', 'circle', 'cardioid', 'heart', 'diamond', 'triangle-forward', 'triangle', 'pentagon', 'star'], '词云形状'); boolean(wordCloud.repeat, '词云重复填充')
    optionalNumber(wordCloud.minFontSize, 1, 256, '词云最小字号'); optionalNumber(wordCloud.maxFontSize, 1, 256, '词云最大字号'); optionalNumber(wordCloud.rotation, -90, 90, '词云旋转角')
    if ((wordCloud.minFontSize ?? 10) > (wordCloud.maxFontSize ?? 54)) throw new RangeError('词云最小字号不能大于最大字号，包括未设置端使用的默认值。')
  }
  boolean(options.aggregation?.aggregate, '分类聚合'); optionalObject(options.aggregation?.topN, '前 N 项')
  if (options.aggregation?.topN) {
    integer(options.aggregation.topN.count, 1, 100_000, '前 N 项数量'); boolean(options.aggregation.topN.includeOthers, '其他分类合并'); string(options.aggregation.topN.othersLabel, '其他分类名称')
  }
  optionalArray(options.trendlines, '趋势线')
  options.trendlines?.forEach(trend => {
    plain(trend, '趋势线'); line(trend, '趋势线样式'); integer(trend.seriesIndex, 0, 1_048_575, '趋势线系列字段'); requiredEnum(trend.type, ['exp', 'linear', 'log', 'movingAvg', 'poly', 'pow'], '趋势线类型')
    string(trend.name, '趋势线名称'); boolean(trend.showFormula, '趋势线公式'); boolean(trend.showR2, '趋势线拟合度')
    optionalInteger(trend.order, 2, 6, '多项式阶数'); optionalInteger(trend.period, 2, 100_000, '移动平均周期')
  })
  optionalArray(options.errorBars, '误差线')
  options.errorBars?.forEach(error => {
    plain(error, '误差线'); integer(error.seriesIndex, 0, 1_048_575, '误差线系列字段')
    requiredEnum(error.direction, ['x', 'y'], '误差线方向'); requiredEnum(error.type, ['both', 'plus', 'minus'], '误差线正负方向'); requiredEnum(error.valueType, ['fixedVal', 'percentage', 'stdDev', 'stdErr', 'cust'], '误差值类型')
    optionalNumber(error.value, 0, Number.MAX_VALUE, '误差值'); optionalInteger(error.plusIndex, 0, 1_048_575, '正误差字段'); optionalInteger(error.minusIndex, 0, 1_048_575, '负误差字段'); boolean(error.noEndCap, '无误差端帽'); line(error.line, '误差线样式')
    if (error.valueType === 'cust' && (error.type !== 'minus' && error.plusIndex === undefined || error.type !== 'plus' && error.minusIndex === undefined)) throw new RangeError('自定义误差线须为所选正负方向配置对应误差字段。')
  })
}

export function validateChartAxis(axisOptions: ChartAxisOptions | undefined, name: string): void { axis(axisOptions, name) }
export function validateChartFont(fontOptions: ChartFontOptions | undefined, name: string): void { font(fontOptions, name) }

/** Both hierarchy chart labels accept public colors, gradients and patterns. */
export function validateChartTreemapBackground(value: unknown, name = '树图标签背景'): void {
  if (value === undefined) return
  if (value && typeof value === 'object' && 'type' in value && value.type === 'pattern') {
    plain(value, `${name}图案`)
    if (typeof value.preset !== 'string' || !value.preset.trim() || value.preset.length > 128) throw new TypeError(`${name}图案需要有效预设名称。`)
    color(value.foregroundColor, `${name}图案前景色`, true); color(value.backgroundColor, `${name}图案底色`, true)
  } else validateChartColor(value, name)
}

function visitValues(options: unknown): void {
  let nodes = 0, characters = 0
  function visit(value: unknown, path: string, depth: number): void {
    if (++nodes > 200_000 || depth > 12) throw new RangeError('图表配置超出允许的大小或嵌套深度。')
    if (value === undefined || value === null || typeof value === 'boolean') return
    if (typeof value === 'string') { characters += value.length; if (characters > 1_000_000) throw new RangeError('图表配置文本超出允许大小。'); return }
    if (typeof value === 'number') { if (!Number.isFinite(value)) throw new RangeError(`${path}须为有限数值。`); return }
    if (Array.isArray(value)) { value.forEach((entry, index) => visit(entry, `${path}[${index}]`, depth + 1)); return }
    plain(value, path)
    for (const [key, entry] of Object.entries(value)) {
      if (key === '__proto__' || key === 'constructor' || key === 'prototype') throw new TypeError('图表配置包含无效属性。')
      visit(entry, `${path}.${key}`, depth + 1)
    }
  }
  visit(options, '图表', 0)
}
function font(value: ChartFontOptions | undefined, name: string, minFontSize = 1, maxFontSize = 256, validateColor = color): void {
  optionalObject(value, name)
  if (!value) return
  string(value.family, `${name}字体`); validateColor(value.color, `${name}颜色`); optionalNumber(value.fontSize, minFontSize, maxFontSize, `${name}字号`); boolean(value.bold, `${name}粗体`); boolean(value.italic, `${name}斜体`)
}
function line(value: ChartLineOptions | undefined, name: string): void {
  optionalObject(value, name)
  if (!value) return
  boolean(value.visible, `${name}显隐`); color(value.color, `${name}颜色`); optionalNumber(value.width, 0, 100, `${name}线宽`); optionalNumber(value.opacity, 0, 1, `${name}不透明度`); enumeration(value.dashType, ['solid', 'dashed', 'dotted'], `${name}线型`)
}
function label(value: ChartLabelOptions | undefined, name: string): void {
  font(value, name)
  if (!value) return
  boolean(value.visible, `${name}显隐`); enumeration(value.position, CHART_LABEL_POSITION_CHOICES.map(([position]) => position), `${name}位置`); chartLabelContentMask(value.contentType, true, `${name}内容`); string(value.format, `${name}数字格式`)
}
function axis(value: ChartAxisOptions | undefined, name: string): void {
  optionalObject(value, name)
  if (!value) return
  string(value.title, `${name}标题`); string(value.format, `${name}数字格式`); enumeration(value.scaleType, ['linear', 'log'], `${name}刻度`); boolean(value.reverse, `${name}方向`)
  font(value.titleStyle, `${name}标题字体`); enumeration(value.titleAlignment, ['start', 'center', 'end'], `${name}标题对齐`)
  optionalNumber(value.logBase, 2, 1000, `${name}对数底数`)
  optionalNumber(value.min, -Number.MAX_VALUE, Number.MAX_VALUE, `${name}最小刻度`); optionalNumber(value.max, -Number.MAX_VALUE, Number.MAX_VALUE, `${name}最大刻度`)
  enumeration(value.baseTimeUnit, ['days', 'months', 'years'], `${name}日期基准单位`)
  enumeration(value.majorTimeUnit, ['days', 'months', 'years'], `${name}日期刻度单位`)
  optionalNumber(value.majorUnit, Number.MIN_VALUE, Number.MAX_VALUE, `${name}主刻度间隔`)
  if (value.scaleType === 'log' && (value.min !== undefined && value.min <= 0 || value.max !== undefined && value.max <= 0)) throw new RangeError(`${name}对数轴刻度须为正数。`)
  if (value.min !== undefined && value.max !== undefined && value.min >= value.max) throw new RangeError(`${name}最小刻度须小于最大刻度。`)
  line(value.line, `${name}轴线`); line(value.gridLine, `${name}网格线`); font(value.label, `${name}标签`); boolean(value.label?.visible, `${name}标签显隐`); optionalNumber(value.label?.rotate, -90, 90, `${name}标签旋转`)
  line(value.tick, `${name}刻度线`); optionalNumber(value.tick?.length, 0, 50, `${name}刻度线长度`); enumeration(value.tick?.position, ['inside', 'outside'], `${name}刻度位置`)
}
function series(value: ChartSeriesOptions | undefined, name: string, allowWaterfallStyles = true): void {
  optionalObject(value, name)
  if (!value) return
  if (!allowWaterfallStyles && Object.hasOwn(value, 'waterfallStyles')) throw new TypeError(`${name}不能嵌套 waterfallStyles。`)
  if (allowWaterfallStyles) waterfallStyles(value.waterfallStyles, `${name}瀑布图角色样式`)
  string(value.name, `${name}名称`); enumeration(value.type, ['column', 'columnStacked', 'line', 'area', 'areaStacked'], `${name}图型`); seriesColor(value.color, `${name}颜色`); optionalNumber(value.fillOpacity, 0, 1, `${name}填充不透明度`); boolean(value.rightAxis, `${name}右轴`)
  fillPattern(value.fillPattern, `${name}填充纹理`)
  line(value.border, `${name}边框`); label(value.label, `${name}标签`); optionalObject(value.point, `${name}标记`)
  enumeration(value.point?.shape, CHART_POINT_SHAPES, `${name}标记形状`); optionalNumber(value.point?.size, 0, 256, `${name}标记大小`); color(value.point?.color, `${name}标记颜色`)
  if (value.dataPoints !== undefined) {
    plain(value.dataPoints, `${name}项目样式`)
    for (const [key, point] of Object.entries(value.dataPoints)) {
      recordIndex(key, `${name}项目索引`); optionalObject(point, `${name}项目 ${key}`)
      if (!point) continue
      seriesColor(point.color, `${name}项目颜色`); optionalNumber(point.fillOpacity, 0, 1, `${name}项目不透明度`); optionalNumber(point.size, 0, 256, `${name}项目大小`); line(point.border, `${name}项目边框`)
      fillPattern(point.fillPattern, `${name}项目填充纹理`)
      enumeration(point.shape, CHART_POINT_SHAPES, `${name}项目标记形状`)
    }
  }
}
function waterfallStyles(value: ChartSeriesOptions['waterfallStyles'], name: string): void {
  optionalObject(value, name)
  if (!value) return
  for (const [target, style] of Object.entries(value)) {
    requiredEnum(target, ['positive', 'negative', 'subtotal'], `${name}目标`)
    series(style, `${name}.${target}`, false)
  }
}
function fillPattern(value: ChartSeriesFillPattern | null | undefined, name: string): void {
  if (value === undefined || value === null) return
  plain(value, name)
  if (!isChartFillPatternPreset(value.preset)) throw new TypeError(`${name}包含无效的 DrawingML 纹理预设。`)
  color(value.foregroundColor, `${name}前景色`, true); color(value.backgroundColor, `${name}背景色`, true)
}
function plain(value: unknown, name: string): asserts value is Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype) throw new TypeError(`${name}须为可序列化的普通配置对象。`)
}
function optionalObject(value: unknown, name: string): void { if (value !== undefined) plain(value, name) }
function optionalArray(value: unknown, name: string): void { if (value !== undefined && !Array.isArray(value)) throw new TypeError(`${name}须为数组。`) }
function indexes(value: number[] | undefined, name: string, nonempty = false): void {
  optionalArray(value, name)
  if (!value) return
  if (nonempty && !value.length || new Set(value).size !== value.length) throw new RangeError(`${name}不能为空或包含重复索引。`)
  value.forEach(entry => integer(entry, 0, 1_048_575, name))
}
function recordIndex(key: string, name: string): void { if (!/^(0|[1-9]\d*)$/.test(key)) throw new TypeError(`${name}须为非负整数索引。`); integer(Number(key), 0, 1_048_575, name) }
function integer(value: unknown, min: number, max: number, name: string): void { if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < min || value > max) throw new RangeError(`${name}须为 ${min} 到 ${max} 的整数。`) }
function optionalInteger(value: unknown, min: number, max: number, name: string): void { if (value !== undefined) integer(value, min, max, name) }
function number(value: unknown, min: number, max: number, name: string): void { if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) throw new RangeError(`${name}须位于 ${min} 到 ${max}。`) }
function optionalNumber(value: unknown, min: number, max: number, name: string): void { if (value !== undefined) number(value, min, max, name) }
function boolean(value: unknown, name: string): void { if (value !== undefined && typeof value !== 'boolean') throw new TypeError(`${name}须为布尔值。`) }
function string(value: unknown, name: string): void { if (value !== undefined && typeof value !== 'string') throw new TypeError(`${name}须为文本。`) }
function color(value: unknown, name: string, required = false): void {
  if (value === undefined && !required) return
  if (typeof value !== 'string' || !value.trim()) throw new TypeError(`${name}须为有效颜色。`)
  chartColorChannels(value)
}
function seriesColor(value: unknown, name: string): void { if (value !== undefined) validateChartColor(value, name) }
function enumeration(value: unknown, choices: readonly string[], name: string): void { if (value !== undefined) requiredEnum(value, choices, name) }
function requiredEnum(value: unknown, choices: readonly string[], name: string): void { if (typeof value !== 'string' || !choices.includes(value)) throw new TypeError(`${name}包含无效选项。`) }
