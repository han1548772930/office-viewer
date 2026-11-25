export const CHART_COMPUTE_LIMITS = {
  maxColumns: 1024, maxSeries: 1024, maxJobs: 4096,
  runtimeAllowance: 8 * 1024 * 1024, scratchAllowance: 128 * 1024,
} as const

export interface ChartStatisticsSeries { field: number; seriesPosition: number; column: number; xColumn?: number }
interface ChartJobKey { seriesIndex: number; slot: number }
export interface ChartWordCloudGlyph {
  width: number; height: number; offset: number; length: number; fontSize: number; weight: number
}
export type ChartStatisticsJob =
  | ChartJobKey & { kind: 'histogram'; binning?: { mode: 'count'; count: number } | { mode: 'width'; width: number }; underflowThreshold?: number; overflowThreshold?: number }
  | ChartJobKey & { kind: 'boxplot'; quartileMethod: 'medianOfHalves' | 'inclusive' | 'exclusive'; whiskerMethod?: 'range' | 'iqr'; whiskerMultiplier: number }
  | ChartJobKey & { kind: 'pareto'; includeZeroValues: boolean }
  | ChartJobKey & { kind: 'hierarchy'; hierarchyColumns: number[]; labelColumns: number[]; keyCount: number; labelCount: number }
  | ChartJobKey & { kind: 'chord'; sourceColumn: number; targetColumn: number; labelCount: number; startAngle: number; clockwise: boolean; padAngle: number; minAngle: number }
  | ChartJobKey & { kind: 'wordCloud'; glyphColumn: number; maskColumn: number; width: number; height: number; glyphs: ChartWordCloudGlyph[]; maxAttempts: number; minFontSize: number }
  | ChartJobKey & { kind: 'relationLayout'; seedColumn: number; sourceColumn: number; targetColumn: number; linkValueColumn: number;
    nodeCount: number; linkCount: number; width: number; height: number; layout: 'none' | 'circular' | 'force';
    useValueAsSymbolSize: boolean; symbolSize: number; repulsion: number; gravity: number; edgeLength: number }
  | ChartJobKey & { kind: 'waterfall'; totalColumn?: number }
  | ChartJobKey & { kind: 'waterfallLayout'; seriesIndices: number[]; stackType: 'arrangement' | 'stacked'; useSubtotal: boolean; totalColumns?: Array<number | null> }
  | ChartJobKey & { kind: 'trendline'; type: 'exp' | 'linear' | 'log' | 'movingAvg' | 'poly' | 'pow'; order?: number; period?: number }
  | ChartJobKey & { kind: 'errorBars'; direction: 'x' | 'y'; type: 'both' | 'plus' | 'minus'; valueType: 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'cust'; value?: number; plusColumn?: number; minusColumn?: number }

export interface ChartStatisticsParams {
  rowCount: number; columnCount: number; categoryCount: number
  /** Job.seriesIndex addresses this array. field identifies the source field;
   * seriesPosition identifies the prepared series, including separate groups
   * that share a Y field and have different X/sample columns. */
  series: ChartStatisticsSeries[]; jobs: ChartStatisticsJob[]
}
export interface ChartStatisticsTask {
  kind: 'chartStatistics'
  /** Column-major retained numeric/state input. categoryIds addresses the main-thread dictionary. */
  values: Float64Array; states: Uint8Array; categoryIds: Uint32Array
  params: ChartStatisticsParams
}
export interface ChartStatisticsFit {
  coefficients: number[]
  xScale: number; xMean: number; xSpread: number; yScale: number; yMean: number
}
export interface ChartStatisticsSegment {
  kind: ChartStatisticsJob['kind']; jobIndex: number; field: number; seriesPosition: number; slot: number
  valueOffset: number; valueLength: number; indexOffset: number; indexLength: number
  validSampleCount: number; available: boolean; fit?: ChartStatisticsFit; r2?: number
}
/** Hist: values[lower,upper], indices[kind,count] (0=underflow,1=regular,2=overflow).
 * Box: values[min,q1,median,q3,max,...outliers].
 * Pareto: values[value,cumulativePercent], indices[dictionaryId].
 * Hierarchy: values[value,selfValue]; indices[N,R,...N*(parent,key,label,sourceOrder),...R*(rootIndex)].
 * The root parent is 0xffffffff; nodes use sorted preorder and rootIndex uses first-source order.
 * Chord: values[N*(value,start,end),...E*(value,sourceStart,sourceEnd,targetStart,targetEnd)];
 * indices[N,E,...N*(displayLabelId),...E*(sourceNodeIndex,targetNodeIndex)]. Angles are radians,
 * in first-valid-source order. Undirected duplicates keep the first endpoint direction.
 * Strings and pixel geometry stay with the owner; only numeric IDs and normalized angles transfer.
 * WordCloud: values[centerX,centerY,scale], indices[glyphIndex]. The glyph and
 * mask inputs use grid cells; placement is decoded with the retained grid size.
 * RelationLayout: values[x,y,symbolSize], indices[nodeIndex]. Positions are
 * relative to the plot; node values, UTF-16 name seeds and link endpoint/value
 * columns have separate counts. The retained graph owns names and link identity.
 * Waterfall: values[start,end,value], indices[categoryIndex,isTotal].
 * WaterfallLayout: values[finalCumulative,...N*(start,end,value),...C*(fromValue,toValue)];
 * indices[N,C,...N*(seriesPosition,dataPointIndex,visibleCategoryIndex,flags),...C*(fromCategory,toCategory)].
 * flags bit0=total, bit1=synthetic. A stacked global subtotal has seriesPosition=0xffffffff.
 * Trend: values[x,y]; error: values[x,y,minus,plus], indices[categoryIndex].
 */
export interface ChartStatisticsOutput {
  kind: 'chartStatistics'; values: Float64Array; indices: Uint32Array
  segments: ChartStatisticsSegment[]
}
