export const CHART_COMPUTE_LIMITS = {
  maxColumns: 1024, maxSeries: 1024, maxJobs: 4096,
  runtimeAllowance: 8 * 1024 * 1024, scratchAllowance: 128 * 1024,
} as const

export interface ChartStatisticsSeries { field: number; seriesPosition: number; column: number; xColumn?: number }
interface ChartJobKey { seriesIndex: number; slot: number }
export type ChartStatisticsJob =
  | ChartJobKey & { kind: 'histogram'; binning?: { mode: 'count'; count: number } | { mode: 'width'; width: number }; underflowThreshold?: number; overflowThreshold?: number }
  | ChartJobKey & { kind: 'boxplot'; quartileMethod: 'inclusive' | 'exclusive'; whiskerMultiplier: number }
  | ChartJobKey & { kind: 'pareto'; includeZeroValues: boolean }
  | ChartJobKey & { kind: 'waterfall'; totalColumn?: number }
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
 * Waterfall: values[start,end,value], indices[categoryIndex,isTotal].
 * Trend: values[x,y]; error: values[x,y,minus,plus], indices[categoryIndex].
 */
export interface ChartStatisticsOutput {
  kind: 'chartStatistics'; values: Float64Array; indices: Uint32Array
  segments: ChartStatisticsSegment[]
}
