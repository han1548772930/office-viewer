import type { SheetAxisSpan } from './sheet-axis.ts'

export type OfficeKind = 'workbook' | 'document'

export type CellScalar = string | number | boolean | null

export interface FontStyle {
  family: string
  size: number
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  color: string | null
}

export interface CellStyle {
  font: FontStyle
  fill: string | null
  horizontal: 'general' | 'left' | 'center' | 'right' | 'justify'
  vertical: 'top' | 'middle' | 'bottom'
  wrap: boolean
  wrapStrategy?: 'overflow' | 'wrap' | 'clip'
  numberFormat: string
  border?: Record<string, unknown>
  rotation?: number
  rotationVertical?: boolean
  indent?: number
  shrinkToFit?: boolean
  metadata?: Record<string, unknown>
}

export interface SpreadsheetCell {
  id: string
  row: number
  col: number
  value: CellScalar
  formula: string | null
  display: string
  style: CellStyle
  dataType?: string
  cachedDataType?: string | null
  richText?: ReadonlyArray<Record<string, unknown>>
  hyperlink?: string | null
  sparkline?: SparklineSpec
  metadata?: Record<string, unknown>
}

export type SparklineType = 'line' | 'column' | 'winloss'
export type SparklineEmptyCells = 'gap' | 'zero' | 'connect'
export interface SparklineSpec {
  type: SparklineType
  sourceRange: string
  color: string
  negativeColor?: string
  highPointColor?: string
  lowPointColor?: string
  emptyCells: SparklineEmptyCells
  axis: 'individual' | 'custom'
  min?: number
  max?: number
}

export type OutlineAxis = 'row' | 'column'
export interface SheetOutlineGroup {
  id: string
  axis: OutlineAxis
  start: number
  end: number
  collapsed: boolean
}
export interface SheetOutlineState {
  summaryAbove: boolean
  summaryLeft: boolean
  groups: SheetOutlineGroup[]
}

export interface CellRange {
  startRow: number
  startCol: number
  endRow: number
  endCol: number
}

export interface SheetTableColumn {
  id: string
  name: string
}

export interface SheetTableModel {
  id: string
  name: string
  sheetId: string
  range: CellRange
  columns: SheetTableColumn[]
  showHeaderRow: boolean
  showTotalsRow: boolean
  styleName: string
  autoFilter?: SheetAutoFilter
}

export type SheetSortDirection = 'ascending' | 'descending'
export interface SheetSortCriterion { column: number; direction: SheetSortDirection }
export type SheetFilterConditionOperator = 'equals' | 'doesNotEqual' | 'greaterThan' | 'greaterThanOrEqual' |
  'lessThan' | 'lessThanOrEqual' | 'between' | 'notBetween' | 'contains' | 'doesNotContain' |
  'beginsWith' | 'endsWith' | 'isBlank' | 'isNotBlank' | 'isError'
export interface SheetValueFilter { kind: 'values'; selectedItems: CellScalar[] }
export interface SheetConditionFilter {
  kind: 'condition'
  operator: SheetFilterConditionOperator
  value?: CellScalar
  secondValue?: CellScalar
  caseSensitive?: boolean
}
export type SheetFilterCriteria = SheetValueFilter | SheetConditionFilter
export interface SheetAutoFilter {
  range: CellRange
  columns: Record<number, SheetFilterCriteria>
  filteredRows: number[]
}

export type SheetDataValidationType = 'any' | 'whole' | 'decimal' | 'textLength' | 'date' | 'time' | 'checkbox' | 'list' | 'listMultiple' | 'custom'
export type SheetDataValidationOperator = 'between' | 'equal' | 'greaterThan' | 'greaterThanOrEqual' |
  'lessThan' | 'lessThanOrEqual' | 'notBetween' | 'notEqual'
export type SheetDataValidationErrorStyle = 'stop' | 'warning' | 'info'
export type SheetDataValidationRenderMode = 'text' | 'arrow' | 'custom'

export interface SheetDataValidationRule {
  id: string
  type: SheetDataValidationType
  ranges: CellRange[]
  /** Retains relative formula operands when rebasing the first range would lose references. */
  formulaAnchor?: { row: number; col: number }
  operator?: SheetDataValidationOperator
  formula1?: string
  formula2?: string
  allowBlank?: boolean
  showDropDown?: boolean
  showErrorMessage?: boolean
  error?: string
  errorTitle?: string
  errorStyle?: SheetDataValidationErrorStyle
  showInputMessage?: boolean
  prompt?: string
  promptTitle?: string
  renderMode?: SheetDataValidationRenderMode
  metadata?: Record<string, unknown>
}

export type SheetProtectionEditState = 'onlyMe' | 'designedUserCanEdit'
export type SheetProtectionViewState = 'othersCanView' | 'noOneElseCanView'
export interface SheetWorksheetProtectionRule {
  id: string
  permissionId: string
  unitType: 'worksheet'
  description?: string
  ownerUserId: string
  allowedUserIds: string[]
  editState: SheetProtectionEditState
  viewState: SheetProtectionViewState
}
export interface SheetRangeProtectionRule {
  id: string
  permissionId: string
  ranges: CellRange[]
  description?: string
  ownerUserId: string
  allowedUserIds: string[]
  editState: SheetProtectionEditState
  viewState: SheetProtectionViewState
  metadata?: Record<string, unknown>
}

export type SheetConditionalFormatOperator = 'equal' | 'notEqual' | 'greaterThan' | 'greaterThanOrEqual' |
  'lessThan' | 'lessThanOrEqual' | 'between' | 'notBetween'
export interface SheetConditionalFormatStyle {
  fill?: string
  fontColor?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
}
export type SheetConditionalFormatType = 'cellIs' | 'highlightCell' | 'top10' | 'average' | 'duplicateValues' | 'uniqueValues' | 'timePeriod' | 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith' | 'expression' | 'dataBar' | 'colorScale' | 'iconSet' | 'office'
export type SheetConditionalFormatSubType = 'number' | 'text' | 'timePeriod' | 'formula' | 'rank' | 'average' | 'duplicateValues' | 'uniqueValues'
export interface SheetConditionalFormatValueConfig { type: string; value?: number | string }
export interface SheetConditionalFormatVisualConfig {
  min?: SheetConditionalFormatValueConfig
  max?: SheetConditionalFormatValueConfig
  isGradient?: boolean
  positiveColor?: string
  nativeColor?: string
  thresholds?: Array<{ value: SheetConditionalFormatValueConfig; color: string }>
  icons?: Array<{ operator?: string; value: SheetConditionalFormatValueConfig; iconType?: string; iconId?: string }>
  isShowValue?: boolean
  iconSet?: string
  reverse?: boolean
  minLength?: number
  maxLength?: number
  border?: boolean
  borderColor?: string
  negativeBorderColor?: string
  axisColor?: string
  negativeBarColorSameAsPositive?: boolean
  negativeBarBorderColorSameAsPositive?: boolean
  direction?: 'context' | 'leftToRight' | 'rightToLeft'
  axisPosition?: 'automatic' | 'middle' | 'none'
}
export interface SheetConditionalFormatFormulaAnchor { row: number; col: number }
export interface SheetConditionalFormatOfficeSource {
  version: 1
  format: 'xlsx'
  originalId: string
  storage: 'base' | 'x14'
  ruleXml: string
  groupXml?: string
  differentialFormatXml?: string
  x14Xml?: string
  themeXml?: string
  /** Only inspectable model fields; provenance and unknown JSON are excluded. */
  baseline: SheetConditionalFormatOfficeBaseline
}
export type SheetConditionalFormatOfficeBaseline = Omit<SheetConditionalFormatRule, 'id' | 'metadata' | 'officeSource'>
export interface SheetConditionalFormatRule {
  id: string
  type: SheetConditionalFormatType
  subType?: SheetConditionalFormatSubType
  ranges: CellRange[]
  /** Formula origin retained when rebasing to the first range would erase partially valid references. */
  formulaAnchor?: SheetConditionalFormatFormulaAnchor
  operator: SheetConditionalFormatOperator | string
  formula1: string
  formula2?: string
  value?: string | number | [number, number]
  isBottom?: boolean
  isPercent?: boolean
  config?: SheetConditionalFormatVisualConfig
  style: SheetConditionalFormatStyle
  priority: number
  stopIfTrue: boolean
  metadata?: Record<string, unknown>
  /** Original Office rule/style/extension identity retained independently of priority and dxf indexes. */
  officeSource?: SheetConditionalFormatOfficeSource
}
export type SheetRangeProtectionDraft = Omit<SheetRangeProtectionRule, 'id' | 'permissionId' | 'ownerUserId'> & {
  id?: string
  permissionId?: string
  ownerUserId?: string
}

export type PivotFieldArea = 'filter' | 'row' | 'column' | 'value'
export type PivotAggregation = 'average' | 'count' | 'countNums' | 'max' | 'min' | 'product' | 'stdDev' | 'stdDevp' | 'sum' | 'var' | 'varp'
export type PivotSortDirection = 'ascending' | 'descending'
export type PivotDateGroupType = 'year' | 'quarter' | 'month' | 'day'
export type PivotDateGroupingMode = 'derived' | 'excelLike'
export interface PivotNumberGroup { start: number; end: number; interval: number }
export interface PivotManualGroup { name: string; members: CellScalar[] }
export type PivotFilterOperator = 'valueGreaterThan' | 'valueGreaterThanOrEqual' | 'valueLessThan' |
  'valueLessThanOrEqual' | 'valueEqual' | 'valueNotEqual' | 'valueBetween' | 'valueNotBetween'

export interface PivotTableSource {
  unitId: string
  subUnitId: string
  sheetName: string
  range: CellRange
}

export interface PivotTableAnchor {
  unitId: string
  subUnitId: string
  row: number
  col: number
}

export interface PivotFieldFilter {
  operator: PivotFilterOperator
  expected: number
  expected2?: number
  valueFieldId: string
}

export interface PivotFieldModel {
  id: string
  sourceColumn: number
  name: string
  area: PivotFieldArea
  position: number
  displayName?: string
  aggregation?: PivotAggregation
  format?: string
  sort?: PivotSortDirection
  selectedItems?: string[]
  valueFilter?: PivotFieldFilter
  numberGroup?: PivotNumberGroup
  dateGroupType?: PivotDateGroupType
  manualGroups?: PivotManualGroup[]
}

export interface PivotTableOptions {
  repeatRowLabels: boolean
  showRowGrandTotal: boolean
  showColGrandTotal: boolean
  showRowSubtotals: boolean
  showColSubtotals: boolean
  valuePosition: 'row' | 'column'
}

export interface PivotTableModel {
  id: string
  source: PivotTableSource
  anchor: PivotTableAnchor
  fields: PivotFieldModel[]
  options: PivotTableOptions
  dateGroupingMode?: PivotDateGroupingMode
  outputRange?: CellRange
  generatedCellKeys?: string[]
  drillDownMap?: Record<string, {
    fieldFilters: Record<string, CellScalar>
    fieldMembers?: Record<string, CellScalar[]>
    valueFieldId: string
  }>
  revision: number
}

/** `bar` and `stacked-column` retain the meaning of existing Office Studio snapshots. */
export type ChartType = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'stacked-column' | 'scatter' | 'radar'
  | 'lineStacked' | 'linePercentStacked' | 'line3D'
  | 'column' | 'columnStacked' | 'columnPercentStacked' | 'column3D' | 'columnStacked3D' | 'columnPercentStacked3D'
  | 'bar-horizontal' | 'barStacked' | 'barPercentStacked' | 'bar3D' | 'barStacked3D' | 'barPercentStacked3D'
  | 'pieOfPie' | 'barOfPie' | 'pie3D' | 'areaStacked' | 'areaPercentStacked' | 'area3D' | 'areaStacked3D' | 'areaPercentStacked3D'
  | 'combination' | 'wordCloud' | 'funnel' | 'bubble' | 'relation' | 'waterfall' | 'pareto' | 'sankey' | 'heatmap'
  | 'boxplot' | 'candlestick' | 'histogram' | 'treemap' | 'sunburst' | 'gauge' | 'chord'
export type ChartSourceOrientation = 'columns' | 'rows'
export type ChartLegendPosition = 'top' | 'right' | 'bottom' | 'left'
export type ChartAxisScaleType = 'linear' | 'log'
export type ChartAxisTitleAlignment = 'start' | 'center' | 'end'

export interface ChartFontOptions { color?: string; fontSize?: number; bold?: boolean; italic?: boolean; family?: string }
export interface ChartLineOptions { visible?: boolean; color?: string; width?: number; opacity?: number; dashType?: 'solid' | 'dashed' | 'dotted' }
export interface ChartLabelOptions extends ChartFontOptions {
  visible?: boolean
  position?: 'inside' | 'outside' | 'top' | 'bottom' | 'left' | 'right' | 'center'
  contentType?: 'value' | 'category' | 'series' | 'percentage'
  format?: string
}
export type ChartFillPatternPreset = 'pct5' | 'pct10' | 'pct20' | 'pct25' | 'pct30' | 'pct40' | 'pct50' | 'pct60' | 'pct70' | 'pct75' | 'pct80' | 'pct90'
  | 'horz' | 'vert' | 'ltHorz' | 'ltVert' | 'dkHorz' | 'dkVert' | 'narHorz' | 'narVert' | 'dashHorz' | 'dashVert' | 'cross'
  | 'dnDiag' | 'upDiag' | 'ltDnDiag' | 'ltUpDiag' | 'dkDnDiag' | 'dkUpDiag' | 'wdDnDiag' | 'wdUpDiag' | 'dashDnDiag' | 'dashUpDiag' | 'diagCross'
  | 'smCheck' | 'lgCheck' | 'smGrid' | 'lgGrid' | 'dotGrid' | 'smConfetti' | 'lgConfetti' | 'horzBrick' | 'diagBrick'
  | 'solidDmnd' | 'openDmnd' | 'dotDmnd' | 'plaid' | 'sphere' | 'weave' | 'divot' | 'shingle' | 'wave' | 'trellis' | 'zigZag'
/** Omission inherits the surrounding style; null explicitly paints a solid fill. */
export interface ChartSeriesFillPattern { preset: ChartFillPatternPreset; foregroundColor: string; backgroundColor: string }
/** Source-authorable mark-local gradient, following the fixed chart-color contract. */
export interface ChartLinearGradient {
  type: 'linear-gradient'
  start: { x: number; y: number }
  end: { x: number; y: number }
  stops: Array<{ offset: number; color: string; opacity?: number }>
}
export type ChartColor = string | ChartLinearGradient
/** Public null patches reset authored fields, except fillPattern:null disables the inherited pattern. */
export type ChartNullablePatch<T> = T extends ChartSeriesFillPattern | ChartLinearGradient ? T : T extends readonly unknown[] ? T : T extends object
  ? { [K in keyof T]?: ChartNullablePatch<T[K]> | null } : T
export type ChartRadarShape = 'polygon' | 'circle'
export interface ChartRadarAxisLabelOptions extends ChartFontOptions { visible?: boolean; format?: string }
export interface ChartRadarOptions { shape?: ChartRadarShape; fill?: boolean; axisLabel?: ChartRadarAxisLabelOptions }
export interface ChartSeriesOptions {
  name?: string
  type?: 'column' | 'columnStacked' | 'line' | 'area' | 'areaStacked'
  color?: ChartColor
  fillOpacity?: number
  fillPattern?: ChartSeriesFillPattern | null
  border?: ChartLineOptions
  point?: { shape?: 'circle' | 'rect' | 'triangle' | 'diamond' | 'none'; size?: number; color?: string }
  rightAxis?: boolean
  label?: ChartLabelOptions
  dataPoints?: Record<number, { color?: ChartColor; fillOpacity?: number; fillPattern?: ChartSeriesFillPattern | null; size?: number; shape?: 'circle' | 'rect' | 'triangle' | 'diamond' | 'none'; border?: ChartLineOptions }>
}
/** Numeric selectors keep the source-field contract; names match source identities exactly. */
export type ChartSeriesSelector = number | { name: string }
export interface ChartNamedSeriesOptions { name: string; style: ChartSeriesOptions }
export interface ChartBubbleMappingOptions { xIndex: number; yIndex: number; sizeIndex?: number; labelIndex?: number; seriesIndex?: number }
export interface ChartAxisPointerOptions {
  indicatorLabelColor?: string
  indicatorLineType?: 'solid' | 'dashed' | 'dotted'
  indicatorLineColor?: string
  indicatorLabelTextColor?: string
}
export interface ChartTrendlineOptions extends ChartLineOptions {
  seriesIndex: number
  type: 'exp' | 'linear' | 'log' | 'movingAvg' | 'poly' | 'pow'
  name?: string
  order?: number
  period?: number
  showFormula?: boolean
  showR2?: boolean
}
export interface ChartErrorBarOptions {
  seriesIndex: number
  direction: 'x' | 'y'
  type: 'both' | 'plus' | 'minus'
  valueType: 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'cust'
  value?: number
  plusIndex?: number
  minusIndex?: number
  noEndCap?: boolean
  line?: ChartLineOptions
}
export type ChartGaugeMode = 'modern' | 'classic' | 'pointer' | 'progress'
export type ChartGaugePointerShape = 'needle' | 'slim' | 'arrow'
export interface ChartGaugeScaleOptions { min?: number; max?: number; splitNumber?: number }
export interface ChartGaugeAnglesOptions { startAngle?: number; endAngle?: number }
export interface ChartGaugeRangeOptions { from: number; to: number; color: string }
export interface ChartGaugeAxisLineOptions { visible?: boolean; width?: number; color?: string }
export interface ChartGaugeAxisTickOptions extends ChartGaugeAxisLineOptions { splitNumber?: number; length?: number }
export interface ChartGaugeSplitLineOptions extends ChartGaugeAxisLineOptions { length?: number }
export interface ChartGaugeAxisLabelOptions extends ChartFontOptions { visible?: boolean; distance?: number; format?: string }
export interface ChartGaugePointerOptions { visible?: boolean; length?: number; width?: number; color?: string; shape?: ChartGaugePointerShape }
export interface ChartGaugeAnchorOptions { visible?: boolean; size?: number; color?: string; borderColor?: string; borderWidth?: number }
export interface ChartGaugeProgressOptions { visible?: boolean; width?: number; color?: string; rounded?: boolean }
export interface ChartGaugeTitleOptions extends ChartFontOptions { visible?: boolean; offsetY?: number }
export interface ChartGaugeDetailOptions extends ChartGaugeTitleOptions { format?: string }
/** Facade patches use null to reset individual authored values. The model saves undefined resets. */
export type ChartGaugeNullablePatch<T> = { [K in keyof T]?: T[K] | null }
export interface ChartGaugeTicksPatch { majorDivisions?: number | null; minorDivisionsPerMajor?: number | null }
export interface ChartGaugeOptions extends ChartGaugeScaleOptions, ChartGaugeAnglesOptions {
  /** Modern/Classic follow the fixed public presets; pointer/progress preserve local historical modes. */
  mode?: ChartGaugeMode
  ranges?: ChartGaugeRangeOptions[]
  axisLine?: ChartGaugeAxisLineOptions
  axisTick?: ChartGaugeAxisTickOptions
  splitLine?: ChartGaugeSplitLineOptions
  axisLabel?: ChartGaugeAxisLabelOptions
  pointer?: ChartGaugePointerOptions
  anchor?: ChartGaugeAnchorOptions
  progress?: ChartGaugeProgressOptions
  title?: ChartGaugeTitleOptions
  detail?: ChartGaugeDetailOptions
  /** Historical shared title/detail font and content configuration. Dedicated groups take precedence. */
  label?: ChartLabelOptions
}
/** Fractions of the entire chart frame, independent of the renderer and axis labels. */
export interface ChartPlotAreaOptions { left: number; top: number; width: number; height: number }
/** Renderer-independent authored configuration. Derived statistics and geometries are not saved here. */
export interface ChartSemanticOptions {
  invalidValueStrategy?: 'zero' | 'break' | 'link'
  plotVisibleOnly?: boolean
  autoGradientFill?: boolean
  plotArea?: ChartPlotAreaOptions
  appearance?: { background?: string; border?: ChartLineOptions; font?: ChartFontOptions }
  titleStyle?: ChartFontOptions
  subtitleStyle?: ChartFontOptions
  allSeriesStyle?: ChartSeriesOptions
  /** Keys identify source fields, so a series keeps its style after other series are removed. */
  series?: Record<number, ChartSeriesOptions>
  /** Applied after field overrides. Display-name changes do not rename these source selectors. */
  namedSeries?: ChartNamedSeriesOptions[]
  axisPointer?: ChartAxisPointerOptions
  rightYAxis?: ChartAxisOptions
  /** Categorical series-depth axis, authored independently of the 3D camera. */
  depthAxis?: ChartDepthAxisOptions
  trendlines?: ChartTrendlineOptions[]
  errorBars?: ChartErrorBarOptions[]
  mapping?: {
    /** Use elapsed date serials on the category axis; the date system belongs to the workbook. */
    useDateAxis?: boolean
    categoryIndexes?: number[]
    hierarchyIndexes?: number[]
    bubble?: ChartBubbleMappingOptions
    edge?: { sourceIndex: number; targetIndex: number; valueIndex?: number }
    candlestick?: { openIndex?: number; highIndex: number; lowIndex: number; closeIndex: number }
    gauge?: { valueIndex: number; labelIndex?: number }
  }
  threeD?: { view?: { rotX?: number; rotY?: number; perspective?: number; rightAngleAxes?: boolean; heightPercent?: number; depthPercent?: number }; grouping?: 'standard' | 'clustered'; gapDepth?: number }
  bar?: { gapWidth?: number; overlap?: number }
  area?: { smooth?: boolean }
  pie?: {
    startAngle?: number; doughnutHole?: number; explosion?: number; rosePie?: boolean; isHalfPie?: boolean
    showLabelLine?: boolean; label?: ChartLabelOptions
    secondaryPlot?: { type?: 'pie' | 'bar'; size?: number; gapWidth?: number; split?: { type: 'auto' | 'position' | 'value' | 'percentage' | 'custom'; count?: number; lessThan?: number; pointIndexes?: number[] } }
    aggregateLabel?: string
  }
  radar?: ChartRadarOptions
  bubble?: { scale?: number; sizeRepresents?: 'area' | 'width' }
  funnel?: { gap?: number; useAbsoluteValue?: boolean; sort?: 'ascending' | 'descending' | 'none' }
  relation?: { layout?: 'none' | 'circular' | 'force'; useValueAsSymbolSize?: boolean; force?: { repulsion?: number; gravity?: number; edgeLength?: number } }
  waterfall?: { pointRoles?: Array<{ seriesIndex: number; dataPointIndex: number; role: 'change' | 'total' }>; connector?: boolean; positiveColor?: string; negativeColor?: string; totalColor?: string }
  pareto?: { includeZeroValues?: boolean }
  histogram?: { binning?: { mode: 'count'; count: number } | { mode: 'width'; width: number }; underflowThreshold?: number; overflowThreshold?: number; binGap?: number }
  boxplot?: { quartileMethod?: 'inclusive' | 'exclusive'; showOutliers?: boolean; whiskerMultiplier?: number }
  candlestick?: { candleWidth?: number; risingColor?: string; fallingColor?: string; hollow?: boolean }
  heatmap?: { visualMapType?: 'continuous' | 'piecewise'; min?: number; max?: number; unit?: string; label?: ChartLabelOptions }
  treemap?: { parentLabelLayout?: 'banner' | 'overlapping' | 'none'; label?: ChartLabelOptions }
  sunburst?: { label?: ChartLabelOptions; centerLabel?: { visible?: boolean; font?: ChartFontOptions } }
  gauge?: ChartGaugeOptions
  chord?: { startAngle?: number; clockwise?: boolean; padAngle?: number; innerRadiusRatio?: number; outerRadiusRatio?: number; ribbonOpacity?: number; label?: ChartLabelOptions }
  wordCloud?: { shape?: 'cloud' | 'rect' | 'circle' | 'cardioid' | 'heart' | 'diamond' | 'triangle-forward' | 'triangle' | 'pentagon' | 'star'; repeat?: boolean; minFontSize?: number; maxFontSize?: number; rotation?: number }
  aggregation?: { aggregate?: boolean; topN?: { count: number; includeOthers?: boolean; othersLabel?: string } }
}

export interface ChartSource {
  range: CellRange
  orientation: ChartSourceOrientation
  /** Data source worksheet, independent of the host. Older snapshots default to ChartModel.sheetId. */
  sheetId?: string
}

export interface ChartAxisOptions {
  title?: string
  titleStyle?: ChartFontOptions
  titleAlignment?: ChartAxisTitleAlignment
  scaleType: ChartAxisScaleType
  logBase?: number
  min?: number
  max?: number
  reverse?: boolean
  line?: ChartLineOptions
  label?: ChartFontOptions & { visible?: boolean; rotate?: number }
  gridLine?: ChartLineOptions
  tick?: ChartLineOptions & { length?: number; position?: 'inside' | 'outside' }
  format?: string
}

export interface ChartLegendOptions {
  visible: boolean
  position: ChartLegendPosition
  font?: ChartFontOptions
}

export type ChartDepthAxisOptions = Omit<ChartAxisOptions, 'scaleType' | 'logBase' | 'min' | 'max' | 'format'>

export interface ChartModel {
  id: string
  sheetId: string
  type: ChartType
  source: ChartSource
  position: { x: number; y: number }
  size: { width: number; height: number }
  title?: string
  subtitle?: string
  palette?: string[]
  options?: ChartSemanticOptions
  legend: ChartLegendOptions
  categoryField: number
  valueFields: number[]
  seriesToCategory?: Record<number, number[]>
  pivotTableId?: string
  xAxis: ChartAxisOptions
  yAxis: ChartAxisOptions
  revision: number
}

export type SlicerType = 'slicer' | 'timeline'
export type TimelineLevel = 'day' | 'month' | 'quarter' | 'year'

export interface SlicerTimelineRange {
  start?: CellScalar
  end?: CellScalar
  level: TimelineLevel
}

/** A persisted filter control bound to one Pivot source field or stable Table column. */
export interface SlicerModel {
  id: string
  unitId: string
  hostSheetId: string
  pivotTableId?: string
  tableId?: string
  sourceFieldId: string
  type: SlicerType
  caption: string
  name?: string
  position: { row: number; column: number }
  size: { rows: number; columns: number }
  layout?: { left: number; top: number; width: number; height: number }
  drawing?: {
    anchorType: 'both' | 'position' | 'none'
    from: { row: number; column: number; offsetX: number; offsetY: number }
    to: { row: number; column: number; offsetX: number; offsetY: number }
    keepRatio?: boolean
    zIndex?: number
  }
  style?: string
  startItem?: number
  columnCount?: number
  rowHeight?: number
  showCaption?: boolean
  lockedPosition?: boolean
  level?: TimelineLevel
  selectionLevel?: TimelineLevel
  showHeader?: boolean
  showSelectionLabel?: boolean
  showTimeLevel?: boolean
  showHorizontalScrollbar?: boolean
  scrollPosition?: number
  selectedItems?: string[]
  timelineRange?: SlicerTimelineRange
  revision: number
}

export interface SpreadsheetSheet {
  id: string
  name: string
  visible: boolean
  tabColor?: string
  cells: Record<string, SpreadsheetCell>
  mergedRanges: CellRange[]
  columnWidths: Record<number, number>
  rowHeights: Record<number, number>
  /** All editable axis geometry is measured in CSS pixels. Older snapshots may omit this marker. */
  axisUnit?: 'cssPixel'
  rowHeightSpans?: SheetAxisSpan<number>[]
  columnWidthSpans?: SheetAxisSpan<number>[]
  /** Undefined rows with no custom height use automatic height. */
  rowAutoHeightSpans?: SheetAxisSpan<boolean>[]
  /** Measured heights of content-bearing automatic rows, separate from manual heights. */
  rowAutoHeights?: Record<number, number>
  frozenRows: number
  frozenCols: number
  showGridlines: boolean
  hiddenRows?: number[]
  /** Compact manual hides; legacy point arrays remain readable and are combined with these ranges. */
  rowHiddenSpans?: SheetAxisSpan<boolean>[]
  columnHiddenSpans?: SheetAxisSpan<boolean>[]
  outline?: SheetOutlineState
  autoFilter?: SheetAutoFilter
  dataValidations?: SheetDataValidationRule[]
  conditionalFormats?: SheetConditionalFormatRule[]
  rangeProtectionRules?: SheetRangeProtectionRule[]
  worksheetProtection?: SheetWorksheetProtectionRule
  hiddenColumns?: number[]
  defaultColumnWidth?: number
  defaultRowHeight?: number
  defaultRowAutoHeight?: boolean
  rowCount?: number
  columnCount?: number
  zoom?: number
  metadata?: Record<string, unknown>
}

export interface WorkbookModel {
  kind: 'workbook'
  name: string
  calculationMode?: 'automatic' | 'manual'
  sheets: SpreadsheetSheet[]
  activeSheetId: string
  date1904: boolean
  modified: boolean
  id?: string
  metadata?: Record<string, unknown>
  pivotTables?: PivotTableModel[]
  tables?: SheetTableModel[]
  charts?: ChartModel[]
  slicers?: SlicerModel[]
}

export type DocumentSectionBreakType = 'unspecified' | 'nextPage' | 'continuous' | 'nextColumn' | 'evenPage' | 'oddPage'
export type DocumentBlockType = 'paragraph' | 'heading' | 'bullet' | 'numbered' | 'checklist' | 'quote' | 'divider' | 'table' | 'image' | 'pageBreak' | 'columnBreak' | 'sectionBreak'

export interface InlineRun {
  id: string
  text: string
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  color: string
  fontFamily: string
  eastAsiaFontFamily?: string
  fontSize: number
  verticalAlign?: 'baseline' | 'subscript' | 'superscript'
  highlight?: string
  metadata?: Record<string, unknown>
}

export interface DocumentParagraphStyle {
  align: 'left' | 'center' | 'right' | 'justify'
  indent: number
  lineHeight: number
  spaceBefore: number
  spaceAfter: number
  indentStart?: number
  indentEnd?: number
  indentFirstLine?: number
  hanging?: number
  spacingRule?: 'auto' | 'atLeast' | 'exact'
  wordWrap?: boolean
  snapToGrid?: boolean
}

/** Resolved document defaults, independent of paragraph and direct run formatting. */
export interface DocumentDefaultStyle {
  version: 1
  text: Pick<InlineRun, 'bold' | 'italic' | 'underline' | 'strike' | 'color' | 'fontFamily' | 'eastAsiaFontFamily' | 'fontSize' | 'verticalAlign'> & { highlight?: string }
  paragraph: DocumentParagraphStyle
}

export interface DocumentSectionSetup {
  marginHeader?: number
  marginFooter?: number
  width?: number
  height?: number
  orientation?: 'portrait' | 'landscape'
  marginTop?: number
  marginRight?: number
  marginBottom?: number
  marginLeft?: number
  columnCount?: number
  columnGap?: number
  columnSeparator?: 'none' | 'line'
  sectionStart?: DocumentSectionBreakType
  pageNumberStart?: number
}

export interface DocumentBlock {
  id: string
  type: DocumentBlockType
  level: number
  runs: InlineRun[]
  style: DocumentParagraphStyle
  metadata: Record<string, unknown>
  checklist?: { listId: string; checked: boolean }
  list?: { id: string; preset: keyof typeof import('./document-list-presets.json') }
  sectionBreak?: { type: DocumentSectionBreakType; setup?: DocumentSectionSetup }
}

export type DocumentContentControlType =
  | 'richText'
  | 'text'
  | 'checkbox'
  | 'date'
  | 'dropDownList'
  | 'comboBox'
  | 'picture'
  | 'group'
  | 'repeatingSection'
  | 'repeatingSectionItem'
  | 'docPartObj'

export type DocumentContentControlPlacement = 'inline' | 'block' | 'table' | 'row' | 'cell' | 'header' | 'footer'

export interface DocumentContentControlListItem {
  displayText?: string
  value: string
}

/**
 * A persisted SDT/content-control range. Offsets are UTF-16 offsets within
 * the referenced paragraph block, so edits do not depend on a rendered DOM.
 */
export interface DocumentContentControl {
  id: string
  type: DocumentContentControlType
  blockId: string
  startOffset: number
  endOffset: number
  placement?: DocumentContentControlPlacement
  alias?: string
  title?: string
  tag?: string
  appearance?: string
  color?: string
  placeholder?: string
  placeholderDocPart?: { name?: string; text?: string }
  showingPlaceholder?: boolean
  temporary?: boolean
  lock?: 'unlocked' | 'controlLocked' | 'contentLocked' | 'sdtLocked' | 'sdtContentLocked'
  value?: string
  checked?: boolean
  checkbox?: { checked: boolean; checkedState?: { value?: string; font?: string }; uncheckedState?: { value?: string; font?: string } }
  multiLine?: boolean
  dateValue?: string
  dateFormat?: string
  dateLocale?: string
  dateCalendar?: string
  dateStorage?: string
  items?: DocumentContentControlListItem[]
  parentId?: string
  repeatingSection?: { sectionTitle?: string; doNotAllowInsertDelete?: boolean }
  docPart?: { gallery?: string; category?: string; unique?: boolean }
  dataBinding?: { prefixMappings?: string; xpath?: string; storeItemID?: string; cachedResult?: string; dirty?: boolean }
  rawSdtPr?: string
  rawSdtEndPr?: string
  metadata?: Record<string, unknown>
}

export interface DocumentPageSetup {
  width: number
  height: number
  marginTop: number
  marginRight: number
  marginBottom: number
  marginLeft: number
}

export interface DocumentSectionHeaderFooter {
  header: string
  footer: string
  firstPageHeader: string
  firstPageFooter: string
  evenPageHeader: string
  evenPageFooter: string
  pageFields?: DocumentPageField[]
  /** Rich formatting over the corresponding literal text; fields remain separate objects. */
  storyContent?: Partial<Record<DocumentPageField['story'], DocumentStoryContent>>
  differentFirstPage: boolean
  differentOddEven: boolean
  /** Per-story association; omitted keys fall back to the legacy whole-section flag. */
  storyLinkToPrevious?: Partial<Record<DocumentPageField['story'], boolean>>
  /** Legacy whole-section association; mixed per-story state aggregates to false. */
  linkToPrevious: boolean
}

export interface DocumentStoryParagraph {
  /** Half-open UTF-16 offsets in literal text, excluding the following paragraph separator. */
  startOffset: number
  endOffset: number
  style: DocumentParagraphStyle
  metadata?: Record<string, unknown>
}

export interface DocumentStoryContent {
  /** When present, this block tree is authoritative; runs/paragraphs are compatibility projections. */
  blocks?: DocumentBlock[]
  /** Concatenated run text must equal the section's corresponding literal story string. */
  runs: InlineRun[]
  paragraphs: DocumentStoryParagraph[]
  locks?: DocumentStoryLock[]
  metadata?: Record<string, unknown>
}

export interface DocumentStoryLock {
  id: string
  /** Input-buffer UTF-16 offsets; page fields each occupy one U+FFFC. */
  startOffset: number
  endOffset: number
  lock: 'contentLocked' | 'sdtLocked' | 'sdtContentLocked'
}

export interface DocumentPageField {
  story: 'header' | 'footer' | 'firstPageHeader' | 'firstPageFooter' | 'evenPageHeader' | 'evenPageFooter'
  field: 'PAGE' | 'NUMPAGES' | 'DATE' | 'TIME'
  /** UTF-16 offset in the story's literal text (field results are not counted). */
  offset?: number
  instruction?: string
  cachedResult?: string
  /** Native field lock prevents automatic refresh; it is separate from an SDT content lock. */
  locked?: boolean
  /** An outer field containing another custom range retains its imported result. */
  containsNestedFields?: boolean
  textStyle?: DocumentDefaultStyle['text']
  /** Parser provenance and source-format replay policy for this field result. */
  metadata?: Record<string, unknown>
}

export interface DocumentNote {
  id: string
  text: string
  /** Rich note body. Its blocks, when present, are authoritative over the plain text projection. */
  content?: DocumentStoryContent
  referenceBlockId: string
  referenceRunId: string
}

export interface DocumentCommentAnchorSegment {
  blockId: string
  startOffset: number
  endOffset: number
  text: string
}

export interface DocumentCommentMessage {
  id: string
  authorId: string
  authorName: string
  text: string
  createdAt: string
  updatedAt?: string
  mentions?: Array<{ userId: string; label: string }>
}

export interface DocumentThreadComment {
  id: string
  anchor: DocumentCommentAnchorSegment[]
  referenceText: string
  messages: DocumentCommentMessage[]
  resolved: boolean
  createdAt: string
  updatedAt: string
  metadata?: Record<string, unknown>
}

export type DocumentPermissionAction = 'edit' | 'copy' | 'comment' | 'print' | 'export'
export type DocumentPermissionScope = 'document' | 'section' | 'paragraph' | 'entity'

export interface DocumentPermissionRule {
  id: string
  scope: DocumentPermissionScope
  targetId?: string
  action: DocumentPermissionAction
  effect: 'allow' | 'deny'
  userIds: string[]
}

export interface DocumentModel {
  kind: 'document'
  name: string
  blocks: DocumentBlock[]
  page: DocumentPageSetup
  sectionSetup?: DocumentSectionSetup
  modified: boolean
  wordCount: number
  defaultStyle?: DocumentDefaultStyle
  contentControls?: DocumentContentControl[]
  sectionHeaderFooters?: Record<string, DocumentSectionHeaderFooter>
  permissionRules?: DocumentPermissionRule[]
  footnotes?: Record<string, DocumentNote>
  endnotes?: Record<string, DocumentNote>
  comments?: Record<string, DocumentThreadComment>
  id?: string
  metadata?: Record<string, unknown>
}

export interface ParserDiagnostic {
  code: string
  severity: 'info' | 'warning' | 'error'
  message: string
  details?: string
}

export interface EditorDocumentState {
  activeKind: OfficeKind
  workbook: WorkbookModel
  document: DocumentModel
  fileName: string
  parserStatus: 'idle' | 'loading' | 'ready' | 'error'
  parserMessage: string
  diagnostics: ParserDiagnostic[]
  selectedRange: CellRange
}

export const defaultCellStyle = (): CellStyle => ({
  font: {
    family: 'Aptos',
    size: 11,
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    color: null,
  },
  fill: null,
  horizontal: 'general',
  vertical: 'middle',
  wrap: false,
  rotation: 0,
  rotationVertical: false,
  indent: 0,
  shrinkToFit: false,
  numberFormat: 'General',
})

export const createCell = (row: number, col: number, value: CellScalar = ''): SpreadsheetCell => ({
  id: `${row}:${col}`,
  row,
  col,
  value,
  formula: null,
  display: value === null ? '' : String(value),
  style: defaultCellStyle(),
})

export const createBlankSheet = (name = 'Sheet1'): SpreadsheetSheet => ({
  id: `sheet-${crypto.randomUUID?.() ?? Date.now()}`,
  name,
  visible: true,
  cells: {},
  mergedRanges: [],
  columnWidths: {},
  rowHeights: {},
  axisUnit: 'cssPixel',
  frozenRows: 0,
  frozenCols: 0,
  showGridlines: true,
})

export const createBlankWorkbook = (): WorkbookModel => {
  const sheet = createBlankSheet()
  return {
    kind: 'workbook',
    name: '未命名工作簿',
    sheets: [sheet],
    activeSheetId: sheet.id,
    date1904: false,
    modified: false,
    pivotTables: [],
  }
}

export const createDocumentDefaultStyle = (): DocumentDefaultStyle => ({
  version: 1,
  text: { fontFamily: 'Aptos', eastAsiaFontFamily: 'Aptos', fontSize: 11, color: '#1f2937', bold: false, italic: false, underline: false, strike: false },
  paragraph: { align: 'left', indent: 0, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 8 },
})

export const createBlankDocument = (): DocumentModel => ({
  kind: 'document',
  name: '未命名文档',
  defaultStyle: createDocumentDefaultStyle(),
  blocks: [{
    id: `block-${Date.now()}`,
    type: 'paragraph',
    level: 0,
    runs: [{
      id: `run-${Date.now()}`,
      text: '',
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      color: '#1f2937',
      fontFamily: 'Aptos',
      fontSize: 11,
    }],
    style: {
      align: 'left',
      indent: 0,
      lineHeight: 1.5,
      spaceBefore: 0,
      spaceAfter: 8,
    },
    metadata: {},
  }],
  page: {
    width: 794,
    height: 1123,
    marginTop: 72,
    marginRight: 72,
    marginBottom: 72,
    marginLeft: 72,
  },
  modified: false,
  wordCount: 0,
  contentControls: [],
})
