import type { SheetAxisSpan } from './sheet-axis.ts'
import type { SheetBackgroundImage, SheetCellImage, SheetFloatingImage } from './sheet-images.ts'
import type { DrawingAffineGeometry } from './drawing-affine.ts'
import type { SheetDrawingShapeDTO, SheetDrawingState } from './sheet-drawings.ts'
import type { PivotRecordValue } from './record-value.ts'
import type { ISlicerStylesDefinition, ITimelineStylesDefinition } from './slicer-styles.ts'
import type { ITimelineRange, ITimelineDateCondition } from './timeline-filters.ts'
import type { OfficeSparklineSource } from './sparkline-office-source.ts'
import type { SparklineSourceReference } from './sparkline-source.ts'
export type { SheetImageResource, SheetCellImage, SheetFloatingImage, SheetBackgroundImage } from './sheet-images.ts'
export type { RecordValue, RecordSelection, PivotRecordValue } from './record-value.ts'
export type { ISlicerStylesDefinition, ITimelineStylesDefinition } from './slicer-styles.ts'
export type { ITimelineRange, ITimelineDateCondition } from './timeline-filters.ts'
export type { OfficeSparklineSource } from './sparkline-office-source.ts'
export type { SparklineBrokenSourceRange, SparklineSourceReference } from './sparkline-source.ts'

export type OfficeKind = 'workbook' | 'document'

/** Univer v1.6.0 core/sheets/sheet-snapshot-utils.ts. Grid extents are
 * worksheet state, independent of the last populated cell or fill target. */
export const DEFAULT_SHEET_ROW_COUNT = 1000
export const DEFAULT_SHEET_COLUMN_COUNT = 20

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

/** Sparse direct formatting; omitted properties continue to inherit. */
export type SheetCellStylePatch = Omit<Partial<CellStyle>, 'font'> & { font?: Partial<FontStyle> }
export type SheetStyleSource = SheetCellStylePatch | string | null
export type SheetCellStyleProperty = Exclude<keyof CellStyle, 'font'> | `font.${keyof FontStyle}`

export interface SpreadsheetCell {
  id: string
  row: number
  col: number
  value: CellScalar
  formula: string | null
  display: string
  style: CellStyle
  /** Persisted materialization marker; a saved cache must never become newly authored formatting. */
  styleMaterialized?: true
  /** Authoritative sparse formatting. Absence preserves legacy fully specified cell styles. */
  authoredStyle?: SheetCellStylePatch
  /** Public cell.s catalog reference; direct edits keep unrelated referenced properties live. */
  authoredStyleReference?: string
  /** Explicit resets suppress these properties of the referenced direct style. */
  clearedStyleProperties?: SheetCellStyleProperty[]
  /** Range-theme formatting is composed before direct cell formatting. */
  themeStyle?: SheetStyleSource
  dataType?: string
  cachedDataType?: string | null
  richText?: ReadonlyArray<Record<string, unknown>>
  hyperlink?: string | null
  sparkline?: SparklineSpec
  image?: SheetCellImage
  metadata?: Record<string, unknown>
}

export type SparklineType = 'line' | 'column' | 'winloss' | 'pie'
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
  groupId?: string
  themeType?: 'standard' | 'specialty' | 'distinctive' | 'soft' | 'fresh' | 'custom'
  points?: Partial<Record<'highPoint' | 'firstPoint' | 'lowPoint' | 'lastPoint' | 'negativePoint' | 'markersPoint', { visible?: boolean; color?: string }>>
  axisOptions?: { visible?: boolean; reverse?: boolean; color?: string }
  extremumMin?: { type: 'individual' | 'group' | 'custom'; value?: number }
  extremumMax?: { type: 'individual' | 'group' | 'custom'; value?: number }
  nonNumericCells?: SparklineEmptyCells
  containHiddenCells?: boolean
  lineWidth?: number
  showGradient?: boolean
  showRadius?: boolean
  officeSource?: OfficeSparklineSource
  sourceReference?: SparklineSourceReference
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
  /** Native OOXML identity stays separate from the domain column ID. */
  officeColumnId?: number
  /** null explicitly removes the corresponding native column property. */
  formula?: string | null
  formulaIsArray?: boolean
  totalsRowLabel?: string | null
  totalsRowFunction?: string | null
  totalsRowFormula?: string | null
  showFilterButton?: boolean
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
  showFirstColumn?: boolean
  showLastColumn?: boolean
  showRowStripes?: boolean
  showColumnStripes?: boolean
  /** Absent retains unprojected native criteria; null explicitly clears them. */
  autoFilter?: SheetAutoFilter | null
  officeTableId?: number
  officeSheetId?: number
  /** Complete native table XML, including fields not projected by the editor. */
  officeSource?: { partName: string; xml: string }
}

export type SheetSortDirection = 'ascending' | 'descending'
export interface SheetSortCriterion { column: number; direction: SheetSortDirection }
export type SheetFilterConditionOperator = 'equals' | 'doesNotEqual' | 'greaterThan' | 'greaterThanOrEqual' |
  'lessThan' | 'lessThanOrEqual' | 'between' | 'notBetween' | 'contains' | 'doesNotContain' |
  'beginsWith' | 'endsWith' | 'isBlank' | 'isNotBlank' | 'isError'
export interface SheetValueFilter {
  kind: 'values'
  selectedItems: CellScalar[]
  selectionMode?: 'include' | 'exclude'
  /** Native typed member rule; selectedItems is its legacy/raw mirror. */
  recordSelection?: PivotRecordSelection
  /** Calendar context for date storage aliases; numeric serial identity is unchanged. */
  recordDate1904?: boolean
}
export interface SheetConditionFilter {
  kind: 'condition'
  operator: SheetFilterConditionOperator
  value?: CellScalar
  secondValue?: CellScalar
  caseSensitive?: boolean
}
/** A native custom-filter column joins one or two non-nested conditions. */
export interface SheetCompoundFilter {
  kind: 'compound'
  join: 'and' | 'or'
  conditions: [SheetConditionFilter] | [SheetConditionFilter, SheetConditionFilter]
}
export interface SheetColorFilter {
  kind: 'color'
  /** null matches the default/no authored fill. Empty arrays match no rows. */
  cellFillColors?: Array<string | null>
  /** The default text color is black. As upstream, fill takes precedence when both arrays exist. */
  cellTextColors?: string[]
}
export type SheetFilterCriteria = SheetValueFilter | SheetConditionFilter | SheetColorFilter | SheetCompoundFilter
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
/** String names from the fixed Show Values As contract; the legacy parent-grand mode is retained separately. */
export type PivotShowDataAsType = 'normal' | 'percentOfColumnTotal' | 'percentOfRowTotal' | 'percentOfGrandTotal'
  | 'percentOfParentColumnTotal' | 'percentOfParentRowTotal' | 'index' | 'percentOfParent' | 'percentOf'
  | 'differenceFrom' | 'percentDifferenceFrom' | 'runningTotal' | 'percentRunningTotal' | 'rankAscending' | 'rankDescending'
  | 'percentOfParentGrandTotal'
export type PivotShowDataAsBaseItemType = 'item' | 'previous' | 'next'
export interface PivotShowDataAsInfo {
  type: PivotShowDataAsType
  /** Report-local row/column instance ID after canonical source IDs have been resolved. */
  baseFieldId: string
  /** Grouped member key; empty for rules that do not select a concrete member. */
  baseItem: string
  baseItemType?: PivotShowDataAsBaseItemType
}
export interface PivotNumberGroup { start: number; end: number; interval: number }
export interface PivotManualGroup { name: string; members: CellScalar[] }
export type PivotFilterOperator = 'valueGreaterThan' | 'valueGreaterThanOrEqual' | 'valueLessThan' |
  'valueLessThanOrEqual' | 'valueEqual' | 'valueNotEqual' | 'valueBetween' | 'valueNotBetween'

export interface PivotTableSource {
  unitId: string
  subUnitId: string
  sheetName: string
  range: CellRange
  /** Imported source identity; report.pivotCacheId is the canonical retained reference. */
  pivotCacheId?: string
}

/** A source reference shared by reports. The source cells remain in the workbook. */
export interface PivotCacheModel {
  id: string
  source: PivotTableSource
  /** Field identity/column metadata survives disconnection from every report. */
  fields?: Array<{ id: string; sourceColumn: number; name?: string }>
  nextFieldId?: number
  revision: number
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

export type PivotDateFilterOperator = 'dateBetween' | 'dateEqual' | 'dateNewerThan' | 'dateNewerThanOrEqual'
  | 'dateNotBetween' | 'dateNotEqual' | 'dateOlderThan' | 'dateOlderThanOrEqual'
export interface PivotDateFilter {
  operator: PivotDateFilterOperator
  expected: number | [number, number]
  /** Omitted means exact serial/time comparison. */
  useWholeDay?: boolean
}
export type PivotLabelConditionOperator = 'captionBeginsWith' | 'captionBetween' | 'captionContains' | 'captionEndsWith'
  | 'captionEqual' | 'captionGreaterThan' | 'captionGreaterThanOrEqual' | 'captionLessThan' | 'captionLessThanOrEqual'
  | 'captionNotBeginsWith' | 'captionNotBetween' | 'captionNotContains' | 'captionNotEndsWith' | 'captionNotEqual'
export interface PivotLabelConditionFilter {
  operator: PivotLabelConditionOperator
  expected: string | [string, string]
}
/** Raw member identity preserves scalar types when one displayed item is only partly selected. */
export type PivotRecordSelection = { mode: 'all' } | { mode: 'include' | 'exclude'; values: PivotRecordValue[] }

export interface PivotFieldModel {
  id: string
  /** Source-cache identity, independent of this report's field instance ID. */
  sourceFieldId?: string
  sourceColumn: number
  name: string
  area: PivotFieldArea
  position: number
  displayName?: string
  aggregation?: PivotAggregation
  showDataAs?: PivotShowDataAsInfo
  format?: string
  sort?: PivotSortDirection
  /** Legacy/Slicer source-string selection; grouped checklist queries use the computed item state. */
  selectedItems?: string[]
  /** Fully selected grouped keys remain rules when new raw source members enter the same group. */
  groupedItems?: string[]
  partialItems?: string[]
  recordSelection?: PivotRecordSelection
  labelFilter?: PivotLabelConditionFilter
  dateFilter?: PivotDateFilter
  /** Canonical Timeline range is a live half-open source rule, independent of enumerated members. */
  timelineDateRange?: ITimelineRange
  /** Typed-date authority for a Timeline condition; dateFilter retains its public Pivot projection. */
  timelineDateCondition?: ITimelineDateCondition
  valueFilter?: PivotFieldFilter
  numberGroup?: PivotNumberGroup
  dateGroupType?: PivotDateGroupType
  manualGroups?: PivotManualGroup[]
}

export type PivotValuePosition = 'none' | 'row' | 'column'

export interface PivotTableOptions {
  repeatRowLabels: boolean
  showRowGrandTotal: boolean
  showColGrandTotal: boolean
  showRowSubtotals: boolean
  showColSubtotals: boolean
  valuePosition: PivotValuePosition
  /** Zero-based ΣValues gap in its dimension axis; omitted/-1 chooses the current tail. */
  valueIndex?: number
  /** Filter categories are exclusive per canonical source unless explicitly enabled. */
  allowMultipleFilters?: boolean
}

/** Report-local hierarchy state; member maps are keyed by grouped member keys. */
export type PivotCollapseState = boolean | Record<string, boolean>
export type PivotCollapseInfo = Record<string, PivotCollapseState>
export interface PivotAxisOutputEntry {
  /** Absolute row/column index within the output matrix, including its headers. */
  index: number
  kind: 'detail' | 'collapsed' | 'subtotal' | 'grand-total'
  fieldId?: string
  /** Zero-based field depth; -1 for a grand total or an empty dimension axis. */
  level: number
  labels: CellScalar[]
  /** Report-local IDs and grouped scalar labels defining the original aggregate scope. */
  fieldFilters: Record<string, CellScalar>
  /** Original source members behind each grouped label in the scope. */
  fieldMembers: Record<string, CellScalar[]>
  valueFieldId?: string
}

export interface PivotTableModel {
  id: string
  pivotCacheId?: string
  source: PivotTableSource
  anchor: PivotTableAnchor
  fields: PivotFieldModel[]
  options: PivotTableOptions
  dateGroupingMode?: PivotDateGroupingMode
  collapseInfo?: PivotCollapseInfo
  outputRange?: CellRange
  generatedCellKeys?: string[]
  /** Derived worksheet output hit/projection metadata; never an independent aggregation state. */
  axisOutput?: { rowAxis: PivotAxisOutputEntry[]; columnAxis: PivotAxisOutputEntry[] }
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
export type ChartAxisTimeUnit = 'days' | 'months' | 'years'

export interface ChartFontOptions { color?: string; fontSize?: number; bold?: boolean; italic?: boolean; family?: string }

export type ChartTitlePosition = 'top' | 'bottom' | 'left' | 'right' | 'hide'
export type ChartTextAlignment = 'left' | 'center' | 'right'
/** Normalized text anchor in the chart frame; clearing it restores automatic layout. */
export interface ChartTitleManualLayout { left: number; top: number }
export interface ChartTitleStyleOptions extends ChartFontOptions { visible?: boolean; position?: ChartTitlePosition; alignment?: ChartTextAlignment; manualLayout?: ChartTitleManualLayout; rotate?: number }
export interface ChartSubtitleStyleOptions extends ChartFontOptions { alignment?: ChartTextAlignment; manualLayout?: ChartTitleManualLayout; rotate?: number }
export interface ChartLineOptions { visible?: boolean; color?: string; width?: number; opacity?: number; dashType?: 'solid' | 'dashed' | 'dotted' }
export interface ChartLabelOptions extends ChartFontOptions {
  visible?: boolean
  position?: 'auto' | 'inside' | 'insideBase' | 'insideEnd' | 'outside' | 'top' | 'bottom' | 'left' | 'right' | 'center'
  /** Numeric masks follow the fixed v1.6.0 public LabelContentType flags. */
  contentType?: 'value' | 'category' | 'series' | 'percentage' | number
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
/** Label-local pattern follows the v1.6.0 Treemap label contract. */
export interface ChartLabelPattern { type: 'pattern'; preset: string; foregroundColor: string; backgroundColor: string }
export interface ChartTreemapLabelOptions extends ChartLabelOptions { backgroundColor?: ChartColor | ChartLabelPattern }
export interface ChartTreemapOptions { parentLabelLayout?: 'banner' | 'overlapping' | 'none'; label?: ChartTreemapLabelOptions }
export type ChartSunburstLabelPosition = 'inside' | 'outside'
export interface ChartSunburstLabelOptions extends ChartLabelOptions {
  position?: ChartSunburstLabelPosition
  backgroundColor?: ChartColor | ChartLabelPattern
}
export interface ChartSunburstCenterLabelOptions extends ChartFontOptions {
  visible?: boolean
  format?: string
  /** Earlier project snapshots stored center-label fonts in this object. */
  font?: ChartFontOptions
}
export interface ChartSunburstOptions { label?: ChartSunburstLabelOptions; centerLabel?: ChartSunburstCenterLabelOptions }
/** Source/Target identity uses the final display label; Value omission counts rows. */
export interface ChartChordMapping { sourceIndex: number; targetIndex: number; valueIndex?: number }
export interface ChartChordLayoutOptions {
  startAngle?: number
  clockwise?: boolean
  padAngle?: number
  minAngle?: number
  innerRadiusRatio?: number
  outerRadiusRatio?: number
}
export interface ChartChordNodeOptions {
  color?: string
  opacity?: number
  borderColor?: string
  borderWidth?: number
  borderType?: 'solid' | 'dashed' | 'dotted'
}
export type ChartChordLabelPosition = 'inside' | 'outside'
export interface ChartChordLabelOptions extends ChartFontOptions {
  visible?: boolean
  position?: ChartChordLabelPosition
  rotation?: number
  distance?: number
}
export type ChartChordRibbonColorMode = 'source' | 'target' | 'gradient'
export interface ChartChordRibbonOptions {
  color?: ChartChordRibbonColorMode
  opacity?: number
  width?: number
  type?: 'solid' | 'dashed' | 'dotted'
}
export type ChartChordFocusMode = 'adjacency' | 'self' | 'none'
export interface ChartChordEmphasisOptions { focus?: ChartChordFocusMode }
export interface ChartChordOptions extends ChartChordLayoutOptions {
  layout?: ChartChordLayoutOptions
  node?: ChartChordNodeOptions
  label?: ChartChordLabelOptions
  ribbon?: ChartChordRibbonOptions
  emphasis?: ChartChordEmphasisOptions
  /** Historical flat layout and ribbonOpacity are read by the runtime only. */
  ribbonOpacity?: number
}
export type ChartPointShape = 'dot' | 'dash' | 'circle' | 'square' | 'triangle' | 'diamond' | 'rect' | 'roundRect' | 'pin' | 'arrow' | 'none'
  | 'emptyCircle' | 'emptySquare' | 'emptyTriangle' | 'emptyDiamond' | 'emptyRect' | 'emptyRoundRect' | 'emptyPin' | 'emptyArrow'
export interface ChartPointOptions { shape?: ChartPointShape; size?: number; color?: string }
export interface ChartDataPointOptions {
  color?: ChartColor
  fillOpacity?: number
  fillPattern?: ChartSeriesFillPattern | null
  size?: number
  shape?: ChartPointShape
  border?: ChartLineOptions
}
/** Slice appearance is independent of legacy series dataPoints. The
 * composite aggregate reuses these fields in its own unindexed author scope. */
export interface ChartPieSliceStyleOptions extends Pick<ChartDataPointOptions, 'color' | 'fillOpacity' | 'fillPattern' | 'border'> {
  explosion?: number
}
export type ChartPieAggregateStyleOptions = ChartPieSliceStyleOptions
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
  point?: ChartPointOptions
  rightAxis?: boolean
  label?: ChartLabelOptions
  dataPoints?: Record<number, ChartDataPointOptions>
  waterfallStyles?: ChartWaterfallStyles
}
export type ChartWaterfallStyleTarget = 'positive' | 'negative' | 'subtotal'
export type ChartWaterfallStackType = 'arrangement' | 'stacked'
export type ChartWaterfallSeriesStyle = Omit<ChartSeriesOptions, 'waterfallStyles'>
export type ChartWaterfallStyles = Partial<Record<ChartWaterfallStyleTarget, ChartWaterfallSeriesStyle>>
export type ChartWaterfallConnectorOptions = Omit<ChartLineOptions, 'visible'>
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
export interface ChartParetoBarStyleOptions {
  color?: ChartColor
  opacity?: number
  border?: ChartLineOptions
  label?: ChartLabelOptions
  dataPoints?: Record<number, ChartDataPointOptions>
}
export interface ChartParetoLineStyleOptions {
  color?: ChartColor
  opacity?: number
  width?: number
  dashType?: ChartLineOptions['dashType']
  label?: ChartLabelOptions
  dataPoints?: Record<number, ChartDataPointOptions>
  point?: ChartPointOptions
}
export interface ChartParetoOptions {
  barStyle?: ChartParetoBarStyleOptions
  cumulativeLineStyle?: ChartParetoLineStyleOptions
  includeZeroValues?: boolean
}
export interface ChartCandlestickStateStyleOptions {
  color?: string
  fillOpacity?: number
  border?: Pick<ChartLineOptions, 'color' | 'width' | 'opacity'>
  hollow?: boolean
}
export interface ChartCandlestickHighLowLineOptions { visible?: boolean; border?: ChartLineOptions }
export interface ChartCandlestickCloseMarkerOptions { visible?: boolean; point?: ChartPointOptions }
export interface ChartCandlestickMappingOptions { openIndex?: number; highIndex: number; lowIndex: number; closeIndex: number }
export interface ChartCandlestickOptions {
  candleWidth?: number
  rising?: ChartCandlestickStateStyleOptions
  falling?: ChartCandlestickStateStyleOptions
  doji?: ChartCandlestickStateStyleOptions
  highLowLine?: ChartCandlestickHighLowLineOptions
  closeMarker?: ChartCandlestickCloseMarkerOptions
  /** Earlier project snapshots retain these values; state groups take precedence. */
  risingColor?: string
  fallingColor?: string
  hollow?: boolean
}
/** Renderer-independent authored configuration. Derived statistics and geometries are not saved here. */
export type ChartWordCloudShape = 'cloud' | 'rect' | 'circle' | 'cardioid' | 'heart' | 'diamond' | 'triangle-forward' | 'triangle' | 'pentagon' | 'star'
/** Persisted by the chart resource adapter; browser handles never enter this DTO. */
export interface ChartImageResource {
  id: string
  sourceType: 'dataUrl' | 'url'
  source: string
  mimeType: string
  sizeBytes?: number
  name?: string
}
export interface ChartWordCloudOptions {
  shape?: ChartWordCloudShape
  maskImageId?: string
  /** Legacy authored alias; canonical writes and resets clear it. */
  maskImage?: string
  /** Detached upload payload. Publication moves it to the workbook registry. */
  maskImageResource?: ChartImageResource
  repeat?: boolean
  /** Local runtime authoring extensions; fixed published defaults are 16..56. */
  minFontSize?: number
  maxFontSize?: number
  rotation?: number
}

export type ChartRelationLayout = 'none' | 'circular' | 'force'
export interface ChartRelationForceOptions { repulsion?: number; gravity?: number; edgeLength?: number }
export interface ChartRelationOptions {
  layout?: ChartRelationLayout
  useValueAsSymbolSize?: boolean
  emphasisEnabled?: boolean
  circularLabelRotation?: boolean
  nodeShape?: ChartPointShape
  force?: ChartRelationForceOptions
}

export type ChartAreaLineStyle = 'line' | 'smooth' | 'step'
export type ChartInvalidValueStrategy = 'zero' | 'break' | 'link'

export interface ChartLineAndAreaOptions {
  /** Shared shape of line/area series, including stacked and projected 3D variants. */
  lineStyle?: ChartAreaLineStyle
  /** Legacy area-only switch. An explicit lineStyle takes precedence. */
  smooth?: boolean
}

export interface ChartSemanticOptions {
  invalidValueStrategy?: ChartInvalidValueStrategy
  plotVisibleOnly?: boolean
  autoGradientFill?: boolean
  plotArea?: ChartPlotAreaOptions
  appearance?: { background?: string; border?: ChartLineOptions; font?: ChartFontOptions }
  titleStyle?: ChartTitleStyleOptions
  subtitleStyle?: ChartSubtitleStyleOptions
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
    /** Local row-major adapter for the published Relation co-word data contract. */
    relationFormat?: 'edgeList' | 'coWordMatrix'
    chord?: ChartChordMapping
    candlestick?: ChartCandlestickMappingOptions
    gauge?: { valueIndex: number; labelIndex?: number }
  }
  threeD?: { view?: { rotX?: number; rotY?: number; perspective?: number; rightAngleAxes?: boolean; heightPercent?: number; depthPercent?: number }; grouping?: 'standard' | 'clustered'; gapDepth?: number }
  bar?: { gapWidth?: number; overlap?: number }
  area?: ChartLineAndAreaOptions
  pie?: {
    startAngle?: number; doughnutHole?: number; explosion?: number; rosePie?: boolean; isHalfPie?: boolean
    showLabelLine?: boolean; label?: ChartLabelOptions
    secondaryPlot?: { type?: 'pie' | 'bar'; size?: number; gapWidth?: number; connector?: ChartLineOptions; split?: { type: 'auto' | 'position' | 'value' | 'percentage' | 'custom'; count?: number; lessThan?: number; pointIndexes?: number[] } }
    aggregateLabel?: string
    aggregateStyle?: ChartPieAggregateStyleOptions
    sliceStyles?: Partial<Record<number, ChartPieSliceStyleOptions>>
  }
  radar?: ChartRadarOptions
  bubble?: { scale?: number; sizeRepresents?: 'area' | 'width' }
  funnel?: { gap?: number; useAbsoluteValue?: boolean; sort?: 'ascending' | 'descending' | 'none' }
  relation?: ChartRelationOptions
  waterfall?: { pointRoles?: Array<{ seriesIndex: number; dataPointIndex: number; role: 'change' | 'total' }>; stackType?: ChartWaterfallStackType; useSubtotal?: boolean; connector?: boolean | ChartWaterfallConnectorOptions; styles?: ChartWaterfallStyles; positiveColor?: string; negativeColor?: string; totalColor?: string }
  pareto?: ChartParetoOptions
  histogram?: { binning?: { mode: 'count'; count: number } | { mode: 'width'; width: number }; underflowThreshold?: number; overflowThreshold?: number; binGap?: number }
  boxplot?: { quartileMethod?: 'medianOfHalves' | 'inclusive' | 'exclusive'; whiskerMethod?: 'range' | 'iqr'; showOutliers?: boolean; whiskerMultiplier?: number }
  candlestick?: ChartCandlestickOptions
  heatmap?: { visualMapType?: 'continuous' | 'piecewise'; min?: number; max?: number; unit?: string; label?: ChartLabelOptions }
  treemap?: ChartTreemapOptions
  sunburst?: ChartSunburstOptions
  gauge?: ChartGaugeOptions
  chord?: ChartChordOptions
  wordCloud?: ChartWordCloudOptions
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
  /** Imported date-axis granularity. Bounds and data retain Excel serials. */
  baseTimeUnit?: ChartAxisTimeUnit
  /** Tick calendar unit; absent inherits baseTimeUnit, then days. */
  majorTimeUnit?: ChartAxisTimeUnit
  /** Positive major-tick interval. Calendar month/year intervals are integers. */
  majorUnit?: number
  reverse?: boolean
  line?: ChartLineOptions
  label?: ChartFontOptions & { visible?: boolean; rotate?: number }
  gridLine?: ChartLineOptions
  tick?: ChartLineOptions & { length?: number; position?: 'inside' | 'outside' }
  format?: string
}

export type ChartLegendSelectMode = 'single' | 'multiple' | 'close'

export interface ChartLegendOptions {
  visible: boolean
  position: ChartLegendPosition
  font?: ChartFontOptions
  /** Authored capability; the selected names belong to the surface view. */
  selectMode?: ChartLegendSelectMode
  /** Horizontal legends wrap; side legends retain pagination. */
  wrap?: boolean
}

export type ChartDepthAxisOptions = Omit<ChartAxisOptions, 'scaleType' | 'logBase' | 'min' | 'max' | 'format' | 'baseTimeUnit' | 'majorTimeUnit' | 'majorUnit'>

export interface PivotChartControls {
  showFieldButtons: boolean
  showExpandCollapseButtons: boolean
}

export interface ChartModel {
  id: string
  sheetId: string
  type: ChartType
  source: ChartSource
  position: { x: number; y: number }
  size: { width: number; height: number }
  drawingGeometry?: DrawingAffineGeometry
  title?: string
  subtitle?: string
  palette?: string[]
  options?: ChartSemanticOptions
  legend: ChartLegendOptions
  categoryField: number
  valueFields: number[]
  seriesToCategory?: Record<number, number[]>
  pivotTableId?: string
  /** Independently owned report. It uses the Pivot engine without writing a worksheet output. */
  ownedPivot?: PivotTableModel
  /** Pivot presentation controls; both default to true when absent in an older snapshot. */
  controls?: PivotChartControls
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

export interface SlicerPivotTableReference {
  subUnitId: string
  pivotTableId: string
}

export interface SlicerOwnedPivotChartReference {
  subUnitId: string
  pivotChartId: string
}

export interface SheetSlicerConnections {
  pivotTables: SlicerPivotTableReference[]
  ownedPivotCharts?: SlicerOwnedPivotChartReference[]
}

export interface SlicerConnectionCandidate {
  kind: 'pivotTable' | 'ownedPivotChart'
  subUnitId: string
  id: string
  name: string
  sheetName: string
}

export type SlicerCacheSource = {
  kind: 'pivot'
  pivotCacheId: string
  sourceFieldId: string
} | {
  kind: 'table'
  subUnitId: string
  tableId: string
  columnId: string
}

export type SlicerCacheSortOrder = 'ascending' | 'descending'
export type SlicerCacheCrossFilter = 'none' | 'showItemsWithDataAtTop' | 'showItemsWithNoData'
export interface SheetSlicerCacheOptions {
  sortOrder?: SlicerCacheSortOrder
  customListSort?: boolean
  crossFilter?: SlicerCacheCrossFilter
  hideItemsWithNoData?: boolean
}

/** Pivot filters/connections are authoritative here; Table values mirror native column filters; drawings own presentation only. */
export interface SlicerCacheModel {
  id: string
  unitId: string
  type: SlicerType
  /** Cache and original-source names are independent of each hosted view caption/name. */
  name?: string
  sourceName?: string
  source: SlicerCacheSource
  connections?: SheetSlicerConnections
  /** Authored options are shared; own undefined in an options patch clears the optional value. */
  options?: SheetSlicerCacheOptions
  /** Typed raw member identity is authoritative; selectedItems is a display/legacy mirror. */
  recordSelection?: PivotRecordSelection
  selectedItems?: string[]
  timelineRange?: SlicerTimelineRange
  timelineDateRange?: ITimelineRange
  timelineDateCondition?: ITimelineDateCondition
  /** Decoded legacy member filters are retired together with the canonical date rule when reports are available. */
  timelineLegacyMemberFilterPending?: boolean
  revision: number
}

/** A persisted filter drawing bound to a shared cache, with legacy source hints retained. */
export interface SlicerModel {
  id: string
  unitId: string
  hostSheetId: string
  cacheId?: string
  pivotTableId?: string
  ownedPivotChartId?: string
  tableId?: string
  sourceFieldId: string
  type: SlicerType
  caption?: string
  name?: string
  position: { row: number; column: number }
  size: { rows: number; columns: number }
  layout?: { left: number; top: number; width: number; height: number }
  /** Local control layout mapped to header-free Sheet CSS pixels; filters remain in this feature model. */
  drawingGeometry?: DrawingAffineGeometry
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
  timelineDateRange?: ITimelineRange
  timelineDateCondition?: ITimelineDateCondition
  revision: number
}

export const SHEET_RANGE_THEME_SECTIONS = ['wholeStyle', 'firstColumnStyle', 'secondColumnStyle', 'firstRowStyle', 'secondRowStyle',
  'headerColumnStyle', 'lastColumnStyle', 'headerRowStyle', 'lastRowStyle'] as const
export type SheetRangeThemeSection = typeof SHEET_RANGE_THEME_SECTIONS[number]
/** Range themes are author rules, composed before cell direct formatting instead of written into every cell. */
export type SheetRangeThemeStyle = { name: string; metadata?: Record<string, unknown> } & Partial<Record<SheetRangeThemeSection, SheetCellStylePatch | null>>
export interface SheetRangeThemeRule {
  id: string
  themeName: string
  range: CellRange
  /** Logical header/stripe/footer geometry retained when a partial move splits physical membership. */
  styleRange?: CellRange
  metadata?: Record<string, unknown>
}

export interface SpreadsheetSheet {
  id: string
  name: string
  visible: boolean
  tabColor?: string
  cells: Record<string, SpreadsheetCell>
  images?: SheetFloatingImage[]
  drawingState?: SheetDrawingState
  shapes?: SheetDrawingShapeDTO[]
  backgroundImage?: SheetBackgroundImage
  defaultStyle?: SheetStyleSource
  rowStyles?: Record<number, SheetStyleSource>
  columnStyles?: Record<number, SheetStyleSource>
  rowStyleSpans?: SheetAxisSpan<SheetStyleSource>[]
  columnStyleSpans?: SheetAxisSpan<SheetStyleSource>[]
  rowStylePriority?: boolean
  rangeThemeRules?: SheetRangeThemeRule[]
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
  defaultStyle?: SheetStyleSource
  styles?: Record<string, SheetCellStylePatch>
  rangeThemes?: Record<string, SheetRangeThemeStyle>
  calculationMode?: 'automatic' | 'manual'
  sheets: SpreadsheetSheet[]
  activeSheetId: string
  date1904: boolean
  modified: boolean
  id?: string
  metadata?: Record<string, unknown>
  pivotTables?: PivotTableModel[]
  pivotCaches?: PivotCacheModel[]
  tables?: SheetTableModel[]
  charts?: ChartModel[]
  slicers?: SlicerModel[]
  slicerCaches?: SlicerCacheModel[]
  slicerStyles?: ISlicerStylesDefinition
  timelineStyles?: ITimelineStylesDefinition
}

export type DocumentSectionBreakType = 'unspecified' | 'nextPage' | 'continuous' | 'nextColumn' | 'evenPage' | 'oddPage'
export type DocumentBlockType = 'paragraph' | 'heading' | 'bullet' | 'numbered' | 'checklist' | 'quote' | 'callout' | 'code' | 'columns' | 'divider' | 'table' | 'image' | 'pageBreak' | 'columnBreak' | 'sectionBreak'

/** Block decoration is independent of the editable paragraphs it contains. */
export interface DocumentCalloutConfig {
  icon: string
  showIcon: boolean
  backgroundColor: string
  borderColor: string
  borderWidth: number
  borderOpacity: number
  borderStyle: 'solid' | 'dashed' | 'dotted'
  borderRadius: number
  paddingTop: number
  paddingRight: number
  paddingBottom: number
  paddingLeft: number
  iconSize: number
  iconGap: number
}

export interface DocumentCodeConfig {
  language: string
  showLineNumbers: boolean
  tabSize: number
  wrap: boolean
}

export interface DocumentColumn {
  id: string
  widthRatio: number
  /** CSS pixels; public number-unit provenance is retained separately. */
  minWidth?: number
  blocks: DocumentBlock[]
  metadata?: Record<string, unknown>
}

export interface DocumentColumnGroup {
  version: 1
  columns: DocumentColumn[]
  gap: number
  /** CSS pixels; omitted values use the fixed Univer defaults. */
  horizontalPadding?: number
  minHeight?: number
  layout: 'fixed' | 'auto'
  responsive: 'stack' | 'shrink'
}

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
  /** Undefined inherits; false explicitly cancels an inherited pagination rule. */
  keepLines?: boolean
  keepNext?: boolean
  widowControl?: boolean
  pageBreakBefore?: boolean
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
  noteProperties?: { footnote?: DocumentFootnoteProperties; endnote?: DocumentEndnoteProperties }
  suppressEndnotes?: boolean
}

export interface DocumentBlock {
  id: string
  type: DocumentBlockType
  level: number
  runs: InlineRun[]
  style: DocumentParagraphStyle
  metadata: Record<string, unknown>
  /** Authoritative children of quote/callout/code blocks. A legacy quote can remain a paragraph. */
  children?: DocumentBlock[]
  callout?: DocumentCalloutConfig
  code?: DocumentCodeConfig
  columns?: DocumentColumnGroup
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

/** A real OPC custom-XML data source, separate from office-studio's own sidecars. */
export interface DocumentCustomXmlPart {
  partName: string
  xml: string
  contentType?: string
  storeItemID?: string
  propertiesPartName?: string
  propertiesXml?: string
  relationshipsXml?: string
  /** Exact related OPC resources, including their own relationship parts. */
  relatedParts?: Array<{ partName: string; contentType: string; dataBase64: string }>
  /** XML stays Unicode in the model; IO retains the source byte encoding. */
  encoding?: 'utf8' | 'utf16le' | 'utf16be'
  byteOrderMark?: boolean
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
  /** Cumulative note-reference boundaries within the owning paragraph. */
  startNoteReferenceCount?: number
  endNoteReferenceCount?: number
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
  /** Row SDTs own an ordered contiguous range in one table. IDs survive edits and snapshot reopen. */
  tableRowRange?: { tableId: string; rowIds: string[] }
  /** Block SDTs own complete, ordered contiguous blocks in one body/story/cell container. */
  blockRange?: { blockIds: string[] }
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
  /** Custom marks do not consume an automatic footnote/endnote number. */
  customMark?: string
  referenceTextStyle?: DocumentDefaultStyle['text']
  showBodyReference?: boolean
  metadata?: Record<string, unknown>
}

export interface DocumentFootnoteProperties {
  position?: 'pageBottom' | 'beneathText'
  restart?: 'continuous' | 'eachSect' | 'eachPage'
  startNumber?: number
  /** OOXML/public format identity; unknown imported formats remain preserved. */
  numberFormat?: string
  columnCount?: number
  metadata?: Record<string, unknown>
}

export interface DocumentEndnoteProperties {
  position?: 'docEnd' | 'sectEnd'
  restart?: 'continuous' | 'eachSect'
  startNumber?: number
  numberFormat?: string
  metadata?: Record<string, unknown>
}

export interface DocumentNoteDecorations {
  separator?: DocumentStoryContent
  continuationSeparator?: DocumentStoryContent
  continuationNotice?: DocumentStoryContent
}

export interface DocumentFootnoteSettings extends DocumentFootnoteProperties, DocumentNoteDecorations {}
export interface DocumentEndnoteSettings extends DocumentEndnoteProperties, DocumentNoteDecorations {}

export interface DocumentCommentAnchorSegment {
  blockId: string
  startOffset: number
  endOffset: number
  /** Cumulative paragraph note references before the endpoint. Distinguishes
   * positions around references with an empty display label at the same offset. */
  startNoteReferenceCount?: number
  endNoteReferenceCount?: number
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

export type DocumentPermissionAction = 'view' | 'edit' | 'copy' | 'comment' | 'print' | 'export'
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
  customXmlParts?: DocumentCustomXmlPart[]
  sectionHeaderFooters?: Record<string, DocumentSectionHeaderFooter>
  permissionRules?: DocumentPermissionRule[]
  footnotes?: Record<string, DocumentNote>
  endnotes?: Record<string, DocumentNote>
  noteSettings?: { footnote?: DocumentFootnoteSettings; endnote?: DocumentEndnoteSettings }
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
  authoredStyle: {},
})

export const createBlankSheet = (name = 'Sheet1'): SpreadsheetSheet => ({
  id: `sheet-${crypto.randomUUID?.() ?? Date.now()}`,
  name,
  visible: true,
  cells: {},
  rowCount: DEFAULT_SHEET_ROW_COUNT,
  columnCount: DEFAULT_SHEET_COLUMN_COUNT,
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
