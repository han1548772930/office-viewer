export type OfficeKind = 'workbook' | 'document'

export type CellScalar = string | number | boolean | null

export interface FontStyle {
  family: string
  size: number
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  color: string
}

export interface CellStyle {
  font: FontStyle
  fill: string
  horizontal: 'left' | 'center' | 'right' | 'justify'
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

export type SheetDataValidationType = 'whole' | 'decimal' | 'textLength' | 'date' | 'time' | 'checkbox' | 'list' | 'listMultiple' | 'custom'
export type SheetDataValidationOperator = 'between' | 'equal' | 'greaterThan' | 'greaterThanOrEqual' |
  'lessThan' | 'lessThanOrEqual' | 'notBetween' | 'notEqual'
export type SheetDataValidationErrorStyle = 'stop' | 'warning' | 'info'
export type SheetDataValidationRenderMode = 'text' | 'arrow' | 'custom'

export interface SheetDataValidationRule {
  id: string
  type: SheetDataValidationType
  ranges: CellRange[]
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
export type SheetConditionalFormatType = 'cellIs' | 'highlightCell' | 'top10' | 'average' | 'duplicateValues' | 'uniqueValues' | 'timePeriod' | 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith' | 'expression' | 'dataBar' | 'colorScale' | 'iconSet'
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
}
export interface SheetConditionalFormatRule {
  id: string
  type: SheetConditionalFormatType
  subType?: SheetConditionalFormatSubType
  ranges: CellRange[]
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

export type ChartType = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'stacked-column' | 'scatter'
export type ChartSourceOrientation = 'columns' | 'rows'
export type ChartLegendPosition = 'top' | 'right' | 'bottom' | 'left'
export type ChartAxisScaleType = 'linear' | 'log'

export interface ChartSource {
  range: CellRange
  orientation: ChartSourceOrientation
}

export interface ChartAxisOptions {
  title?: string
  scaleType: ChartAxisScaleType
  logBase?: number
}

export interface ChartLegendOptions {
  visible: boolean
  position: ChartLegendPosition
}

export interface ChartModel {
  id: string
  sheetId: string
  type: ChartType
  source: ChartSource
  position: { x: number; y: number }
  size: { width: number; height: number }
  title?: string
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
  frozenRows: number
  frozenCols: number
  showGridlines: boolean
  hiddenRows?: number[]
  outline?: SheetOutlineState
  autoFilter?: SheetAutoFilter
  dataValidations?: SheetDataValidationRule[]
  conditionalFormats?: SheetConditionalFormatRule[]
  rangeProtectionRules?: SheetRangeProtectionRule[]
  worksheetProtection?: SheetWorksheetProtectionRule
  hiddenColumns?: number[]
  defaultColumnWidth?: number
  defaultRowHeight?: number
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
  fontSize: number
  verticalAlign?: 'baseline' | 'subscript' | 'superscript'
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

export interface DocumentSectionSetup {
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
  differentFirstPage: boolean
  differentOddEven: boolean
  linkToPrevious: boolean
}

export interface DocumentPageField {
  story: 'header' | 'footer' | 'firstPageHeader' | 'firstPageFooter' | 'evenPageHeader' | 'evenPageFooter'
  field: 'PAGE' | 'NUMPAGES'
}

export interface DocumentNote {
  id: string
  text: string
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
    color: '#1f2937',
  },
  fill: '#ffffff',
  horizontal: 'left',
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

export const createBlankDocument = (): DocumentModel => ({
  kind: 'document',
  name: '未命名文档',
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
