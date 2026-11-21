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
  metadata?: Record<string, unknown>
}

export interface CellRange {
  startRow: number
  startCol: number
  endRow: number
  endCol: number
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
export type SheetRangeProtectionDraft = Omit<SheetRangeProtectionRule, 'id' | 'permissionId' | 'ownerUserId'> & {
  id?: string
  permissionId?: string
  ownerUserId?: string
}

export type PivotFieldArea = 'filter' | 'row' | 'column' | 'value'
export type PivotAggregation = 'sum' | 'count' | 'average' | 'min' | 'max'
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

export type ChartType = 'bar' | 'line'
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

/** A persisted filter control bound to one pivot source field. */
export interface SlicerModel {
  id: string
  unitId: string
  hostSheetId: string
  pivotTableId: string
  sourceFieldId: string
  type: SlicerType
  caption: string
  position: { row: number; column: number }
  size: { rows: number; columns: number }
  selectedItems?: string[]
  timelineRange?: SlicerTimelineRange
  revision: number
}

export interface SpreadsheetSheet {
  id: string
  name: string
  visible: boolean
  cells: Record<string, SpreadsheetCell>
  mergedRanges: CellRange[]
  columnWidths: Record<number, number>
  rowHeights: Record<number, number>
  frozenRows: number
  frozenCols: number
  showGridlines: boolean
  hiddenRows?: number[]
  autoFilter?: SheetAutoFilter
  dataValidations?: SheetDataValidationRule[]
  rangeProtectionRules?: SheetRangeProtectionRule[]
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
  sheets: SpreadsheetSheet[]
  activeSheetId: string
  date1904: boolean
  modified: boolean
  id?: string
  metadata?: Record<string, unknown>
  pivotTables?: PivotTableModel[]
  charts?: ChartModel[]
  slicers?: SlicerModel[]
}

export type DocumentSectionBreakType = 'nextPage' | 'continuous' | 'nextColumn' | 'evenPage' | 'oddPage'
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
}

export interface DocumentBlock {
  id: string
  type: DocumentBlockType
  level: number
  runs: InlineRun[]
  style: DocumentParagraphStyle
  metadata: Record<string, unknown>
  checklist?: { listId: string; checked: boolean }
  sectionBreak?: { type: DocumentSectionBreakType }
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

export interface DocumentModel {
  kind: 'document'
  name: string
  blocks: DocumentBlock[]
  page: DocumentPageSetup
  modified: boolean
  wordCount: number
  contentControls?: DocumentContentControl[]
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
