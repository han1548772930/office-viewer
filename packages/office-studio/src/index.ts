import './style.css'

export { OfficeStudio, type OfficeStudioOptions, type OfficeCommandHandler } from './api/office-studio'
export {
  FBlob,
  FDocument,
  FDocumentContentControl,
  FDocumentParagraph,
  FDocumentTextRange,
  FSheetSlicer,
  FPivotTable,
  FRange,
  FSelection,
  FUniver,
  FWorkbook,
  FWorksheet,
  PositionTypeEnum,
  type CreateUnitOptions,
  type DocumentBodySegment,
  type DocumentBodySnapshot,
  type DocumentContentControlCreateInfo,
  type DocumentContentControlInfo,
  type DocumentContentControlPatch,
  type DocumentParagraphIdentity,
  type DocumentParagraphInfo,
  type DocumentParagraphQuery,
  type DocumentTextRange,
  type DocumentTextStyle,
  type DocumentTextStylePatch,
  type DocumentTextStyleRun,
  type FRangeCoordinates,
  type PivotTableSourceInfo,
  type PivotTableAnchorInfo,
  type SlicerInfo,
} from './api/facade'
export { FEventName, EventHub, type ApiEventMap, type ApiRange, type Disposable } from './api/events'
export { type ApiCellData, type ApiCellValue } from './api/model-utils'

export { default as DocumentWorkbench } from './components/DocumentWorkbench.vue'
export { default as ExcelWorkbench } from './components/ExcelWorkbench.vue'
export type { ExcelCell, ExcelSheet, ExcelWorkbook } from './components/ExcelWorkbench.vue'

export * from './domain/model'
export { cloneModel } from './domain/serialization'
export {
  getPivotSourceFields,
  getPivotTableAtCell,
  pivotDrillDown,
  pivotFieldIds,
  pivotOutputValues,
  createPivotTable,
  refreshPivotTable,
  setPivotField,
  setPivotFieldSetting,
  setPivotLabelFilter,
  setPivotOptions,
  setPivotSort,
  setPivotValueFilter,
  setPivotNumberGroup,
  setPivotDateGroupingMode,
  setPivotDateGroupType,
  setPivotExcelDateGroup,
  setPivotManualGroups,
} from './features/pivot/services/pivot-engine'
export {
  clearSlicerFilter,
  createSlicer,
  getSlicer,
  getSlicerItems,
  removeSlicer,
  setSlicerItems,
  setSlicerTimelineRange,
} from './features/slicer/slicer-engine'
export {
  clearFilter,
  expandVisibleSortOrder,
  getFilteredRows,
  getNumericSortColumns,
  getSortableRowOffsets,
  getSortedRowOrder,
  normalizeAutoFilterRange,
  refreshAutoFilter,
  setAutoFilter,
  setFilterCriteria,
  sortRange,
} from './features/sheet-sort-filter/sort-filter-engine'
export type {
  FormulaReferenceTranslator,
  NumericSortColumnData,
  SortRangeOptions,
} from './features/sheet-sort-filter/sort-filter-engine'
export { useEditorStore } from './stores/editorStore'
export { useWorkbookStore } from './stores/workbookStore'
export { useDocumentStore } from './stores/documentStore'
export { useHistoryStore, type HistoryEntry, type UnitHistory } from './stores/historyStore'
export { useParserStore } from './stores/parserStore'

export {
  applyExcelWorkbookView,
  cellKey,
  parseCellKey,
  toExcelWorkbookView,
  type ExcelCellView,
  type ExcelSheetView,
  type ExcelWorkbookView,
} from './services/excelViewAdapter'
export {
  downloadOfficeBytes,
  exportOfficeModel,
  officeDownloadName,
  OfficeExportError,
  type ExportOptions,
} from './services/exportService'
export { ParserError } from './parser/errors'
export { type ExportOutcome, type ExportProgress } from './parser/export-contracts'

