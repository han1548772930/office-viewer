import './style.css'
export { OfficeFontListKey, type OfficeFontList } from './platform/ribbon/ribbon-fonts'

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
  FOffice,
  FWorkbook,
  FWorksheet,
  PositionTypeEnum,
  type CreateUnitOptions,
  type DataValidationError,
  type DataValidationRule,
  type DataValidationRuleInput,
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
export { default as DocumentTableOfContentsDialog } from './components/document/DocumentTableOfContentsDialog.vue'
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
export { useEditorStore, type EditorSavepoint } from './stores/editorStore'
export { useWorkbookStore } from './stores/workbookStore'
export { useDocumentStore } from './stores/documentStore'
export { useHistoryStore, type HistoryEntry, type HistorySavepoint, type UnitHistory } from './stores/historyStore'
export { useParserStore } from './stores/parserStore'
export { default as WorkbenchRibbon } from './platform/ribbon/WorkbenchRibbon.vue'
export { createOfficeI18n, OfficeDirectionKey, OfficeLocaleKey, OfficeRegionKey, OfficeTranslatorKey, supportedOfficeLocales, createOfficeMessages, createOfficeTranslator, createRibbonTranslator, localizeRibbon } from './i18n/office-i18n.ts'
export type { OfficeLocale, OfficeTranslator } from './i18n/office-i18n.ts'
export type { RibbonTab, RibbonGroup, RibbonItem, RibbonVariant } from './platform/ribbon/ribbon-schema'
export { RIBBON_VARIANTS, isRibbonVariant } from './platform/ribbon/ribbon-schema'
export { WORKBENCH_CHROME, WORKBENCH_ACCENTS, WORKBENCH_ZOOMS, WorkbenchDisplayKey } from './platform/workbench-display'
export type { WorkbenchChrome, WorkbenchAccent, WorkbenchZoom } from './platform/workbench-display'

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
export { createSparklineGroupPatches, createSparklineSpec, translateSparklineForMove, type SparklineGroupRequest } from './features/sparkline/sparkline-model.ts'
export { createSparklineCommands, parseSparklineRange, readSparklineValues, sparklineCommandsForCell } from './features/sparkline/sparkline-render.ts'
export { default as SparklinePanel } from './components/spreadsheet/SparklinePanel.vue'

