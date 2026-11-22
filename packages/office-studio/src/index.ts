import './style.css'
export { OfficeFontListKey, type OfficeFontList } from './platform/ribbon/ribbon-fonts'

export { OfficeStudio, type OfficeStudioOptions, type OfficeCommandHandler, type WorkbookCommitOptions } from './api/office-studio'
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
  type FSelectionCell,
  type PivotTableSourceInfo,
  type PivotTableAnchorInfo,
  type SlicerInfo,
} from './api/facade'
export { FEventName, EventHub, type ApiEventMap, type ApiRange, type ApiSheetSelection, type Disposable } from './api/events'
export { normalizeSheetSelection, reconcileSheetSelection, resolveVisibleSheetSelection, type SheetSelectionState, type SheetSelectionRange,
  type SheetSelectionCell, type SheetSelectionOptions, type SheetSelectionSheet } from './features/sheet-workbench/sheet-selection.ts'
export { SHEET_CLEAR_COMMANDS, SHEET_CLEAR_COMMAND_ALIASES, type SheetClearOptions, type SheetClearMode,
  type SheetClearCommandInput } from './features/sheet-workbench/sheet-clear-command.ts'
export { SHEET_MERGE_COMMANDS, SHEET_MERGE_COMMAND_ALIASES, type SheetMergeOptions, type SheetMergeMode,
  type SheetMergeCommandInput, type SheetMergeResult, type SheetMergeConfirmationRequest } from './features/sheet-workbench/sheet-merge-command.ts'
export { SHEET_FORMAT_PAINTER_COMMANDS, SHEET_FORMAT_PAINTER_ALIASES, SHEET_FORMAT_PAINTER_COMMAND_IDS,
  type SheetFormatPainterMode, type SheetFormatPainterState, type SheetFormatPainterCommandInput,
  type SheetFormatPainterConfirmationRequest } from './features/sheet-workbench/sheet-format-painter-session.ts'
export type { SheetFormatPainterSource, SheetFormatPainterCanvasState, SheetFormatPainterOptions,
  SheetFormatPainterPlan, SheetFormatPainterResult } from './features/sheet-workbench/sheet-format-painter.ts'
export type { SheetConditionalFormatPainterSource } from './features/sheet-workbench/sheet-format-painter-conditional.ts'
export { SheetBorderTypes, SheetBorderStyleTypes, type SheetBorderType, type SheetBorderStyleType,
  type SheetBorderPreset, type SheetBorderStyle, type SheetBorderLineStyle, type SheetBorderOptions } from './features/sheet-workbench/sheet-cell-border.ts'
export type { SheetBorderFormatState, SheetBorderFormatCommandInput } from './features/sheet-workbench/sheet-border-format-command.ts'
export { SHEET_AXIS_DIMENSION_COMMANDS, SHEET_AXIS_DIMENSION_COMMAND_IDS, type SheetAxisDimensionCommandInput, type SheetAxisDimensionAction,
  type SheetAxisDimensionOptions } from './features/sheet-workbench/sheet-axis-dimensions.ts'
export type { SheetAxisSpan, SheetAxisDimensions } from './domain/sheet-axis.ts'
export { sheetRowHeight, sheetColumnWidth, sheetRowIsAutoHeight, SparseAxisIndex } from './domain/sheet-axis.ts'
export { transformSheetAxisDimensions, transformSheetAxisSpans, type SheetAxisStructureEdit, type SheetAxisDimensionChanges } from './domain/sheet-axis-transform.ts'
export { SHEET_AXIS_STRUCTURE_COMMAND_IDS, type SheetAxisStructureCommandInput, type SheetAxisStructureAction,
  type SheetAxisStructurePlan } from './features/sheet-workbench/sheet-axis-structure-command.ts'
export { SHEET_AXIS_VISIBILITY_COMMANDS, SHEET_AXIS_VISIBILITY_COMMAND_IDS, SheetAxisVisibilityError, planSheetAxisVisibilityCommand,
  type SheetAxisVisibilityCommandInput, type SheetAxisVisibilityAction, type SheetAxisVisibilityPlan } from './features/sheet-workbench/sheet-axis-visibility.ts'
export { type SheetAxis, type SheetAxisEdit, type SheetAxisInterval, sheetAxisMaximum, mapSheetAxisIndex,
  transformSheetAxisInterval, transformSheetAxisRange, transformSheetAxisReferenceInterval } from './domain/sheet-axis-structure.ts'
export { mapFormulaReferenceAreas, translateFormulaReferenceAreas, transformFormulaReferenceAreasForAxisEdit,
  type FormulaReferenceArea } from './features/formula/formula-reference-areas.ts'
export { SET_SHEET_CONDITIONAL_FORMATS_COMMAND, type SetSheetConditionalFormatsInput,
  type SheetConditionalFormatSubmitResult } from './features/conditional-formatting/conditional-format-command.ts'
export type { SheetFontFormatAction, SheetFontFormatCommandInput, SheetFontFormatCommandParams,
  SheetFontPrimaryPosition, SheetFontRangeInput, SheetFontMark } from './features/sheet-workbench/sheet-font-format-command.ts'
export { type ApiCellData, type ApiCellValue } from './api/model-utils'
export type { DocumentEditorCommandRequest, DocumentEditorCommandResult } from './api/document-editor-commands.ts'
export type { DocumentCommand } from './features/document/commands/document-commands.ts'
export type { DocumentPoint, DocumentSelection } from './features/document/model/document-model.ts'
export type { DocumentStoryRange } from './features/document/model/document-story.ts'
export { prepareDocumentStoryClearFormatting, applyDocumentStoryClearFormatting,
  type DocumentStoryClearFormattingPlan } from './features/document/services/document-clear-formatting.ts'
export { splitSheetA1Reference, parseSheetA1Range, type SheetA1Reference, type SheetA1Range } from './features/sheet-workbench/sheet-a1-reference.ts'

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

