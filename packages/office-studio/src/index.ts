import './style.css'
export { OfficeFontListKey, type OfficeFontList } from './platform/ribbon/ribbon-fonts'
export { WORKBENCH_CONTROL_THEME } from './platform/workbench-theme'

export { OfficeStudio, type OfficeStudioOptions, type OfficeCommandHandler, type WorkbookCommitOptions } from './api/office-studio'
export type { WorkbookInput } from './api/model-utils.ts'
export { FOverGridImage, FOverGridImageBuilder, type PreparedSheetImage, type SheetImageAsyncOptions } from './api/sheet-images.ts'
export { SHEET_IMAGE_COMMANDS, type SheetImageCommandInput, type SheetImageCommandId } from './features/sheet-images/sheet-image-commands.ts'
export { SHEET_DRAWING_PLUGIN, SheetImageSourceTypes, workbookWithPublicSheetImages, sheetCellImageToPublicBody,
  type SheetImageSourceType, type UniverSheetImageDrawing, type UniverSheetImageBody, type UniverWorksheetBackgroundImage,
  type UniverSheetCellData, type UniverWorksheetData, type UniverWorkbookData } from './features/sheet-images/sheet-image-public.ts'
export { prepareWorkbookSheetImageInput, type SheetImageInputPreparationOptions } from './features/sheet-images/sheet-image-input-preparation.ts'
export { prepareSheetImageFiles, prepareSheetImageUrl, prepareSheetImageResources } from './features/sheet-images/sheet-image-io.ts'
export { prepareSheetClipboardHtmlImages, type SheetClipboardImagePreparationOptions } from './features/sheet-images/sheet-image-html-preparation.ts'
export { SHEET_IMAGE_OBJECT_COMMANDS, sheetImageObjectClipboardFromHtml, type SheetImageObjectClipboard } from './features/sheet-images/sheet-image-object-clipboard.ts'
export { SHEET_DRAWING_ALIGN_TYPES } from './domain/sheet-drawings.ts'
export type { SheetDrawingAlignType, SheetDrawingMoveDirection, SheetDrawingLeafKind, SheetDrawingKind, SheetDrawingLeafRef, SheetDrawingGroupRef, SheetDrawingRef,
  SheetDrawingMatrix, SheetDrawingAffineGeometry, SheetDrawingGroup, SheetDrawingState, SheetDrawingShapeDTO } from './domain/sheet-drawings.ts'
export { DRAWING_AFFINE_IDENTITY, normalizeDrawingAffineMatrix, normalizeDrawingAffineGeometry, multiplyDrawingAffine,
  invertDrawingAffine, drawingAffinePoint, drawingAffineBounds, drawingAffineFrame, drawingAffineWithFrame, drawingAffineFromFrame } from './domain/drawing-affine.ts'
export type { DrawingAffineMatrix, DrawingAffineGeometry, DrawingAffineBounds } from './domain/drawing-affine.ts'
export { chartDrawingWorldPoint, chartDrawingViewportPoint, chartDrawingViewportBounds, chartDrawingLocalPoint, chartDrawingLocalDelta,
  chartDrawingOutlineCommands, chartDrawingKeyboardTranslation } from './features/charts/chart-drawing-affine.ts'
export type { ChartDrawingProjection, ChartDrawingViewport } from './features/charts/chart-drawing-affine.ts'
export { SHEET_DRAWING_GROUP_COMMAND_IDS, DRAWING_GROUP_MENU, DRAWING_ARRANGE_MENU, DRAWING_ALIGN_MENU, drawingPaintOrder, drawingHitOrder,
  planDrawingAlign, planDrawingMove, sheetDrawingMoveDirectionForKey, planDrawingGroup, planDrawingUngroup, planDrawingArrange, planDrawingGroupTransform, planDrawingSelectionTransform, applyDrawingTransaction } from './features/drawing-groups/index.ts'
export type { SheetDrawingAlignRequest, SheetDrawingMoveRequest } from './features/drawing-groups/index.ts'
export type { SheetDrawingArrangeType, SheetDrawingScene, SheetDrawingObject, SheetDrawingTransaction, SheetDrawingPlanResult } from './features/drawing-groups/index.ts'
export type { SheetDrawingCommandId, SheetDrawingCommandInput, SheetDrawingCommandParams, SheetDrawingPublicRef, SheetDrawingPublicGroupUpdate } from './api/sheet-drawing-commands.ts'
export type { SlicerCommandContext } from './api/slicer-command-context.ts'
export { SheetDrawingAnchorTypes } from './domain/sheet-images.ts'
export type { SheetImageResource, SheetCellImage, SheetFloatingImage, SheetBackgroundImage, SheetDrawingAnchorType,
  SheetImageAnchor, SheetImageCellMarker, SheetImageBounds, SheetImagePlacement, SheetImagePlacementInput, SheetImageSourceRect } from './domain/sheet-images.ts'
export type { SheetFloatingImagePatch } from './features/sheet-images/sheet-image-placement.ts'
export { SHEET_IMAGE_CROP_TYPES, SHEET_IMAGE_CROP_PRESETS, type SheetImageCropType } from './features/sheet-images/sheet-image-crop.ts'
export { SHEET_IMAGE_CROP_OPERATIONS, type SheetImageCropOperation, type SheetImageCropResult } from './features/sheet-images/sheet-image-crop-session.ts'
export { SHEET_IMAGE_ALIGN_TYPES, SHEET_IMAGE_ALIGN_OPERATIONS, SHEET_IMAGE_ALIGN_PRESETS,
  sheetImageSelectionBounds, transformSheetImageSelection, alignSheetImageSelection,
  type SheetImageAlignType } from './features/sheet-images/sheet-image-selection.ts'
export type { SaveCellImagesOptions, SheetImageSaveConfig } from './features/sheet-images/sheet-image-save.ts'
export { SHEET_DEFAULT_STYLE_COMMANDS, type SheetDefaultStyleCommandInput } from './api/sheet-default-style-commands.ts'
export { composeSheetCellStyles, resolveSheetCellStyle, resolveSheetStyleSource, sheetInheritedCellStyle,
  normalizeSheetStyleInput, sheetCellDirectStyle, type SheetStyleInput, type SheetStyleContext } from './features/sheet-workbench/sheet-style-inheritance.ts'
export type { SheetCellStylePatch, SheetCellStyleProperty, SheetStyleSource } from './domain/model.ts'
export { sheetStyleToPublicData, sheetStyleDataWithPublicFields,
  type SheetPublicStyleData, type SheetApiStyleData } from './features/sheet-workbench/sheet-public-style.ts'
export { normalizeSheetPublicAxisData, applySheetPublicAxisInput,
  type SheetPublicAxisData } from './features/sheet-workbench/sheet-public-axis.ts'
export { SHEET_RANGE_THEME_COMMANDS, type SheetRangeThemeCommandInput } from './api/sheet-range-theme-commands.ts'
export { OPEN_SHEET_RANGE_THEME_OPERATION, planSheetRangeThemeCommand, sheetRangeThemeCommandInput, type SheetRangeThemeCommandPlan } from './features/sheet-workbench/sheet-range-theme-command.ts'
export { RangeThemeStyle, normalizeSheetRangeThemeStyle, sheetRangeThemeToJson, type SheetRangeThemeStyleJson } from './features/sheet-workbench/sheet-range-theme-style.ts'
export { SHEET_BUILTIN_RANGE_THEME_NAMES } from './features/sheet-workbench/sheet-range-theme-presets.ts'
export { resolveSheetRangeThemeStyle, sheetRangeThemeHiddenRowSpans, type SheetRangeThemeRegionSnapshot } from './features/sheet-workbench/sheet-range-theme-resolver.ts'
export { planSheetRangeThemeCellShift, planSheetRangeThemeAxisEdit, planSheetRangeThemeCut, sheetRangeThemeChangedRanges,
  assertSheetRangeThemeStyleRange, type SheetRangeThemeCellShift } from './features/sheet-workbench/sheet-range-theme-structure.ts'
export { SHEET_RANGE_THEME_SECTIONS, type SheetRangeThemeSection, type SheetRangeThemeStyle, type SheetRangeThemeRule } from './domain/model.ts'
export { captureSheetStyleRegion, sheetStyleRegionCellStyle, type SheetStyleRegionSnapshot } from './features/sheet-workbench/sheet-style-region.ts'
export { COPY_WORKSHEET_COMMAND, planWorksheetCopy, createWorksheetCopyName, remapWorksheetCopyFormula,
  type CopyWorksheetCommandInput, type WorksheetCopyPlan } from './features/sheet-workbench/sheet-worksheet-copy.ts'
export { WORKSHEET_LIFECYCLE_COMMANDS, planWorksheetLifecycle, worksheetHiddenState,
  type WorksheetHiddenState, type WorksheetLifecycleCommandId, type WorksheetLifecycleCommandInput,
  type WorksheetLifecyclePlan } from './features/sheet-workbench/sheet-worksheet-lifecycle.ts'
export { SET_DOCUMENT_CUSTOM_XML_PARTS_COMMAND, type DocumentCustomXmlPartsWriteOptions,
  type DocumentCustomXmlPartsCommandInput } from './api/document-custom-xml-commands'
export { DOCUMENT_NOTE_COMMANDS, DocumentNoteCommandRejectedError, type DocumentNoteCommandResult,
  type DocumentNoteCommandContext, type DocumentNoteCommandParams } from './api/document-note-commands'
export type { DocumentNoteMutationRequest, DocumentNoteInsertOptions, DocumentNoteKind, DocumentNotePatch } from './features/document/notes/document-note-types'
export { decodeDocumentBody, encodeDocumentBody } from './api/document-body-codec'
export { FDocumentLatex } from './api/document-latex'
export { FDocumentImage } from './api/document-images'
export type { DocumentImageInsertOptions, DocumentImageData, DocumentImageTextRange } from './api/document-images'
export { DOCUMENT_IMAGE_PUBLIC_COMMANDS } from './api/document-image-commands'
export type { DocumentImageCommandBase, DocumentImageOperationRequest, DocumentImagePublicInsertRequest, DocumentImageTransformItem,
  DocumentImagePublicTransformRequest, DocumentImagePublicWrappingRequest, DocumentImagePublicArrangeRequest, DocumentImagePublicRemoveRequest,
  DocumentImagePublicAlignRequest, DocumentImagePublicBatchTransformRequest } from './api/document-image-commands'
export { DOCUMENT_IMAGE_ALIGN_TYPES, documentImageBatchMembers, documentImageBatchBounds } from './features/document/images/document-image-batch'
export type { DocumentImageAlign, DocumentImageBatchTransform, DocumentImageBatchMember, DocumentImageBatchUpdate } from './features/document/images/document-image-batch'
export type { DocumentImageGeometryReadOptions, DocumentImageLayoutRead } from './api/document-image-layout'
export type { DocumentImageEditor, DocumentImageEditorHandle, DocumentImageEditorTarget } from './api/document-image-editors'
export type { DocumentImagePosition, DocumentImageWrappingStyle, DocumentImagePatch, DocumentImageArrange, DocumentImageSourceRect } from './features/document/images/document-image-model'
export { DOCUMENT_LATEX_PUBLIC_COMMANDS } from './api/document-latex-commands'
export { DOCUMENT_LATEX_TEMPLATE_CATEGORIES, DOCUMENT_LATEX_CATEGORY_LABELS, DOCUMENT_LATEX_TEMPLATES,
  DOCUMENT_LATEX_TEMPLATE_OPERATION, DOCUMENT_LATEX_TEMPLATE_LOCAL_OPERATION, getDocumentLatexTemplate,
  getDocumentLatexTemplates, isDocumentLatexTemplateCategory, insertDocumentLatexTemplate } from './features/document/latex/latex-templates'
export type { DocumentLatexTemplateCategory, DocumentLatexTemplate, DocumentLatexTemplateInsertion } from './features/document/latex/latex-templates'
export type { DocumentLatexTemplateEditorHandle } from './api/document-latex-template-editors'
export { DOCUMENT_LATEX_POPUP_OPERATIONS } from './features/document/controllers/document-latex-popup'
export type { DocumentLatexPopupOperationRequest, DocumentLatexPopupTarget, DocumentLatexPopupTextRange } from './features/document/controllers/document-latex-popup'
export type { DocumentLatexPopupEditor, DocumentLatexPopupEditorHandle } from './api/document-latex-popup-editors'
export type { DocumentLatexTemplateOperationRequest, DocumentLatexPublicInsertRequest, DocumentLatexPublicUpdateRequest,
  DocumentLatexPublicRemoveRequest, DocumentLatexPublicReplaceRequest } from './api/document-latex-commands'
export type { DocumentLatexFormulaConfig, DocumentLatexCreateOptions, DocumentLatexInsertOptions, DocumentLatexUpdateOptions,
  DocumentLatexFindQuery, DocumentLatexRange, DocumentLatexInfo } from './api/document-latex'
export type { UniverDocumentBody, UniverDocumentBodyData, UniverDocumentBodyResources, DocumentBodyDecodeResult,
  DocumentBodyDecodeOptions, DocumentBodyEncodeOptions, DocumentBodyEncodeResult } from './api/document-body-codec'
export { documentToPublicSnapshot } from './api/document-output'
export { documentParagraphPagination, documentUsesTraditionalPagination, DOCUMENT_PARAGRAPH_PAGINATION_FIELDS,
  type DocumentParagraphPagination, type DocumentParagraphPaginationField } from './features/document/model/document-paragraph-pagination'
export type { UniverDocumentData, UniverDocumentNoteData, UniverDocumentNoteSource, OfficeDocumentSnapshotSource } from './api/document-snapshot-types'
export type { UniverDocumentTextStyle, UniverDocumentTextStyleRun, UniverDocumentTextRangeDescription,
  UniverDocumentColorStyle, UniverDocumentTextDecoration } from './api/document-text-style'
export type { DocumentNoteContentInput, PublicDocumentNoteInput, PublicDocumentNoteDecoration,
  PublicDocumentFootnoteSettings, PublicDocumentEndnoteSettings } from './api/document-note-body'
export { DocumentCustomXmlError, normalizeDocumentCustomXmlParts, canonicalDocumentCustomXmlPartName, isDocumentCustomXmlContentType,
  documentCustomXmlPartKey, documentCustomXmlRelationshipPart, documentCustomXmlRelationshipOwner,
  resolveDocumentCustomXmlTarget, documentCustomXmlUtf8Bytes, documentCustomXmlEncodedBytes, documentCustomXmlResourceByteLength,
  DOCUMENT_CUSTOM_XML_MAX_PART_BYTES, DOCUMENT_CUSTOM_XML_MAX_TOTAL_BYTES, DOCUMENT_CUSTOM_XML_MAX_PARTS,
  DOCUMENT_CUSTOM_XML_MAX_RELATED_PARTS, DOCUMENT_CUSTOM_XML_MAX_ELEMENTS, DOCUMENT_CUSTOM_XML_MAX_DEPTH,
  DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH } from './domain/document-custom-xml'
export { CONTENT_HISTORY_COMMANDS, CONTENT_HISTORY_COMMAND_IDS, ContentHistoryCanceledError, contentHistoryCommandAction, resolveContentHistoryTarget,
  type ContentHistoryCommandInput, type ContentHistoryDirection, type ContentHistoryTarget } from './api/content-history-commands.ts'
export {
  FBlob,
  FChart,
  FSheetPivotChart,
  ChartConfigBuilder,
  PivotChartConfigBuilder,
  FDocument,
  FDocumentContentControl,
  FDocumentParagraph,
  FDocumentTextRange,
  FSheetSlicer,
  FSheetTimeline,
  FPivotTable,
  FRange,
  FSelection,
  FOffice,
  FWorkbook,
  FWorksheet,
  PositionTypeEnum,
  type CreateUnitOptions,
  type ChartInfo,
  type ChartPatch,
  type ChartComputationReadOptions,
  type DataValidationError,
  type DataValidationRule,
  type DataValidationRuleInput,
  type DocumentBodySegment,
  type DocumentBodySnapshot,
  type DocumentContentControlCreateInfo,
  type DocumentContentControlInfo,
  type DocumentContentControlPatch,
  type DocumentNoteInfo,
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
export { FSheetSlicerBuilder, FSheetTimelineBuilder } from './api/slicer-builders.ts'
export type { FLinkedSheetSlicerBuilder, FBoundSheetSlicerBuilder, FLinkedSheetTimelineBuilder, FBoundSheetTimelineBuilder } from './api/slicer-builders.ts'
export { FEventName, EventHub, type ApiEventMap, type ApiRange, type ApiSheetSelection, type EventDeliveryError, type Disposable } from './api/events'
export { SET_SHEET_SELECTIONS_OPERATION, SELECT_SHEET_RANGE_COMMAND, MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND,
  type SheetSelectionOperationRequest, type SheetSelectionShortcutRequest } from './api/sheet-selection-commands.ts'
export { normalizeSheetSelection, reconcileSheetSelection, resolveVisibleSheetSelection, type SheetSelectionState, type SheetSelectionRange,
  type SheetSelectionCell, type SheetSelectionOptions, type SheetSelectionSheet } from './features/sheet-workbench/sheet-selection.ts'
export { nextSheetSelectionCell, planSheetSelectionShortcut, SheetSelectionDirections, type SheetSelectionDirection,
  type SheetSelectionNavigationResult, type SheetSelectionNavigationSheet, type SheetSelectionShortcutKey,
  type SheetSelectionShortcutPlan } from './features/sheet-workbench/sheet-selection-navigation.ts'
export { COMMIT_SHEET_CELL_INPUT_COMMAND, planSheetCellInput, sheetCellInputCommandInput, sheetCellInputConflictMessage,
  type SheetCellInputCommandInput, type SheetCellInputCommandResult, type SheetCellInputOptions,
  type SheetCellInputPlan, type SheetCellInputDefaults } from './features/sheet-workbench/sheet-cell-input-command.ts'
export { createSheetCellEditSession, getSheetCellEditConflict, resolveSheetDate1904,
  type SheetCellEditSession, type SheetCellEditContext, type SheetCellEditConflict } from './features/sheet-workbench/sheet-cell-edit-session.ts'
export { sheetCellContentKey, sheetCellInputText, createSheetCellEditorStart } from './features/sheet-workbench/sheet-cell-edit.ts'
export { editSheetCellRichText } from './features/sheet-workbench/sheet-cell-rich-text-edit.ts'
export { sheetCellRichTextFontValue, sheetCellRichTextRunView, sheetCellFontValues, setSheetCellRichTextFont, clearSheetCellRichTextFonts,
  type SheetRichTextFontField, type SheetRichTextFontValue } from './features/sheet-workbench/sheet-cell-rich-text-style.ts'
export { layoutSheetCellRichText, type SheetCellRichTextLayout, type SheetCellRichTextLayoutOptions,
  type SheetRichTextFragment, type SheetRichTextLine, type SheetRichTextColumn } from './features/sheet-workbench/sheet-cell-rich-text-layout.ts'
export { formatSheetCellRichTextRange, replaySheetCellInlineFormats, readSheetCellInlineFormats, SheetCellInlineFormatDraft,
  type SheetCellInlineFormatEdit, type SheetCellInlineFormatReplay } from './features/sheet-workbench/sheet-cell-inline-format.ts'
export { FORMAT_SHEET_CELL_INPUT_OPERATION, FORMAT_SHEET_CELL_RICH_TEXT_COMMAND, planSheetCellInlineFontFormat, planSheetCellInlineStyle,
  type SheetCellInlineStylePatch, type SheetCellInlineStyleCommandInput, type SheetCellInlineFontFormatInput,
  type SheetCellInlineFontFormatResult } from './features/sheet-workbench/sheet-cell-inline-format-command.ts'
export { SHEET_AUTO_FILL_COMMAND, SHEET_REFILL_COMMAND, SHEET_COPY_DOWN_COMMAND, SHEET_COPY_RIGHT_COMMAND,
  SHEET_AUTO_CLEAR_COMMAND, SHEET_FILL_COMMAND_IDS, SHEET_FILL_SYNC_CELL_BUDGET, createSheetAutoFillPlan,
  type SheetAutoFillResult, type SheetAutoFillPlan } from './features/sheet-workbench/sheet-auto-fill.ts'
export { parseSheetFillCommandInput, sheetFillApplyType, type SheetFillCommandInput } from './features/sheet-workbench/sheet-fill-command.ts'
export { resolveSheetFillTarget, sheetFillPointerTarget, sheetCopyFillRanges, type SheetFillTarget } from './features/sheet-workbench/sheet-fill-target.ts'
export type { SheetFillApplyType } from './features/sheet-workbench/sheet-fill.ts'
export type { SheetFillOptionsView } from './features/sheet-workbench/sheet-fill-store.ts'
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
export { SHEET_CELL_SHIFT_COMMANDS, SHEET_CELL_SHIFT_CONFIRM_COMMANDS, SHEET_CELL_SHIFT_COMMAND_IDS, SheetCellShiftDimensions,
  planWorkbookCellShiftCommand, sheetCellShiftCommandInput, sheetCellShiftCommandRange, sheetCellShiftDirectionForCommand, sheetCellShiftEdit,
  type SheetCellShiftCommandInput, type SheetCellShiftConfirmationRequest, type SheetCellShiftDimension, type SheetCellShiftDirection,
  type SheetCellShiftPlan } from './features/sheet-workbench/sheet-cell-shift-command.ts'
export { assertSheetCellShiftEdit, sheetCellShiftAffectedRange, sheetCellShiftContainsBand, mapSheetCellShiftPosition,
  inverseSheetCellShiftPosition, transformSheetCellShiftRange, type SheetCellShiftEdit } from './domain/sheet-cell-shift.ts'
export { SHEET_AXIS_VISIBILITY_COMMANDS, SHEET_AXIS_VISIBILITY_COMMAND_IDS, SheetAxisVisibilityError, planSheetAxisVisibilityCommand,
  type SheetAxisVisibilityCommandInput, type SheetAxisVisibilityAction, type SheetAxisVisibilityPlan } from './features/sheet-workbench/sheet-axis-visibility.ts'
export { type SheetAxis, type SheetAxisEdit, type SheetAxisInterval, sheetAxisMaximum, mapSheetAxisIndex,
  transformSheetAxisInterval, transformSheetAxisRange, transformSheetAxisReferenceInterval } from './domain/sheet-axis-structure.ts'
export { mapFormulaReferenceAreas, translateFormulaReferenceAreas, transformFormulaReferenceAreasForAxisEdit, transformFormulaReferenceAreasForCellShift,
  type FormulaReferenceArea } from './features/formula/formula-reference-areas.ts'
export { SET_SHEET_CONDITIONAL_FORMATS_COMMAND, type SetSheetConditionalFormatsInput,
  type SheetConditionalFormatSubmitResult } from './features/conditional-formatting/conditional-format-command.ts'
export type { SheetFontFormatAction, SheetFontFormatCommandInput, SheetFontFormatCommandParams,
  SheetFontPrimaryPosition, SheetFontRangeInput, SheetFontMark } from './features/sheet-workbench/sheet-font-format-command.ts'
export { type ApiCellData, type ApiCellValue } from './api/model-utils'
export type { DocumentEditorCommandRequest, DocumentEditorCommandResult } from './api/document-editor-commands.ts'
export type { DocumentFieldCacheAdapter, DocumentFieldCacheState, DocumentFieldCacheRefreshRequest,
  DocumentFieldCacheRefreshResult } from './features/document/services/document-field-open-refresh.ts'
export type { DocumentCommand } from './features/document/commands/document-commands.ts'
export type { DocumentStructureCommand, DocumentColumnGroupCreateOptions } from './features/document/commands/document-structure-commands.ts'
export { registerDocumentCodeFontFamilies, type DocumentCodeFontFamilyOption } from './features/document/code-blocks/document-code-fonts.ts'
export { DOCUMENT_STRUCTURED_CODE_LANGUAGES, type DocumentStructuredCodeLanguage } from './features/document/code-blocks/document-code-languages.ts'
export { FDocumentStructuredBlock, FDocumentQuote, FDocumentCallout, FDocumentCode,
  type DocumentStructuredKind, type DocumentStructuredFindQuery, type DocumentStructuredRange, type DocumentStructuredInsertOptions,
  type DocumentQuoteInsertOptions, type DocumentCalloutInsertOptions, type DocumentCodeInsertOptions,
  type DocumentQuoteStyle, type DocumentCalloutStyle, type DocumentCalloutBorderStyle, type DocumentCalloutBorderPatch } from './api/document-structured-blocks.ts'
export { FDocumentColumnGroup, FDocumentColumn, DocsColumnPositions,
  type DocumentColumnFindQuery, type DocumentColumnInsertOptions, type DocumentColumnOffsetRange,
  type DocumentColumnGroupOffsetRange, type DocumentColumnContentRange } from './api/document-columns.ts'
export type { DocumentPoint, DocumentSelection, DocumentTableViewSelection, DocumentTable, DocumentTableRow, DocumentCell, DocumentTableContinuationContent,
  DocumentTableStyle, DocumentTableCellStyle, DocumentTableBorder, DocumentTableBorders, DocumentTableCellMargin,
  DocumentTableColumnTypeConfig, RunPatch, DocumentRunClearField } from './features/document/model/document-model.ts'
export { FDocumentTable, FDocumentTableCell, DocsTableSelectionKinds, DocsTableRowInsertPositions, DocsTableColumnInsertPositions,
  DocsTableSortDirections, DocsTableColumnTypes, DocsTableRowHeightRules, DocsTableBorderPresets,
  type DocumentTableFindQuery, type DocumentTableInfo, type DocumentTableDescription, type DocumentTableMetadata,
  type DocumentTableCellContentRange, type DocumentTableCellOffsetRange, type DocumentTableOffsetRange } from './api/document-tables.ts'
export type { DocumentTableInsertOptions, DocumentTableBorderOptions, DocumentTableStylePatch, DocumentTableCommand,
  DocumentTableCommandTarget } from './features/document/commands/table-command-types.ts'
export type { DocumentTableRowPropertiesPatch } from './features/document/model/document-table-style.ts'
export type { DocumentTableCellRange, DocumentTableGridCell } from './features/document/model/document-table-grid.ts'
export { documentTableCellContentContainers, documentTableCellBlocks, documentTableContinuationEntries,
  assertDocumentTableContinuationRows } from './features/document/model/document-table-continuations.ts'
export type { DocumentStoryRange } from './features/document/model/document-story.ts'
export { SET_DOCUMENT_INPUT_SELECTION_OPERATION, CLEAR_DOCUMENT_INPUT_MARKS_OPERATION, ADVANCE_DOCUMENT_INPUT_MARKS_OPERATION,
  type DocumentInputLocation, type DocumentInputSession, type DocumentStoryInputCommit } from './features/document/model/document-input-session.ts'
export type { DocumentInputSelectionRequest } from './api/document-input-commands.ts'
export { useDocumentInputStore } from './features/document/stores/document-input-store.ts'
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
  getPivotShowDataAs,
  getPivotShowDataAsBaseItems,
  calculatePivotReport,
  PivotOutputConflictError,
  getPivotTableAtCell,
  pivotDrillDown,
  pivotFieldIds,
  pivotOutputValues,
  createPivotTable,
  refreshPivotTable,
  setPivotField,
  setPivotFieldSetting,
  setPivotShowDataAs,
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
export { applyPivotChartEdit, createPivotChartReport, getPivotChartReport, getPivotChartConfiguredField, getPivotChartFieldItems, getPivotChartShowDataAsBaseItems, PivotChartEditError,
  type PivotChartFieldsConfig, type PivotChartFieldSetting, type PivotChartReportEdit } from './features/pivot/services/pivot-chart-engine.ts'
export { PIVOT_CHART_COMMANDS, type PivotChartCommandContext, type PivotChartConfigureInput, type PivotChartInfo,
  type PivotChartInsertInfo, type PivotChartInsertOptions, type LinkedPivotChartInfo, type PivotChartSourceSpec, type PivotChartTableReference } from './features/pivot/services/pivot-chart-api-types.ts'
export { PIVOT_SHOW_DATA_AS_RULES, getPivotShowDataAsRule, getPivotShowDataAsBaseFieldIds, type PivotShowDataAsRule } from './features/pivot/services/pivot-show-data-as.ts'
export { assertPivotReportReadPermission, assertPivotReportEditPermission } from './features/pivot/services/pivot-report-permissions.ts'
export { synchronizeLinkedPivotChartReports } from './features/pivot/services/pivot-chart-report-lifecycle.ts'
export { getPivotAxisState, getPivotCollapse, getPivotCollapseItems, setPivotCollapse, setPivotCollapseInfo, setPivotCategoryLevel, stepPivotCategoryLevel } from './features/pivot/services/pivot-engine.ts'
export { getPivotChartAxisState, getPivotChartCollapseItems } from './features/pivot/services/pivot-chart-engine.ts'
export { PIVOT_HIERARCHY_COMMANDS, type PivotChartHierarchyTarget, type PivotReportHierarchyTarget } from './features/pivot/services/pivot-hierarchy-api-types.ts'
export type { PivotAxisArea, PivotAxisFieldState, PivotAxisState } from './features/pivot/services/pivot-axis-state.ts'
export { normalizePivotChartControls, pivotChartControls } from './features/pivot/services/pivot-chart-controls.ts'
export { PIVOT_CHART_FIELD_COMMANDS } from './features/pivot/services/pivot-chart-field-command-types.ts'
export type { PivotChartFieldCommandTarget, PivotChartFieldCommandParams, PivotChartFieldCommandInput } from './features/pivot/services/pivot-chart-field-command-types.ts'
export type { PivotReportCommandContext, PivotReportCommandTarget } from './features/pivot/services/pivot-report-api-types.ts'
export { getPivotValuesPosition, setPivotValuesPosition, getPivotSourceAvailableRows } from './features/pivot/services/pivot-engine.ts'
export { getPivotValuesLayout, normalizePivotValuesPosition, normalizePivotValuePositionInput, PIVOT_VALUE_POSITIONS } from './features/pivot/services/pivot-values-position.ts'
export type { PivotValuesPositionState } from './features/pivot/services/pivot-values-position.ts'
export { normalizePivotDateFilterInput, normalizePivotLabelConditionInput } from './features/pivot/services/pivot-filter-api.ts'
export { PIVOT_REPORT_FILTER_COMMANDS } from './api/pivot-filter-commands.ts'
export type { PivotReportManualFilterInput, PivotReportDateFilterInput, PivotReportLabelConditionInput } from './api/pivot-filter-commands.ts'
export type { PivotDateFilterInput, PivotDateFilterApiInfo, PivotDateFilterCode, PivotLabelConditionInput,
  PivotLabelConditionApiInfo, PivotLabelConditionCode } from './features/pivot/services/pivot-filter-api.ts'
export { getPivotDateFilter, getPivotLabelConditionFilter, getPivotManualFilterItemState, getPivotFieldFilterCapabilities,
  getPivotGroupedFieldItems, getPivotSourceRecordSelection, setPivotDateFilter, setPivotLabelConditionFilter, setPivotSourceMemberFilter } from './features/pivot/services/pivot-engine.ts'
export type { PivotChartControls } from './domain/model.ts'
export type { PivotChartAxisMetadata } from './features/charts/chart-pivot-axis.ts'
export {
  clearSlicerFilter,
  createSlicer,
  getSlicer,
  getSlicerItems,
  readSlicerItemState,
  readSlicerRecordSelection,
  readSlicerSelection,
  removeSlicer,
  setSlicerItems,
  setSlicerItemStates,
  setSlicerRecordSelection,
  setSlicerTimelineRange,
  setSlicerConnections,
} from './features/slicer/slicer-engine'
export { getSlicerConnections, getSlicerConnectionCandidates, getSlicerCache, getSlicerSourceValues, getSlicerSourceRecords,
  getSlicerSourceRecordValues, getSlicerSourceCellRange, getSlicerTableNativeFilter, planSlicerConnections, SlicerConnectionError } from './features/slicer/slicer-connections.ts'
export { normalizeSlicerCacheOptions, normalizeSlicerCacheOptionsPatch, readSlicerCacheOptions, patchSlicerCacheOptions,
  effectiveSlicerCacheOptions, applySlicerCacheItemOptions, createSlicerMemberComparator } from './features/slicer/slicer-cache-options.ts'
export { normalizeSlicerSettingsPatch, readSlicerSettings } from './features/slicer/slicer-settings.ts'
export type { SheetSlicerSettings } from './features/slicer/slicer-settings.ts'
export { SLICER_PUBLIC_COMMANDS, normalizeSheetSlicerAnchor, normalizeSheetSlicerPlacement, normalizeSheetSlicerSource,
  normalizeSheetSlicerInfo, normalizeSheetTimelineInfo, resolveSheetSlicerSource, readSheetSlicerInfo, readSheetTimelineInfo,
  readSheetSlicerSource, readSheetSlicerDrawing, sheetSlicerCacheDefinition, sheetTimelineCacheDefinition } from './features/slicer/slicer-public-info.ts'
export type { ISheetSlicerInfo, ISheetTimelineInfo, ISheetSlicerDrawing, SheetSlicerAnchorSpec, SheetSlicerSourceSpec,
  SheetTimelineSourceSpec, SheetSlicerPlacement, SlicerCacheDefinition, ITimelineCacheDefinition, TimelinePresentationPatch } from './features/slicer/slicer-public-info.ts'
export { normalizeTimelineSettingsPatch, readTimelineSettings } from './features/slicer/timeline-settings.ts'
export type { SheetTimelineSettings, TimelineSettingsInputPatch } from './features/slicer/timeline-settings.ts'
export { TimelinePublicLevel, TimelineDateFilterOperator } from './domain/timeline-filters.ts'
export type { ITimelineRange, ITimelineDateCondition, TimelineDateFilterState } from './domain/timeline-filters.ts'
export { normalizeTimelineLevel, toTimelinePublicLevel, normalizeTimelineDateRange, effectiveTimelineDateRange,
  compileTimelineSerialRange, compileTimelineRange, normalizeTimelineDateCondition, compileTimelineSerialCondition,
  compileTimelineCondition, timelineConditionIntersectsPeriod, timelineDateConditionToPivot, timelineDateConditionFromPivot } from './features/slicer/timeline-date-filter.ts'
export { timelinePeriodRange, timelineFilterStateFromLegacy } from './features/slicer/timeline-periods.ts'
export { readTimelineFilterState, readTimelineDateRange, readTimelineDateCondition, setTimelineDateRange,
  setTimelineDateCondition, updateTimelineCacheState, applyTimelineReportFilter } from './features/slicer/timeline-filter.ts'
export { slicerMemberViewport } from './features/slicer/slicer-member-viewport.ts'
export { SlicerCanvasDataCache } from './features/slicer/slicer-canvas-data.ts'
export type { SlicerCanvasDataOwner, SlicerCanvasDataLoader, SlicerCanvasMemberWindow, SlicerCanvasDataState,
  SlicerCanvasDataOptions } from './features/slicer/slicer-canvas-data.ts'
export { timelineCanvasSelectionRange, timelineCanvasHandleRange } from './features/slicer/slicer-canvas-gesture.ts'
export { RecordValueType, FilterSelectionMode, getRecordValueKey, normalizeRecordValue, normalizeRecordSelection, compileRecordSelection } from './domain/record-value.ts'
export type { RecordValue, RecordSelection, PivotRecordValue } from './domain/record-value.ts'
export { SlicerMemberQueryService, SlicerMemberQueryError, normalizeSlicerMemberWindowRequest } from './features/slicer/slicer-member-query.ts'
export type { ISheetSlicerItemsRequest, ISlicerSourceMemberWindow, ISlicerDisplayMemberWindow, ISlicerSourceMetadata, SlicerDisplayMember,
  SlicerMemberQueryOptions, SlicerMemberQueryMetrics } from './features/slicer/slicer-member-query.ts'
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
  applyExcelCellView,
  applyExcelWorkbookView,
  cellKey,
  parseCellKey,
  toExcelWorkbookView,
  toExcelCellView,
  type ExcelCellViewMappingOptions,
  type ExcelCellView,
  type ExcelSheetView,
  type ExcelWorkbookView,
} from './services/excelViewAdapter'
export {
  downloadOfficeBytes,
  exportOfficeModel,
  officeDownloadName,
  OfficeExportError,
  resetOfficeWriter,
  type ExportOptions,
} from './services/exportService'
export { ParserError } from './parser/errors'
export { parseOfficeFile, resetParser, type ImportOptions } from './services/wasmAdapter.ts'
export { type ExportOutcome, type ExportProgress, type ExportWarning } from './parser/export-contracts'
export type { ImportedOfficeModel, ParseEnvelope, ParseProgress } from './parser/contracts.ts'
export { WASM_PROTOCOL_VERSION, createWasmTaskControl, WasmProtocolError } from './platform/wasm/task-protocol.ts'
export type { WasmTaskKind, WasmTaskControl, WasmTaskBudget, WasmTaskMetrics, BoundaryCopyStats, WorkerTransportStats } from './platform/wasm/task-protocol.ts'
export * from './compute/index.ts'
export { ChartStatisticsService, chartStatisticsSignature, type ChartStatisticsServiceOptions, type ChartStatisticsRequestContext,
  type ChartStatisticsRequest, type ChartStatisticsSnapshot, type ChartStatisticsInvalidation } from './features/charts/services/chart-statistics-service.ts'
export { prepareChartStatisticsTask, decodeChartStatistics, type ChartStatisticsPlan } from './features/charts/chart-statistics-compute.ts'
export { prepareChartComputeInputs } from './features/charts/chart-input-preparation.ts'
export { readChartRecordsAsync } from './features/charts/chart-source-compute.ts'
export type { ChartComputePreparationContext } from './features/charts/chart-compute-lifecycle.ts'
export type { ChartComputedStatistics } from './features/charts/chart-derived.ts'
export type { ChartConfig, ChartData, ChartSeriesData } from './features/charts/chart-engine.ts'
export { CHART_EDIT_PANEL_TABS, CHART_EDIT_PANEL_SECTIONS, CHART_EDIT_PANEL_AXES, CHART_EDIT_PANEL_TITLES,
  normalizeChartEditPanelRoute, resolveChartEditPanelRoute, chartEditPanelSeriesId,
  type ChartEditPanelRoute, type ChartEditPanelTab, type ChartEditPanelSection, type ChartEditPanelAxis,
  type ChartEditPanelTitle, type ChartElementPanelTarget } from './features/charts/chart-edit-panel-route.ts'
export { ChartLegendPositions, ChartLegendSelectModes, type ChartLegendSpec } from './features/charts/chart-legend.ts'
export { ChartAreaLineStyles, ChartInvalidValueTypes, CHART_AREA_LINE_STYLE_DEFAULT, CHART_INVALID_VALUE_STRATEGY_DEFAULT,
  supportsChartAreaLineStyle, supportsChartLineAndAreaSection, resolveChartAreaLineStyle } from './features/charts/chart-line-and-area.ts'
export type { ChartAreaLineStyle, ChartInvalidValueStrategy, ChartLineAndAreaOptions } from './domain/model'
export { ChartTitlePositions, ChartTextAlignments, type ChartTitleSpec, type ChartSubtitleSpec, type ChartTitleObjectSpec, type ChartTitleKind } from './features/charts/chart-titles.ts'
export type { ChartTitlePosition, ChartTextAlignment, ChartTitleStyleOptions, ChartSubtitleStyleOptions, ChartTitleManualLayout } from './domain/model'
export type { ChartPlotAreaOptions } from './domain/model'
export { describeChartConfiguration, chartDescriptionFieldIndexes,
  type ChartDescription, type ChartDescriptionType, type ChartDescriptionContext, type ChartDescriptionConfig,
  type ChartSeriesDescription, type ChartAxisBindingDescription, type ChartDescriptionAxis } from './features/charts/chart-description.ts'
export { CHART_LINEAR_GRADIENT_TYPE, CHART_GRADIENT_MAX_STOPS, chartLinearGradient, isChartLinearGradient,
  validateChartColor, chartColorRepresentative, type ChartLinearGradientInput } from './features/charts/chart-color.ts'
export { supportsChartExplicitGradient } from './features/charts/chart-gradient.ts'
export { ChartGaugeModes, ChartGaugePointerShapes, CHART_GAUGE_PRESETS, chartGaugePresetOptions,
  type ChartGaugePreset } from './features/charts/chart-gauge.ts'
export { ChartRadarShapes } from './features/charts/chart-radar.ts'
export { ChartVisualMapTypes, type ChartHeatmapVisualMapType, type ChartHeatmapLabelOptions } from './features/charts/chart-heatmap-api.ts'
export { ChartWaterfallStyleTargets, initializeChartWaterfallOptions } from './features/charts/chart-waterfall.ts'
export { chartParetoValuePercentages } from './features/charts/chart-pareto.ts'
export { CHART_POINT_SHAPE_CHOICES, CHART_POINT_SHAPES, isChartPointShape, chartMarkerShape, chartMarkerPoints,
  chartMarkerRequiresSharedGeometry, chartNativeMarkerSymbol, type ChartMarkerShape, type ChartMarkerVector } from './features/charts/chart-markers.ts'
export { CHART_LABEL_CONTENT_CHOICES, CHART_LABEL_POSITION_CHOICES, chartLabelContentMask, chartLabelContentValues,
  chartLabelContentFromValues, chartPointLabelAnchor, type ChartLabelContent, type ChartPointLabelAnchor,
  type ChartPointLabelLocation } from './features/charts/chart-labels.ts'
export { layoutChartWaterfall } from './features/charts/chart-waterfall-layout.ts'
export { CHART_WATERFALL_POINT_CHOICE_LIMIT, chartWaterfallLayoutControls, chartWaterfallPointRole, chartWaterfallPointRoleControls,
  chartWaterfallStylePointIsActive, chartWaterfallStylePoints, chartWaterfallStyleGroup, chartWaterfallStyleControls,
  chartWaterfallPointControls, chartWaterfallConnectorControls } from './features/charts/chart-waterfall-settings-controls.ts'
export { ChartAxisTargets, ChartAxisTitleAlignments, ChartAxisPointerTargets, ChartSeriesClearTargets, ChartAllSeriesStyleTargets,
  chartAxisPointerValues, clearChartAxisPointer,
  type ChartAxisTarget, type ChartAxisPointerTarget, type ChartSeriesClearTarget, type ChartAllSeriesStyleTarget } from './features/charts/chart-cartesian-api.ts'
export { chartDateSerial, prepareChartDateData, supportsChartDateAxis } from './features/charts/chart-date-axis.ts'
export { createSparklineGroupPatches, createSparklineSpec, createSparklinePlacementPatches, normalizeSparklineSpec, createSparklineGroupId,
  translateSparklineForMove, resolveSparklineSettings, mergeSparklineSettings, sparklineSettingsInput, readSparklineGroups,
  planSparklineMutation, assertSparklineGroupConsistency, SPARKLINE_POINT_ORDER, SPARKLINE_THEMES } from './features/sparkline/sparkline-model.ts'
export type { SparklineGroupRequest, SparklineSettingsInput, ResolvedSparklineSettings, SparklinePointName, SparklineTheme,
  SparklineExtremumMode, SparklinePlacementRequest, SparklineGroupInfo, SparklineMutationRequest, SparklineMutationPlan,
  SparklineMutationOptions } from './features/sparkline/sparkline-model.ts'
export { SET_SPARKLINES_COMMAND, normalizeSparklinePatches, assertSparklineReadPermission, assertSparklineMutationPermission,
  applySparklinePatches } from './features/sparkline/sparkline-command.ts'
export type { SparklineMutationContext, SparklineCommandParams } from './features/sparkline/sparkline-command.ts'
export { FSparkline, FSparklineGroup } from './api/sparklines.ts'
export { SparklineTypeEnum, SparklineThemeTypeEnum, SpacialShowAsEnum, SparklineExtremumTypeEnum,
  sparklineTypeFromPublic, sparklineConfigFromPublic, sparklineConfigToPublic } from './features/sparkline/sparkline-api-types.ts'
export type { ISparklinePointConfig, ISparklineAxisConfig, ISparklineExtremumConfig, ISparklinePoints,
  ISparklineGroupConfig, ISparklineSourceRange, ISparklineGroup, ISparklineSerializedGroup } from './features/sparkline/sparkline-api-types.ts'
export { SparklineRangeMatrix } from './features/sparkline/sparkline-range-matrix.ts'
export type { SparklineMatrixRange, SparklineMatrixData } from './features/sparkline/sparkline-range-matrix.ts'
export { SPARKLINE_PUBLIC_COMMANDS } from './api/sparkline-public-commands.ts'
export type { AddSheetSparklineCommandInput, RemoveSheetSparklineCommandInput, SetSheetSparklineCommandInput } from './api/sparkline-public-commands.ts'
export { SHEET_SPARKLINE_PLUGIN, serializeWorkbookSparklineResource, deserializeWorkbookSparklineResource,
  importWorkbookSparklineResources } from './features/sparkline/sparkline-resources.ts'
export type { SparklineResource, SparklineResourceGroup } from './features/sparkline/sparkline-resources.ts'
export { createSparklineSourceChangeDetector } from './features/sparkline/sparkline-source-changes.ts'
export { createSparklineCommands, parseSparklineRange, readSparklineValues, readSparklineSamples, sparklineCommandsForCell } from './features/sparkline/sparkline-render.ts'
export type { SparklineRenderContext, SparklineExtrema, SparklineSample } from './features/sparkline/sparkline-render.ts'
export { parseSparklineReference, parseSparklineRangeList, sparklineCellAddress, sparklineRangeString,
  sparklineSheetQualifier, shiftSparklineReference } from './features/sparkline/sparkline-reference.ts'
export type { SparklineAddress, SparklineReference } from './features/sparkline/sparkline-reference.ts'
export { normalizeSparklineSourceReference, parseSparklineSource, normalizeSparklineSourceText, resolveSparklineSource,
  sparklineSourceToPublicRange, shiftSparklineSource } from './features/sparkline/sparkline-source.ts'
export type { SparklineSourceSheet, SparklineSourceWorkbook, SparklineSource, SparklineSourceInput,
  SparklineSourceResolution, SparklinePublicSourceRange } from './features/sparkline/sparkline-source.ts'
export { default as SparklinePanel } from './components/spreadsheet/SparklinePanel.vue'

export { DimensionOutlineAxis, DimensionOutlineErrorReason, SheetOutlineError, MAX_SHEET_OUTLINE_DEPTH,
  SHEET_OUTLINE_PLUGIN, sheetOutlineErrorMessage } from './features/sheet-outline/sheet-outline-types.ts'
export type { IDimensionOutline, SheetOutlineExtent, SheetOutlineScope, SheetOutlineHiddenRange, SheetOutlineTreeNode,
  SheetDimensionOutlineResource, SheetsOutlineResource } from './features/sheet-outline/sheet-outline-types.ts'
export { CHANGE_SHEET_OUTLINE_COMMAND, createSheetOutlineState, addSheetOutlineGroup, removeSheetOutlineGroup,
  setSheetOutlineCollapsed, setSheetOutlineSummaryPosition, setSheetOutlineLevel, getSheetOutlineUpdate,
  getSheetOutlineHiddenSpans, validateSheetOutlineState, buildSheetOutlineTree } from './features/sheet-outline/sheet-outline.ts'
export type { OutlineAxisState, SheetOutlineOperation, OutlineRange, SheetOutlineAction, ChangeSheetOutlineInput,
  OutlineWorksheet } from './features/sheet-outline/sheet-outline.ts'
export { getSheetDimensionOutlines, serializeSheetOutlineResource, deserializeSheetOutlineResource,
  serializeWorkbookOutlineResource, deserializeWorkbookOutlineResource } from './features/sheet-outline/sheet-outline-resources.ts'
export { transformSheetOutlineForAxisEdit } from './features/sheet-outline/sheet-outline-structure.ts'
export { getSheetOutlineViewModels, getSheetOutlineLevelControls } from './features/sheet-outline/sheet-outline-view.ts'
export type { SheetOutlineViewModel } from './features/sheet-outline/sheet-outline-view.ts'
export { SHEET_OUTLINE_CORE_COMMAND_IDS, SHEET_OUTLINE_UI_COMMAND_IDS, SHEET_OUTLINE_LOCAL_COMMAND_IDS,
  SHEET_OUTLINE_COMMAND_IDS, resolveSheetOutlineCommand, getSheetOutlinePromptChoices } from './features/sheet-outline/sheet-outline-command.ts'
export type { SheetOutlineCommandInput, SheetOutlineCommandContext, SheetOutlinePromptChoice,
  SheetOutlineCommandResolution } from './features/sheet-outline/sheet-outline-command.ts'

