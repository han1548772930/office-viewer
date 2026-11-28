import type { OfficeDirection } from '../i18n/regional-settings.ts'
import type { DocumentCharacterPreferencesStore } from '../features/document/stores/document-character-preferences-store'
import type {
  CellScalar,
  ChartAxisOptions,
  ChartAxisPointerOptions,
  ChartModel,
  ChartImageResource,
  ChartPlotAreaOptions,
  ChartParetoBarStyleOptions,
  ChartParetoLineStyleOptions,
  ChartSemanticOptions,
  ChartAreaLineStyle,
  ChartInvalidValueStrategy,
  ChartSeriesOptions,
  ChartSeriesSelector,
  ChartWaterfallConnectorOptions,
  ChartWaterfallSeriesStyle,
  ChartWaterfallStackType,
  ChartBubbleMappingOptions,
  ChartCandlestickCloseMarkerOptions,
  ChartCandlestickHighLowLineOptions,
  ChartCandlestickMappingOptions,
  ChartCandlestickStateStyleOptions,
  ChartTreemapOptions,
  ChartTreemapLabelOptions,
  ChartSunburstLabelOptions,
  ChartSunburstCenterLabelOptions,
  ChartChordMapping,
  ChartChordLayoutOptions,
  ChartChordNodeOptions,
  ChartChordLabelOptions,
  ChartChordRibbonOptions,
  ChartChordEmphasisOptions,
  ChartTrendlineOptions,
  ChartErrorBarOptions,
  ChartGaugeOptions,
  ChartGaugeMode,
  ChartGaugeScaleOptions,
  ChartGaugeAnglesOptions,
  ChartGaugeRangeOptions,
  ChartGaugeAxisLineOptions,
  ChartGaugeAxisTickOptions,
  ChartGaugeSplitLineOptions,
  ChartGaugeAxisLabelOptions,
  ChartGaugePointerOptions,
  ChartGaugeAnchorOptions,
  ChartGaugeProgressOptions,
  ChartGaugeTitleOptions,
  ChartGaugeDetailOptions,
  ChartGaugeNullablePatch,
  ChartGaugeTicksPatch,
  ChartNullablePatch,
  ChartRadarOptions,
  ChartRadarShape,
  ChartWordCloudShape,
  ChartPointShape,
  ChartRelationLayout,
  ChartRelationForceOptions,
  ChartRadarAxisLabelOptions,
  ChartSource,
  ChartSourceOrientation,
  ChartType,
  CellRange,
  DocumentContentControl,
  DocumentContentControlType,
  DocumentCustomXmlPart,
  DocumentFootnoteSettings,
  DocumentEndnoteSettings,
  DocumentModel,
  DocumentParagraphStyle,
  PivotAggregation,
  PivotChartControls,
  PivotDateFilter,
  PivotDateGroupType,
  PivotDateGroupingMode,
  PivotFieldArea,
  PivotFieldFilter,
  PivotFieldModel,
  PivotLabelConditionFilter,
  PivotManualGroup,
  PivotNumberGroup,
  PivotRecordSelection,
  PivotShowDataAsInfo,
  PivotTableModel,
  PivotTableOptions,
  SlicerModel,
  SheetSlicerConnections,
  SlicerConnectionCandidate,
  SlicerTimelineRange,
  SpreadsheetCell,
  SpreadsheetSheet,
  SparklineSpec,
  SheetDataValidationRule,
  SheetConditionalFormatRule,
  SheetRangeProtectionDraft,
  SheetRangeProtectionRule,
  SheetFilterCriteria,
  SheetSortCriterion,
  WorkbookModel,
} from '../domain/model'
import { createCell } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { worksheetHiddenState, type WorksheetHiddenState } from '../features/sheet-workbench/sheet-worksheet-lifecycle.ts'
import { SHEET_AUTO_FILL_COMMAND } from '../features/sheet-workbench/sheet-auto-fill.ts'
import { resolveSheetFillTarget } from '../features/sheet-workbench/sheet-fill-target.ts'
import { sheetFillApplyType } from '../features/sheet-workbench/sheet-fill-command.ts'
import type { SheetFillApplyType } from '../features/sheet-workbench/sheet-fill.ts'
import { normalizeDocumentCustomXmlParts } from '../domain/document-custom-xml'
import { SET_DOCUMENT_CUSTOM_XML_PARTS_COMMAND, type DocumentCustomXmlPartsWriteOptions } from './document-custom-xml-commands'
import { assertDocumentPermission, permissionTargetsForSelection } from '../features/document/permissions/document-permissions'
import { cloneChartOptions, mergeChartOptions } from '../features/charts/chart-options'
import { ChartAreaLineStyles, ChartInvalidValueTypes, CHART_INVALID_VALUE_STRATEGY_DEFAULT,
  chartAreaLineStylePatch, chartAreaLineStyleValue, chartInvalidValueStrategyValue, resolveChartAreaLineStyle } from '../features/charts/chart-line-and-area'
import { chartTypeSelectOptions, resolveChartType } from '../features/charts/chart-catalog'
import { convertChartType } from '../features/charts/chart-type-conversion'
import { chartGaugeNullableValues, chartGaugeRangesOptions, chartGaugeTicksOptions, patchChartGaugeOptions } from '../features/charts/chart-gauge-api'
import { assertChartCandlestickType, chartCandleWidth, chartCandlestickFields, chartCandlestickStateStyleValues,
  chartCandlestickHighLowLineValues, chartCandlestickCloseMarkerValues } from '../features/charts/chart-candlestick-api'
import { assertChartHistogramType, isChartHistogramAuthoredNumber } from '../features/charts/chart-histogram-api'
import { assertChartTreemapType, chartTreemapHierarchyFields, chartTreemapValueField, chartTreemapParentLabelLayout,
  chartTreemapLabelValues } from '../features/charts/chart-treemap-api'
import { assertChartSunburstType, chartSunburstHierarchyFields, chartSunburstValueField, chartSunburstLabelValues,
  chartSunburstCenterLabelValues } from '../features/charts/chart-sunburst-api'
import { assertChartChordType, chartChordMappingValues, chartChordLayoutPatch, chartChordNodeValues,
  chartChordLabelValues, chartChordRibbonPatch, chartChordEmphasisValues } from '../features/charts/chart-chord-api'
import { ChartGaugeModes, ChartGaugePointerShapes, chartGaugePresetOptions, type ChartGaugePreset } from '../features/charts/chart-gauge'
import { ChartRadarShapes, patchChartRadarOptions } from '../features/charts/chart-radar'
import { ChartWordCloudShapes, ChartWordCloudShapeSources, assertChartWordCloudType, chartWordCloudShape,
  chartWordCloudMaskImageId, chartWordCloudRepeat } from '../features/charts/chart-word-cloud'
import { normalizeChartImageResource } from '../features/charts/chart-image-resources'
import { ChartRelationLayouts, assertChartRelationType, chartRelationLayoutValue, chartRelationBoolean,
  chartRelationForce, chartRelationNodeShape } from '../features/charts/chart-relation'
import { assertChartWaterfallType, chartWaterfallConnectorValues, patchChartWaterfallStyle } from '../features/charts/chart-waterfall'
import { assertChartParetoType, chartParetoBarStyleValues, chartParetoLineStyleValues } from '../features/charts/chart-pareto'
import { ChartVisualMapTypes, assertChartHeatmapType, chartHeatmapVisualMapType, chartHeatmapValueRange, chartHeatmapCellLabelValues, chartHeatmapValueUnit,
  type ChartHeatmapVisualMapType, type ChartHeatmapLabelOptions } from '../features/charts/chart-heatmap-api'
import { ChartAllSeriesStyleTargets, ChartAxisTargets, ChartAxisTitleAlignments, ChartAxisPointerTargets, ChartSeriesClearTargets, chartCartesianAxisPatch, chartAxisPointerValues, chartSeriesStyleValues, clearChartAllSeriesStyle, clearChartAxis, clearChartAxisPointer, clearChartSeriesStyle, patchChartSeriesOverride,
  type ChartAllSeriesStyleTarget, type ChartAxisTarget, type ChartAxisPointerTarget, type ChartSeriesClearTarget } from '../features/charts/chart-cartesian-api'
import { chartSeriesOverride, validateChartSeriesSelector } from '../features/charts/chart-series-selector'
import { validateChartPlotArea } from '../features/charts/chart-plot-area'
import { validateChartGradientAuthoring, validateChartGradientSeriesPatch } from '../features/charts/chart-gradient'
import { describeChartConfiguration, type ChartDescription, type ChartDescriptionContext } from '../features/charts/chart-description'
import { chartDescriptionSourceContext } from './chart-description-host'
import type { ChartComputedStatistics } from '../features/charts/chart-derived'
import { assertChartReadPermission, chartStatisticsSignature, type ChartStatisticsRequestContext } from '../features/charts/services/chart-statistics-service'
import { ComputeError } from '../compute/errors'
import { SHEET_DATA_VALIDATION_ASYNC_COMMANDS } from '../features/sheet-validation/validation-rule-command.ts'
import type { SheetDataValidationAsyncOptions } from '../features/sheet-validation/validation-checkbox-rule-cells-async.ts'
import { CONTENT_HISTORY_COMMANDS, ContentHistoryCanceledError, type ContentHistoryDirection } from './content-history-commands.ts'
import {
  getCommonDocumentTextStyle,
  getDocumentParagraphStyle,
  getDocumentTextStyleRuns,
  documentBodyText,
  saveDocumentSnapshot,
  setDocumentTextStyle,
  type DocumentParagraphIdentity,
  type DocumentParagraphInfo,
  type DocumentTextStyle,
  type DocumentTextStylePatch,
  type DocumentTextStyleRun,
} from '../features/document/services/document-facade'
import {
  contentControlRange,
  contentControlText,
  getDocumentContentControl,
  getDocumentContentControls,
  type ContentControlInput,
  type ContentControlPatch,
} from '../features/document/services/content-controls'
import type { OfficeStudio } from './office-studio'
import { COMMIT_SHEET_CELL_INPUT_COMMAND, type SheetCellInputCommandResult,
  type SheetCellInputOptions } from '../features/sheet-workbench/sheet-cell-input-command.ts'
import { SHEET_AXIS_DIMENSION_COMMANDS } from '../features/sheet-workbench/sheet-axis-dimensions.ts'
import { SHEET_CELL_SHIFT_COMMANDS, SheetCellShiftDimensions, type SheetCellShiftDimension } from '../features/sheet-workbench/sheet-cell-shift-command.ts'
import type { SheetAxisStructureCommandInput } from '../features/sheet-workbench/sheet-axis-structure-command.ts'
import { SHEET_AXIS_VISIBILITY_COMMANDS } from '../features/sheet-workbench/sheet-axis-visibility.ts'
import { SET_SHEET_CONDITIONAL_FORMATS_COMMAND } from '../features/conditional-formatting/conditional-format-command.ts'
import { sheetRowHeight, sheetColumnWidth, sheetRowIsAutoHeight } from '../domain/sheet-axis.ts'
import type { DocumentCommand } from '../features/document/commands/document-commands.ts'
import { textTargets, type DocumentSelection, type RunPatch } from '../features/document/model/document-model.ts'
import type { DocumentStoryFieldRefreshRequest } from '../features/document/services/document-story-fields.ts'
import type { DocumentStoryRange } from '../features/document/model/document-story.ts'
import type { DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { FDocumentTable, getDocumentTables, findDocumentTables, getDocumentTableAtSelection, insertDocumentTable,
  DocsTableSelectionKinds, DocsTableRowInsertPositions, DocsTableColumnInsertPositions, DocsTableSortDirections,
  DocsTableColumnTypes, DocsTableRowHeightRules, DocsTableBorderPresets,
  type DocumentTableFindQuery } from './document-tables.ts'
import type { DocumentTableInsertOptions } from '../features/document/commands/table-command-types.ts'
import { FDocumentQuote, FDocumentCallout, FDocumentCode, documentStructuredBlocks, findDocumentStructuredBlocks, documentStructuredBlockAt,
  insertDocumentQuote, insertDocumentCallout, insertDocumentCode, type DocumentStructuredFindQuery,
  type DocumentQuoteInsertOptions, type DocumentCalloutInsertOptions, type DocumentCodeInsertOptions } from './document-structured-blocks'
import { FDocumentColumnGroup, insertDocumentColumnGroup, type DocumentColumnFindQuery, type DocumentColumnInsertOptions } from './document-columns'
import { FDocumentLatex, documentLatexFormulaTargets, documentLatexRangeAt, findDocumentLatexFormulas, insertDocumentLatexAtRange,
  insertDocumentLatexAtSelection, type DocumentLatexCreateOptions, type DocumentLatexInsertOptions, type DocumentLatexFindQuery } from './document-latex'
import { resolveDocumentSegmentHost } from './document-segments'
import { registerDocumentLatexTemplateEditor, type DocumentLatexTemplateEditorHandle } from './document-latex-template-editors'
import { registerDocumentLatexPopupEditor, type DocumentLatexPopupEditor, type DocumentLatexPopupEditorHandle } from './document-latex-popup-editors'
import { DOCUMENT_LATEX_POPUP_OPERATIONS, type DocumentLatexPopupOperationRequest } from '../features/document/controllers/document-latex-popup'
import { DOCUMENT_LATEX_TEMPLATE_OPERATION } from '../features/document/latex/latex-templates'
import { documentBlockHostId } from '../features/document/model/document-block-host'
import { parseSheetA1Range, splitSheetA1Reference } from '../features/sheet-workbench/sheet-a1-reference.ts'
import type { DocumentEditorCommandResult } from './document-editor-commands.ts'
import { DocumentNoteCommandRejectedError, type DocumentNoteCommandResult } from './document-note-commands'
import { assertDocumentPublicRange, documentPublicBodyIndex, documentPublicOffsetAtFlat, documentPublicParagraphs, documentPublicRangeToFlat,
  documentSelectionAtPublicBodyRange, type DocumentPublicBodyIndex } from './document-public-body'
import type { UniverDocumentBody, UniverDocumentBodyData } from './document-body-types'
import type { UniverDocumentData } from './document-snapshot-types'
import { documentToPublicSnapshot } from './document-output'
import { assertDocumentSnapshotViewPermission } from './document-snapshot-permissions'
import { commonPublicDocumentExplicitTextStyle, isLocalDocumentTextStyle, publicDocumentExplicitTextStyleRuns,
  setPublicDocumentTextStyle, validatePublicDocumentTextStyle, type UniverDocumentTextStyle, type UniverDocumentTextStyleRun,
  type UniverDocumentTextRangeDescription } from './document-text-style'
import type { DocumentNoteKind, DocumentNoteMutationRequest } from '../features/document/notes/document-note-types'
import { getDocumentNotesInReferenceOrder } from '../features/document/notes/document-note-references'
import { getDocumentNoteSettings } from '../features/document/notes/document-note-settings'
import { documentNotePermissionTargets } from '../features/document/permissions/document-note-permissions'
import { applyDocumentBlockHostModel, documentModelForBlockHost } from '../features/document/services/document-block-host'
import { documentNoteInputFromPublic, documentNoteSettingsFromPublic, documentNoteSettingsToPublic, documentNoteToPublic,
  type DocumentNoteContentInput, type DocumentNoteInfo, type DocumentNoteSettingsInput, type PublicDocumentFootnoteSettings,
  type PublicDocumentEndnoteSettings } from './document-note-body'
export type { DocumentNoteInfo, DocumentNoteContentInput, PublicDocumentNoteInput, PublicDocumentNoteDecoration,
  PublicDocumentFootnoteSettings, PublicDocumentEndnoteSettings } from './document-note-body'
import { FEventName, type ApiEventMap, type Disposable } from './events'
import { applyCellValue, randomId, type ApiCellValue, type WorkbookInput } from './model-utils'
import { SHEET_IMAGE_COMMANDS } from '../features/sheet-images/sheet-image-commands.ts'
import { sheetBackgroundImageFromPublic, sheetBackgroundImageToPublic, sheetCellDataWithPublicImage, workbookWithPublicSheetImages,
  SheetImageSourceTypes, type SheetImageSourceType, type UniverWorksheetBackgroundImage, type UniverWorkbookData, type UniverSheetCellData,
  type UniverSheetImageDrawing } from '../features/sheet-images/sheet-image-public.ts'
import { sheetImageSaveConfig, type SaveCellImagesOptions } from '../features/sheet-images/sheet-image-save.ts'
import type { SheetImageInputPreparationOptions } from '../features/sheet-images/sheet-image-input-preparation.ts'
import { FOverGridImage, FOverGridImageBuilder, canViewSheetImage, insertWorksheetImage, preparedSheetImages, type SheetImageAsyncOptions } from './sheet-images.ts'
import { SheetDrawingAnchorTypes, type SheetFloatingImage } from '../domain/sheet-images.ts'
import { SHEET_IMAGE_CROP_TYPES } from '../features/sheet-images/sheet-image-crop.ts'
import { SHEET_IMAGE_ALIGN_TYPES, SHEET_IMAGE_ALIGN_OPERATIONS, type SheetImageAlignType } from '../features/sheet-images/sheet-image-selection.ts'
import { FDocumentImage, insertDocumentImage, getDocumentImage, getDocumentImages, alignDocumentImages, transformDocumentImages,
  arrangeDocumentImages, removeDocumentImages, type DocumentImageInsertOptions } from './document-images'
import { DOCUMENT_IMAGE_ALIGN_TYPES, type DocumentImageAlign, type DocumentImageBatchTransform } from '../features/document/images/document-image-batch'
import type { DocumentImageGeometryReadOptions } from './document-image-layout'
import type { DocumentImageArrange } from '../features/document/images/document-image-model'
import { SHEET_DRAWING_ALIGN_TYPES, type SheetDrawingAlignType, type SheetDrawingRef, type SheetDrawingAffineGeometry, type SheetDrawingMatrix, type SheetDrawingMoveDirection } from '../domain/sheet-drawings.ts'
import { SHEET_DRAWING_GROUP_COMMAND_IDS, drawingPaintOrder, planDrawingArrange, type SheetDrawingArrangeType } from '../features/drawing-groups/index.ts'
import { createSheetDrawingScene } from '../features/sheet-drawings/sheet-drawing-host.ts'
import { registerDocumentImageEditor, type DocumentImageEditor, type DocumentImageEditorHandle } from './document-image-editors'
import { createSparklineGroupPatches, readSparklineGroups, type SparklineGroupInfo, type SparklineGroupRequest, type SparklineMutationRequest } from '../features/sparkline/sparkline-model.ts'
import { SET_SPARKLINES_COMMAND, assertSparklineReadPermission, type SparklineMutationContext } from '../features/sparkline/sparkline-command.ts'
import { createSparklineGroupId } from '../features/sparkline/sparkline-model.ts'
import { SparklineTypeEnum, SparklineThemeTypeEnum, SparklineExtremumTypeEnum, SpacialShowAsEnum, sparklineTypeFromPublic, type ISparklineGroup } from '../features/sparkline/sparkline-api-types.ts'
import { FSparkline, FSparklineGroup, sparklineApiRangeReference, sparklineGroupToPublic } from './sparklines.ts'
import { sparklineCellAddress } from '../features/sparkline/sparkline-reference.ts'
import { SHEET_SPARKLINE_PLUGIN, serializeWorkbookSparklineResource } from '../features/sparkline/sparkline-resources.ts'
import { getSheetDimensionOutlines, serializeWorkbookOutlineResource, sheetOutlinePublicAddRange, sheetOutlineScope } from '../features/sheet-outline/sheet-outline-resources.ts'
import { DimensionOutlineAxis, SHEET_OUTLINE_PLUGIN, isIgnoredOutlineAddError, type IDimensionOutline } from '../features/sheet-outline/sheet-outline-types.ts'
import { parseSheetHyperlinkTarget } from '../features/sheet-workbench/sheet-hyperlink.ts'
import { SHEET_NUMBER_FORMAT_COMMANDS, type SheetNumberFormatCellInput } from '../features/sheet-workbench/sheet-number-format-command.ts'
import { SHEET_FONT_FORMAT_COMMANDS } from '../features/sheet-workbench/sheet-font-format-command.ts'
import { SHEET_DEFAULT_STYLE_COMMANDS } from './sheet-default-style-commands.ts'
import { SHEET_RANGE_THEME_COMMANDS } from './sheet-range-theme-commands.ts'
import { RangeThemeStyle, sheetRangeThemeToJson, type SheetRangeThemeStyleJson } from '../features/sheet-workbench/sheet-range-theme-style.ts'
import { SHEET_BUILTIN_RANGE_THEME_NAMES } from '../features/sheet-workbench/sheet-range-theme-presets.ts'
import { sheetRangeThemeByName, sheetUsedRangeThemeRule } from '../features/sheet-workbench/sheet-range-theme-resolver.ts'
import { resolveSheetCellStyle, resolveSheetStyleSource, sheetAxisStyleSource, sheetCellDirectStyle, type SheetStyleInput } from '../features/sheet-workbench/sheet-style-inheritance.ts'
import type { CellStyle, SheetStyleSource } from '../domain/model.ts'
import { sheetStyleDataWithPublicFields, type SheetApiStyleData } from '../features/sheet-workbench/sheet-public-style.ts'
import { FORMAT_SHEET_CELL_RICH_TEXT_COMMAND, type SheetCellInlineStylePatch } from '../features/sheet-workbench/sheet-cell-inline-format-command.ts'
import { SHEET_BORDER_FORMAT_COMMANDS } from '../features/sheet-workbench/sheet-border-format-command.ts'
import { SheetBorderTypes, SheetBorderStyleTypes, type SheetBorderType, type SheetBorderPreset, type SheetBorderStyleType, type SheetBorderLineStyle } from '../features/sheet-workbench/sheet-cell-border.ts'
import { SHEET_LAYOUT_FORMAT_COMMANDS, sheetCellTextRotation, sheetCellWrapStrategy, type SheetTextRotation, type SheetWrapStrategy } from '../features/sheet-workbench/sheet-layout-format-command.ts'
import { SHEET_MERGE_COMMANDS, type SheetMergeMode, type SheetMergeOptions } from '../features/sheet-workbench/sheet-merge-command.ts'
import { SHEET_CLEAR_COMMANDS, type SheetClearOptions } from '../features/sheet-workbench/sheet-clear-command.ts'
import { sheetRangesEqual, sheetRangesIntersect } from '../features/sheet-workbench/sheet-range-command.ts'
import type { SheetSelectionOptions } from '../features/sheet-workbench/sheet-selection.ts'
import { nextSheetSelectionCell, SheetSelectionDirections, type SheetSelectionDirection } from '../features/sheet-workbench/sheet-selection-navigation.ts'
import { formatSheetValue } from '../features/sheet-workbench/sheet-number-format.ts'

export type {
  DocumentBodySegment,
  DocumentBodySnapshot,
  DocumentParagraphIdentity,
  DocumentParagraphInfo,
  DocumentParagraphQuery,
  DocumentTextRange,
  DocumentTextStyle,
  DocumentTextStylePatch,
  DocumentTextStyleRun,
} from '../features/document/services/document-facade'
import {
  getPivotSourceFields,
  getPivotShowDataAs,
  getPivotShowDataAsBaseItems,
  getPivotAxisState,
  getPivotCollapse,
  getPivotCollapseItems,
  getPivotDateFilter,
  getPivotLabelConditionFilter,
  getPivotManualFilterItemState,
  getPivotFieldFilterCapabilities,
  getPivotTableAtCell,
  pivotDrillDown,
  pivotFieldIds,
  pivotOutputValues,
  type PivotTableAnchorInput,
  type PivotTableSourceInput,
} from '../features/pivot/services/pivot-engine'
import { getPivotChartReport, getPivotChartConfiguredField, getPivotChartFieldItems, getPivotChartShowDataAsBaseItems, getPivotChartAxisState, getPivotChartCollapseItems, type PivotChartFieldsConfig, type PivotChartFieldSetting,
  type PivotChartReportEdit } from '../features/pivot/services/pivot-chart-engine.ts'
import { PIVOT_CHART_COMMANDS, type PivotChartCommandContext, type PivotChartConfigureInput,
  type PivotChartInfo, type PivotChartInsertInfo, type PivotChartInsertOptions, type PivotChartSourceSpec, type PivotChartTableReference } from '../features/pivot/services/pivot-chart-api-types.ts'
import { assertPivotReportReadPermission } from '../features/pivot/services/pivot-report-permissions.ts'
import { PIVOT_HIERARCHY_COMMANDS } from '../features/pivot/services/pivot-hierarchy-api-types.ts'
import { PIVOT_CHART_FIELD_COMMANDS } from '../features/pivot/services/pivot-chart-field-command-types.ts'
import type { PivotDateFilterInput, PivotLabelConditionInput } from '../features/pivot/services/pivot-filter-api.ts'
import type { PivotReportCommandContext } from '../features/pivot/services/pivot-report-api-types.ts'
import { PIVOT_REPORT_FILTER_COMMANDS } from './pivot-filter-commands.ts'
import type { PivotAxisArea, PivotAxisState } from '../features/pivot/services/pivot-axis-state.ts'
import { pivotChartControls } from '../features/pivot/services/pivot-chart-controls.ts'
import {
  chartLayout,
  chartRenderCommands,
  createChartDefinition,
  defaultChartLegend,
  readChartData,
  type ChartData,
  type ChartConfig,
} from '../features/charts/chart-engine'
import { ChartLegendPositions, ChartLegendSelectModes, chartLegendPatch, type ChartLegendSpec } from '../features/charts/chart-legend'
import { ChartTitlePositions, ChartTextAlignments, chartTitlePatch, type ChartTitleSpec, type ChartSubtitleSpec } from '../features/charts/chart-titles'
import { getSlicer, getSlicerItems, readSlicerItemState, readSlicerRecordSelection, readSlicerSelection, type SlicerCreateInput } from '../features/slicer/slicer-engine'
import { RecordValueType, FilterSelectionMode, isRecordValue, normalizeRecordSelection, type RecordSelection } from '../domain/record-value.ts'
import { getPivotValuesPosition, normalizePivotValuesPosition, normalizePivotValuePositionInput, PIVOT_VALUE_POSITIONS } from '../features/pivot/services/pivot-values-position.ts'
import type { PivotValuePosition } from '../domain/model.ts'
import { getSlicerCache, getSlicerConnectionCandidates, getSlicerPivotSource, getSlicerSourceValues,
  getSlicerSourceRecordValues, getSlicerTableNativeFilter, SlicerConnectionError } from '../features/slicer/slicer-connections.ts'
import { slicerPublicRecordSelection } from '../features/slicer/slicer-record-selection.ts'
import { SlicerMemberQueryError, getSlicerDisplayMembers, type SlicerDisplayMember, type ISheetSlicerItemsRequest, type ISlicerSourceMemberWindow, type ISlicerSourceMetadata,
  type SlicerMemberQueryContext } from '../features/slicer/slicer-member-query.ts'
import { normalizeSlicerCacheOptions } from '../features/slicer/slicer-cache-options.ts'
import { readSlicerSettings, type SheetSlicerSettings } from '../features/slicer/slicer-settings.ts'
import { listSlicerStyleNames } from '../features/slicer/slicer-style-catalog.ts'
import { FSheetSlicerBuilder, FSheetTimelineBuilder, type FBoundSheetSlicerBuilder, type FBoundSheetTimelineBuilder } from './slicer-builders.ts'
import { readTimelineSettings, type SheetTimelineSettings, type TimelineSettingsInputPatch } from '../features/slicer/timeline-settings.ts'
import { readTimelineDateCondition, readTimelineDateRange } from '../features/slicer/timeline-filter.ts'
import { readSheetSlicerDrawing, readSheetSlicerInfo, readSheetSlicerSource, readSheetTimelineInfo, sheetSlicerCacheDefinition, sheetTimelineCacheDefinition, SLICER_PUBLIC_COMMANDS,
  type ISheetSlicerDrawing, type ISheetSlicerInfo, type ISheetTimelineInfo, type ITimelineCacheDefinition, type SheetSlicerAnchorSpec,
  type SheetSlicerPlacement, type SheetSlicerSourceSpec, type SlicerCacheDefinition, type TimelinePresentationPatch } from '../features/slicer/slicer-public-info.ts'
import type { ITimelineDateCondition, ITimelineRange, TimelinePublicLevel } from '../domain/timeline-filters.ts'
import { canViewProtectedRanges, canViewWorksheetProtection } from '../features/sheet-protection/range-protection.ts'
import { assertSlicerSourceReadPermission, assertSlicerViewPermission } from '../features/slicer/slicer-permissions.ts'
import { expandVisibleSortOrder, getFilteredRows, getNumericSortColumns, getSortedRowOrder, type SortRangeOptions } from '../features/sheet-sort-filter/sort-filter-engine'
import { createSheetDataValidationRule, sheetDataValidationOptions, validateSheetDataValidationChanges, type SheetDataValidationDraft, type SheetValidationIssue } from '../features/sheet-validation/validation-engine.ts'
import { sheetValidationSourceReader } from '../features/sheet-validation/validation-value-read.ts'
import { createSheetRangeProtectionRule } from '../features/sheet-protection/range-protection.ts'

export interface CreateUnitOptions { makeCurrent?: boolean }
/** Local asynchronous read options. Identity, revision and permission checks
 * come from the Facade host and cannot be replaced by caller options. */
export type ChartComputationReadOptions = Omit<ChartStatisticsRequestContext, 'documentId' | 'modelRevision' | 'userId' | 'assertCurrent'>
export interface FRangeCoordinates { startRow: number; startColumn: number; endRow: number; endColumn: number; rangeType?: 0 | 1 | 2 | 3 }
export interface FSelectionCell extends FRangeCoordinates { actualRow: number; actualColumn: number; isMerged: boolean; isMergedMainCell: boolean }
export type PivotTableSourceInfo = PivotTableSourceInput
export type PivotTableAnchorInfo = PivotTableAnchorInput
export type ChartInfo = ChartConfig
export type ChartPatch = Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>
export type SlicerInfo = Omit<SlicerCreateInput, 'id' | 'unitId' | 'hostSheetId'> & { id?: string }
export interface SortRangeApiOptions extends SortRangeOptions { signal?: AbortSignal; timeoutMs?: number; onProgress?: (progress: import('../compute/protocol.ts').ComputeProgress) => void }
export type DataValidationRuleInput = SheetDataValidationDraft
export type DataValidationRule = SheetDataValidationRule
export type DataValidationError = SheetValidationIssue
export type SortColumnSpec = { column: number; ascending: boolean } | number
export type DocumentContentControlInfo = DocumentContentControl
export type DocumentContentControlCreateInfo = Omit<ContentControlInput, 'id'> & { id?: string }
export type DocumentContentControlPatch = ContentControlPatch

/**
 * Public root facade for the headless OfficeStudio runtime.
 */
export class FOffice {
  readonly Event = new FEventName()
  readonly Enum = Object.freeze({ BorderType: SheetBorderTypes, BorderStyleTypes: SheetBorderStyleTypes, Direction: SheetSelectionDirections,
    SparklineTypeEnum, SparklineThemeTypeEnum, SparklineExtremumTypeEnum, SpacialShowAsEnum, DimensionOutlineAxis,
    RecordValueType: Object.freeze(RecordValueType), FilterSelectionMode: Object.freeze(FilterSelectionMode), PivotValuePosition: PivotValuePositions,
    SlicerCacheSortOrder: Object.freeze({ Ascending: 'ascending', Descending: 'descending' } as const),
    SlicerCacheCrossFilter: Object.freeze({ None: 'none', ShowItemsWithDataAtTop: 'showItemsWithDataAtTop', ShowItemsWithNoData: 'showItemsWithNoData' } as const),
    ImageSourceType: SheetImageSourceTypes, SheetDrawingAnchorType: Object.freeze(SheetDrawingAnchorTypes), CropType: Object.freeze(SHEET_IMAGE_CROP_TYPES),
    ArrangeType: Object.freeze({ forward: 0, backward: 1, front: 2, back: 3 } as const), AlignType: SHEET_IMAGE_ALIGN_TYPES,
    DocumentImageAlignType: DOCUMENT_IMAGE_ALIGN_TYPES, DrawingAlignType: SHEET_DRAWING_ALIGN_TYPES, DrawingType: Object.freeze({ UNRECOGNIZED: -1, DRAWING_IMAGE: 0, DRAWING_SHAPE: 1,
      DRAWING_CHART: 2, DRAWING_TABLE: 3, DRAWING_SMART_ART: 4, DRAWING_VIDEO: 5, DRAWING_GROUP: 6, DRAWING_UNIT: 7,
      DRAWING_DOM: 8, DRAWING_BLOCK: 9, DRAWING_SLICER: 10, DRAWING_TIMELINE: 11, DRAWING_OLE: 12 } as const),
    Dimension: SheetCellShiftDimensions,
    WorksheetHiddenState: Object.freeze({ VISIBLE: 0, HIDDEN: 1, VERY_HIDDEN: 2 } as const),
    AUTO_FILL_APPLY_TYPE: Object.freeze({ COPY: 'COPY', SERIES: 'SERIES', ONLY_FORMAT: 'ONLY_FORMAT', NO_FORMAT: 'NO_FORMAT' } as const),
    ChartGaugeMode: ChartGaugeModes, ChartGaugePointerShape: ChartGaugePointerShapes, ChartRadarShape: ChartRadarShapes,
    ChartWordCloudShapeEnum: ChartWordCloudShapes, ChartWordCloudShapeSource: ChartWordCloudShapeSources,
    ChartRelationChartLayoutEnum: ChartRelationLayouts,
    ChartVisualMapType: ChartVisualMapTypes,
    ChartSelectModeEnum: ChartLegendSelectModes,
    ChartAreaLineStyle: ChartAreaLineStyles, ChartInvalidValueType: ChartInvalidValueTypes,
    ChartLegendPositionEnum: ChartLegendPositions,
    ChartTitlePositionEnum: ChartTitlePositions,
    ChartLabelAlignEnum: ChartTextAlignments,
    ChartSunburstLabelPosition: Object.freeze({ Inside: 'inside', Outside: 'outside' } as const),
    ChartChordLabelPosition: Object.freeze({ Inside: 'inside', Outside: 'outside' } as const),
    ChartChordRibbonColorMode: Object.freeze({ Source: 'source', Target: 'target', Gradient: 'gradient' } as const),
    ChartChordFocusMode: Object.freeze({ Adjacency: 'adjacency', Self: 'self', None: 'none' } as const),
    ChartAxisTarget: ChartAxisTargets, ChartSeriesClearTarget: ChartSeriesClearTargets, ChartAllSeriesStyleTarget: ChartAllSeriesStyleTargets,
    ChartAxisPointerTarget: ChartAxisPointerTargets, AxisAlignEnum: ChartAxisTitleAlignments,
    DocsTableSelectionKind: DocsTableSelectionKinds, DocsTableRowInsertPosition: DocsTableRowInsertPositions,
    DocsTableColumnInsertPosition: DocsTableColumnInsertPositions, DocsTableSortDirection: DocsTableSortDirections,
    DocsTableColumnType: DocsTableColumnTypes, DocsTableRowHeightRule: DocsTableRowHeightRules,
    DocsTableBorderPreset: DocsTableBorderPresets })
  readonly studio: OfficeStudio

  private constructor(studio: OfficeStudio) {
    this.studio = studio
  }

  static newAPI(studio: OfficeStudio): FOffice {
    return new FOffice(studio)
  }

  setLocale(locale: string): void { this.studio.regionalSettings.setLocale(locale) }
  setRegion(region: string): void { this.studio.regionalSettings.setRegion(region) }
  setDirection(direction: OfficeDirection): void { this.studio.regionalSettings.setDirection(direction) }
  getCurrentLocale(): string { return this.studio.regionalSettings.locale.value.replace('-', '') }
  getCurrentRegion(): string { return this.studio.regionalSettings.region.value.replace('-', '') }

  createWorkbook(data: WorkbookInput = {}, options: CreateUnitOptions = {}): FWorkbook {
    const model = this.studio.createWorkbook(data, options.makeCurrent ?? true)
    return new FWorkbook(this.studio, model.id!)
  }
  async createWorkbookAsync(data: WorkbookInput = {}, options: CreateUnitOptions & SheetImageInputPreparationOptions = {}): Promise<FWorkbook> {
    const model = await this.studio.createWorkbookAsync(data, options.makeCurrent ?? true, options)
    return new FWorkbook(this.studio, model.id!)
  }

  getActiveWorkbook(): FWorkbook | null {
    return this.getWorkbook(this.studio.activeWorkbookId)
  }

  getWorkbook(id: string): FWorkbook | null {
    return this.studio.getWorkbookModel(id) ? new FWorkbook(this.studio, id) : null
  }

  createDocument(data: Partial<DocumentModel> & Record<string, unknown> = {}, options: CreateUnitOptions = {}): FDocument {
    const model = this.studio.createDocument(data, options.makeCurrent ?? true)
    return new FDocument(this.studio, model.id!)
  }

  getActiveDocument(): FDocument | null {
    return this.getDocument(this.studio.activeDocumentId)
  }

  getDocument(id: string): FDocument | null {
    return this.studio.getDocumentModel(id) ? new FDocument(this.studio, id) : null
  }

  getActiveSheet(): { workbook: FWorkbook; worksheet: FWorksheet } | null {
    const workbook = this.getActiveWorkbook()
    const worksheet = workbook?.getActiveSheet() ?? null
    return workbook && worksheet ? { workbook, worksheet } : null
  }

  getSheetCommandTarget(params: { unitId?: string; subUnitId?: string; sheetId?: string } = {}): { workbook: FWorkbook; worksheet: FWorksheet; unitId: string; subUnitId: string } | null {
    const workbook = params.unitId ? this.getWorkbook(params.unitId) : this.getActiveWorkbook()
    if (!workbook) return null
    const sheet = params.subUnitId || params.sheetId
    const worksheet = sheet ? workbook.getSheetBySheetId(sheet) : workbook.getActiveSheet()
    return worksheet ? { workbook, worksheet, unitId: workbook.getId(), subUnitId: worksheet.getSheetId() } : null
  }

  getPivotTableByCell(unitId: string, subUnitId: string, row: number, col: number): FPivotTable | null {
    const workbook = this.getWorkbook(unitId)
    return workbook?.getPivotTableByCell(subUnitId, row, col) ?? null
  }

  disposeUnit(unitId: string): boolean { return this.studio.disposeUnit(unitId) }
  addEvent<K extends keyof ApiEventMap>(event: K, listener: (payload: ApiEventMap[K]) => void): Disposable { return this.studio.events.addEvent(event, listener) }
  registerCommand<P, R>(id: string, handler: (params: P) => R): Disposable { return this.studio.registerCommand(id, handler) }
  executeCommand<P = Record<string, unknown>, R = boolean>(id: string, params?: P): Promise<R> { return this.studio.executeCommand(id, params) }
  executeCommandSync<P, R>(id: string, params: P): R { return this.studio.executeCommandSync(id, params) }
  undo(): Promise<boolean> { return this.replayFocusedHistory('undo') }
  redo(): Promise<boolean> { return this.replayFocusedHistory('redo') }
  private replayFocusedHistory(direction: ContentHistoryDirection): Promise<boolean> {
    return this.studio.executeCommand(CONTENT_HISTORY_COMMANDS[direction]).catch((error: unknown) => {
      if (error instanceof ContentHistoryCanceledError) return false
      throw error
    })
  }
  setActiveKind(kind: 'workbook' | 'document'): void { this.studio.setActiveKind(kind) }
  getCurrentLifecycleStage(): 'ready' | 'disposed' { return this.studio.isDisposed ? 'disposed' : 'ready' }
  dispose(): void { this.studio.dispose() }
}

export class FWorkbook {
  private readonly api: OfficeStudio
  readonly id: string

  constructor(api: OfficeStudio, id: string) {
    this.api = api
    this.id = id
  }

  getId(): string { return this.id }
  getName(): string { return this.model().name }
  getDefaultStyle(): SheetStyleSource { return cloneModel(this.model().defaultStyle ?? null) }
  getRegisteredRangeThemes(): string[] { return [...SHEET_BUILTIN_RANGE_THEME_NAMES] }
  getCustomRangeThemes(): string[] { return Object.keys(this.model().rangeThemes ?? {}) }
  createRangeThemeStyle(name: string, options?: Omit<SheetRangeThemeStyleJson, 'name'>): RangeThemeStyle { return new RangeThemeStyle(name, options) }
  getRangeThemeStyle(name: string): RangeThemeStyle | null {
    const style = sheetRangeThemeByName(name, this.model())
    return style ? new RangeThemeStyle(style.name, sheetRangeThemeToJson(style)) : null
  }
  registerRangeTheme(style: RangeThemeStyle): void {
    if (!(style instanceof RangeThemeStyle)) throw new TypeError('请提供 RangeThemeStyle 配置。')
    this.api.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.register, { unitId: this.id, rangeThemeStyleJson: style.toJson() })
  }
  unregisterRangeTheme(name: string): void { this.api.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.unregister, { unitId: this.id, themeName: name }) }
  setDefaultStyle(style: SheetStyleInput): this {
    this.api.executeCommandSync(SHEET_DEFAULT_STYLE_COMMANDS.workbook, { unitId: this.id, style })
    return this
  }
  getCalculationMode(): 'automatic' | 'manual' { return this.model().calculationMode ?? 'automatic' }
  setCalculationMode(mode: 'automatic' | 'manual'): this {
    this.api.executeCommandSync('sheet.command.set-calculation-mode', { unitId: this.id, mode })
    return this
  }
  recalculate(): this {
    this.api.executeCommandSync('sheet.command.recalculate-workbook', { unitId: this.id })
    return this
  }
  setName(name: string): this {
    this.api.executeCommandSync('sheet.command.set-workbook-name', { unitId: this.id, name })
    return this
  }
  save(): WorkbookModel {
    for (const sheet of this.model().sheets) this.api.assertCanViewWorksheet(this.id, sheet.id)
    const snapshot = cloneModel(this.model())
    snapshot.modified = false
    this.api.markSaved('workbook', this.id)
    return snapshot
  }
  getSnapshot(): UniverWorkbookData {
    for (const sheet of this.model().sheets) this.api.assertCanViewWorksheet(this.id, sheet.id)
    const workbook = this.model(), snapshot = workbookWithPublicSheetImages(workbook)
    for (const sheet of workbook.sheets) assertSparklineReadPermission(sheet, readSparklineGroups(sheet).flatMap(group => group.targetRanges), this.api.currentUserId)
    for (const sheet of workbook.sheets) new FWorksheet(this.api, this.id, sheet.id).getDimensionOutlines()
    snapshot.resources = [...(snapshot.resources ?? []).filter(resource => resource.name !== SHEET_SPARKLINE_PLUGIN && resource.name !== SHEET_OUTLINE_PLUGIN),
      { name: SHEET_SPARKLINE_PLUGIN, data: JSON.stringify(serializeWorkbookSparklineResource(workbook)) },
      { name: SHEET_OUTLINE_PLUGIN, data: JSON.stringify(serializeWorkbookOutlineResource(workbook)) }]
    return snapshot
  }
  savePublicSnapshot(): UniverWorkbookData {
    const snapshot = this.getSnapshot()
    snapshot.modified = false
    this.api.markSaved('workbook', this.id)
    return snapshot
  }
  undo(): this { this.api.undo('workbook', this.id); return this }
  redo(): this { this.api.redo('workbook', this.id); return this }
  getActiveSheet(): FWorksheet | null {
    const sheet = this.model().sheets.find((item) => item.id === this.model().activeSheetId)
    return sheet ? new FWorksheet(this.api, this.id, sheet.id) : null
  }
  getActiveRange(): FRange | null { return this.getActiveSheet()?.getActiveRange() ?? null }
  getActiveRangeList(): FRange[] { return this.getActiveSheet()?.getActiveRangeList() ?? [] }
  getActiveCell(): FRange | null { return this.getActiveSheet()?.getActiveCell() ?? null }
  getSheets(): FWorksheet[] { return this.model().sheets.map((sheet) => new FWorksheet(this.api, this.id, sheet.id)) }
  getSheetBySheetId(sheetId: string): FWorksheet | null {
    return this.model().sheets.some((sheet) => sheet.id === sheetId) ? new FWorksheet(this.api, this.id, sheetId) : null
  }
  getSheetByName(name: string): FWorksheet | null {
    const sheet = this.model().sheets.find((item) => item.name === name)
    return sheet ? new FWorksheet(this.api, this.id, sheet.id) : null
  }
  setActiveSheet(sheet: FWorksheet | string): FWorksheet {
    const id = typeof sheet === 'string' ? sheet : sheet.getSheetId()
    if (!this.model().sheets.some((item) => item.id === id)) throw new Error(`Worksheet not found: ${id}`)
    this.api.assertCanViewWorksheet(this.id, id)
    this.api.mutateWorkbook(this.id, 'sheet.operation.set-active-sheet', 'Set active worksheet', (model) => { model.activeSheetId = id })
    this.api.activateWorkbook(this.id)
    return new FWorksheet(this.api, this.id, id)
  }
  navigateToSheetHyperlink(hyperlink: string): this {
    const target = parseSheetHyperlinkTarget(hyperlink, this.model())
    if (!target || target.type !== 'internal') throw new Error('Invalid internal worksheet hyperlink')
    const sheet = this.model().sheets.find((candidate) => candidate.id === target.sheetId)
    if (!sheet || sheet.visible === false) throw new Error('Worksheet hyperlink target is unavailable')
    this.setActiveSheet(sheet.id)
    this.api.setSelection(this.id, sheet.id, {
      startRow: target.range.startRow,
      startColumn: target.range.startCol,
      endRow: target.range.endRow,
      endColumn: target.range.endCol,
    })
    return this
  }
  insertSheet(sheetName?: string, options: { index?: number; sheet?: Partial<SpreadsheetSheet> } = {}): FWorksheet {
    const sheet = { ...(options.sheet ?? {}), id: options.sheet?.id ?? randomId('sheet') }
    const changed = this.api.executeCommandSync('sheet.command.insert-sheet', { unitId: this.id, name: sheetName, index: options.index, sheet })
    if (!changed || !this.api.getWorkbookModel(this.id)?.sheets.some(item => item.id === sheet.id)) throw new Error('工作表插入未完成。')
    return new FWorksheet(this.api, this.id, sheet.id)
  }
  deleteSheet(sheet: FWorksheet | string): boolean {
    const id = typeof sheet === 'string' ? sheet : sheet.getSheetId()
    if (typeof sheet !== 'string' && sheet.getWorkbook().getId() !== this.id) throw new Error('工作表必须属于当前工作簿。')
    if (!this.model().sheets.some(item => item.id === id)) return false
    return this.api.executeCommandSync('sheet.command.remove-sheet', { unitId: this.id, subUnitId: id })
  }
  getNumSheets(): number { return this.model().sheets.length }

  duplicateSheet(sheet: FWorksheet | string): FWorksheet {
    const sourceSheetId = typeof sheet === 'string' ? sheet : sheet.getSheetId()
    if (typeof sheet !== 'string' && sheet.getWorkbook().getId() !== this.id) throw new Error('工作表必须属于当前工作簿。')
    const targetSheetId = randomId('sheet')
    const changed = this.api.executeCommandSync('sheet.command.copy-sheet', { unitId: this.id, subUnitId: sourceSheetId, targetSheetId })
    if (!changed || !this.api.getWorkbookModel(this.id)?.sheets.some(item => item.id === targetSheetId)) throw new Error('工作表复制未完成。')
    return new FWorksheet(this.api, this.id, targetSheetId)
  }
  duplicateActiveSheet(): FWorksheet {
    const sheet = this.getActiveSheet()
    if (!sheet) throw new Error('当前没有活动工作表。')
    return this.duplicateSheet(sheet)
  }
  copySheet(sheet: FWorksheet | string): FWorksheet { return this.duplicateSheet(sheet) }

  async addPivotTable(
    sourceInfo: PivotTableSourceInfo,
    positionType: PositionTypeEnum | string = PositionTypeEnum.Existing,
    anchorCellInfo: PivotTableAnchorInfo,
  ): Promise<FPivotTable> {
    if (!sourceInfo || !anchorCellInfo) throw new TypeError('Pivot source and anchor are required')
    if (positionType !== PositionTypeEnum.Existing) throw new Error(`Unsupported pivot table position: ${positionType}`)
    const pivotTableId = randomId('pivot')
    this.api.executeCommandSync('sheet.command.add-pivot-table', {
      unitId: this.id,
      pivotTableId,
      source: normalizePivotSourceInfo(sourceInfo, this.id),
      anchor: normalizePivotAnchorInfo(anchorCellInfo, this.id),
    })
    return new FPivotTable(this.api, this.id, pivotTableId)
  }

  getPivotTables(): FPivotTable[] {
    return (this.model().pivotTables ?? []).map((pivot) => new FPivotTable(this.api, this.id, pivot.id))
  }

  getPivotChartById(pivotChartId: string): FSheetPivotChart | null {
    const chart = this.model().charts?.find(candidate => candidate.id === pivotChartId && (candidate.ownedPivot || candidate.pivotTableId))
    return chart ? new FSheetPivotChart(this.api, this.id, chart.id) : null
  }

  getPivotCharts(): FSheetPivotChart[] {
    return (this.model().charts ?? []).filter(chart => chart.ownedPivot || chart.pivotTableId)
      .map(chart => new FSheetPivotChart(this.api, this.id, chart.id))
  }

  getSlicers(): FSheetSlicer[] {
    return this.getSheets().flatMap(sheet => sheet.getSlicers())
  }

  getSlicer(slicerId: string): FSheetSlicer | null {
    return this.model().slicers?.some(view => view.id === slicerId && view.type === 'slicer') ? new FSheetSlicer(this.api, this.id, slicerId) : null
  }
  getSlicerById(slicerId: string): FSheetSlicer | null { return this.getSlicer(slicerId) }
  getTimelines(): FSheetTimeline[] { return this.getSheets().flatMap(sheet => sheet.getTimelines()) }
  getTimeline(timelineId: string): FSheetTimeline | null {
    return this.model().slicers?.some(view => view.id === timelineId && view.type === 'timeline') ? new FSheetTimeline(this.api, this.id, timelineId) : null
  }
  getTimelineById(timelineId: string): FSheetTimeline | null { return this.getTimeline(timelineId) }
  getSlicerStyleNames(): string[] { return listSlicerStyleNames(this.model(), 'slicer') }
  getTimelineStyleNames(): string[] { return listSlicerStyleNames(this.model(), 'timeline') }
  /** Local mixed drawing API for Canvas and property panels. */
  getSlicerControls(): FSheetSlicer[] { return this.getSheets().flatMap(sheet => sheet.getSlicerControls()) }
  getSlicerControl(controlId: string): FSheetSlicer | null {
    return this.model().slicers?.some(view => view.id === controlId) ? new FSheetSlicer(this.api, this.id, controlId) : null
  }

  getPivotTableByCell(subUnitId: string, row: number, col: number): FPivotTable | null
  getPivotTableByCell(unitId: string, subUnitId: string, row: number, col: number): FPivotTable | null
  getPivotTableByCell(unitIdOrSheetId: string, subUnitIdOrRow: string | number, rowOrCol: number, col?: number): FPivotTable | null {
    const sheetId = col == null ? unitIdOrSheetId : String(subUnitIdOrRow)
    const row = col == null ? Number(subUnitIdOrRow) : rowOrCol
    const column = col ?? rowOrCol
    if (col != null && unitIdOrSheetId !== this.id) return null
    const pivot = getPivotTableAtCell(this.model(), sheetId, row, column)
    return pivot ? new FPivotTable(this.api, this.id, pivot.id) : null
  }

  private model(): WorkbookModel {
    const model = this.api.getWorkbookModel(this.id)
    if (!model) throw new Error(`Workbook not found: ${this.id}`)
    return model
  }
}

export class FWorksheet {
  private readonly api: OfficeStudio
  readonly unitId: string
  readonly id: string

  constructor(api: OfficeStudio, unitId: string, id: string) {
    this.api = api
    this.unitId = unitId
    this.id = id
  }

  getUnitId(): string { return this.unitId }
  getSheetId(): string { return this.id }
  getSheetName(): string { return this.sheet().name }
  getBackgroundImage(): UniverWorksheetBackgroundImage | null {
    this.api.assertCanViewWorksheet(this.unitId, this.id)
    const image = this.sheet().backgroundImage
    return image ? sheetBackgroundImageToPublic(image) : null
  }
  setBackgroundImage(source: string, imageSourceType: SheetImageSourceType = 'URL'): this {
    const backgroundImage = sheetBackgroundImageFromPublic({ source, imageSourceType })
    this.api.executeCommandSync(SHEET_IMAGE_COMMANDS.setBackground, { unitId: this.unitId, sheetId: this.id, backgroundImage })
    return this
  }
  async setBackgroundImageAsync(source: string | File, options: SheetImageAsyncOptions = {}): Promise<boolean> {
    return this.api.executeCommand(SHEET_IMAGE_COMMANDS.addBackground, { ...options, unitId: this.unitId, sheetId: this.id,
      ...(typeof source === 'string' ? { source } : { files: [source] }) })
  }
  deleteBackgroundImage(): this { this.api.executeCommandSync(SHEET_IMAGE_COMMANDS.deleteBackground, { unitId: this.unitId, sheetId: this.id }); return this }
  insertImage(source: string | File, column = 0, row = 0, offsetX = 0, offsetY = 0, options: SheetImageAsyncOptions = {}): Promise<boolean> {
    return insertWorksheetImage(this.api, this.unitId, this.id, source, column, row, offsetX, offsetY, options)
  }
  newOverGridImage(): FOverGridImageBuilder { this.sheet(); return new FOverGridImageBuilder(this.api, this.unitId, this.id) }
  insertImages(images: readonly (SheetFloatingImage | UniverSheetImageDrawing)[]): this {
    this.api.executeCommandSync(SHEET_IMAGE_COMMANDS.insertPrepared, { unitId: this.unitId, sheetId: this.id, images: preparedSheetImages(this.api, this.unitId, this.id, images) })
    return this
  }
  updateImages(images: readonly (SheetFloatingImage | UniverSheetImageDrawing)[]): this {
    this.api.executeCommandSync(SHEET_IMAGE_COMMANDS.updateBatch, { unitId: this.unitId, sheetId: this.id, images: preparedSheetImages(this.api, this.unitId, this.id, images) })
    return this
  }
  getImages(): FOverGridImage[] {
    const workbook = this.workbook(), sheet = this.sheet(workbook)
    return (sheet.images ?? []).filter(image => canViewSheetImage(this.api, workbook, sheet, image))
      .map(image => new FOverGridImage(this.api, this.unitId, this.id, image.id))
  }
  getImageById(id: string): FOverGridImage | null {
    const workbook = this.workbook(), sheet = this.sheet(workbook), image = sheet.images?.find(item => item.id === id)
    return image && canViewSheetImage(this.api, workbook, sheet, image) ? new FOverGridImage(this.api, this.unitId, this.id, id) : null
  }
  deleteImages(images: readonly FOverGridImage[]): this {
    if (images.some(image => image.getUnitId() !== this.unitId || image.getSubUnitId() !== this.id)) throw new Error('图片必须属于当前工作表。')
    this.api.executeCommandSync(SHEET_IMAGE_COMMANDS.remove, { unitId: this.unitId, sheetId: this.id, imageIds: images.map(image => image.getId()) })
    return this
  }
  alignImages(images: readonly FOverGridImage[], alignType: SheetImageAlignType): this {
    if (images.some(image => image.getUnitId() !== this.unitId || image.getSubUnitId() !== this.id)) throw new Error('图片必须属于当前工作表。')
    this.api.executeCommandSync(SHEET_IMAGE_ALIGN_OPERATIONS.align, { unitId: this.unitId, sheetId: this.id,
      drawings: images.map(image => ({ unitId: this.unitId, subUnitId: this.id, drawingId: image.getId() })), alignType })
    return this
  }
  groupDrawings(selection: readonly SheetDrawingRef[], groupId?: string): this {
    this.api.executeCommandSync(SHEET_DRAWING_GROUP_COMMAND_IDS.group, { unitId: this.unitId, sheetId: this.id,
      selection: selection.map(ref => ({ ...ref })), ...(groupId === undefined ? {} : { groupId }) })
    return this
  }
  ungroupDrawings(selection: readonly SheetDrawingRef[]): this {
    this.api.executeCommandSync(SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup, { unitId: this.unitId, sheetId: this.id, selection: selection.map(ref => ({ ...ref })) })
    return this
  }
  transformDrawings(selection: readonly SheetDrawingRef[], transform: { frame?: SheetDrawingAffineGeometry; delta?: SheetDrawingMatrix }): this {
    this.api.executeCommandSync(SHEET_DRAWING_GROUP_COMMAND_IDS.transform, { ...transform, unitId: this.unitId, sheetId: this.id,
      selection: selection.map(ref => ({ ...ref })) })
    return this
  }
  arrangeDrawings(selection: readonly SheetDrawingRef[], arrangeType: SheetDrawingArrangeType | 0 | 1 | 2 | 3): this {
    this.api.executeCommandSync(SHEET_DRAWING_GROUP_COMMAND_IDS.arrange, { unitId: this.unitId, sheetId: this.id,
      selection: selection.map(ref => ({ ...ref })), arrangeType })
    return this
  }
  alignDrawings(selection: readonly SheetDrawingRef[], alignType: SheetDrawingAlignType): this {
    this.api.executeCommandSync(SHEET_DRAWING_GROUP_COMMAND_IDS.align, { unitId: this.unitId, sheetId: this.id,
      selection: selection.map(ref => ({ ...ref })), alignType })
    return this
  }
  moveDrawings(selection: readonly SheetDrawingRef[], direction: SheetDrawingMoveDirection, step: 1 | 10 = 1): this {
    this.api.executeCommandSync(SHEET_DRAWING_GROUP_COMMAND_IDS.move, { unitId: this.unitId, sheetId: this.id,
      selection: selection.map(ref => ({ ...ref })), direction, step })
    return this
  }
  removeDrawings(selection: readonly SheetDrawingRef[]): this {
    this.api.executeCommandSync(SHEET_DRAWING_GROUP_COMMAND_IDS.remove, { unitId: this.unitId, sheetId: this.id, selection: selection.map(ref => ({ ...ref })) })
    return this
  }
  saveCellImagesAsync(options: SaveCellImagesOptions = {}, ranges?: readonly FRange[]): Promise<boolean> {
    if (ranges?.some(range => range.getUnitId() !== this.unitId || range.getSheetId() !== this.id)) throw new Error('图片范围必须属于当前工作表。')
    return this.api.executeCommand(SHEET_IMAGE_COMMANDS.saveCell, { unitId: this.unitId, sheetId: this.id, saveConfig: sheetImageSaveConfig(options),
      ranges: ranges ? ranges.map(range => toCellRange(range.getRange())) : [{ startRow: 0, startCol: 0, endRow: 1_048_575, endCol: 16_383 }] })
  }
  getDefaultStyle(): SheetStyleSource { return cloneModel(this.sheet().defaultStyle ?? null) }
  setDefaultStyle(style: SheetStyleInput): this {
    this.api.executeCommandSync(SHEET_DEFAULT_STYLE_COMMANDS.worksheet, { unitId: this.unitId, subUnitId: this.id, defaultStyle: style })
    return this
  }
  getRowDefaultStyle(index: number, keepRaw = false): SheetStyleSource {
    return this.axisDefaultStyle(index, 'rowStyles', keepRaw)
  }
  getColumnDefaultStyle(index: number, keepRaw = false): SheetStyleSource {
    return this.axisDefaultStyle(index, 'columnStyles', keepRaw)
  }
  setRowDefaultStyle(index: number, style: SheetStyleInput): this {
    this.api.executeCommandSync(SHEET_DEFAULT_STYLE_COMMANDS.row, { unitId: this.unitId, subUnitId: this.id, index, style })
    return this
  }
  setColumnDefaultStyle(index: number, style: SheetStyleInput): this {
    this.api.executeCommandSync(SHEET_DEFAULT_STYLE_COMMANDS.column, { unitId: this.unitId, subUnitId: this.id, index, style })
    return this
  }
  setRowStylePriority(rowPriority: boolean): this {
    this.api.executeCommandSync(SHEET_DEFAULT_STYLE_COMMANDS.rowPriority, { unitId: this.unitId, subUnitId: this.id, rowPriority })
    return this
  }
  private axisDefaultStyle(index: number, key: 'rowStyles' | 'columnStyles', keepRaw: boolean): SheetStyleSource {
    if (!Number.isSafeInteger(index) || index < 0 || index >= (key === 'rowStyles' ? 1_048_576 : 16_384)) throw new RangeError('默认样式的行列索引越界。')
    const style = sheetAxisStyleSource(this.sheet(), key === 'rowStyles' ? 'row' : 'column', index)
    return cloneModel((keepRaw ? style : resolveSheetStyleSource(style, this.workbook())) ?? null)
  }
  getDataValidations(): DataValidationRule[] {
    this.api.assertCanViewWorksheet(this.unitId, this.id)
    const sheet = this.sheet(), canRead = sheetValidationSourceReader(this.api.currentUserId)!
    return cloneModel((sheet.dataValidations ?? []).filter(rule => rule.ranges.every(range => canRead(sheet, range))))
  }
  getDataValidation(range: FRange | FRangeCoordinates | string): DataValidationRule | null {
    const bounds = this.ownedRange(range)
    this.api.assertCanViewRange(this.unitId, this.id, bounds)
    return this.getDataValidations().find(rule => rule.ranges.some(item => rangesIntersect(item, bounds))) ?? null
  }
  setDataValidation(range: FRange | FRangeCoordinates | string, input: DataValidationRuleInput | null): this {
    const bounds = this.ownedRange(range)
    if (input === null) this.api.executeCommandSync('sheet.command.clear-data-validation', { unitId: this.unitId, sheetId: this.id, ranges: [bounds] })
    else {
      const rule = createSheetDataValidationRule(input, [bounds])
      this.api.executeCommandSync('sheet.command.set-data-validation', { unitId: this.unitId, sheetId: this.id, rule })
    }
    return this
  }
  /** Large checkbox rules use the shared cancellable compute transaction. */
  async setDataValidationAsync(range: FRange | FRangeCoordinates | string, input: DataValidationRuleInput | null,
    options: SheetDataValidationAsyncOptions = {}): Promise<this> {
    if (options.signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '数据验证操作已取消。')
    const bounds = this.ownedRange(range)
    if (input === null) await this.api.executeCommand('sheet.command.clear-data-validation', { ...options, unitId: this.unitId, sheetId: this.id, ranges: [bounds] })
    else {
      const rule = createSheetDataValidationRule(input, [bounds])
      await this.api.executeCommand(SHEET_DATA_VALIDATION_ASYNC_COMMANDS.set, { ...options, unitId: this.unitId, sheetId: this.id, rule })
    }
    return this
  }
  getDataValidationErrors(range: FRange | FRangeCoordinates | string): DataValidationError[] {
    const bounds = this.ownedRange(range)
    this.api.assertCanViewRange(this.unitId, this.id, bounds)
    const count = (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1)
    if (!Number.isSafeInteger(count) || count > 250_000) throw new RangeError('数据验证错误查询超出 250000 格预算，请按区域查询。')
    const positions: Array<{ row: number; column: number }> = []
    for (let row = bounds.startRow; row <= bounds.endRow; row++) for (let column = bounds.startCol; column <= bounds.endCol; column++) positions.push({ row, column })
    return validateSheetDataValidationChanges(this.workbook(), this.id, positions, sheetValidationSourceReader(this.api.currentUserId), this.api.regionalSettings.region.value).issues
  }
  addRowOutline(startRow: number, numRows: number): this { return this.addDimensionOutline(DimensionOutlineAxis.ROW, startRow, numRows) }
  addColumnOutline(startColumn: number, numColumns: number): this { return this.addDimensionOutline(DimensionOutlineAxis.COLUMN, startColumn, numColumns) }
  private addDimensionOutline(axis: DimensionOutlineAxis, start: number, count: number): this {
    const sheet = this.sheet()
    try {
      const range = sheetOutlinePublicAddRange(axis, start, count, axis === DimensionOutlineAxis.ROW ? sheet.rowCount : sheet.columnCount)
      this.api.executeCommandSync('sheet.command.add-dimension-outline', { unitId: this.unitId, subUnitId: this.id, ...range,
        expectedRevision: this.api.getWorkbookRevision(this.unitId), expectedHistoryGeneration: this.api.editor.history.captureSavepoint(this.unitId).generation,
        expectedUserId: this.api.currentUserId, host: 'api' })
    } catch (error) { if (!isIgnoredOutlineAddError(error)) throw error }
    return this
  }
  removeDimensionOutline(outlineId: string): this {
    this.sheet()
    this.api.executeCommandSync('sheet.command.remove-dimension-outline', { unitId: this.unitId, subUnitId: this.id, outlineId,
      expectedRevision: this.api.getWorkbookRevision(this.unitId), expectedHistoryGeneration: this.api.editor.history.captureSavepoint(this.unitId).generation,
      expectedUserId: this.api.currentUserId, host: 'api' })
    return this
  }
  setDimensionOutlineCollapsed(outlineId: string, collapsed: boolean): this {
    this.sheet()
    this.api.executeCommandSync('sheet.command.set-dimension-outline-collapsed', { unitId: this.unitId, subUnitId: this.id, outlineId, collapsed,
      expectedRevision: this.api.getWorkbookRevision(this.unitId), expectedHistoryGeneration: this.api.editor.history.captureSavepoint(this.unitId).generation,
      expectedUserId: this.api.currentUserId, host: 'api' })
    return this
  }
  clearDimensionOutlines(axis: DimensionOutlineAxis, start: number, end: number): this {
    this.sheet()
    this.api.executeCommandSync('sheet.command.clear-dimension-outlines', { unitId: this.unitId, subUnitId: this.id, axis, start, end,
      expectedRevision: this.api.getWorkbookRevision(this.unitId), expectedHistoryGeneration: this.api.editor.history.captureSavepoint(this.unitId).generation,
      expectedUserId: this.api.currentUserId, host: 'api' })
    return this
  }
  getDimensionOutlines(axis?: DimensionOutlineAxis): IDimensionOutline[] {
    const sheet = this.sheet(), outlines = getSheetDimensionOutlines(sheet.outline, sheetOutlineScope(this.workbook(), sheet), axis)
    for (const group of outlines) this.api.assertCanViewRange(this.unitId, this.id, group.axis === 'row'
      ? { startRow: group.start, endRow: group.end, startCol: 0, endCol: sheet.columnCount - 1 }
      : { startRow: 0, endRow: sheet.rowCount - 1, startCol: group.start, endCol: group.end })
    return outlines
  }
  getSparkline(address: string): SparklineSpec | null {
    const cell = parseA1Cell(address)
    const sheet = this.sheet()
    assertSparklineReadPermission(sheet, [cellAddress(cell.row, cell.column)], this.api.currentUserId)
    const value = sheet.cells[`${cell.row}:${cell.column}`]?.sparkline
    return cloneModel(value ?? null)
  }
  getSparklineGroups(ranges?: readonly string[]): SparklineGroupInfo[] {
    const sheet = this.sheet(), groups = readSparklineGroups(sheet, ranges)
    assertSparklineReadPermission(sheet, groups.flatMap(group => group.members.map(member => member.address)), this.api.currentUserId)
    return cloneModel(groups)
  }
  addSparkline(sourceRanges: ApiRange[], targetRanges: ApiRange[], type: SparklineTypeEnum = SparklineTypeEnum.LINE_CHART): FSparklineGroup | undefined {
    this.sheet()
    if (!Array.isArray(sourceRanges) || !Array.isArray(targetRanges)) throw new TypeError('迷你图 API 须提供源数据和放置区域数组。')
    const groupId = createSparklineGroupId()
    const changed = this.applySparklineMutation({ kind: 'insert', groupId, settings: { type: sparklineTypeFromPublic(type) },
      sourceRanges: sourceRanges.map(range => sparklineApiRangeReference(this.api, this.unitId, this.id, range)),
      targetRanges: targetRanges.map(range => sparklineApiRangeReference(this.api, this.unitId, this.id, range, true)) })
    return changed ? new FSparklineGroup(this.api, this, groupId) : undefined
  }
  getAllSubSparkline(): Map<string, ISparklineGroup> | undefined {
    const groups = this.getSparklineGroups(), workbook = this.workbook()
    if (!groups.length) return undefined
    return new Map(groups.map(group => [group.groupId, sparklineGroupToPublic(group,
      name => workbook.sheets.find(sheet => sheet.id === name || sheet.name.toLocaleLowerCase() === name.toLocaleLowerCase())?.id,
      { host: this.sheet(), workbook })]))
  }
  composeSparkline(ranges: ApiRange[]): void {
    if (!Array.isArray(ranges)) throw new TypeError('迷你图组合 API 须提供区域数组。')
    this.applySparklineMutation({ kind: 'combine', ranges: ranges.map(range => sparklineApiRangeReference(this.api, this.unitId, this.id, range, true)) })
  }
  unComposeSparkline(ranges: ApiRange[]): void {
    if (!Array.isArray(ranges)) throw new TypeError('迷你图取消组合 API 须提供区域数组。')
    this.applySparklineMutation({ kind: 'uncombine', ranges: ranges.map(range => sparklineApiRangeReference(this.api, this.unitId, this.id, range, true)) })
  }
  getSparklineByCell(row: number, col: number): FSparkline | undefined {
    const address = sparklineCellAddress(row, col), group = this.getSparklineGroups([address])[0]
    return group ? new FSparkline(this.api, this, group.groupId, row, col) : undefined
  }
  getSparklineGroupByCell(row: number, col: number): FSparklineGroup | undefined {
    const group = this.getSparklineGroups([sparklineCellAddress(row, col)])[0]
    return group ? new FSparklineGroup(this.api, this, group.groupId) : undefined
  }
  applySparklineMutation(request: SparklineMutationRequest, context: SparklineMutationContext = {}): boolean {
    this.sheet()
    return this.api.executeCommandSync(SET_SPARKLINES_COMMAND, { unitId: this.unitId, sheetId: this.id, request,
      expectedRevision: context.expectedRevision ?? this.api.getWorkbookRevision(this.unitId),
      historyGeneration: context.historyGeneration ?? this.api.editor.history.captureSavepoint(this.unitId).generation,
      expectedUserId: context.expectedUserId ?? this.api.currentUserId, host: context.host })
  }
  setSparklines(request: SparklineGroupRequest): this {
    const patches = createSparklineGroupPatches(request)
    this.api.executeCommandSync(SET_SPARKLINES_COMMAND, { unitId: this.unitId, sheetId: this.id, patches,
      expectedRevision: this.api.getWorkbookRevision(this.unitId), historyGeneration: this.api.editor.history.captureSavepoint(this.unitId).generation,
      expectedUserId: this.api.currentUserId })
    return this
  }
  clearSparklines(range: FRange | FRangeCoordinates | string): this {
    const bounds = this.ownedRange(range)
    const patches: Record<string, null> = {}
    for (const cell of Object.values(this.sheet().cells)) if (cell.sparkline && cell.row >= bounds.startRow && cell.row <= bounds.endRow
      && cell.col >= bounds.startCol && cell.col <= bounds.endCol) patches[cellAddress(cell.row, cell.col)] = null
    this.api.executeCommandSync(SET_SPARKLINES_COMMAND, { unitId: this.unitId, sheetId: this.id, patches,
      expectedRevision: this.api.getWorkbookRevision(this.unitId), historyGeneration: this.api.editor.history.captureSavepoint(this.unitId).generation,
      expectedUserId: this.api.currentUserId })
    return this
  }
  getRangeProtectionRules(): SheetRangeProtectionRule[] { return cloneModel(this.sheet().rangeProtectionRules ?? []) }
  getRangeProtectionRulesForRange(range: FRange | FRangeCoordinates | string): SheetRangeProtectionRule[] {
    const bounds = this.ownedRange(range)
    return cloneModel((this.sheet().rangeProtectionRules ?? []).filter(rule => rule.ranges.some(item => rangesIntersect(item, bounds))))
  }
  setRangeProtection(range: FRange | FRangeCoordinates | string, draft: SheetRangeProtectionDraft): SheetRangeProtectionRule {
    const rule = createSheetRangeProtectionRule(draft, [this.ownedRange(range)], this.api.currentUserId)
    this.api.executeCommandSync('sheet.command.set-range-protection', { unitId: this.unitId, sheetId: this.id, rule })
    return cloneModel(rule)
  }
  removeRangeProtection(ruleId: string): boolean {
    return this.api.executeCommandSync('sheet.command.remove-range-protection', { unitId: this.unitId, sheetId: this.id, ruleId })
  }
  getCellValidationOptions(row: number, column: number): string[] {
    const bounds = this.ownedRange({ startRow: row, endRow: row, startColumn: column, endColumn: column })
    this.api.assertCanViewRange(this.unitId, this.id, bounds)
    const merge = this.sheet().mergedRanges.find(range => rangesIntersect(range, bounds))
    const position = merge ? { row: merge.startRow, column: merge.startCol } : { row, column }
    const rules = this.getDataValidations().filter(rule => rule.ranges.some(range => position.row >= range.startRow && position.row <= range.endRow
      && position.column >= range.startCol && position.column <= range.endCol))
    return rules.flatMap(rule => sheetDataValidationOptions(rule, this.workbook(), this.id, sheetValidationSourceReader(this.api.currentUserId), this.api.regionalSettings.region.value))
  }
  setAutoFilter(range: CellRange | FRangeCoordinates | FRange | string | null): this {
    const normalized = range === null ? null : this.ownedRange(typeof range === 'object' && !(range instanceof FRange) && 'startCol' in range
      ? { startRow: range.startRow, startColumn: range.startCol, endRow: range.endRow, endColumn: range.endCol } : range)
    this.api.executeCommandSync('sheet.command.set-auto-filter', { unitId: this.unitId, sheetId: this.id, range: normalized })
    return this
  }
  getAutoFilter(): { range: CellRange; columns: Record<number, SheetFilterCriteria>; filteredRows: number[] } | null {
    const filter = this.sheet().autoFilter
    return filter ? cloneModel({ ...filter, filteredRows: getFilteredRows(this.sheet()) }) : null
  }
  setFilterCriteria(column: number, criteria: SheetFilterCriteria | null): this {
    this.api.executeCommandSync('sheet.command.set-filter-criteria', { unitId: this.unitId, sheetId: this.id, column, criteria })
    return this
  }
  clearFilter(column?: number): this {
    this.api.executeCommandSync('sheet.command.clear-filter', { unitId: this.unitId, sheetId: this.id, column })
    return this
  }
  getFilteredRows(): number[] { return getFilteredRows(this.sheet()) }
  async sortRange(range: FRange | FRangeCoordinates | string, criteria: readonly SheetSortCriterion[], options: SortRangeApiOptions = {}): Promise<this> {
    const normalized = this.ownedRange(range)
    const sheet = this.sheet()
    const expectedRevision = this.api.getWorkbookRevision(this.unitId)
    const numeric = getNumericSortColumns(sheet, normalized, criteria, options)
    let order: number[] | undefined
    if (numeric && numeric.rowCount > 0) {
      const result = await this.api.numericIndexService.sortRows(numeric, {
        documentId: this.unitId,
        modelRevision: expectedRevision,
        signal: options.signal,
        timeoutMs: options.timeoutMs,
        onProgress: options.onProgress,
      })
      if (result) {
        const dataRowCount = normalized.endRow - normalized.startRow + 1 - (options.hasHeader ? 1 : 0)
        order = expandVisibleSortOrder(dataRowCount, numeric.destinationOffsets, Array.from(result.indices))
      }
    }
    if (this.sheet() !== sheet || this.api.getWorkbookRevision(this.unitId) !== expectedRevision) throw new Error('Sort result is stale because the workbook changed during calculation')
    order ??= getSortedRowOrder(sheet, normalized, criteria, options)
    this.api.executeCommandSync('sheet.command.sort-range', { unitId: this.unitId, sheetId: this.id, range: normalized, criteria: [...criteria], options: { hasHeader: options.hasHeader, locale: options.locale }, order, expectedRevision })
    return this
  }
  getPivotTableByCell(row: number, col: number): FPivotTable | null {
    const pivot = getPivotTableAtCell(this.workbook(), this.id, row, col)
    return pivot ? new FPivotTable(this.api, this.unitId, pivot.id) : null
  }
  getChartTypes(): ReturnType<typeof chartTypeSelectOptions> { return chartTypeSelectOptions() }
  newChart(type: ChartType): ChartConfigBuilder { return new ChartConfigBuilder(type) }
  newPivotChart(type: ChartType): PivotChartConfigBuilder { return new PivotChartConfigBuilder(this.api, this.unitId, this.id, type) }
  /** Local draft-source read; it never inserts a chart or changes content history. */
  async getChartSettingsData(config: ChartConfig, options: ChartComputationReadOptions = {}): Promise<ChartData> {
    const { workbook, chart, context } = this.chartSettingsReadContext(config, options)
    const data = await this.api.chartStatisticsService.requestSourceData(workbook, chart, context).promise
    context.assertCurrent?.()
    // This request owns its ephemeral data; no cache or other consumer shares it.
    return data
  }
  /** Local read of materialized statistical identities for an unpublished draft. */
  async getChartSettingsStatistics(config: ChartConfig, options: ChartComputationReadOptions = {}): Promise<ChartComputedStatistics> {
    const { workbook, chart, context } = this.chartSettingsReadContext(config, options)
    const snapshot = await this.api.chartStatisticsService.requestSettingsStatistics(workbook, chart, context).promise
    context.assertCurrent?.()
    return snapshot.statistics
  }
  private chartSettingsReadContext(config: ChartConfig, options: ChartComputationReadOptions): {
    workbook: WorkbookModel; chart: ChartModel; context: ChartStatisticsRequestContext
  } {
    const workbook = this.workbook(), sheet = this.sheet(), userId = this.api.currentUserId
    const modelRevision = this.api.getWorkbookRevision(this.unitId)
    const chart = createChartDefinition(workbook, this.id, config, randomId('chart-source-read'))
    const context: ChartStatisticsRequestContext = { ...options, documentId: this.unitId, modelRevision, userId,
      priority: options.priority ?? 'interactive', assertCurrent: () => {
        if (options.signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '图表设置数据读取已取消')
        const current = this.api.getWorkbookModel(this.unitId)
        if (current !== workbook || current.sheets.find(candidate => candidate.id === this.id) !== sheet
          || this.api.currentUserId !== userId || this.api.getWorkbookRevision(this.unitId) !== modelRevision) {
          throw new ComputeError('COMPUTE_STALE_RESULT', '图表设置来源已在数据准备期间更新')
        }
        this.api.assertCanViewWorksheet(this.unitId, this.id)
        assertChartReadPermission(current, chart, userId)
      } }
    context.assertCurrent?.()
    return { workbook, chart, context }
  }
  async insertChart(info: ChartInfo): Promise<FChart> {
    const chartId = randomId('chart')
    await this.api.executeCommand('sheet.command.add-chart', { unitId: this.unitId, sheetId: this.id, chartId, chart: info })
    return new FChart(this.api, this.unitId, chartId)
  }
  async insertPivotChart(info: PivotChartInsertInfo, context: PivotChartCommandContext = {}, options?: PivotChartInsertOptions): Promise<FSheetPivotChart> {
    const chartId = randomId('pivot-chart')
    const accepted = await this.api.executeCommand(PIVOT_CHART_COMMANDS.insert, { ...context, unitId: this.unitId, sheetId: this.id, pivotChartId: chartId, chart: info, options })
    if (!accepted || !this.workbook().charts?.some(chart => chart.id === chartId && chart.sheetId === this.id)) throw new Error('透视图插入未能提交。')
    return new FSheetPivotChart(this.api, this.unitId, chartId)
  }
  getPivotCharts(): FSheetPivotChart[] { return (this.workbook().charts ?? []).filter(chart => chart.sheetId === this.id && (chart.ownedPivot || chart.pivotTableId))
    .map(chart => new FSheetPivotChart(this.api, this.unitId, chart.id)) }
  getPivotChart(pivotChartId: string): FSheetPivotChart | null { return this.workbook().charts?.some(chart => chart.id === pivotChartId
    && chart.sheetId === this.id && (chart.ownedPivot || chart.pivotTableId)) ? new FSheetPivotChart(this.api, this.unitId, pivotChartId) : null }
  getCharts(): FChart[] { return (this.workbook().charts ?? []).filter((chart) => chart.sheetId === this.id).map((chart) => new FChart(this.api, this.unitId, chart.id)) }
  getChartById(chartId: string): FChart | null { return this.workbook().charts?.some((chart) => chart.id === chartId && chart.sheetId === this.id) ? new FChart(this.api, this.unitId, chartId) : null }
  newSlicer(fieldId: string): FSheetSlicerBuilder { return new FSheetSlicerBuilder(this.api, this.unitId, fieldId) }
  newTimeline(fieldId: string): FSheetTimelineBuilder { return new FSheetTimelineBuilder(this.api, this.unitId, fieldId) }
  async insertSlicer(info: ISheetSlicerInfo, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<FSheetSlicer> {
    const slicerId = randomId('slicer')
    const accepted = await this.api.executeCommand(SLICER_PUBLIC_COMMANDS.insert, { unitId: this.unitId, hostSheetId: this.id, slicerId, info,
      expectedRevision: context.expectedRevision, host: context.host })
    if (!accepted || !this.getSlicer(slicerId)) throw new Error('切片器插入未能提交。')
    return new FSheetSlicer(this.api, this.unitId, slicerId)
  }
  async insertTimeline(info: ISheetTimelineInfo, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<FSheetTimeline> {
    const slicerId = randomId('timeline')
    const accepted = await this.api.executeCommand(SLICER_PUBLIC_COMMANDS.insertTimeline, { unitId: this.unitId, hostSheetId: this.id, slicerId, info,
      expectedRevision: context.expectedRevision, host: context.host })
    if (!accepted || !this.getTimeline(slicerId)) throw new Error('时间线插入未能提交。')
    return new FSheetTimeline(this.api, this.unitId, slicerId)
  }
  /** Local model-based insertion shared by existing Canvas/Ribbon adapters. */
  async insertSlicerControl(info: SlicerInfo, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<FSheetSlicer> {
    const slicerId = info.id ?? randomId(info.type === 'timeline' ? 'timeline' : 'slicer')
    const slicer: SlicerCreateInput = { ...info, id: slicerId, unitId: this.unitId, hostSheetId: this.id }
    const accepted = await this.api.executeCommand('sheet.command.add-slicer', { unitId: this.unitId, slicer,
      expectedRevision: context.expectedRevision, host: context.host })
    if (!accepted || !this.getSlicerControl(slicerId)) throw new Error('切片器或时间线插入未能提交。')
    return new FSheetSlicer(this.api, this.unitId, slicerId)
  }
  getSlicers(): FSheetSlicer[] { return this.orderedSlicerIds('slicer').map(id => new FSheetSlicer(this.api, this.unitId, id)) }
  getSlicer(slicerId: string): FSheetSlicer | null { return this.workbook().slicers?.some(slicer => slicer.hostSheetId === this.id && slicer.id === slicerId && slicer.type === 'slicer')
    ? new FSheetSlicer(this.api, this.unitId, slicerId) : null }
  getSlicerById(slicerId: string): FSheetSlicer | null { return this.getSlicer(slicerId) }
  getTimelines(): FSheetTimeline[] { return this.orderedSlicerIds('timeline').map(id => new FSheetTimeline(this.api, this.unitId, id)) }
  getTimeline(timelineId: string): FSheetTimeline | null { return this.workbook().slicers?.some(slicer => slicer.hostSheetId === this.id && slicer.id === timelineId && slicer.type === 'timeline')
    ? new FSheetTimeline(this.api, this.unitId, timelineId) : null }
  getTimelineById(timelineId: string): FSheetTimeline | null { return this.getTimeline(timelineId) }
  getSlicerControls(): FSheetSlicer[] { return this.orderedSlicerIds().map(id => new FSheetSlicer(this.api, this.unitId, id)) }
  getSlicerControl(controlId: string): FSheetSlicer | null { return this.workbook().slicers?.some(slicer => slicer.hostSheetId === this.id && slicer.id === controlId)
    ? new FSheetSlicer(this.api, this.unitId, controlId) : null }
  private orderedSlicerIds(type?: 'slicer' | 'timeline'): string[] {
    this.api.assertCanViewWorksheet(this.unitId, this.id)
    return drawingPaintOrder(createSheetDrawingScene(this.workbook(), this.id))
      .filter(ref => (ref.kind === 'slicer' || ref.kind === 'timeline') && (type === undefined || ref.kind === type)).map(ref => ref.id)
  }
  setName(name: string): this {
    const oldName = this.sheet().name
    const changed = this.api.executeCommandSync('sheet.command.set-worksheet-name', { unitId: this.unitId, sheetId: this.id, name })
    if (!changed && oldName !== name) throw new Error('Worksheet name change was canceled')
    return this
  }
  getTabColor(): string | undefined { return this.sheet().tabColor }
  setTabColor(value: string): this {
    this.api.executeCommandSync('sheet.command.set-tab-color', { unitId: this.unitId, subUnitId: this.id, value })
    return this
  }
  getHiddenState(): WorksheetHiddenState { return worksheetHiddenState(this.sheet()) }
  isSheetHidden(): boolean { return this.getHiddenState() !== 0 }
  hideSheet(): this { return this.isSheetHidden() ? this : this.setHiddenState(1) }
  showSheet(): this {
    this.api.executeCommandSync('sheet.command.set-worksheet-show', { unitId: this.unitId, subUnitId: this.id })
    return this
  }
  setHiddenState(hidden: WorksheetHiddenState): this {
    if (hidden === 0) return this.showSheet()
    if (hidden !== 1 && hidden !== 2) throw new RangeError('工作表隐藏状态无效。')
    this.api.executeCommandSync('sheet.command.set-worksheet-hidden', { unitId: this.unitId, subUnitId: this.id, hidden })
    return this
  }
  setIndex(order: number): this {
    this.api.executeCommandSync('sheet.command.set-worksheet-order', { unitId: this.unitId, subUnitId: this.id, order })
    return this
  }
  getIndex(): number { return this.workbook().sheets.findIndex((sheet) => sheet.id === this.id) }
  getMaxRows(): number { return this.sheet().rowCount ?? 1_048_576 }
  getMaxColumns(): number { return this.sheet().columnCount ?? 16_384 }
  insertRowBefore(beforePosition: number): this { return this.insertRowsBefore(beforePosition, 1) }
  insertRowAfter(afterPosition: number): this { return this.insertRowsAfter(afterPosition, 1) }
  insertRows(rowIndex: number, numRows = 1): this { return this.insertRowsBefore(rowIndex, numRows) }
  insertRowsBefore(beforePosition: number, howMany: number): this { return this.insertAxes('row', beforePosition, howMany, false) }
  insertRowsAfter(afterPosition: number, howMany: number): this { return this.insertAxes('row', afterPosition, howMany, true) }
  insertColumnBefore(beforePosition: number): this { return this.insertColumnsBefore(beforePosition, 1) }
  insertColumnAfter(afterPosition: number): this { return this.insertColumnsAfter(afterPosition, 1) }
  insertColumns(columnIndex: number, numColumns = 1): this { return this.insertColumnsBefore(columnIndex, numColumns) }
  insertColumnsBefore(beforePosition: number, howMany: number): this { return this.insertAxes('column', beforePosition, howMany, false) }
  insertColumnsAfter(afterPosition: number, howMany: number): this { return this.insertAxes('column', afterPosition, howMany, true) }
  deleteRow(rowPosition: number): this { return this.deleteRows(rowPosition, 1) }
  deleteRows(rowPosition: number, howMany: number): this {
    return this.executeAxisStructure('sheet.command.remove-row-by-range', { range: this.axisDimensionRange('row', rowPosition, howMany) })
  }
  deleteColumn(columnPosition: number): this { return this.deleteColumns(columnPosition, 1) }
  deleteColumns(columnPosition: number, howMany: number): this {
    return this.executeAxisStructure('sheet.command.remove-col-by-range', { range: this.axisDimensionRange('column', columnPosition, howMany) })
  }
  deleteRowsByPoints(points: readonly (number | readonly [number, number])[]): this { return this.deleteAxesByPoints('row', points) }
  deleteColumnsByPoints(points: readonly (number | readonly [number, number])[]): this { return this.deleteAxesByPoints('column', points) }
  moveRows(rowSpec: FRange, destinationIndex: number): this { return this.transferAxes('row', rowSpec, destinationIndex, false) }
  moveColumns(columnSpec: FRange, destinationIndex: number): this { return this.transferAxes('column', columnSpec, destinationIndex, false) }
  copyRows(rowSpec: FRange, destinationIndex: number): this { return this.transferAxes('row', rowSpec, destinationIndex, true) }
  copyColumns(columnSpec: FRange, destinationIndex: number): this { return this.transferAxes('column', columnSpec, destinationIndex, true) }
  hideRow(row: FRange): this { return this.setAxesVisibleRange('row', row, false) }
  unhideRow(row: FRange): this { return this.setAxesVisibleRange('row', row, true) }
  hideRows(rowIndex: number, numRows = 1): this { return this.setAxesVisible('row', rowIndex, numRows, false) }
  showRows(rowIndex: number, numRows = 1): this { return this.setAxesVisible('row', rowIndex, numRows, true) }
  hideColumn(column: FRange): this { return this.setAxesVisibleRange('column', column, false) }
  unhideColumn(column: FRange): this { return this.setAxesVisibleRange('column', column, true) }
  hideColumns(columnIndex: number, numColumns = 1): this { return this.setAxesVisible('column', columnIndex, numColumns, false) }
  showColumns(columnIndex: number, numColumns = 1): this { return this.setAxesVisible('column', columnIndex, numColumns, true) }
  getConditionalFormats(): SheetConditionalFormatRule[] { return cloneModel(this.sheet().conditionalFormats ?? []) }
  setConditionalFormats(rules: readonly SheetConditionalFormatRule[], expectedRevision?: number): this {
    this.api.executeCommandSync(SET_SHEET_CONDITIONAL_FORMATS_COMMAND, { unitId: this.unitId, sheetId: this.id, rules, expectedRevision })
    return this
  }
  private executeAxisStructure(commandId: string, input: SheetAxisStructureCommandInput): this {
    this.api.executeCommandSync(commandId, { ...input, unitId: this.unitId, sheetId: this.id })
    return this
  }
  private setAxesVisible(axis: 'row' | 'column', index: number, count: number, visible: boolean): this {
    const range = this.axisDimensionRange(axis, index, count)
    const command = axis === 'row' ? visible ? SHEET_AXIS_VISIBILITY_COMMANDS.showRows : SHEET_AXIS_VISIBILITY_COMMANDS.hideRows
      : visible ? SHEET_AXIS_VISIBILITY_COMMANDS.showColumns : SHEET_AXIS_VISIBILITY_COMMANDS.hideColumns
    this.api.executeCommandSync(command, { unitId: this.unitId, sheetId: this.id, ranges: [range] })
    return this
  }
  private setAxesVisibleRange(axis: 'row' | 'column', source: FRange, visible: boolean): this {
    if (!(source instanceof FRange) || source.getUnitId() !== this.unitId || source.getSheetId() !== this.id) throw new Error('目标选区必须属于当前工作表。')
    const range = source.getRange(), index = axis === 'row' ? range.startRow : range.startColumn, end = axis === 'row' ? range.endRow : range.endColumn
    return this.setAxesVisible(axis, index, end - index + 1, visible)
  }
  private insertAxes(axis: 'row' | 'column', position: number, count: number, after: boolean): this {
    const anchor = this.axisDimensionRange(axis, position, 1)
    validateSize(count, 'count')
    const start = position + (after ? 1 : 0)
    const range: CellRange = axis === 'row' ? { ...anchor, startRow: start, endRow: start + count - 1 }
      : { ...anchor, startCol: start, endCol: start + count - 1 }
    return this.executeAxisStructure(axis === 'row' ? 'sheet.command.insert-row-by-range' : 'sheet.command.insert-col-by-range', {
      range, direction: axis === 'row' ? after ? 2 : 0 : after ? 1 : 3,
    })
  }
  private deleteAxesByPoints(axis: 'row' | 'column', points: readonly (number | readonly [number, number])[]): this {
    if (!Array.isArray(points)) throw new TypeError('行列删除位置必须是数组。')
    const ranges = points.map(point => {
      if (typeof point === 'number') return this.axisDimensionRange(axis, point, 1)
      if (!Array.isArray(point) || point.length !== 2) throw new TypeError('行列删除区间必须包含开始和结束位置。')
      return this.axisDimensionRange(axis, point[0], point[1] - point[0] + 1)
    })
    if (!ranges.length) return this
    // Overlapping points are merged and removed in reverse order by one transaction.
    return this.executeAxisStructure(axis === 'row' ? 'sheet.command.remove-row' : 'sheet.command.remove-col', { ranges })
  }
  private transferAxes(axis: 'row' | 'column', source: FRange, destinationIndex: number, copy: boolean): this {
    if (!(source instanceof FRange) || source.getUnitId() !== this.unitId || source.getSheetId() !== this.id) throw new Error('源选区必须属于当前工作表。')
    const range = source.getRange(), start = axis === 'row' ? range.startRow : range.startColumn
    const end = axis === 'row' ? range.endRow : range.endColumn
    const fromRange = this.axisDimensionRange(axis, start, end - start + 1)
    return this.executeAxisStructure(`sheet.command.${copy ? 'copy' : 'move'}-${axis === 'row' ? 'rows' : 'cols'}`, { fromRange, destinationIndex })
  }
  getRowHeight(row: number): number { this.axisDimensionRange('row', row, 1); return sheetRowHeight(this.sheet(), row) }
  getColumnWidth(column: number): number { this.axisDimensionRange('column', column, 1); return sheetColumnWidth(this.sheet(), column) }
  getRowIsAutoHeight(row: number): boolean { this.axisDimensionRange('row', row, 1); return sheetRowIsAutoHeight(this.sheet(), row) }
  setRowHeight(row: number, height: number): this { return this.setRowHeights(row, 1, height) }
  setRowHeights(startRow: number, numRows: number, height: number): this {
    this.api.executeCommandSync(SHEET_AXIS_DIMENSION_COMMANDS.growingRowHeight, { unitId: this.unitId, sheetId: this.id,
      ranges: [this.axisDimensionRange('row', startRow, numRows)], value: height })
    return this
  }
  setRowHeightsForced(startRow: number, numRows: number, height: number): this {
    this.api.executeCommandSync(SHEET_AXIS_DIMENSION_COMMANDS.rowHeight, { unitId: this.unitId, sheetId: this.id,
      ranges: [this.axisDimensionRange('row', startRow, numRows)], value: height })
    return this
  }
  setRowAutoHeight(startRow: number, numRows = 1): this {
    this.api.executeCommandSync(SHEET_AXIS_DIMENSION_COMMANDS.rowAutoHeight, { unitId: this.unitId, sheetId: this.id,
      ranges: [this.axisDimensionRange('row', startRow, numRows)] })
    return this
  }
  setRangesAutoHeight(ranges: readonly (CellRange | FRangeCoordinates)[]): this {
    this.api.executeCommandSync(SHEET_AXIS_DIMENSION_COMMANDS.rowAutoHeight, { unitId: this.unitId, sheetId: this.id, ranges })
    return this
  }
  setColumnWidth(column: number, width: number): this { return this.setColumnWidths(column, 1, width) }
  setColumnWidths(startColumn: number, numColumns: number, width: number): this {
    this.api.executeCommandSync(SHEET_AXIS_DIMENSION_COMMANDS.columnWidth, { unitId: this.unitId, sheetId: this.id,
      ranges: [this.axisDimensionRange('column', startColumn, numColumns)], value: width })
    return this
  }
  setColumnAutoWidth(startColumn: number, numColumns = 1): this {
    this.api.executeCommandSync(SHEET_AXIS_DIMENSION_COMMANDS.columnAutoWidth, { unitId: this.unitId, sheetId: this.id,
      ranges: [this.axisDimensionRange('column', startColumn, numColumns)] })
    return this
  }
  private axisDimensionRange(axis: 'row' | 'column', start: number, count: number): CellRange {
    const sheet = this.sheet(), rows = sheet.rowCount ?? 1_048_576, columns = sheet.columnCount ?? 16_384
    if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(count) || count <= 0 || start + count > (axis === 'row' ? rows : columns)) throw new RangeError('行列尺寸目标超出工作表。')
    return axis === 'row' ? { startRow: start, endRow: start + count - 1, startCol: 0, endCol: columns - 1 }
      : { startRow: 0, endRow: rows - 1, startCol: start, endCol: start + count - 1 }
  }
  getRange(row: number, column: number): FRange
  getRange(row: number, column: number, numRows: number): FRange
  getRange(row: number, column: number, numRows: number, numColumns: number): FRange
  getRange(a1Notation: string): FRange
  getRange(range: FRangeCoordinates): FRange
  getRange(rowOrRange: number | string | FRangeCoordinates, column?: number, numRows = 1, numColumns = 1): FRange {
    if (typeof rowOrRange === 'string') {
      const { sheetName } = splitSheetA1Reference(rowOrRange)
      const sheet = sheetName === undefined ? this.sheet() : this.workbook().sheets.find(candidate => candidate.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase())
      if (!sheet) throw new Error(`找不到范围所在的工作表：${sheetName}`)
      this.api.assertCanViewWorksheet(this.unitId, sheet.id)
      return new FRange(this.api, this.unitId, sheet.id, parseSheetA1Range(rowOrRange, sheet))
    }
    if (typeof rowOrRange === 'object') return new FRange(this.api, this.unitId, this.id, normalizeRange(rowOrRange))
    if (column == null) throw new TypeError('A column index is required')
    validateIndex(rowOrRange, 'row')
    validateIndex(column, 'column')
    validateSize(numRows, 'numRows')
    validateSize(numColumns, 'numColumns')
    return new FRange(this.api, this.unitId, this.id, {
      startRow: rowOrRange,
      startColumn: column,
      endRow: rowOrRange + numRows - 1,
      endColumn: column + numColumns - 1,
    })
  }
  getActiveRange(): FRange | null {
    const selection = this.api.getSelection()
    if (!selection || selection.sheetId !== this.id || selection.unitId !== this.unitId) return null
    return new FRange(this.api, this.unitId, this.id, {
      startRow: selection.startRow,
      startColumn: selection.startColumn,
      endRow: selection.endRow,
      endColumn: selection.endColumn,
      ...(selection.rangeType === undefined ? {} : { rangeType: selection.rangeType }),
    })
  }
  getActiveRangeList(): FRange[] {
    const active = this.api.getSelection()
    if (active?.unitId !== this.unitId || active.sheetId !== this.id) return []
    return this.api.getSelections().map(range => new FRange(this.api, this.unitId, this.id, range))
  }
  getPrimaryCell(): FSelectionCell | null {
    const active = this.api.getSelection(), primary = this.api.getSelectionState()?.primary
    if (active?.unitId !== this.unitId || active.sheetId !== this.id || !primary) return null
    const merge = this.sheet().mergedRanges.find(range => primary.row >= range.startRow && primary.row <= range.endRow
      && primary.column >= range.startCol && primary.column <= range.endCol)
    const main = Boolean(merge && primary.row === merge.startRow && primary.column === merge.startCol)
    return { startRow: merge?.startRow ?? primary.row, endRow: merge?.endRow ?? primary.row,
      startColumn: merge?.startCol ?? primary.column, endColumn: merge?.endCol ?? primary.column,
      actualRow: primary.row, actualColumn: primary.column, rangeType: 0, isMerged: Boolean(merge), isMergedMainCell: main }
  }
  getActiveCell(): FRange | null {
    const primary = this.getPrimaryCell()
    return primary ? this.getRange({ startRow: primary.startRow, endRow: primary.endRow,
      startColumn: primary.startColumn, endColumn: primary.endColumn, rangeType: 0 }) : null
  }
  setActiveRange(range: FRange): this {
    this.ownedRange(range)
    this.api.setSelection(this.unitId, this.id, range.getRange())
    return this
  }
  setActiveRangeList(ranges: readonly FRange[], options: SheetSelectionOptions = {}): this {
    if (!Array.isArray(ranges)) throw new TypeError('选区必须是范围数组。')
    for (const range of ranges) this.ownedRange(range)
    this.api.setSelections(this.unitId, this.id, ranges.map(range => range.getRange()), options)
    return this
  }
  getSelection(): FSelection {
    const active = this.api.getSelection()
    const index = active?.unitId === this.unitId && active.sheetId === this.id ? this.api.getSelectionState()?.activeRangeIndex : -1
    return new FSelection(this, index)
  }
  /** @internal Resolve a captured Facade selection against the current sheet geometry. */
  nextSelectionRange(ranges: readonly FRange[], active: FRange | null, primary: FSelectionCell | null, direction: SheetSelectionDirection): FRange | null {
    if (!active || !primary) return null
    const sheet = this.sheet(), workbook = this.workbook()
    const filtered = (workbook.tables ?? []).filter(table => table.sheetId === this.id).flatMap(table => table.autoFilter?.filteredRows ?? [])
    const next = nextSheetSelectionCell({ ranges: ranges.map(range => ({ ...toCellRange(range.getRange()), rangeType: range.getRange().rangeType })),
      activeRangeIndex: ranges.indexOf(active), primary: { row: primary.actualRow, column: primary.actualColumn } }, direction,
      { ...sheet, rowCount: sheet.rowCount ?? 1_048_576, columnCount: sheet.columnCount ?? 16_384, filteredRows: sheet.autoFilter?.filteredRows }, filtered)
    return next ? this.getRange({ startRow: next.range.startRow, endRow: next.range.endRow, startColumn: next.range.startCol,
      endColumn: next.range.endCol, rangeType: 0 }) : null
  }
  setFrozenRows(rows: number): this {
    validateIndex(rows, 'rows', true)
    this.api.mutateWorkbook(this.unitId, 'sheet.command.set-freeze', 'Freeze worksheet rows', (model) => { this.sheet(model).frozenRows = rows })
    return this
  }
  setFrozenColumns(columns: number): this {
    validateIndex(columns, 'columns', true)
    this.api.mutateWorkbook(this.unitId, 'sheet.command.set-freeze', 'Freeze worksheet columns', (model) => { this.sheet(model).frozenCols = columns })
    return this
  }
  setFreeze(freeze: { rows?: number; columns?: number }): this {
    const rows = freeze.rows ?? 0
    const columns = freeze.columns ?? 0
    validateIndex(rows, 'rows', true)
    validateIndex(columns, 'columns', true)
    this.api.mutateWorkbook(this.unitId, 'sheet.command.set-freeze', 'Freeze worksheet panes', (model) => {
      const sheet = this.sheet(model)
      sheet.frozenRows = rows
      sheet.frozenCols = columns
    })
    return this
  }
  getFrozenRows(): number { return this.sheet().frozenRows }
  getFrozenColumns(): number { return this.sheet().frozenCols }
  setHiddenGridlines(hidden: boolean): this {
    this.api.mutateWorkbook(this.unitId, 'sheet.command.set-gridlines', 'Set worksheet gridlines', (model) => { this.sheet(model).showGridlines = !hidden })
    return this
  }
  getHiddenGridlines(): boolean { return !this.sheet().showGridlines }
  activate(): this { new FWorkbook(this.api, this.unitId).setActiveSheet(this); return this }

  /** Sheet-scoped commands must keep the identity resolved by the A1/FRange boundary. */
  private ownedRange(value: FRange | FRangeCoordinates | string): CellRange {
    const range = value instanceof FRange ? value : typeof value === 'string' ? this.getRange(value) : this.getRange(value)
    if (range.getUnitId() !== this.unitId || range.getSheetId() !== this.id) throw new Error('目标选区必须属于当前工作表。')
    return toCellRange(range.getRange())
  }

  private workbook(): WorkbookModel {
    const model = this.api.getWorkbookModel(this.unitId)
    if (!model) throw new Error(`Workbook not found: ${this.unitId}`)
    return model
  }
  private sheet(model = this.workbook()): SpreadsheetSheet {
    const sheet = model.sheets.find((item) => item.id === this.id)
    if (!sheet) throw new Error(`Worksheet not found: ${this.id}`)
    this.api.assertCanViewWorksheet(this.unitId, this.id)
    return sheet
  }
}

export class FSelection {
  private readonly worksheet: FWorksheet
  private readonly ranges: FRange[]
  private readonly active: FRange | null
  private readonly primary: FSelectionCell | null

  constructor(worksheet: FWorksheet, activeRangeIndex?: number) {
    this.worksheet = worksheet
    this.ranges = worksheet.getActiveRangeList()
    const active = worksheet.getActiveRange()?.getRange()
    this.active = activeRangeIndex === undefined ? this.ranges.find(range => {
      const value = range.getRange()
      return active && value.startRow === active.startRow && value.endRow === active.endRow
        && value.startColumn === active.startColumn && value.endColumn === active.endColumn
    }) ?? null : this.ranges[activeRangeIndex] ?? null
    this.primary = worksheet.getPrimaryCell()
  }
  getActiveRange(): FRange | null { return this.active }
  getCurrentCell(): FSelectionCell | null { return this.primary ? { ...this.primary } : null }
  getCurrentCellRange(): FRange | null { return this.primary ? this.worksheet.getRange({ startRow: this.primary.startRow, endRow: this.primary.endRow,
    startColumn: this.primary.startColumn, endColumn: this.primary.endColumn, rangeType: 0 }) : null }
  getActiveRangeList(): FRange[] { return [...this.ranges] }
  getActiveSheet(): FWorksheet { return this.worksheet }
  updatePrimaryCell(cell: FRange): FSelection {
    if (cell.getUnitId() !== this.worksheet.getUnitId() || cell.getSheetId() !== this.worksheet.getSheetId()) throw new Error('活动单元格必须属于当前工作表。')
    const next = cell.getRange()
    let index = -1
    for (let candidate = this.ranges.length - 1; candidate >= 0; candidate--) {
      const range = this.ranges[candidate]!.getRange()
      if (next.startRow >= range.startRow && next.endRow <= range.endRow && next.startColumn >= range.startColumn && next.endColumn <= range.endColumn) {
        index = candidate; break
      }
    }
    this.worksheet.setActiveRangeList(index < 0 ? [cell] : this.ranges,
      { activeRangeIndex: index < 0 ? 0 : index, primary: { row: next.startRow, column: next.startColumn } })
    return this.worksheet.getSelection()
  }
  getNextDataRange(direction: SheetSelectionDirection): FRange | null {
    return this.worksheet.nextSelectionRange(this.ranges, this.active, this.primary, direction)
  }
}

export class ChartConfigBuilder {
  private value: ChartConfig
  private readonly descriptionId = randomId('chart-draft')

  constructor(type: ChartType) {
    resolveChartType(type)
    this.value = { type, source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, orientation: 'columns' }, position: { x: 0, y: 0 }, size: { width: 480, height: 300 } }
  }
  setType(type: ChartType): this { this.value = convertChartType(this.value, type); return this }
  setSource(range: string | ChartSource, orientation: ChartSourceOrientation = 'columns'): this {
    const source = typeof range === 'string' ? { range: parseChartA1Range(range), orientation } : range
    this.value = { ...this.value, source: { ...source, range: { ...source.range } } }
    return this
  }
  setPosition(cell: string): this {
    parseA1Cell(cell)
    this.value = { ...this.value, positionCell: cell }
    return this
  }
  setAbsolutePosition(x: number, y: number): this { this.value = { ...this.value, position: { x, y }, positionCell: undefined }; return this }
  setSize(width: number, height: number): this { this.value = { ...this.value, size: { width, height } }; return this }
  setCategoryField(index: number): this {
    const mapping = this.value.options?.mapping
    const options = this.value.type === 'gauge' && mapping?.gauge
      ? mergeChartOptions(this.value.options, { mapping: { ...mapping, gauge: { ...mapping.gauge, labelIndex: index } } }) : this.value.options
    this.value = { ...this.value, categoryField: index, options }; return this
  }
  setPivotTableId(pivotTableId: string): this {
    if (!pivotTableId.trim()) throw new TypeError('Pivot table id is required')
    this.value = { ...this.value, pivotTableId }
    return this
  }
  setValueFields(fields: number[]): this {
    if (this.value.type === 'gauge') {
      if (!Array.isArray(fields) || fields.length !== 1) throw new RangeError('仪表仅支持一个数值字段。')
      validateChartSeriesField(fields[0]!)
      if (this.value.options?.mapping?.gauge) return this.setGaugeValueField(fields[0]!)
    }
    this.value = { ...this.value, valueFields: [...fields] }; return this
  }
  setSeriesToCategory(mapping: Record<number, number[]>): this {
    this.value = { ...this.value, seriesToCategory: Object.fromEntries(Object.entries(mapping).map(([field, categories]) => [field, [...categories]])) }
    return this
  }
  setTitle(title: ChartTitleSpec | undefined): this { const patch = chartTitlePatch(title); this.value = { ...this.value, ...patch, options: mergeChartOptions(this.value.options, patch.options) }; return this }
  clearTitle(): this { return this.setTitle(undefined) }
  setSubtitle(subtitle: ChartSubtitleSpec | undefined): this { const patch = chartTitlePatch(subtitle, 'subtitle'); this.value = { ...this.value, ...patch, options: mergeChartOptions(this.value.options, patch.options) }; return this }
  clearSubtitle(): this { return this.setSubtitle(undefined) }
  setPalette(palette: readonly string[]): this { this.value = { ...this.value, palette: [...palette] }; return this }
  setOptions(options: ChartSemanticOptions): this {
    const next = mergeChartOptions(this.value.options, options)
    validateChartGradientAuthoring(this.value.type, next, this.value.options)
    this.value = { ...this.value, options: next }; return this
  }
  setOption<K extends keyof ChartSemanticOptions>(key: K, value: ChartSemanticOptions[K]): this { return this.setOptions(chartOptionPatch(key, value)) }
  setAreaLineStyle(value: ChartAreaLineStyle): this { return this.setOptions(chartAreaLineStylePatch(this.value.type, chartAreaLineStyleValue(value))) }
  clearAreaLineStyle(): this { return this.setOptions(chartAreaLineStylePatch(this.value.type, undefined)) }
  setInvalidValueStrategy(value: ChartInvalidValueStrategy): this { return this.setOption('invalidValueStrategy', chartInvalidValueStrategyValue(value)) }
  resetInvalidValueStrategy(): this { return this.setOption('invalidValueStrategy', undefined) }
  setAutoGradientFill(value: boolean): this { validateChartAutomaticGradientValue(value); return this.setOption('autoGradientFill', value) }
  resetAutoGradientFill(): this { return this.setOption('autoGradientFill', undefined) }
  /** Local convenience for the upstream plotArea configuration property. */
  setPlotArea(value: ChartPlotAreaOptions): this { validateChartPlotArea(value); return this.setOption('plotArea', value) }
  resetPlotArea(): this { return this.setOption('plotArea', undefined) }
  setMapping(mapping: ChartSemanticOptions['mapping']): this { return this.setOption('mapping', mapping) }
  setUseDateAxis(value: boolean): this { return this.setOption('mapping', { useDateAxis: value }) }
  resetUseDateAxis(): this { return this.setUseDateAxis(false) }
  setAxisPointer(value: ChartNullablePatch<ChartAxisPointerOptions> | undefined): this { return this.setOption('axisPointer', chartAxisPointerValues(value)) }
  clearAxisPointer(target?: ChartAxisPointerTarget): this { return this.setOption('axisPointer', clearChartAxisPointer(this.value.options?.axisPointer, target)) }
  setBubbleMapping(mapping: ChartBubbleMappingOptions): this { assertChartBubbleType(this.value.type); return this.setOption('mapping', { bubble: mapping }) }
  clearBubbleMapping(): this { assertChartBubbleType(this.value.type); return this.setOption('mapping', { bubble: undefined }) }
  setHistogramField(index: number): this { assertChartHistogramType(this.value.type); validateChartSeriesField(index); return this.setValueFields([index]) }
  clearHistogramField(): this { assertChartHistogramType(this.value.type); return this.setValueFields([]) }
  setBinCount(count: number): this { assertChartHistogramType(this.value.type); return isChartHistogramAuthoredNumber(count) ? this.setOption('histogram', { binning: { mode: 'count', count } }) : this }
  setBinWidth(width: number): this { assertChartHistogramType(this.value.type); return isChartHistogramAuthoredNumber(width) ? this.setOption('histogram', { binning: { mode: 'width', width } }) : this }
  resetBinning(): this { assertChartHistogramType(this.value.type); return this.setOption('histogram', { binning: undefined }) }
  setUnderflowThreshold(value: number): this { assertChartHistogramType(this.value.type); return isChartHistogramAuthoredNumber(value) ? this.setOption('histogram', { underflowThreshold: value }) : this }
  clearUnderflowThreshold(): this { assertChartHistogramType(this.value.type); return this.setOption('histogram', { underflowThreshold: undefined }) }
  setOverflowThreshold(value: number): this { assertChartHistogramType(this.value.type); return isChartHistogramAuthoredNumber(value) ? this.setOption('histogram', { overflowThreshold: value }) : this }
  clearOverflowThreshold(): this { assertChartHistogramType(this.value.type); return this.setOption('histogram', { overflowThreshold: undefined }) }
  setBinGap(value: number): this { assertChartHistogramType(this.value.type); return isChartHistogramAuthoredNumber(value) ? this.setOption('histogram', { binGap: value }) : this }
  resetBinGap(): this { assertChartHistogramType(this.value.type); return this.setOption('histogram', { binGap: undefined }) }
  setHierarchyFields(indexes: readonly number[]): this {
    const fields = this.value.type === 'sunburst' ? chartSunburstHierarchyFields(indexes)
      : (assertChartTreemapType(this.value.type), chartTreemapHierarchyFields(indexes))
    return this.setOption('mapping', { hierarchyIndexes: fields })
  }
  setValueField(index: number): this {
    const field = this.value.type === 'sunburst' ? chartSunburstValueField(index)
      : (assertChartTreemapType(this.value.type), chartTreemapValueField(index))
    return this.setValueFields([field])
  }
  setParentLabelLayout(layout: NonNullable<ChartTreemapOptions['parentLabelLayout']>): this { assertChartTreemapType(this.value.type); return this.setOption('treemap', { parentLabelLayout: chartTreemapParentLabelLayout(layout) }) }
  resetParentLabelLayout(): this { assertChartTreemapType(this.value.type); return this.setOption('treemap', { parentLabelLayout: undefined }) }
  setTreemapLabel(label: ChartNullablePatch<ChartTreemapLabelOptions>): this { assertChartTreemapType(this.value.type); return this.setOption('treemap', { label: chartTreemapLabelValues(label) }) }
  resetTreemapLabel(): this { assertChartTreemapType(this.value.type); return this.setOption('treemap', { label: undefined }) }
  setSunburstLabel(label: ChartNullablePatch<ChartSunburstLabelOptions>): this { assertChartSunburstType(this.value.type); return this.setOption('sunburst', { label: chartSunburstLabelValues(label) }) }
  resetSunburstLabel(): this { assertChartSunburstType(this.value.type); return this.setOption('sunburst', { label: undefined }) }
  setCenterLabel(label: ChartNullablePatch<ChartSunburstCenterLabelOptions>): this { assertChartSunburstType(this.value.type); return this.setOption('sunburst', { centerLabel: chartSunburstCenterLabelValues(label) }) }
  resetCenterLabel(): this { assertChartSunburstType(this.value.type); return this.setOption('sunburst', { centerLabel: undefined }) }
  setChordMapping(mapping: ChartChordMapping): this { assertChartChordType(this.value.type); return this.setOption('mapping', { chord: { ...chartChordMappingValues(mapping), valueIndex: mapping.valueIndex }, edge: undefined }) }
  setChordLayout(layout: ChartNullablePatch<ChartChordLayoutOptions>): this {
    assertChartChordType(this.value.type)
    return this.setOption('chord', chartChordLayoutPatch(this.value.options?.chord, layout))
  }
  resetChordLayout(): this {
    assertChartChordType(this.value.type)
    return this.setOption('chord', { layout: undefined, startAngle: undefined, clockwise: undefined,
      padAngle: undefined, minAngle: undefined, innerRadiusRatio: undefined, outerRadiusRatio: undefined })
  }
  setChordNode(node: ChartNullablePatch<ChartChordNodeOptions>): this { assertChartChordType(this.value.type); return this.setOption('chord', { node: chartChordNodeValues(node) }) }
  resetChordNode(): this { assertChartChordType(this.value.type); return this.setOption('chord', { node: undefined }) }
  setChordLabel(label: ChartNullablePatch<ChartChordLabelOptions>): this { assertChartChordType(this.value.type); return this.setOption('chord', { label: chartChordLabelValues(label) }) }
  resetChordLabel(): this { assertChartChordType(this.value.type); return this.setOption('chord', { label: undefined }) }
  setChordRibbon(ribbon: ChartNullablePatch<ChartChordRibbonOptions>): this { assertChartChordType(this.value.type); return this.setOption('chord', chartChordRibbonPatch(this.value.options?.chord, ribbon)) }
  resetChordRibbon(): this { assertChartChordType(this.value.type); return this.setOption('chord', { ribbon: undefined, ribbonOpacity: undefined }) }
  setChordEmphasis(emphasis: ChartNullablePatch<ChartChordEmphasisOptions>): this { assertChartChordType(this.value.type); return this.setOption('chord', { emphasis: chartChordEmphasisValues(emphasis) }) }
  resetChordEmphasis(): this { assertChartChordType(this.value.type); return this.setOption('chord', { emphasis: undefined }) }
  setCandlestickFields(fields: ChartCandlestickMappingOptions): this { assertChartCandlestickType(this.value.type); return this.setOption('mapping', { candlestick: chartCandlestickFields(fields) }) }
  clearCandlestickFields(): this { assertChartCandlestickType(this.value.type); return this.setOption('mapping', { candlestick: undefined }) }
  setRisingStyle(style: ChartNullablePatch<ChartCandlestickStateStyleOptions>): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { rising: chartCandlestickStateStyleValues(style) }) }
  clearRisingStyle(): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { rising: undefined }) }
  setFallingStyle(style: ChartNullablePatch<ChartCandlestickStateStyleOptions>): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { falling: chartCandlestickStateStyleValues(style) }) }
  clearFallingStyle(): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { falling: undefined }) }
  setDojiStyle(style: ChartNullablePatch<ChartCandlestickStateStyleOptions>): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { doji: chartCandlestickStateStyleValues(style) }) }
  clearDojiStyle(): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { doji: undefined }) }
  setCandleWidth(width: number): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { candleWidth: chartCandleWidth(width) }) }
  resetCandleWidth(): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { candleWidth: undefined }) }
  setHighLowLineStyle(style: ChartNullablePatch<ChartCandlestickHighLowLineOptions>): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { highLowLine: chartCandlestickHighLowLineValues(style) }) }
  clearHighLowLineStyle(): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { highLowLine: undefined }) }
  setCloseMarkerStyle(style: ChartNullablePatch<ChartCandlestickCloseMarkerOptions>): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { closeMarker: chartCandlestickCloseMarkerValues(style) }) }
  clearCloseMarkerStyle(): this { assertChartCandlestickType(this.value.type); return this.setOption('candlestick', { closeMarker: undefined }) }
  setAllSeriesStyle(style: ChartNullablePatch<ChartSeriesOptions> | undefined): this {
    const next = chartSeriesStyleValues(style)
    validateChartGradientSeriesPatch(this.value.type, next, '所有系列')
    return this.setOption('allSeriesStyle', next)
  }
  clearAllSeriesStyle(target?: ChartAllSeriesStyleTarget): this { return this.setOption('allSeriesStyle', clearChartAllSeriesStyle(this.value.options?.allSeriesStyle, target)) }
  setSeries(selector: ChartSeriesSelector, style: ChartNullablePatch<ChartSeriesOptions> | undefined): this {
    const next = chartSeriesStyleValues(style)
    validateChartGradientSeriesPatch(this.value.type, next)
    return this.setOptions(patchChartSeriesOverride(this.value.options, selector, next))
  }
  clearSeries(selector: ChartSeriesSelector, target?: ChartSeriesClearTarget): this {
    validateChartSeriesSelector(selector)
    return this.setOptions(patchChartSeriesOverride(this.value.options, selector, clearChartSeriesStyle(chartSeriesOverride(this.value.options, selector), target)))
  }
  clearAllSeriesOverrides(): this { return this.setOptions({ series: {}, namedSeries: [] }) }
  setPositiveStyle(selector: ChartSeriesSelector, value: ChartNullablePatch<ChartWaterfallSeriesStyle>): this {
    assertChartWaterfallType(this.value.type)
    return this.setOptions(patchChartWaterfallStyle(this.value.options, selector, 'positive', value))
  }
  clearPositiveStyle(selector: ChartSeriesSelector): this {
    assertChartWaterfallType(this.value.type)
    return this.setOptions(patchChartWaterfallStyle(this.value.options, selector, 'positive', undefined))
  }
  setNegativeStyle(selector: ChartSeriesSelector, value: ChartNullablePatch<ChartWaterfallSeriesStyle>): this {
    assertChartWaterfallType(this.value.type)
    return this.setOptions(patchChartWaterfallStyle(this.value.options, selector, 'negative', value))
  }
  clearNegativeStyle(selector: ChartSeriesSelector): this {
    assertChartWaterfallType(this.value.type)
    return this.setOptions(patchChartWaterfallStyle(this.value.options, selector, 'negative', undefined))
  }
  setSubtotalStyle(selector: ChartSeriesSelector, value: ChartNullablePatch<ChartWaterfallSeriesStyle>): this {
    assertChartWaterfallType(this.value.type)
    return this.setOptions(patchChartWaterfallStyle(this.value.options, selector, 'subtotal', value))
  }
  clearSubtotalStyle(selector: ChartSeriesSelector): this {
    assertChartWaterfallType(this.value.type)
    return this.setOptions(patchChartWaterfallStyle(this.value.options, selector, 'subtotal', undefined))
  }
  setConnector(value: boolean | ChartNullablePatch<ChartWaterfallConnectorOptions>): this {
    assertChartWaterfallType(this.value.type)
    return this.setOption('waterfall', { connector: chartWaterfallConnectorValues(value) })
  }
  clearConnector(): this { assertChartWaterfallType(this.value.type); return this.setOption('waterfall', { connector: undefined }) }
  setStackType(value: ChartWaterfallStackType): this {
    assertChartWaterfallType(this.value.type); validateChartWaterfallStackType(value)
    return this.setOption('waterfall', { stackType: value, ...(value === 'stacked' ? { pointRoles: undefined } : {}) })
  }
  resetStackType(): this { assertChartWaterfallType(this.value.type); return this.setOption('waterfall', { stackType: undefined }) }
  setUseSubtotal(value: boolean): this {
    assertChartWaterfallType(this.value.type); validateChartWaterfallSubtotalValue(value)
    return this.setOption('waterfall', { useSubtotal: value })
  }
  resetUseSubtotal(): this { assertChartWaterfallType(this.value.type); return this.setOption('waterfall', { useSubtotal: undefined }) }
  setBarStyle(value: ChartNullablePatch<ChartParetoBarStyleOptions>): this {
    assertChartParetoType(this.value.type)
    return this.setOption('pareto', { barStyle: chartParetoBarStyleValues(value) })
  }
  clearBarStyle(): this { assertChartParetoType(this.value.type); return this.setOption('pareto', { barStyle: undefined }) }
  setCumulativeLineStyle(value: ChartNullablePatch<ChartParetoLineStyleOptions>): this {
    assertChartParetoType(this.value.type)
    return this.setOption('pareto', { cumulativeLineStyle: chartParetoLineStyleValues(value) })
  }
  clearCumulativeLineStyle(): this { assertChartParetoType(this.value.type); return this.setOption('pareto', { cumulativeLineStyle: undefined }) }
  setIncludeZeroValues(value: boolean): this {
    assertChartParetoType(this.value.type)
    if (typeof value !== 'boolean') throw new TypeError('是否包含零合计必须是布尔值。')
    return this.setOption('pareto', { includeZeroValues: value })
  }
  resetIncludeZeroValues(): this { assertChartParetoType(this.value.type); return this.setOption('pareto', { includeZeroValues: undefined }) }
  setVisualMapType(value: ChartHeatmapVisualMapType): this {
    assertChartHeatmapType(this.value.type)
    return this.setOption('heatmap', { visualMapType: chartHeatmapVisualMapType(value) })
  }
  resetVisualMapType(): this { assertChartHeatmapType(this.value.type); return this.setOption('heatmap', { visualMapType: undefined }) }
  setValueRange(min: number, max: number): this {
    assertChartHeatmapType(this.value.type)
    return this.setOption('heatmap', chartHeatmapValueRange(min, max))
  }
  clearValueRange(): this { assertChartHeatmapType(this.value.type); return this.setOption('heatmap', { min: undefined, max: undefined }) }
  setCellLabel(value: ChartNullablePatch<ChartHeatmapLabelOptions>): this {
    assertChartHeatmapType(this.value.type)
    return this.setOption('heatmap', { label: chartHeatmapCellLabelValues(value) })
  }
  clearCellLabel(): this { assertChartHeatmapType(this.value.type); return this.setOption('heatmap', { label: undefined }) }
  setValueUnit(value: string): this {
    assertChartHeatmapType(this.value.type)
    return this.setOption('heatmap', { unit: chartHeatmapValueUnit(value) })
  }
  clearValueUnit(): this { assertChartHeatmapType(this.value.type); return this.setOption('heatmap', { unit: undefined }) }
  setTrendlines(trendlines: readonly ChartTrendlineOptions[]): this { return this.setOption('trendlines', [...trendlines]) }
  setErrorBars(errorBars: readonly ChartErrorBarOptions[]): this { return this.setOption('errorBars', [...errorBars]) }
  setThreeD(threeD: ChartSemanticOptions['threeD']): this { return this.setOption('threeD', threeD) }
  clearThreeD(): this { return this.setOption('threeD', undefined) }
  setDepthAxis(axis: ChartSemanticOptions['depthAxis']): this { return this.setOption('depthAxis', axis) }
  clearDepthAxis(): this { return this.setOption('depthAxis', undefined) }
  setRightYAxis(axis: ChartNullablePatch<ChartAxisOptions> | undefined): this { return this.setOption('rightYAxis', axis === undefined ? undefined : chartAxisPatch(this.value.options?.rightYAxis, axis)) }
  clearRightYAxis(target?: ChartAxisTarget): this { return this.setOption('rightYAxis', clearChartAxis(this.value.options?.rightYAxis, target)) }
  setRadar(value: ChartNullablePatch<ChartRadarOptions> | undefined): this { assertChartRadarType(this.value.type); return this.setOption('radar', patchChartRadarOptions(this.value.options?.radar, value)) }
  setShape(value: ChartRadarShape | ChartWordCloudShape): this {
    if (this.value.type === 'wordCloud') return this.setOption('wordCloud', { shape: chartWordCloudShape(value) })
    if (value !== 'circle' && value !== 'polygon') throw new TypeError('雷达图形状无效。')
    return this.setRadar({ shape: value })
  }
  resetShape(): this { return this.value.type === 'wordCloud' ? this.setOption('wordCloud', { shape: undefined }) : this.setRadar({ shape: undefined }) }
  setMaskImage(resourceId: string): this { assertChartWordCloudType(this.value.type); return this.setOption('wordCloud', { maskImageId: chartWordCloudMaskImageId(resourceId), maskImage: undefined, maskImageResource: undefined }) }
  /** Local host resource adapter; upstream setMaskImage keeps its ID contract. */
  setMaskImageResource(value: ChartImageResource): this { assertChartWordCloudType(this.value.type); const resource = normalizeChartImageResource(value); return this.setOption('wordCloud', { maskImageId: resource.id, maskImage: undefined, maskImageResource: resource }) }
  clearMaskImage(): this { assertChartWordCloudType(this.value.type); return this.setOption('wordCloud', { maskImageId: undefined, maskImage: undefined, maskImageResource: undefined }) }
  setRepeat(value: boolean): this { assertChartWordCloudType(this.value.type); return this.setOption('wordCloud', { repeat: chartWordCloudRepeat(value) }) }
  resetRepeat(): this { assertChartWordCloudType(this.value.type); return this.setOption('wordCloud', { repeat: undefined }) }
  setLayout(value: ChartRelationLayout): this { assertChartRelationType(this.value.type); return this.setOption('relation', { layout: chartRelationLayoutValue(value) }) }
  resetLayout(): this { assertChartRelationType(this.value.type); return this.setOption('relation', { layout: undefined }) }
  setUseValueAsSymbolSize(value: boolean): this { assertChartRelationType(this.value.type); return this.setOption('relation', { useValueAsSymbolSize: chartRelationBoolean(value) }) }
  resetUseValueAsSymbolSize(): this { assertChartRelationType(this.value.type); return this.setOption('relation', { useValueAsSymbolSize: undefined }) }
  setEmphasisEnabled(value: boolean): this { assertChartRelationType(this.value.type); return this.setOption('relation', { emphasisEnabled: chartRelationBoolean(value) }) }
  resetEmphasisEnabled(): this { assertChartRelationType(this.value.type); return this.setOption('relation', { emphasisEnabled: undefined }) }
  setForce(value: ChartRelationForceOptions): this { assertChartRelationType(this.value.type); return this.setOption('relation', { force: chartRelationForce(value) }) }
  clearForce(): this { assertChartRelationType(this.value.type); return this.setOption('relation', { force: undefined }) }
  setCircularLabelRotation(value: boolean): this { assertChartRelationType(this.value.type); return this.setOption('relation', { circularLabelRotation: chartRelationBoolean(value) }) }
  resetCircularLabelRotation(): this { assertChartRelationType(this.value.type); return this.setOption('relation', { circularLabelRotation: undefined }) }
  setNodeShape(value: ChartPointShape): this { assertChartRelationType(this.value.type); return this.setOption('relation', { nodeShape: chartRelationNodeShape(value) }) }
  resetNodeShape(): this { assertChartRelationType(this.value.type); return this.setOption('relation', { nodeShape: undefined }) }
  setFill(value: boolean): this { return this.setRadar({ fill: value }) }
  resetFill(): this { return this.setRadar({ fill: undefined }) }
  setRadarAxisLabel(value: ChartNullablePatch<ChartRadarAxisLabelOptions>): this { return this.setRadar({ axisLabel: value }) }
  resetRadarAxisLabel(): this { return this.setRadar({ axisLabel: undefined }) }
  setGauge(value: ChartGaugeOptions | undefined): this { return this.setOption('gauge', patchChartGaugeOptions(this.value.options?.gauge, value)) }
  setGaugePreset(preset: ChartGaugePreset): this { return this.setGauge(chartGaugePresetOptions(preset)) }
  setGaugeValueField(index: number): this {
    const mapping = chartGaugeValueFieldMapping(this.value.options?.mapping, index, this.value.categoryField ?? 0)
    const options = mergeChartOptions(this.value.options, { mapping })
    this.value = { ...this.value, valueFields: [index], options }; return this
  }
  setGaugeMode(value: ChartGaugeMode): this { return this.setGauge({ mode: value }) }
  resetGaugeMode(): this { return this.setGauge({ mode: undefined }) }
  setGaugeScale(value: ChartGaugeNullablePatch<ChartGaugeScaleOptions>): this { return this.setGauge(chartGaugeNullableValues<ChartGaugeScaleOptions>(value)) }
  resetGaugeScale(): this { return this.setGauge({ min: undefined, max: undefined, splitNumber: undefined }) }
  setGaugeTicks(value: ChartGaugeTicksPatch): this { return this.setGauge(chartGaugeTicksOptions(value)) }
  setGaugeAngles(value: ChartGaugeNullablePatch<ChartGaugeAnglesOptions>): this { return this.setGauge(chartGaugeNullableValues<ChartGaugeAnglesOptions>(value)) }
  resetGaugeAngles(): this { return this.setGauge({ startAngle: undefined, endAngle: undefined }) }
  setGaugeRanges(value: readonly ChartGaugeRangeOptions[]): this { return this.setGauge(chartGaugeRangesOptions(value)) }
  resetGaugeRanges(): this { return this.setGauge({ ranges: undefined }) }
  setGaugeAxisLine(value: ChartGaugeNullablePatch<ChartGaugeAxisLineOptions>): this { return this.setGauge({ axisLine: chartGaugeNullableValues<ChartGaugeAxisLineOptions>(value) }) }
  resetGaugeAxisLine(): this { return this.setGauge({ axisLine: undefined }) }
  setGaugeAxisTick(value: ChartGaugeNullablePatch<ChartGaugeAxisTickOptions>): this { return this.setGauge({ axisTick: chartGaugeNullableValues<ChartGaugeAxisTickOptions>(value) }) }
  resetGaugeAxisTick(): this { return this.setGauge({ axisTick: undefined }) }
  setGaugeSplitLine(value: ChartGaugeNullablePatch<ChartGaugeSplitLineOptions>): this { return this.setGauge({ splitLine: chartGaugeNullableValues<ChartGaugeSplitLineOptions>(value) }) }
  resetGaugeSplitLine(): this { return this.setGauge({ splitLine: undefined }) }
  setGaugeAxisLabel(value: ChartGaugeNullablePatch<ChartGaugeAxisLabelOptions>): this { return this.setGauge({ axisLabel: chartGaugeNullableValues<ChartGaugeAxisLabelOptions>(value) }) }
  resetGaugeAxisLabel(): this { return this.setGauge({ axisLabel: undefined }) }
  setGaugePointer(value: ChartGaugeNullablePatch<ChartGaugePointerOptions>): this { return this.setGauge({ pointer: chartGaugeNullableValues<ChartGaugePointerOptions>(value) }) }
  resetGaugePointer(): this { return this.setGauge({ pointer: undefined }) }
  setGaugeAnchor(value: ChartGaugeNullablePatch<ChartGaugeAnchorOptions>): this { return this.setGauge({ anchor: chartGaugeNullableValues<ChartGaugeAnchorOptions>(value) }) }
  resetGaugeAnchor(): this { return this.setGauge({ anchor: undefined }) }
  setGaugeProgress(value: ChartGaugeNullablePatch<ChartGaugeProgressOptions>): this { return this.setGauge({ progress: chartGaugeNullableValues<ChartGaugeProgressOptions>(value) }) }
  resetGaugeProgress(): this { return this.setGauge({ progress: undefined }) }
  setGaugeTitle(value: ChartGaugeNullablePatch<ChartGaugeTitleOptions>): this { return this.setGauge({ title: chartGaugeNullableValues<ChartGaugeTitleOptions>(value) }) }
  resetGaugeTitle(): this { return this.setGauge({ title: undefined }) }
  setGaugeDetail(value: ChartGaugeNullablePatch<ChartGaugeDetailOptions>): this { return this.setGauge({ detail: chartGaugeNullableValues<ChartGaugeDetailOptions>(value) }) }
  resetGaugeDetail(): this { return this.setGauge({ detail: undefined }) }
  setLegend(legend: ChartLegendSpec): this { this.value = { ...this.value, legend: chartLegendPatch(this.value.legend, legend) }; return this }
  clearLegend(): this { this.value = { ...this.value, legend: undefined }; return this }
  setXAxis(axis: ChartNullablePatch<ChartAxisOptions>): this { this.value = { ...this.value, xAxis: chartAxisPatch(this.value.xAxis, axis) }; return this }
  setYAxis(axis: ChartNullablePatch<ChartAxisOptions>): this { this.value = { ...this.value, yAxis: chartAxisPatch(this.value.yAxis, axis) }; return this }
  clearXAxis(target?: ChartAxisTarget): this { this.value = { ...this.value, xAxis: clearChartAxis(this.value.xAxis, target) }; return this }
  clearYAxis(target?: ChartAxisTarget): this { this.value = { ...this.value, yAxis: clearChartAxis(this.value.yAxis, target) }; return this }
  setXAxisTitle(title: string): this { return this.setXAxis({ title }) }
  clearXAxisTitle(): this { return this.setXAxis({ title: undefined }) }
  setYAxisTitle(title: string): this { return this.setYAxis({ title }) }
  clearYAxisTitle(): this { return this.setYAxis({ title: undefined }) }
  setRightYAxisTitle(title: string): this { return this.setOption('rightYAxis', chartAxisPatch(this.value.options?.rightYAxis, { title })) }
  clearRightYAxisTitle(): this { return this.setOption('rightYAxis', chartAxisPatch(this.value.options?.rightYAxis, { title: undefined })) }
  describe(context: Omit<ChartDescriptionContext, 'id'> = {}): ChartDescription {
    return describeChartConfiguration(this.value, { ...context, id: this.descriptionId })
  }
  build(): ChartInfo { return cloneModel(this.value) }
}

class ChartVisualFacade {
  protected readonly api: OfficeStudio
  protected readonly unitId: string
  readonly id: string
  constructor(api: OfficeStudio, unitId: string, id: string) { this.api = api; this.unitId = unitId; this.id = id }
  getChartId(): string { return this.id }
  getType(): ChartType { return this.model().type }
  getInfo(): ChartModel { return cloneModel(this.model()) }
  protected chartOptions(): ChartSemanticOptions { return cloneChartOptions(this.model().options) ?? {} }
  getAreaLineStyle(): ChartAreaLineStyle { return resolveChartAreaLineStyle(this.model()) }
  getInvalidValueStrategy(): ChartInvalidValueStrategy { return this.model().options?.invalidValueStrategy ?? CHART_INVALID_VALUE_STRATEGY_DEFAULT }
  /** Local live description API; upstream names describe() on its builder. */
  getDescription(): ChartDescription {
    const model = this.model()
    this.assertCanReadSource(model)
    return describeChartConfiguration(model, chartDescriptionSourceContext(this.workbook(), model))
  }
  getDataSource(): ChartSource { return cloneModel(this.model().source) }
  getData(): ReturnType<typeof readChartData> { const model = this.model(); this.assertCanReadSource(model); return readChartData(this.workbook(), model) }
  getLayout(): ReturnType<typeof chartLayout> {
    const model = this.model(); this.assertCanReadSource(model)
    const computed = this.api.chartStatisticsService.getCached(model, this.computationContext(model), this.workbook())?.statistics
    return cloneModel(chartLayout(this.workbook(), model, computed))
  }
  getRenderCommands(): ReturnType<typeof chartRenderCommands> {
    const model = this.model(); this.assertCanReadSource(model)
    const computed = this.api.chartStatisticsService.getCached(model, this.computationContext(model), this.workbook())?.statistics
    return cloneModel(chartRenderCommands(this.workbook(), model, computed))
  }
  async getComputedData(options: ChartComputationReadOptions = {}): Promise<ReturnType<typeof readChartData>> {
    const model = this.model(), context = this.computationContext(model, options)
    const snapshot = await this.api.chartStatisticsService.request(this.workbook(), model, context).promise
    context.assertCurrent?.()
    return cloneModel(snapshot.statistics.preparedData ?? readChartData(this.workbook(), model))
  }
  async getComputedLayout(options: ChartComputationReadOptions = {}): Promise<ReturnType<typeof chartLayout>> {
    const model = this.model(), context = this.computationContext(model, options)
    const snapshot = await this.api.chartStatisticsService.request(this.workbook(), model, context).promise
    context.assertCurrent?.()
    return cloneModel(chartLayout(this.workbook(), model, snapshot.statistics))
  }
  async getComputedRenderCommands(options: ChartComputationReadOptions = {}): Promise<ReturnType<typeof chartRenderCommands>> {
    const model = this.model(), context = this.computationContext(model, options)
    const snapshot = await this.api.chartStatisticsService.request(this.workbook(), model, context).promise
    context.assertCurrent?.()
    return cloneModel(chartRenderCommands(this.workbook(), model, snapshot.statistics))
  }
  async update(info: ChartPatch): Promise<this> {
    this.model()
    const committed = await this.api.executeCommand('sheet.command.update-chart', { unitId: this.unitId, chartId: this.id, patch: info })
    if (!committed) throw new Error('图表设置未能提交。')
    return this
  }
  async setDataSource(source: ChartSource): Promise<this> { return this.update({ source }) }
  async setType(type: ChartType): Promise<this> { resolveChartType(type); return this.update({ type }) }
  async setTitle(title: ChartTitleSpec | undefined): Promise<this> { const patch = chartTitlePatch(title); return Object.keys(patch).length ? this.update(patch) : this }
  async clearTitle(): Promise<this> { return this.setTitle(undefined) }
  async setSubtitle(subtitle: ChartSubtitleSpec | undefined): Promise<this> { const patch = chartTitlePatch(subtitle, 'subtitle'); return Object.keys(patch).length ? this.update(patch) : this }
  async clearSubtitle(): Promise<this> { return this.setSubtitle(undefined) }
  async setPalette(palette: readonly string[]): Promise<this> { return this.update({ palette: [...palette] }) }
  protected async updateChartOptions(options: ChartSemanticOptions): Promise<this> { return this.update({ options: cloneChartOptions(options) }) }
  async setOption<K extends keyof ChartSemanticOptions>(key: K, value: ChartSemanticOptions[K]): Promise<this> { return this.updateChartOptions(chartOptionPatch(key, value)) }
  async setAreaLineStyle(value: ChartAreaLineStyle): Promise<this> { return this.updateChartOptions(chartAreaLineStylePatch(this.model().type, chartAreaLineStyleValue(value))) }
  async clearAreaLineStyle(): Promise<this> { return this.updateChartOptions(chartAreaLineStylePatch(this.model().type, undefined)) }
  async setInvalidValueStrategy(value: ChartInvalidValueStrategy): Promise<this> { return this.setOption('invalidValueStrategy', chartInvalidValueStrategyValue(value)) }
  async resetInvalidValueStrategy(): Promise<this> { return this.setOption('invalidValueStrategy', undefined) }
  async setAutoGradientFill(value: boolean): Promise<this> { validateChartAutomaticGradientValue(value); return this.setOption('autoGradientFill', value) }
  async resetAutoGradientFill(): Promise<this> { return this.setOption('autoGradientFill', undefined) }
  async setPlotArea(value: ChartPlotAreaOptions): Promise<this> { validateChartPlotArea(value); return this.setOption('plotArea', value) }
  async resetPlotArea(): Promise<this> { return this.setOption('plotArea', undefined) }
  async setMapping(mapping: ChartSemanticOptions['mapping']): Promise<this> { return this.setOption('mapping', mapping) }
  async setUseDateAxis(value: boolean): Promise<this> { return this.setOption('mapping', { useDateAxis: value }) }
  async resetUseDateAxis(): Promise<this> { return this.setUseDateAxis(false) }
  async setAxisPointer(value: ChartNullablePatch<ChartAxisPointerOptions> | undefined): Promise<this> { return this.setOption('axisPointer', chartAxisPointerValues(value)) }
  async clearAxisPointer(target?: ChartAxisPointerTarget): Promise<this> { return this.setOption('axisPointer', clearChartAxisPointer(this.model().options?.axisPointer, target)) }
  async setBubbleMapping(mapping: ChartBubbleMappingOptions): Promise<this> { assertChartBubbleType(this.model().type); return this.setOption('mapping', { bubble: mapping }) }
  async clearBubbleMapping(): Promise<this> { assertChartBubbleType(this.model().type); return this.setOption('mapping', { bubble: undefined }) }
  async setHistogramField(index: number): Promise<this> { assertChartHistogramType(this.model().type); validateChartSeriesField(index); return this.setValueFields([index]) }
  async clearHistogramField(): Promise<this> { assertChartHistogramType(this.model().type); return this.setValueFields([]) }
  async setBinCount(count: number): Promise<this> { assertChartHistogramType(this.model().type); return isChartHistogramAuthoredNumber(count) ? this.setOption('histogram', { binning: { mode: 'count', count } }) : this }
  async setBinWidth(width: number): Promise<this> { assertChartHistogramType(this.model().type); return isChartHistogramAuthoredNumber(width) ? this.setOption('histogram', { binning: { mode: 'width', width } }) : this }
  async resetBinning(): Promise<this> { assertChartHistogramType(this.model().type); return this.setOption('histogram', { binning: undefined }) }
  async setUnderflowThreshold(value: number): Promise<this> { assertChartHistogramType(this.model().type); return isChartHistogramAuthoredNumber(value) ? this.setOption('histogram', { underflowThreshold: value }) : this }
  async clearUnderflowThreshold(): Promise<this> { assertChartHistogramType(this.model().type); return this.setOption('histogram', { underflowThreshold: undefined }) }
  async setOverflowThreshold(value: number): Promise<this> { assertChartHistogramType(this.model().type); return isChartHistogramAuthoredNumber(value) ? this.setOption('histogram', { overflowThreshold: value }) : this }
  async clearOverflowThreshold(): Promise<this> { assertChartHistogramType(this.model().type); return this.setOption('histogram', { overflowThreshold: undefined }) }
  async setBinGap(value: number): Promise<this> { assertChartHistogramType(this.model().type); return isChartHistogramAuthoredNumber(value) ? this.setOption('histogram', { binGap: value }) : this }
  async resetBinGap(): Promise<this> { assertChartHistogramType(this.model().type); return this.setOption('histogram', { binGap: undefined }) }
  async setHierarchyFields(indexes: readonly number[]): Promise<this> {
    const type = this.model().type, fields = type === 'sunburst' ? chartSunburstHierarchyFields(indexes)
      : (assertChartTreemapType(type), chartTreemapHierarchyFields(indexes))
    return this.setOption('mapping', { hierarchyIndexes: fields })
  }
  async setValueField(index: number): Promise<this> {
    const type = this.model().type, field = type === 'sunburst' ? chartSunburstValueField(index)
      : (assertChartTreemapType(type), chartTreemapValueField(index))
    return this.setValueFields([field])
  }
  async setParentLabelLayout(layout: NonNullable<ChartTreemapOptions['parentLabelLayout']>): Promise<this> { assertChartTreemapType(this.model().type); return this.setOption('treemap', { parentLabelLayout: chartTreemapParentLabelLayout(layout) }) }
  async resetParentLabelLayout(): Promise<this> { assertChartTreemapType(this.model().type); return this.setOption('treemap', { parentLabelLayout: undefined }) }
  async setTreemapLabel(label: ChartNullablePatch<ChartTreemapLabelOptions>): Promise<this> { assertChartTreemapType(this.model().type); return this.setOption('treemap', { label: chartTreemapLabelValues(label) }) }
  async resetTreemapLabel(): Promise<this> { assertChartTreemapType(this.model().type); return this.setOption('treemap', { label: undefined }) }
  async setSunburstLabel(label: ChartNullablePatch<ChartSunburstLabelOptions>): Promise<this> { assertChartSunburstType(this.model().type); return this.setOption('sunburst', { label: chartSunburstLabelValues(label) }) }
  async resetSunburstLabel(): Promise<this> { assertChartSunburstType(this.model().type); return this.setOption('sunburst', { label: undefined }) }
  async setCenterLabel(label: ChartNullablePatch<ChartSunburstCenterLabelOptions>): Promise<this> { assertChartSunburstType(this.model().type); return this.setOption('sunburst', { centerLabel: chartSunburstCenterLabelValues(label) }) }
  async resetCenterLabel(): Promise<this> { assertChartSunburstType(this.model().type); return this.setOption('sunburst', { centerLabel: undefined }) }
  async setChordMapping(mapping: ChartChordMapping): Promise<this> { assertChartChordType(this.model().type); return this.setOption('mapping', { chord: { ...chartChordMappingValues(mapping), valueIndex: mapping.valueIndex }, edge: undefined }) }
  async setChordLayout(layout: ChartNullablePatch<ChartChordLayoutOptions>): Promise<this> {
    const model = this.model(); assertChartChordType(model.type)
    return this.setOption('chord', chartChordLayoutPatch(model.options?.chord, layout))
  }
  async resetChordLayout(): Promise<this> {
    assertChartChordType(this.model().type)
    return this.setOption('chord', { layout: undefined, startAngle: undefined, clockwise: undefined,
      padAngle: undefined, minAngle: undefined, innerRadiusRatio: undefined, outerRadiusRatio: undefined })
  }
  async setChordNode(node: ChartNullablePatch<ChartChordNodeOptions>): Promise<this> { assertChartChordType(this.model().type); return this.setOption('chord', { node: chartChordNodeValues(node) }) }
  async resetChordNode(): Promise<this> { assertChartChordType(this.model().type); return this.setOption('chord', { node: undefined }) }
  async setChordLabel(label: ChartNullablePatch<ChartChordLabelOptions>): Promise<this> { assertChartChordType(this.model().type); return this.setOption('chord', { label: chartChordLabelValues(label) }) }
  async resetChordLabel(): Promise<this> { assertChartChordType(this.model().type); return this.setOption('chord', { label: undefined }) }
  async setChordRibbon(ribbon: ChartNullablePatch<ChartChordRibbonOptions>): Promise<this> { const model = this.model(); assertChartChordType(model.type); return this.setOption('chord', chartChordRibbonPatch(model.options?.chord, ribbon)) }
  async resetChordRibbon(): Promise<this> { assertChartChordType(this.model().type); return this.setOption('chord', { ribbon: undefined, ribbonOpacity: undefined }) }
  async setChordEmphasis(emphasis: ChartNullablePatch<ChartChordEmphasisOptions>): Promise<this> { assertChartChordType(this.model().type); return this.setOption('chord', { emphasis: chartChordEmphasisValues(emphasis) }) }
  async resetChordEmphasis(): Promise<this> { assertChartChordType(this.model().type); return this.setOption('chord', { emphasis: undefined }) }
  async setCandlestickFields(fields: ChartCandlestickMappingOptions): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('mapping', { candlestick: chartCandlestickFields(fields) }) }
  async clearCandlestickFields(): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('mapping', { candlestick: undefined }) }
  async setRisingStyle(style: ChartNullablePatch<ChartCandlestickStateStyleOptions>): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { rising: chartCandlestickStateStyleValues(style) }) }
  async clearRisingStyle(): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { rising: undefined }) }
  async setFallingStyle(style: ChartNullablePatch<ChartCandlestickStateStyleOptions>): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { falling: chartCandlestickStateStyleValues(style) }) }
  async clearFallingStyle(): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { falling: undefined }) }
  async setDojiStyle(style: ChartNullablePatch<ChartCandlestickStateStyleOptions>): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { doji: chartCandlestickStateStyleValues(style) }) }
  async clearDojiStyle(): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { doji: undefined }) }
  async setCandleWidth(width: number): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { candleWidth: chartCandleWidth(width) }) }
  async resetCandleWidth(): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { candleWidth: undefined }) }
  async setHighLowLineStyle(style: ChartNullablePatch<ChartCandlestickHighLowLineOptions>): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { highLowLine: chartCandlestickHighLowLineValues(style) }) }
  async clearHighLowLineStyle(): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { highLowLine: undefined }) }
  async setCloseMarkerStyle(style: ChartNullablePatch<ChartCandlestickCloseMarkerOptions>): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { closeMarker: chartCandlestickCloseMarkerValues(style) }) }
  async clearCloseMarkerStyle(): Promise<this> { assertChartCandlestickType(this.model().type); return this.setOption('candlestick', { closeMarker: undefined }) }
  async setAllSeriesStyle(style: ChartNullablePatch<ChartSeriesOptions> | undefined): Promise<this> {
    const next = chartSeriesStyleValues(style)
    validateChartGradientSeriesPatch(this.model().type, next, '所有系列')
    return this.setOption('allSeriesStyle', next)
  }
  async clearAllSeriesStyle(target?: ChartAllSeriesStyleTarget): Promise<this> { return this.setOption('allSeriesStyle', clearChartAllSeriesStyle(this.model().options?.allSeriesStyle, target)) }
  async setSeries(selector: ChartSeriesSelector, style: ChartNullablePatch<ChartSeriesOptions> | undefined): Promise<this> {
    const model = this.model(), next = chartSeriesStyleValues(style)
    validateChartGradientSeriesPatch(model.type, next)
    return this.updateChartOptions(patchChartSeriesOverride(model.options, selector, next))
  }
  async clearSeries(selector: ChartSeriesSelector, target?: ChartSeriesClearTarget): Promise<this> {
    validateChartSeriesSelector(selector)
    return this.updateChartOptions(patchChartSeriesOverride(this.model().options, selector, clearChartSeriesStyle(chartSeriesOverride(this.model().options, selector), target)))
  }
  async clearAllSeriesOverrides(): Promise<this> { return this.updateChartOptions({ series: {}, namedSeries: [] }) }
  async setPositiveStyle(selector: ChartSeriesSelector, value: ChartNullablePatch<ChartWaterfallSeriesStyle>): Promise<this> {
    const model = this.model(); assertChartWaterfallType(model.type)
    return this.updateChartOptions(patchChartWaterfallStyle(model.options, selector, 'positive', value))
  }
  async clearPositiveStyle(selector: ChartSeriesSelector): Promise<this> {
    const model = this.model(); assertChartWaterfallType(model.type)
    return this.updateChartOptions(patchChartWaterfallStyle(model.options, selector, 'positive', undefined))
  }
  async setNegativeStyle(selector: ChartSeriesSelector, value: ChartNullablePatch<ChartWaterfallSeriesStyle>): Promise<this> {
    const model = this.model(); assertChartWaterfallType(model.type)
    return this.updateChartOptions(patchChartWaterfallStyle(model.options, selector, 'negative', value))
  }
  async clearNegativeStyle(selector: ChartSeriesSelector): Promise<this> {
    const model = this.model(); assertChartWaterfallType(model.type)
    return this.updateChartOptions(patchChartWaterfallStyle(model.options, selector, 'negative', undefined))
  }
  async setSubtotalStyle(selector: ChartSeriesSelector, value: ChartNullablePatch<ChartWaterfallSeriesStyle>): Promise<this> {
    const model = this.model(); assertChartWaterfallType(model.type)
    return this.updateChartOptions(patchChartWaterfallStyle(model.options, selector, 'subtotal', value))
  }
  async clearSubtotalStyle(selector: ChartSeriesSelector): Promise<this> {
    const model = this.model(); assertChartWaterfallType(model.type)
    return this.updateChartOptions(patchChartWaterfallStyle(model.options, selector, 'subtotal', undefined))
  }
  async setConnector(value: boolean | ChartNullablePatch<ChartWaterfallConnectorOptions>): Promise<this> {
    assertChartWaterfallType(this.model().type)
    return this.setOption('waterfall', { connector: chartWaterfallConnectorValues(value) })
  }
  async clearConnector(): Promise<this> { assertChartWaterfallType(this.model().type); return this.setOption('waterfall', { connector: undefined }) }
  async setStackType(value: ChartWaterfallStackType): Promise<this> {
    assertChartWaterfallType(this.model().type); validateChartWaterfallStackType(value)
    return this.setOption('waterfall', { stackType: value, ...(value === 'stacked' ? { pointRoles: undefined } : {}) })
  }
  async resetStackType(): Promise<this> { assertChartWaterfallType(this.model().type); return this.setOption('waterfall', { stackType: undefined }) }
  async setUseSubtotal(value: boolean): Promise<this> {
    assertChartWaterfallType(this.model().type); validateChartWaterfallSubtotalValue(value)
    return this.setOption('waterfall', { useSubtotal: value })
  }
  async resetUseSubtotal(): Promise<this> { assertChartWaterfallType(this.model().type); return this.setOption('waterfall', { useSubtotal: undefined }) }
  async setBarStyle(value: ChartNullablePatch<ChartParetoBarStyleOptions>): Promise<this> {
    assertChartParetoType(this.model().type)
    return this.setOption('pareto', { barStyle: chartParetoBarStyleValues(value) })
  }
  async clearBarStyle(): Promise<this> { assertChartParetoType(this.model().type); return this.setOption('pareto', { barStyle: undefined }) }
  async setCumulativeLineStyle(value: ChartNullablePatch<ChartParetoLineStyleOptions>): Promise<this> {
    assertChartParetoType(this.model().type)
    return this.setOption('pareto', { cumulativeLineStyle: chartParetoLineStyleValues(value) })
  }
  async clearCumulativeLineStyle(): Promise<this> { assertChartParetoType(this.model().type); return this.setOption('pareto', { cumulativeLineStyle: undefined }) }
  async setIncludeZeroValues(value: boolean): Promise<this> {
    assertChartParetoType(this.model().type)
    if (typeof value !== 'boolean') throw new TypeError('是否包含零合计必须是布尔值。')
    return this.setOption('pareto', { includeZeroValues: value })
  }
  async resetIncludeZeroValues(): Promise<this> { assertChartParetoType(this.model().type); return this.setOption('pareto', { includeZeroValues: undefined }) }
  async setVisualMapType(value: ChartHeatmapVisualMapType): Promise<this> {
    assertChartHeatmapType(this.model().type)
    return this.setOption('heatmap', { visualMapType: chartHeatmapVisualMapType(value) })
  }
  async resetVisualMapType(): Promise<this> { assertChartHeatmapType(this.model().type); return this.setOption('heatmap', { visualMapType: undefined }) }
  async setValueRange(min: number, max: number): Promise<this> {
    assertChartHeatmapType(this.model().type)
    return this.setOption('heatmap', chartHeatmapValueRange(min, max))
  }
  async clearValueRange(): Promise<this> { assertChartHeatmapType(this.model().type); return this.setOption('heatmap', { min: undefined, max: undefined }) }
  async setCellLabel(value: ChartNullablePatch<ChartHeatmapLabelOptions>): Promise<this> {
    assertChartHeatmapType(this.model().type)
    return this.setOption('heatmap', { label: chartHeatmapCellLabelValues(value) })
  }
  async clearCellLabel(): Promise<this> { assertChartHeatmapType(this.model().type); return this.setOption('heatmap', { label: undefined }) }
  async setValueUnit(value: string): Promise<this> {
    assertChartHeatmapType(this.model().type)
    return this.setOption('heatmap', { unit: chartHeatmapValueUnit(value) })
  }
  async clearValueUnit(): Promise<this> { assertChartHeatmapType(this.model().type); return this.setOption('heatmap', { unit: undefined }) }
  async setTrendlines(trendlines: readonly ChartTrendlineOptions[]): Promise<this> { return this.setOption('trendlines', [...trendlines]) }
  async setErrorBars(errorBars: readonly ChartErrorBarOptions[]): Promise<this> { return this.setOption('errorBars', [...errorBars]) }
  async setRightYAxis(axis: ChartNullablePatch<ChartAxisOptions> | undefined): Promise<this> { return this.setOption('rightYAxis', axis === undefined ? undefined : chartAxisPatch(this.model().options?.rightYAxis, axis)) }
  async clearRightYAxis(target?: ChartAxisTarget): Promise<this> { return this.setOption('rightYAxis', clearChartAxis(this.model().options?.rightYAxis, target)) }
  async setRadar(value: ChartNullablePatch<ChartRadarOptions> | undefined): Promise<this> { const model = this.model(); assertChartRadarType(model.type); return this.setOption('radar', patchChartRadarOptions(model.options?.radar, value)) }
  async setShape(value: ChartRadarShape | ChartWordCloudShape): Promise<this> {
    if (this.model().type === 'wordCloud') return this.setOption('wordCloud', { shape: chartWordCloudShape(value) })
    if (value !== 'circle' && value !== 'polygon') throw new TypeError('雷达图形状无效。')
    return this.setRadar({ shape: value })
  }
  async resetShape(): Promise<this> { return this.model().type === 'wordCloud' ? this.setOption('wordCloud', { shape: undefined }) : this.setRadar({ shape: undefined }) }
  async setMaskImage(resourceId: string): Promise<this> { assertChartWordCloudType(this.model().type); return this.setOption('wordCloud', { maskImageId: chartWordCloudMaskImageId(resourceId), maskImage: undefined, maskImageResource: undefined }) }
  async setMaskImageResource(value: ChartImageResource): Promise<this> { assertChartWordCloudType(this.model().type); const resource = normalizeChartImageResource(value); return this.setOption('wordCloud', { maskImageId: resource.id, maskImage: undefined, maskImageResource: resource }) }
  async clearMaskImage(): Promise<this> { assertChartWordCloudType(this.model().type); return this.setOption('wordCloud', { maskImageId: undefined, maskImage: undefined, maskImageResource: undefined }) }
  async setRepeat(value: boolean): Promise<this> { assertChartWordCloudType(this.model().type); return this.setOption('wordCloud', { repeat: chartWordCloudRepeat(value) }) }
  async resetRepeat(): Promise<this> { assertChartWordCloudType(this.model().type); return this.setOption('wordCloud', { repeat: undefined }) }
  async setLayout(value: ChartRelationLayout): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { layout: chartRelationLayoutValue(value) }) }
  async resetLayout(): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { layout: undefined }) }
  async setUseValueAsSymbolSize(value: boolean): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { useValueAsSymbolSize: chartRelationBoolean(value) }) }
  async resetUseValueAsSymbolSize(): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { useValueAsSymbolSize: undefined }) }
  async setEmphasisEnabled(value: boolean): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { emphasisEnabled: chartRelationBoolean(value) }) }
  async resetEmphasisEnabled(): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { emphasisEnabled: undefined }) }
  async setForce(value: ChartRelationForceOptions): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { force: chartRelationForce(value) }) }
  async clearForce(): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { force: undefined }) }
  async setCircularLabelRotation(value: boolean): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { circularLabelRotation: chartRelationBoolean(value) }) }
  async resetCircularLabelRotation(): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { circularLabelRotation: undefined }) }
  async setNodeShape(value: ChartPointShape): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { nodeShape: chartRelationNodeShape(value) }) }
  async resetNodeShape(): Promise<this> { assertChartRelationType(this.model().type); return this.setOption('relation', { nodeShape: undefined }) }
  async setFill(value: boolean): Promise<this> { return this.setRadar({ fill: value }) }
  async resetFill(): Promise<this> { return this.setRadar({ fill: undefined }) }
  async setRadarAxisLabel(value: ChartNullablePatch<ChartRadarAxisLabelOptions>): Promise<this> { return this.setRadar({ axisLabel: value }) }
  async resetRadarAxisLabel(): Promise<this> { return this.setRadar({ axisLabel: undefined }) }
  async setGauge(value: ChartGaugeOptions | undefined): Promise<this> { return this.setOption('gauge', patchChartGaugeOptions(this.model().options?.gauge, value)) }
  async setGaugePreset(preset: ChartGaugePreset): Promise<this> { return this.setGauge(chartGaugePresetOptions(preset)) }
  async setGaugeValueField(index: number): Promise<this> {
    const model = this.model(), mapping = chartGaugeValueFieldMapping(model.options?.mapping, index, model.categoryField)
    return this.update({ valueFields: [index], options: { mapping } })
  }
  async setGaugeMode(value: ChartGaugeMode): Promise<this> { return this.setGauge({ mode: value }) }
  async resetGaugeMode(): Promise<this> { return this.setGauge({ mode: undefined }) }
  async setGaugeScale(value: ChartGaugeNullablePatch<ChartGaugeScaleOptions>): Promise<this> { return this.setGauge(chartGaugeNullableValues<ChartGaugeScaleOptions>(value)) }
  async resetGaugeScale(): Promise<this> { return this.setGauge({ min: undefined, max: undefined, splitNumber: undefined }) }
  async setGaugeTicks(value: ChartGaugeTicksPatch): Promise<this> { return this.setGauge(chartGaugeTicksOptions(value)) }
  async setGaugeAngles(value: ChartGaugeNullablePatch<ChartGaugeAnglesOptions>): Promise<this> { return this.setGauge(chartGaugeNullableValues<ChartGaugeAnglesOptions>(value)) }
  async resetGaugeAngles(): Promise<this> { return this.setGauge({ startAngle: undefined, endAngle: undefined }) }
  async setGaugeRanges(value: readonly ChartGaugeRangeOptions[]): Promise<this> { return this.setGauge(chartGaugeRangesOptions(value)) }
  async resetGaugeRanges(): Promise<this> { return this.setGauge({ ranges: undefined }) }
  async setGaugeAxisLine(value: ChartGaugeNullablePatch<ChartGaugeAxisLineOptions>): Promise<this> { return this.setGauge({ axisLine: chartGaugeNullableValues<ChartGaugeAxisLineOptions>(value) }) }
  async resetGaugeAxisLine(): Promise<this> { return this.setGauge({ axisLine: undefined }) }
  async setGaugeAxisTick(value: ChartGaugeNullablePatch<ChartGaugeAxisTickOptions>): Promise<this> { return this.setGauge({ axisTick: chartGaugeNullableValues<ChartGaugeAxisTickOptions>(value) }) }
  async resetGaugeAxisTick(): Promise<this> { return this.setGauge({ axisTick: undefined }) }
  async setGaugeSplitLine(value: ChartGaugeNullablePatch<ChartGaugeSplitLineOptions>): Promise<this> { return this.setGauge({ splitLine: chartGaugeNullableValues<ChartGaugeSplitLineOptions>(value) }) }
  async resetGaugeSplitLine(): Promise<this> { return this.setGauge({ splitLine: undefined }) }
  async setGaugeAxisLabel(value: ChartGaugeNullablePatch<ChartGaugeAxisLabelOptions>): Promise<this> { return this.setGauge({ axisLabel: chartGaugeNullableValues<ChartGaugeAxisLabelOptions>(value) }) }
  async resetGaugeAxisLabel(): Promise<this> { return this.setGauge({ axisLabel: undefined }) }
  async setGaugePointer(value: ChartGaugeNullablePatch<ChartGaugePointerOptions>): Promise<this> { return this.setGauge({ pointer: chartGaugeNullableValues<ChartGaugePointerOptions>(value) }) }
  async resetGaugePointer(): Promise<this> { return this.setGauge({ pointer: undefined }) }
  async setGaugeAnchor(value: ChartGaugeNullablePatch<ChartGaugeAnchorOptions>): Promise<this> { return this.setGauge({ anchor: chartGaugeNullableValues<ChartGaugeAnchorOptions>(value) }) }
  async resetGaugeAnchor(): Promise<this> { return this.setGauge({ anchor: undefined }) }
  async setGaugeProgress(value: ChartGaugeNullablePatch<ChartGaugeProgressOptions>): Promise<this> { return this.setGauge({ progress: chartGaugeNullableValues<ChartGaugeProgressOptions>(value) }) }
  async resetGaugeProgress(): Promise<this> { return this.setGauge({ progress: undefined }) }
  async setGaugeTitle(value: ChartGaugeNullablePatch<ChartGaugeTitleOptions>): Promise<this> { return this.setGauge({ title: chartGaugeNullableValues<ChartGaugeTitleOptions>(value) }) }
  async resetGaugeTitle(): Promise<this> { return this.setGauge({ title: undefined }) }
  async setGaugeDetail(value: ChartGaugeNullablePatch<ChartGaugeDetailOptions>): Promise<this> { return this.setGauge({ detail: chartGaugeNullableValues<ChartGaugeDetailOptions>(value) }) }
  async resetGaugeDetail(): Promise<this> { return this.setGauge({ detail: undefined }) }
  async setLegend(legend: ChartLegendSpec): Promise<this> { return this.update({ legend: { ...this.model().legend, ...chartLegendPatch(this.model().legend, legend) } }) }
  async clearLegend(): Promise<this> { return this.update({ legend: { ...defaultChartLegend(this.model().type), font: undefined, selectMode: undefined, wrap: undefined } }) }
  async setXAxis(axis: ChartNullablePatch<ChartAxisOptions>): Promise<this> { return this.update({ xAxis: chartAxisPatch(this.model().xAxis, axis) }) }
  async setYAxis(axis: ChartNullablePatch<ChartAxisOptions>): Promise<this> { return this.update({ yAxis: chartAxisPatch(this.model().yAxis, axis) }) }
  async clearXAxis(target?: ChartAxisTarget): Promise<this> { return this.update({ xAxis: clearChartAxis(this.model().xAxis, target) }) }
  async clearYAxis(target?: ChartAxisTarget): Promise<this> { return this.update({ yAxis: clearChartAxis(this.model().yAxis, target) }) }
  async setXAxisTitle(title: string): Promise<this> { return this.setXAxis({ title }) }
  async clearXAxisTitle(): Promise<this> { return this.setXAxis({ title: undefined }) }
  async setYAxisTitle(title: string): Promise<this> { return this.setYAxis({ title }) }
  async clearYAxisTitle(): Promise<this> { return this.setYAxis({ title: undefined }) }
  async setRightYAxisTitle(title: string): Promise<this> { return this.setOption('rightYAxis', chartAxisPatch(this.model().options?.rightYAxis, { title })) }
  async clearRightYAxisTitle(): Promise<this> { return this.setOption('rightYAxis', chartAxisPatch(this.model().options?.rightYAxis, { title: undefined })) }
  async setAbsolutePosition(x: number, y: number): Promise<this> { return this.update({ position: { x, y } }) }
  async setSize(width: number, height: number): Promise<this> { return this.update({ size: { width, height } }) }
  async remove(): Promise<boolean> { return this.api.executeCommand('sheet.command.remove-chart', { unitId: this.unitId, chartId: this.id }) }
  protected model(): ChartModel {
    const chart = this.workbook().charts?.find((item) => item.id === this.id)
    if (!chart) throw new Error(`Chart not found: ${this.id}`)
    this.api.assertCanViewWorksheet(this.unitId, chart.sheetId)
    return chart
  }
  protected assertCanReadSource(chart: ChartModel): void {
    assertChartReadPermission(this.workbook(), chart, this.api.currentUserId)
    const sourceSheetId = chart.pivotTableId ? this.workbook().pivotTables?.find(pivot => pivot.id === chart.pivotTableId)?.source.subUnitId : chart.source.sheetId ?? chart.sheetId
    if (sourceSheetId) this.api.assertCanViewWorksheet(this.unitId, sourceSheetId)
  }
  private computationContext(chart: ChartModel, options: ChartComputationReadOptions = {}): ChartStatisticsRequestContext {
    this.assertCanReadSource(chart)
    const modelRevision = this.api.getWorkbookRevision(this.unitId), signature = chartStatisticsSignature(chart)
    const signal = options.signal, userId = this.api.currentUserId
    return { ...options, documentId: this.unitId, modelRevision, userId,
      assertCurrent: () => {
        if (signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '图表计算读取已取消')
        const workbook = this.api.getWorkbookModel(this.unitId), current = workbook?.charts?.find(candidate => candidate.id === this.id)
        if (!workbook || !current || this.api.currentUserId !== userId || this.api.getWorkbookRevision(this.unitId) !== modelRevision || chartStatisticsSignature(current) !== signature) {
          throw new ComputeError('COMPUTE_STALE_RESULT', '图表或来源已在计算期间更新')
        }
        this.assertCanReadSource(current)
      } }
  }
  protected workbook(): WorkbookModel {
    const workbook = this.api.getWorkbookModel(this.unitId)
    if (!workbook) throw new Error(`Workbook not found: ${this.unitId}`)
    return workbook
  }
}

export class FChart extends ChartVisualFacade {
  getOptions(): ChartSemanticOptions { return this.chartOptions() }
  async setOptions(options: ChartSemanticOptions): Promise<this> { return this.updateChartOptions(options) }
}

/** A Pivot report facade over the same drawing, chart engine and content history. */
export class FSheetPivotChart extends ChartVisualFacade {
  getId(): string { return this.id }
  getUnitId(): string { return this.unitId }
  getSubUnitId(): string { return this.model().sheetId }
  getSourceKind(): 'range' | 'pivotTable' { return this.model().ownedPivot ? 'range' : 'pivotTable' }
  isLinkedToPivotTable(): boolean { return this.getSourceKind() === 'pivotTable' }
  getSourceRangeInfo(): PivotTableSourceInput {
    const report = this.reportModel(), source = report.source
    return { unitId: source.unitId, subUnitId: source.subUnitId, sheetName: source.sheetName,
      ...(report.pivotCacheId ? { pivotCacheId: report.pivotCacheId } : {}),
      range: { startRow: source.range.startRow, startColumn: source.range.startCol, endRow: source.range.endRow, endColumn: source.range.endCol } }
  }
  getPivotTableReference(): PivotChartTableReference | undefined {
    const chart = this.model()
    if (!chart.pivotTableId) return undefined
    const report = this.reportModel()
    return { unitId: this.unitId, subUnitId: report.anchor.subUnitId, pivotTableId: report.id }
  }
  getPivotTable(): FPivotTable | null {
    const reference = this.getPivotTableReference()
    return reference ? new FPivotTable(this.api, this.unitId, reference.pivotTableId) : null
  }
  getFieldsConfig(): PivotChartFieldsConfig {
    const report = this.reportModel()
    return cloneModel({ fields: report.fields, options: report.options, valuePosition: report.options.valuePosition, valueIndex: report.options.valueIndex ?? -1,
      ...(report.dateGroupingMode ? { dateGroupingMode: report.dateGroupingMode } : {}),
      ...(report.collapseInfo ? { collapseInfo: report.collapseInfo } : {}) })
  }
  getSourceFieldsInfo(): Array<{ id: string; name: string; fieldId: string; fieldName: string; sourceColumn: number }> {
    return getPivotSourceFields(this.workbook(), this.reportModel()).map(field => ({ ...field, id: field.fieldId, name: field.fieldName }))
  }
  getFieldItems(fieldId: string): string[] {
    const chart = this.model()
    this.assertCanReadSource(chart)
    return getPivotChartFieldItems(this.workbook(), chart, fieldId)
  }
  getFieldSetting(fieldId: string): PivotFieldModel | undefined {
    const chart = this.model()
    this.assertCanReadSource(chart)
    const field = getPivotChartConfiguredField(this.workbook(), chart, fieldId)
    return field ? cloneModel(field) : undefined
  }
  getFieldIdsByArea(area: PivotFieldArea): string[] { return pivotFieldIds(this.reportModel(), area) }
  getOptions(): PivotTableOptions { return cloneModel(this.reportModel().options) }
  getReportOptions(): PivotTableOptions { return cloneModel(this.reportModel().options) }
  getChartOptions(): ChartSemanticOptions { return this.chartOptions() }
  getControls(): PivotChartControls { return pivotChartControls(this.model()) }
  getManualFilterItemState(fieldId: string): { checkedItems: string[]; partialItems: string[] } {
    return getPivotManualFilterItemState(this.workbook(), this.reportModel(), fieldId)
  }
  getFilterCapabilities(fieldId: string): { supportsDateFilter: boolean; date1904: boolean } {
    return getPivotFieldFilterCapabilities(this.workbook(), this.reportModel(), fieldId)
  }
  getDateFilter(fieldId: string): PivotDateFilter | undefined { return getPivotDateFilter(this.reportModel(), fieldId) }
  getLabelConditionFilter(fieldId: string): PivotLabelConditionFilter | undefined { return getPivotLabelConditionFilter(this.reportModel(), fieldId) }
  async setControls(controls: Partial<PivotChartControls>, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_COMMANDS.setControls, this.commandTarget(context, { controls }))
  }
  getValueFilter(fieldId: string): PivotFieldFilter | undefined { return this.getFieldSetting(fieldId)?.valueFilter }
  getValueFilters(): Array<{ fieldId: string; filter: PivotFieldFilter }> {
    return this.reportModel().fields.filter(field => field.valueFilter).map(field => ({ fieldId: field.id, filter: cloneModel(field.valueFilter!) }))
  }
  async applyReportEdit(edit: PivotChartReportEdit, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_COMMANDS.edit, this.commandTarget(context, { edit }))
  }
  async configure(input: PivotChartConfigureInput, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_COMMANDS.configure, this.commandTarget(context, input))
  }
  override async update(info: ChartPatch): Promise<this> {
    if (!await this.configure({ patch: info })) throw new Error('透视图设置未能提交。')
    return this
  }
  async setFieldsConfig(config: PivotChartFieldsConfig, context: PivotChartCommandContext = {}): Promise<boolean> { return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setFieldsConfig, this.commandTarget(context, { pivotTableConfig: config })) }
  async resetFields(area?: PivotFieldArea, context: PivotChartCommandContext = {}): Promise<boolean> { return this.applyReportEdit({ type: 'resetFields', area }, context) }
  async addField(dataFieldId: string | number, fieldArea: PivotFieldArea, index?: number, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.add, this.commandTarget(context, { dataFieldId, fieldArea, index }))
  }
  async removeField(fieldIds: string[], context: PivotChartCommandContext = {}): Promise<boolean> { return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.remove, this.commandTarget(context, { fieldIds })) }
  async updateFieldPosition(fieldId: string, fieldArea: PivotFieldArea, index: number, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.move, this.commandTarget(context, { fieldId, fieldArea, index }))
  }
  async setOptions(options: Partial<PivotTableOptions>, context: PivotChartCommandContext = {}): Promise<boolean> { return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setOptions, this.commandTarget(context, { options })) }
  async setReportOptions(options: Partial<PivotTableOptions>, context: PivotChartCommandContext = {}): Promise<boolean> { return this.setOptions(options, context) }
  async setChartOptions(options: ChartSemanticOptions): Promise<this> { return this.updateChartOptions(options) }
  async setFieldSetting(fieldId: string, setting: PivotChartFieldSetting, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setFieldSetting, this.commandTarget(context, { fieldId, setting }))
  }
  async renameField(fieldId: string, displayName: string, context: PivotChartCommandContext = {}): Promise<boolean> { return this.setFieldSetting(fieldId, { displayName }, context) }
  async setSubtotalType(fieldId: string, subtotalType: PivotAggregation, context: PivotChartCommandContext = {}): Promise<boolean> { return this.setFieldSetting(fieldId, { subtotalType }, context) }
  getShowDataAs(fieldId: string): PivotShowDataAsInfo | undefined { return this.getFieldSetting(fieldId)?.showDataAs }
  getShowDataAsBaseItems(baseFieldId: string): string[] {
    const chart = this.model()
    this.assertCanReadSource(chart)
    return getPivotChartShowDataAsBaseItems(this.workbook(), chart, baseFieldId)
  }
  async setShowDataAs(fieldId: string, showDataAs: PivotShowDataAsInfo | null, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.applyReportEdit({ type: 'setShowDataAs', fieldId, showDataAs }, context)
  }
  async resetShowDataAs(fieldId: string, context: PivotChartCommandContext = {}): Promise<boolean> { return this.setShowDataAs(fieldId, null, context) }
  getAxisState(area: PivotAxisArea = 'row'): PivotAxisState {
    const chart = this.model()
    this.assertCanReadSource(chart)
    return cloneModel(getPivotChartAxisState(this.workbook(), chart, area))
  }
  getCollapse(fieldId: string, item?: string): boolean {
    const report = this.reportModel(), field = getPivotChartConfiguredField(this.workbook(), this.model(), fieldId)
    if (!field || field.area !== 'row') throw new Error('透视图展开折叠需要行字段。')
    return getPivotCollapse(report, field.id, item)
  }
  getCollapseItems(fieldId: string): string[] {
    const chart = this.model()
    this.assertCanReadSource(chart)
    return getPivotChartCollapseItems(this.workbook(), chart, fieldId)
  }
  async setCollapse(fieldId: string, collapse: boolean, item?: string, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_HIERARCHY_COMMANDS.chartSetCollapse, { ...this.commandTarget(context), fieldId, collapse, item })
  }
  async setCategoryLevel(fieldId: string, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_HIERARCHY_COMMANDS.chartSetCategoryLevel, { ...this.commandTarget(context), fieldId })
  }
  async stepCategoryLevel(direction: 'expand' | 'collapse', context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_HIERARCHY_COMMANDS.chartStepCategoryLevel, { ...this.commandTarget(context), direction })
  }
  async setLabelSort(fieldId: string, info: { type: 'ascending' | 'descending' }, context: PivotChartCommandContext = {}): Promise<boolean> { return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setSort, this.commandTarget(context, { fieldId, info })) }
  async clearLabelSort(fieldId: string, context: PivotChartCommandContext = {}): Promise<boolean> { return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setSort, this.commandTarget(context, { fieldId })) }
  async setLabelManualFilter(fieldId: string, items: string[] | null, isAll = false, context: PivotChartCommandContext = {}, partialItems?: string[]): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setFilter, this.commandTarget(context, { fieldId, items, isAll, partialItems }))
  }
  async clearFilters(fieldId: string, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setFilter, this.commandTarget(context, { fieldId, items: [], clearAllFilters: true }))
  }
  async setDateFilter(fieldId: string, filterInfo?: PivotDateFilterInput | null, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setDateFilter, this.commandTarget(context, { fieldId, filterInfo }))
  }
  async clearDateFilter(fieldId: string, context: PivotChartCommandContext = {}): Promise<boolean> { return this.setDateFilter(fieldId, undefined, context) }
  async setLabelConditionFilter(fieldId: string, filterInfo?: PivotLabelConditionInput | null, context: PivotChartCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setLabelConditionFilter, this.commandTarget(context, { fieldId, filterInfo }))
  }
  async clearLabelConditionFilter(fieldId: string, context: PivotChartCommandContext = {}): Promise<boolean> { return this.setLabelConditionFilter(fieldId, undefined, context) }
  async setValueFilter(fieldId: string, filter: PivotFieldFilter | null, context: PivotChartCommandContext = {}): Promise<boolean> { return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setValueFilter, this.commandTarget(context, { fieldId, valueFilterInfo: filter })) }
  async clearValueFilter(fieldId: string, context: PivotChartCommandContext = {}): Promise<boolean> { return this.setValueFilter(fieldId, null, context) }
  async setNumberGroup(fieldId: string, group: PivotNumberGroup | null, context: PivotChartCommandContext = {}): Promise<boolean> { return this.applyReportEdit({ type: 'setNumberGroup', fieldId, group }, context) }
  async setDateGroupingMode(mode: PivotDateGroupingMode, context: PivotChartCommandContext = {}): Promise<boolean> { return this.applyReportEdit({ type: 'setDateGroupingMode', mode }, context) }
  async setDateGroupType(fieldId: string, dateType: PivotDateGroupType | null, context: PivotChartCommandContext = {}): Promise<boolean> { return this.api.executeCommand(PIVOT_CHART_FIELD_COMMANDS.setDateGroup, this.commandTarget(context, { fieldId, dateType })) }
  async setExcelDateGroup(fieldId: string, dateType: PivotDateGroupType, context: PivotChartCommandContext = {}): Promise<boolean> { return this.applyReportEdit({ type: 'setDateGroup', fieldId, dateType, excelLike: true }, context) }
  async setManualGroups(fieldId: string, groups: PivotManualGroup[], context: PivotChartCommandContext = {}): Promise<boolean> { return this.applyReportEdit({ type: 'setManualGroups', fieldId, groups }, context) }
  getValuesPosition(): { position: PivotValuePosition; index: number } { return getPivotValuesPosition(this.reportModel()) }
  async updateValuePosition(position: PivotValuePositionInput, indexOrContext: number | PivotChartCommandContext = -1,
    context: PivotChartCommandContext = {}): Promise<boolean> {
    const index = typeof indexOrContext === 'number' ? indexOrContext : -1
    const target = typeof indexOrContext === 'number' ? context : indexOrContext
    const state = normalizePivotValuesPosition(normalizePivotValuePositionInput(position), index)
    return this.applyReportEdit({ type: 'updateValuePosition', position: state.position, index: state.index }, target)
  }
  async refresh(context: PivotChartCommandContext = {}): Promise<boolean> { return this.api.executeCommand(PIVOT_CHART_COMMANDS.refresh, this.commandTarget(context)) }
  override async remove(context: PivotChartCommandContext = {}): Promise<boolean> { return this.api.executeCommand(PIVOT_CHART_COMMANDS.remove, this.commandTarget(context)) }
  private reportModel(): PivotTableModel {
    const chart = this.model()
    this.assertCanReadSource(chart)
    return getPivotChartReport(this.workbook(), chart)
  }
  private commandTarget(context: PivotChartCommandContext, extra: object = {}): object {
    const chart = this.model()
    return { ...context, ...extra, unitId: this.unitId, sheetId: chart.sheetId, pivotChartId: this.id }
  }
}

/** Visual configuration delegates to the existing Chart builder. Report source remains typed. */
export class PivotChartConfigBuilder {
  private readonly visual: ChartConfigBuilder
  private source?: PivotChartSourceSpec
  private fields?: PivotChartFieldsConfig
  constructor(private readonly api: OfficeStudio, private readonly unitId: string, private readonly sheetId: string, type: ChartType) { this.visual = new ChartConfigBuilder(type) }
  setSource(input: string | FRange | FRangeCoordinates | PivotTableSourceInput): this {
    if (typeof input === 'object' && !(input instanceof FRange) && 'subUnitId' in input) {
      if (input.unitId !== this.unitId) throw new Error('透视图来源必须属于同一工作簿。')
      this.source = { kind: 'range', sourceRangeInfo: cloneModel(input) }
      return this
    }
    const range = input instanceof FRange ? input : typeof input === 'string'
      ? new FWorksheet(this.api, this.unitId, this.sheetId).getRange(input) : new FWorksheet(this.api, this.unitId, this.sheetId).getRange(input)
    if (range.getUnitId() !== this.unitId) throw new Error('透视图来源必须属于同一工作簿。')
    const sheet = this.api.getWorkbookModel(this.unitId)?.sheets.find(candidate => candidate.id === range.getSheetId())
    if (!sheet) throw new Error('透视图来源工作表不存在。')
    this.source = { kind: 'range', sourceRangeInfo: { unitId: this.unitId, subUnitId: sheet.id, sheetName: sheet.name, range: range.getRange() } }
    return this
  }
  setPivotTable(pivot: FPivotTable): this { return this.setPivotTableReference({ unitId: pivot.getUnitId(), subUnitId: pivot.getSubUnitId(), pivotTableId: pivot.getPivotTableId() }) }
  setPivotTableReference(reference: PivotChartTableReference): this {
    if (reference.unitId !== this.unitId) throw new Error('关联透视表必须属于同一工作簿。')
    this.source = { kind: 'pivotTable', pivotTableRef: cloneModel(reference) }
    return this
  }
  setInitialFieldsConfig(fields: PivotChartFieldsConfig): this { this.fields = cloneModel(fields); return this }
  setType(type: ChartType): this { this.visual.setType(type); return this }
  setPosition(cell: string): this { this.visual.setPosition(cell); return this }
  setAbsolutePosition(x: number, y: number): this { this.visual.setAbsolutePosition(x, y); return this }
  setSize(width: number, height: number): this { this.visual.setSize(width, height); return this }
  setTitle(title: ChartTitleSpec | undefined): this { this.visual.setTitle(title); return this }
  clearTitle(): this { this.visual.clearTitle(); return this }
  setSubtitle(title: ChartSubtitleSpec | undefined): this { this.visual.setSubtitle(title); return this }
  clearSubtitle(): this { this.visual.clearSubtitle(); return this }
  setLegend(legend: ChartLegendSpec): this { this.visual.setLegend(legend); return this }
  clearLegend(): this { this.visual.clearLegend(); return this }
  setPalette(palette: readonly string[]): this { this.visual.setPalette(palette); return this }
  setVisualOptions(options: ChartSemanticOptions): this { this.visual.setOptions(options); return this }
  build(): PivotChartInfo {
    if (!this.source) throw new TypeError('请选择独立范围或关联透视表来源。')
    const { source: _source, pivotTableId: _pivotTableId, ownedPivot: _ownedPivot, ...visual } = this.visual.build()
    return { ...visual, source: cloneModel(this.source), ...(this.fields ? { initialFieldsConfig: cloneModel(this.fields) } : {}) }
  }
}

function chartOptionPatch<K extends keyof ChartSemanticOptions>(key: K, value: ChartSemanticOptions[K]): ChartSemanticOptions {
  const patch: ChartSemanticOptions = {}
  Object.assign(patch, { [key]: value })
  return cloneChartOptions(patch)!
}
function validateChartSeriesField(field: number): void { if (!Number.isSafeInteger(field) || field < 0) throw new RangeError('图表系列字段须为非负安全整数。') }
function validateChartAutomaticGradientValue(value: boolean): void { if (typeof value !== 'boolean') throw new TypeError('自动渐变设置须为布尔值。') }
function validateChartWaterfallStackType(value: ChartWaterfallStackType): void {
  if (value !== 'arrangement' && value !== 'stacked') throw new TypeError('瀑布图排列方式无效。')
}
function validateChartWaterfallSubtotalValue(value: boolean): void { if (typeof value !== 'boolean') throw new TypeError('瀑布图自动小计设置须为布尔值。') }
function assertChartRadarType(type: ChartType): void { if (type !== 'radar') throw new TypeError('雷达专属 API 仅适用于雷达图。') }
function assertChartBubbleType(type: ChartType): void { if (type !== 'bubble') throw new TypeError('气泡映射 API 仅适用于气泡图。') }
function chartGaugeValueFieldMapping(current: ChartSemanticOptions['mapping'], index: number, category: number): NonNullable<ChartSemanticOptions['mapping']> {
  validateChartSeriesField(index)
  return { ...current, gauge: { ...(current?.gauge ?? { labelIndex: category }), valueIndex: index } }
}
function chartAxisPatch(current: Partial<ChartAxisOptions> | undefined, patch: ChartNullablePatch<ChartAxisOptions>): ChartAxisOptions {
  return chartCartesianAxisPatch(current, patch)
}

export class FRange {
  private readonly api: OfficeStudio
  readonly unitId: string
  readonly sheetId: string
  private readonly range: FRangeCoordinates

  constructor(
    api: OfficeStudio,
    unitId: string,
    sheetId: string,
    range: FRangeCoordinates
  ) {
    this.api = api
    this.unitId = unitId
    this.sheetId = sheetId
    this.range = normalizeRange(range)
  }

  getUnitId(): string { return this.unitId }
  getSheetId(): string { return this.sheetId }
  getRange(): FRangeCoordinates { return { ...this.range } }
  insertCells(shiftDimension: SheetCellShiftDimension): void {
    if (shiftDimension !== SheetCellShiftDimensions.ROWS && shiftDimension !== SheetCellShiftDimensions.COLUMNS) throw new TypeError('单元格插入方向无效。')
    this.api.executeCommandSync(shiftDimension === SheetCellShiftDimensions.ROWS ? SHEET_CELL_SHIFT_COMMANDS.down : SHEET_CELL_SHIFT_COMMANDS.right,
      { unitId: this.unitId, subUnitId: this.sheetId, range: this.range })
  }
  deleteCells(shiftDimension: SheetCellShiftDimension): void {
    if (shiftDimension !== SheetCellShiftDimensions.ROWS && shiftDimension !== SheetCellShiftDimensions.COLUMNS) throw new TypeError('单元格删除方向无效。')
    this.api.executeCommandSync(shiftDimension === SheetCellShiftDimensions.ROWS ? SHEET_CELL_SHIFT_COMMANDS.up : SHEET_CELL_SHIFT_COMMANDS.left,
      { unitId: this.unitId, subUnitId: this.sheetId, range: this.range })
  }
  useThemeStyle(themeName: string | null | undefined): void {
    this.api.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.set, { unitId: this.unitId, subUnitId: this.sheetId, range: this.range, themeName })
  }
  removeThemeStyle(themeName: string): void {
    this.api.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.remove, { unitId: this.unitId, subUnitId: this.sheetId, range: this.range, themeName })
  }
  getUsedThemeStyle(): string | undefined { return sheetUsedRangeThemeRule(this.sheet(), toCellRange(this.range))?.themeName }
  getRow(): number { return this.range.startRow }
  getLastRow(): number { return this.range.endRow }
  getColumn(): number { return this.range.startColumn }
  getLastColumn(): number { return this.range.endColumn }
  getHeight(): number { return this.range.endRow - this.range.startRow + 1 }
  getWidth(): number { return this.range.endColumn - this.range.startColumn + 1 }
  getA1Notation(withSheet = false): string {
    const first = cellAddress(this.range.startRow, this.range.startColumn)
    const last = cellAddress(this.range.endRow, this.range.endColumn)
    const ref = this.range.rangeType === 1 ? `${this.range.startRow + 1}:${this.range.endRow + 1}`
      : this.range.rangeType === 2 ? `${first.replace(/\d+$/, '')}:${last.replace(/\d+$/, '')}`
        : first === last ? first : `${first}:${last}`
    const sheetName = this.sheet().name.replaceAll("'", "''")
    return withSheet ? `'${sheetName}'!${ref}` : ref
  }
  getValue(): CellScalar { return this.cellAt(this.range.startRow, this.range.startColumn)?.value ?? null }
  getValues(): CellScalar[][] { return this.mapCells((cell) => cell?.value ?? null) }
  getRawValue(): CellScalar { return this.getValue() }
  getRawValues(): CellScalar[][] { return this.getValues() }
  getDisplayValue(): string { return this.formattedCell(this.cellAt(this.range.startRow, this.range.startColumn)) }
  getDisplayValues(): string[][] { return this.mapCells(cell => this.formattedCell(cell)) }
  getFormula(): string { return this.cellAt(this.range.startRow, this.range.startColumn)?.formula ?? '' }
  getFormulas(): string[][] { return this.mapCells((cell) => cell?.formula ?? '') }
  autoFill(targetRange: FRange, applyType?: SheetFillApplyType): Promise<boolean> {
    if (!(targetRange instanceof FRange) || targetRange.api !== this.api || targetRange.unitId !== this.unitId || targetRange.sheetId !== this.sheetId) {
      throw new TypeError('自动填充的源范围和目标范围必须属于同一工作簿及工作表。')
    }
    const source = toCellRange(this.range), target = toCellRange(targetRange.getRange())
    if (target.startRow > source.startRow || target.endRow < source.endRow || target.startCol > source.startCol || target.endCol < source.endCol) {
      throw new RangeError('自动填充的目标必须包含源范围。')
    }
    resolveSheetFillTarget(source, target, false)
    return this.api.executeCommand(SHEET_AUTO_FILL_COMMAND, { unitId: this.unitId, subUnitId: this.sheetId,
      sourceRange: source, targetRange: target, applyType: sheetFillApplyType(applyType) })
  }
  getDataValidation(): DataValidationRule | null { return new FWorksheet(this.api, this.unitId, this.sheetId).getDataValidation(this.range) }
  setDataValidation(rule: DataValidationRuleInput | null): this {
    new FWorksheet(this.api, this.unitId, this.sheetId).setDataValidation(this.range, rule)
    return this
  }
  async setDataValidationAsync(rule: DataValidationRuleInput | null, options: SheetDataValidationAsyncOptions = {}): Promise<this> {
    await new FWorksheet(this.api, this.unitId, this.sheetId).setDataValidationAsync(this.range, rule, options)
    return this
  }
  getDataValidationErrors(): DataValidationError[] { return new FWorksheet(this.api, this.unitId, this.sheetId).getDataValidationErrors(this.range) }
  getRangeProtectionRules(): SheetRangeProtectionRule[] { return new FWorksheet(this.api, this.unitId, this.sheetId).getRangeProtectionRulesForRange(this.range) }
  setRangeProtection(draft: SheetRangeProtectionDraft): this {
    new FWorksheet(this.api, this.unitId, this.sheetId).setRangeProtection(this.range, draft)
    return this
  }
  sort(column: SortColumnSpec | SortColumnSpec[], options: SortRangeOptions = {}): this {
    const specs = Array.isArray(column) ? column : [column]
    const criteria: SheetSortCriterion[] = specs.map((spec) => {
      const relative = typeof spec === 'number' ? { column: spec, direction: 'ascending' as const } : { column: spec.column, direction: spec.ascending ? 'ascending' as const : 'descending' as const }
      if (!Number.isInteger(relative.column) || relative.column < 0 || relative.column >= this.getWidth()) throw new RangeError('Sort column must be relative to the range')
      return { column: this.range.startColumn + relative.column, direction: relative.direction }
    })
    const range: CellRange = { startRow: this.range.startRow, startCol: this.range.startColumn, endRow: this.range.endRow, endCol: this.range.endColumn }
    const order = getSortedRowOrder(this.sheet(), range, criteria, options)
    this.api.executeCommandSync('sheet.command.sort-range', { unitId: this.unitId, sheetId: this.sheetId, range, criteria, options, order })
    return this
  }
  getCellData(): (SpreadsheetCell & UniverSheetCellData) | null { const cell = this.cellAt(this.range.startRow, this.range.startColumn); return cell ? sheetCellDataWithPublicImage(cell) : null }
  getCellDatas(): ((SpreadsheetCell & UniverSheetCellData) | null)[][] { return this.mapCells((cell) => cell ? sheetCellDataWithPublicImage(cell) : null) }
  insertCellImageAsync(source: string | File, options: SheetImageAsyncOptions = {}): Promise<boolean> {
    return this.api.executeCommand(SHEET_IMAGE_COMMANDS.insertCell, { ...options, unitId: this.unitId, sheetId: this.sheetId,
      row: this.range.startRow, column: this.range.startColumn, ranges: [toCellRange(this.range)],
      ...(typeof source === 'string' ? { source } : { files: [source] }) })
  }
  saveCellImagesAsync(options: SaveCellImagesOptions = {}): Promise<boolean> {
    return this.api.executeCommand(SHEET_IMAGE_COMMANDS.saveCell, { unitId: this.unitId, sheetId: this.sheetId,
      ranges: [toCellRange(this.range)], saveConfig: sheetImageSaveConfig(options) })
  }
  /** Explicit editor-text interpretation; setValue/setValues retain their scalar/DTO API semantics. */
  setInputText(input: string, options: SheetCellInputOptions = {}): SheetCellInputCommandResult {
    if (this.getHeight() !== 1 || this.getWidth() !== 1) {
      const merge = this.sheet().mergedRanges.find(item => item.startRow === this.range.startRow && item.endRow === this.range.endRow
        && item.startCol === this.range.startColumn && item.endCol === this.range.endColumn)
      if (!merge) throw new RangeError('编辑文本需要单个单元格或完整合并区。')
    }
    return this.api.executeCommandSync(COMMIT_SHEET_CELL_INPUT_COMMAND, { ...options, unitId: this.unitId, sheetId: this.sheetId,
      row: this.range.startRow, column: this.range.startColumn, input })
  }
  setValue(value: ApiCellValue): this { return this.setValues(this.fill(value)) }
  setValues(values: ApiCellValue[][]): this {
    assertMatrixSize(values, this.getHeight(), this.getWidth())
    const current = this.api.getWorkbookModel(this.unitId)
    if (!current) throw new Error(`Workbook not found: ${this.unitId}`)
    const candidate = cloneModel(current)
    const candidateSheet = candidate.sheets.find(sheet => sheet.id === this.sheetId)
    if (!candidateSheet) throw new Error(`Worksheet not found: ${this.sheetId}`)
    const positions: Array<{ row: number; column: number }> = []
    for (let row = 0; row < this.getHeight(); row++) for (let column = 0; column < this.getWidth(); column++) {
      const targetRow = this.range.startRow + row
      const targetColumn = this.range.startColumn + column
      candidateSheet.cells[`${targetRow}:${targetColumn}`] = applyCellValue(candidateSheet.cells[`${targetRow}:${targetColumn}`], targetRow, targetColumn, values[row]![column]!, candidate, candidateSheet)
      positions.push({ row: targetRow, column: targetColumn })
    }
    const validation = validateSheetDataValidationChanges(candidate, this.sheetId, positions, sheetValidationSourceReader(this.api.currentUserId), this.api.regionalSettings.region.value)
    if (validation.blocked) throw new Error(validation.issues.find(issue => issue.severity === 'stop')?.message ?? 'Input violates data validation')
    this.mutate('sheet.command.set-range-values', 'Set range values', (model) => {
      const sheet = this.sheet(model)
      for (let row = 0; row < this.getHeight(); row++) for (let column = 0; column < this.getWidth(); column++) {
        const targetRow = this.range.startRow + row
        const targetColumn = this.range.startColumn + column
        const key = `${targetRow}:${targetColumn}`
        sheet.cells[key] = applyCellValue(sheet.cells[key], targetRow, targetColumn, values[row]![column]!, model, sheet)
      }
    }, { range: this.range, values })
    this.api.events.emit('SheetValueChanged', { workbookId: this.unitId, sheetId: this.sheetId, ranges: [{ ...this.range, sheetId: this.sheetId, unitId: this.unitId }] })
    return this
  }
  setFormula(formula: string): this { return this.setValue({ f: formula }) }
  setFormulas(formulas: string[][]): this { return this.setValues(formulas.map((row) => row.map((formula) => ({ f: formula || null }))) as ApiCellValue[][]) }
  setNumberFormat(pattern: string): this {
    this.api.executeCommandSync(SHEET_NUMBER_FORMAT_COMMANDS.set, { unitId: this.unitId, sheetId: this.sheetId, ranges: [toCellRange(this.range)], pattern })
    return this
  }
  setNumberFormats(patterns: string[][]): this {
    assertMatrixSize(patterns, this.getHeight(), this.getWidth())
    const values: SheetNumberFormatCellInput[] = []
    for (let row = 0; row < this.getHeight(); row += 1) for (let col = 0; col < this.getWidth(); col += 1) {
      const pattern = patterns[row]![col]!
      if (typeof pattern !== 'string') throw new TypeError('数字格式矩阵必须包含字符串。')
      values.push({ row: this.range.startRow + row, col: this.range.startColumn + col, pattern })
    }
    this.api.executeCommandSync(SHEET_NUMBER_FORMAT_COMMANDS.setCells, { unitId: this.unitId, subUnitId: this.sheetId, values })
    return this
  }
  getNumberFormat(): string {
    const pattern = this.styleAt(this.range.startRow, this.range.startColumn).numberFormat
    return !pattern || /^general$/i.test(pattern) ? '' : pattern
  }
  getNumberFormats(): string[][] { return this.mapStyles(style => /^general$/i.test(style.numberFormat) ? '' : style.numberFormat) }
  getCellStyleData(type: 'row' | 'col' | 'cell' = 'row'): SheetApiStyleData | null {
    if (!['row', 'col', 'cell'].includes(type)) throw new TypeError('单元格样式读取类型无效。')
    const sheet = this.sheet(), row = this.range.startRow, col = this.range.startColumn, cell = sheet.cells[`${row}:${col}`]
    if (type === 'cell') return cell ? sheetStyleDataWithPublicFields(sheetCellDirectStyle(cell, this.api.getWorkbookModel(this.unitId)) ?? {}) : null
    return sheetStyleDataWithPublicFields(resolveSheetCellStyle({ ...sheet, rowStylePriority: type === 'row' }, row, col, cell, this.api.getWorkbookModel(this.unitId)))
  }
  getFontFamily(type: 'row' | 'col' | 'cell' = 'row'): string | null { return this.getCellStyleData(type)?.ff ?? null }
  getFontSize(type: 'row' | 'col' | 'cell' = 'row'): number | null { return this.getCellStyleData(type)?.fs ?? null }
  setFontWeight(weight: 'normal' | 'bold' | null): this { return this.formatFont(SHEET_FONT_FORMAT_COMMANDS.fontWeight, weight) }
  setFontStyle(style: 'normal' | 'italic' | null): this { return this.formatFont(SHEET_FONT_FORMAT_COMMANDS.fontStyle, style) }
  setFontFamily(family: string | null): this { return this.formatFont(SHEET_FONT_FORMAT_COMMANDS.fontFamily, family) }
  setFontSize(size: number | null): this { return this.formatFont(SHEET_FONT_FORMAT_COMMANDS.fontSize, size) }
  /** UTF-16 offsets in a text cell; shares the editor's font/runs planner and workbook history. */
  setRichTextStyle(start: number, end: number, style: SheetCellInlineStylePatch): SheetCellInputCommandResult {
    if (this.getHeight() !== 1 || this.getWidth() !== 1) {
      const merge = this.sheet().mergedRanges.find(item => item.startRow === this.range.startRow && item.endRow === this.range.endRow
        && item.startCol === this.range.startColumn && item.endCol === this.range.endColumn)
      if (!merge) throw new RangeError('局部文字格式需要单个文本单元格或完整合并区。')
    }
    return this.api.executeCommandSync(FORMAT_SHEET_CELL_RICH_TEXT_COMMAND, { unitId: this.unitId, sheetId: this.sheetId,
      row: this.range.startRow, column: this.range.startColumn, start, end, style })
  }
  setFontColor(color: string | null): this { return this.formatFont(color === null ? SHEET_FONT_FORMAT_COMMANDS.resetTextColor : SHEET_FONT_FORMAT_COMMANDS.fontColor, color) }
  setBackground(color: string | null): this { return this.formatFont(color === null ? SHEET_FONT_FORMAT_COMMANDS.resetBackground : SHEET_FONT_FORMAT_COMMANDS.background, color) }
  /** Border topology is relative to this complete range, including merged cells and shared edges. */
  setBorder(type: SheetBorderType | SheetBorderPreset, style: SheetBorderStyleType | SheetBorderLineStyle, color?: string): this {
    this.api.executeCommandSync(SHEET_BORDER_FORMAT_COMMANDS.basic, {
      unitId: this.unitId, subUnitId: this.sheetId, ranges: [toCellRange(this.range)], value: { type, style, ...(color !== undefined ? { color } : {}) },
    })
    return this
  }
  private formatFont(command: string, value: unknown): this {
    this.api.executeCommandSync(command, { unitId: this.unitId, sheetId: this.sheetId, ranges: [toCellRange(this.range)], value })
    return this
  }
  setHorizontalAlignment(alignment: 'general' | 'left' | 'center' | 'right' | 'justify'): this { return this.formatLayout('sheet.command.set-horizontal-alignment', alignment) }
  setVerticalAlignment(alignment: 'top' | 'middle' | 'bottom'): this { return this.formatLayout('sheet.command.set-vertical-alignment', alignment) }
  setWrap(enabled: boolean): this { return this.formatLayout(SHEET_LAYOUT_FORMAT_COMMANDS.wrap, enabled) }
  setWrapStrategy(strategy: SheetWrapStrategy): this { return this.formatLayout(SHEET_LAYOUT_FORMAT_COMMANDS.wrapStrategy, strategy) }
  setShrinkToFit(enabled: boolean): this { return this.formatLayout(SHEET_LAYOUT_FORMAT_COMMANDS.shrink, enabled) }
  setTextRotation(rotation: SheetTextRotation): this { return this.formatLayout(SHEET_LAYOUT_FORMAT_COMMANDS.rotation, rotation) }
  getHorizontalAlignment(): 'general' | 'left' | 'center' | 'right' | 'justify' { return this.styleAt(this.range.startRow, this.range.startColumn).horizontal }
  getVerticalAlignment(): 'top' | 'middle' | 'bottom' { return this.styleAt(this.range.startRow, this.range.startColumn).vertical }
  getWrap(): boolean { return this.getWrapStrategy() === 'wrap' }
  getWrapStrategy(): SheetWrapStrategy { return sheetCellWrapStrategy(this.styleAt(this.range.startRow, this.range.startColumn)) }
  getTextRotation(): SheetTextRotation { return sheetCellTextRotation(this.styleAt(this.range.startRow, this.range.startColumn)) }
  private formatLayout(command: string, value: unknown): this {
    this.api.executeCommandSync(command, { unitId: this.unitId, sheetId: this.sheetId, ranges: [toCellRange(this.range)], value })
    return this
  }
  getBackground(): string { return this.styleAt(this.range.startRow, this.range.startColumn).fill ?? '#ffffff' }
  getFontWeight(): 'normal' | 'bold' { return this.styleAt(this.range.startRow, this.range.startColumn).font.bold ? 'bold' : 'normal' }
  merge(options: SheetMergeOptions = {}): this { return this.mergeWithMode('all', options) }
  mergeAcross(options: SheetMergeOptions = {}): this { return this.mergeWithMode('horizontal', options) }
  mergeVertically(options: SheetMergeOptions = {}): this { return this.mergeWithMode('vertical', options) }
  private mergeWithMode(mode: SheetMergeMode, options: SheetMergeOptions): this {
    this.api.executeCommandSync(SHEET_MERGE_COMMANDS.merge, { unitId: this.unitId, sheetId: this.sheetId, ranges: [toCellRange(this.range)],
      mode, defaultMerge: options.defaultMerge ?? true, isForceMerge: options.isForceMerge ?? false })
    return this
  }
  breakApart(): this {
    this.api.executeCommandSync(SHEET_MERGE_COMMANDS.unmerge, { unitId: this.unitId, sheetId: this.sheetId, ranges: [toCellRange(this.range)] })
    return this
  }
  isMerged(): boolean { return this.sheet().mergedRanges.some(range => sheetRangesEqual(range, toCellRange(this.range))) }
  isPartOfMerge(): boolean { return this.sheet().mergedRanges.some(range => sheetRangesIntersect(range, toCellRange(this.range))) }
  clear(options: SheetClearOptions = {}): this {
    this.api.executeCommandSync(SHEET_CLEAR_COMMANDS.range, { unitId: this.unitId, sheetId: this.sheetId, ranges: [toCellRange(this.range)], options })
    return this
  }
  clearContent(): this {
    this.api.executeCommandSync(SHEET_CLEAR_COMMANDS.content, { unitId: this.unitId, sheetId: this.sheetId, ranges: [toCellRange(this.range)] })
    return this
  }
  clearFormat(): this {
    this.api.executeCommandSync(SHEET_CLEAR_COMMANDS.format, { unitId: this.unitId, sheetId: this.sheetId, ranges: [toCellRange(this.range)] })
    return this
  }
  activate(): this { this.api.setSelection(this.unitId, this.sheetId, this.range); return this }

  private mutateStyle(id: string, label: string, change: (cell: SpreadsheetCell) => void): this {
    this.mutate(id, label, (model) => {
      const sheet = this.sheet(model)
      for (let row = this.range.startRow; row <= this.range.endRow; row++) for (let column = this.range.startColumn; column <= this.range.endColumn; column++) {
        const key = `${row}:${column}`
        const cell = sheet.cells[key] ?? createCell(row, column, null)
        change(cell)
        sheet.cells[key] = cell
      }
    })
    return this
  }

  private formattedCell(cell: SpreadsheetCell | undefined): string {
    if (!cell) return ''
    if (cell.formula && cell.value === null) return cell.display
    const workbook = this.api.getWorkbookModel(this.unitId)
    return formatSheetValue(cell.value, cell.style.numberFormat, workbook?.date1904, this.api.regionalSettings.region.value) ?? cell.display
  }
  private mutate(id: string, label: string, update: (model: WorkbookModel) => void, params: unknown = { range: this.range }): void {
    this.api.mutateWorkbook(this.unitId, id, label, update, { ...asRecord(params), unitId: this.unitId, sheetId: this.sheetId })
  }
  private mapCells<T>(read: (cell: SpreadsheetCell | undefined) => T): T[][] {
    const sheet = this.sheet()
    return Array.from({ length: this.getHeight() }, (_, row) => Array.from({ length: this.getWidth() }, (_, column) => read(sheet.cells[`${this.range.startRow + row}:${this.range.startColumn + column}`])))
  }
  private styleAt(row: number, column: number): CellStyle {
    const sheet = this.sheet()
    return resolveSheetCellStyle(sheet, row, column, sheet.cells[`${row}:${column}`], this.api.getWorkbookModel(this.unitId))
  }
  private mapStyles<T>(read: (style: CellStyle) => T): T[][] {
    return Array.from({ length: this.getHeight() }, (_, row) => Array.from({ length: this.getWidth() }, (_, column) =>
      read(this.styleAt(this.range.startRow + row, this.range.startColumn + column))))
  }
  private cellAt(row: number, column: number): SpreadsheetCell | undefined {
    const cell = this.sheet().cells[`${row}:${column}`]
    return cell ? { ...cell, style: this.styleAt(row, column) } : undefined
  }
  private sheet(model = this.api.getWorkbookModel(this.unitId)): SpreadsheetSheet {
    if (!model) throw new Error(`Workbook not found: ${this.unitId}`)
    const sheet = model.sheets.find((item) => item.id === this.sheetId)
    if (!sheet) throw new Error(`Worksheet not found: ${this.sheetId}`)
    this.api.assertCanViewWorksheet(this.unitId, this.sheetId)
    return sheet
  }
  private fill(value: ApiCellValue): ApiCellValue[][] {
    return Array.from({ length: this.getHeight() }, () => Array.from({ length: this.getWidth() }, () => value))
  }
}

export const PositionTypeEnum = { Existing: 'existing' } as const
export type PositionTypeEnum = typeof PositionTypeEnum[keyof typeof PositionTypeEnum]
export const PivotValuePositions = PIVOT_VALUE_POSITIONS
export type PivotValuePositionInput = PivotValuePosition | typeof PivotValuePositions[keyof typeof PivotValuePositions]

export class FPivotTable {
  private readonly api: OfficeStudio
  private readonly unitId: string
  readonly id: string

  constructor(api: OfficeStudio, unitId: string, id: string) {
    this.api = api
    this.unitId = unitId
    this.id = id
  }

  getPivotTableId(): string { return this.id }
  getUnitId(): string { return this.unitId }
  getSubUnitId(): string { return this.model().anchor.subUnitId }
  newChart(type: ChartType): ChartConfigBuilder { return new ChartConfigBuilder(type).setPivotTableId(this.id) }
  async insertChart(info: ChartInfo): Promise<FSheetPivotChart> {
    const sheet = new FWorksheet(this.api, this.unitId, this.model().anchor.subUnitId)
    return sheet.insertPivotChart({ ...info, pivotTableId: this.id })
  }
  getSourceFieldsInfo(): Array<{ fieldId: string; fieldName: string; sourceColumn: number }> {
    return getPivotSourceFields(this.workbook(), this.model())
  }
  getFieldIdsByArea(area: PivotFieldArea): string[] { return pivotFieldIds(this.model(), area) }
  getData(): CellScalar[][] { return pivotOutputValues(this.workbook(), this.model()) }
  getRange(): FRange | null {
    const range = this.model().outputRange
    return range ? new FRange(this.api, this.unitId, this.model().anchor.subUnitId, {
      startRow: range.startRow,
      startColumn: range.startCol,
      endRow: range.endRow,
      endColumn: range.endCol,
    }) : null
  }
  getOptions(): PivotTableOptions { return cloneModel(this.model().options) }
  getFieldItems(fieldId: string): string[] {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    return getPivotGroupedFieldItems(workbook, report, fieldId)
  }
  getManualFilterItemState(fieldId: string): { checkedItems: string[]; partialItems: string[] } {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    return getPivotManualFilterItemState(workbook, report, fieldId)
  }
  getFilterCapabilities(fieldId: string): { supportsDateFilter: boolean; date1904: boolean } {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    return getPivotFieldFilterCapabilities(workbook, report, fieldId)
  }
  getDateFilter(fieldId: string): PivotDateFilter | undefined {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    return getPivotDateFilter(report, fieldId)
  }
  getLabelConditionFilter(fieldId: string): PivotLabelConditionFilter | undefined {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    return getPivotLabelConditionFilter(report, fieldId)
  }
  async setManualFilterItems(fieldId: string, items: string[] | null, isAll = false, partialItems?: string[], context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute(PIVOT_REPORT_FILTER_COMMANDS.setFilter, { ...context, tableFieldId: fieldId, items, isAll, partialItems })
  }
  async clearFilters(fieldId: string, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute(PIVOT_REPORT_FILTER_COMMANDS.setFilter, { ...context, tableFieldId: fieldId, items: [], clearAllFilters: true })
  }
  async setDateFilter(fieldId: string, filterInfo?: PivotDateFilterInput | null, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute(PIVOT_REPORT_FILTER_COMMANDS.setDateFilter, { ...context, tableFieldId: fieldId, filterInfo })
  }
  async clearDateFilter(fieldId: string, context: PivotReportCommandContext = {}): Promise<this> { return this.setDateFilter(fieldId, undefined, context) }
  async setLabelConditionFilter(fieldId: string, filterInfo?: PivotLabelConditionInput | null, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute(PIVOT_REPORT_FILTER_COMMANDS.setLabelConditionFilter, { ...context, tableFieldId: fieldId, filterInfo })
  }
  async clearLabelConditionFilter(fieldId: string, context: PivotReportCommandContext = {}): Promise<this> { return this.setLabelConditionFilter(fieldId, undefined, context) }
  async addField(fieldIdOrIndex: string | number, area: PivotFieldArea, index?: number, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-field', { ...context, fieldId: fieldIdOrIndex, area, index })
  }
  async removeField(fieldIds: string | string[], context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.remove-pivot-field', { ...context, fieldIds: Array.isArray(fieldIds) ? fieldIds : [fieldIds] })
  }
  async setOptions(options: Partial<PivotTableOptions>, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-options', { ...context, options })
  }
  async setLabelSort(fieldId: string, options: { type: 'ascending' | 'descending' }, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-label-sort', { ...context, fieldId, direction: options.type })
  }
  async clearLabelSort(fieldId: string, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-label-sort', { ...context, fieldId, direction: null })
  }
  async setLabelManualFilter(fieldId: string, selectedItems: string[] | null, isAll = false, context: PivotReportCommandContext = {}, partialItems?: string[]): Promise<this> {
    return this.setManualFilterItems(fieldId, selectedItems, isAll, partialItems, context)
  }
  async setFieldSetting(fieldId: string, setting: { displayName?: string; format?: string; subtotalType?: PivotAggregation; showDataAs?: PivotShowDataAsInfo | null }, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-field-setting', { ...context, fieldId, setting })
  }
  async setSubtotalType(fieldId: string, subtotalType: PivotAggregation, context: PivotReportCommandContext = {}): Promise<this> {
    return this.setFieldSetting(fieldId, { subtotalType }, context)
  }
  getShowDataAs(fieldId: string): PivotShowDataAsInfo | undefined {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    const info = getPivotShowDataAs(report, fieldId)
    return info ? cloneModel(info) : undefined
  }
  getShowDataAsBaseItems(baseFieldId: string): string[] {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    return getPivotShowDataAsBaseItems(workbook, report, baseFieldId)
  }
  async setShowDataAs(fieldId: string, showDataAs: PivotShowDataAsInfo | null, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-show-data-as', { ...context, fieldId, showDataAs })
  }
  async resetShowDataAs(fieldId: string, context: PivotReportCommandContext = {}): Promise<this> { return this.setShowDataAs(fieldId, null, context) }
  getAxisState(area: PivotAxisArea = 'row'): PivotAxisState {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    return cloneModel(getPivotAxisState(workbook, report, area))
  }
  getCollapse(fieldId: string, item?: string): boolean {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    return getPivotCollapse(report, fieldId, item)
  }
  getCollapseItems(fieldId: string): string[] {
    const workbook = this.workbook(), report = this.model()
    assertPivotReportReadPermission(workbook, report, this.api.currentUserId)
    return getPivotCollapseItems(workbook, report, fieldId)
  }
  async setCollapse(fieldId: string, collapse: boolean, item?: string, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute(PIVOT_HIERARCHY_COMMANDS.setCollapse, { ...context, fieldId, collapse, item })
  }
  async setCategoryLevel(fieldId: string, context: PivotReportCommandContext = {}): Promise<this> { return this.execute(PIVOT_HIERARCHY_COMMANDS.setCategoryLevel, { ...context, fieldId }) }
  async stepCategoryLevel(direction: 'expand' | 'collapse', context: PivotReportCommandContext = {}): Promise<this> { return this.execute(PIVOT_HIERARCHY_COMMANDS.stepCategoryLevel, { ...context, direction }) }
  async setValueFilter(fieldId: string, filter: PivotFieldFilter | null, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-value-filter', { ...context, fieldId, filter })
  }
  async clearValueFilter(fieldId: string, context: PivotReportCommandContext = {}): Promise<this> { return this.setValueFilter(fieldId, null, context) }
  async setNumberGroup(fieldId: string, settings: PivotNumberGroup | null, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-number-group', { ...context, fieldId, settings })
  }
  async setDateGroupingMode(mode: PivotDateGroupingMode, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-date-grouping-mode', { ...context, mode })
  }
  async setDateGroupType(fieldId: string, type: PivotDateGroupType | null, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-date-group-type', { ...context, fieldId, type })
  }
  async setExcelDateGroup(fieldId: string, type: PivotDateGroupType, context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-excel-date-group', { ...context, fieldId, type })
  }
  async setManualGroups(fieldId: string, groups: PivotManualGroup[], context: PivotReportCommandContext = {}): Promise<this> {
    return this.execute('sheet.command.set-pivot-manual-groups', { ...context, fieldId, groups })
  }
  getValuesPosition(): { position: PivotValuePosition; index: number } { return getPivotValuesPosition(this.model()) }
  async updateValuePosition(position: PivotValuePositionInput, index = -1, context: PivotReportCommandContext = {}): Promise<this> {
    const state = normalizePivotValuesPosition(normalizePivotValuePositionInput(position), index)
    return this.execute('sheet.command.update-pivot-value-position', { ...context, position: state.position, index: state.index })
  }
  async refresh(context: PivotReportCommandContext = {}): Promise<this> { return this.execute('sheet.command.refresh-pivot-table', { ...context }) }
  async remove(context: PivotReportCommandContext = {}): Promise<void> {
    await this.api.executeCommand('sheet.command.remove-pivot-table', { ...context, unitId: this.unitId, pivotTableId: this.id })
  }
  drillDown(rowOffset: number, columnOffset: number): { headers: string[]; rows: CellScalar[][] } {
    return pivotDrillDown(this.workbook(), this.model(), rowOffset, columnOffset)
  }

  private async execute(commandId: string, params: Record<string, unknown>): Promise<this> {
    // A successful command can leave identical content and return false. The
    // command itself still rejects cancellation, permissions and stale owners.
    await this.api.executeCommand(commandId, { ...params, unitId: this.unitId, pivotTableId: this.id })
    return this
  }
  private model(): PivotTableModel {
    const model = this.workbook().pivotTables?.find((pivot) => pivot.id === this.id)
    if (!model) throw new Error(`Pivot table not found: ${this.id}`)
    return model
  }
  private workbook(): WorkbookModel {
    const workbook = this.api.getWorkbookModel(this.unitId)
    if (!workbook) throw new Error(`Workbook not found: ${this.unitId}`)
    return workbook
  }
}

export class FSheetSlicer {
  private readonly api: OfficeStudio
  readonly unitId: string
  readonly subUnitId: string
  readonly id: string

  constructor(api: OfficeStudio, unitId: string, id: string) {
    this.api = api
    this.unitId = unitId
    this.id = id
    this.subUnitId = api.getWorkbookModel(unitId)?.slicers?.find(control => control.id === id)?.hostSheetId ?? ''
  }

  getId(): string { return this.id }
  getUnitId(): string { return this.unitId }
  getSubUnitId(): string { return this.subUnitId }
  getDrawingId(): string { return this.id }
  /** Local hosted-model projection. It remains readable when a source/cache is unavailable. */
  getModel(): SlicerModel { return cloneModel(this.model()) }
  getInfo(): ISheetSlicerInfo { return readSheetSlicerInfo(this.workbook(), this.model()) }
  getDrawing(): ISheetSlicerDrawing | undefined {
    const model = this.optionalModel()
    return model?.drawing ? readSheetSlicerDrawing(this.workbook(), model) : undefined
  }
  getSource(): SheetSlicerSourceSpec | undefined {
    const model = this.optionalModel()
    return model ? readSheetSlicerSource(this.workbook(), model) : undefined
  }
  getCache(): SlicerCacheDefinition | undefined {
    const cache = this.optionalCache()
    return cache?.type === 'slicer' ? sheetSlicerCacheDefinition(this.workbook(), cache) : undefined
  }
  /** Local retained cache projection used by the independent Timeline facade. */
  getCacheModel(): import('../domain/model.ts').SlicerCacheModel | undefined {
    const cache = this.optionalCache()
    return cache ? cloneModel(cache) : undefined
  }
  toBuilder(): FBoundSheetSlicerBuilder { return FSheetSlicerBuilder.fromInfo(this.api, this.unitId, this.getInfo()) }
  async update(info: ISheetSlicerInfo, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<this> {
    const accepted = await this.api.executeCommand(SLICER_PUBLIC_COMMANDS.update, { unitId: this.unitId, slicerId: this.id, info,
      expectedRevision: context.expectedRevision, host: context.host })
    if (!accepted) throw new Error('切片器信息更新未能提交。')
    return this
  }
  /** Local display projection shares raw identity and cross-filter availability with public member windows. */
  getDisplayMembers(): SlicerDisplayMember[] {
    const workbook = this.workbook(), model = this.model()
    assertSlicerSourceReadPermission(workbook, model, this.api.currentUserId)
    return getSlicerDisplayMembers(workbook, model, this.api.regionalSettings.locale.value)
  }
  /** Local Canvas/panel label projection; public getItems is an asynchronous typed window. */
  getItemLabels(): string[] {
    const workbook = this.workbook(), model = this.model()
    assertSlicerSourceReadPermission(workbook, model, this.api.currentUserId)
    return getSlicerItems(workbook, model, this.api.regionalSettings.locale.value)
  }
  async getItems(request: ISheetSlicerItemsRequest = {}): Promise<ISlicerSourceMemberWindow> {
    const workbook = this.workbook(), model = this.model(), context = this.queryContext(workbook, model)
    return this.api.slicerMemberQueryService.getItems(workbook, model, request, context)
  }
  /** Local Canvas window folds raw captions and retains checked/partial states; the public typed member API is independent. */
  async getDisplayMemberWindow(request: ISheetSlicerItemsRequest = {}): Promise<import('../features/slicer/slicer-member-query.ts').ISlicerDisplayMemberWindow> {
    const workbook = this.workbook(), model = this.model(), context = this.queryContext(workbook, model)
    return this.api.slicerMemberQueryService.getDisplayMemberWindow(workbook, model, request, context)
  }
  async getSourceMetadata(): Promise<ISlicerSourceMetadata> {
    const workbook = this.workbook(), model = this.model(), context = this.queryContext(workbook, model)
    return this.api.slicerMemberQueryService.getSourceMetadata(workbook, model, context)
  }
  /** Owner-bound Canvas reads can cancel their own query without cancelling another API caller's task. */
  async getCanvasSourceMetadata(signal?: AbortSignal): Promise<ISlicerSourceMetadata> {
    const workbook = this.workbook(), model = this.model(), context = this.queryContext(workbook, model)
    return this.api.slicerMemberQueryService.getSourceMetadata(workbook, model, context, signal)
  }
  getSelection(): RecordSelection | { mode: 'condition' | 'color' | 'manual' } | undefined {
    const workbook = this.api.getWorkbookModel(this.unitId), model = this.optionalModel()
    if (!workbook || !model || model.type !== 'slicer') return undefined
    assertSlicerViewPermission(workbook, model, this.api.currentUserId, false)
    const retained = this.optionalCache()
    let cache: import('../domain/model.ts').SlicerCacheModel
    try { cache = getSlicerCache(workbook, model) }
    catch (error) {
      if (!(error instanceof SlicerConnectionError)) throw error
      if (error.code === 'source-missing' || error.code === 'field-missing') {
        return retained?.source.kind === 'pivot' ? slicerPublicRecordSelection(retained.recordSelection) : undefined
      }
      if (error.code === 'invalid-resource' && !retained) return undefined
      throw error
    }
    assertSlicerSourceReadPermission(workbook, model, this.api.currentUserId)
    if (cache.source.kind === 'table') {
      const filter = getSlicerTableNativeFilter(workbook, model)
      if (filter?.kind === 'condition') return { mode: 'condition' }
    }
    const selection = readSlicerRecordSelection(workbook, model)
    if (selection && selection.mode !== 'all' && selection.values.every(isRecordValue)) return normalizeRecordSelection(selection)
    return slicerPublicRecordSelection(selection, getSlicerSourceValues(workbook, model), getSlicerSourceRecordValues(workbook, model))
  }
  getSelectedItems(): string[] | undefined {
    const workbook = this.workbook(), model = this.model()
    assertSlicerSourceReadPermission(workbook, model, this.api.currentUserId)
    return readSlicerSelection(workbook, model)
  }
  getRecordSelection(): PivotRecordSelection | undefined {
    const workbook = this.workbook(), model = this.model()
    assertSlicerSourceReadPermission(workbook, model, this.api.currentUserId)
    return readSlicerRecordSelection(workbook, model)
  }
  getItemState(): { checkedItems: string[]; partialItems: string[] } {
    const workbook = this.workbook(), model = this.model()
    assertSlicerSourceReadPermission(workbook, model, this.api.currentUserId)
    return readSlicerItemState(workbook, model)
  }
  getCacheId(): string | undefined { return this.optionalCache()?.id }
  getConnections(): SheetSlicerConnections | undefined {
    const cache = this.optionalCache()
    return cache?.source.kind === 'pivot' ? cloneModel(cache.connections ?? { pivotTables: [] }) : undefined
  }
  getSettings(): SheetSlicerSettings | undefined {
    const model = this.optionalModel()
    if (!model || model.type !== 'slicer') return undefined
    return readSlicerSettings(model)
  }
  getCacheOptions(): import('../domain/model.ts').SheetSlicerCacheOptions | undefined {
    const cache = this.optionalCache()
    return cache?.type === 'slicer' ? normalizeSlicerCacheOptions(cache.options) : undefined
  }
  async updateSettings(patch: Partial<SheetSlicerSettings>, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand('sheet.command.update-slicer-settings', { unitId: this.unitId, slicerId: this.id, patch,
      expectedRevision: context.expectedRevision, host: context.host })
  }
  async setCacheOptions(patch: Partial<import('../domain/model.ts').SheetSlicerCacheOptions>, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand('sheet.command.patch-slicer-cache-options', { unitId: this.unitId, slicerId: this.id, patch,
      expectedRevision: context.expectedRevision, host: context.host })
  }
  setName(name: string): Promise<boolean> { return this.updateSettings({ name }) }
  setCaption(caption: string): Promise<boolean> { return this.updateSettings({ caption }) }
  setStyle(style: string): Promise<boolean> { return this.updateSettings({ style }) }
  clearStyle(): Promise<boolean> { return this.updateSettings({ style: undefined }) }
  setColumnCount(columnCount: number): Promise<boolean> { return this.updateSettings({ columnCount }) }
  setRowHeight(rowHeight: number): Promise<boolean> { return this.updateSettings({ rowHeight }) }
  setShowCaption(showCaption: boolean): Promise<boolean> { return this.updateSettings({ showCaption }) }
  setLockedPosition(lockedPosition: boolean): Promise<boolean> { return this.updateSettings({ lockedPosition }) }
  setStartItem(startItem: number): Promise<boolean> { return this.updateSettings({ startItem }) }
  setPosition(anchor: SheetSlicerAnchorSpec): this { return this.setPlacement({ anchor }) }
  setAbsolutePosition(x: number, y: number): this { return this.setPlacement({ position: { x, y } }) }
  setSize(width: number, height: number): this { return this.setPlacement({ size: { width, height } }) }
  bringToFront(): this { return this.setOrder({ arrangeType: 'front' }) }
  bringForward(): this { return this.setOrder({ arrangeType: 'forward' }) }
  sendBackward(): this { return this.setOrder({ arrangeType: 'backward' }) }
  sendToBack(): this { return this.setOrder({ arrangeType: 'back' }) }
  setZOrder(index: number): this {
    if (typeof index !== 'number' || !Number.isFinite(index)) throw new TypeError('绘图层级必须是有限数值。')
    return this.setOrder({ zOrder: index })
  }
  private setPlacement(placement: SheetSlicerPlacement): this {
    this.model()
    const accepted = this.api.executeCommandSync(SLICER_PUBLIC_COMMANDS.placement, { unitId: this.unitId, slicerId: this.id, placement })
    if (!accepted) throw new Error('切片器或时间线绘图位置更新未能提交。')
    return this
  }
  private setOrder(order: { arrangeType: SheetDrawingArrangeType; zOrder?: never } | { zOrder: number; arrangeType?: never }): this {
    const workbook = this.workbook(), model = this.model(), selection: SheetDrawingRef[] = [{ kind: model.type, id: this.id }]
    assertSlicerViewPermission(workbook, model, this.api.currentUserId, true)
    const plan = planDrawingArrange(createSheetDrawingScene(workbook, this.subUnitId), { selection, ...order })
    const accepted = this.api.executeCommandSync(SHEET_DRAWING_GROUP_COMMAND_IDS.arrange, { unitId: this.unitId, sheetId: this.subUnitId, selection, ...order })
    if (!accepted && plan.status !== 'noop') throw new Error('切片器或时间线绘图层级更新未能提交。')
    return this
  }
  async getConnectionCandidates(): Promise<SlicerConnectionCandidate[]> {
    const workbook = this.workbook(), model = this.model(), cache = getSlicerCache(workbook, model)
    this.api.assertCanViewWorksheet(this.unitId, model.hostSheetId)
    if (cache.source.kind === 'table') return []
    const source = getSlicerPivotSource(workbook, model)
    this.api.assertCanViewWorksheet(this.unitId, source.subUnitId)
    const sourceSheet = workbook.sheets.find(sheet => sheet.id === source.subUnitId)!
    if (!canViewProtectedRanges(sourceSheet, [source.range], this.api.currentUserId)) throw new SlicerConnectionError('permission-denied', '没有查看切片器来源范围的权限。')
    return getSlicerConnectionCandidates(workbook, model).filter(candidate => {
      const sheet = workbook.sheets.find(item => item.id === candidate.subUnitId)
      return sheet && canViewWorksheetProtection(sheet, this.api.currentUserId)
    })
  }
  async setConnections(connections: SheetSlicerConnections, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand('sheet.command.set-slicer-connections', { unitId: this.unitId, slicerId: this.id, connections,
      expectedRevision: context.expectedRevision, host: context.host })
  }
  async setItems(items: string[] | null, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<this> {
    const accepted = await this.api.executeCommand('sheet.command.set-slicer-items', { unitId: this.unitId, slicerId: this.id, selectedItems: items,
      expectedRevision: context.expectedRevision, host: context.host })
    if (!accepted) throw new Error('切片器成员筛选未能提交。')
    return this
  }
  async setSelection(selection: RecordSelection, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    let normalized: RecordSelection
    try { normalized = normalizeRecordSelection(selection) }
    catch (error) { if (error instanceof TypeError || error instanceof RangeError) return false; throw error }
    return this.setRecordSelection(normalized, context)
  }
  async setItemStates(checkedItems: string[], partialItems: string[], context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand('sheet.command.set-slicer-item-states', { unitId: this.unitId, slicerId: this.id, checkedItems, partialItems,
      expectedRevision: context.expectedRevision, host: context.host })
  }
  async setRecordSelection(selection: PivotRecordSelection | null, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand('sheet.command.set-slicer-record-selection', { unitId: this.unitId, slicerId: this.id, selection,
      expectedRevision: context.expectedRevision, host: context.host })
  }
  async setTimelineRange(range: SlicerTimelineRange | null, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<this> {
    const accepted = await this.api.executeCommand('sheet.command.set-timeline-range', { unitId: this.unitId, slicerId: this.id, range,
      expectedRevision: context.expectedRevision, host: context.host })
    if (!accepted) throw new Error('时间线筛选未能提交。')
    return this
  }
  async updatePresentation(patch: import('../features/slicer/slicer-engine').SlicerPresentationPatch, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<this> {
    const accepted = await this.api.executeCommand('sheet.command.update-slicer-presentation', { unitId: this.unitId, slicerId: this.id, patch,
      expectedRevision: context.expectedRevision, host: context.host })
    if (!accepted) throw new Error('切片器显示设置未能提交。')
    return this
  }
  async clearFilter(context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand('sheet.command.clear-slicer-filter', { unitId: this.unitId, slicerId: this.id,
      expectedRevision: context.expectedRevision, host: context.host })
  }
  async remove(context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand('sheet.command.remove-slicer', { unitId: this.unitId, slicerId: this.id,
      expectedRevision: context.expectedRevision, host: context.host })
  }
  private optionalCache(): import('../domain/model.ts').SlicerCacheModel | undefined {
    const workbook = this.api.getWorkbookModel(this.unitId), model = this.optionalModel()
    if (!workbook || !model) return undefined
    if (model.cacheId) return workbook.slicerCaches?.find(cache => cache.id === model.cacheId && cache.unitId === this.unitId && cache.type === model.type
      && (cache.source.kind === 'table' ? cache.source.columnId : cache.source.sourceFieldId) === model.sourceFieldId)
    try { return getSlicerCache(workbook, model) }
    catch (error) { if (error instanceof SlicerConnectionError) return undefined; throw error }
  }
  private queryContext(workbook: WorkbookModel, model: SlicerModel): SlicerMemberQueryContext {
    assertSlicerSourceReadPermission(workbook, model, this.api.currentUserId)
    const revision = this.api.getWorkbookRevision(this.unitId), userId = this.api.currentUserId
    const cache = getSlicerCache(workbook, model), identity = JSON.stringify({ cacheId: cache.id, source: cache.source, hostSheetId: model.hostSheetId })
    return { unitId: this.unitId, baseRevision: revision, locale: this.api.regionalSettings.locale.value, assertCurrent: () => {
      if (this.api.isDisposed) throw new SlicerMemberQueryError('SLICER_RPC_DISPOSED', '切片器工作台已释放。')
      const current = this.api.getWorkbookModel(this.unitId), view = current?.slicers?.find(control => control.id === this.id)
      if (!current || !view) throw new SlicerMemberQueryError('SLICER_RPC_UNIT_DISPOSED', '切片器所属工作簿或对象已关闭。')
      if (view.hostSheetId !== this.subUnitId) throw new SlicerMemberQueryError('SLICER_SOURCE_INVALIDATED', '切片器所属工作表已更改。')
      if (current !== workbook || this.api.getWorkbookRevision(this.unitId) !== revision || this.api.currentUserId !== userId) {
        throw new SlicerMemberQueryError('SLICER_SOURCE_INVALIDATED', '切片器或其来源已在查询期间更改。')
      }
      const liveCache = getSlicerCache(current, view)
      if (JSON.stringify({ cacheId: liveCache.id, source: liveCache.source, hostSheetId: view.hostSheetId }) !== identity) {
        throw new SlicerMemberQueryError('SLICER_SOURCE_INVALIDATED', '切片器来源身份已更改。')
      }
      assertSlicerSourceReadPermission(current, view, userId)
    } }
  }
  private model(): SlicerModel {
    const model = this.optionalModel()
    if (!model) throw new Error(`Slicer not found: ${this.id}`)
    return model
  }
  private optionalModel(): SlicerModel | undefined {
    const workbook = this.api.getWorkbookModel(this.unitId), model = workbook?.slicers?.find(control => control.id === this.id)
    if (!workbook || !model) return undefined
    if (model.unitId !== this.unitId || model.hostSheetId !== this.subUnitId || !workbook.sheets.some(sheet => sheet.id === this.subUnitId)) {
      throw new SlicerConnectionError('invalid-target', '切片器的工作簿或绘图宿主已更改。')
    }
    assertSlicerViewPermission(workbook, model, this.api.currentUserId, false)
    return model
  }
  private workbook(): WorkbookModel {
    const workbook = this.api.getWorkbookModel(this.unitId)
    if (!workbook) throw new Error(`Workbook not found: ${this.unitId}`)
    return workbook
  }
}

/** The fixed public Timeline contract has date filtering and per-view settings, independently of Slicer member selections. */
export class FSheetTimeline {
  private readonly control: FSheetSlicer
  readonly subUnitId: string
  constructor(private readonly api: OfficeStudio, readonly unitId: string, readonly id: string) {
    this.control = new FSheetSlicer(api, unitId, id)
    this.subUnitId = this.control.getSubUnitId()
  }
  getId(): string { return this.id }
  getUnitId(): string { return this.unitId }
  getSubUnitId(): string { return this.subUnitId }
  getDrawingId(): string { return this.id }
  getModel(): SlicerModel {
    const model = this.control.getModel()
    if (model.type !== 'timeline') throw new SlicerConnectionError('invalid-target', '该对象不是时间线。')
    return model
  }
  private optionalModel(): SlicerModel | undefined {
    const model = this.api.getWorkbookModel(this.unitId)?.slicers?.find(view => view.id === this.id)
    return model ? this.getModel() : undefined
  }
  private workbook(): WorkbookModel {
    const workbook = this.api.getWorkbookModel(this.unitId)
    if (!workbook) throw new Error(`Workbook not found: ${this.unitId}`)
    return workbook
  }
  getDrawing(): ISheetSlicerDrawing | undefined { return this.optionalModel() ? this.control.getDrawing() : undefined }
  getSource(): SheetSlicerSourceSpec | undefined { return this.optionalModel() ? this.control.getSource() : undefined }
  getCacheId(): string | undefined { return this.optionalModel() ? this.control.getCacheId() : undefined }
  getCache(): ITimelineCacheDefinition | undefined {
    if (!this.optionalModel()) return undefined
    const cache = this.control.getCacheModel()
    return cache ? sheetTimelineCacheDefinition(this.workbook(), cache) : undefined
  }
  getSettings(): SheetTimelineSettings | undefined {
    const model = this.optionalModel()
    return model ? readTimelineSettings(model) : undefined
  }
  getInfo(): ISheetTimelineInfo { return readSheetTimelineInfo(this.workbook(), this.getModel()) }
  getSourceMetadata(): Promise<ISlicerSourceMetadata> { this.getModel(); return this.control.getSourceMetadata() }
  toBuilder(): FBoundSheetTimelineBuilder { return FSheetTimelineBuilder.fromInfo(this.api, this.unitId, this.getInfo()) }
  async update(info: ISheetTimelineInfo, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<this> {
    const accepted = await this.api.executeCommand(SLICER_PUBLIC_COMMANDS.updateTimeline, { unitId: this.unitId, slicerId: this.id, info,
      expectedRevision: context.expectedRevision, host: context.host })
    if (!accepted) throw new Error('时间线信息更新未能提交。')
    return this
  }
  updateSettings(patch: TimelineSettingsInputPatch, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(SLICER_PUBLIC_COMMANDS.timelineSettings, { unitId: this.unitId, slicerId: this.id, patch,
      expectedRevision: context.expectedRevision, host: context.host })
  }
  setName(name: string): Promise<boolean> { return this.updateSettings({ name }) }
  setCaption(caption: string): Promise<boolean> { return this.updateSettings({ caption }) }
  setStyle(style: string): Promise<boolean> { return this.updateSettings({ style }) }
  clearStyle(): Promise<boolean> { return this.updateSettings({ style: undefined }) }
  setLevel(level: TimelinePublicLevel | import('../domain/model.ts').TimelineLevel): Promise<boolean> { return this.updateSettings({ level }) }
  setSelectionLevel(selectionLevel: TimelinePublicLevel | import('../domain/model.ts').TimelineLevel): Promise<boolean> { return this.updateSettings({ selectionLevel }) }
  setScrollPosition(scrollPosition: number): Promise<boolean> { return this.updateSettings({ scrollPosition }) }
  setShowHeader(showHeader: boolean): Promise<boolean> { return this.updateSettings({ showHeader }) }
  setShowSelectionLabel(showSelectionLabel: boolean): Promise<boolean> { return this.updateSettings({ showSelectionLabel }) }
  setShowTimeLevel(showTimeLevel: boolean): Promise<boolean> { return this.updateSettings({ showTimeLevel }) }
  setShowHorizontalScrollbar(showHorizontalScrollbar: boolean): Promise<boolean> { return this.updateSettings({ showHorizontalScrollbar }) }
  getConnections(): SheetSlicerConnections | undefined { return this.optionalModel() ? this.control.getConnections() : undefined }
  getConnectionCandidates(): Promise<SlicerConnectionCandidate[]> { this.getModel(); return this.control.getConnectionCandidates() }
  setConnections(connections: SheetSlicerConnections, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    this.getModel(); return this.control.setConnections(connections, context)
  }
  setPosition(anchor: SheetSlicerAnchorSpec): this { this.getModel(); this.control.setPosition(anchor); return this }
  setAbsolutePosition(x: number, y: number): this { this.getModel(); this.control.setAbsolutePosition(x, y); return this }
  setSize(width: number, height: number): this { this.getModel(); this.control.setSize(width, height); return this }
  bringToFront(): this { this.getModel(); this.control.bringToFront(); return this }
  bringForward(): this { this.getModel(); this.control.bringForward(); return this }
  sendBackward(): this { this.getModel(); this.control.sendBackward(); return this }
  sendToBack(): this { this.getModel(); this.control.sendToBack(); return this }
  setZOrder(index: number): this { this.getModel(); this.control.setZOrder(index); return this }
  getDateRange(): ITimelineRange | undefined {
    const model = this.optionalModel()
    return model ? readTimelineDateRange(this.workbook(), model) : undefined
  }
  getDateCondition(): ITimelineDateCondition | undefined {
    const model = this.optionalModel()
    return model ? readTimelineDateCondition(this.workbook(), model) : undefined
  }
  setDateRange(range: ITimelineRange | undefined, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> {
    return this.api.executeCommand(SLICER_PUBLIC_COMMANDS.dateRange, { unitId: this.unitId, slicerId: this.id, range,
      expectedRevision: context.expectedRevision, host: context.host })
  }
  async getDateBounds(): Promise<{ minDate?: number; maxDate?: number; dateSystem: 'date1900' | 'date1904' }> {
    const dateSystem = this.workbook().date1904 ? 'date1904' : 'date1900'
    const metadata = await this.getSourceMetadata()
    return { dateSystem, ...(metadata.timeline?.minDate === undefined ? {} : { minDate: metadata.timeline.minDate }),
      ...(metadata.timeline?.maxDate === undefined ? {} : { maxDate: metadata.timeline.maxDate }) }
  }
  /** Local combined properties transaction for Canvas/property-panel geometry and settings. */
  async updatePresentation(patch: TimelinePresentationPatch, context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<this> {
    const accepted = await this.api.executeCommand('sheet.command.update-slicer-presentation', { unitId: this.unitId, slicerId: this.id, patch,
      expectedRevision: context.expectedRevision, host: context.host })
    if (!accepted) throw new Error('时间线显示设置未能提交。')
    return this
  }
  clearFilter(context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> { return this.setDateRange(undefined, context) }
  remove(context: import('./slicer-command-context.ts').SlicerCommandContext = {}): Promise<boolean> { this.getModel(); return this.control.remove(context) }
}

export class FDocument {
  get regionalSettings() { return this.api.regionalSettings }
  workbenchUiHost() { return this.api.workbenchUiHost(this.id) }
  private readonly api: OfficeStudio
  readonly id: string

  constructor(api: OfficeStudio, id: string) {
    this.api = api
    this.id = id
  }

  getId(): string { return this.id }
  getCurrentUserId(): string { return this.api.currentUserId }
  getPermissionAdminUserId(): string { return this.api.permissionAdminUserId }
  getName(): string { return this.model().name }
  setName(name: string): this { this.api.executeCommandSync('doc.command.set-name', { unitId: this.id, name }); return this }
  getZoomRatio(): number { return this.api.getDocumentZoomRatio(this.id) }
  getCharacterPreferences(): DocumentCharacterPreferencesStore {
    this.model()
    return this.api.documentCharacterPreferences
  }
  setZoomRatio(ratio: number): boolean { return this.api.setDocumentZoomRatio(this.id, ratio) }
  onZoomRatioChanged(listener: (event: ApiEventMap['DocumentZoomRatioChanged']) => void): Disposable {
    return this.api.events.addEvent('DocumentZoomRatioChanged', event => { if (event.documentId === this.id) listener(event) })
  }
  /** Read the complete detached public snapshot without creating a savepoint. */
  getSnapshot(): UniverDocumentData {
    const model = this.model()
    assertDocumentSnapshotViewPermission(model, this.getCurrentUserId())
    return documentToPublicSnapshot(model, this.getModelRevision())
  }
  save(): UniverDocumentData {
    const snapshot = this.getSnapshot()
    this.api.markSaved('document', this.id)
    return snapshot
  }
  /** Project model snapshot for internal persistence and native Office writing. */
  saveModel(): DocumentModel {
    assertDocumentSnapshotViewPermission(this.model(), this.getCurrentUserId())
    const snapshot = saveDocumentSnapshot(this.model())
    this.api.markSaved('document', this.id)
    return snapshot
  }
  undo(): boolean { return this.api.undo('document', this.id) }
  redo(): boolean { return this.api.redo('document', this.id) }
  getText(): string { return documentBodyText(this.model()) }
  getBody(segmentId = ''): UniverDocumentBody { return this.getBodyData(segmentId).body }
  getBodyData(segmentId = ''): UniverDocumentBodyData { return this.getPublicBodyIndex(segmentId).data }
  getPublicBodyIndex(segmentId = ''): DocumentPublicBodyIndex {
    this.getSegmentModel(segmentId)
    return documentPublicBodyIndex(this.model(), this.getSegmentHost(segmentId))
  }
  getTextRange(startOffset: number, endOffset: number, segmentId = ''): FDocumentTextRange {
    assertDocumentPublicRange(this.getPublicBodyIndex(segmentId), startOffset, endOffset)
    return new FDocumentTextRange(this, startOffset, endOffset, segmentId)
  }
  insertText(index: number, text: string, segmentId = ''): boolean {
    const host = this.getSegmentHost(segmentId)
    this.getSegmentModel(segmentId)
    const selection = documentSelectionAtPublicBodyRange(this.model(), index, index, host)
    if (!selection && !segmentId && index === 0 && !textTargets(this.model()).length) {
      return this.api.executeCommandSync('doc.command.insert-text', { unitId: this.id, index, text, expectedRevision: this.getModelRevision() })
    }
    return !!selection && this.executeCommand({ id: 'doc.command.insert-text', text, ...(host ? { host } : {}) }, selection).changed
  }
  getParagraphInfos(segmentId = ''): DocumentParagraphInfo[] { return documentPublicParagraphs(this.getPublicBodyIndex(segmentId)) }
  getParagraphs(segmentId = ''): FDocumentParagraph[] { return this.getParagraphInfos(segmentId).map((item) => new FDocumentParagraph(this, item, segmentId)) }
  /** Paragraph APIs and the settings panel share the registered content command. */
  formatParagraph(paragraphId: string, patch: Partial<DocumentParagraphStyle>, segmentId = ''): boolean {
    const scoped = this.getSegmentModel(segmentId), target = textTargets(scoped).find(candidate => (candidate.cellBlockId ?? candidate.blockId) === paragraphId)
    if (!target) throw new RangeError(`Document paragraph with id ${paragraphId} not found`)
    const host = this.getSegmentHost(segmentId)
    const effectivePatch = Object.fromEntries(Object.entries(patch).filter(([, value]) => value !== undefined)) as Partial<DocumentParagraphStyle>
    return this.executeCommand({ id: 'doc.command.format-paragraph', patch: effectivePatch, ...(host ? { host } : {}) },
      { anchor: { ...target, offset: 0 }, focus: { ...target, offset: 0 } }).changed
  }
  getParagraph(paragraphId: string, segmentId = ''): FDocumentParagraph | null { return this.findParagraphs({ paragraphId }, segmentId)[0] ?? null }
  findParagraphs(query: string | { text?: string; paragraphId?: string; blockId?: string }, segmentId = ''): FDocumentParagraph[] {
    const match = typeof query === 'string' ? { text: query } : query
    return this.getParagraphInfos(segmentId).filter(info => (!match.text || info.text.includes(match.text))
      && (!match.paragraphId || info.paragraphId === match.paragraphId) && (!match.blockId || info.blockId === match.blockId))
      .map(info => new FDocumentParagraph(this, info, segmentId))
  }
  appendParagraph(text = '', segmentId = ''): FDocumentParagraph {
    const paragraphs = this.getParagraphInfos(segmentId), last = paragraphs.at(-1)
    this.insertText(last?.endOffset ?? 0, `${last && (last.text.length || paragraphs.length > 1) ? '\n' : ''}${text}`, segmentId)
    return this.getParagraphs(segmentId).at(-1)!
  }
  getModel(): DocumentModel { return this.model() }
  getLatexFormulas(segmentId = ''): FDocumentLatex[] { return documentLatexFormulaTargets(this, segmentId).map(item => new FDocumentLatex(this, item.rangeId, segmentId)) }
  getLatexFormula(rangeId: string, segmentId = ''): FDocumentLatex | null { return this.findLatexFormulas({ rangeId }, segmentId)[0] ?? null }
  getLatexFormulaAt(offset: number, segmentId = ''): FDocumentLatex | null {
    const id = documentLatexRangeAt(this.getPublicBodyIndex(segmentId), offset)
    return id ? this.getLatexFormula(id, segmentId) : null
  }
  findLatexFormulaByText(latex: string, segmentId = ''): FDocumentLatex | null { return this.findLatexFormulas(latex, segmentId)[0] ?? null }
  findLatexFormulas(query: string | DocumentLatexFindQuery, segmentId = ''): FDocumentLatex[] { return findDocumentLatexFormulas(this, query, segmentId) }
  insertLatexAtOffset(offset: number, latex: string, options: DocumentLatexInsertOptions = {}): FDocumentLatex | null { return insertDocumentLatexAtRange(this, offset, offset, latex, options) }
  insertLatexAtSelection(latex: string, options: DocumentLatexCreateOptions = {}): FDocumentLatex | null { return insertDocumentLatexAtSelection(this, latex, options) }
  /** Local UI binding; source drafts are not persisted until a formula command applies them. */
  bindLatexTemplateEditor(insert: (templateId: string) => boolean): DocumentLatexTemplateEditorHandle {
    return registerDocumentLatexTemplateEditor(this.api, this.id, insert)
  }
  /** Local UI binding; the fixed upstream contract is the popup operation IDs. */
  bindLatexPopupEditor(editor: DocumentLatexPopupEditor): DocumentLatexPopupEditorHandle {
    return registerDocumentLatexPopupEditor(this.api, this.id, editor)
  }
  bindImageEditor(editor: DocumentImageEditor): DocumentImageEditorHandle { return registerDocumentImageEditor(this.api, this.id, editor) }
  insertImage(options: DocumentImageInsertOptions): Promise<FDocumentImage | null> { return insertDocumentImage(this, options) }
  getImage(imageId: string, segmentId = ''): FDocumentImage | null { return getDocumentImage(this, imageId, segmentId) }
  getImages(segmentId = ''): FDocumentImage[] { return getDocumentImages(this, segmentId) }
  alignImages(images: readonly (FDocumentImage | string)[], align: DocumentImageAlign, segmentId = '', options: DocumentImageGeometryReadOptions = {}): boolean {
    return alignDocumentImages(this, images, align, segmentId, options)
  }
  transformImages(images: readonly (FDocumentImage | string)[], transform: DocumentImageBatchTransform, segmentId = '', options: DocumentImageGeometryReadOptions = {}): boolean {
    return transformDocumentImages(this, images, transform, segmentId, options)
  }
  arrangeImages(images: readonly (FDocumentImage | string)[], arrange: DocumentImageArrange | number, segmentId = ''): boolean {
    return arrangeDocumentImages(this, images, arrange, segmentId)
  }
  removeImages(images: readonly (FDocumentImage | string)[], segmentId = ''): boolean { return removeDocumentImages(this, images, segmentId) }
  openLatexPopup(request: Omit<DocumentLatexPopupOperationRequest, 'unitId'> = {}): boolean {
    return this.api.executeCommandSync(DOCUMENT_LATEX_POPUP_OPERATIONS.open, { ...request, unitId: this.id })
  }
  closeLatexPopup(): boolean {
    return this.api.executeCommandSync(DOCUMENT_LATEX_POPUP_OPERATIONS.close, { unitId: this.id })
  }
  insertLatexTemplate(templateId: string): boolean {
    return this.api.executeCommandSync(DOCUMENT_LATEX_TEMPLATE_OPERATION, { unitId: this.id, templateId })
  }
  getQuotes(segmentId = ''): FDocumentQuote[] { return documentStructuredBlocks(this, 'quote', segmentId).map(block => new FDocumentQuote(this, block.id, segmentId)) }
  getQuote(blockId: string, segmentId = ''): FDocumentQuote | null { return this.findQuotes({ blockId }, segmentId)[0] ?? null }
  getQuoteAt(offset: number, segmentId = ''): FDocumentQuote | null {
    const block = documentStructuredBlockAt(this, 'quote', offset, segmentId); return block ? new FDocumentQuote(this, block.id, segmentId) : null
  }
  findQuotes(query: string | DocumentStructuredFindQuery, segmentId = ''): FDocumentQuote[] { return findDocumentStructuredBlocks(this, 'quote', query, segmentId).map(block => new FDocumentQuote(this, block.id, segmentId)) }
  findQuoteByText(text: string, segmentId = ''): FDocumentQuote | null { return this.findQuotes(text, segmentId)[0] ?? null }
  insertQuote(paragraph: FDocumentParagraph, options?: DocumentQuoteInsertOptions): FDocumentQuote | null
  insertQuote(options?: DocumentQuoteInsertOptions): FDocumentQuote | null
  insertQuote(paragraphOrOptions: FDocumentParagraph | DocumentQuoteInsertOptions = {}, options: DocumentQuoteInsertOptions = {}): FDocumentQuote | null { return insertDocumentQuote(this, paragraphOrOptions, options) }
  getCallouts(segmentId = ''): FDocumentCallout[] { return documentStructuredBlocks(this, 'callout', segmentId).map(block => new FDocumentCallout(this, block.id, segmentId)) }
  getCallout(blockId: string, segmentId = ''): FDocumentCallout | null { return this.findCallouts({ blockId }, segmentId)[0] ?? null }
  getCalloutAt(offset: number, segmentId = ''): FDocumentCallout | null {
    const block = documentStructuredBlockAt(this, 'callout', offset, segmentId); return block ? new FDocumentCallout(this, block.id, segmentId) : null
  }
  findCallouts(query: string | DocumentStructuredFindQuery, segmentId = ''): FDocumentCallout[] { return findDocumentStructuredBlocks(this, 'callout', query, segmentId).map(block => new FDocumentCallout(this, block.id, segmentId)) }
  findCalloutByText(text: string, segmentId = ''): FDocumentCallout | null { return this.findCallouts(text, segmentId)[0] ?? null }
  insertCallout(paragraph: FDocumentParagraph, options?: DocumentCalloutInsertOptions): FDocumentCallout | null
  insertCallout(options?: DocumentCalloutInsertOptions): FDocumentCallout | null
  insertCallout(paragraphOrOptions: FDocumentParagraph | DocumentCalloutInsertOptions = {}, options: DocumentCalloutInsertOptions = {}): FDocumentCallout | null { return insertDocumentCallout(this, paragraphOrOptions, options) }
  getCodes(segmentId = ''): FDocumentCode[] { return documentStructuredBlocks(this, 'code', segmentId).map(block => new FDocumentCode(this, block.id, segmentId)) }
  getCode(blockId: string, segmentId = ''): FDocumentCode | null { return this.findCodes({ blockId }, segmentId)[0] ?? null }
  getCodeAt(offset: number, segmentId = ''): FDocumentCode | null {
    const block = documentStructuredBlockAt(this, 'code', offset, segmentId); return block ? new FDocumentCode(this, block.id, segmentId) : null
  }
  findCodes(query: string | DocumentStructuredFindQuery, segmentId = ''): FDocumentCode[] { return findDocumentStructuredBlocks(this, 'code', query, segmentId).map(block => new FDocumentCode(this, block.id, segmentId)) }
  findCodeByText(text: string, segmentId = ''): FDocumentCode | null { return this.findCodes(text, segmentId)[0] ?? null }
  insertCode(paragraph: FDocumentParagraph, options?: DocumentCodeInsertOptions): FDocumentCode | null
  insertCode(options?: DocumentCodeInsertOptions): FDocumentCode | null
  insertCode(paragraphOrOptions: FDocumentParagraph | DocumentCodeInsertOptions = {}, options: DocumentCodeInsertOptions = {}): FDocumentCode | null { return insertDocumentCode(this, paragraphOrOptions, options) }
  getColumnGroups(segmentId = ''): FDocumentColumnGroup[] { return documentStructuredBlocks(this, 'columns', segmentId).map(block => new FDocumentColumnGroup(this, block.id, segmentId)) }
  getColumnGroup(columnGroupId: string, segmentId = ''): FDocumentColumnGroup | null { return this.findColumnGroups({ columnGroupId }, segmentId)[0] ?? null }
  getColumnGroupAt(offset: number, segmentId = ''): FDocumentColumnGroup | null {
    const block = documentStructuredBlockAt(this, 'columns', offset, segmentId); return block ? new FDocumentColumnGroup(this, block.id, segmentId) : null
  }
  findColumnGroups(query: string | DocumentColumnFindQuery, segmentId = ''): FDocumentColumnGroup[] {
    const filter: DocumentColumnFindQuery = typeof query === 'string' ? { text: query } : query
    if (!filter || typeof filter !== 'object' || Array.isArray(filter)
      || [filter.columnGroupId, filter.text].some(value => value !== undefined && typeof value !== 'string')) throw new TypeError('多栏组查询参数无效。')
    return findDocumentStructuredBlocks(this, 'columns', { blockId: filter.columnGroupId, text: filter.text }, segmentId).map(block => new FDocumentColumnGroup(this, block.id, segmentId))
  }
  findColumnGroupByText(text: string, segmentId = ''): FDocumentColumnGroup | null { return this.findColumnGroups(text, segmentId)[0] ?? null }
  insertColumnGroup(columnCount: number, options: DocumentColumnInsertOptions = {}): FDocumentColumnGroup | null { return insertDocumentColumnGroup(this, columnCount, options) }
  getTables(segmentId = ''): FDocumentTable[] { return getDocumentTables(this, segmentId) }
  getTable(tableId: string, segmentId = ''): FDocumentTable | null {
    return findDocumentTables(this, { tableId }, segmentId)[0] ?? null
  }
  /** index is the table ordinal in recursive body order, not a text offset. */
  getTableAt(index: number, segmentId = ''): FDocumentTable | null {
    return Number.isSafeInteger(index) && index >= 0 ? getDocumentTables(this, segmentId)[index] ?? null : null
  }
  getTableAtSelection(): FDocumentTable | null { return getDocumentTableAtSelection(this) }
  findTableByText(text: string, segmentId = ''): FDocumentTable | null { return findDocumentTables(this, text, segmentId)[0] ?? null }
  findTables(query: string | DocumentTableFindQuery, segmentId = ''): FDocumentTable[] { return findDocumentTables(this, query, segmentId) }
  insertTable(rows: number, columns: number, options: DocumentTableInsertOptions = {}): FDocumentTable | null {
    return insertDocumentTable(this, rows, columns, options)
  }
  insertTableFromData(data: string[][], options: Omit<DocumentTableInsertOptions, 'values'> = {}): FDocumentTable | null {
    return insertDocumentTable(this, undefined, undefined, { ...options, values: data })
  }
  executeCommand(command: DocumentCommand, selection: DocumentSelection, sourceCommandId?: 'doc.command.insert-special-character'): DocumentEditorCommandResult {
    return this.api.executeDocumentEditorCommand(this.id, command, selection, sourceCommandId)
  }
  executeNoteCommand(request: DocumentNoteMutationRequest): DocumentNoteCommandResult { return this.api.executeDocumentNoteCommand(this.id, request) }
  insertFootnote(offset: number, content?: DocumentNoteContentInput): string | null { return this.insertNote('footnote', offset, content) }
  insertEndnote(offset: number, content?: DocumentNoteContentInput): string | null { return this.insertNote('endnote', offset, content) }
  getFootnotes(): DocumentNoteInfo[] { return this.getNotes('footnote') }
  getEndnotes(): DocumentNoteInfo[] { return this.getNotes('endnote') }
  deleteFootnote(noteId: string): boolean { return this.tryNoteMutation({ type: 'delete', kind: 'footnote', noteId })?.changed ?? false }
  deleteEndnote(noteId: string): boolean { return this.tryNoteMutation({ type: 'delete', kind: 'endnote', noteId })?.changed ?? false }
  getFootnoteSettings(sectionId?: string): PublicDocumentFootnoteSettings | undefined { return this.readNoteSettings('footnote', sectionId) as PublicDocumentFootnoteSettings | undefined }
  getEndnoteSettings(sectionId?: string): PublicDocumentEndnoteSettings | undefined { return this.readNoteSettings('endnote', sectionId) as PublicDocumentEndnoteSettings | undefined }
  setFootnoteSettings(settings: DocumentFootnoteSettings | PublicDocumentFootnoteSettings | null, sectionId?: string): boolean {
    return this.setNoteSettings('footnote', settings, sectionId)
  }
  setEndnoteSettings(settings: DocumentEndnoteSettings | PublicDocumentEndnoteSettings | null, sectionId?: string, suppressEndnotes?: boolean): boolean {
    return this.setNoteSettings('endnote', settings, sectionId, suppressEndnotes)
  }
  /** Pending input overrides; clearing color/highlight preserves an own undefined property. */
  getInputMarks(selection?: DocumentSelection, story?: DocumentStoryRange, host?: DocumentAuxiliaryBlockHost): RunPatch { return this.api.getDocumentInputMarks(this.id, selection, story, host) }
  getModelRevision(): number { return this.api.getDocumentRevision(this.id) }
  getHistoryGeneration(): number { return this.api.editor.history.captureSavepoint(this.id).generation }
  getFieldCacheAdapter() { return this.api.getDocumentFieldCacheAdapter(this.id) }
  getInputState() { return this.api.getDocumentInputState(this.id) }
  onInputStateChanged(listener: (event: ApiEventMap['DocumentInputStateChanged']) => void): Disposable {
    return this.api.events.addEvent('DocumentInputStateChanged', event => { if (event.unitId === this.id) listener(event) })
  }
  setInputSelection(selection: DocumentSelection, story?: DocumentStoryRange, host?: DocumentAuxiliaryBlockHost): boolean { return this.api.setDocumentInputSelection(this.id, selection, story, host) }
  clearInputMarks(): boolean { return this.api.clearDocumentInputMarks(this.id) }
  advanceInputMarks(selection: DocumentSelection, story?: DocumentStoryRange, host?: DocumentAuxiliaryBlockHost): boolean { return this.api.advanceDocumentInputMarks(this.id, selection, story, host) }
  clearFormatting(selection: DocumentSelection, story?: DocumentStoryRange): DocumentEditorCommandResult {
    return this.executeCommand({ id: 'doc.command.clear-formatting', ...(story ? { story } : {}) }, selection)
  }
  refreshFields(options: Partial<Omit<DocumentStoryFieldRefreshRequest, 'stories' | 'bodySelection'>> & { selection?: DocumentSelection } = {}): DocumentEditorCommandResult {
    const state = this.getInputState(), target = textTargets(this.model())[0]
    const selection = options.selection ?? (state?.location.kind === 'body' ? state.location.selection : undefined)
      ?? (target ? { anchor: { ...target, offset: 0 }, focus: { ...target, offset: 0 } } : undefined)
    if (!selection) throw new Error('文档没有可定位的字段更新范围。')
    return this.executeCommand({ id: 'doc.command.refresh-header-footer-fields',
      locale: options.locale ?? this.regionalSettings.locale.value.replaceAll('_', '-'), timestamp: options.timestamp ?? Date.now(),
      ...(options.onOpen === undefined ? {} : { onOpen: options.onOpen }) }, selection)
  }
  updateSelectedFields(selection: DocumentSelection, options: Partial<DocumentStoryFieldRefreshRequest> = {}): DocumentEditorCommandResult {
    return this.executeCommand({ id: 'doc.command.update-selected-fields',
      locale: options.locale ?? this.regionalSettings.locale.value.replaceAll('_', '-'), timestamp: options.timestamp ?? Date.now(),
      ...(options.stories === undefined ? {} : { stories: options.stories }),
      ...(options.bodySelection === undefined ? {} : { bodySelection: options.bodySelection }),
      ...(options.onOpen === undefined ? {} : { onOpen: options.onOpen }) }, selection)
  }
  getPermissionRules(): DocumentModel['permissionRules'] { return cloneModel(this.model().permissionRules ?? []) }
  setPermissionRules(rules: NonNullable<DocumentModel['permissionRules']>): this {
    this.api.executeCommandSync('doc.command.set-document-permission-rules', { unitId: this.id, rules })
    return this
  }
  /** Local source API. Returned resources are detached from the live document. */
  getCustomXmlParts(): DocumentCustomXmlPart[] {
    const model = this.model()
    assertDocumentPermission(model.permissionRules ?? [], { documentId: this.id, userId: this.getCurrentUserId() }, 'view')
    return normalizeDocumentCustomXmlParts(model.customXmlParts)
  }
  /** Replace sources and refresh mapped controls in one guarded content-history transaction. */
  setCustomXmlParts(parts: readonly DocumentCustomXmlPart[], options: DocumentCustomXmlPartsWriteOptions = {}): this {
    this.api.executeCommandSync(SET_DOCUMENT_CUSTOM_XML_PARTS_COMMAND, {
      expectedRevision: this.getModelRevision(), expectedHistoryGeneration: this.getHistoryGeneration(), ...options, unitId: this.id, parts,
    })
    return this
  }
  getContentControls(): FDocumentContentControl[] {
    return getDocumentContentControls(this.model()).map((control) => new FDocumentContentControl(this, control.id))
  }
  getContentControl(id: string): FDocumentContentControl | null {
    return getDocumentContentControl(this.model(), id) ? new FDocumentContentControl(this, id) : null
  }
  addContentControl(info: DocumentContentControlCreateInfo): FDocumentContentControl {
    const id = info.id ?? randomId('content-control')
    this.api.executeCommandSync('doc.command.add-content-control', { unitId: this.id, control: { ...info, id } })
    return new FDocumentContentControl(this, id)
  }
  replaceRange(startOffset: number, endOffset: number, text: string, segmentId = ''): boolean {
    const host = this.getSegmentHost(segmentId)
    this.getSegmentModel(segmentId)
    const selection = documentSelectionAtPublicBodyRange(this.model(), startOffset, endOffset, host)
    return !!selection && this.executeCommand({ id: 'doc.command.insert-text', text, ...(host ? { host } : {}) }, selection).changed
  }
  mutate(label: string, commandId: string, update: (model: DocumentModel) => void): boolean {
    return this.api.mutateDocument(this.id, commandId, label, update)
  }
  getSegmentModel(segmentId = ''): DocumentModel {
    if (!segmentId) {
      const model = this.model()
      assertDocumentPermission(model.permissionRules ?? [], { documentId: this.id, userId: this.getCurrentUserId() }, 'view')
      return model
    }
    const model = this.model(), host = this.getSegmentHost(segmentId)!, scoped = documentModelForBlockHost(model, host)
    const targets = textTargets(scoped), first = targets[0], last = targets.at(-1)
    assertDocumentPermission(model.permissionRules ?? [], { documentId: this.id, userId: this.getCurrentUserId() }, 'view', host.kind === 'note'
      ? documentNotePermissionTargets(model, host.noteKind, host.noteId)
      : [{ scope: 'entity', id: documentBlockHostId(host) }, ...(first && last ? permissionTargetsForSelection(scoped,
        { anchor: { ...first, offset: 0 }, focus: { ...last, offset: 0 } }) : [])])
    return scoped
  }
  getSegmentHost(segmentId = ''): DocumentAuxiliaryBlockHost | undefined { return resolveDocumentSegmentHost(this.model(), segmentId) }
  mutateSegment(label: string, commandId: string, update: (model: DocumentModel) => void, segmentId = ''): boolean {
    if (!segmentId) return this.mutate(label, commandId, update)
    const host = this.getSegmentHost(segmentId)!
    return this.mutate(label, commandId, model => {
      const scoped = documentModelForBlockHost(model, host, true)
      update(scoped); applyDocumentBlockHostModel(model, host, scoped)
    })
  }
  executeContentControlCommand<P extends Record<string, unknown>>(commandId: string, params: P): boolean {
    return this.api.executeCommandSync(commandId, { ...params, unitId: this.id })
  }

  private insertNote(kind: DocumentNoteKind, offset: number, input?: DocumentNoteContentInput): string | null {
    let selection: DocumentSelection | undefined, options: ReturnType<typeof documentNoteInputFromPublic>
    try {
      selection = documentSelectionAtPublicBodyRange(this.model(), offset, offset)
      options = documentNoteInputFromPublic(this.model(), input, { kind })
    } catch (error) { if (error instanceof TypeError || error instanceof RangeError) return null; throw error }
    if (!selection) return null
    const result = this.tryNoteMutation({ type: 'insert', kind, selection, ...(options === undefined ? {} : { options }) })
    return result?.changed ? result.noteId ?? null : null
  }
  private tryNoteMutation(request: DocumentNoteMutationRequest): DocumentNoteCommandResult | undefined {
    try { return this.executeNoteCommand(request) }
    catch (error) { if (error instanceof DocumentNoteCommandRejectedError) return undefined; throw error }
  }
  private getNotes(kind: DocumentNoteKind): DocumentNoteInfo[] {
    const model = this.model(), permission = { documentId: this.id, userId: this.getCurrentUserId() }
    assertDocumentPermission(model.permissionRules ?? [], permission, 'view')
    return getDocumentNotesInReferenceOrder(model, kind).map(note => {
      assertDocumentPermission(model.permissionRules ?? [], permission, 'view', documentNotePermissionTargets(model, kind, note.id))
      return documentNoteToPublic(model, note, kind)
    })
  }
  private setNoteSettings(kind: DocumentNoteKind, input: DocumentNoteSettingsInput | null, sectionId?: string, suppressEndnotes?: boolean): boolean {
    let normalized: ReturnType<typeof documentNoteSettingsFromPublic>
    try { normalized = documentNoteSettingsFromPublic(this.model(), input) }
    catch (error) { if (error instanceof TypeError || error instanceof RangeError) return false; throw error }
    return this.tryNoteMutation({ type: 'settings', kind, ...normalized, ...(sectionId === undefined ? {} : { sectionId }),
      ...(suppressEndnotes === undefined ? {} : { suppressEndnotes }) })?.changed ?? false
  }
  private readNoteSettings(kind: DocumentNoteKind, sectionId?: string): PublicDocumentFootnoteSettings | PublicDocumentEndnoteSettings | undefined {
    const model = this.model()
    assertDocumentPermission(model.permissionRules ?? [], { documentId: this.id, userId: this.getCurrentUserId() }, 'view', sectionId === undefined ? [] : [{ scope: 'section', id: sectionId }])
    const settings = getDocumentNoteSettings(model, kind, sectionId)
    return settings === undefined ? undefined : documentNoteSettingsToPublic(model, settings)
  }

  private model(): DocumentModel {
    const model = this.api.getDocumentModel(this.id)
    if (!model) throw new Error(`Document not found: ${this.id}`)
    return model
  }
}

export class FDocumentContentControl {
  private readonly document: FDocument
  readonly id: string

  constructor(document: FDocument, id: string) {
    this.document = document
    this.id = id
  }
  getId(): string { return this.id }
  getInfo(): DocumentContentControlInfo {
    const control = getDocumentContentControl(this.document.getModel(), this.id)
    if (!control) throw new Error(`Content control not found: ${this.id}`)
    return control
  }
  getType(): DocumentContentControlType { return this.getInfo().type }
  getText(): string { return contentControlText(this.document.getModel(), this.getInfo()) }
  getRange(): { blockId: string; startOffset: number; endOffset: number } { return contentControlRange(this.getInfo()) }
  setValue(value: string): this { this.document.executeContentControlCommand('doc.command.set-content-control-value', { controlId: this.id, value }); return this }
  replacePicture(image: { src: string; alt: string; width: number; height: number }): boolean { return this.document.executeContentControlCommand('doc.command.replace-picture-content-control-image', { controlId: this.id, image }) }
  setChecked(checked: boolean): this { this.document.executeContentControlCommand('doc.command.set-content-control-checked', { controlId: this.id, checked }); return this }
  resetPlaceholder(): boolean { return this.document.executeContentControlCommand('doc.command.reset-content-control-placeholder', { controlId: this.id }) }
  duplicateRepeatingItem(): boolean { return this.document.executeContentControlCommand('doc.command.duplicate-repeating-section-item', { controlId: this.id }) }
  deleteRepeatingItem(): boolean { return this.document.executeContentControlCommand('doc.command.delete-repeating-section-item', { controlId: this.id }) }
  selectItem(value: string): this { this.document.executeContentControlCommand('doc.command.select-content-control-item', { controlId: this.id, value }); return this }
  update(patch: DocumentContentControlPatch): this { this.document.executeContentControlCommand('doc.command.update-content-control', { controlId: this.id, patch }); return this }
  remove(): boolean { return this.document.executeContentControlCommand('doc.command.remove-content-control', { controlId: this.id }) }
}

export class FDocumentTextRange {
  private readonly document: FDocument
  readonly startOffset: number
  readonly endOffset: number
  readonly segmentId: string

  constructor(document: FDocument, startOffset: number, endOffset: number, segmentId = '') {
    this.document = document
    this.startOffset = startOffset
    this.endOffset = endOffset
    this.segmentId = segmentId
  }
  getRange(): { startOffset: number; endOffset: number; segmentId: string } { return { startOffset: this.startOffset, endOffset: this.endOffset, segmentId: this.segmentId } }
  getText(): string {
    const index = this.document.getPublicBodyIndex(this.segmentId)
    assertDocumentPublicRange(index, this.startOffset, this.endOffset)
    return index.data.body.dataStream.slice(this.startOffset, this.endOffset)
  }
  setText(text: string): boolean { return this.document.replaceRange(this.startOffset, this.endOffset, text, this.segmentId) }
  replaceWithLatex(latex: string, options: DocumentLatexCreateOptions = {}): FDocumentLatex | null {
    return insertDocumentLatexAtRange(this.document, this.startOffset, this.endOffset, latex, { ...options, segmentId: this.segmentId })
  }
  getTextStyleRuns(): DocumentTextStyleRun[] {
    const index = this.document.getPublicBodyIndex(this.segmentId), range = documentPublicRangeToFlat(index, this.startOffset, this.endOffset)
    return getDocumentTextStyleRuns(index.model, range.startOffset, range.endOffset).map(run => ({ ...run,
      startOffset: Math.max(this.startOffset, documentPublicOffsetAtFlat(index, run.startOffset)),
      endOffset: Math.min(this.endOffset, documentPublicOffsetAtFlat(index, run.endOffset, 'end')) }))
  }
  getTextStyle(): DocumentTextStyle {
    const index = this.document.getPublicBodyIndex(this.segmentId), range = documentPublicRangeToFlat(index, this.startOffset, this.endOffset)
    return getCommonDocumentTextStyle(index.model, range.startOffset, range.endOffset)
  }
  getExplicitTextStyleRuns(): UniverDocumentTextStyleRun[] {
    return publicDocumentExplicitTextStyleRuns(this.document.getPublicBodyIndex(this.segmentId), this.startOffset, this.endOffset)
  }
  getCommonExplicitTextStyle(): UniverDocumentTextStyle {
    return commonPublicDocumentExplicitTextStyle(this.document.getPublicBodyIndex(this.segmentId), this.startOffset, this.endOffset)
  }
  describe(): UniverDocumentTextRangeDescription {
    return { ...this.getRange(), text: this.getText(), length: this.endOffset - this.startOffset,
      explicitTextStyleRuns: this.getExplicitTextStyleRuns(), commonExplicitTextStyle: this.getCommonExplicitTextStyle() }
  }
  setTextStyle(style: UniverDocumentTextStyle | DocumentTextStylePatch): boolean {
    if (this.startOffset === this.endOffset) return false
    const local = isLocalDocumentTextStyle(style), patch = local ? style : validatePublicDocumentTextStyle(style)
    const index = this.document.getPublicBodyIndex(this.segmentId), range = documentPublicRangeToFlat(index, this.startOffset, this.endOffset)
    return this.document.mutateSegment('Set document text style', 'doc.command.set-text-style', (model) => {
      if (local) setDocumentTextStyle(model, range.startOffset, range.endOffset, patch as DocumentTextStylePatch)
      else setPublicDocumentTextStyle(model, index, this.startOffset, this.endOffset, patch as UniverDocumentTextStyle)
    }, this.segmentId)
  }
}

export class FDocumentParagraph {
  private readonly document: FDocument
  private readonly info: DocumentParagraphInfo
  readonly segmentId: string

  constructor(document: FDocument, info: DocumentParagraphInfo, segmentId = '') {
    this.document = document
    this.info = info
    this.segmentId = segmentId
  }
  getId(): string { return this.info.paragraphId }
  getDocumentId(): string { return this.document.getId() }
  getInfo(): DocumentParagraphInfo {
    const current = this.document.getParagraphInfos(this.segmentId).find(info => info.paragraphId === this.info.paragraphId)
    if (!current) throw new Error(`Document paragraph with id ${this.info.paragraphId} not found`)
    return { ...current }
  }
  getText(): string { return this.getInfo().text }
  getTextRange(): FDocumentTextRange {
    const info = this.getInfo()
    return this.document.getTextRange(info.startOffset, info.endOffset, this.segmentId)
  }
  setText(text: string): boolean { return this.getTextRange().setText(text) }
  appendText(text: string): boolean { return this.document.insertText(this.getInfo().endOffset, text, this.segmentId) }
  appendLatex(latex: string, options: DocumentLatexCreateOptions = {}): FDocumentLatex | null {
    return this.document.insertLatexAtOffset(this.getInfo().endOffset, latex, { ...options, segmentId: this.segmentId })
  }
  getTextStyle(): DocumentTextStyle { return this.getTextRange().getTextStyle() }
  setTextStyle(style: UniverDocumentTextStyle | DocumentTextStylePatch): boolean { return this.getTextRange().setTextStyle(style) }
  getParagraphStyle(): DocumentParagraphStyle {
    const style = getDocumentParagraphStyle(this.document.getSegmentModel(this.segmentId), this.identity())
    if (!style) throw new Error(`Document paragraph with id ${this.info.paragraphId} not found`)
    return style
  }
  setParagraphStyle(style: Partial<DocumentParagraphStyle>): boolean {
    return this.document.formatParagraph(this.info.paragraphId, style, this.segmentId)
  }

  private identity(): DocumentParagraphIdentity {
    return { paragraphId: this.info.paragraphId }
  }
}

export class FBlob {
  private value: Blob | undefined
  constructor(value?: Blob | Uint8Array | string, contentType = 'text/plain') {
    if (value instanceof Blob) this.value = value
    else if (typeof value === 'string') this.value = new Blob([value], { type: contentType })
    else if (value) this.setBytes(value)
  }
  copyBlob(): FBlob { return new FBlob(this.value) }
  getAs(contentType: string): FBlob { const blob = this.copyBlob(); blob.setContentType(contentType); return blob }
  getDataAsString(charset?: string): Promise<string> {
    if (!this.value) return Promise.resolve('')
    if (!charset) return this.value.text()
    return this.value.arrayBuffer().then((buffer) => new TextDecoder(charset).decode(buffer))
  }
  getBytes(): Promise<Uint8Array> {
    if (!this.value) return Promise.reject(new Error('Blob is undefined or null.'))
    return this.value.arrayBuffer().then((buffer) => new Uint8Array(buffer))
  }
  setBytes(bytes: Uint8Array): this { this.value = new Blob([Uint8Array.from(bytes)]); return this }
  setDataFromString(data: string, contentType = 'text/plain'): this { this.value = new Blob([data], { type: contentType }); return this }
  getContentType(): string | undefined { return this.value?.type }
  setContentType(contentType: string): this { if (this.value) this.value = this.value.slice(0, this.value.size, contentType); return this }
  getBlob(): Blob { if (!this.value) throw new Error('Blob is undefined or null.'); return this.value }
  getSize(): number { return this.value?.size ?? 0 }
}

function normalizeRange(range: FRangeCoordinates): FRangeCoordinates {
  if (!range || typeof range !== 'object' || Array.isArray(range)) throw new TypeError('Range specification is invalid')
  if (range.rangeType !== undefined && ![0, 1, 2, 3].includes(range.rangeType)) throw new TypeError('Range type is invalid')
  const next = { startRow: range.startRow, startColumn: range.startColumn, endRow: range.endRow, endColumn: range.endColumn,
    ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) }
  validateIndex(next.startRow, 'startRow')
  validateIndex(next.startColumn, 'startColumn')
  validateIndex(next.endRow, 'endRow')
  validateIndex(next.endColumn, 'endColumn')
  if (next.endRow < next.startRow || next.endColumn < next.startColumn) throw new RangeError('Range end must not precede its start')
  if (next.endRow >= 1_048_576 || next.endColumn >= 16_384) throw new RangeError('Range exceeds the Office row or column limit')
  return next
}

function normalizePivotSourceInfo(source: PivotTableSourceInfo, workbookId: string): PivotTableSourceInfo {
  if (source.unitId !== workbookId) throw new Error('Pivot source must belong to the target workbook')
  const sheet = source.subUnitId
  if (!sheet || !source.sheetName) throw new TypeError('Pivot source worksheet identity is required')
  const range = source.range
  for (const [name, value] of Object.entries({
    startRow: range.startRow,
    startColumn: range.startColumn,
    endRow: range.endRow,
    endColumn: range.endColumn,
  })) validateIndex(value, `pivot source ${name}`)
  if (range.endRow < range.startRow || range.endColumn < range.startColumn) throw new RangeError('Invalid pivot source range')
  return { ...source, range: { ...range } }
}

function normalizePivotAnchorInfo(anchor: PivotTableAnchorInfo, workbookId: string): PivotTableAnchorInfo {
  if (anchor.unitId !== workbookId) throw new Error('Pivot anchor must belong to the target workbook')
  validateIndex(anchor.row, 'pivot anchor row')
  validateIndex(anchor.col, 'pivot anchor column')
  return { ...anchor }
}

function toCellRange(range: FRangeCoordinates): CellRange {
  const normalized = normalizeToCellRange(range)
  return normalized
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  return left.startRow <= right.endRow && left.endRow >= right.startRow
    && left.startCol <= right.endCol && left.endCol >= right.startCol
}

function normalizeToCellRange(range: FRangeCoordinates): CellRange {
  const normalized = normalizeRange(range)
  return { startRow: normalized.startRow, startCol: normalized.startColumn, endRow: normalized.endRow, endCol: normalized.endColumn }
}

function parseA1Range(value: string): FRangeCoordinates {
  return parseSheetA1Range(value)
}

function parseChartA1Range(value: string): ChartSource['range'] {
  const range = parseA1Range(value)
  return { startRow: range.startRow, startCol: range.startColumn, endRow: range.endRow, endCol: range.endColumn }
}

function parseA1Cell(value: string): { row: number; column: number } {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(value.trim())
  if (!match) throw new TypeError(`Invalid A1 cell: ${value}`)
  let column = 0
  for (const letter of match[1]!.toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64
  const row = Number(match[2])
  if (row < 1) throw new RangeError(`Invalid A1 row: ${row}`)
  return { row: row - 1, column: column - 1 }
}

function cellAddress(row: number, column: number): string {
  let letters = ''
  for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + ((value - 1) % 26)) + letters
  return `${letters}${row + 1}`
}

function validateIndex(value: number, name: string, allowZero = false): void {
  if (!Number.isSafeInteger(value) || value < (allowZero ? 0 : 0)) throw new RangeError(`${name} must be a non-negative safe integer`)
}

function validateSize(value: number, name: string): void {
  if (!Number.isSafeInteger(value) || value < 1) throw new RangeError(`${name} must be a positive safe integer`)
}

function assertMatrixSize(values: unknown[][], rows: number, columns: number): void {
  if (values.length !== rows || values.some((row) => row.length !== columns)) throw new RangeError(`Expected a ${rows}x${columns} value matrix`)
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' ? value as Record<string, unknown> : {}
}
