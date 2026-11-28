import { createPinia } from 'pinia'
import type { Pinia } from 'pinia'
import { shallowReactive, watch } from 'vue'
import { OfficeRegionalSettings } from '../i18n/regional-settings.ts'
import type { ChartModel, CellRange, DocumentModel, OfficeKind, PivotAggregation, PivotDateGroupType, PivotDateGroupingMode, PivotFieldArea, PivotFieldFilter, PivotManualGroup, PivotNumberGroup, PivotTableOptions, SheetFilterCriteria, SheetRangeProtectionRule, SheetWorksheetProtectionRule, SheetSortCriterion, SheetSlicerConnections, SlicerTimelineRange, WorkbookModel } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { createCell } from '../domain/model'
import { registerSparklineCommands } from '../features/sparkline/sparkline-command.ts'
import { publishWorkbookSparklineChanges } from './sparkline-events.ts'
import { registerPublicSparklineCommands } from './sparkline-public-commands.ts'
import { applySheetFontFormatCommand, assertSheetFontFormatRanges, fontFormatActionForCommand, SHEET_FONT_FORMAT_COMMAND_IDS, sheetFontFormatCommandInput, sheetFontFormatInputRanges, type SheetFontFormatCommandParams } from '../features/sheet-workbench/sheet-font-format-command.ts'
import { applySheetNumberFormatCommand, numberFormatActionForCommand, SHEET_NUMBER_FORMAT_ALIASES, SHEET_NUMBER_FORMAT_COMMANDS, type SheetNumberFormatCommandInput } from '../features/sheet-workbench/sheet-number-format-command.ts'
import { applySheetLayoutFormatCommand, assertSheetLayoutFormatRanges, layoutFormatActionForCommand, SHEET_LAYOUT_FORMAT_COMMAND_IDS, sheetLayoutFormatInputRanges, type SheetLayoutFormatCommandInput } from '../features/sheet-workbench/sheet-layout-format-command.ts'
import { applySheetBorderFormatCommand, borderFormatActionForCommand, createSheetBorderFormatState, SHEET_BORDER_FORMAT_COMMAND_IDS, sheetBorderFormatCommandInput, sheetBorderFormatInputRanges, type SheetBorderFormatCommandInput, type SheetBorderFormatState } from '../features/sheet-workbench/sheet-border-format-command.ts'
import { SET_SHEET_HYPERLINK_COMMAND, REMOVE_SHEET_HYPERLINK_COMMAND, setSheetCellHyperlink, removeSheetCellHyperlink, type SheetHyperlinkDraft } from '../features/sheet-workbench/sheet-hyperlink.ts'
import { equivalentModelData } from '../platform/commands/model-equivalence'
import { OPEN_FEATURE_SEARCH, TOGGLE_SHORTCUT_PANEL, type WorkbenchUiHost, type WorkbenchUiOperationHandler } from '../platform/ribbon/workbench-operations'
import type { DocumentCommand } from '../features/document/commands/document-commands.ts'
import { prepareDocumentStoryFieldRefresh, type DocumentStoryFieldRefreshPlan, type DocumentStoryFieldRefreshRequest } from '../features/document/services/document-story-fields.ts'
import { applyDocumentFieldRefreshPlan, createDocumentSnapshotHistory, prepareDocumentHistoryCacheRebase, prepareDocumentSnapshotHistoryAmend } from '../features/document/services/document-history-cache.ts'
import type { DocumentFieldCacheAdapter, DocumentFieldCacheState } from '../features/document/services/document-field-open-refresh.ts'
import { clampPoint, createParagraph, runText, runsForTarget, textTargets, type DocumentSelection, type RunPatch } from '../features/document/model/document-model.ts'
import { documentEffectiveZoomRatio, equivalentDocumentContent, preserveDocumentZoomRatio } from '../features/document/services/document-zoom'
import { DOCUMENT_ZOOM_COMMANDS, registerDocumentZoomCommands, resolveDocumentZoomCommand, type DocumentZoomOperationParams } from './document-zoom-commands'
import { DOCUMENT_PAGE_SETUP_COMMAND, registerDocumentPageSetupCommand } from './document-page-setup-commands'
import { DOCUMENT_TOC_INSERT_COMMAND, registerDocumentTocInsertCommand } from './document-toc-commands'
import { DOCUMENT_TOC_UPDATE_COMMAND, DOCUMENT_TOC_DELETE_COMMAND, registerDocumentTocActionCommands } from './document-toc-action-commands'
import { DOCUMENT_TOC_DIALOG_OPERATION, registerDocumentTocDialogOperation, disposeDocumentTocDialogEditors } from './document-toc-dialog-editors'
import { SET_DOCUMENT_HEADING_COLLAPSED_COMMAND, registerDocumentHeadingCollapseCommand } from './document-heading-commands'
import { disposeDocumentTocLayouts } from './document-toc-layout'
import { cancelDocumentTocPagination } from './document-toc-pagination'
import { registerDocumentHeaderFooterCommands } from './document-header-footer-commands'
import { captureDocumentHeaderFooterEditorTarget, documentHeaderFooterEditorTargetCurrent, disposeDocumentHeaderFooterEditors } from './document-header-footer-editors'
import { isDocumentHeaderFooterCommand, resolveDocumentHeaderFooterRequest } from '../features/document/controllers/document-header-footer-editor'
import { workbookCommandGuard } from './workbook-command-guards'
import { cloneDocumentSelection } from '../features/document/model/document-table-selection.ts'
import type { DocumentStoryRange } from '../features/document/model/document-story.ts'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host.ts'
import { documentModelForBlockHost, validateDocumentBlockHost } from '../features/document/services/document-block-host.ts'
import { useDocumentInputStore } from '../features/document/stores/document-input-store.ts'
import { useDocumentCharacterPreferencesStore, type DocumentCharacterPreferencesStore } from '../features/document/stores/document-character-preferences-store'
import { ADVANCE_DOCUMENT_INPUT_MARKS_OPERATION, CLEAR_DOCUMENT_INPUT_MARKS_OPERATION, SET_DOCUMENT_INPUT_SELECTION_OPERATION,
  documentInputLocation, mergeDocumentInputMarks, sameDocumentInputLocation, sanitizeDocumentInputSelection, type DocumentInputLocation,
  type DocumentInputSession } from '../features/document/model/document-input-session.ts'
import { cloneDocumentInputSession } from '../features/document/model/document-input-session.ts'
import { parseDocumentInputSelectionRequest, prepareCollapsedDocumentInputFormat, prepareDocumentStoryInputAdvance, resolveDocumentInputLocation,
  type DocumentInputSelectionRequest } from './document-input-commands.ts'
import { documentSelectionAtBodyRange } from './document-indexed-input.ts'
import { documentContentControlCommandIsTextInput } from '../features/document/services/content-control-editing.ts'
import { applyDocumentEditorCommand, assertDocumentEditorCommandRequest, DOCUMENT_EDITOR_COMMAND_LABELS, resolveDocumentEditorSelection,
  type DocumentEditorCommandRequest, type DocumentEditorCommandResult } from './document-editor-commands.ts'
import type { DocumentImage } from '../features/document/model/document-model'
import { useEditorStore } from '../stores/editorStore'
import { applyCellValue, documentFromInput, workbookFromInput, type WorkbookInput } from './model-utils'
import { prepareWorkbookSheetImageInput, type SheetImageInputPreparationOptions } from '../features/sheet-images/sheet-image-input-preparation.ts'
import { EventHub, type ApiEventMap, type ApiRange, type ApiSheetSelection, type Disposable } from './events'
import { registerSheetRangeEditorCommands } from './sheet-range-editor-commands.ts'
import { registerSheetCellInputCommands } from './sheet-cell-input-commands.ts'
import { registerSheetFilterPanelCommands, disposeSheetFilterPanelEditors } from './sheet-filter-panel-commands.ts'
import { registerSheetFunctionCommands, disposeSheetFunctionEditors } from './sheet-function-commands.ts'
import { sheetFilterCommandContext } from './sheet-filter-command-context.ts'
import { registerSheetCellInlineFormatCommands } from './sheet-cell-inline-format-commands.ts'
import { registerSheetDefaultStyleCommands } from './sheet-default-style-commands.ts'
import { registerSheetRangeThemeCommands } from './sheet-range-theme-commands.ts'
import { registerSheetImageCommands } from './sheet-image-commands.ts'
import { registerSheetDrawingCommands } from './sheet-drawing-commands.ts'
import type { SlicerCommandContext } from './slicer-command-context.ts'
import { reconcileWorkbookDrawingStates } from '../features/sheet-drawings/sheet-drawing-model-lifecycle.ts'
import { refreshWorkbookCellStyles } from '../features/sheet-workbench/sheet-style-inheritance.ts'
import { registerSheetClipboardCommands } from './sheet-clipboard-commands.ts'
import { registerPivotChartCommands } from './pivot-chart-commands.ts'
import { registerPivotChartFieldCommands } from './pivot-chart-field-commands.ts'
import { registerPivotHierarchyCommands } from './pivot-hierarchy-commands.ts'
import { registerPivotFilterCommands } from './pivot-filter-commands.ts'
import type { PivotReportCommandContext } from '../features/pivot/services/pivot-report-api-types.ts'
import type { PivotChartCommandContext } from '../features/pivot/services/pivot-chart-api-types.ts'
import type { PivotRecordSelection, PivotShowDataAsInfo, PivotTableModel } from '../domain/model.ts'
import { assertPivotReportEditPermission } from '../features/pivot/services/pivot-report-permissions.ts'
import { synchronizeLinkedPivotChartReports } from '../features/pivot/services/pivot-chart-report-lifecycle.ts'
import { registerSheetFillCommands } from './sheet-fill-commands.ts'
import { registerSheetLifecycleCommands } from './sheet-lifecycle-commands.ts'
import { CONTENT_HISTORY_COMMAND_IDS, CONTENT_HISTORY_COMMANDS, ContentHistoryCanceledError, contentHistoryCommandAction, resolveContentHistoryTarget } from './content-history-commands.ts'
import { MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND, parseSheetSelectionOperation, parseSheetSelectionShortcut, SELECT_SHEET_RANGE_COMMAND, SET_SHEET_SELECTIONS_OPERATION } from './sheet-selection-commands.ts'
import { registerSheetAxisDimensionCommands } from './sheet-axis-dimension-commands.ts'
import { registerSheetAxisVisibilityCommands } from './sheet-axis-visibility-commands.ts'
import { registerSheetAxisStructureCommands } from './sheet-axis-structure-commands.ts'
import { registerSheetCellShiftCommands } from './sheet-cell-shift-commands.ts'
import { registerSheetConditionalFormattingCommands } from './sheet-conditional-formatting-commands.ts'
import { registerSheetTableCommands } from './sheet-table-commands.ts'
import { registerSheetValidationInputCommands } from './sheet-validation-input-commands.ts'
import { registerSheetDataValidationCommands } from './sheet-data-validation-commands.ts'
import { registerSheetValidationDropdownOperations } from './sheet-validation-dropdown-operations.ts'
import { assertWorkbookConditionalFormatBudget } from '../features/conditional-formatting/conditional-format-workbook-budget.ts'
import { refreshWorkbookAutoRowHeights } from '../features/sheet-workbench/sheet-axis-dimensions.ts'
import type { SheetTextMeasure } from '../features/sheet-workbench/sheet-cell-text-layout.ts'
import { measureSpreadsheetText } from '../rendering/spreadsheet/text-metrics.ts'
import { registerSheetFormatPainterCommands } from './sheet-format-painter-commands.ts'
import { useSheetFormatPainterStore } from '../features/sheet-workbench/sheet-format-painter-store.ts'
import { useSheetSelectionStore } from '../features/sheet-workbench/sheet-selection-store.ts'
import { useSheetClipboardStore } from '../features/sheet-workbench/sheet-clipboard-store.ts'
import { useSheetFillStore } from '../features/sheet-workbench/sheet-fill-store.ts'
import type { HistorySavepoint } from '../platform/history/history-types.ts'
import { normalizeSheetSelection, resolveVisibleSheetSelection, type SheetSelectionOptions, type SheetSelectionState } from '../features/sheet-workbench/sheet-selection.ts'
import { planSheetSelectionShortcut } from '../features/sheet-workbench/sheet-selection-navigation.ts'
import {
  createDetachedPivotReport,
  createPivotTable,
  refreshPivotTable,
  removePivotFields,
  setPivotField,
  setPivotFieldSetting,
  setPivotShowDataAs,
  setPivotLabelFilter,
  setPivotOptions,
  setPivotValuesPosition,
  setPivotSort,
  setPivotValueFilter,
  setPivotNumberGroup,
  setPivotDateGroupingMode,
  setPivotDateGroupType,
  setPivotExcelDateGroup,
  setPivotManualGroups,
  type PivotTableAnchorInput,
  type PivotTableSourceInput,
} from '../features/pivot/services/pivot-engine'
import { normalizePivotValuePositionInput } from '../features/pivot/services/pivot-values-position.ts'
import { createChartDefinition, updateChartDefinition, type ChartConfig, type ChartUpdatePatch } from '../features/charts/chart-engine'
import { publishChartImageResourceCandidate, pruneChartImageResources } from '../features/charts/chart-image-resources.ts'
import { CHART_OBJECT_COMMANDS, PASTE_COPIED_CHART_COMMAND, assertChartObjectAction, assertChartObjectRevision, chartObjectSheetId,
  chartObjectUIHost, chartObjectUISession, clearChartObjectUIHosts, isPrimaryChartObject, resolveChartObjectTarget, type ChartObjectCommandInput, type ChartObjectPasteInput } from '../features/charts/chart-object-commands.ts'
import { assertChartObjectClipboardSource, captureChartObjectClipboard, chartObjectClipboardContent, chartObjectPasteDefinition, createChartObjectId } from '../features/charts/chart-object-clipboard.ts'
import { assertChartSourceImportSource, planChartSourceImportAsync, type ChartSourceImportPlan } from '../features/charts/chart-source-import.ts'
import { normalizeChartEditPanelRoute, planChartEditPanelRoute } from '../features/charts/chart-edit-panel-route.ts'
import { chartObjectClipboardStore } from '../features/charts/chart-object-clipboard-store.ts'
import { CONFIGURE_PIVOT_COMMAND, configurePivotFromPanel, type ConfigurePivotFromPanelInput } from '../features/pivot/services/pivot-panel-state'
import { clearSlicerFilter, createSlicer, patchSlicerSettings, removeSlicer, setSlicerConnections, setSlicerItems, setSlicerItemStates, setSlicerRecordSelection, setSlicerTimelineRange, synchronizeSlicerSelections, updateSlicerPresentation, type SlicerCreateInput, type SlicerPresentationInputPatch } from '../features/slicer/slicer-engine'
import { patchSlicerCacheOptions } from '../features/slicer/slicer-cache-options.ts'
import type { SheetSlicerSettings } from '../features/slicer/slicer-settings.ts'
import { patchTimelineSettings, type TimelineSettingsInputPatch } from '../features/slicer/timeline-settings.ts'
import { setTimelineDateCondition, setTimelineDateRange } from '../features/slicer/timeline-filter.ts'
import { normalizeTimelineDateCondition, normalizeTimelineDateRange } from '../features/slicer/timeline-date-filter.ts'
import type { ITimelineDateCondition, ITimelineRange } from '../domain/timeline-filters.ts'
import { createSheetSlicerFromInfo, updateSheetSlicerFromInfo } from '../features/slicer/slicer-public-control.ts'
import { SLICER_PUBLIC_COMMANDS, sheetSlicerPlacementPatch, type ISheetSlicerInfo, type ISheetTimelineInfo,
  type SheetSlicerPlacement } from '../features/slicer/slicer-public-info.ts'
import { normalizeSlicerCaches, planSlicerConnections, SlicerConnectionError } from '../features/slicer/slicer-connections.ts'
import { SlicerMemberQueryService, type SlicerMemberQueryOptions } from '../features/slicer/slicer-member-query.ts'
import { reconcileSlicerConnectionResources } from '../features/slicer/slicer-connections-lifecycle.ts'
import { assertSlicerFilterPermission, assertSlicerViewPermission } from '../features/slicer/slicer-permissions.ts'
import { recalculateWorkbookFormulas } from '../features/formula/index.ts'
import { clearFilter, refreshAutoFilter, setAutoFilter, setFilterCriteria, sortRange, type SortRangeOptions } from '../features/sheet-sort-filter/sort-filter-engine'
import { CONFIGURE_SHEET_FILTER_COMMAND, configureFilterFromPanel, type ConfigureSheetFilterInput } from '../features/sheet-sort-filter/sheet-data-panel-state'
import { OfficeComputeClient } from '../compute/client.ts'
import { ComputeError } from '../compute/errors.ts'
import { NumericIndexService } from '../features/sheet-compute/numeric-index-service.ts'
import { assertChartReadPermission, ChartStatisticsService, type ChartAuthoringPlan, type ChartStatisticsServiceOptions } from '../features/charts/services/chart-statistics-service.ts'
import { translateFormulaReferences } from '../features/formula/formula-references.ts'
import { registerSheetOutlineCommands } from './sheet-outline-commands.ts'
import { validateSheetDataValidationChanges } from '../features/sheet-validation/validation-engine.ts'
import { sheetValidationSourceReader } from '../features/sheet-validation/validation-value-read.ts'
import { assertWorksheetProtectionChanges, canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection, removeSheetRangeProtectionRule, setSheetRangeProtectionRule, setWorksheetProtectionRule } from '../features/sheet-protection/range-protection.ts'
import { assertDocumentPermission, permissionTargetsForRange, validateDocumentPermissionRules } from '../features/document/permissions/document-permissions.ts'
import { registerDocumentContentControlCommands } from './document-content-control-commands'
import { registerDocumentCustomXmlCommands } from './document-custom-xml-commands'
import { registerDocumentLatexCommands } from './document-latex-commands'
import { disposeDocumentLatexTemplateEditors } from './document-latex-template-editors'
import { disposeDocumentLatexPopupEditors } from './document-latex-popup-editors'
import { registerDocumentImageCommands } from './document-image-commands'
import { disposeDocumentImageEditors, cancelDocumentImageTasksForSource } from './document-image-editors'
import { documentNoteCommandId, DocumentNoteCommandRejectedError, mapDocumentNoteLocation, registerDocumentNoteCommands,
  type DocumentNoteCommandContext, type DocumentNoteCommandParams, type DocumentNoteCommandResult } from './document-note-commands'
import { prepareDocumentNoteMutation, type DocumentNoteMutationRequest } from '../features/document/notes/document-note-mutations'
import { documentNoteMutationPermissionTargets } from '../features/document/permissions/document-note-permissions'
import { DocumentPermissionError } from '../features/document/permissions/document-permissions'
import { prepareDocumentCustomXmlCandidate } from '../features/document/services/document-custom-xml-candidate'
import { assertDocumentCustomXmlBindings, validateDocumentCustomXmlParts } from '../features/document/services/document-custom-xml-validation'

export interface OfficeStudioOptions {
  locale?: string
  region?: string
  workbook?: WorkbookInput
  document?: Partial<DocumentModel> & Record<string, unknown>
  computeClient?: OfficeComputeClient
  numericIndexMinRows?: number
  chartStatisticsOptions?: ChartStatisticsServiceOptions
  slicerOptions?: SlicerMemberQueryOptions
  currentUserId?: string
  /** Local permission-rule manager. Defaults to currentUserId and is not server authorization. */
  permissionAdminUserId?: string
  /** Inject font measurement for deterministic/headless axis layout. Browser defaults to shared Canvas metrics. */
  sheetTextMeasure?: SheetTextMeasure
}

export type OfficeCommandHandler<P = unknown, R = unknown> = (params: P) => R

export interface WorkbookCommitOptions {
  recalculate?: boolean
  /** Axis relocation already carries authored formatting; recomposition must not freeze inherited values. */
  preserveStyleInheritance?: boolean
  expectedRevision?: number
  expectedHistoryGeneration?: number
  /** Runtime owner/cancellation context; excluded from history snapshots. */
  assertCurrent?: () => void
  /** Recent paste options or a continuous drawing move retain the first undo and publish the newest snapshot. */
  amend?: HistorySavepoint
  /** Receives the committed entry even if a settled listener immediately makes another edit. */
  onHistoryCommitted?: (savepoint: HistorySavepoint) => void
  selection?: { sheetId: string; before: ApiRange | null; after: ApiRange }
  selections?: { sheetId: string; before: ApiSheetSelection; after: ApiSheetSelection }
  /** Prepare detached view data before the model and history become observable. */
  prepareView?: (model: WorkbookModel) => void
}

interface PreparedWorkbookChange {
  candidate: WorkbookModel
  baseRevision: number
  assertCurrent(): void
  publish(): boolean
}

export class OfficeStudio {
  readonly regionalSettings: OfficeRegionalSettings
  readonly pinia: Pinia
  readonly editor: ReturnType<typeof useEditorStore>
  readonly sheetFormatPainter: ReturnType<typeof useSheetFormatPainterStore>
  readonly sheetSelections: ReturnType<typeof useSheetSelectionStore>
  readonly sheetClipboard: ReturnType<typeof useSheetClipboardStore>
  readonly sheetFill: ReturnType<typeof useSheetFillStore>
  readonly documentInput: ReturnType<typeof useDocumentInputStore>
  readonly documentCharacterPreferences: DocumentCharacterPreferencesStore
  readonly events = new EventHub<ApiEventMap>()
  readonly numericIndexService: NumericIndexService
  readonly chartStatisticsService: ChartStatisticsService
  readonly slicerMemberQueryService: SlicerMemberQueryService
  readonly currentUserId: string
  readonly permissionAdminUserId: string
  readonly sheetTextMeasure: SheetTextMeasure
  private readonly workbooks = new Map<string, WorkbookModel>()
  private readonly documents = new Map<string, DocumentModel>()
  private readonly commands = new Map<string, OfficeCommandHandler<never, unknown>>()
  private readonly asynchronousCommands = new Set<string>()
  private readonly workbenchUiOwners = new Map<string, WorkbenchUiOperationHandler>()
  private readonly workbookCreationTasks = new Set<AbortController>()
  private _activeWorkbookId: string
  private _activeDocumentId: string
  private disposed = false
  private readonly workbookRevisions = new Map<string, number>()
  private readonly documentRevisions = shallowReactive(new Map<string, number>())
  private readonly documentFieldCacheAdapters = new Map<string, DocumentFieldCacheAdapter>()
  private readonly borderFormatStates = new Map<string, Map<string, SheetBorderFormatState>>()
  private readonly computeClient: OfficeComputeClient
  private readonly ownsComputeClient: boolean
  private readonly stopModelSync: Array<() => void> = []
  private workbookSelectionPublicationDepth = 0
  private selectionPublication = 0
  private pendingSheetSelection: { unitId: string; sheetId: string } | undefined
  private readonly pendingDocumentInputPublications = new Set<string>()
  private publishedSheetSelection: { unitId: string; sheetId: string; selection: ApiSheetSelection | null } | undefined
  private readonly eventListenerErrors: ApiEventMap['EventListenerError'][] = []

  constructor(options: OfficeStudioOptions = {}) {
    this.regionalSettings = new OfficeRegionalSettings(options.locale, options.region)
    this.pinia = createPinia()
    this.editor = useEditorStore(this.pinia)
    this.sheetFormatPainter = useSheetFormatPainterStore(this.pinia)
    this.sheetSelections = useSheetSelectionStore(this.pinia)
    this.sheetClipboard = useSheetClipboardStore(this.pinia)
    this.sheetFill = useSheetFillStore(this.pinia)
    this.documentInput = useDocumentInputStore(this.pinia)
    this.documentCharacterPreferences = useDocumentCharacterPreferencesStore(this.pinia)
    const workbook = workbookFromInput(options.workbook ?? this.editor.workbook.model)
    const document = documentFromInput(options.document ?? this.editor.document.model)
    if (document.customXmlParts !== undefined) document.customXmlParts = validateDocumentCustomXmlParts(document.customXmlParts)
    assertDocumentCustomXmlBindings(document)
    if (workbook.id === document.id) throw new TypeError('工作簿与文档必须使用不同标识，不能共用内容历史。')
    this._activeWorkbookId = workbook.id!
    this._activeDocumentId = document.id!
    this.currentUserId = options.currentUserId?.trim() || 'local-user'
    this.permissionAdminUserId = options.permissionAdminUserId?.trim() || this.currentUserId
    workbook.activeSheetId = this.activeWorksheetIdForView(workbook)
    this.sheetTextMeasure = options.sheetTextMeasure ?? measureSpreadsheetText
    this.workbooks.set(workbook.id!, workbook)
    this.workbookRevisions.set(workbook.id!, 0)
    this.ownsComputeClient = options.computeClient === undefined
    this.computeClient = options.computeClient ?? new OfficeComputeClient()
    this.numericIndexService = new NumericIndexService(this.computeClient, { minRows: options.numericIndexMinRows })
    this.chartStatisticsService = new ChartStatisticsService(this.computeClient, options.chartStatisticsOptions)
    this.slicerMemberQueryService = new SlicerMemberQueryService({ ...options.slicerOptions,
      computeClient: options.slicerOptions?.computeClient ?? this.computeClient })
    this.numericIndexService.setDocumentRevision(workbook.id!, 0)
    this.chartStatisticsService.setDocumentRevision(workbook.id!, 0)
    this.documents.set(document.id!, document)
    this.documentRevisions.set(document.id!, 0)
    this.editor.workbook.replace(workbook)
    this.editor.document.replace(document)
    this.editor.setActiveKind('workbook')
    this.editor.history.activate(workbook.id!)
    const initialSheet = workbook.sheets.find(sheet => sheet.id === workbook.activeSheetId)
    if (initialSheet && initialSheet.visible !== false && canViewWorksheetProtection(initialSheet, this.currentUserId)) {
      this.sheetSelections.set(workbook.id!, workbook.activeSheetId,
        this.selectionStateForSheet(workbook, workbook.activeSheetId,
          normalizeSheetSelection([{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }], initialSheet)))
    }
    this.registerBuiltinCommands()
    this.stopModelSync.push(
      watch(() => this.editor.history.isApplying, applying => {
        if (applying || this.disposed) return
        const documentIds = [...this.pendingDocumentInputPublications]
        this.pendingDocumentInputPublications.clear()
        if (this.pendingSheetSelection) {
          this.pendingSheetSelection = undefined
          // History callbacks may switch the active unit or sheet. Publish after the stack settles.
          const workbook = this.getWorkbookModel(this._activeWorkbookId)
          if (workbook && this.editor.state.activeKind === 'workbook') this.publishSheetSelection(this._activeWorkbookId, workbook.activeSheetId)
        }
        for (const id of documentIds) if (this.getDocumentModel(id)) this.publishDocumentInputState(id)
      }, { flush: 'sync' }),
      watch(() => this.editor.workbook.model, (next, previous) => {
        const id = next.id
        if (!id || this.disposed) return
        next.activeSheetId = this.activeWorksheetIdForView(next)
        if (id !== this._activeWorkbookId) {
          this.workbooks.set(this._activeWorkbookId, cloneModel(previous))
          const created = !this.workbooks.has(id)
          this._activeWorkbookId = id
          this.workbooks.set(id, next)
          this.advanceWorkbookRevision(id)
          if (created) this.events.emit('WorkbookCreated', { workbookId: id })
        } else {
          this.workbooks.set(id, next)
          if (previous.id === id && !equivalentModelData(next, previous, ['modified', 'activeSheetId'])) this.advanceWorkbookRevision(id)
        }
        this.reconcileWorkbookSelections(id, next)
        if (!this.workbookSelectionPublicationDepth) this.publishSheetSelection(id, next.activeSheetId)
      }, { flush: 'sync' }),
      watch(() => this.editor.document.model, (next, previous) => {
        const id = next.id
        if (!id || this.disposed) return
        if (previous !== next) cancelDocumentImageTasksForSource(previous)
        if (id !== this._activeDocumentId) {
          this.documents.set(this._activeDocumentId, cloneModel(previous))
          const created = !this.documents.has(id)
          this._activeDocumentId = id
          this.documents.set(id, next)
          if (!this.documentRevisions.has(id)) this.documentRevisions.set(id, 0)
          this.documentInput.clearMarks(previous.id ?? '')
          this.documentInput.clearMarks(id)
          if (created) this.events.emit('DocCreated', { unitId: id, documentId: id })
        } else this.documents.set(id, next)
        if (previous.id === id && !equivalentDocumentContent(next, previous)) this.advanceDocumentRevision(id)
      }, { flush: 'sync' }),
      watch(() => this.editor.state.activeKind, (kind, previous) => {
        if (kind !== previous && this.documentInput.clearMarks(this._activeDocumentId)) this.publishDocumentInputState(this._activeDocumentId)
      }, { flush: 'sync' }),
    )
    this.events.emit('LifeCycleChanged', { stage: 'ready' })
  }

  get activeWorkbookId(): string { return this._activeWorkbookId }
  get activeDocumentId(): string { return this._activeDocumentId }
  get isDisposed(): boolean { return this.disposed }
  getEventListenerErrors(): ApiEventMap['EventListenerError'][] { return this.eventListenerErrors.map(failure => ({ ...failure })) }

  /** Feature observations report subscriber failures without turning a completed edit into a failed command. */
  publishObserverEvent<K extends keyof ApiEventMap>(eventName: K, payload: () => ApiEventMap[K], shouldContinue?: () => boolean): void {
    const failures = this.events.emitSettled(eventName, payload, shouldContinue)
    for (const failure of failures) this.reportObserverError(failure.eventName, failure.error)
  }

  /** Feature payload preparation has the same failure isolation as listeners. */
  reportObserverError<K extends keyof ApiEventMap>(eventName: K, error: unknown): void {
    const failure: ApiEventMap['EventListenerError'] = { eventName, error }
    this.eventListenerErrors.push(failure)
    this.eventListenerErrors.push(...this.events.emitSettled('EventListenerError', () => ({ ...failure })))
    if (this.eventListenerErrors.length > 50) this.eventListenerErrors.splice(0, this.eventListenerErrors.length - 50)
  }

  getWorkbookRevision(id: string): number {
    if (!this.getWorkbookModel(id)) throw new Error(`Workbook not found: ${id}`)
    return this.workbookRevisions.get(id) ?? 0
  }

  getDocumentRevision(id: string): number {
    if (!this.getDocumentModel(id)) throw new Error(`Document not found: ${id}`)
    return this.documentRevisions.get(id) ?? 0
  }

  getDocumentZoomRatio(id: string = this._activeDocumentId): number {
    this.assertAlive()
    const model = this.getDocumentModel(id)
    if (!model) throw new Error(`Document not found: ${id}`)
    assertDocumentPermission(model.permissionRules ?? [], { documentId: id, userId: this.currentUserId }, 'view')
    return documentEffectiveZoomRatio(model)
  }

  setDocumentZoomRatio(id: string, zoomRatio: number): boolean {
    return this.executeCommandSync<DocumentZoomOperationParams, boolean>(DOCUMENT_ZOOM_COMMANDS.SetRatio, { unitId: id, zoomRatio })
  }

  private publishDocumentZoomRatio({ unitId, zoomRatio }: DocumentZoomOperationParams): boolean {
    this.assertContentUnitIdentity('document', unitId)
    const model = this.getDocumentModel(unitId)
    if (!model) return false
    assertDocumentPermission(model.permissionRules ?? [], { documentId: unitId, userId: this.currentUserId }, 'view')
    const previousZoomRatio = documentEffectiveZoomRatio(model)
    if (previousZoomRatio === zoomRatio) return true
    // The model remains the same owner; view publication does not revalidate or
    // invalidate the input session, create a savepoint, or mark content dirty.
    model.metadata = { ...model.metadata, documentZoomRatio: zoomRatio }
    if (unitId === this._activeDocumentId) this.editor.document.notifyViewChange()
    this.publishObserverEvent('DocumentZoomRatioChanged', () => ({ documentId: unitId, previousZoomRatio, zoomRatio }),
      () => this.getDocumentModel(unitId) === model && documentEffectiveZoomRatio(model) === zoomRatio)
    return true
  }

  getDocumentInputMarks(unitId: string, selection?: DocumentSelection, story?: DocumentStoryRange, host?: DocumentAuxiliaryBlockHost): RunPatch {
    if (!this.getDocumentModel(unitId)) return {}
    if (host) validateDocumentBlockHost(host)
    if (story && host) throw new RangeError('输入选区不能同时指定旧文字范围和块内容区。')
    const sessionLocation = this.documentInput.get(unitId)?.location
    if (host && !selection && (sessionLocation?.kind !== 'body' || !sameDocumentBlockHost(sessionLocation.host, host))) return {}
    const location = story ? { kind: 'story' as const, range: { ...story, start: Math.min(story.start, story.end), end: Math.max(story.start, story.end) } }
      : selection ? documentInputLocation(selection, undefined, host) : host ? sessionLocation : undefined
    return this.documentInput.getMarks(unitId, this.getDocumentRevision(unitId), location)
  }

  getDocumentInputState(unitId: string): DocumentInputSession | undefined {
    if (!this.getDocumentModel(unitId)) return undefined
    const state = this.documentInput.get(unitId)
    return state ? { ...state, marks: this.getDocumentInputMarks(unitId) } : undefined
  }

  /** Hosts without FDocument still publish automatic fields through the same
   * command, permissions and derived-history transaction as the workbench. */
  getDocumentFieldCacheAdapter(unitId: string = this._activeDocumentId): DocumentFieldCacheAdapter {
    this.assertAlive()
    if (!this.documents.has(unitId)) throw new Error(`Document not found: ${unitId}`)
    const existing = this.documentFieldCacheAdapters.get(unitId)
    if (existing) return existing
    const getState = (): DocumentFieldCacheState | undefined => {
      // During DocDisposed the editor can still hold the outgoing model.
      if (this.disposed || !this.documents.has(unitId)) return undefined
      const model = this.getDocumentModel(unitId)
      if (!model) return undefined
      return { unitId, model, modelRevision: this.getDocumentRevision(unitId),
        historyGeneration: this.editor.history.captureSavepoint(unitId).generation,
        inputGeneration: this.documentInput.get(unitId)?.generation }
    }
    const adapter: DocumentFieldCacheAdapter = {
      identity: this.regionalSettings,
      getState,
      refresh: request => {
        const source = getState()
        if (!source || request.unitId !== unitId || request.expectedRevision !== source.modelRevision
          || request.expectedHistoryGeneration !== source.historyGeneration || request.inputGeneration !== source.inputGeneration) {
          return { accepted: false, changed: false }
        }
        const activeId = this._activeDocumentId, activeKind = this.editor.state.activeKind
        const result = this.executeCommandSync<DocumentEditorCommandRequest, DocumentEditorCommandResult>('doc.command.refresh-header-footer-fields', {
          unitId, command: { id: 'doc.command.refresh-header-footer-fields', locale: request.locale, timestamp: request.timestamp, onOpen: true },
          selection: request.selection, expectedRevision: request.expectedRevision,
          expectedHistoryGeneration: request.expectedHistoryGeneration, inputGeneration: request.inputGeneration,
        })
        const committed = result.fieldCache
        if (!committed) throw new Error('字段缓存命令没有返回已提交的缓存状态。')
        const latest = getState()
        return { accepted: true, changed: result.changed, model: committed.model, selection: result.selection,
          modelRevision: committed.modelRevision, historyGeneration: committed.historyGeneration,
          selectionCurrent: result.selectionCurrent === true && latest !== undefined
            && this._activeDocumentId === activeId && this.editor.state.activeKind === activeKind
            && latest.modelRevision === committed.modelRevision && latest.historyGeneration === committed.historyGeneration
            && latest.inputGeneration === committed.inputGeneration }
      },
    }
    this.documentFieldCacheAdapters.set(unitId, adapter)
    return adapter
  }

  setDocumentInputSelection(unitId: string, selection: DocumentSelection, story?: DocumentStoryRange, host?: DocumentAuxiliaryBlockHost): boolean {
    return this.executeCommandSync<DocumentInputSelectionRequest, boolean>(SET_DOCUMENT_INPUT_SELECTION_OPERATION, { unitId, selection, story, host })
  }

  clearDocumentInputMarks(unitId: string): boolean {
    return this.executeCommandSync<{ unitId: string }, boolean>(CLEAR_DOCUMENT_INPUT_MARKS_OPERATION, { unitId })
  }

  /** Consume a checked input commit once; navigation or a subsequent edit invalidates its caret and overrides. */
  advanceDocumentInputMarks(unitId: string, selection: DocumentSelection, story?: DocumentStoryRange, host?: DocumentAuxiliaryBlockHost): boolean {
    return this.executeCommandSync<DocumentInputSelectionRequest, boolean>(ADVANCE_DOCUMENT_INPUT_MARKS_OPERATION, { unitId, selection, story, host })
  }

  getWorkbookModel(id: string): WorkbookModel | undefined {
    if (this.disposed) return undefined
    return id === this._activeWorkbookId ? this.editor.workbook.model : this.workbooks.get(id)
  }

  assertCanEditRange(unitId: string, sheetId: string, range: CellRange): void {
    const workbook = this.getWorkbookModel(unitId)
    if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
    const sheet = requireSheet(workbook, sheetId)
    const scopes = [range, ...sheet.mergedRanges.filter(merge => merge.startRow <= range.endRow && merge.endRow >= range.startRow
      && merge.startCol <= range.endCol && merge.endCol >= range.startCol)]
    if (!canEditWorksheetProtection(sheet, this.currentUserId) || !canEditProtectedRanges(sheet, scopes, this.currentUserId)) {
      throw new Error('该区域受保护，当前用户没有编辑权限。')
    }
  }

  assertCanViewWorksheet(unitId: string, sheetId: string): void {
    const workbook = this.getWorkbookModel(unitId)
    if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
    if (!canViewWorksheetProtection(requireSheet(workbook, sheetId), this.currentUserId)) {
      throw new Error('该工作表受保护，当前用户没有查看权限。')
    }
  }

  assertCanViewRange(unitId: string, sheetId: string, range: CellRange): void {
    const workbook = this.getWorkbookModel(unitId)
    if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
    const sheet = requireSheet(workbook, sheetId)
    const scopes = [range, ...sheet.mergedRanges.filter(merge => merge.startRow <= range.endRow && merge.endRow >= range.startRow
      && merge.startCol <= range.endCol && merge.endCol >= range.startCol)]
    if (!canViewWorksheetProtection(sheet, this.currentUserId) || !canViewProtectedRanges(sheet, scopes, this.currentUserId)) {
      throw new Error('该区域受保护，当前用户没有查看权限。')
    }
  }

  assertCanEditWorksheet(unitId: string, sheetId: string): void {
    const workbook = this.getWorkbookModel(unitId)
    if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
    if (!canEditWorksheetProtection(requireSheet(workbook, sheetId), this.currentUserId)) {
      throw new Error('该工作表受保护，当前用户没有编辑权限。')
    }
  }

  getDocumentModel(id: string): DocumentModel | undefined {
    if (this.disposed) return undefined
    return id === this._activeDocumentId ? this.editor.document.model : this.documents.get(id)
  }

  createWorkbook(input: WorkbookInput = {}, makeCurrent = true): WorkbookModel {
    this.assertAlive()
    const model = workbookFromInput(input)
    if (this.workbooks.has(model.id!)) throw new Error(`Workbook id already exists: ${model.id}`)
    this.assertContentUnitIdentity('workbook', model.id!)
    model.activeSheetId = this.activeWorksheetIdForView(model)
    this.workbooks.set(model.id!, model)
    this.workbookRevisions.set(model.id!, 0)
    this.numericIndexService.setDocumentRevision(model.id!, 0)
    this.chartStatisticsService.setDocumentRevision(model.id!, 0)
    if (makeCurrent) this.activateWorkbook(model.id!)
    this.events.emit('WorkbookCreated', { workbookId: model.id! })
    return model
  }

  async createWorkbookAsync(input: WorkbookInput = {}, makeCurrent = true, options: SheetImageInputPreparationOptions = {}): Promise<WorkbookModel> {
    this.assertAlive()
    const controller = new AbortController(), cancel = () => controller.abort()
    this.workbookCreationTasks.add(controller)
    options.signal?.addEventListener('abort', cancel, { once: true })
    if (options.signal?.aborted) controller.abort()
    try {
      const prepared = await prepareWorkbookSheetImageInput(input, { ...options, signal: controller.signal,
        current: () => !this.disposed && (options.current?.() ?? true) })
      this.assertAlive()
      if (controller.signal.aborted) throw new DOMException('工作簿图片读取已取消。', 'AbortError')
      return this.createWorkbook(prepared, makeCurrent)
    } finally { options.signal?.removeEventListener('abort', cancel); this.workbookCreationTasks.delete(controller) }
  }

  createDocument(input: Partial<DocumentModel> & Record<string, unknown> = {}, makeCurrent = true): DocumentModel {
    this.assertAlive()
    const model = documentFromInput(input)
    if (this.documents.has(model.id!)) throw new Error(`Document id already exists: ${model.id}`)
    this.assertContentUnitIdentity('document', model.id!)
    if (model.customXmlParts !== undefined) model.customXmlParts = validateDocumentCustomXmlParts(model.customXmlParts)
    assertDocumentCustomXmlBindings(model)
    this.documents.set(model.id!, model)
    this.documentRevisions.set(model.id!, 0)
    if (makeCurrent) this.activateDocument(model.id!)
    this.events.emit('DocCreated', { unitId: model.id!, documentId: model.id! })
    return model
  }

  activateWorkbook(id: string): void {
    this.assertAlive()
    this.assertContentUnitIdentity('workbook', id)
    const current = this.getWorkbookModel(id)
    if (!current) throw new Error(`Workbook not found: ${id}`)
    if (this.editor.state.activeKind === 'document' && this.documentInput.clearMarks(this._activeDocumentId)) this.publishDocumentInputState(this._activeDocumentId)
    const activeSheetId = this.activeWorksheetIdForView(current)
    const model = current.activeSheetId === activeSheetId ? current : { ...current, activeSheetId }
    if (this._activeWorkbookId !== id && this.workbooks.has(this._activeWorkbookId)) this.workbooks.set(this._activeWorkbookId, this.editor.workbook.model)
    this._activeWorkbookId = id
    this.workbooks.set(id, model)
    const modified = this.editor.history.isDirty(id)
    if (this.editor.workbook.model !== model || model.modified !== modified) {
      this.replaceWorkbookModel(id, { ...model, modified }, false)
    }
    this.editor.setActiveKind('workbook')
    this.editor.history.activate(id)
    this.publishSheetSelection(id, model.activeSheetId)
  }

  activateDocument(id: string): void {
    this.assertAlive()
    this.assertContentUnitIdentity('document', id)
    const model = this.getDocumentModel(id)
    if (!model) throw new Error(`Document not found: ${id}`)
    if (id !== this._activeDocumentId || this.editor.state.activeKind !== 'document') {
      if (this.documentInput.clearMarks(this._activeDocumentId)) this.publishDocumentInputState(this._activeDocumentId)
      if (id !== this._activeDocumentId && this.documentInput.clearMarks(id)) this.publishDocumentInputState(id)
    }
    if (this._activeDocumentId !== id && this.documents.has(this._activeDocumentId)) this.documents.set(this._activeDocumentId, this.editor.document.model)
    this._activeDocumentId = id
    this.documents.set(id, model)
    this.editor.document.replace({ ...model, modified: this.editor.history.isDirty(id) })
    this.editor.setActiveKind('document')
    this.editor.history.activate(id)
  }

  setActiveKind(kind: OfficeKind): void {
    this.assertAlive()
    this.assertContentUnitIdentity(kind, kind === 'workbook' ? this._activeWorkbookId : this._activeDocumentId)
    this.editor.setActiveKind(kind)
    this.editor.history.activate(kind === 'workbook' ? this._activeWorkbookId : this._activeDocumentId)
    if (kind === 'workbook') this.publishSheetSelection(this._activeWorkbookId, this.editor.workbook.model.activeSheetId)
  }

  getWorkbook(id: string): WorkbookModel | undefined {
    return this.getWorkbookModel(id)
  }

  getDocument(id: string): DocumentModel | undefined {
    return this.getDocumentModel(id)
  }

  getSelection(): ApiRange | null {
    const selection = this.getSelectionState()
    return selection?.ranges[selection.activeRangeIndex] ?? null
  }

  getSelectionState(unitId = this._activeWorkbookId, sheetId?: string): ApiSheetSelection | null {
    const workbook = this.getWorkbookModel(unitId)
    if (!workbook) return null
    const sheet = workbook.sheets.find(item => item.id === (sheetId ?? workbook.activeSheetId))
    if (!sheet || sheet.visible === false || !canViewWorksheetProtection(sheet, this.currentUserId)) return null
    const saved = this.sheetSelections.get(unitId, sheet.id)
    const state = this.selectionStateForSheet(workbook, sheet.id,
      saved ?? normalizeSheetSelection([{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }], sheet))
    return this.apiSheetSelection(unitId, sheet.id, state)
  }

  getSelections(unitId = this._activeWorkbookId, sheetId?: string): ApiRange[] {
    return this.getSelectionState(unitId, sheetId)?.ranges ?? []
  }

  private commandSelectionRanges(unitId: string, sheetId: string): CellRange[] {
    const active = this.getSelection()
    return active?.unitId === unitId && active.sheetId === sheetId ? this.getSelections().map(range => ({
      startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn,
    })) : []
  }

  getSheetBorderFormatState(unitId: string, sheetId: string): SheetBorderFormatState {
    this.assertCanViewWorksheet(unitId, sheetId)
    return { ...(this.borderFormatStates.get(unitId)?.get(sheetId) ?? createSheetBorderFormatState()) }
  }

  setSelection(unitId: string, sheetId: string, range: ApiRange): void {
    this.executeCommandSync(SELECT_SHEET_RANGE_COMMAND, { unitId, sheetId, range })
  }

  setSelections(unitId: string, sheetId: string, ranges: readonly ApiRange[], options: SheetSelectionOptions = {}): void {
    this.executeCommandSync(SET_SHEET_SELECTIONS_OPERATION, { unitId, sheetId, ranges, options })
  }

  private applySelections(unitId: string, sheetId: string, ranges: readonly ApiRange[], options: SheetSelectionOptions): boolean {
    this.assertAlive()
    const workbook = this.getWorkbookModel(unitId)
    if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
    if (!workbook.sheets.some((sheet) => sheet.id === sheetId)) throw new Error(`Worksheet not found: ${sheetId}`)
    this.assertCanViewWorksheet(unitId, sheetId)
    if (requireSheet(workbook, sheetId).visible === false) throw new Error('工作表已隐藏，请先取消隐藏再选择。')
    if (!Array.isArray(ranges)) throw new TypeError('选区必须是范围数组。')
    for (const range of ranges) if (!range || range.unitId !== undefined && range.unitId !== unitId || range.sheetId !== undefined && range.sheetId !== sheetId) {
      throw new Error('选区必须属于指定的工作簿和工作表。')
    }
    const requested = normalizeSheetSelection(ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow,
      startCol: range.startColumn, endCol: range.endColumn, ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) })), requireSheet(workbook, sheetId), options)
    const next = this.selectionStateForSheet(workbook, sheetId, requested)
    const switched = unitId !== this._activeWorkbookId || workbook.activeSheetId !== sheetId || this.editor.state.activeKind !== 'workbook'
    // Complete the model operation's validation before publishing any selection state.
    if (workbook.activeSheetId !== sheetId) {
      this.workbookSelectionPublicationDepth++
      try { this.applyWorkbookOperation(unitId, draft => { draft.activeSheetId = sheetId }) }
      finally { this.workbookSelectionPublicationDepth-- }
    }
    const changed = this.sheetSelections.set(unitId, sheetId, next)
    this.sheetSelections.setTabStart(unitId, sheetId)
    if (switched) this.activateWorkbook(unitId)
    else if (changed) this.publishSheetSelection(unitId, sheetId)
    return changed || switched
  }

  private selectionStateForSheet(workbook: WorkbookModel, sheetId: string, state: SheetSelectionState): SheetSelectionState {
    const sheet = requireSheet(workbook, sheetId)
    const filtered = (workbook.tables ?? []).filter(table => table.sheetId === sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])
    return resolveVisibleSheetSelection(state, { ...sheet, rowCount: sheet.rowCount ?? 1_048_576, columnCount: sheet.columnCount ?? 16_384,
      filteredRows: sheet.autoFilter?.filteredRows }, filtered)
  }

  private reconcileWorkbookSelections(unitId: string, workbook: WorkbookModel): void {
    this.sheetSelections.retainSheets(unitId, workbook.sheets.map(sheet => sheet.id))
    for (const sheet of workbook.sheets) {
      const saved = this.sheetSelections.get(unitId, sheet.id)
      if (saved) this.sheetSelections.set(unitId, sheet.id, this.selectionStateForSheet(workbook, sheet.id, saved))
    }
  }

  private apiSheetSelection(unitId: string, sheetId: string, state: SheetSelectionState): ApiSheetSelection {
    return { ranges: state.ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol,
      unitId, sheetId, ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) })), activeRangeIndex: state.activeRangeIndex,
      ...(state.primary ? { primary: { ...state.primary } } : {}) }
  }

  private activeWorksheetIdForView(workbook: WorkbookModel): string {
    const available = (sheet: WorkbookModel['sheets'][number]) => sheet.visible !== false && canViewWorksheetProtection(sheet, this.currentUserId)
    return workbook.sheets.find(sheet => sheet.id === workbook.activeSheetId && available(sheet))?.id
      ?? workbook.sheets.find(available)?.id ?? workbook.activeSheetId
  }

  private publishSheetSelection(unitId: string, sheetId: string): void {
    if (unitId !== this._activeWorkbookId || this.editor.state.activeKind !== 'workbook'
      || this.getWorkbookModel(unitId)?.activeSheetId !== sheetId) return
    if (this.editor.history.isApplying) { this.pendingSheetSelection = { unitId, sheetId }; return }
    const selection = this.getSelectionState(unitId, sheetId)
    const notification = { unitId, sheetId, selection }
    if (JSON.stringify(notification) === JSON.stringify(this.publishedSheetSelection)) return
    const publication = ++this.selectionPublication
    this.publishedSheetSelection = cloneModel(notification)
    const range = selection ? selection.ranges[selection.activeRangeIndex] : undefined
    this.editor.setSelection({
      startRow: range?.startRow ?? 0,
      startCol: range?.startColumn ?? 0,
      endRow: range?.endRow ?? 0,
      endCol: range?.endColumn ?? 0,
    })
    if (!selection) return
    const legacyRange = range ? { ...range } : undefined
    const current = () => publication === this.selectionPublication && unitId === this._activeWorkbookId
      && this.editor.state.activeKind === 'workbook' && this.getWorkbookModel(unitId)?.activeSheetId === sheetId
    this.publishObserverEvent('SelectionsChanged', () => ({ workbookId: unitId, sheetId, selection: cloneModel(selection) }), current)
    if (legacyRange && current()) this.publishObserverEvent('SelectionChanged',
      () => ({ workbookId: unitId, sheetId, range: { ...legacyRange } }), current)
  }

  mutateWorkbook(
    id: string,
    commandId: string,
    label: string,
    update: (draft: WorkbookModel) => void,
    params: unknown = {},
  ): boolean {
    const event: ApiEventMap['BeforeCommandExecute'] = {
      id: commandId,
      params,
      unitId: id,
      kind: 'workbook',
    }
    this.events.emit('BeforeCommandExecute', event)
    if (event.cancel) throw new Error(`Command was canceled: ${commandId}`)
    const rangeEdit = rangeEditFromParams(params)
    if (rangeEdit) this.assertCanEditRange(id, rangeEdit.sheetId, rangeEdit.range)
    const changed = commandId.includes('.operation.')
      ? this.applyWorkbookOperation(id, update)
      : this.commitWorkbook(id, label, update)
    if (changed) this.publishObserverEvent('CommandExecuted', () => ({ id: commandId, params: cloneModel(params), unitId: id, kind: 'workbook' }))
    return changed
  }

  mutateDocument(
    id: string,
    commandId: string,
    label: string,
    update: (draft: DocumentModel) => void,
    params: unknown = {},
  ): boolean {
    const event: ApiEventMap['BeforeCommandExecute'] = {
      id: commandId,
      params,
      unitId: id,
      kind: 'document',
    }
    this.events.emit('BeforeCommandExecute', event)
    if (event.cancel) throw new Error(`Command was canceled: ${commandId}`)
    const changed = this.commitDocument(id, label, update)
    if (changed) this.publishObserverEvent('CommandExecuted', () => ({ id: commandId, params: cloneModel(params), unitId: id, kind: 'document' }))
    return changed
  }

  commitWorkbook(id: string, label: string, update: (draft: WorkbookModel) => void, options: boolean | WorkbookCommitOptions = false): boolean {
    return this.prepareWorkbookChange(id, label, update, options)?.publish() ?? false
  }

  async commitChartChange(id: string, label: string, chartId: string, update: (draft: WorkbookModel) => void,
    authoring: ChartAuthoringPlan, assertHostCurrent?: () => void, options: boolean | WorkbookCommitOptions = false): Promise<boolean> {
    this.assertAlive()
    this.assertContentUnitIdentity('workbook', id)
    const userId = this.currentUserId, savepoint = this.editor.history.captureSavepoint(id)
    // Prepare source-free configuration and the normal formula/filter candidate
    // first. No chart source arrays are allocated by this synchronous mutator.
    const plan = this.prepareWorkbookChange(id, label, update, options)
    if (!plan) return false
    const seed = requireChart(plan.candidate, chartId)
    const assertCurrent = (): void => {
      this.assertAlive()
      assertHostCurrent?.()
      const current = this.getWorkbookModel(id), latestSavepoint = this.editor.history.captureSavepoint(id)
      if (!current || this.currentUserId !== userId || this.getWorkbookRevision(id) !== plan.baseRevision
        || latestSavepoint.revision !== savepoint.revision || latestSavepoint.generation !== savepoint.generation) {
        throw new ComputeError('COMPUTE_STALE_RESULT', '图表准备期间工作簿、用户或内容历史已更新')
      }
      this.assertCanEditWorksheet(id, seed.sheetId)
      const previous = current.charts?.find(chart => chart.id === chartId)
      if (previous) assertChartReadPermission(current, previous, userId)
      // Existing sources use their live permissions. Newly imported sheets and
      // reports must be checked in the detached, fully bound candidate.
      const permissionWorkbook: WorkbookModel = { ...plan.candidate,
        sheets: plan.candidate.sheets.map(sheet => current.sheets.find(live => live.id === sheet.id) ?? sheet) }
      assertChartReadPermission(permissionWorkbook, seed, userId)
      assertChartReadPermission(plan.candidate, seed, userId)
    }
    assertCurrent()
    const chart = await this.chartStatisticsService.requestAuthoredCandidate(plan.candidate, seed,
      { documentId: id, modelRevision: plan.baseRevision, userId, priority: 'interactive', assertCurrent }, authoring).promise
    assertCurrent()
    // Check complete content equivalence once at publication, rather than
    // serializing the whole workbook at each cooperative preparation boundary.
    plan.assertCurrent()
    if (chart.id !== seed.id || chart.sheetId !== seed.sheetId || chart.revision !== seed.revision) {
      throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '图表准备结果与候选身份不匹配')
    }
    Object.assign(seed, cloneModel(chart))
    publishChartImageResourceCandidate(plan.candidate, seed)
    assertHostCurrent?.()
    return plan.publish()
  }

  async commitWorkbookWithValidation(id: string, label: string, update: (draft: WorkbookModel) => void,
    validate: (candidate: WorkbookModel, context: { baseRevision: number; assertCurrent(): void }) => Promise<void>,
    options: boolean | WorkbookCommitOptions = false): Promise<boolean> {
    const plan = this.prepareWorkbookChange(id, label, update, options)
    if (!plan) return false
    const candidateSignature = JSON.stringify(plan.candidate)
    await validate(plan.candidate, { baseRevision: plan.baseRevision, assertCurrent: () => plan.assertCurrent() })
    if (JSON.stringify(plan.candidate) !== candidateSignature) throw new Error('异步校验修改了工作簿候选，请重新准备编辑。')
    return plan.publish()
  }

  private prepareWorkbookChange(id: string, label: string, update: (draft: WorkbookModel) => void,
    options: boolean | WorkbookCommitOptions): PreparedWorkbookChange | undefined {
    this.assertAlive()
    this.assertContentUnitIdentity('workbook', id)
    const current = this.getWorkbookModel(id)
    if (!current) throw new Error(`Workbook not found: ${id}`)
    const baseRevision = this.getWorkbookRevision(id), baseSavepoint = this.editor.history.captureSavepoint(id)
    if (typeof options !== 'boolean') {
      options.assertCurrent?.()
      if (options.expectedRevision !== undefined && (!Number.isSafeInteger(options.expectedRevision) || options.expectedRevision < 0)) throw new TypeError('工作簿版本无效。')
      if (options.expectedHistoryGeneration !== undefined && (!Number.isSafeInteger(options.expectedHistoryGeneration) || options.expectedHistoryGeneration <= 0)) throw new TypeError('工作簿历史版本无效。')
      if (options.expectedRevision !== undefined && options.expectedRevision !== baseRevision
        || options.expectedHistoryGeneration !== undefined && options.expectedHistoryGeneration !== baseSavepoint.generation) throw new Error('工作簿内容或历史已更改，请重新执行。')
      if (options.amend && !this.editor.history.canAmend(id, options.amend)) return undefined
    }
    const before = cloneModel(current)
    const after = cloneModel(current)
    update(after)
    const removedReportResource = (['pivotTables', 'charts', 'tables'] as const).some(key => {
      const surviving = new Set((after[key] ?? []).map(resource => resource.id))
      return (before[key] ?? []).some(resource => !surviving.has(resource.id))
    })
    if (removedReportResource) reconcileSlicerConnectionResources(before, after)
    normalizeSlicerCaches(after)
    synchronizeSlicerSelections(after, before)
    synchronizeLinkedPivotChartReports(before, after)
    for (const source of [before, after]) for (const control of source.slicers ?? []) {
      const peer = (source === before ? after : before).slicers?.find(view => view.id === control.id)
      if (!equivalentModelData(control, peer)) assertSlicerViewPermission(source, control, this.currentUserId, true)
    }
    refreshWorkbookCellStyles(after, before, typeof options === 'boolean' || options.preserveStyleInheritance !== true, this.regionalSettings.region.value)
    assertWorkbookConditionalFormatBudget(after.sheets)
    assertWorksheetProtectionChanges(before, after, this.currentUserId)
    this.assertChartHostChanges(before, after)
    const explicitRecalculate = typeof options === 'boolean' ? options : options.recalculate === true
    if (!explicitRecalculate && JSON.stringify(before) === JSON.stringify(after)) return undefined
    const automatic = before.calculationMode !== 'manual' || after.calculationMode !== 'manual'
    const shouldRecalculate = typeof options === 'boolean' ? options || automatic : options.recalculate ?? automatic
    const recalculated = shouldRecalculate
      ? recalculateWorkbookFormulas(after, before, { numberFormatLocale: this.regionalSettings.region.value }).workbook
      : after
    for (const sheet of recalculated.sheets) if (sheet.autoFilter) refreshAutoFilter(sheet, recalculated)
    refreshWorkbookAutoRowHeights(recalculated, before, { measure: this.sheetTextMeasure, region: this.regionalSettings.region.value })
    reconcileWorkbookDrawingStates(before, recalculated)
    recalculated.id = id
    if (JSON.stringify(before) === JSON.stringify(recalculated)) return undefined
    recalculated.modified = true
    const single = typeof options === 'boolean' ? undefined : options.selection
    const currentSelection = this.getSelectionState(id, current.activeSheetId)
    const selectedSheetId = current.activeSheetId
    const defaultSelections = currentSelection && recalculated.sheets.some(sheet => sheet.id === selectedSheetId)
      ? { sheetId: selectedSheetId, before: currentSelection, after: this.apiSheetSelection(id, selectedSheetId,
        this.selectionStateForSheet(recalculated, selectedSheetId, {
          ranges: currentSelection.ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn,
            endCol: range.endColumn, ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) })),
          activeRangeIndex: currentSelection.activeRangeIndex, primary: currentSelection.primary,
        })) }
      : undefined
    const historySelection = typeof options === 'boolean' ? undefined : options.selections ?? (single ? { sheetId: single.sheetId,
      before: { ranges: single.before ? [{ ...single.before }] : [], activeRangeIndex: single.before ? 0 : -1 },
      after: { ranges: [{ ...single.after }], activeRangeIndex: 0 } } : undefined)
    const selectedHistory = historySelection ?? defaultSelections
    const selectionSnapshot = selectedHistory ? cloneModel(selectedHistory) : undefined
    if (selectionSnapshot) {
      for (const [model, state] of [[before, selectionSnapshot.before], [recalculated, selectionSnapshot.after]] as const) {
        for (const range of state.ranges) if (range.unitId !== undefined && range.unitId !== id || range.sheetId !== undefined && range.sheetId !== selectionSnapshot.sheetId) {
          throw new Error('历史选区必须属于目标工作簿和工作表。')
        }
        normalizeSheetSelection(state.ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn,
          ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) })), requireSheet(model, selectionSnapshot.sheetId), state)
      }
    }
    const restoreSelection = (state: ApiSheetSelection): void => {
      const model = this.getWorkbookModel(id)
      if (selectionSnapshot && model?.sheets.some(sheet => sheet.id === selectionSnapshot.sheetId)) {
        const normalized = normalizeSheetSelection(state.ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow,
          startCol: range.startColumn, endCol: range.endColumn, ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) })),
          requireSheet(model, selectionSnapshot.sheetId), state)
        this.sheetSelections.set(id, selectionSnapshot.sheetId, this.selectionStateForSheet(model, selectionSnapshot.sheetId, normalized))
      }
      if (id === this._activeWorkbookId && model) this.publishSheetSelection(id, model.activeSheetId)
    }
    if (typeof options !== 'boolean' && options.prepareView) options.prepareView(cloneModel(recalculated))
    const assertCurrent = (): void => {
      this.assertAlive()
      if (typeof options !== 'boolean') options.assertCurrent?.()
      const latest = this.getWorkbookModel(id)
      const latestSavepoint = this.editor.history.captureSavepoint(id)
      if (!latest || !equivalentModelData(before, latest, ['modified', 'activeSheetId']) || this.getWorkbookRevision(id) !== baseRevision
        || latestSavepoint.revision !== baseSavepoint.revision || latestSavepoint.generation !== baseSavepoint.generation) {
        throw new Error('准备编辑期间工作簿内容已更改，请重新执行。')
      }
    }
    return { candidate: recalculated, baseRevision, assertCurrent, publish: () => {
      assertCurrent()
      const latest = this.getWorkbookModel(id)!
      if (recalculated.activeSheetId === before.activeSheetId) recalculated.activeSheetId = latest.activeSheetId
      let initial = true
      const redo = (): void => {
        if (initial) {
          assertCurrent()
          const latest = this.getWorkbookModel(id)
          if (!latest) throw new Error(`Workbook not found: ${id}`)
          this.assertChartHostChanges(latest, recalculated)
          this.replaceWorkbookModel(id, cloneModel(recalculated), !selectionSnapshot); initial = false
        }
        else this.replaceWorkbookHistorySnapshot(id, recalculated, !selectionSnapshot)
        if (selectionSnapshot) restoreSelection(selectionSnapshot.after)
      }
      const undo = (): void => {
        this.replaceWorkbookHistorySnapshot(id, before, !selectionSnapshot)
        if (selectionSnapshot) restoreSelection(selectionSnapshot.before)
      }
      const amendment = typeof options === 'boolean' ? undefined : options.amend
      const committed = amendment ? this.editor.history.amendWithReceipt(id, amendment, label, redo)
        : this.editor.history.execute(id, label, redo, undo)
      if (!committed) return false
      publishWorkbookSparklineChanges(this, before, recalculated)
      if (typeof options !== 'boolean' && options.onHistoryCommitted) {
        try { options.onHistoryCommitted(committed) }
        catch (error) {
          const failure: ApiEventMap['EventListenerError'] = { eventName: 'CommandExecuted', error }
          this.eventListenerErrors.push(failure)
          this.publishObserverEvent('EventListenerError', () => ({ ...failure }))
          if (this.eventListenerErrors.length > 50) this.eventListenerErrors.splice(0, this.eventListenerErrors.length - 50)
        }
      }
      return true
    } }
  }

  commitDocument(id: string, label: string, update: (draft: DocumentModel) => void): boolean {
    return this.commitDocumentTransaction(id, label, update, false)
  }

  commitDocumentPermissionRules(id: string, rules: DocumentModel['permissionRules']): boolean {
    if (this.currentUserId !== this.permissionAdminUserId) throw new Error('Only the document permission administrator can update permission rules')
    return this.commitDocumentTransaction(id, 'Set document permission rules', (model) => {
      const next = rules ?? []
      validateDocumentPermissionRules(next)
      model.permissionRules = next.map(rule => ({ ...rule, userIds: [...rule.userIds] }))
    }, true)
  }

  private commitDocumentTransaction(id: string, label: string, update: (draft: DocumentModel) => void, permissionAdminCommand: boolean,
    amendment?: HistorySavepoint, restoreHistorySelection?: (direction: 'undo' | 'redo') => void): boolean {
    this.assertAlive()
    this.assertContentUnitIdentity('document', id)
    const current = this.getDocumentModel(id)
    if (!current) throw new Error(`Document not found: ${id}`)
    const revision = this.getDocumentRevision(id)
    const userId = this.currentUserId, savepoint = this.editor.history.captureSavepoint(id), historyUnit = this.editor.history.units[id]
    if (amendment && !this.editor.history.canAmend(id, amendment)) throw new Error('目录插入历史已不在栈顶，不能合并重排。')
    const before = cloneModel(current)
    const after = cloneModel(current)
    update(after)
    prepareDocumentCustomXmlCandidate(before, after, { documentId: id, userId })
    if (!permissionAdminCommand) {
      if (JSON.stringify(before.permissionRules ?? []) !== JSON.stringify(after.permissionRules ?? [])) throw new Error('请通过文档权限命令修改权限规则。')
      assertDocumentPermission(current.permissionRules ?? [], { documentId: id, userId }, 'edit', documentNoteMutationPermissionTargets(before, after))
    }
    after.id = id
    if (JSON.stringify(before) === JSON.stringify(after)) return false
    after.modified = true
    const latest = this.getDocumentModel(id)
    const latestSavepoint = this.editor.history.captureSavepoint(id)
    if (!latest || this.getDocumentRevision(id) !== revision || this.currentUserId !== userId || this.editor.history.units[id] !== historyUnit
      || latestSavepoint.generation !== savepoint.generation || latestSavepoint.revision !== savepoint.revision
      || !equivalentModelData(before, latest, ['modified'])) throw new Error('准备编辑期间文档内容、用户或历史已更改，请重新执行。')
    const publish = (model: DocumentModel) => this.replaceDocumentHistorySnapshot(id, model, permissionAdminCommand)
    if (amendment) {
      const entry = this.editor.history.units[id]!.undo.at(-1)!
      const callbacks = prepareDocumentSnapshotHistoryAmend(entry, after, publish)
      if (!this.editor.history.amendWithReceipt(id, amendment, entry.label, callbacks.redo)) throw new Error('目录插入历史已变化，不能合并重排。')
    } else {
      const callbacks = createDocumentSnapshotHistory(before, after, publish, restoreHistorySelection)
      this.editor.history.execute(id, label, callbacks.redo, callbacks.undo)
    }
    return true
  }

  private publishDocumentFieldCaches(id: string, request: DocumentStoryFieldRefreshRequest, plan?: DocumentStoryFieldRefreshPlan): { changed: boolean; state: DocumentFieldCacheState } {
    const current = this.getDocumentModel(id)
    if (!current) throw new Error(`Document not found: ${id}`)
    const revision = this.getDocumentRevision(id), savepoint = this.editor.history.captureSavepoint(id)
    const stack = this.editor.history.units[id]!
    const entries = [...stack.undo, ...stack.redo]
    const context = { documentId: id, userId: this.currentUserId, permissionAdminUserId: this.permissionAdminUserId }
    const refresh = (snapshot: DocumentModel): DocumentModel => {
      const refreshed = prepareDocumentStoryFieldRefresh(snapshot, { locale: request.locale, timestamp: request.timestamp, onOpen: true }, context)
      if (refreshed) applyDocumentFieldRefreshPlan(snapshot, refreshed)
      return snapshot
    }
    const rebase = prepareDocumentHistoryCacheRebase(entries, refresh)
    const next = cloneModel(current)
    if (plan) applyDocumentFieldRefreshPlan(next, plan)
    next.modified = current.modified
    const changed = !equivalentModelData(current, next, ['modified'])
    const latestSavepoint = this.editor.history.captureSavepoint(id)
    if (this.getDocumentModel(id) !== current || this.getDocumentRevision(id) !== revision
      || latestSavepoint.revision !== savepoint.revision || latestSavepoint.generation !== savepoint.generation
      || this.editor.history.units[id] !== stack || stack.undo.length + stack.redo.length !== entries.length
      || [...stack.undo, ...stack.redo].some((entry, index) => entry !== entries[index])) {
      throw new Error('准备字段缓存期间文档或历史已更改，请重新打开。')
    }
    const state: DocumentFieldCacheState = { unitId: id, model: changed ? next : current,
      modelRevision: revision + (changed ? 1 : 0), historyGeneration: savepoint.generation }
    this.editor.history.publishDerived(id, () => {
      rebase()
      if (changed) this.replaceDocumentModel(id, next)
    })
    return { changed, state }
  }

  executeCommandSync<P, R>(id: string, params: P): R {
    this.assertAlive()
    const handler = this.commands.get(id)
    if (!handler) throw new Error(`Unknown command: ${id}`)
    if (this.asynchronousCommands.has(id)) throw new Error(`命令 ${id} 必须通过 executeCommand 异步执行。`)
    if ([DOCUMENT_TOC_INSERT_COMMAND, DOCUMENT_TOC_UPDATE_COMMAND, DOCUMENT_TOC_DELETE_COMMAND].includes(id)
      && (!params || typeof params !== 'object' || !('command' in params))) {
      throw new Error('公开目录命令通过 executeCommand 异步执行；插入和更新会等待完整排版。')
    }
    // Fixed Univer facade exposes history through its own BeforeUndo/Redo and
    // Undo/Redo events, excluding these commands from generic command events.
    if (contentHistoryCommandAction(id)) return (handler as OfficeCommandHandler<P, R>)(params)
    const commandTarget = this.captureDocumentTransientCommandTarget(id, params) ?? this.captureChartPasteCommandTarget(id, params)
      ?? this.captureWorkbookCommandTarget(id, params)
    const event: ApiEventMap['BeforeCommandExecute'] = {
      id,
      params,
      unitId: commandTarget?.unitId ?? commandUnitId(params),
      kind: this.commandKind(id, params),
    }
    this.events.emit('BeforeCommandExecute', event)
    if (event.cancel) throw new Error(`Command was canceled: ${id}`)
    commandTarget?.assertCurrent()
    const result = (handler as OfficeCommandHandler<P, R>)(params)
    this.publishObserverEvent('CommandExecuted', () => ({ ...event, ...(params === undefined ? {} : { params: cloneModel(params) }) }))
    return result
  }

  executeCommand<P = Record<string, unknown>, R = boolean>(id: string, params?: P): Promise<R> {
    return Promise.resolve().then(async () => {
      this.assertAlive()
      const handler = this.commands.get(id)
      if (!handler) throw new Error(`Unknown command: ${id}`)
      if (contentHistoryCommandAction(id)) return (handler as OfficeCommandHandler<P | undefined, R | Promise<R>>)(params)
      const commandTarget = this.captureDocumentTransientCommandTarget(id, params) ?? this.captureChartPasteCommandTarget(id, params)
        ?? this.captureWorkbookCommandTarget(id, params)
      const event: ApiEventMap['BeforeCommandExecute'] = { id, params, unitId: commandTarget?.unitId ?? commandUnitId(params), kind: this.commandKind(id, params) }
      this.events.emit('BeforeCommandExecute', event)
      if (event.cancel) throw new Error(`Command was canceled: ${id}`)
      commandTarget?.assertCurrent()
      const result = await (handler as OfficeCommandHandler<P | undefined, R | Promise<R>>)(params)
      this.publishObserverEvent('CommandExecuted', () => ({ ...event, ...(params === undefined ? {} : { params: cloneModel(params) }) }))
      return result
    })
  }

  executeDocumentEditorCommand(unitId: string, command: DocumentCommand, selection: DocumentSelection, sourceCommandId?: 'doc.command.insert-special-character'): DocumentEditorCommandResult {
    this.assertAlive()
    if (command.id === 'doc.command.set-document-permission-rules') {
      const changed = this.executeCommandSync<{ unitId: string; rules: DocumentModel['permissionRules'] }, boolean>(command.id, { unitId, rules: command.rules })
      return { changed, selection: cloneDocumentSelection(selection) }
    }
    if (!Object.hasOwn(DOCUMENT_EDITOR_COMMAND_LABELS, command.id)) throw new Error(`Unknown document editor command: ${command.id}`)
    const expectedRevision = this.getDocumentRevision(unitId), inputGeneration = this.documentInput.get(unitId)?.generation
    const expectedHistoryGeneration = this.editor.history.captureSavepoint(unitId).generation
    const effective = command.id === 'doc.command.insert-text' || command.id === 'doc.command.insert-latex-formula'
      ? { ...command, marks: mergeDocumentInputMarks(this.getDocumentInputMarks(unitId, selection, undefined, command.host), command.marks ?? {}) } : command
    return this.executeCommandSync<DocumentEditorCommandRequest & { value?: string }, DocumentEditorCommandResult>(sourceCommandId ?? command.id,
      { unitId, command: effective, selection, expectedRevision, expectedHistoryGeneration, inputGeneration,
        ...(sourceCommandId && effective.id === 'doc.command.insert-text' ? { value: effective.text } : {}) })
  }

  executeDocumentNoteCommand(unitId: string, request: DocumentNoteMutationRequest): DocumentNoteCommandResult {
    this.assertAlive()
    const session = this.documentInput.get(unitId), expectedRevision = this.getDocumentRevision(unitId)
    const expectedHistoryGeneration = this.editor.history.captureSavepoint(unitId).generation
    const activeId = this._activeDocumentId, activeKind = this.editor.state.activeKind, userId = this.currentUserId
    const params: DocumentNoteCommandParams = { ...request, unitId, expectedRevision, expectedHistoryGeneration,
      inputGeneration: session?.generation, ...(session?.location.kind === 'body' ? { host: session.location.host } : {}), feedback: true }
    const result = this.executeCommandSync<DocumentNoteCommandParams, DocumentNoteCommandResult>(documentNoteCommandId(request), params)
    const latest = this.documentInput.get(unitId)
    // Generic CommandExecuted observers run after the handler; recheck before any controller restores focus.
    const selectionCurrent = result.selectionCurrent && !this.disposed && !!this.getDocumentModel(unitId)
      && this._activeDocumentId === activeId && this.editor.state.activeKind === activeKind && this.currentUserId === userId
      && this.getDocumentRevision(unitId) === result.modelRevision && this.editor.history.captureSavepoint(unitId).generation === expectedHistoryGeneration
      && latest?.generation === result.inputGeneration
      && (!result.selection || latest?.valid === true && sameDocumentInputLocation(latest.location, documentInputLocation(result.selection, undefined, result.host)))
    return { ...result, selectionCurrent }
  }

  /** The ten note commands share one candidate, content-history entry and checked input publication. */
  commitDocumentNoteMutation(input: DocumentNoteCommandContext, request: DocumentNoteMutationRequest): DocumentNoteCommandResult {
    this.assertAlive()
    const current = this.getDocumentModel(input.unitId)
    if (!current) throw new DocumentNoteCommandRejectedError('注释所属文档已关闭。')
    const revision = this.getDocumentRevision(input.unitId), savepoint = this.editor.history.captureSavepoint(input.unitId)
    const session = this.documentInput.get(input.unitId), generation = session?.generation
    for (const [name, value, expected] of [
      ['expectedRevision', input.expectedRevision, revision], ['expectedHistoryGeneration', input.expectedHistoryGeneration, savepoint.generation],
      ['inputGeneration', input.inputGeneration, generation],
    ] as const) {
      if (Object.hasOwn(input, name) && (value !== undefined && (!Number.isSafeInteger(value) || value < 0) || value !== expected)) {
        throw new DocumentNoteCommandRejectedError('注释所属文档、历史或输入位置已更改，请重新执行。')
      }
    }
    if (Object.hasOwn(input, 'host') && session?.valid && (session.location.kind !== 'body' || !sameDocumentBlockHost(session.location.host, input.host))) {
      throw new DocumentNoteCommandRejectedError('注释命令的输入内容区已更改。')
    }
    const activeId = this._activeDocumentId, activeKind = this.editor.state.activeKind, userId = this.currentUserId
    let plan: ReturnType<typeof prepareDocumentNoteMutation>
    try { plan = prepareDocumentNoteMutation(current, request, { documentId: input.unitId, userId }) }
    catch (error) {
      if (error instanceof TypeError || error instanceof RangeError || error instanceof DocumentPermissionError) {
        throw new DocumentNoteCommandRejectedError(error.message, error)
      }
      throw error
    }
    const sourceLocation = session?.valid && session.modelRevision === revision ? session.location : undefined
    const usesPlanSelection = request.type === 'insert'
      ? !sourceLocation || sourceLocation.kind === 'body' && !sourceLocation.host && sameDocumentInputLocation(sourceLocation, documentInputLocation(request.selection))
      : request.type === 'delete' && (!sourceLocation || sourceLocation.kind === 'body' && sourceLocation.host?.kind === 'note'
        && sourceLocation.host.noteKind === request.kind && sourceLocation.host.noteId === request.noteId)
    const changed = plan.changed && this.commitDocument(input.unitId, request.type === 'insert' ? '插入注释' : request.type === 'delete' ? '删除注释'
      : request.type === 'settings' ? '设置注释格式' : '修改注释', model => {
        for (const key of Object.keys(model)) if (!Object.hasOwn(plan.model, key)) Reflect.deleteProperty(model, key)
        Object.assign(model, plan.model)
        model.wordCount = textTargets(model).reduce((count, target) => count + runText(runsForTarget(model, target)).replace(/\s/g, '').length, 0)
      })
    const modelRevision = revision + (changed ? 1 : 0)
    let selection: DocumentSelection | undefined, host: DocumentAuxiliaryBlockHost | undefined, expectedGeneration = generation
    const currentContext = () => !this.disposed && !!this.getDocumentModel(input.unitId) && this.currentUserId === userId
      && this.getDocumentRevision(input.unitId) === modelRevision && this._activeDocumentId === activeId && this.editor.state.activeKind === activeKind
      && this.editor.history.captureSavepoint(input.unitId).generation === savepoint.generation
      && this.documentInput.get(input.unitId)?.generation === expectedGeneration
    if (changed && currentContext() && activeKind === 'document' && activeId === input.unitId) {
      if (usesPlanSelection && plan.selection) {
        selection = cloneDocumentSelection(plan.selection); host = plan.host
        this.documentInput.select(input.unitId, documentInputLocation(selection, undefined, host), modelRevision)
        expectedGeneration = (generation ?? 0) + 1
        this.publishDocumentInputState(input.unitId)
      } else if (sourceLocation?.kind === 'body' && session) {
        try {
          const location = mapDocumentNoteLocation(current, this.getDocumentModel(input.unitId)!, sourceLocation.selection, sourceLocation.host)
          selection = location.selection; host = location.host
        }
        catch (error) { if (!(error instanceof RangeError)) throw error }
        if (selection && this.documentInput.completeUnchangedSelection(input.unitId, session, modelRevision, documentInputLocation(selection, undefined, host), true)) {
          expectedGeneration = (generation ?? 0) + 1
          this.publishDocumentInputState(input.unitId)
        } else { selection = undefined; host = undefined }
      }
    } else if (!changed && sourceLocation?.kind === 'body') {
      selection = cloneDocumentSelection(sourceLocation.selection); host = sourceLocation.host
    }
    return { changed, selectionCurrent: !!selection && currentContext(), modelRevision, inputGeneration: expectedGeneration,
      ...(selection ? { selection } : {}), ...(host ? { host: { ...host } } : {}), ...(plan.noteId ? { noteId: plan.noteId } : {}) }
  }

  registerCommand<P, R>(id: string, handler: OfficeCommandHandler<P, R>, options: { asynchronous?: boolean } = {}): Disposable {
    this.assertAlive()
    if (!id || this.commands.has(id)) throw new Error(`Command already registered: ${id}`)
    this.commands.set(id, handler as OfficeCommandHandler<never, unknown>)
    if (options.asynchronous) this.asynchronousCommands.add(id)
    let active = true
    return { dispose: () => { if (active) { active = false; this.commands.delete(id); this.asynchronousCommands.delete(id) } } }
  }

  workbenchUiHost(unitId: string): WorkbenchUiHost {
    return {
      bind: run => {
        this.workbenchUiOwners.set(unitId, run)
        return { dispose: () => { if (this.workbenchUiOwners.get(unitId) === run) this.workbenchUiOwners.delete(unitId) } }
      },
      execute: id => this.executeCommandSync<{ unitId: string }, boolean>(id, { unitId }),
    }
  }

  undo(kind: OfficeKind, id: string): boolean {
    return this.executeCommandSync(CONTENT_HISTORY_COMMANDS.undo, { kind, unitId: id })
  }

  redo(kind: OfficeKind, id: string): boolean {
    return this.executeCommandSync(CONTENT_HISTORY_COMMANDS.redo, { kind, unitId: id })
  }

  private replayContentHistory(kind: OfficeKind, id: string, direction: 'undo' | 'redo'): boolean {
    this.assertAlive()
    if (kind === 'workbook') this.activateWorkbook(id)
    else this.activateDocument(id)
    const history = this.editor.history, stack = history.units[id], entry = stack?.[direction].at(-1)
    const revision = kind === 'workbook' ? this.getWorkbookRevision(id) : this.getDocumentRevision(id)
    const hostCurrent = () => !this.disposed && this.editor.state.activeKind === kind && history.activeUnit === id
      && (kind === 'workbook' ? this._activeWorkbookId : this._activeDocumentId) === id
    if (!hostCurrent()) return false
    const savepoint = history.captureSavepoint(id)
    const event: ApiEventMap['BeforeUndo'] = { unitId: id, kind }
    this.events.emit(direction === 'undo' ? 'BeforeUndo' : 'BeforeRedo', event)
    if (event.cancel) throw new ContentHistoryCanceledError(direction)
    const model = kind === 'workbook' ? this.getWorkbookModel(id) : this.getDocumentModel(id)
    if (!stack || !entry || !model || !hostCurrent() || history.units[id] !== stack || stack.revision !== savepoint.revision
      || stack.generation !== savepoint.generation || stack[direction].at(-1) !== entry
      || (kind === 'workbook' ? this.getWorkbookRevision(id) : this.getDocumentRevision(id)) !== revision) return false
    const generation = this.documentInput.get(id)?.generation
    const label = direction === 'undo' ? this.editor.undo() : this.editor.redo()
    if (!label) return false
    if (kind === 'workbook') {
      const next = this.getWorkbookModel(id)
      if (next) publishWorkbookSparklineChanges(this, model as WorkbookModel, next)
    }
    if (kind === 'document' && this.getDocumentModel(id) && this.getDocumentRevision(id) === revision + 1
      && this.documentInput.get(id)?.generation === generation) {
      this.documentInput.clearMarks(id)
      this.publishDocumentInputState(id)
    }
    this.publishObserverEvent(direction === 'undo' ? 'Undo' : 'Redo', () => ({ unitId: id, kind, label }))
    return true
  }

  markSaved(kind: OfficeKind, id: string): void {
    this.assertAlive()
    this.assertContentUnitIdentity(kind, id)
    if (kind === 'workbook') {
      const model = this.getWorkbookModel(id)
      if (!model) throw new Error(`Workbook not found: ${id}`)
      this.editor.history.markSaved(id)
      this.replaceWorkbookModel(id, { ...model, modified: false })
    } else {
      const model = this.getDocumentModel(id)
      if (!model) throw new Error(`Document not found: ${id}`)
      this.editor.history.markSaved(id)
      this.replaceDocumentModel(id, { ...model, modified: false })
    }
  }

  disposeUnit(id: string): boolean {
    this.assertAlive()
    if (this.workbooks.has(id)) {
      disposeSheetFilterPanelEditors(this, id)
      disposeSheetFunctionEditors(this, id)
      this.sheetClipboard.clearUnit(id)
      chartObjectClipboardStore(this.pinia).clearUnit(id)
      this.sheetFill.clear(id)
      this.sheetSelections.removeUnit(id)
      this.chartStatisticsService.forgetDocument(id)
      this.slicerMemberQueryService.forgetDocument(id)
      this.computeClient.forgetDocument(id)
      this.workbookRevisions.delete(id)
      this.borderFormatStates.delete(id)
      this.workbooks.delete(id)
      this.editor.history.clear(id)
      this.events.emit('WorkbookDisposed', { workbookId: id })
      if (id === this._activeWorkbookId) {
        const next = this.createWorkbook({}, false)
        this.activateWorkbook(next.id!)
      }
      return true
    }
    if (this.documents.has(id)) {
      disposeDocumentImageEditors(this, id)
      disposeDocumentLatexTemplateEditors(this, id)
      disposeDocumentLatexPopupEditors(this, id)
      disposeDocumentHeaderFooterEditors(this, id)
      cancelDocumentTocPagination(this, id); disposeDocumentTocLayouts(this, id)
      disposeDocumentTocDialogEditors(this, id)
      this.documentFieldCacheAdapters.delete(id)
      this.documentInput.removeUnit(id)
      this.pendingDocumentInputPublications.delete(id)
      this.documentRevisions.delete(id)
      this.documents.delete(id)
      this.editor.history.clear(id)
      this.events.emit('DocDisposed', { unitId: id, documentId: id })
      if (id === this._activeDocumentId) {
        const next = this.createDocument({}, false)
        this.activateDocument(next.id!)
      }
      return true
    }
    return false
  }

  dispose(): void {
    if (this.disposed) return
    disposeSheetFilterPanelEditors(this)
    disposeSheetFunctionEditors(this)
    disposeDocumentImageEditors(this)
    disposeDocumentHeaderFooterEditors(this)
    cancelDocumentTocPagination(this); disposeDocumentTocLayouts(this)
    disposeDocumentTocDialogEditors(this)
    this.disposed = true
    for (const controller of this.workbookCreationTasks) controller.abort()
    this.workbookCreationTasks.clear()
    disposeDocumentLatexTemplateEditors(this)
    disposeDocumentLatexPopupEditors(this)
    for (const stop of this.stopModelSync) stop()
    this.editor.parser.cancel()
    this.chartStatisticsService.dispose()
    this.slicerMemberQueryService.dispose()
    if (this.ownsComputeClient) this.computeClient.dispose()
    this.events.emit('LifeCycleChanged', { stage: 'disposed' })
    this.events.clear()
    this.commands.clear()
    this.asynchronousCommands.clear()
    this.workbenchUiOwners.clear()
    this.workbooks.clear()
    this.borderFormatStates.clear()
    this.sheetSelections.clear()
    this.sheetSelections.$dispose()
    this.sheetClipboard.clear()
    this.sheetClipboard.$dispose()
    const chartClipboard = chartObjectClipboardStore(this.pinia)
    chartClipboard.clear()
    chartClipboard.$dispose()
    clearChartObjectUIHosts(this)
    this.sheetFill.clear()
    this.sheetFill.$dispose()
    this.documentInput.clear()
    this.pendingDocumentInputPublications.clear()
    this.documentInput.$dispose()
    this.documentCharacterPreferences.releaseStorage()
    this.documentCharacterPreferences.$dispose()
    this.documentRevisions.clear()
    this.documentFieldCacheAdapters.clear()
    this.documents.clear()
  }

  private replaceWorkbookModel(id: string, model: WorkbookModel, publishSelection = true): void {
    const activeSheetId = this.activeWorksheetIdForView(model)
    if (model.activeSheetId !== activeSheetId) model = { ...model, activeSheetId }
    this.workbooks.set(id, model)
    if (id === this._activeWorkbookId) {
      this.workbookSelectionPublicationDepth++
      try { this.editor.workbook.replace(model) }
      finally { this.workbookSelectionPublicationDepth-- }
      if (publishSelection && !this.workbookSelectionPublicationDepth) this.publishSheetSelection(id, model.activeSheetId)
    } else {
      this.advanceWorkbookRevision(id)
      this.reconcileWorkbookSelections(id, model)
    }
  }

  private applyWorkbookOperation(id: string, update: (draft: WorkbookModel) => void): boolean {
    this.assertAlive()
    this.assertContentUnitIdentity('workbook', id)
    const current = this.getWorkbookModel(id)
    if (!current) throw new Error(`Workbook not found: ${id}`)
    const next = cloneModel(current)
    update(next)
    assertWorkbookConditionalFormatBudget(next.sheets)
    assertWorksheetProtectionChanges(current, next, this.currentUserId)
    this.assertChartHostChanges(current, next)
    if (next.activeSheetId && next.activeSheetId !== current.activeSheetId) {
      this.assertCanViewWorksheet(id, next.activeSheetId)
      if (requireSheet(next, next.activeSheetId).visible === false) throw new Error('工作表已隐藏，请先取消隐藏再激活。')
    }
    next.id = id
    next.modified = current.modified
    if (JSON.stringify(current) === JSON.stringify(next)) return false
    this.replaceWorkbookModel(id, next)
    return true
  }

  private replaceWorkbookHistorySnapshot(id: string, snapshot: WorkbookModel, publishSelection = true): void {
    this.assertAlive()
    this.assertContentUnitIdentity('workbook', id)
    const current = this.getWorkbookModel(id)
    const next = cloneModel(snapshot)
    if (!current) throw new Error(`Workbook not found: ${id}`)
    this.assertChartHostChanges(current, next)
    if (current && next.sheets.some(sheet => sheet.id === current.activeSheetId && sheet.visible !== false
      && canViewWorksheetProtection(sheet, this.currentUserId))) next.activeSheetId = current.activeSheetId
    for (const sheet of next.sheets) this.sheetSelections.setTabStart(id, sheet.id)
    this.replaceWorkbookModel(id, next, publishSelection)
  }

  /** Chart mutations and history share actual old/new host-sheet authorization. */
  private assertChartHostChanges(before: WorkbookModel, after: WorkbookModel): void {
    if (JSON.stringify(before.charts ?? []) === JSON.stringify(after.charts ?? [])) return
    const previous = new Map((before.charts ?? []).map(chart => [chart.id, chart])), next = new Map((after.charts ?? []).map(chart => [chart.id, chart]))
    if (previous.size !== (before.charts?.length ?? 0) || next.size !== (after.charts?.length ?? 0)) throw new Error('图表 ID 必须唯一。')
    const changed = new Set([...new Set([...previous.keys(), ...next.keys()])]
      .filter(id => JSON.stringify(previous.get(id)) !== JSON.stringify(next.get(id))))
    for (const sheetId of new Set([...(before.charts ?? []).map(chart => chart.sheetId), ...(after.charts ?? []).map(chart => chart.sheetId)])) {
      const previousOrder = (before.charts ?? []).filter(chart => chart.sheetId === sheetId && next.get(chart.id)?.sheetId === sheetId).map(chart => chart.id)
      const nextOrder = (after.charts ?? []).filter(chart => chart.sheetId === sheetId && previous.get(chart.id)?.sheetId === sheetId).map(chart => chart.id)
      if (JSON.stringify(previousOrder) !== JSON.stringify(nextOrder)) for (const id of previousOrder) changed.add(id)
    }
    for (const id of changed) for (const chart of [previous.get(id), next.get(id)]) {
      if (!chart) continue
      const hosts = [before.sheets.find(sheet => sheet.id === chart.sheetId), after.sheets.find(sheet => sheet.id === chart.sheetId)].filter((sheet): sheet is SpreadsheetSheet => !!sheet)
      if (!hosts.length) throw new Error('图表宿主工作表不存在：' + chart.sheetId)
      if (hosts.some(sheet => !canEditWorksheetProtection(sheet, this.currentUserId))) throw new Error('图表宿主工作表受保护，当前用户没有编辑权限。')
    }
  }

  private advanceWorkbookRevision(id: string): void {
    const revision = (this.workbookRevisions.get(id) ?? 0) + 1
    this.workbookRevisions.set(id, revision)
    this.sheetFill.clearSession(id)
    this.numericIndexService.setDocumentRevision(id, revision)
    this.chartStatisticsService.setDocumentRevision(id, revision)
  }

  private replaceDocumentModel(id: string, model: DocumentModel): void {
    const previous = this.getDocumentModel(id)
    if (previous && previous !== model) cancelDocumentImageTasksForSource(previous)
    this.documents.set(id, model)
    if (id === this._activeDocumentId) this.editor.document.replace(model)
    else if (previous && !equivalentDocumentContent(previous, model)) this.advanceDocumentRevision(id)
    if (this.documentInput.revalidateModel(id, model, previous)) this.publishDocumentInputState(id)
  }

  private replaceDocumentHistorySnapshot(id: string, snapshot: DocumentModel, permissionAdminCommand: boolean): void {
    this.assertAlive()
    this.assertContentUnitIdentity('document', id)
    const current = this.getDocumentModel(id)
    if (!current) throw new Error(`Document not found: ${id}`)
    const next = cloneModel(snapshot)
    preserveDocumentZoomRatio(next, current)
    if (permissionAdminCommand) {
      if (this.currentUserId !== this.permissionAdminUserId) throw new Error('只有文档权限管理员可以撤销或重做权限规则。')
    } else {
      if (JSON.stringify(current.permissionRules ?? []) !== JSON.stringify(next.permissionRules ?? [])) throw new Error('内容历史不能覆盖当前的文档权限规则。')
      assertDocumentPermission(current.permissionRules ?? [], { documentId: id, userId: this.currentUserId }, 'edit', documentNoteMutationPermissionTargets(current, next))
    }
    this.replaceDocumentModel(id, next)
  }

  private advanceDocumentRevision(id: string): void {
    this.documentRevisions.set(id, (this.documentRevisions.get(id) ?? 0) + 1)
    if (this.documentInput.invalidate(id)) this.publishDocumentInputState(id)
  }

  private publishDocumentInputState(unitId: string): void {
    if (this.editor.history.isApplying) { this.pendingDocumentInputPublications.add(unitId); return }
    this.pendingDocumentInputPublications.delete(unitId)
    const session = this.documentInput.get(unitId), revision = this.getDocumentModel(unitId) ? this.getDocumentRevision(unitId) : undefined
    const snapshot: DocumentInputSession | undefined = session ? { ...session, marks: this.getDocumentInputMarks(unitId) } : undefined
    this.publishObserverEvent('DocumentInputStateChanged', () => ({ unitId, ...(snapshot ? { state: cloneDocumentInputSession(snapshot) } : {}) }),
      () => this.documentInput.get(unitId)?.generation === session?.generation
        && this.documentInput.get(unitId)?.valid === session?.valid
        && (!this.getDocumentModel(unitId) ? revision === undefined : this.getDocumentRevision(unitId) === revision))
  }

  private commitDocumentEditorCommand(commandId: string, request: DocumentEditorCommandRequest, amendment?: HistorySavepoint): DocumentEditorCommandResult {
    assertDocumentEditorCommandRequest(request, commandId)
    if (request.command.id === 'doc.command.set-document-permission-rules') throw new Error('Permission rules use their administrator command')
    const current = this.getDocumentModel(request.unitId)
    if (!current) throw new Error(`Document not found: ${request.unitId}`)
    const revision = this.getDocumentRevision(request.unitId), session = this.documentInput.get(request.unitId)
    const generation = session?.generation
    if (request.expectedRevision !== undefined && request.expectedRevision !== revision) throw new Error('文档内容已更改，请重新执行编辑。')
    if (request.expectedHistoryGeneration !== undefined
      && request.expectedHistoryGeneration !== this.editor.history.captureSavepoint(request.unitId).generation) {
      throw new Error('文档历史或打开代数已更改，请重新执行编辑。')
    }
    const sourceCurrent = !Object.hasOwn(request, 'inputGeneration') || request.inputGeneration === generation
    if (documentContentControlCommandIsTextInput(request.command) && !sourceCurrent) throw new Error('输入选区已更改，请重新执行编辑。')
    const fieldRefresh = request.command.id === 'doc.command.refresh-header-footer-fields' || request.command.id === 'doc.command.update-selected-fields' ? request.command : undefined
    const tocRefresh = request.command.id === 'doc.command.update-table-of-contents' || request.command.id === 'doc.command.delete-table-of-contents'
    if (fieldRefresh && !sourceCurrent) throw new Error('字段更新选区已更改，请重新执行。')
    if (tocRefresh && !sourceCurrent) throw new Error('目录更新选区已更改，请重新执行。')
    const activeId = this._activeDocumentId, activeKind = this.editor.state.activeKind
    const hostCurrent = () => this._activeDocumentId === activeId && this.editor.state.activeKind === activeKind
    const context = { documentId: request.unitId, userId: this.currentUserId, permissionAdminUserId: this.permissionAdminUserId }
    const prepared = prepareCollapsedDocumentInputFormat(current, request, context)
    if (prepared) {
      if (!sourceCurrent) throw new Error('输入选区已更改，请重新设置格式。')
      if (this.getDocumentRevision(request.unitId) !== revision) throw new Error('文档内容已更改，请重新设置输入格式。')
      const inputChanged = this.documentInput.format(request.unitId, prepared.location, prepared.patch, revision)
      const expectedGeneration = inputChanged ? (generation ?? 0) + 1 : generation
      if (inputChanged) this.publishDocumentInputState(request.unitId)
      const latest = this.documentInput.get(request.unitId)
      const selectionCurrent = !!this.getDocumentModel(request.unitId) && this.getDocumentRevision(request.unitId) === revision
        && hostCurrent() && latest?.valid === true && latest.generation === expectedGeneration
        && sameDocumentInputLocation(latest.location, prepared.location)
      return { changed: false, selection: prepared.selection, inputChanged, selectionCurrent, modelRevision: revision,
        ...(prepared.location.kind === 'body' && prepared.location.host ? { host: { ...prepared.location.host } } : {}) }
    }
    const globalCommand = request.command.id === 'doc.command.set-section-header-footer' || request.command.id === 'doc.command.set-section-setup'
      || request.command.id === 'docs.command.page-setup' || request.command.id === SET_DOCUMENT_HEADING_COLLAPSED_COMMAND || !!fieldRefresh || tocRefresh
    // Global commands still operate on the full model. Their view selection
    // retains its independent block host instead of being resolved as body text.
    const selectionHost = request.selectionHost !== undefined ? request.selectionHost ?? undefined
      : request.command.host ?? (globalCommand && sourceCurrent && session?.valid && session.modelRevision === revision
        && session.location.kind === 'body' ? session.location.host : undefined)
    const sourceModel = selectionHost ? documentModelForBlockHost(current, selectionHost) : current
    const sourceSelectionUsable = sourceModel.blocks.length > 0
    const sourceSelection = request.command.id === DOCUMENT_PAGE_SETUP_COMMAND && !textTargets(sourceModel).length
      ? sourceSelectionUsable ? { anchor: clampPoint(sourceModel, request.selection.anchor), focus: clampPoint(sourceModel, request.selection.focus) }
        : cloneDocumentSelection(request.selection)
      : resolveDocumentEditorSelection(current, request.selection, selectionHost)
    const sessionStory = globalCommand && request.selectionHost === undefined && sourceCurrent && session?.valid
      && session.modelRevision === revision && session.location.kind === 'story' ? session.location.range : undefined
    const story = request.command.id === 'doc.command.set-section-header-footer' ? request.command.input?.range ?? sessionStory
      : request.command.id === 'doc.command.format-text' || request.command.id === 'doc.command.clear-formatting' ? request.command.story
      : sessionStory
    const preserveTocAuxiliaryInput = (request.command.id === SET_DOCUMENT_HEADING_COLLAPSED_COMMAND
      || (request.command.id === DOCUMENT_TOC_UPDATE_COMMAND || request.command.id === DOCUMENT_TOC_DELETE_COMMAND)
        && request.command.preserveInputSelection === true) && sourceCurrent && session?.valid === true
      && (session.location.kind === 'story' || !!session.location.host)
    const sourceLocation = preserveTocAuxiliaryInput ? cloneDocumentInputSession(session!).location : documentInputLocation(sourceSelection, story, selectionHost)
    const tracksSelection = sourceSelectionUsable && sourceCurrent && (!session?.valid || sameDocumentInputLocation(session.location, sourceLocation))
    let input: RunPatch | undefined
    let command = request.command
    if (command.id === 'doc.command.insert-text' || command.id === 'doc.command.insert-latex-formula') {
      input = mergeDocumentInputMarks(sourceCurrent ? this.getDocumentInputMarks(request.unitId, sourceSelection, undefined, selectionHost) : {}, command.marks ?? {})
      command = { ...command, marks: input }
    }
    if (command.id === DOCUMENT_TOC_INSERT_COMMAND) input = {}
    const mutationSelection = selectionHost && !command.host && command.id !== DOCUMENT_PAGE_SETUP_COMMAND
      ? sanitizeDocumentInputSelection(current, sourceSelection) : sourceSelection
    const effective: DocumentEditorCommandRequest = { ...request, selection: mutationSelection, command }
    let selection = sourceSelection, storyAdvance: DocumentInputLocation | undefined
    const selectedFieldHost = fieldRefresh?.id === 'doc.command.update-selected-fields'
      ? fieldRefresh.host ?? (!fieldRefresh.bodySelection ? selectionHost : undefined) : fieldRefresh?.host
    const fieldRequest: DocumentStoryFieldRefreshRequest | undefined = fieldRefresh ? { locale: fieldRefresh.locale, timestamp: fieldRefresh.timestamp,
      onOpen: fieldRefresh.onOpen, ...(selectedFieldHost ? { host: selectedFieldHost } : {}),
      ...(fieldRefresh.id === 'doc.command.update-selected-fields' ? { stories: fieldRefresh.stories ?? [],
        bodySelection: fieldRefresh.bodySelection ?? (!fieldRefresh.stories?.length ? sourceSelection : undefined) } : {}) } : undefined
    // An opening host can have a different local caret from the existing input
    // session. Map both through the same field replacements; cache publication
    // must not discard the authoritative session's typing preferences.
    const cacheSession = fieldRequest?.onOpen && session?.valid && session.modelRevision === revision ? session : undefined
    const fieldRefreshPlan = fieldRequest ? prepareDocumentStoryFieldRefresh(current, fieldRequest, context, sourceSelection,
      cacheSession?.location.kind === 'body' ? [{ selection: cacheSession.location.selection, host: cacheSession.location.host }] : undefined, selectionHost) : undefined
    const cacheInputLocation = cacheSession && (cacheSession.location.kind === 'body'
      ? documentInputLocation(fieldRefreshPlan?.mappedAdditionalBodySelections?.[0] ?? cacheSession.location.selection, undefined,
        fieldRefreshPlan?.mappedAdditionalBodySelectionHosts?.[0] ?? cacheSession.location.host)
      : cacheSession.location)
    // Read-only/unsupported/no-selection refreshes succeed without entering a content mutation.
    if (fieldRefreshPlan) selection = fieldRefreshPlan.selection ?? sourceSelection
    const cachePublication = fieldRequest?.onOpen ? this.publishDocumentFieldCaches(request.unitId, fieldRequest, fieldRefreshPlan) : undefined
    let initialHeadingPublication = true
    const restoreHeadingSelection = request.command.id === SET_DOCUMENT_HEADING_COLLAPSED_COMMAND && session?.valid && tracksSelection
      ? (direction: 'undo' | 'redo') => {
        if (initialHeadingPublication) { initialHeadingPublication = false; return }
        const model = this.getDocumentModel(request.unitId)
        if (!model) return
        const original = cloneDocumentInputSession(session).location
        const location = direction === 'undo' || original.kind === 'story' || original.host
          ? original : documentInputLocation(cloneDocumentSelection(selection))
        this.documentInput.select(request.unitId, location, this.getDocumentRevision(request.unitId))
        this.publishDocumentInputState(request.unitId)
      } : undefined
    const changed = cachePublication ? cachePublication.changed
      : (!fieldRefresh || fieldRefreshPlan !== undefined) && this.commitDocumentTransaction(request.unitId, DOCUMENT_EDITOR_COMMAND_LABELS[request.command.id], model => {
      if (fieldRefreshPlan) {
        applyDocumentFieldRefreshPlan(model, fieldRefreshPlan)
        selection = fieldRefreshPlan.selection ?? sourceSelection
        return
      }
      const appliedSelection = applyDocumentEditorCommand(model, effective, context)
      const scoped = selectionHost && !command.host ? documentModelForBlockHost(model, selectionHost) : undefined
      selection = scoped?.blocks.length ? sanitizeDocumentInputSelection(scoped, sourceSelection, true)
        : scoped ? cloneDocumentSelection(sourceSelection) : appliedSelection
      if (request.command.id === 'doc.command.set-section-header-footer' && request.command.input) {
        storyAdvance = prepareDocumentStoryInputAdvance(current, model, sourceSelection, request.command.sectionId,
          request.command.input, session, revision, context)
      }
    }, false, amendment, restoreHeadingSelection)
    let expectedGeneration = generation
    let fieldInputPreserved = true
    const modelRevision = revision + (changed ? 1 : 0)
    if (this.getDocumentModel(request.unitId) && this.getDocumentRevision(request.unitId) === modelRevision
      && this.documentInput.get(request.unitId)?.generation === generation && hostCurrent()) {
      if (!sourceSelectionUsable) {
        fieldInputPreserved = false
        if (this.documentInput.invalidate(request.unitId)) this.publishDocumentInputState(request.unitId)
      } else if (changed && cacheInputLocation && cacheSession) {
        fieldInputPreserved = this.documentInput.completeUnchangedSelection(request.unitId, cacheSession, modelRevision,
          cacheInputLocation, cacheInputLocation.kind === 'body')
        if (!fieldInputPreserved) this.documentInput.clearMarks(request.unitId)
        expectedGeneration = (generation ?? 0) + 1
        this.publishDocumentInputState(request.unitId)
      } else if (changed && fieldRequest?.onOpen) {
        // Already invalid drafts stay invalid. A derived cache update neither
        // revives a stale location nor clears its retained typing preferences.
      } else if (changed && globalCommand && tracksSelection && session?.valid) {
        fieldInputPreserved = this.documentInput.completeUnchangedSelection(request.unitId, session, modelRevision,
          preserveTocAuxiliaryInput ? sourceLocation : sourceLocation.kind === 'body'
            ? documentInputLocation(selection, undefined, fieldRefreshPlan?.host ?? selectionHost) : sourceLocation, sourceLocation.kind === 'body')
        if (!fieldInputPreserved) this.documentInput.clearMarks(request.unitId)
        expectedGeneration = (generation ?? 0) + 1
        this.publishDocumentInputState(request.unitId)
      } else if (storyAdvance && generation !== undefined) {
        this.documentInput.allowAdvance(request.unitId, storyAdvance, modelRevision, generation)
      } else if (changed && input && tracksSelection) {
        const location = documentInputLocation(selection, undefined, selectionHost)
        if (this.documentInput.completeInput(request.unitId, location, input, revision + 1, generation)) {
          expectedGeneration = (generation ?? 0) + 1
          this.publishDocumentInputState(request.unitId)
        }
      } else if (changed) {
        this.documentInput.clearMarks(request.unitId)
        expectedGeneration = generation === undefined ? undefined : generation + 1
        this.publishDocumentInputState(request.unitId)
      } else if (tracksSelection && sourceLocation.kind === 'body' && !sameDocumentInputLocation(sourceLocation, documentInputLocation(selection, undefined, selectionHost))) {
        this.documentInput.select(request.unitId, documentInputLocation(selection, undefined, selectionHost), revision)
        expectedGeneration = (generation ?? 0) + 1
        this.publishDocumentInputState(request.unitId)
      }
    }
    const selectionCurrent = tracksSelection && fieldInputPreserved && !!this.getDocumentModel(request.unitId) && hostCurrent()
      && this.getDocumentRevision(request.unitId) === modelRevision && this.documentInput.get(request.unitId)?.generation === expectedGeneration
    const resultHost = fieldRefreshPlan?.host ?? selectionHost
    return { changed, selection, selectionCurrent, modelRevision, ...(resultHost ? { host: { ...resultHost } } : {}),
      ...(cachePublication ? { fieldCache: { ...cachePublication.state, inputGeneration: expectedGeneration } } : {}) }
  }

  private commitIndexedDocumentInput(unitId: string, startOffset: number, endOffset: number, text: string, expectedRevision?: number): boolean {
    if (typeof unitId !== 'string' || !unitId.trim() || typeof text !== 'string') throw new TypeError('文档输入参数无效。')
    const current = this.getDocumentModel(unitId)
    if (!current) throw new Error(`Document not found: ${unitId}`)
    if (expectedRevision !== undefined && (!Number.isSafeInteger(expectedRevision) || expectedRevision !== this.getDocumentRevision(unitId))) {
      throw new Error('文档内容已更改，请重新执行编辑。')
    }
    const selection = documentSelectionAtBodyRange(current, startOffset, endOffset)
    assertDocumentPermission(current.permissionRules ?? [], { documentId: unitId, userId: this.currentUserId }, 'edit',
      permissionTargetsForRange(current, startOffset, endOffset))
    if (selection) return this.commitDocumentEditorCommand('doc.command.insert-text', {
      unitId, selection, command: { id: 'doc.command.insert-text', text },
    }).changed
    if (!text) return false
    return this.commitDocument(unitId, '输入文本', model => {
      const paragraph = createParagraph('', model.defaultStyle?.text)
      if (model.defaultStyle) paragraph.style = { ...model.defaultStyle.paragraph }
      model.blocks.unshift(paragraph)
      const point = { blockId: paragraph.id, offset: 0 }
      applyDocumentEditorCommand(model, { unitId, selection: { anchor: point, focus: point }, command: { id: 'doc.command.insert-text', text } },
        { documentId: unitId, userId: this.currentUserId })
    })
  }

  private registerBuiltinCommands(): void {
    for (const id of [OPEN_FEATURE_SEARCH, TOGGLE_SHORTCUT_PANEL] as const) {
      this.registerCommand<unknown, boolean>(id, params => {
        const activeId = this.editor.state.activeKind === 'workbook' ? this._activeWorkbookId : this._activeDocumentId
        const requestedId = commandUnitId(params)
        if (requestedId !== undefined && requestedId !== activeId) return false
        return this.workbenchUiOwners.get(activeId)?.(id) ?? false
      })
    }
    for (const id of CONTENT_HISTORY_COMMAND_IDS) {
      const action = contentHistoryCommandAction(id)
      if (!action) throw new Error(`Unknown history command: ${id}`)
      this.registerCommand<unknown, boolean>(id, params => {
        const target = resolveContentHistoryTarget(id, params, {
          kind: this.editor.state.activeKind, workbookId: this._activeWorkbookId, documentId: this._activeDocumentId,
        }, (kind, unitId) => Boolean(kind === 'workbook' ? this.getWorkbookModel(unitId) : this.getDocumentModel(unitId)))
        return this.replayContentHistory(target.kind, target.unitId, action.direction)
      })
    }
    this.registerCommand<DocumentInputSelectionRequest, boolean>(SET_DOCUMENT_INPUT_SELECTION_OPERATION, params => {
      const request = parseDocumentInputSelectionRequest(params), model = this.getDocumentModel(request.unitId)
      if (!model) throw new Error(`Document not found: ${request.unitId}`)
      const context = { documentId: request.unitId, userId: this.currentUserId }
      resolveDocumentInputLocation(model, request.selection, request.story, context, undefined, request.host)
      if (this._activeDocumentId !== request.unitId || this.editor.state.activeKind !== 'document') this.activateDocument(request.unitId)
      if (this._activeDocumentId !== request.unitId || this.editor.state.activeKind !== 'document') return false
      const current = this.getDocumentModel(request.unitId)
      if (!current) return false
      const { location } = resolveDocumentInputLocation(current, request.selection, request.story, context, undefined, request.host)
      const changed = this.documentInput.select(request.unitId, location, this.getDocumentRevision(request.unitId))
      this.publishDocumentInputState(request.unitId)
      return changed
    })
    this.registerCommand<{ unitId: string }, boolean>(CLEAR_DOCUMENT_INPUT_MARKS_OPERATION, request => {
      if (!request || typeof request.unitId !== 'string' || !request.unitId.trim()) throw new TypeError('文档标识无效。')
      if (!this.getDocumentModel(request.unitId)) throw new Error(`Document not found: ${request.unitId}`)
      const changed = this.documentInput.clearMarks(request.unitId)
      this.publishDocumentInputState(request.unitId)
      return changed
    })
    this.registerCommand<DocumentInputSelectionRequest, boolean>(ADVANCE_DOCUMENT_INPUT_MARKS_OPERATION, params => {
      const request = parseDocumentInputSelectionRequest(params), model = this.getDocumentModel(request.unitId)
      if (!model) throw new Error(`Document not found: ${request.unitId}`)
      if (this._activeDocumentId !== request.unitId || this.editor.state.activeKind !== 'document') return false
      const { location } = resolveDocumentInputLocation(model, request.selection, request.story,
        { documentId: request.unitId, userId: this.currentUserId }, 'edit', request.host)
      const changed = this.documentInput.advance(request.unitId, location, this.getDocumentRevision(request.unitId))
      if (changed) this.publishDocumentInputState(request.unitId)
      return changed
    })
    this.registerCommand<unknown, boolean>(SET_SHEET_SELECTIONS_OPERATION, params => {
      const request = parseSheetSelectionOperation(params)
      return this.applySelections(request.unitId, request.sheetId, request.ranges, request.options)
    })
    this.registerCommand<unknown, boolean>(SELECT_SHEET_RANGE_COMMAND, params => {
      const request = parseSheetSelectionOperation(params, true)
      return this.executeCommandSync(SET_SHEET_SELECTIONS_OPERATION, request)
    })
    this.registerCommand<unknown, boolean>(MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND, params => {
      const request = parseSheetSelectionShortcut(params, { unitId: this._activeWorkbookId, sheetId: this.editor.workbook.model.activeSheetId }, unitId => {
        const workbook = this.getWorkbookModel(unitId)
        if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
        return workbook.activeSheetId
      })
      const workbook = this.getWorkbookModel(request.unitId)
      if (!workbook) throw new Error(`Workbook not found: ${request.unitId}`)
      this.assertCanViewWorksheet(request.unitId, request.sheetId)
      const sheet = requireSheet(workbook, request.sheetId), current = this.getSelectionState(request.unitId, request.sheetId)
      if (!current) return false
      const state = normalizeSheetSelection(current.ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow,
        startCol: range.startColumn, endCol: range.endColumn, ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) })), sheet, current)
      const filtered = (workbook.tables ?? []).filter(table => table.sheetId === sheet.id).flatMap(table => table.autoFilter?.filteredRows ?? [])
      const plan = planSheetSelectionShortcut(state, request.direction, request.key,
        { ...sheet, rowCount: sheet.rowCount ?? 1_048_576, columnCount: sheet.columnCount ?? 16_384, filteredRows: sheet.autoFilter?.filteredRows },
        this.sheetSelections.getTabStart(request.unitId, request.sheetId), filtered)
      if (!plan) return false
      const expected = this.apiSheetSelection(request.unitId, request.sheetId, plan.selection)
      const changed = this.executeCommandSync<unknown, boolean>(SET_SHEET_SELECTIONS_OPERATION,
        { unitId: request.unitId, sheetId: request.sheetId, ranges: expected.ranges, options: expected })
      // A listener may select another cell while the operation publishes. Keep its newer selection authoritative.
      if (JSON.stringify(this.getSelectionState(request.unitId, request.sheetId)) === JSON.stringify(expected)) {
        this.sheetSelections.setTabStart(request.unitId, request.sheetId, plan.tabStart)
      }
      return changed
    })
    registerSheetRangeEditorCommands(this)
    registerSheetFilterPanelCommands(this)
    this.stopModelSync.push(registerSheetFunctionCommands(this).dispose)
    this.stopModelSync.push(registerSheetCellInputCommands(this).dispose)
    this.stopModelSync.push(registerSheetCellInlineFormatCommands(this).dispose)
    this.stopModelSync.push(registerSheetDefaultStyleCommands(this).dispose)
    this.stopModelSync.push(registerSheetRangeThemeCommands(this).dispose)
    this.stopModelSync.push(registerSheetImageCommands(this).dispose)
    this.stopModelSync.push(registerSheetDrawingCommands(this).dispose)
    this.stopModelSync.push(registerSheetClipboardCommands(this).dispose)
    this.stopModelSync.push(registerSheetFillCommands(this).dispose)
    this.stopModelSync.push(registerSheetLifecycleCommands(this).dispose)
    this.stopModelSync.push(registerSheetAxisDimensionCommands(this).dispose)
    this.stopModelSync.push(registerSheetAxisVisibilityCommands(this).dispose)
    this.stopModelSync.push(registerSheetAxisStructureCommands(this).dispose)
    this.stopModelSync.push(registerSheetCellShiftCommands(this).dispose)
    this.stopModelSync.push(registerSheetConditionalFormattingCommands(this).dispose)
    this.stopModelSync.push(registerSheetTableCommands(this).dispose)
    this.stopModelSync.push(registerSheetValidationInputCommands(this).dispose)
    this.stopModelSync.push(registerSheetDataValidationCommands(this, this.computeClient).dispose)
    this.stopModelSync.push(registerSheetValidationDropdownOperations(this).dispose)
    this.stopModelSync.push(registerSheetFormatPainterCommands(this).dispose)
    for (const id of SHEET_BORDER_FORMAT_COMMAND_IDS) {
      this.registerCommand<SheetBorderFormatCommandInput | string | number | undefined, boolean>(id, params => {
        const input = sheetBorderFormatCommandInput(params)
        const unitId = input.unitId ?? this.activeWorkbookId
        const workbook = this.getWorkbookModel(unitId)
        if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
        const selected = this.getSelection()
        const sheetId = input.sheetId ?? input.subUnitId ?? (selected?.unitId === unitId ? selected.sheetId : undefined) ?? workbook.activeSheetId
        const ranges = sheetBorderFormatInputRanges(input) ?? this.commandSelectionRanges(unitId, sheetId)
        if (!ranges.length) return false
        this.assertCanEditWorksheet(unitId, sheetId)
        for (const range of ranges) this.assertCanEditRange(unitId, sheetId, range)
        const previousState = this.getSheetBorderFormatState(unitId, sheetId)
        const action = borderFormatActionForCommand(id, input, previousState)
        if (!action) return false
        if (action.apply) this.commitWorkbook(unitId, action.state.type === 'none' ? '清除边框' : '设置边框', model => {
          applySheetBorderFormatCommand(requireSheet(model, sheetId), ranges, action, this.currentUserId, model)
        }, { recalculate: false })
        let states = this.borderFormatStates.get(unitId)
        if (!states) { states = new Map(); this.borderFormatStates.set(unitId, states) }
        states.set(sheetId, { ...action.state })
        if (action.state.type !== previousState.type || action.state.style !== previousState.style || action.state.color !== previousState.color || action.state.active !== previousState.active) {
          this.events.emit('SheetBorderFormatStateChanged', { workbookId: unitId, sheetId, state: { ...action.state } })
        }
        // The fixed border command reports accepted redundant mutations; history remains unchanged.
        return action.apply
      })
    }
    for (const id of SHEET_LAYOUT_FORMAT_COMMAND_IDS) {
      this.registerCommand<SheetLayoutFormatCommandInput | undefined, boolean>(id, input => {
        const action = layoutFormatActionForCommand(id, input)
        if (!action) return false
        const unitId = input?.unitId ?? this.activeWorkbookId
        const workbook = this.getWorkbookModel(unitId)
        if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
        const selected = this.getSelection()
        const sheetId = input?.sheetId ?? input?.subUnitId ?? (selected?.unitId === unitId ? selected.sheetId : undefined) ?? workbook.activeSheetId
        const explicitRanges = sheetLayoutFormatInputRanges(input)
        const ranges = explicitRanges ?? this.commandSelectionRanges(unitId, sheetId)
        const selectedPrimary = !explicitRanges && selected?.unitId === unitId && selected.sheetId === sheetId ? this.getSelectionState()?.primary : undefined
        const primary = input?.primary !== undefined ? input.primary : selectedPrimary ? { row: selectedPrimary.row, col: selectedPrimary.column } : undefined
        if (!ranges.length) return false
        assertSheetLayoutFormatRanges(ranges)
        this.assertCanEditWorksheet(unitId, sheetId)
        for (const range of ranges) this.assertCanEditRange(unitId, sheetId, range)
        const label = action.kind === 'horizontal' ? '设置水平对齐' : action.kind === 'vertical' ? '设置垂直对齐'
          : action.kind === 'rotation' ? '设置文字方向' : action.kind === 'shrink' ? '设置缩小字体填充' : '设置文本换行'
        return this.commitWorkbook(unitId, label, model => {
          const sheet = requireSheet(model, sheetId)
          const filteredRows = [...(sheet.autoFilter?.filteredRows ?? []), ...(model.tables ?? []).filter(table => table.sheetId === sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])]
          applySheetLayoutFormatCommand(sheet, ranges, action, this.currentUserId, filteredRows, model, primary)
        }, { recalculate: false })
      })
    }
    for (const id of [...Object.values(SHEET_NUMBER_FORMAT_COMMANDS), ...Object.values(SHEET_NUMBER_FORMAT_ALIASES)]) {
      this.registerCommand<SheetNumberFormatCommandInput | undefined, boolean>(id, input => {
        const action = numberFormatActionForCommand(id, input, this.regionalSettings.region.value)
        if (!action) return false
        const unitId = input?.unitId ?? this.activeWorkbookId
        const workbook = this.getWorkbookModel(unitId)
        if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
        const selected = this.getSelection()
        const sheetId = input?.sheetId ?? input?.subUnitId ?? (selected?.unitId === unitId ? selected.sheetId : undefined) ?? workbook.activeSheetId
        const ranges = input?.ranges ?? this.commandSelectionRanges(unitId, sheetId)
        const label = action.kind === 'decimals' ? action.delta === 1 ? '增加小数位' : '减少小数位' : '设置数字格式'
        return this.commitWorkbook(unitId, label, model => {
          applySheetNumberFormatCommand(requireSheet(model, sheetId), ranges, action, this.regionalSettings.region.value, model.date1904, this.currentUserId, model)
        }, { recalculate: false })
      })
    }
    registerSheetOutlineCommands(this)
    for (const id of SHEET_FONT_FORMAT_COMMAND_IDS) {
      this.registerCommand<SheetFontFormatCommandParams, boolean>(id, params => {
        const input = sheetFontFormatCommandInput(params)
        const action = fontFormatActionForCommand(id, input)
        if (!action) return false
        const unitId = input?.unitId ?? this.activeWorkbookId
        const workbook = this.getWorkbookModel(unitId)
        if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
        const selected = this.getSelection()
        const sheetId = input?.sheetId ?? input?.subUnitId ?? (selected?.unitId === unitId ? selected.sheetId : undefined) ?? workbook.activeSheetId
        const explicitRanges = sheetFontFormatInputRanges(input)
        const ranges = explicitRanges ?? this.commandSelectionRanges(unitId, sheetId)
        const selectedPrimary = !explicitRanges && selected?.unitId === unitId && selected.sheetId === sheetId ? this.getSelectionState()?.primary : undefined
        const primary = input.primary ?? (selectedPrimary ? { row: selectedPrimary.row, col: selectedPrimary.column } : undefined)
        if (!ranges.length) return false
        assertSheetFontFormatRanges(ranges)
        this.assertCanEditWorksheet(unitId, sheetId)
        for (const range of ranges) this.assertCanEditRange(unitId, sheetId, range)
        const label = action.kind === 'mark' ? `设置${{ bold: '粗体', italic: '斜体', underline: '下划线', strike: '删除线' }[action.field]}`
          : action.kind === 'family' ? action.value === null ? '重置字体' : '设置字体'
            : action.kind === 'size' ? action.value === null ? '重置字号' : '设置字号'
              : action.kind === 'size-step' ? action.direction === 1 ? '增大字号' : '减小字号'
                : action.field === 'color' ? action.value === null ? '重置文字颜色' : '设置文字颜色' : action.value === null ? '清除填充颜色' : '设置填充颜色'
        return this.commitWorkbook(unitId, label, model => {
          const sheet = requireSheet(model, sheetId)
          const filteredRows = [...(sheet.autoFilter?.filteredRows ?? []), ...(model.tables ?? []).filter(table => table.sheetId === sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])]
          applySheetFontFormatCommand(sheet, ranges, action, this.currentUserId, filteredRows, primary, model)
        }, { recalculate: false })
      })
    }
    this.registerCommand<{ unitId: string; sheetId: string; row: number; col: number; draft: SheetHyperlinkDraft }, boolean>(SET_SHEET_HYPERLINK_COMMAND, ({ unitId, sheetId, row, col, draft }) => this.commitWorkbook(unitId, '设置单元格超链接', (model) => {
      this.assertCanEditWorksheet(unitId, sheetId)
      this.assertCanEditRange(unitId, sheetId, { startRow: row, endRow: row, startCol: col, endCol: col })
      setSheetCellHyperlink(model, requireSheet(model, sheetId), row, col, draft)
      const validation = validateSheetDataValidationChanges(model, sheetId, [{ row, column: col }], sheetValidationSourceReader(this.currentUserId), this.regionalSettings.region.value)
      if (validation.blocked) throw new Error(validation.issues[0]?.message ?? '超链接显示文字未通过数据验证。')
    }))
    this.registerCommand<{ unitId: string; sheetId: string; row: number; col: number }, boolean>(REMOVE_SHEET_HYPERLINK_COMMAND, ({ unitId, sheetId, row, col }) => this.commitWorkbook(unitId, '移除单元格超链接', (model) => {
      this.assertCanEditWorksheet(unitId, sheetId)
      this.assertCanEditRange(unitId, sheetId, { startRow: row, endRow: row, startCol: col, endCol: col })
      removeSheetCellHyperlink(requireSheet(model, sheetId), row, col)
    }))
    this.registerCommand<{ unitId: string; mode: 'automatic' | 'manual' }, boolean>('sheet.command.set-calculation-mode', ({ unitId, mode }) => {
      if (mode !== 'automatic' && mode !== 'manual') throw new TypeError('Calculation mode must be automatic or manual')
      return this.commitWorkbook(unitId, 'Set calculation mode', (model) => { model.calculationMode = mode })
    })
    this.registerCommand<{ unitId: string }, boolean>('sheet.command.recalculate-workbook', ({ unitId }) =>
      this.commitWorkbook(unitId, 'Recalculate workbook', () => {}, true),
    )
    this.registerCommand<ConfigureSheetFilterInput, boolean>(CONFIGURE_SHEET_FILTER_COMMAND, (input) => {
      const context = sheetFilterCommandContext(this, input, input.range)
      return this.commitWorkbook(input.unitId, '配置工作表筛选', (model) => configureFilterFromPanel(model, input),
        context)
    })
    this.registerCommand<ConfigurePivotFromPanelInput & PivotChartCommandContext & { hostSheetId?: string }, boolean>(CONFIGURE_PIVOT_COMMAND, raw => {
      const input = cloneModel(raw), workbook = this.getWorkbookModel(input.unitId)
      if (!workbook) throw new Error('透视报告工作簿已关闭。')
      const prepared = createDetachedPivotReport(workbook, input.source, input.anchor, input.pivotTableId)
      assertPivotReportEditPermission(workbook, prepared, this.currentUserId)
      return this.commitPivotChange(input.unitId, input.pivotTableId, input.anchor.subUnitId, CONFIGURE_PIVOT_COMMAND,
        input.mode === 'create' ? '插入数据透视表' : '配置数据透视表',
        draft => { configurePivotFromPanel(draft, input) }, input, input.hostSheetId ?? workbook.activeSheetId, prepared)
    })
    this.registerCommand<{ unitId: string; name: string }, boolean>('sheet.command.set-workbook-name', ({ unitId, name }) => this.commitWorkbook(unitId, 'Rename workbook', (model) => { model.name = name }))
    this.registerCommand<{ unitId: string; sheetId: string; range: { startRow: number; startColumn: number; endRow: number; endColumn: number }; values: Array<Array<unknown>> }, boolean>('sheet.command.set-range-values', ({ unitId, sheetId, range, values }) => this.commitWorkbook(unitId, 'Set range values', (model) => {
      this.assertCanEditWorksheet(unitId, sheetId)
      this.assertCanEditRange(unitId, sheetId, { startRow: range.startRow, startCol: range.startColumn, endRow: range.endRow, endCol: range.endColumn })
      const sheet = model.sheets.find((item) => item.id === sheetId)
      if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
      const positions: Array<{ row: number; column: number }> = []
      for (let row = range.startRow; row <= range.endRow; row++) for (let col = range.startColumn; col <= range.endColumn; col++) {
        const cell = values[row - range.startRow]?.[col - range.startColumn]
        if (cell === undefined) continue
        const key = `${row}:${col}`
        sheet.cells[key] = applyCellValue(sheet.cells[key], row, col, cell, model, sheet)
        positions.push({ row, column: col })
      }
      const validation = validateSheetDataValidationChanges(model, sheetId, positions, sheetValidationSourceReader(this.currentUserId), this.regionalSettings.region.value)
      if (validation.blocked) throw new Error(validation.issues[0]?.message ?? '输入值未通过数据验证，数据没有修改。')
    }))
    this.registerCommand<{ unitId: string; sheetId: string; range: CellRange | null }, boolean>('sheet.command.set-auto-filter', input => {
      const context = sheetFilterCommandContext(this, input, input.range)
      return this.commitWorkbook(input.unitId, 'Set auto-filter', model => {
        setAutoFilter(requireSheet(model, input.sheetId), input.range, model)
      }, context)
    })
    this.registerCommand<{ unitId: string; sheetId: string; range: CellRange; expectedRevision?: number }, boolean>('sheet.command.toggle-auto-filter', request => {
      const { unitId, sheetId, range } = request
      const removing = this.getWorkbookModel(unitId)?.sheets.find(sheet => sheet.id === sheetId)?.autoFilter !== undefined
      const context = sheetFilterCommandContext(this, request, removing ? null : range)
      return this.commitWorkbook(unitId, 'Toggle auto-filter', model => {
        const sheet = requireSheet(model, sheetId)
        setAutoFilter(sheet, sheet.autoFilter ? null : range, model)
      }, context)
    })
    this.registerCommand<{ unitId: string; sheetId: string; column: number; criteria: SheetFilterCriteria | null }, boolean>('sheet.command.set-filter-criteria', input => {
      const context = sheetFilterCommandContext(this, input)
      return this.commitWorkbook(input.unitId, 'Filter worksheet rows', model => {
        setFilterCriteria(requireSheet(model, input.sheetId), input.column, input.criteria, model)
      }, context)
    })
    this.registerCommand<{ unitId: string; sheetId: string; column?: number }, boolean>('sheet.command.clear-filter', input => {
      const context = sheetFilterCommandContext(this, input)
      return this.commitWorkbook(input.unitId, 'Clear worksheet filter', model => {
        clearFilter(requireSheet(model, input.sheetId), input.column, model)
      }, context)
    })
    this.registerCommand<{ unitId: string; sheetId: string }, boolean>('sheet.command.reapply-filter', input => {
      const context = sheetFilterCommandContext(this, input)
      return this.applyWorkbookOperation(input.unitId, model => {
        context.assertCurrent()
        refreshAutoFilter(requireSheet(model, input.sheetId), model)
        context.assertCurrent()
      })
    })
    this.registerCommand<{ unitId: string; sheetId: string; range: CellRange; criteria: SheetSortCriterion[]; options: SortRangeOptions; order?: number[]; expectedRevision?: number }, boolean>('sheet.command.sort-range', request => {
      const { unitId, sheetId, range, criteria, options, order, expectedRevision } = request
      const assertCurrent = workbookCommandGuard(request)
      assertCurrent?.()
      if (expectedRevision !== undefined && this.getWorkbookRevision(unitId) !== expectedRevision) throw new Error('Sort result is stale because the workbook changed during calculation')
      this.assertCanEditRange(unitId, sheetId, range)
      return this.commitWorkbook(unitId, 'Sort worksheet range', (model) => sortRange(requireSheet(model, sheetId), range, criteria, order, options, translateFormulaReferences, model),
        { expectedRevision, assertCurrent })
    })
    this.registerCommand<{ unitId: string; sheetId: string; rule: SheetRangeProtectionRule }, boolean>('sheet.command.set-range-protection', ({ unitId, sheetId, rule }) => this.commitWorkbook(unitId, 'Set range protection', (model) => {
      const sheet = requireSheet(model, sheetId)
      const existing = sheet.rangeProtectionRules?.find(item => item.id === rule.id)
      if (existing && existing.ownerUserId !== this.currentUserId) throw new Error('只有保护规则创建者可以修改该规则。')
      if (rule.ownerUserId !== this.currentUserId) throw new Error('保护规则创建者必须是当前用户。')
      setSheetRangeProtectionRule(sheet, rule)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; ruleId: string }, boolean>('sheet.command.remove-range-protection', ({ unitId, sheetId, ruleId }) => this.commitWorkbook(unitId, 'Remove range protection', (model) => {
      removeSheetRangeProtectionRule(requireSheet(model, sheetId), ruleId, this.currentUserId)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; rule: SheetWorksheetProtectionRule }, boolean>('sheet.command.add-worksheet-protection', ({ unitId, sheetId, rule }) => this.commitWorkbook(unitId, 'Protect worksheet', (model) => {
      const sheet = requireSheet(model, sheetId)
      const existing = sheet.worksheetProtection
      if (existing && existing.ownerUserId !== this.currentUserId) throw new Error('只有工作表保护创建者可以修改该保护。')
      if (rule.ownerUserId !== this.currentUserId) throw new Error('工作表保护创建者必须是当前用户。')
      if (rule.unitType !== 'worksheet') throw new TypeError('工作表保护范围必须是整个工作表。')
      setWorksheetProtectionRule(sheet, rule)
    }))
    this.registerCommand<{ unitId: string; sheetId: string }, boolean>('sheet.command.delete-worksheet-protection', ({ unitId, sheetId }) => this.commitWorkbook(unitId, 'Unprotect worksheet', (model) => {
      const sheet = requireSheet(model, sheetId)
      const rule = sheet.worksheetProtection
      if (!rule) return
      if (rule.ownerUserId !== this.currentUserId) throw new Error('只有工作表保护创建者可以移除该保护。')
      delete sheet.worksheetProtection
    }))
    registerSparklineCommands(this)
    registerPublicSparklineCommands(this)
    registerPivotChartCommands(this)
    registerPivotChartFieldCommands(this)
    registerPivotHierarchyCommands(this)
    registerPivotFilterCommands(this)
    this.registerCommand<{ unitId: string; sheetId: string; chartId: string; chart: ChartConfig }, Promise<boolean>>('sheet.command.add-chart', ({ unitId, sheetId, chartId, chart }) => this.commitChartChange(unitId, 'Insert chart', chartId, (model) => {
      this.assertCanEditWorksheet(unitId, sheetId)
      if (!model.sheets.some((item) => item.id === sheetId)) throw new Error(`Worksheet not found: ${sheetId}`)
      if (model.charts?.some((item) => item.id === chartId)) throw new Error(`Chart already exists: ${chartId}`)
      const created = createChartDefinition(model, sheetId, chart, chartId)
      model.charts ??= []
      model.charts.push(created)
    }, { initializeDefaults: true }), { asynchronous: true })
    this.registerCommand<{ unitId: string; chartId: string; patch: ChartUpdatePatch }, Promise<boolean>>('sheet.command.update-chart', ({ unitId, chartId, patch }) => {
      const authoring: ChartAuthoringPlan = { patch: cloneModel(patch) }
      return this.commitChartChange(unitId, 'Update chart', chartId, model => {
        const chart = requireChart(model, chartId)
        this.assertCanEditWorksheet(unitId, chart.sheetId)
        authoring.previous = cloneModel(chart)
        Object.assign(chart, updateChartDefinition(model, chart, authoring.patch!))
      }, authoring)
    }, { asynchronous: true })
    const removeChartObject = (input: ChartObjectCommandInput): boolean => {
      const target = resolveChartObjectTarget(this, input)
      if (!target) return false
      const workbook = this.getWorkbookModel(target.unitId)!, chart = requireChart(workbook, target.chartId)
      if (isPrimaryChartObject(workbook, chart)) return false
      assertChartObjectAction(this, target, 'delete')
      return this.commitWorkbook(target.unitId, 'Remove chart', model => {
        this.assertCanEditWorksheet(target.unitId, target.sheetId)
        model.charts = model.charts?.filter(item => item.id !== target.chartId)
        pruneChartImageResources(model)
      }, { expectedRevision: target.expectedRevision })
    }
    this.registerCommand<{ unitId: string; chartId: string }, boolean>('sheet.command.remove-chart', input => {
      const workbook = this.getWorkbookModel(input.unitId)
      const chart = workbook && requireChart(workbook, input.chartId)
      if (!chart) throw new Error('图表所属工作簿已经关闭。')
      return removeChartObject({ ...input, sheetId: chart.sheetId })
    })
    this.registerCommand<ChartObjectCommandInput, boolean>(CHART_OBJECT_COMMANDS.edit, input => {
      const target = resolveChartObjectTarget(this, input)
      if (!target) return false
      const chart = assertChartObjectAction(this, target, 'edit')
      const route = input?.route === undefined ? undefined : normalizeChartEditPanelRoute(input.route)
      if (route) {
        const data = this.chartStatisticsService.getCached(chart, { documentId: target.unitId, modelRevision: target.expectedRevision,
          userId: this.currentUserId }, this.getWorkbookModel(target.unitId))?.statistics.preparedData
        planChartEditPanelRoute(chart, route, data)
      }
      const host = chartObjectUIHost(this, target)
      if (!host) throw new Error('图表属性面板宿主尚未绑定。')
      return host.openPanel(target, route)
    })
    this.registerCommand<ChartObjectCommandInput, Promise<boolean>>(CHART_OBJECT_COMMANDS.copy, async input => {
      const target = resolveChartObjectTarget(this, input)
      if (!target) return false
      assertChartObjectAction(this, target, 'copy')
      const clipboard = chartObjectClipboardStore(this.pinia), ticket = clipboard.beginCopy(target.unitId)
      const captured = captureChartObjectClipboard(this.getWorkbookModel(target.unitId)!, target)
      const host = chartObjectUIHost(this, target)
      const written = host ? await host.writeClipboard(chartObjectClipboardContent(captured)) : false
      this.assertAlive()
      assertChartObjectAction(this, target, 'copy')
      if (host && (!host.current(target) || chartObjectUIHost(this, target) !== host)) throw new Error('图表宿主已切换，复制结果已停止。')
      return clipboard.publish(ticket, captured, written)
    }, { asynchronous: true })
    this.registerCommand<ChartObjectCommandInput, Promise<boolean>>(CHART_OBJECT_COMMANDS.exportImage, async input => {
      const target = resolveChartObjectTarget(this, input)
      if (!target) return false
      assertChartObjectAction(this, target, 'exportImage')
      const host = chartObjectUIHost(this, target)
      if (!host) throw new Error('图表图片导出需要已绘制的活动 Canvas 宿主。')
      return host.exportImage(target)
    }, { asynchronous: true })
    this.registerCommand<ChartObjectCommandInput, boolean>(CHART_OBJECT_COMMANDS.delete, removeChartObject)
    this.registerCommand<ChartObjectPasteInput, Promise<boolean>>(PASTE_COPIED_CHART_COMMAND, async (input = {}) => {
      if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('图表粘贴参数必须为对象。')
      input = cloneModel(input)
      const unitId = input.unitId ?? this.activeWorkbookId, workbook = this.getWorkbookModel(unitId)
      if (!workbook) throw new Error('粘贴目标工作簿已经关闭。')
      const sheetId = chartObjectSheetId(input) ?? workbook.activeSheetId
      const clipboardStore = chartObjectClipboardStore(this.pinia), clipboard = clipboardStore.clipboard, clipboardTicket = clipboardStore.request
      if (!clipboard) return false
      if (input.clipboardId !== undefined && input.clipboardId !== clipboard.id) throw new Error('图表剪贴板已更新，请重新粘贴。')
      const source = this.getWorkbookModel(clipboard.source.unitId)
      if (!source) throw new Error('复制来源工作簿已经关闭，请重新复制图表。')
      const userId = this.currentUserId, sourceRevision = this.getWorkbookRevision(source.id!), targetRevision = this.getWorkbookRevision(unitId)
      const sourceSavepoint = this.editor.history.captureSavepoint(source.id!), targetSavepoint = this.editor.history.captureSavepoint(unitId)
      let sourceImport: ChartSourceImportPlan | undefined
      const newChartId = input.newChartId ?? createChartObjectId()
      const host = input.hostSessionId ? chartObjectUISession(this, input.hostSessionId) : undefined
      const assertCurrent = (): void => {
        this.assertAlive()
        assertChartObjectRevision(this.getWorkbookRevision(unitId), input.expectedRevision)
        const currentSource = this.getWorkbookModel(source.id!), currentTarget = this.getWorkbookModel(unitId)
        const sourceHistory = this.editor.history.captureSavepoint(source.id!), targetHistory = this.editor.history.captureSavepoint(unitId)
        if (!currentSource || !currentTarget || this.currentUserId !== userId || currentSource !== source || currentTarget !== workbook
          || this.getWorkbookRevision(source.id!) !== sourceRevision || this.getWorkbookRevision(unitId) !== targetRevision
          || sourceHistory.revision !== sourceSavepoint.revision || sourceHistory.generation !== sourceSavepoint.generation
          || targetHistory.revision !== targetSavepoint.revision || targetHistory.generation !== targetSavepoint.generation) {
          throw new ComputeError('COMPUTE_STALE_RESULT', '图表粘贴期间来源、目标、权限用户或内容历史已更改')
        }
        if (input.unitId === undefined && this.activeWorkbookId !== unitId
          || input.sheetId === undefined && input.subUnitId === undefined && currentTarget.activeSheetId !== sheetId) throw new Error('图表粘贴的默认目标已经切换。')
        if (input.hostSessionId && (!host || chartObjectUISession(this, input.hostSessionId) !== host || !host.currentSheet(unitId, sheetId))) throw new Error('图表粘贴宿主已经切换。')
        if (clipboardStore.clipboard?.id !== clipboard.id || clipboardStore.request !== clipboardTicket) throw new Error('图表剪贴板已更新，请重新粘贴。')
        this.assertCanEditWorksheet(unitId, sheetId)
        assertChartObjectClipboardSource(currentSource, clipboard)
        assertChartReadPermission(currentSource, clipboard.chart, userId)
        if (sourceImport) assertChartSourceImportSource(currentSource, sourceImport, userId)
      }
      assertCurrent()
      if (source.id !== unitId) sourceImport = await planChartSourceImportAsync(source, workbook, clipboard, { userId, targetSheetId: sheetId, assertCurrent })
      assertCurrent()
      return this.commitChartChange(unitId, 'Paste chart', newChartId, model => {
        const created = chartObjectPasteDefinition(model, clipboard, { ...input, unitId, sheetId, newChartId }, { sourceWorkbook: source, sourceImport, userId })
        assertChartReadPermission(model, created, this.currentUserId)
        model.charts ??= []
        model.charts.push(created)
      }, {}, assertCurrent, sourceImport ? { expectedRevision: targetRevision, expectedHistoryGeneration: targetSavepoint.generation } : false)
    }, { asynchronous: true })
    this.registerCommand<{ unitId: string; pivotTableId: string; source: PivotTableSourceInput; anchor: PivotTableAnchorInput }, boolean>('sheet.command.add-pivot-table', ({ unitId, pivotTableId, source, anchor }) => this.commitPivotChange(unitId, pivotTableId, anchor.subUnitId, 'sheet.command.add-pivot-table', 'Add pivot table', (model) => {
      if (model.pivotTables?.some((pivot) => pivot.id === pivotTableId)) throw new Error(`Pivot table already exists: ${pivotTableId}`)
      createPivotTable(model, source, anchor, pivotTableId)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string } & PivotReportCommandContext, boolean>('sheet.command.refresh-pivot-table', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.refresh-pivot-table', 'Refresh pivot table', model => {
        refreshPivotTable(model, requirePivot(model, input.pivotTableId))
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string | number; area: PivotFieldArea; index?: number } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-field', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-field', 'Set pivot field', model => {
        setPivotField(model, requirePivot(model, input.pivotTableId), input.fieldId, input.area, input.index)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldIds: string[] } & PivotReportCommandContext, boolean>('sheet.command.remove-pivot-field', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.remove-pivot-field', 'Remove pivot field', model => {
        removePivotFields(model, requirePivot(model, input.pivotTableId), input.fieldIds)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; options: Partial<PivotTableOptions> } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-options', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-options', 'Set pivot options', model => {
        setPivotOptions(model, requirePivot(model, input.pivotTableId), input.options)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; direction: 'ascending' | 'descending' | null } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-label-sort', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-label-sort', 'Sort pivot labels', model => {
        setPivotSort(model, requirePivot(model, input.pivotTableId), input.fieldId, input.direction)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; selectedItems: string[] | null } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-label-filter', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-label-filter', 'Filter pivot labels', model => {
        setPivotLabelFilter(model, requirePivot(model, input.pivotTableId), input.fieldId, input.selectedItems)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; setting: { displayName?: string; format?: string; subtotalType?: PivotAggregation; showDataAs?: PivotShowDataAsInfo | null } } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-field-setting', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-field-setting', 'Set pivot field setting', model => {
        setPivotFieldSetting(model, requirePivot(model, input.pivotTableId), input.fieldId, input.setting)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; showDataAs: PivotShowDataAsInfo | null } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-show-data-as', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-show-data-as', '设置透视值显示方式', model => {
        setPivotShowDataAs(model, requirePivot(model, input.pivotTableId), input.fieldId, input.showDataAs)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; filter: PivotFieldFilter | null } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-value-filter', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-value-filter', 'Filter pivot values', model => {
        setPivotValueFilter(model, requirePivot(model, input.pivotTableId), input.fieldId, input.filter)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; settings: PivotNumberGroup | null } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-number-group', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-number-group', 'Group pivot numbers', model => {
        setPivotNumberGroup(model, requirePivot(model, input.pivotTableId), input.fieldId, input.settings)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; mode: PivotDateGroupingMode } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-date-grouping-mode', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-date-grouping-mode', 'Set pivot date grouping mode', model => {
        setPivotDateGroupingMode(model, requirePivot(model, input.pivotTableId), input.mode)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; type: PivotDateGroupType | null } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-date-group-type', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-date-group-type', 'Group pivot dates', model => {
        setPivotDateGroupType(model, requirePivot(model, input.pivotTableId), input.fieldId, input.type)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; type: PivotDateGroupType } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-excel-date-group', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-excel-date-group', 'Group pivot dates (Excel)', model => {
        setPivotExcelDateGroup(model, requirePivot(model, input.pivotTableId), input.fieldId, input.type)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; groups: PivotManualGroup[] } & PivotReportCommandContext, boolean>('sheet.command.set-pivot-manual-groups', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.set-pivot-manual-groups', 'Group pivot members', model => {
        setPivotManualGroups(model, requirePivot(model, input.pivotTableId), input.fieldId, input.groups)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string; position: import('../domain/model.ts').PivotValuePosition | -1 | 1 | 2;
      index?: number } & PivotReportCommandContext, boolean>('sheet.command.update-pivot-value-position', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.update-pivot-value-position', 'Update pivot value position', model => {
        setPivotValuesPosition(model, requirePivot(model, input.pivotTableId), normalizePivotValuePositionInput(input.position), input.index === undefined ? -1 : input.index)
      }, input, input.hostSheetId)
    })
    this.registerCommand<{ unitId: string; pivotTableId: string } & PivotReportCommandContext, boolean>('sheet.command.remove-pivot-table', raw => {
      const input = cloneModel(raw)
      return this.commitPivotChange(input.unitId, input.pivotTableId, undefined, 'sheet.command.remove-pivot-table', 'Remove pivot table', model => {
        const pivot = requirePivot(model, input.pivotTableId), target = model.sheets.find(sheet => sheet.id === pivot.anchor.subUnitId)
        if (target) for (const key of pivot.generatedCellKeys ?? []) if (target.cells[key]?.metadata?.pivotTableId === pivot.id) delete target.cells[key]
        model.pivotTables = (model.pivotTables ?? []).filter(candidate => candidate.id !== input.pivotTableId)
      }, input, input.hostSheetId)
    })
    for (const [commandId, kind] of [[SLICER_PUBLIC_COMMANDS.insert, 'slicer'], [SLICER_PUBLIC_COMMANDS.insertTimeline, 'timeline']] as const) {
      this.registerCommand<{ unitId: string; hostSheetId: string; slicerId: string; info: ISheetSlicerInfo | ISheetTimelineInfo } & SlicerCommandContext, boolean>(commandId, ({ unitId, hostSheetId, slicerId, info, expectedRevision, host }) => {
        const detached = cloneModel(info)
        return this.commitSlicerChange(unitId, slicerId, hostSheetId, commandId, kind === 'timeline' ? '插入时间线' : '插入切片器',
          model => { createSheetSlicerFromInfo(model, hostSheetId, slicerId, kind, detached) }, expectedRevision, host, true)
      })
    }
    for (const [commandId, kind] of [[SLICER_PUBLIC_COMMANDS.update, 'slicer'], [SLICER_PUBLIC_COMMANDS.updateTimeline, 'timeline']] as const) {
      this.registerCommand<{ unitId: string; slicerId: string; info: ISheetSlicerInfo | ISheetTimelineInfo } & SlicerCommandContext, boolean>(commandId, ({ unitId, slicerId, info, expectedRevision, host }) => {
        const detached = cloneModel(info)
        return this.commitSlicerChange(unitId, slicerId, undefined, commandId, kind === 'timeline' ? '更新时间线信息' : '更新切片器信息', model => {
          const view = requireSlicer(model, slicerId)
          if (view.type !== kind) throw new SlicerConnectionError('invalid-target', '切片器和时间线使用各自的信息命令。')
          updateSheetSlicerFromInfo(model, view, detached)
        }, expectedRevision, host, true)
      })
    }
    this.registerCommand<{ unitId: string; slicerId: string; patch: TimelineSettingsInputPatch } & SlicerCommandContext, boolean>(SLICER_PUBLIC_COMMANDS.timelineSettings, ({ unitId, slicerId, patch, expectedRevision, host }) => {
      const detached = cloneModel(patch)
      return this.commitSlicerChange(unitId, slicerId, undefined, SLICER_PUBLIC_COMMANDS.timelineSettings, '更新时间线显示设置',
        model => { patchTimelineSettings(model, requireSlicer(model, slicerId), detached) }, expectedRevision, host, true)
    })
    this.registerCommand<{ unitId: string; slicerId: string; range: ITimelineRange | undefined } & SlicerCommandContext, boolean>(SLICER_PUBLIC_COMMANDS.dateRange, ({ unitId, slicerId, range, expectedRevision, host }) => {
      let normalized: ITimelineRange | undefined
      try { normalized = range === undefined ? undefined : normalizeTimelineDateRange(range) }
      catch (error) { if (error instanceof TypeError || error instanceof RangeError) return false; throw error }
      return this.commitSlicerChange(unitId, slicerId, undefined, SLICER_PUBLIC_COMMANDS.dateRange, '筛选时间线日期范围',
        model => { setTimelineDateRange(model, requireSlicer(model, slicerId), normalized) }, expectedRevision, host, true)
    })
    this.registerCommand<{ unitId: string; slicerId: string; condition: ITimelineDateCondition | undefined } & SlicerCommandContext, boolean>(SLICER_PUBLIC_COMMANDS.dateCondition, ({ unitId, slicerId, condition, expectedRevision, host }) => {
      let normalized: ITimelineDateCondition | undefined
      try { normalized = condition === undefined ? undefined : normalizeTimelineDateCondition(condition) }
      catch (error) { if (error instanceof TypeError || error instanceof RangeError) return false; throw error }
      return this.commitSlicerChange(unitId, slicerId, undefined, SLICER_PUBLIC_COMMANDS.dateCondition, '筛选时间线日期条件',
        model => { setTimelineDateCondition(model, requireSlicer(model, slicerId), normalized) }, expectedRevision, host, true)
    })
    this.registerCommand<{ unitId: string; slicerId: string; placement: SheetSlicerPlacement } & SlicerCommandContext, boolean>(SLICER_PUBLIC_COMMANDS.placement, ({ unitId, slicerId, placement, expectedRevision, host }) => {
      const detached = cloneModel(placement)
      return this.commitSlicerChange(unitId, slicerId, undefined, SLICER_PUBLIC_COMMANDS.placement, '更新切片器绘图位置', model => {
        const view = requireSlicer(model, slicerId)
        updateSlicerPresentation(view, sheetSlicerPlacementPatch(model, view, detached), model)
      }, expectedRevision, host, true)
    })
    this.registerCommand<{ unitId: string; slicer: SlicerCreateInput } & SlicerCommandContext, boolean>('sheet.command.add-slicer', ({ unitId, slicer, expectedRevision, host }) => this.commitSlicerChange(unitId, slicer.id, slicer.hostSheetId, 'sheet.command.add-slicer', 'Add slicer', (model) => {
      if (slicer.unitId !== unitId) throw new Error('Slicer workbook does not match the command workbook')
      createSlicer(model, slicer)
    }, expectedRevision, host))
    this.registerCommand<{ unitId: string; slicerId: string; connections: SheetSlicerConnections } & SlicerCommandContext, boolean>('sheet.command.set-slicer-connections', ({ unitId, slicerId, connections, expectedRevision, host }) => {
      const model = this.getWorkbookModel(unitId), slicer = model?.slicers?.find(control => control.id === slicerId)
      if (!model || !slicer) return false
      try {
        const plan = planSlicerConnections(model, slicer, connections)
        assertSlicerFilterPermission(model, slicer, this.currentUserId, [...plan.reports, ...plan.removedReports])
      } catch (error) { if (error instanceof SlicerConnectionError) return false; throw error }
      return this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.set-slicer-connections', 'Update slicer report connections',
        draft => { setSlicerConnections(draft, requireSlicer(draft, slicerId), connections) }, expectedRevision, host, true)
    })
    this.registerCommand<{ unitId: string; slicerId: string; selectedItems: string[] | null } & SlicerCommandContext, boolean>('sheet.command.set-slicer-items', ({ unitId, slicerId, selectedItems, expectedRevision, host }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.set-slicer-items', 'Filter with slicer', (model) => {
      setSlicerItems(model, requireSlicer(model, slicerId), selectedItems)
    }, expectedRevision, host))
    this.registerCommand<{ unitId: string; slicerId: string; checkedItems: string[]; partialItems: string[] } & SlicerCommandContext, boolean>('sheet.command.set-slicer-item-states', ({ unitId, slicerId, checkedItems, partialItems, expectedRevision, host }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.set-slicer-item-states', 'Filter slicer members', (model) => {
      setSlicerItemStates(model, requireSlicer(model, slicerId), checkedItems, partialItems)
    }, expectedRevision, host, true))
    this.registerCommand<{ unitId: string; slicerId: string; selection: PivotRecordSelection | null } & SlicerCommandContext, boolean>('sheet.command.set-slicer-record-selection', ({ unitId, slicerId, selection, expectedRevision, host }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.set-slicer-record-selection', 'Filter typed slicer members', (model) => {
      setSlicerRecordSelection(model, requireSlicer(model, slicerId), selection)
    }, expectedRevision, host, true))
    this.registerCommand<{ unitId: string; slicerId: string; range: SlicerTimelineRange | null } & SlicerCommandContext, boolean>('sheet.command.set-timeline-range', ({ unitId, slicerId, range, expectedRevision, host }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.set-timeline-range', 'Filter with timeline', (model) => {
      setSlicerTimelineRange(model, requireSlicer(model, slicerId), range)
    }, expectedRevision, host))
    this.registerCommand<{ unitId: string; slicerId: string; patch: SlicerPresentationInputPatch } & SlicerCommandContext, boolean>('sheet.command.update-slicer-presentation', ({ unitId, slicerId, patch, expectedRevision, host }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.update-slicer-presentation', 'Update slicer presentation', (model) => {
      updateSlicerPresentation(requireSlicer(model, slicerId), patch, model)
    }, expectedRevision, host))
    this.registerCommand<{ unitId: string; slicerId: string; patch: Partial<SheetSlicerSettings> } & SlicerCommandContext, boolean>('sheet.command.update-slicer-settings', ({ unitId, slicerId, patch, expectedRevision, host }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.update-slicer-settings', 'Update slicer view settings', model => {
      patchSlicerSettings(model, requireSlicer(model, slicerId), patch)
    }, expectedRevision, host, true))
    this.registerCommand<{ unitId: string; slicerId: string; patch: Partial<import('../domain/model.ts').SheetSlicerCacheOptions> } & SlicerCommandContext, boolean>('sheet.command.patch-slicer-cache-options', ({ unitId, slicerId, patch, expectedRevision, host }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.patch-slicer-cache-options', 'Update shared slicer cache options', model => {
      patchSlicerCacheOptions(model, requireSlicer(model, slicerId), patch)
    }, expectedRevision, host, true))
    this.registerCommand<{ unitId: string; slicerId: string } & SlicerCommandContext, boolean>('sheet.command.clear-slicer-filter', ({ unitId, slicerId, expectedRevision, host }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.clear-slicer-filter', 'Clear slicer filter', (model) => {
      clearSlicerFilter(model, requireSlicer(model, slicerId))
    }, expectedRevision, host))
    this.registerCommand<{ unitId: string; slicerId: string } & SlicerCommandContext, boolean>('sheet.command.remove-slicer', ({ unitId, slicerId, expectedRevision, host }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.remove-slicer', 'Remove slicer', (model) => removeSlicer(model, slicerId), expectedRevision, host))
    this.registerCommand<{ unitId: string; name: string }, boolean>('doc.command.set-name', ({ unitId, name }) => this.commitDocument(unitId, 'Rename document', (model) => { model.name = name }))
    this.registerCommand<{ unitId: string; rules: DocumentModel['permissionRules'] }, boolean>('doc.command.set-document-permission-rules', ({ unitId, rules }) => this.commitDocumentPermissionRules(unitId, rules))
    this.registerCommand<{ unitId: string; index: number; text: string; expectedRevision?: number } | DocumentEditorCommandRequest, boolean | DocumentEditorCommandResult>('doc.command.insert-text', input => {
      if (!input || typeof input !== 'object') throw new TypeError('文档输入参数无效。')
      if ('command' in input) return this.commitDocumentEditorCommand('doc.command.insert-text', input)
      const { unitId, index, text, expectedRevision } = input
      return this.commitIndexedDocumentInput(unitId, index, index, text, expectedRevision)
    })
    this.registerCommand<{ value: string } | (DocumentEditorCommandRequest & { value: string }), boolean | DocumentEditorCommandResult>('doc.command.insert-special-character', input => {
      if (!input || typeof input !== 'object' || typeof input.value !== 'string') throw new TypeError('特殊字符插入参数无效。')
      if ('command' in input) {
        const request = { ...input, command: { ...input.command, text: input.value } }
        if (!input.value.length) {
          assertDocumentEditorCommandRequest(request, 'doc.command.insert-text')
          return { changed: false, selection: cloneDocumentSelection(request.selection) }
        }
        return this.commitDocumentEditorCommand('doc.command.insert-text', request)
      }
      if (!input.value.length || this.editor.state.activeKind !== 'document') return false
      const unitId = this._activeDocumentId, session = this.documentInput.get(unitId)
      if (!session?.valid || session.location.kind !== 'body' || session.modelRevision !== this.getDocumentRevision(unitId)) return false
      return this.commitDocumentEditorCommand('doc.command.insert-text', {
        unitId, selection: session.location.selection,
        command: { id: 'doc.command.insert-text', text: input.value,
          ...(session.location.host ? { host: session.location.host } : {}) },
        expectedRevision: session.modelRevision,
        expectedHistoryGeneration: this.editor.history.captureSavepoint(unitId).generation,
        inputGeneration: session.generation,
      }).changed
    })
    this.registerCommand<{ rowCount: number; colCount: number } | DocumentEditorCommandRequest, boolean | DocumentEditorCommandResult>('doc.command.create-table', input => {
      if (!input || typeof input !== 'object') throw new TypeError('表格插入参数无效。')
      if ('command' in input) return this.commitDocumentEditorCommand('doc.command.create-table', input)
      if (this.editor.state.activeKind !== 'document') return false
      const unitId = this._activeDocumentId, session = this.documentInput.get(unitId)
      if (!session?.valid || session.location.kind !== 'body' || session.modelRevision !== this.getDocumentRevision(unitId)) return false
      return this.commitDocumentEditorCommand('doc.command.create-table', {
        unitId, selection: session.location.selection,
        command: { id: 'doc.command.create-table', rowCount: input.rowCount, colCount: input.colCount,
          ...(session.location.host ? { host: session.location.host } : {}) },
        expectedRevision: session.modelRevision,
        expectedHistoryGeneration: this.editor.history.captureSavepoint(unitId).generation,
        inputGeneration: session.generation,
      }).changed
    })
    for (const commandId of Object.keys(DOCUMENT_EDITOR_COMMAND_LABELS)) {
      if (commandId === 'doc.command.insert-text' || commandId === 'doc.command.create-table'
        || commandId === DOCUMENT_PAGE_SETUP_COMMAND || commandId === DOCUMENT_TOC_INSERT_COMMAND
        || commandId === DOCUMENT_TOC_UPDATE_COMMAND || commandId === DOCUMENT_TOC_DELETE_COMMAND || commandId === SET_DOCUMENT_HEADING_COLLAPSED_COMMAND) continue
      this.registerCommand<DocumentEditorCommandRequest, DocumentEditorCommandResult>(commandId, request => this.commitDocumentEditorCommand(commandId, request))
    }
    registerDocumentPageSetupCommand(this, request => this.commitDocumentEditorCommand(DOCUMENT_PAGE_SETUP_COMMAND, request))
    registerDocumentTocInsertCommand(this, request => this.commitDocumentEditorCommand(DOCUMENT_TOC_INSERT_COMMAND, request))
    registerDocumentTocActionCommands(this, (request, amendment) => this.commitDocumentEditorCommand(request.command.id, request, amendment))
    registerDocumentTocDialogOperation(this)
    registerDocumentHeadingCollapseCommand(this, request => this.commitDocumentEditorCommand(SET_DOCUMENT_HEADING_COLLAPSED_COMMAND, request))
    registerDocumentZoomCommands(this, params => this.publishDocumentZoomRatio(params))
    registerDocumentHeaderFooterCommands(this)
    this.registerCommand<{ unitId: string; startOffset: number; endOffset: number; text?: string; expectedRevision?: number }, boolean>('doc.command.replace-range', ({ unitId, startOffset, endOffset, text = '', expectedRevision }) => this.commitIndexedDocumentInput(unitId, startOffset, endOffset, text, expectedRevision))
    registerDocumentContentControlCommands(this)
    registerDocumentCustomXmlCommands(this)
    registerDocumentLatexCommands(this)
    registerDocumentImageCommands(this)
    registerDocumentNoteCommands(this)
  }

  commitPivotChange(
    unitId: string,
    pivotTableId: string,
    subUnitId: string | undefined,
    commandId: string,
    label: string,
    update: (model: WorkbookModel) => void,
    context?: PivotChartCommandContext,
    contextHostSheetId?: string,
    creationReport?: PivotTableModel,
  ): boolean {
    const current = this.getWorkbookModel(unitId)
    if (!current) throw new Error(`Workbook not found: ${unitId}`)
    const currentPivot = current.pivotTables?.find((pivot) => pivot.id === pivotTableId)
    const revision = this.getWorkbookRevision(unitId), savepoint = this.editor.history.captureSavepoint(unitId),
      signature = JSON.stringify(current), userId = this.currentUserId
    if (context) {
      if (context.expectedRevision !== undefined && (!Number.isSafeInteger(context.expectedRevision) || context.expectedRevision < 0)
        || context.expectedHistoryGeneration !== undefined && (!Number.isSafeInteger(context.expectedHistoryGeneration) || context.expectedHistoryGeneration <= 0)
        || context.expectedUserId !== undefined && (typeof context.expectedUserId !== 'string' || !context.expectedUserId.trim())
        || context.host !== undefined && context.host !== 'workbench') throw new TypeError('透视报告交互上下文无效。')
      if (contextHostSheetId !== undefined && (typeof contextHostSheetId !== 'string' || !current.sheets.some(sheet => sheet.id === contextHostSheetId))) throw new TypeError('透视报告工作台宿主无效。')
      const permissionReport = currentPivot ?? creationReport
      if (!permissionReport) throw new Error('透视报告不存在。')
      assertPivotReportEditPermission(current, permissionReport, userId)
    }
    const assertCurrent = () => {
      this.assertAlive()
      if (this.getWorkbookModel(unitId) !== current || this.getWorkbookRevision(unitId) !== revision || JSON.stringify(current) !== signature
        || this.editor.history.captureSavepoint(unitId).generation !== savepoint.generation || this.currentUserId !== userId
        || context?.expectedRevision !== undefined && context.expectedRevision !== revision
        || context?.expectedHistoryGeneration !== undefined && context.expectedHistoryGeneration !== savepoint.generation
        || context?.expectedUserId !== undefined && context.expectedUserId !== userId) {
        throw new Error('透视报告目标内容、用户或历史已更改。')
      }
      if (context?.host === 'workbench' && (this._activeWorkbookId !== unitId || this.editor.state.activeKind !== 'workbook'
        || current.activeSheetId !== (contextHostSheetId ?? currentPivot?.anchor.subUnitId ?? creationReport?.anchor.subUnitId))) throw new Error('透视报告工作台宿主已更改。')
    }
    assertCurrent()
    const event: ApiEventMap['BeforePivotTableChange'] = {
      unitId,
      subUnitId: subUnitId ?? currentPivot?.anchor.subUnitId ?? '',
      pivotTableId,
      commandId,
    }
    this.events.emit('BeforePivotTableChange', event)
    if (event.cancel) throw new Error(`Pivot table change was canceled: ${commandId}`)
    assertCurrent()
    let pivots: ApiEventMap['PivotTableRendered'][] = [], controls: ApiEventMap['SlicerChanged'][] = [], charts: ApiEventMap['PivotChartChanged'][] = []
    const changed = this.commitWorkbook(unitId, label, draft => { assertCurrent(); update(draft) }, {
      expectedRevision: revision, expectedHistoryGeneration: savepoint.generation,
      prepareView: candidate => {
        assertCurrent()
        if (context) {
          const report = candidate.pivotTables?.find(pivot => pivot.id === pivotTableId)
          if (report) assertPivotReportEditPermission(candidate, report, userId)
        }
        for (const view of candidate.slicers ?? []) {
          const previous = current.slicers?.find(control => control.id === view.id)
          if (!previous || equivalentModelData(previous, view)) continue
          assertSlicerFilterPermission(candidate, view, userId)
        }
        controls = (candidate.slicers ?? []).filter(view => !equivalentModelData(current.slicers?.find(previous => previous.id === view.id), view))
          .map(view => ({ unitId, hostSheetId: view.hostSheetId, slicerId: view.id, revision: view.revision }))
        pivots = (candidate.pivotTables ?? []).filter(pivot => !equivalentModelData(current.pivotTables?.find(previous => previous.id === pivot.id), pivot))
          .map(pivot => ({ unitId, subUnitId: pivot.anchor.subUnitId, pivotTableId: pivot.id, revision: pivot.revision,
            ...(pivot.outputRange ? { outputRange: { ...pivot.outputRange } } : {}) }))
        charts = (candidate.charts ?? []).filter(chart => chart.pivotTableId && !chart.ownedPivot
          && !equivalentModelData(current.charts?.find(previous => previous.id === chart.id), chart))
          .flatMap(chart => {
            const report = candidate.pivotTables?.find(pivot => pivot.id === chart.pivotTableId)
            return report ? [{ unitId, subUnitId: chart.sheetId, pivotChartId: chart.id, revision: chart.revision, reportRevision: report.revision }] : []
          })
      },
    })
    if (!changed) return false
    for (const payload of controls) this.publishObserverEvent('SlicerChanged', () => ({ ...payload }))
    for (const payload of pivots) {
      this.publishObserverEvent('PivotTableChanged', () => ({ unitId, subUnitId: payload.subUnitId, pivotTableId: payload.pivotTableId, revision: payload.revision }))
      this.publishObserverEvent('PivotTableRendered', () => cloneModel(payload))
    }
    for (const payload of charts) this.publishObserverEvent('PivotChartChanged', () => ({ ...payload }))
    return true
  }

  private commitSlicerChange(
    unitId: string,
    slicerId: string,
    hostSheetId: string | undefined,
    commandId: string,
    label: string,
    update: (model: WorkbookModel) => void,
    expectedRevision?: number,
    host?: 'workbench',
    ordinaryRefusals = false,
  ): boolean {
    if (expectedRevision !== undefined && (!Number.isSafeInteger(expectedRevision) || expectedRevision < 0)) throw new TypeError('切片器内容版本无效。')
    if (host !== undefined && host !== 'workbench') throw new TypeError('切片器交互宿主无效。')
    const current = this.getWorkbookModel(unitId)
    if (!current) { if (ordinaryRefusals) return false; throw new Error(`Workbook not found: ${unitId}`) }
    const currentSlicer = current.slicers?.find((slicer) => slicer.id === slicerId)
    if (!currentSlicer && !hostSheetId && ordinaryRefusals) return false
    const changesFilter = ['sheet.command.add-slicer', SLICER_PUBLIC_COMMANDS.insert, SLICER_PUBLIC_COMMANDS.insertTimeline,
      SLICER_PUBLIC_COMMANDS.dateRange, SLICER_PUBLIC_COMMANDS.dateCondition,
      'sheet.command.set-slicer-items', 'sheet.command.set-slicer-item-states', 'sheet.command.set-slicer-record-selection',
      'sheet.command.set-timeline-range', 'sheet.command.clear-slicer-filter',
      'sheet.command.set-slicer-connections'].includes(commandId)
    if (changesFilter && currentSlicer) {
      try { assertSlicerFilterPermission(current, currentSlicer, this.currentUserId) }
      catch (error) { if (ordinaryRefusals && error instanceof SlicerConnectionError) return false; throw error }
    }
    const revision = this.getWorkbookRevision(unitId), savepoint = this.editor.history.captureSavepoint(unitId),
      signature = JSON.stringify(current), userId = this.currentUserId
    const assertCurrent = () => {
      this.assertAlive()
      if (host === 'workbench' && (this._activeWorkbookId !== unitId || this.editor.state.activeKind !== 'workbook'
        || current.activeSheetId !== (hostSheetId ?? currentSlicer?.hostSheetId))) throw new Error('切片器工作台宿主已更改。')
      if (this.getWorkbookModel(unitId) !== current || this.getWorkbookRevision(unitId) !== revision || JSON.stringify(current) !== signature
        || this.editor.history.captureSavepoint(unitId).generation !== savepoint.generation || this.currentUserId !== userId
        || expectedRevision !== undefined && expectedRevision !== revision) throw new Error('切片器目标内容或历史已更改。')
    }
    assertCurrent()
    const event: ApiEventMap['BeforeSlicerChange'] = {
      unitId,
      hostSheetId: hostSheetId ?? currentSlicer?.hostSheetId ?? '',
      slicerId,
      commandId,
    }
    this.events.emit('BeforeSlicerChange', event)
    if (event.cancel) throw new Error(`Slicer change was canceled: ${commandId}`)
    assertCurrent()
    let added: ApiEventMap['SlicerAdded'][] = [], removed: ApiEventMap['SlicerRemoved'][] = [],
      updated: ApiEventMap['SlicerChanged'][] = [], pivots: ApiEventMap['PivotTableRendered'][] = [], charts: ApiEventMap['PivotChartChanged'][] = []
    let changed: boolean
    try { changed = this.commitWorkbook(unitId, label, draft => { assertCurrent(); update(draft) }, {
      expectedRevision: revision, expectedHistoryGeneration: savepoint.generation,
      prepareView: candidate => {
        assertCurrent()
        const beforeControls = new Map((current.slicers ?? []).map(control => [control.id, control])),
          afterControls = new Map((candidate.slicers ?? []).map(control => [control.id, control]))
        for (const id of new Set([...beforeControls.keys(), ...afterControls.keys()])) {
          if (equivalentModelData(beforeControls.get(id), afterControls.get(id))) continue
          for (const [model, control] of [[current, beforeControls.get(id)], [candidate, afterControls.get(id)]] as const) {
            if (!control) continue
            assertSlicerViewPermission(model, control, userId, true)
          }
        }
        const control = candidate.slicers?.find(slicer => slicer.id === slicerId)
        if (changesFilter && control) assertSlicerFilterPermission(candidate, control, userId)
        added = (candidate.slicers ?? []).filter(view => !beforeControls.has(view.id)).map(view => ({ unitId, hostSheetId: view.hostSheetId, slicerId: view.id, type: view.type }))
        removed = (current.slicers ?? []).filter(view => !afterControls.has(view.id)).map(view => ({ unitId, hostSheetId: view.hostSheetId, slicerId: view.id }))
        updated = (candidate.slicers ?? []).filter(view => beforeControls.has(view.id) && !equivalentModelData(beforeControls.get(view.id), view))
          .map(view => ({ unitId, hostSheetId: view.hostSheetId, slicerId: view.id, revision: view.revision }))
        pivots = (candidate.pivotTables ?? []).filter(pivot => !equivalentModelData(current.pivotTables?.find(before => before.id === pivot.id), pivot))
          .map(pivot => ({ unitId, subUnitId: pivot.anchor.subUnitId, pivotTableId: pivot.id, revision: pivot.revision,
            ...(pivot.outputRange ? { outputRange: { ...pivot.outputRange } } : {}) }))
        charts = (candidate.charts ?? []).filter(chart => (chart.ownedPivot || chart.pivotTableId)
          && !equivalentModelData(current.charts?.find(before => before.id === chart.id), chart))
          .map(chart => ({ unitId, subUnitId: chart.sheetId, pivotChartId: chart.id, revision: chart.revision,
            reportRevision: chart.ownedPivot?.revision ?? candidate.pivotTables?.find(pivot => pivot.id === chart.pivotTableId)?.revision ?? 0 }))
      },
    }) } catch (error) { if (ordinaryRefusals && error instanceof SlicerConnectionError) return false; throw error }
    if (!changed) return ordinaryRefusals
    for (const payload of added) this.publishObserverEvent('SlicerAdded', () => ({ ...payload }))
    for (const payload of removed) this.publishObserverEvent('SlicerRemoved', () => ({ ...payload }))
    for (const payload of updated) this.publishObserverEvent('SlicerChanged', () => ({ ...payload }))
    for (const payload of pivots) {
      this.publishObserverEvent('PivotTableChanged', () => ({ unitId: payload.unitId, subUnitId: payload.subUnitId, pivotTableId: payload.pivotTableId, revision: payload.revision }))
      this.publishObserverEvent('PivotTableRendered', () => cloneModel(payload))
    }
    for (const payload of charts) this.publishObserverEvent('PivotChartChanged', () => ({ ...payload }))
    return true
  }

  private commandKind(commandId: string, params: unknown): OfficeKind | undefined {
    const unitId = commandUnitId(params)
    if (unitId && this.documents.has(unitId)) return 'document'
    if (unitId && this.workbooks.has(unitId)) return 'workbook'
    if (commandId.startsWith('doc.') || commandId.startsWith('docs.') || commandId.startsWith('docs-latex.')) return 'document'
    if (commandId.startsWith('sheet.') || commandId.startsWith('workbook.')) return 'workbook'
    return undefined
  }

  /** A Before listener may reopen the same ID with revision zero. Capture the
   * model owner and history generation before dispatching view/page operations. */
  private captureDocumentTransientCommandTarget(commandId: string, params: unknown): { unitId: string; assertCurrent(): void } | undefined {
    const zoom = commandId === DOCUMENT_ZOOM_COMMANDS.SetRatio || commandId === DOCUMENT_ZOOM_COMMANDS.SetRatioCommand
    const headerFooter = isDocumentHeaderFooterCommand(commandId)
    const tocCommand = commandId === DOCUMENT_TOC_INSERT_COMMAND || commandId === DOCUMENT_TOC_UPDATE_COMMAND
      || commandId === DOCUMENT_TOC_DELETE_COMMAND || commandId === DOCUMENT_TOC_DIALOG_OPERATION || commandId === SET_DOCUMENT_HEADING_COLLAPSED_COMMAND
    if (!zoom && !headerFooter && !tocCommand && commandId !== DOCUMENT_PAGE_SETUP_COMMAND) return undefined
    const headerFooterTarget = headerFooter ? captureDocumentHeaderFooterEditorTarget(this,
      resolveDocumentHeaderFooterRequest(commandId, params)) : undefined
    const resolveUnitId = () => zoom ? resolveDocumentZoomCommand(commandId, params, this._activeDocumentId).unitId
      : headerFooter ? resolveDocumentHeaderFooterRequest(commandId, params).unitId ?? this._activeDocumentId
      : commandUnitId(params) ?? this._activeDocumentId
    const unitId = resolveUnitId(), source = this.getDocumentModel(unitId)
    const generation = source ? this.editor.history.captureSavepoint(unitId).generation : undefined
    const history = this.editor.history.units[unitId]
    const revision = tocCommand && source ? this.getDocumentRevision(unitId) : undefined
    const inputGeneration = tocCommand && source ? this.getDocumentInputState(unitId)?.generation : undefined
    const userId = this.currentUserId, activeId = this._activeDocumentId, activeKind = this.editor.state.activeKind
    return { unitId, assertCurrent: () => {
      if (resolveUnitId() !== unitId || this.getDocumentModel(unitId) !== source || this.editor.history.units[unitId] !== history
        || source && this.editor.history.captureSavepoint(unitId).generation !== generation
        || tocCommand && (this.currentUserId !== userId || this._activeDocumentId !== activeId || this.editor.state.activeKind !== activeKind
          || source && (this.getDocumentRevision(unitId) !== revision || this.getDocumentInputState(unitId)?.generation !== inputGeneration))
        || headerFooterTarget && !documentHeaderFooterEditorTargetCurrent(this, headerFooterTarget)) {
        throw new Error('命令执行前文档或打开代数已更改，请重新执行。')
      }
    } }
  }

  private captureWorkbookCommandTarget(commandId: string, params: unknown): { unitId: string; assertCurrent(): void } | undefined {
    const guard = workbookCommandGuard(params)
    if (!guard && commandId !== 'sheet.command.sort-range' && commandId !== 'sheet.command.toggle-auto-filter'
      && commandId !== CONFIGURE_SHEET_FILTER_COMMAND) return undefined
    const unitId = commandUnitId(params) ?? this._activeWorkbookId
    const source = this.getWorkbookModel(unitId), revision = this.getWorkbookRevision(unitId), userId = this.currentUserId
    const generation = source ? this.editor.history.captureSavepoint(unitId).generation : undefined
    const history = this.editor.history.units[unitId]
    const assertCurrent = (): void => {
      guard?.()
      if ((commandUnitId(params) ?? this._activeWorkbookId) !== unitId || this.currentUserId !== userId
        || this.getWorkbookModel(unitId) !== source || this.getWorkbookRevision(unitId) !== revision || this.editor.history.units[unitId] !== history
        || source && this.editor.history.captureSavepoint(unitId).generation !== generation) {
        throw new Error('命令执行前工作簿、用户或打开代数已更改，请重新执行。')
      }
    }
    assertCurrent()
    return { unitId, assertCurrent }
  }

  /** Capture both units and the clipboard before extension hooks run. A hook
   * cannot redirect an implicit paste or replace a unit with the same ID. */
  private captureChartPasteCommandTarget(commandId: string, params: unknown): { unitId: string; assertCurrent(): void } | undefined {
    if (commandId !== PASTE_COPIED_CHART_COMMAND) return undefined
    if (params !== undefined && (!params || typeof params !== 'object' || Array.isArray(params))) throw new TypeError('图表粘贴参数必须为对象。')
    const input = (params ?? {}) as ChartObjectPasteInput
    const resolveTarget = () => {
      const unitId = input.unitId ?? this.activeWorkbookId
      return { unitId, sheetId: chartObjectSheetId(input) ?? this.getWorkbookModel(unitId)?.activeSheetId }
    }
    const target = resolveTarget(), targetOwner = this.getWorkbookModel(target.unitId), userId = this.currentUserId
    const store = chartObjectClipboardStore(this.pinia), clipboardId = store.clipboard?.id, ticket = store.request
    const sourceUnitId = store.clipboard?.source.unitId, sourceOwner = sourceUnitId ? this.getWorkbookModel(sourceUnitId) : undefined
    const targetHistory = this.editor.history.units[target.unitId], sourceHistory = sourceUnitId ? this.editor.history.units[sourceUnitId] : undefined
    const targetGeneration = targetOwner ? this.editor.history.captureSavepoint(target.unitId).generation : undefined
    const sourceGeneration = sourceOwner && sourceUnitId ? this.editor.history.captureSavepoint(sourceUnitId).generation : undefined
    return { unitId: target.unitId, assertCurrent: () => {
      const resolved = resolveTarget()
      if (resolved.unitId !== target.unitId || resolved.sheetId !== target.sheetId || this.currentUserId !== userId
        || this.getWorkbookModel(target.unitId) !== targetOwner || this.editor.history.units[target.unitId] !== targetHistory
        || targetOwner && this.editor.history.captureSavepoint(target.unitId).generation !== targetGeneration
        || sourceUnitId && (this.getWorkbookModel(sourceUnitId) !== sourceOwner || this.editor.history.units[sourceUnitId] !== sourceHistory
          || sourceOwner && this.editor.history.captureSavepoint(sourceUnitId).generation !== sourceGeneration)
        || store.clipboard?.id !== clipboardId || store.request !== ticket) {
        throw new Error('命令执行前图表粘贴的来源、目标或剪贴板已更改，请重新执行。')
      }
    } }
  }

  private assertContentUnitIdentity(kind: OfficeKind, id: string): void {
    const opposite = kind === 'workbook' ? this.documents.has(id) || this.editor.document.model.id === id
      : this.workbooks.has(id) || this.editor.workbook.model.id === id
    if (opposite) throw new TypeError('工作簿与文档必须使用不同标识，不能共用内容历史。')
  }

  private assertAlive(): void {
    if (this.disposed) throw new Error('OfficeStudio has been disposed')
  }
}

function commandUnitId(params: unknown): string | undefined {
  if (!params || typeof params !== 'object') return undefined
  if ('unitId' in params && typeof params.unitId === 'string') return params.unitId
  return 'documentId' in params && typeof params.documentId === 'string' ? params.documentId : undefined
}

function requirePivot(model: WorkbookModel, pivotTableId: string) {
  const pivot = model.pivotTables?.find((candidate) => candidate.id === pivotTableId)
  if (!pivot) throw new Error(`Pivot table not found: ${pivotTableId}`)
  return pivot
}

function requireSheet(model: WorkbookModel, sheetId: string) {
  const sheet = model.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
  return sheet
}

function rangeEditFromParams(params: unknown): { sheetId: string; range: CellRange } | undefined {
  if (!params || typeof params !== 'object') return undefined
  const input = params as Record<string, unknown>
  if (typeof input.sheetId !== 'string' || !input.range || typeof input.range !== 'object') return undefined
  const source = input.range as Record<string, unknown>
  const startCol = source.startCol ?? source.startColumn
  const endCol = source.endCol ?? source.endColumn
  const values = [source.startRow, startCol, source.endRow, endCol]
  if (!values.every(value => typeof value === 'number' && Number.isInteger(value))) return undefined
  return {
    sheetId: input.sheetId,
    range: { startRow: source.startRow as number, startCol: startCol as number, endRow: source.endRow as number, endCol: endCol as number },
  }
}

function requireSlicer(model: WorkbookModel, slicerId: string) {
  const slicer = model.slicers?.find((candidate) => candidate.id === slicerId)
  if (!slicer) throw new Error(`Slicer not found: ${slicerId}`)
  return slicer
}

function requireChart(model: WorkbookModel, chartId: string): ChartModel {
  const chart = model.charts?.find((candidate) => candidate.id === chartId)
  if (!chart) throw new Error(`Chart not found: ${chartId}`)
  return chart
}

