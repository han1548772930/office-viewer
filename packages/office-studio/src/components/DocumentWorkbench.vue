<script setup lang="ts">
import { ChevronLeft, ChevronRight, FileText, FileUp, Printer, Save, Search, X } from 'lucide-vue-next'
import { Button, Checkbox, ConfigProvider, Form, FormItem, Input, Menu, Modal, Splitter, SplitterPanel, Tooltip, Upload, theme } from 'antdv-next'
import type { MenuItemType } from 'antdv-next'
import { computed, inject, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue'
import type { CSSProperties } from 'vue'
import { OfficeFontListKey } from '../platform/ribbon/ribbon-fonts'
import { selectionFontPatch } from '../features/document/model/document-font'
import type { DocumentContentControl, DocumentContentControlType, DocumentModel, DocumentPageSetup, DocumentParagraphStyle, DocumentPermissionRule } from '../domain/model'
import type { FDocument } from '../api/facade'
import type { ApiEventMap } from '../api/events'
import { createBlankDocument } from '../domain/model'
import DocumentCanvasViewport from './document/DocumentCanvasViewport.vue'
import DocumentContentControlPanel from './document/DocumentContentControlPanel.vue'
import DocumentOutline from './document/DocumentOutline.vue'
import DocumentFooter from './document/DocumentFooter.vue'
import DocumentPageSettings from './document/DocumentPageSettings.vue'
import DocumentHeaderFooterPanel from './document/DocumentHeaderFooterPanel.vue'
import DocumentCodeBlockPanel from './document/DocumentCodeBlockPanel.vue'
import DocumentLatexPanel from './document/DocumentLatexPanel.vue'
import DocumentNotesPanel from './document/DocumentNotesPanel.vue'
import DocumentPermissionPanel from './document/DocumentPermissionPanel.vue'
import DocumentParagraphSettingsPanel from './document/DocumentParagraphSettingsPanel.vue'
import DocumentSectionSettingsPanel from './document/DocumentSectionSettingsPanel.vue'
import DocumentCommentsPanel from './document/DocumentCommentsPanel.vue'
import DocumentTableOfContentsDialog from './document/DocumentTableOfContentsDialog.vue'
import DocumentTableInsertDialog from './document/DocumentTableInsertDialog.vue'
import DocumentTablePropertiesDialog, { type DocumentTablePropertiesKind } from './document/DocumentTablePropertiesDialog.vue'
import { createDocumentTableRibbon, documentTableExtraContextItems, type DocumentTableRibbonContext } from './document/document-table-ribbon'
import { documentTableCommandState } from '../features/document/commands/table-commands'
import { isDocumentTableCommand, type DocumentTableCommand, type DocumentTableInsertOptions } from '../features/document/commands/table-command-types'
import { documentTableGrid, documentTableWholeRange, type DocumentTableCellRange } from '../features/document/model/document-table-grid'
import { findDocumentBlock, tableForBlock } from '../features/document/model/document-model'
import WorkbenchRibbon from '../platform/ribbon/WorkbenchRibbon.vue'
import type { RibbonItem, RibbonVariant } from '../platform/ribbon/ribbon-schema'
import { useWorkbenchDisplay, workbenchVisibility } from '../platform/workbench-display'
import { WORKBENCH_CONTROL_THEME } from '../platform/workbench-theme'
import { activeOutlineIdForBlock, buildDocumentOutline, documentOutlineEnabled, documentOutlinePlacement, type DocumentOutlineItem } from '../features/document/services/document-outline'
import { createDocumentRibbon, documentCommandFromRibbon, documentTableContextMenuItems, documentTocContextMenuItems, type DocumentTocContextActionStates } from './document/document-ribbon'
import { clampDocumentZoom, documentCharacterCount, documentSelectionCharacterCount, documentStoryToolbarState, documentToolbarState, type DocumentToolbarState } from './document/document-toolbar-state'
import { documentParagraphSettingsStyle, type DocumentParagraphLineMetrics } from './document/document-paragraph-settings-state'
import { documentOrientedPage, documentPageSettingsError, documentPaperSize, type DocumentPageMode, type DocumentPaperSize } from './document/document-page-settings'
import { blockAtPoint, canChangeDocumentParagraphIndent, canContinueDocumentNumbering, canRestartDocumentNumbering, documentCommandSupportsBlockHost, executeDocumentCommand, type DocumentCommand } from '../features/document/commands/document-commands'
import { emptyLayout, type DocumentLayout } from '../features/document/layout/document-layout'
import { findDocumentStoryPage } from '../features/document/layout/document-story-navigation'
import { createDocumentNumericPageField } from '../features/document/model/document-page-field-result'
import type { DocumentViewportLocation } from '../features/document/layout/document-viewport'
import { blocksFromHtml, clampPoint, cloneDocument, collapsedSelection, documentHtml, documentRunFormatPatch, marksAtPoint, normalizeDocument, runText, runsForTarget, safeUrl, sameTarget, selectedPlainText, selectionCollapsed, textTargets, type DocumentPoint, type DocumentSelection, type RunPatch } from '../features/document/model/document-model'
import { DOCUMENT_DEFAULT_TEXT_COLOR, rememberDocumentColors, type DocumentColorMemory } from '../features/document/services/document-colors'
import { printDocumentModel } from '../features/document/services/document-print'
import { DOCUMENT_PASTE_OPTIONS, documentTableClipboardData, documentTablePasteCommand, documentPasteSpecialCommand, readDocumentClipboard, writeDocumentClipboard, type DocumentClipboardData, type DocumentPasteMode } from '../features/document/services/document-clipboard'
import { documentSectionIds, resolveSectionHeaderFooter, sectionIdForBlock, sectionStoryLinkedToPrevious, type SectionHeaderFooterPatch } from '../features/document/services/section-header-footer'
import { editStoryBuffer, replaceStoryRange, storyBuffer, storyDateField, type StoryLocation } from '../features/document/services/story-editor'
import { canRefreshDocumentStoryFields, canRefreshSelectedDocumentBodyFields, documentBodySelectionHasField, documentStoryRangeHasField } from '../features/document/services/document-story-fields'
import { releaseDocumentFieldOpenRefresh, reserveDocumentFieldOpenRefresh, type DocumentFieldCacheAdapter } from '../features/document/services/document-field-open-refresh'
import { headerFooterActionDisabled, headerFooterRequestScopeMatches, parseHeaderFooterAction } from './document/document-header-footer-capabilities'
import { documentSectionSetup } from '../features/document/services/document-settings.ts'
import { selectedDocumentContentControl } from '../features/document/services/content-controls'
import { CONTENT_CONTROL_CHECKBOX_STYLES, CONTENT_CONTROL_DATE_FORMATS, contentControlCheckboxStyle } from '../features/document/services/content-control-presentation'
import type { DocumentContentControlCommandRequest, DocumentContentControlSelectionRequest } from '../features/document/commands/content-control-command-types'
import { contentControlValueCommandState } from '../features/document/commands/content-control-command-state'
import { documentContentControlInteractionCurrent } from '../features/document/controllers/content-control-interaction'
import { codeBlockData, editCodeBlockText, exitCodeBlock, insertCodeBlock, updateCodeBlock, type DocumentCodeBlockData } from '../features/document/code-blocks/code-blocks.ts'
import { getLatexFormula, insertLatexFormula, updateLatexFormula, type LatexFormulaMode } from '../features/document/latex/latex-formula.ts'
import { deleteDocumentNote, insertDocumentNote, noteReferencePoint, updateDocumentNote, type DocumentNoteKind } from '../features/document/notes/document-notes.ts'
import { createDocumentThreadComment, documentCommentAnchorsForSelection, documentCommentTableSelection, deleteDocumentCommentMessage, deleteDocumentThreadComment, editDocumentCommentMessage, replyToDocumentThreadComment, setDocumentThreadCommentResolved } from '../features/document/comments/thread-comments.ts'
import { canManageDocumentPermissions, permissionTargetsForSelection, resolveDocumentPermission } from '../features/document/permissions/document-permissions.ts'
import { documentPermissionTargetOptions } from '../features/document/permissions/permission-rules.ts'
import { findDocumentMatches, pointForMatch, replaceDocumentMatches } from '../features/document/services/document-search.ts'
import { buildDocumentTocBlocks, DEFAULT_DOCUMENT_TOC_OPTIONS, documentTocIdForSelection, documentTocIdAtBlock, documentTocOptionsAtBlock, documentTocPageNumbers, supportsDocumentTableOfContents, type DocumentTocOptions, type DocumentTocUpdateMode } from '../features/document/toc/document-toc.ts'
import { documentTocActionState, type DocumentTocActionState } from '../features/document/toc/document-toc-command'
import { createRibbonTranslator, provideWorkbenchRegionalSettings } from '../i18n/office-i18n'
import { captureDocumentFormat, documentPainterTarget, type DocumentPainterFormat } from '../features/document/services/document-format-painter'
import { documentClearFormattingState, prepareDocumentStoryClearFormatting } from '../features/document/services/document-clear-formatting'
import { documentStoryTextFormattingState, prepareDocumentStoryTextFormatting } from '../features/document/services/document-story-formatting'
import { assertDocumentTextSelectionEditable, documentTextFormattingState } from '../features/document/services/document-text-formatting'
import { validateDocumentRunPatch } from '../features/document/model/document-run-range'
import { DOCUMENT_STORY_KEYS, clampDocumentTextRange, type DocumentStoryKey, type DocumentStoryRange } from '../features/document/model/document-story'
import type { DocumentInputSession, DocumentStoryInputCommit } from '../features/document/model/document-input-session'
import { sanitizeDocumentInputSelection } from '../features/document/model/document-input-session'
import { cloneDocumentSelection, documentTableSelectedTextTargets, sameDocumentTableViewSelection } from '../features/document/model/document-table-selection'
import { documentBlockHostForId, documentBlockHostId, sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { applyDocumentBlockHostModel, documentBlockHostSelectionFromLegacyRange, documentModelForBlockHost, resolveDocumentBlockHost } from '../features/document/services/document-block-host'
import { documentLayoutForBlockHost } from '../features/document/layout/document-auxiliary-layout'
import { assertDocumentStoryLocksPreserved } from '../features/document/services/document-story-locks'
import { permissionTargetsForChangedBlocks } from '../features/document/permissions/document-permissions'

defineOptions({ inheritAttrs: false })

// Splitter creates its root and panel DOM in a ResizeObserver slot without this component's scope attribute.
const documentSplitterStyles: { root: CSSProperties; panel: CSSProperties } = {
  root: { flex: '1 1 0', height: 'auto', minWidth: 0, minHeight: 0, overflow: 'hidden', alignItems: 'stretch' },
  panel: { position: 'relative', display: 'flex', height: 'auto', minWidth: 0, minHeight: 0, overflow: 'hidden' },
}

const props = withDefaults(defineProps<{ modelValue?: DocumentModel | string; fileName?: string; canUndo?: boolean; canRedo?: boolean; documentApi?: FDocument; fieldCacheAdapter?: DocumentFieldCacheAdapter; showHeader?: boolean; showOutline?: boolean; permissionUserId?: string; permissionAdminUserId?: string; permissionDefaultAllow?: boolean; ribbonVariant?: RibbonVariant }>(), { showHeader: false, showOutline: undefined, permissionUserId: 'local-user', permissionDefaultAllow: true, ribbonVariant: 'classic' })
const { locale: officeLocale, direction: officeDirection } = provideWorkbenchRegionalSettings(() => props.documentApi?.regionalSettings)
const emit = defineEmits<{
  'update:modelValue': [value: DocumentModel]
  'update:fileName': [value: string]
  'document-change': [value: DocumentModel]
  'document-cache-change': [value: DocumentModel]
  'import-file': [file: File]
  save: [payload: { fileName: string; html: string; model: DocumentModel }]
  undo: [unitId?: string]
  redo: [unitId?: string]
  notify: [message: string]
}>()
const model = shallowRef<DocumentModel>(normalizeDocument(createBlankDocument()))
const firstTarget = textTargets(model.value)[0]!
const selection = shallowRef<DocumentSelection>(collapsedSelection({ ...firstTarget, offset: 0 }))
const layout = shallowRef<DocumentLayout>(emptyLayout())
const layoutSource = shallowRef<DocumentModel>()
const editingHost = shallowRef<DocumentAuxiliaryBlockHost>()
const returnBodySelection = shallowRef<DocumentSelection>()
const editingModel = computed(() => editingHost.value ? documentModelForBlockHost(model.value, editingHost.value) : model.value)
const editingLayout = computed(() => documentLayoutForBlockHost(layout.value, editingHost.value))
const canvasRef = ref<InstanceType<typeof DocumentCanvasViewport> | null>(null)
const activeTab = ref('home')
const { token: themeToken } = theme.useToken()
const workbenchStyle = computed(() => ({
  '--document-text': themeToken.value.colorText,
  '--document-secondary-text': themeToken.value.colorTextSecondary,
  '--document-background': themeToken.value.colorBgLayout,
  '--document-surface': themeToken.value.colorBgContainer,
  '--document-border': themeToken.value.colorBorder,
  '--document-secondary-border': themeToken.value.colorBorderSecondary,
  '--document-primary': themeToken.value.colorPrimary,
  '--document-error': themeToken.value.colorError,
  '--document-info-text': themeToken.value.colorInfoText,
  '--document-info-background': themeToken.value.colorInfoBg,
  '--document-info-border': themeToken.value.colorInfoBorder,
  '--document-shadow': themeToken.value.boxShadowSecondary,
}))
const colorMemory = shallowRef<DocumentColorMemory>({})
const documentTitle = ref(props.fileName || '无标题文档')
const displayPreference = useWorkbenchDisplay()
const chrome = computed(() => workbenchVisibility(displayPreference?.value.chrome ?? 'full'))
const zoom = ref<number>(displayPreference?.value.zoom ?? 100)
watch(() => displayPreference?.value.zoom, value => { if (value !== undefined) setZoom(value) })
const paperSize = ref<DocumentPaperSize>('A4')
const orientation = ref<'portrait' | 'landscape'>('portrait')
const outlineOverride = ref<boolean>()
const showOutline = computed({ get: () => documentOutlineEnabled(model.value, outlineOverride.value ?? props.showOutline), set: (value: boolean) => { outlineOverride.value = value } })
const isFocusMode = ref(false)
const isFullscreen = ref(false)
const workbenchRef = ref<HTMLElement | null>(null)
const workbenchWidth = ref(1440)
const visibleBlockId = ref<string>()
const outlineNavigationTarget = ref<string>()
const outlineMode = computed(() => documentOutlinePlacement({ availableWidth: workbenchWidth.value, pageWidth: model.value.page.width, zoom: zoom.value }))
const outlineItems = computed(() => buildDocumentOutline(model.value))
const activeOutlineId = computed(() => activeOutlineIdForBlock(outlineItems.value, model.value, visibleBlockId.value ?? selection.value.focus.blockId))
const showFind = ref(false)
const searchText = ref('')
const replaceText = ref('')
const searchCaseSensitive = ref(false)
const searchWholeWord = ref(false)
const searchInSelection = ref(false)
const searchScopeSelection = shallowRef<DocumentSelection>()
watch(searchInSelection, (enabled) => {
  searchScopeSelection.value = enabled && !selectionCollapsed(selection.value)
    ? cloneDocumentSelection(selection.value)
    : undefined
})
const searchMatches = computed(() => showFind.value ? findDocumentMatches(editingModel.value, searchText.value, {
  caseSensitive: searchCaseSensitive.value,
  wholeWord: searchWholeWord.value,
  locale: officeLocale.value,
  selection: searchInSelection.value ? searchScopeSelection.value : undefined,
}) : [])
const activePageIndex = ref(0)
const standalonePendingMarks = shallowRef<RunPatch>({})
let modelReceived = false
let inputStateActive = true
let inputStateEpoch = 0
const storyActionVersion = ref(0)
let applyingInputState = false
let applyingOpenFieldCache = false
const pendingMarks = computed<RunPatch>(() => props.documentApi
  ? props.documentApi.id === model.value.id ? props.documentApi.getInputMarks(selection.value, currentStoryRange.value, editingHost.value) : {}
  : standalonePendingMarks.value)
const modalOpen = ref(false)
const modalKind = ref<'image' | 'link'>('image')
const modalValue = ref('')
const modalError = ref('')
const tableInsertOpen = ref(false)
const tableFloatingToolbarOpen = ref(false)
const tableColors = shallowRef<{ background?: string; border?: string }>({})
const tableProperties = shallowRef<{ open: boolean; tableId: string; range: DocumentTableCellRange; kind: DocumentTablePropertiesKind; host?: DocumentAuxiliaryBlockHost; type?: 'text' | 'number' | 'date' | 'checkbox' | 'dropdown' }>()
const notice = ref('')
const documentContextMenu = shallowRef<{ left: number; top: number; fieldRequest?: { source: DocumentModel; story?: DocumentStoryRange; host?: DocumentAuxiliaryBlockHost; bodySelection?: DocumentSelection; requestScope: number } }>()
const contentControlPanelOpen = ref(false)
const activeContentControlId = ref<string>()
const pictureReplacementControlId = ref<string>()
const selectedContentControlId = computed(() => {
  if (selection.value.tableSelection) return undefined
  if (activeContentControlId.value && model.value.contentControls?.some(control => control.id === activeContentControlId.value)) return activeContentControlId.value
  const anchor = selection.value.anchor
  const focus = selection.value.focus
  const blockId = focus.cellBlockId ?? focus.blockId
  if ((anchor.cellBlockId ?? anchor.blockId) !== blockId) return undefined
  const start = Math.min(anchor.offset, focus.offset)
  const end = Math.max(anchor.offset, focus.offset)
  return selectedDocumentContentControl(editingModel.value, blockId, start, end)?.id
})
const selectedContentControl = computed(() => selectedContentControlId.value
  ? model.value.contentControls?.find(control => control.id === selectedContentControlId.value)
  : undefined)
const pageSettingsOpen = ref(false)
const headerFooterPanelOpen = ref(false)
const storyLocation = ref<StoryLocation>()
const legacyStoryActive = computed(() => !!storyLocation.value && !editingHost.value)
const storySelection = shallowRef({ start: 0, end: 0 })
const storyComposing = ref(false)
let storyOptionsPendingEpoch: number | undefined
const previousStoryTab = ref('home')
const storyPage = computed(() => {
  const page = storyLocation.value && layout.value.pageHeaderFooters[storyLocation.value.page]
  return page && !page.isParityBlank ? page : undefined
})
const headerFooterSectionId = computed(() => editingHost.value?.kind === 'story' ? editingHost.value.sectionId
  : storyPage.value?.sectionId ?? sectionIdForBlock(model.value, returnBodySelection.value?.focus.blockId ?? selection.value.focus.blockId))
const storyKey = computed(() => editingHost.value?.kind === 'story' ? editingHost.value.story
  : storyLocation.value?.area === 'footer' ? storyPage.value?.footerStory : storyPage.value?.headerStory)
const storySection = computed(() => resolveSectionHeaderFooter(model.value, headerFooterSectionId.value))
const storyLinkedToPrevious = computed(() => !!storyKey.value && sectionStoryLinkedToPrevious(storySection.value, storyKey.value))
const storyValue = computed(() => storyKey.value && storyBuffer(storySection.value, storyKey.value).value)
const paragraphSettingsPanelOpen = ref(false)
const sectionSettingsPanelOpen = ref(false)
const activeParagraphBlock = computed(() => {
  const point = selection.value.tableSelection ? documentTableSelectedTextTargets(editingModel.value, selection.value.tableSelection)[0] : selection.value.focus
  return point && blockAtPoint(editingModel.value, point)
})
const paragraphSettingsStyle = computed(() => documentParagraphSettingsStyle(activeParagraphBlock.value?.style ?? {}))
const paragraphSettingsFontSize = computed(() => activeParagraphBlock.value?.runs[0]?.fontSize ?? 11)
const paragraphSettingsSelectionKey = computed(() => `${model.value.id}:${editingHost.value ? documentBlockHostId(editingHost.value) : 'body'}:${JSON.stringify(selection.value)}`)
const paragraphSettingsLineMetrics = computed<DocumentParagraphLineMetrics | undefined>(() => {
  const focus = selection.value.focus
  const lines = editingLayout.value.lines.filter(line => sameTarget(line.target, focus))
  const line = lines.find(candidate => focus.offset >= candidate.start && focus.offset < candidate.end) ?? lines.at(-1)
  return line?.glyphLineHeight === undefined ? undefined : { glyphLineHeight: line.glyphLineHeight, renderedLineHeight: line.height }
})
const activeSectionSetup = computed(() => documentSectionSetup(model.value, headerFooterSectionId.value))
const sectionOptions = computed(() => documentSectionIds(model.value).map((value, index) => ({ value, label: `#${index + 1}` })))
const pageMode = computed<DocumentPageMode>(() => documentOutlineEnabled(model.value) ? 'modern' : 'traditional')
const codeBlockPanelOpen = ref(false)
const latexPanelOpen = ref(false)
const notesPanelOpen = ref(false)
const permissionPanelOpen = ref(false)
const commentsPanelOpen = ref(false)
const documentDockWidth = ref(384)
const documentDockOpen = computed(() => paragraphSettingsPanelOpen.value || commentsPanelOpen.value)
const activeCommentId = ref('')
const commentDraftSelection = shallowRef<DocumentSelection>()
const commentDraftHost = shallowRef<DocumentAuxiliaryBlockHost>()
const commentDraftReference = ref('')
const permissionContext = computed(() => ({
  documentId: model.value.id ?? '',
  userId: props.documentApi?.getCurrentUserId() ?? props.permissionUserId,
  permissionAdminUserId: props.permissionAdminUserId ?? props.documentApi?.getPermissionAdminUserId(),
  defaultAllow: props.permissionDefaultAllow,
}))
const permissionTargets = computed(() => documentPermissionTargetOptions(model.value, createRibbonTranslator(officeLocale.value, 'doc')))
const canManagePermissionRules = computed(() => canManageDocumentPermissions(permissionContext.value))
const canComment = computed(() => resolveDocumentPermission(model.value.permissionRules ?? [], permissionContext.value, 'comment').allowed)
const documentInputComposing = computed(() => storyComposing.value || canvasRef.value?.isBodyComposing() === true)
const canRefreshFields = computed(() => !documentInputComposing.value && canRefreshDocumentStoryFields(model.value, permissionContext.value))
const fieldCacheState = computed(() => props.fieldCacheAdapter?.getState())
const fieldOpenHistoryGeneration = computed(() => props.documentApi
  ? props.documentApi.id === model.value.id ? props.documentApi.getHistoryGeneration() : undefined
  : fieldCacheState.value?.unitId === model.value.id ? fieldCacheState.value.historyGeneration : undefined)
const canCreateComment = computed(() => {
  if (legacyStoryActive.value || editingHost.value?.kind === 'story' && storyLinkedToPrevious.value || !canComment.value || selectionCollapsed(selection.value)
    || !resolveDocumentPermission(model.value.permissionRules ?? [], permissionContext.value, 'comment', permissionTargetsForSelection(editingModel.value, selection.value)).allowed) return false
  try { return documentCommentAnchorsForSelection(editingModel.value, selection.value).length > 0 }
  catch (error) { if (error instanceof RangeError || error instanceof TypeError) return false; throw error }
})
const deletableCommentIds = computed(() => Object.values(model.value.comments ?? {})
  .filter(comment => canComment.value && comment.messages[0]?.authorId === permissionContext.value.userId)
  .map(comment => comment.id))
const activeCodeBlock = computed(() => !selection.value.tableSelection && !legacyStoryActive.value ? blockAtPoint(editingModel.value, selection.value.focus) : undefined)
const activeCodeBlockData = computed(() => activeCodeBlock.value && codeBlockData(activeCodeBlock.value))
const activeLatexFormula = computed(() => {
  const block = activeCodeBlock.value
  if (!block || !sameTarget(selection.value.anchor, selection.value.focus)) return undefined
  const start = Math.min(selection.value.anchor.offset, selection.value.focus.offset)
  const end = Math.max(selection.value.anchor.offset, selection.value.focus.offset)
  let offset = 0
  for (const run of block.runs) {
    const formula = getLatexFormula(run)
    const runEnd = offset + run.text.length
    if (formula && (start === end ? start >= offset && start <= runEnd : start <= offset && end >= runEnd)) return { runId: run.id, formula }
    offset = runEnd
  }
  return undefined
})
const pages = computed(() => layout.value.pages)
const wordCount = computed(() => documentCharacterCount(textTargets(model.value).map((target) => runText(runsForTarget(model.value, target))).join('\n')))
const selectedCharacters = computed(() => documentSelectionCharacterCount(editingModel.value, selection.value))
const toolbarState = computed<DocumentToolbarState>(() => {
  const story = currentStoryRange.value
  if (story) return documentStoryToolbarState(model.value, story, pendingMarks.value)
  if (legacyStoryActive.value) return { marks: {}, mixedMarks: [], paragraph: {}, hasSelection: false }
  return documentToolbarState(editingModel.value, selection.value, pendingMarks.value)
})
const tableRibbonContext = computed<DocumentTableRibbonContext>(() => {
  const id = toolbarState.value.table?.blockId, source = id && findDocumentBlock(editingModel.value, id), table = source && tableForBlock(source)
  const fragment = layoutSource.value === model.value ? editingLayout.value.tableFragments?.find(fragment => fragment.tableId === id && !fragment.repeatedHeader) : undefined
  const heights = fragment?.rows.map(row => row.contentHeight ?? row.height)
  return { model: editingModel.value, selection: selection.value, permission: permissionContext.value, mode: pageMode.value,
    composing: documentInputComposing.value || legacyStoryActive.value || !!editingHost.value && storyLinkedToPrevious.value, floatingToolbarOpen: tableFloatingToolbarOpen.value,
    lastBackgroundColor: tableColors.value.background, lastBorderColor: tableColors.value.border,
    ...(table && heights?.length === table.rows.length ? { measuredRowHeights: heights } : {}) }
})
const tableContextActions = computed<RibbonItem[]>(() => {
  if (!toolbarState.value.table) return []
  const context = tableRibbonContext.value, tab = createDocumentTableRibbon(context)
  const basic = documentTableContextMenuItems(createRibbonTranslator(officeLocale.value, 'doc')).map(item => {
    const command = documentCommandFromRibbon(item.commandId ?? '', item.params)
    const state = command && isDocumentTableCommand(command) ? documentTableCommandState(editingModel.value, selection.value, command, permissionContext.value) : { available: false, reason: '表格命令无效' }
    return { ...item, disabled: context.composing || !state.available, disabledReason: state.reason }
  })
  const extra = tab?.groups.flatMap(group => group.items).filter(item => !['doc-table-insert', 'doc-table-delete'].includes(item.id) && !item.hidden).map(item => {
    if (item.inlineComponent) return { ...item, type: 'button' as const, inlineComponent: undefined, commandId: 'doc.operation.open-table-properties',
      children: undefined, params: { tableId: toolbarState.value.table!.blockId, range: { startRow: toolbarState.value.table!.startRow, endRow: toolbarState.value.table!.endRow,
        startColumn: toolbarState.value.table!.startColumn, endColumn: toolbarState.value.table!.endColumn }, kind: item.id === 'doc-table-background' ? 'background' : 'border' } }
    // Width/style selectors are represented by the same complete border form in a menu.
    if (item.type === 'select') return { ...item, type: 'button' as const, commandId: 'doc.operation.open-table-properties',
      params: { tableId: toolbarState.value.table!.blockId, range: { startRow: toolbarState.value.table!.startRow, endRow: toolbarState.value.table!.endRow,
        startColumn: toolbarState.value.table!.startColumn, endColumn: toolbarState.value.table!.endColumn }, kind: 'border' } }
    return item
  }) ?? []
  return [...basic, ...extra, ...documentTableExtraContextItems(context)]
})
const selectedTocId = computed(() => !editingHost.value && !legacyStoryActive.value && supportsDocumentTableOfContents(model.value)
  ? documentTocIdForSelection(model.value, selection.value)
  : undefined)
const selectedTocOptions = computed(() => documentTocOptionsAtBlock(blockAtPoint(model.value, selection.value.focus)) ?? { ...DEFAULT_DOCUMENT_TOC_OPTIONS })
const previousTabBeforeToc = ref<string>()
const canEditContextSelection = computed(() => resolveDocumentPermission(
  model.value.permissionRules ?? [], permissionContext.value, 'edit', permissionTargetsForSelection(editingModel.value, selection.value),
).allowed && !(editingHost.value?.kind === 'story' && storyLinkedToPrevious.value))
const canCopyContextSelection = computed(() => resolveDocumentPermission(model.value.permissionRules ?? [], permissionContext.value,
  'copy', permissionTargetsForSelection(editingModel.value, selection.value)).allowed)
const currentStoryRange = computed<DocumentStoryRange | undefined>(() => legacyStoryActive.value && storyKey.value
  ? { sectionId: headerFooterSectionId.value, story: storyKey.value, ...clampDocumentTextRange(storyValue.value ?? '', storySelection.value) } : undefined)
watch([model, () => props.documentApi, currentStoryRange, editingHost, selection, pendingMarks], () => { storyActionVersion.value++ }, { flush: 'sync' })
const currentStorySelection = computed(() => currentStoryRange.value ?? storySelection.value)
const textFormattingState = computed(() => storyComposing.value ? { available: false, reason: '请先完成当前输入法输入。' }
  : currentStoryRange.value ? documentStoryTextFormattingState(model.value, currentStoryRange.value, permissionContext.value)
  : legacyStoryActive.value ? { available: false, reason: '当前页眉页脚片段不可用。' }
  : documentTextFormattingState(editingModel.value, selection.value, permissionContext.value))
const clearFormattingState = computed(() => storyComposing.value ? { available: false, reason: '请先完成当前输入法输入。' }
  : documentClearFormattingState(currentStoryRange.value ? model.value : editingModel.value, selection.value, permissionContext.value, currentStoryRange.value))
const currentTocPageNumbers = computed(() => layoutSource.value === model.value ? documentTocPageNumbers(layout.value.headings, layout.value.pages) : {})
const selectedTocActionStates = computed<Record<DocumentTocUpdateMode | 'delete', DocumentTocActionState>>(() => {
  const tocId = selectedTocId.value
  if (!tocId || documentInputComposing.value) {
    const unavailable = { available: false, reason: tocId ? '请先完成当前输入法输入。' : '请先选择目录。' }
    return { pageNumbersOnly: unavailable, all: unavailable, delete: unavailable }
  }
  return {
    pageNumbersOnly: documentTocActionState(model.value, selection.value, tocId, 'pageNumbersOnly', currentTocPageNumbers.value, permissionContext.value),
    all: documentTocActionState(model.value, selection.value, tocId, 'all', currentTocPageNumbers.value, permissionContext.value),
    delete: documentTocActionState(model.value, selection.value, tocId, 'delete', currentTocPageNumbers.value, permissionContext.value),
  }
})
const tocContextActionStates = computed<DocumentTocContextActionStates>(() => ({
  update: selectedTocActionStates.value.pageNumbersOnly.available || selectedTocActionStates.value.all.available
    ? { available: true } : { available: false, reason: selectedTocActionStates.value.pageNumbersOnly.reason ?? selectedTocActionStates.value.all.reason },
  delete: selectedTocActionStates.value.delete,
}))
const canEditSelectedToc = computed(() => tocContextActionStates.value.update.available)
const canInsertToc = computed(() => !editingHost.value && !legacyStoryActive.value && supportsDocumentTableOfContents(model.value)
  && selection.value.focus.row == null
  && selectionCollapsed(selection.value)
  && !documentTocIdAtBlock(blockAtPoint(model.value, selection.value.focus))
  && canEditContextSelection.value)
watch(selectedTocId, (tocId) => {
  if (tocId) {
    if (activeTab.value !== 'table-of-contents') previousTabBeforeToc.value = activeTab.value
    activeTab.value = 'table-of-contents'
  } else if (activeTab.value === 'table-of-contents') {
    activeTab.value = previousTabBeforeToc.value ?? 'home'
    previousTabBeforeToc.value = undefined
  }
})
const tocDialog = ref<{ open: boolean; mode: 'insert' | 'update'; insertMode: 'automatic' | 'custom'; tocId?: string; updateMode: DocumentTocUpdateMode }>({
  open: false,
  mode: 'insert',
  insertMode: 'custom',
  updateMode: 'pageNumbersOnly',
})
const tocDialogOptions = ref<DocumentTocOptions>({ ...DEFAULT_DOCUMENT_TOC_OPTIONS })
const contextMenuActions = computed(() => {
  const translate = createRibbonTranslator(officeLocale.value, 'doc')
  const story = currentStoryRange.value
  if (story) return documentStoryRangeHasField(resolveSectionHeaderFooter(model.value, story.sectionId), story.story, story) ? [{
    id: 'doc-context-update-fields', title: translate('docs-ui.headerFooter.refreshFields', '更新字段'),
    commandId: 'doc.command.update-selected-fields', params: undefined,
    disabled: !canEditStory.value || storyLinkedToPrevious.value || storyComposing.value, group: 'format' as const,
  }] : []
  const pasteSpecialActions = DOCUMENT_PASTE_OPTIONS.map(({ value: mode, label }) => ({
    id: `doc-context-paste-special-${mode}`,
    title: translate(`doc.context.paste-special.${mode}`, label),
    commandId: 'doc.command.paste-special',
    params: { value: mode },
    disabled: !canEditContextSelection.value || typeof navigator === 'undefined' || !navigator.clipboard
      || (mode === 'destination' && !blockAtPoint(editingModel.value, selection.value.focus)),
    group: 'paste-special' as const,
  }))
  const quickActions = [
    { id: 'doc-context-copy', title: translate('doc.context.copy', '复制'), commandId: 'doc.command.copy', params: undefined, disabled: !toolbarState.value.hasSelection || !canCopyContextSelection.value, group: 'quick' as const },
    { id: 'doc-context-cut', title: translate('doc.context.cut', '剪切'), commandId: 'doc.command.cut', params: undefined, disabled: !toolbarState.value.hasSelection || !canCopyContextSelection.value || !canEditContextSelection.value, group: 'quick' as const },
    { id: 'doc-context-paste', title: translate('doc.context.paste', '粘贴'), commandId: 'doc.command.paste', params: undefined, disabled: !canEditContextSelection.value || typeof navigator === 'undefined' || !navigator.clipboard, group: 'quick' as const },
    { id: 'doc-context-delete', title: translate('doc.context.delete', '删除'), commandId: 'doc.command.delete', params: { direction: -1 }, disabled: !toolbarState.value.hasSelection || !canEditContextSelection.value, group: 'format' as const },
  ]
  const paragraphSettings = {
    id: 'doc-context-paragraph-settings', title: translate('doc.context.paragraph-settings', 'Paragraph settings'),
    commandId: 'sidebar.operation.doc-paragraph-setting-panel', params: undefined,
    disabled: !canEditContextSelection.value, group: 'layout' as const,
  }
  const layoutActions = [paragraphSettings]
  if (!editingHost.value && supportsDocumentTableOfContents(model.value) && selection.value.focus.row == null) layoutActions.push({
    id: 'doc-context-section-settings', title: translate('doc.context.section-settings', 'Section settings'),
    commandId: 'sidebar.operation.doc-section-setting-panel', params: undefined,
    disabled: !canEditContextSelection.value, group: 'layout' as const,
  })
  const tableActions = toolbarState.value.table
    ? tableContextActions.value.map(action => ({ ...action, group: 'table' as const }))
    : []
  const tocActions = selectedTocId.value
    ? documentTocContextMenuItems(selectedTocId.value, tocContextActionStates.value, translate).map(action => ({ ...action, group: 'layout' as const }))
    : []
  const fieldActions = documentBodySelectionHasField(editingModel.value, selection.value) ? [{
    id: 'doc-context-update-fields', title: translate('docs-ui.headerFooter.refreshFields', '更新字段'),
    commandId: 'doc.command.update-selected-fields', params: undefined,
    disabled: documentInputComposing.value || !canRefreshSelectedDocumentBodyFields(editingModel.value, selection.value, permissionContext.value), group: 'format' as const,
  }] : []
  return [...quickActions, ...pasteSpecialActions, ...fieldActions, ...layoutActions, ...tocActions, ...tableActions]
})
const documentContextMenuItems = computed<MenuItemType[]>(() => {
  const quickItems = contextMenuActions.value.filter(item => item.group === 'quick')
  const pasteSpecialActions = contextMenuActions.value.filter(item => item.group === 'paste-special')
  const formatItems = contextMenuActions.value.filter(item => item.group === 'format')
  const layoutItems = contextMenuActions.value.filter(item => item.group === 'layout')
  const tableItems = contextMenuActions.value.filter(item => item.group === 'table')
  return [
    ...quickItems.map(action => ({ key: action.id, label: action.title, disabled: action.disabled })),
    ...(pasteSpecialActions.length || formatItems.length ? [
      { type: 'divider' as const },
      ...(pasteSpecialActions.length ? [{
        key: 'doc-context-paste-special',
        label: createRibbonTranslator(officeLocale.value, 'doc')('doc.context.paste-special', '选择性粘贴'),
        children: pasteSpecialActions.map(action => ({ key: action.id, label: action.title, disabled: action.disabled })),
      }] : []),
      ...formatItems.map(action => ({ key: action.id, label: action.title, disabled: action.disabled })),
    ] : []),
    ...(layoutItems.length ? [{ type: 'divider' as const }, ...layoutItems.map(action => ({ key: action.id, label: action.title, disabled: action.disabled,
      title: 'disabledReason' in action && typeof action.disabledReason === 'string' ? action.disabledReason : undefined }))] : []),
    ...(tableItems.length ? [{ type: 'divider' as const }, ...tableItems.map(action => tableContextMenuItem(action))] : []),
  ]
})
function tableContextMenuItem(item: RibbonItem): MenuItemType {
  return { key: item.id, label: item.title, disabled: item.disabled, title: item.disabledReason,
    ...(item.children?.length ? { children: item.children.filter(child => !child.hidden).map(tableContextMenuItem) } : {}) }
}
const currentMarks = computed(() => toolbarState.value.marks)
const officeFonts = inject(OfficeFontListKey, undefined)
const painter = shallowRef<{ format: DocumentPainterFormat; continuous: boolean }>()
const modalTitle = computed(() => ({ image: '插入图片', link: '插入链接' })[modalKind.value])
const ribbonTabs = computed(() => {
  const control = selectedContentControl.value
  return createDocumentRibbon({
    table: !legacyStoryActive.value ? tableRibbonContext.value : undefined,
    fontList: officeFonts?.value,
    lastTextColor: colorMemory.value.color, lastHighlightColor: colorMemory.value.highlight,
    defaultTextColor: DOCUMENT_DEFAULT_TEXT_COLOR, defaultHighlightColor: themeToken.value.colorPrimary,
    numbering: { canRestart: canRestartDocumentNumbering(editingModel.value, selection.value, permissionContext.value), canContinue: canContinueDocumentNumbering(editingModel.value, selection.value, permissionContext.value) },
    indentation: { canDecrease: canChangeDocumentParagraphIndent(editingModel.value, selection.value, -1, permissionContext.value), canIncrease: canChangeDocumentParagraphIndent(editingModel.value, selection.value, 1, permissionContext.value) },
    painterActive: !!painter.value, painterUnavailable: !canEditContextSelection.value || legacyStoryActive.value,
    clearFormatting: clearFormattingState.value,
    mode: pageMode.value,
    editingDisabled: legacyStoryActive.value ? !canEditStory.value || storyLinkedToPrevious.value : !canEditContextSelection.value,
    textFormattingDisabled: !textFormattingState.value.available,
    editingHeaderFooter: !!storyLocation.value,
    editingAuxiliaryContent: !!storyLocation.value || !!editingHost.value,
    headerFooter: storyLocation.value && {
      visible: true, area: storyLocation.value.area, disabled: !canEditStory.value, sectionDisabled: !canEditHeaderFooterSection.value,
      refreshDisabled: storyComposing.value || !canRefreshFields.value,
      requestScope: storyActionVersion.value,
      contentDisabled: storyLinkedToPrevious.value, canLinkToPrevious: headerFooterSectionId.value !== 'first',
      linkedToPrevious: storyLinkedToPrevious.value, differentFirstPage: storySection.value.differentFirstPage,
      differentOddEven: storySection.value.differentOddEven, marginHeader: activeSectionSetup.value.marginHeader ?? 48,
      marginFooter: activeSectionSetup.value.marginFooter ?? 48,
    },
    state: toolbarState.value,
    canUndo: !!props.canUndo,
    canRedo: !!props.canRedo,
    outline: showOutline.value,
    focus: isFocusMode.value,
    fullscreen: isFullscreen.value,
    paper: paperSize.value,
    orientation: orientation.value,
    zoom: zoom.value,
    contentControlsAvailable: !!props.documentApi,
    selectedContentControl: control ? {
      id: control.id,
      type: control.type,
      lock: control.lock,
      placeholder: control.placeholder,
      showingPlaceholder: control.showingPlaceholder,
      appearance: control.appearance,
      multiLine: control.multiLine,
      temporary: control.temporary,
      checkboxStyle: contentControlCheckboxStyle(control),
      dateFormat: control.dateFormat,
      parentId: control.parentId,
      placement: control.placement,
      canSelectParent: !!control.parentId && !!model.value.contentControls?.some(candidate => candidate.id === control.parentId),
      canResetPlaceholder: !!control.placeholder && !control.showingPlaceholder && !['contentLocked', 'sdtContentLocked'].includes(control.lock ?? ''),
      canDuplicateItem: control.type === 'repeatingSectionItem'
        && !['row', 'cell', 'table'].includes(control.placement ?? ''),
      canDeleteItem: control.type === 'repeatingSectionItem'
        && !['cell', 'table'].includes(control.placement ?? ''),
    } : undefined,
    commentsPanelOpen: commentsPanelOpen.value,
    canCreateComment: canCreateComment.value,
    tableOfContents: { available: !editingHost.value && !legacyStoryActive.value && supportsDocumentTableOfContents(model.value), insertDisabled: !canInsertToc.value, selectedId: selectedTocId.value, editDisabled: !canEditSelectedToc.value,
      updateDisabled: !tocContextActionStates.value.update.available, updateReason: tocContextActionStates.value.update.reason,
      deleteDisabled: !tocContextActionStates.value.delete.available, deleteReason: tocContextActionStates.value.delete.reason },
  }, createRibbonTranslator(officeLocale.value, 'doc'))
})
let resizeObserver: ResizeObserver | undefined
let imageRequest = 0
let clipboardRequest = 0
onMounted(() => {
  document.addEventListener('fullscreenchange', syncFullscreen)
  const api = props.documentApi
  if (api) applyInputState(api, { unitId: api.id, state: api.getInputState() })
  if (!workbenchRef.value) return
  resizeObserver = new ResizeObserver(([entry]) => { if (entry) workbenchWidth.value = entry.contentRect.width })
  resizeObserver.observe(workbenchRef.value)
})
onActivated(() => {
  inputStateActive = true
  storyActionVersion.value++
  const api = props.documentApi
  if (api) applyInputState(api, { unitId: api.id, state: api.getInputState() })
  refreshDocumentFieldsOnOpen()
})
onUnmounted(() => { inputStateActive = false; inputStateEpoch++; storyActionVersion.value++; document.removeEventListener('fullscreenchange', syncFullscreen); resizeObserver?.disconnect(); clearInputMarks(); imageRequest++; clipboardRequest++ })
onDeactivated(() => {
  inputStateActive = false
  inputStateEpoch++
  storyActionVersion.value++
  painter.value = undefined
  clearInputMarks()
  closeStory(false)
  imageRequest++
  clipboardRequest++
  modalOpen.value = false
  contentControlPanelOpen.value = false
  pageSettingsOpen.value = false
  paragraphSettingsPanelOpen.value = false
  sectionSettingsPanelOpen.value = false
  commentsPanelOpen.value = false
  latexPanelOpen.value = false
  showFind.value = false
  if (document.fullscreenElement === workbenchRef.value) {
    void document.exitFullscreen().catch((error: unknown) => notify(error instanceof Error ? error.message : '无法退出全屏'))
  }
})
const previousTabBeforeTable = ref('home')
watch(() => toolbarState.value.table?.blockId, (id, previous) => {
  if (id && id !== previous) { if (activeTab.value !== 'table') previousTabBeforeTable.value = activeTab.value; activeTab.value = 'table' }
  else if (!id && activeTab.value === 'table') activeTab.value = previousTabBeforeTable.value === 'table' ? 'home' : previousTabBeforeTable.value
})
watch(selectedContentControlId, (id) => {
  if (id) activeTab.value = 'content-control'
  else if (activeTab.value === 'content-control') activeTab.value = 'home'
})

watch(() => props.fileName, (value) => { if (value) documentTitle.value = value })
watch(() => props.documentApi, (next, previous, onCleanup) => {
  inputStateEpoch++
  if (previous && previous !== next) canvasRef.value?.cancelInput(false)
  if (previous && previous.id !== next?.id && previous.getInputState()) previous.clearInputMarks()
  standalonePendingMarks.value = {}
  if (!next) return
  const subscription = next.onInputStateChanged(event => applyInputState(next, event))
  onCleanup(() => { inputStateEpoch++; subscription.dispose() })
  void nextTick(() => {
    if (inputStateActive && next === props.documentApi) applyInputState(next, { unitId: next.id, state: next.getInputState() })
  })
}, { immediate: true, flush: 'sync' })
watch(() => props.modelValue, (value) => {
  if (value == null) return
  // A parent history undo/redo can legitimately return to a previously
  // emitted snapshot. Skip only when the local model is already at that
  // snapshot; comparing against lastEmitted alone would leave the canvas on
  // the undo state and make redo appear to do nothing.
  if (typeof value !== 'string' && JSON.stringify(value) === JSON.stringify(model.value)) return
  const next = typeof value === 'string' ? { ...createBlankDocument(), blocks: blocksFromHtml(value) } : value
  inputStateEpoch++
  if (modelReceived && !props.documentApi) clearInputMarks()
  modelReceived = true
  if (typeof value === 'string' || next.id !== model.value.id) painter.value = undefined
  const sameDocument = next.id === model.value.id
  if (!sameDocument) { canvasRef.value?.cancelInput(false); editingHost.value = undefined; storyLocation.value = undefined; returnBodySelection.value = undefined }
  model.value = normalizeDocument(next)
  documentTitle.value = next.name
  if (editingHost.value) {
    try { resolveDocumentBlockHost(model.value, editingHost.value) }
    catch (error) { if (error instanceof RangeError) closeStory(false); else throw error }
  }
  selection.value = sanitizeDocumentInputSelection(editingModel.value, sameDocument ? selection.value : { anchor: selection.value.anchor, focus: selection.value.focus })
  if (returnBodySelection.value) returnBodySelection.value = sanitizeDocumentInputSelection(model.value, returnBodySelection.value)
  if (!sameDocument) { commentDraftSelection.value = undefined; commentDraftHost.value = undefined; commentDraftReference.value = ''; searchScopeSelection.value = undefined; searchInSelection.value = false }
  paperSize.value = documentPaperSize(next.page)
  orientation.value = next.page.width > next.page.height ? 'landscape' : 'portrait'
}, { immediate: true })
watch([() => props.documentApi, () => props.fieldCacheAdapter, () => model.value,
  () => fieldCacheState.value?.modelRevision, fieldOpenHistoryGeneration, canRefreshFields], () => refreshDocumentFieldsOnOpen(), { immediate: true, flush: 'post' })
function refreshDocumentFieldsOnOpen() {
  const api = props.documentApi
  const adapter = props.fieldCacheAdapter, state = api ? undefined : fieldCacheState.value
  if (!inputStateActive || !canRefreshFields.value || documentInputComposing.value
    || api && api.id !== model.value.id || !api && (!adapter || state?.unitId !== model.value.id)) return
  const identity = api?.regionalSettings ?? adapter!.identity, unitId = model.value.id!
  const generation = fieldOpenHistoryGeneration.value
  if (generation === undefined || !reserveDocumentFieldOpenRefresh(identity, unitId, generation)) return
  const previous = applyingOpenFieldCache
  applyingOpenFieldCache = true
  let accepted = false
  try {
    accepted = execute({ id: 'doc.command.refresh-header-footer-fields', locale: officeLocale.value.replaceAll('_', '-'), timestamp: Date.now(), onOpen: true }, false)
  } finally {
    applyingOpenFieldCache = previous
    if (!accepted) releaseDocumentFieldOpenRefresh(identity, unitId, generation)
  }
}
function clearInputMarks() {
  standalonePendingMarks.value = {}
  const api = props.documentApi
  if (api?.getInputState()) api.clearInputMarks()
}
function applyInputState(api: FDocument, event: ApiEventMap['DocumentInputStateChanged']) {
  const epoch = ++inputStateEpoch
  const state = event.state
  const restoreFocus = !applyingOpenFieldCache
  const current = () => {
    if (!inputStateActive || api !== props.documentApi || api.id !== model.value.id || event.unitId !== api.id || !state?.valid) return false
    const latest = api.getInputState()
    return latest?.valid === true && latest.generation === state.generation && latest.modelRevision === state.modelRevision
      && state.modelRevision === api.getModelRevision() && epoch === inputStateEpoch
  }
  if (!current() || !state) return
  const next = api.getModel()
  applyingInputState = true
  try {
    if (JSON.stringify(next) !== JSON.stringify(model.value)) { model.value = next; documentTitle.value = next.name }
    if (state.location.kind === 'body') {
      const range = state.location.selection, host = state.location.host
      if (!sameDocumentBlockHost(host, editingHost.value) || legacyStoryActive.value || JSON.stringify(selection.value) !== JSON.stringify(range)) canvasRef.value?.cancelInput(false)
      if (host && !editingHost.value && !returnBodySelection.value) returnBodySelection.value = sanitizeDocumentInputSelection(model.value, selection.value)
      editingHost.value = host ? { ...host } : undefined
      if (host?.kind === 'story') {
        const area = host.story.toLowerCase().includes('header') ? 'header' : 'footer'
        const page = findDocumentStoryPage(layout.value, { area, sectionId: host.sectionId, story: host.story, preferredPage: storyLocation.value?.page })
        if (!storyLocation.value) previousStoryTab.value = activeTab.value
        if (page !== undefined) storyLocation.value = { page, area }
      } else {
        storyLocation.value = undefined
        if (activeTab.value === 'header-footer') activeTab.value = previousStoryTab.value
        if (!host) returnBodySelection.value = undefined
      }
      selection.value = cloneDocumentSelection(range)
      visibleBlockId.value = range.focus.blockId
      void nextTick(() => { if (restoreFocus && current() && sameDocumentBlockHost(host, editingHost.value)) canvasRef.value?.focus() })
      return
    }
  } finally { applyingInputState = false }
  if (state.location.kind !== 'story') return
  const range = state.location.range
  void nextTick(() => {
    if (!current()) return
    const area = range.story.toLowerCase().includes('header') ? 'header' : 'footer'
    const page = findDocumentStoryPage(layout.value, { area, sectionId: range.sectionId, story: range.story, preferredPage: storyLocation.value?.page })
    if (pageMode.value !== 'traditional' || page === undefined) return
    const previous = currentStoryRange.value
    if (!previous || previous.sectionId !== range.sectionId || previous.story !== range.story || previous.start !== range.start || previous.end !== range.end) canvasRef.value?.cancelInput(false)
    applyingInputState = true
    try {
      if (!storyLocation.value) previousStoryTab.value = activeTab.value
      if (!returnBodySelection.value) returnBodySelection.value = sanitizeDocumentInputSelection(model.value, selection.value)
      editingHost.value = undefined
      selection.value = sanitizeDocumentInputSelection(model.value, returnBodySelection.value)
      storyLocation.value = { page, area }
      storySelection.value = { start: range.start, end: range.end }
      activeTab.value = 'header-footer'
    } finally { applyingInputState = false }
    void nextTick(() => { if (restoreFocus && current()) canvasRef.value?.focusStory(range.start, range.end) })
  })
}
function publish(next: DocumentModel, preserveInputStyle = false) {
  if (!preserveInputStyle) clearInputMarks()
  next.name = documentTitle.value
  next.modified = true
  next.wordCount = textTargets(next).reduce((count, target) => count + runText(runsForTarget(next, target)).replace(/\s/g, '').length, 0)
  model.value = next
  emit('update:modelValue', next)
  emit('document-change', next)
}
function editingCandidate(next: DocumentModel, host = editingHost.value): DocumentModel {
  return host ? documentModelForBlockHost(next, host, true) : next
}
function publishEditingCandidate(next: DocumentModel, scoped: DocumentModel, preserveInputStyle = false, host = editingHost.value) {
  if (host) {
    if (!sameDocumentBlockHost(host, editingHost.value)) throw new Error('内容区已切换，请重新执行编辑。')
    applyDocumentBlockHostModel(next, host, scoped)
    if (host.kind === 'story') assertDocumentStoryLocksPreserved(resolveSectionHeaderFooter(model.value, host.sectionId), resolveSectionHeaderFooter(next, host.sectionId), host.story, model.value)
    if (!canDocumentPermission('edit', permissionTargetsForChangedBlocks(model.value, next))) return false
  }
  publish(next, preserveInputStyle)
  return true
}
function activateStoryBlockEditing(point?: DocumentSelection) {
  if (editingHost.value || !storyLocation.value || !storyKey.value) return
  const host: DocumentAuxiliaryBlockHost = { kind: 'story', sectionId: headerFooterSectionId.value, story: storyKey.value }
  const range = currentStoryRange.value ?? { start: 0, end: 0 }
  const scoped = documentModelForBlockHost(model.value, host)
  canvasRef.value?.cancelInput(false)
  selection.value = sanitizeDocumentInputSelection(scoped, point ?? documentBlockHostSelectionFromLegacyRange(model.value, host.sectionId, host.story, range.start, range.end))
  editingHost.value = host
  standalonePendingMarks.value = {}
  props.documentApi?.setInputSelection(selection.value, undefined, host)
}
function scheduleInputFocus(focus: () => void) {
  const epoch = inputStateEpoch, documentId = model.value.id, api = props.documentApi
  void nextTick(() => {
    if (inputStateActive && epoch === inputStateEpoch && documentId === model.value.id && api === props.documentApi) focus()
  })
}
function syncCurrentInput(api: FDocument) {
  if (!inputStateActive || api !== props.documentApi || api.id !== model.value.id) return
  inputStateEpoch++
  canvasRef.value?.cancelInput(false)
  // A commit observer can close the document. Its missing session must not
  // turn a completed edit into a failed model read or restore an old focus.
  const state = api.getInputState()
  if (!state) return
  const next = api.getModel()
  model.value = next
  documentTitle.value = next.name
  if (state?.valid) applyInputState(api, { unitId: api.id, state })
  else selection.value = sanitizeDocumentInputSelection(editingModel.value, selection.value)
}
function refreshStandaloneDocumentFieldCaches(locale: string, timestamp: number, feedback?: { selectionCurrent: boolean; modelRevision?: number }): boolean {
  const adapter = props.fieldCacheAdapter, source = model.value, epoch = inputStateEpoch
  const identity = adapter?.identity
  const state = adapter?.getState(), selected = cloneDocumentSelection(selection.value)
  if (!adapter || !state || state.unitId !== source.id || JSON.stringify(state.model) !== JSON.stringify(source)
    || !inputStateActive || documentInputComposing.value) return false
  const result = adapter.refresh({ unitId: state.unitId, expectedRevision: state.modelRevision,
    expectedHistoryGeneration: state.historyGeneration, inputGeneration: state.inputGeneration, locale, timestamp, selection: selected })
  if (!result.accepted) return false
  if (feedback) feedback.modelRevision = result.modelRevision
  const latest = adapter.getState()
  if (!result.selectionCurrent || !inputStateActive || documentInputComposing.value || props.documentApi
    || adapter !== props.fieldCacheAdapter || identity !== adapter.identity || source !== model.value || epoch !== inputStateEpoch
    || JSON.stringify(selection.value) !== JSON.stringify(selected) || !latest || latest.unitId !== state.unitId
    || latest.historyGeneration !== result.historyGeneration || latest.modelRevision !== result.modelRevision
    || JSON.stringify(latest.model) !== JSON.stringify(result.model)) {
    if (feedback) feedback.selectionCurrent = false
    return true
  }
  selection.value = result.selection
  if (result.changed) {
    model.value = result.model
    documentTitle.value = result.model.name
    emit('update:modelValue', result.model)
    emit('document-cache-change', result.model)
  }
  return true
}
function execute(command: DocumentCommand, restoreFocus = true, feedback?: { selectionCurrent: boolean; modelRevision?: number }) {
  if (!command.host && legacyStoryActive.value && documentCommandSupportsBlockHost(command)
    && command.id !== 'doc.command.format-text' && command.id !== 'doc.command.clear-formatting') activateStoryBlockEditing()
  if (editingHost.value && documentCommandSupportsBlockHost(command)) {
    if (command.host && !sameDocumentBlockHost(command.host, editingHost.value)) return false
    command = { ...command, host: { ...editingHost.value } }
  }
  if (isDocumentTableCommand(command)) {
    const state = documentInputComposing.value ? { available: false, reason: '请先完成当前输入法输入' }
      : documentTableCommandState(editingModel.value, selection.value, command, permissionContext.value)
    if (!state.available) { if (state.reason) notify(state.reason); return false }
  }
  if (command.id === 'doc.command.clear-formatting' && !clearFormattingState.value.available) return false
  const documentApi = props.documentApi
  const fieldRefresh = command.id === 'doc.command.refresh-header-footer-fields' || command.id === 'doc.command.update-selected-fields'
  const derivedFieldRefresh = (command.id === 'doc.command.refresh-header-footer-fields' || command.id === 'doc.command.update-selected-fields') && command.onOpen === true
  if (documentApi && documentApi.id !== model.value.id) { notify('文档切换尚未完成，请稍后重试。'); return false }
  if ((command.id === 'doc.command.clear-formatting' || command.id === 'doc.command.format-text') && !command.story && currentStoryRange.value) command = { ...command, story: currentStoryRange.value }
  const targets = (command.id === 'doc.command.clear-formatting' || command.id === 'doc.command.format-text') && command.story ? [{ scope: 'section' as const, id: command.story.sectionId }]
    : command.host ? permissionTargetsForSelection(editingModel.value, selection.value)
    : 'sectionId' in command ? [{ scope: 'section' as const, id: command.sectionId }] : permissionTargetsForSelection(editingModel.value, selection.value)
  if (command.id !== 'doc.command.set-document-permission-rules' && command.id !== 'doc.command.insert-table' && !isDocumentTableCommand(command) && !fieldRefresh && !canDocumentPermission('edit', targets)) return false
  try {
    if (derivedFieldRefresh && !documentApi) {
      if (command.id !== 'doc.command.refresh-header-footer-fields') return false
      return refreshStandaloneDocumentFieldCaches(command.locale, command.timestamp, feedback)
    }
    if (!documentApi && !selection.value.tableSelection && (command.id === 'doc.command.insert-text' || command.id === 'doc.command.delete')) {
      assertDocumentTextSelectionEditable(editingModel.value, selection.value)
      const next = cloneDocument(model.value), scoped = editingCandidate(next)
      const marks = command.id === 'doc.command.insert-text' ? { ...pendingMarks.value, ...command.marks } : undefined
      if (marks) validateDocumentRunPatch(marks)
      const codeSelection = command.id === 'doc.command.insert-text'
        ? editCodeBlockText(scoped, selection.value, { text: command.text, marks })
        : editCodeBlockText(scoped, selection.value, { deleteDirection: command.direction })
      if (codeSelection) {
        selection.value = codeSelection
        if (JSON.stringify(scoped) !== JSON.stringify(editingModel.value) && !publishEditingCandidate(next, scoped, command.id === 'doc.command.insert-text')) return false
        if (restoreFocus) scheduleInputFocus(() => canvasRef.value?.focus())
        return true
      }
    }
    const clearedStory = command.id === 'doc.command.clear-formatting' && command.story
      ? prepareDocumentStoryClearFormatting(model.value, command.story, permissionContext.value)?.range : undefined
    const storyFormat = command.id === 'doc.command.format-text' && command.story
      ? prepareDocumentStoryTextFormatting(model.value, command.story, documentRunFormatPatch(command.patch, command.clearFields), permissionContext.value) : undefined
    const collapsedFormat = command.id === 'doc.command.format-text'
      && (storyFormat ? storyFormat.range.start === storyFormat.range.end : !selection.value.tableSelection && selectionCollapsed(selection.value))
    if (!documentApi && command.id === 'doc.command.format-text' && collapsedFormat) {
      if (!storyFormat) executeDocumentCommand(cloneDocument(model.value), selection.value, command, permissionContext.value)
      const typingPatch = storyFormat?.typingPatch ?? selectionFontPatch(editingModel.value, selection.value, documentRunFormatPatch(command.patch, command.clearFields))
      standalonePendingMarks.value = { ...standalonePendingMarks.value, ...typingPatch }
      colorMemory.value = rememberDocumentColors(colorMemory.value, command)
      if (restoreFocus) scheduleInputFocus(() => storyFormat ? canvasRef.value?.focusStory(storyFormat.range.start, storyFormat.range.end) : canvasRef.value?.focus())
      return true
    }
    const nextCommand = !documentApi && command.id === 'doc.command.insert-text' ? { ...command, marks: { ...pendingMarks.value, ...command.marks } } : command
    if (documentApi) {
      const result = documentApi.executeCommand(nextCommand, selection.value)
      if (feedback) feedback.modelRevision = result.modelRevision
      if (result.selectionCurrent === false || documentApi !== props.documentApi || documentApi.id !== model.value.id || !inputStateActive) {
        if (feedback) feedback.selectionCurrent = false
        syncCurrentInput(documentApi)
        return true
      }
      selection.value = result.selection
      if (result.changed) {
        const next = documentApi.getModel()
        model.value = next
        documentTitle.value = next.name
        emit('update:modelValue', next)
        if (derivedFieldRefresh) emit('document-cache-change', next)
        else emit('document-change', next)
      }
    } else {
      const next = cloneDocument(model.value)
      selection.value = executeDocumentCommand(next, selection.value, nextCommand, permissionContext.value)
      if (JSON.stringify(next) !== JSON.stringify(model.value)) publish(next, command.id === 'doc.command.insert-text')
    }
    if (command.id === 'doc.command.clear-formatting') {
      clearInputMarks(); painter.value = undefined
      if (clearedStory && clearedStory.sectionId === headerFooterSectionId.value && clearedStory.story === storyKey.value) {
        storySelection.value = { start: clearedStory.start, end: clearedStory.end }
      }
    }
    if (!documentApi && command.id !== 'doc.command.insert-text' && !collapsedFormat && !fieldRefresh) standalonePendingMarks.value = {}
    if (outlineNavigationTarget.value !== selection.value.focus.blockId) outlineNavigationTarget.value = undefined
    colorMemory.value = rememberDocumentColors(colorMemory.value, command)
    if (command.id === 'doc.command.table-set-background' && typeof command.color === 'string') tableColors.value = { ...tableColors.value, background: command.color }
    if (command.id === 'doc.command.table-set-border-color' && typeof command.color === 'string') tableColors.value = { ...tableColors.value, border: command.color }
    if (restoreFocus) scheduleInputFocus(() => {
      const range = clearedStory ?? storyFormat?.range
      if (range && range.sectionId === headerFooterSectionId.value && range.story === storyKey.value) canvasRef.value?.focusStory(range.start, range.end)
      else canvasRef.value?.focus()
    })
    return true
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
    if (restoreFocus) scheduleInputFocus(() => canvasRef.value?.focus())
    return false
  }
}
function openCodeBlockPanel() { codeBlockPanelOpen.value = true }
function applyCodeBlock(data: DocumentCodeBlockData) {
  if (documentInputComposing.value || selection.value.tableSelection) { notify('请先在段落或单元格文字中定位代码块插入位置'); return }
  if (legacyStoryActive.value) activateStoryBlockEditing()
  if (!canEditContextSelection.value) return
  if (!canDocumentPermission('edit', permissionTargetsForSelection(editingModel.value, selection.value))) return
  const next = cloneDocument(model.value), scoped = editingCandidate(next)
  try {
    const existing = blockAtPoint(scoped, selection.value.focus)
    if (existing && codeBlockData(existing)) {
      updateCodeBlock(scoped, existing.id, data)
    } else {
      const block = insertCodeBlock(scoped, selection.value.focus.cellBlockId ?? selection.value.focus.blockId, data.source, data.language)
      updateCodeBlock(scoped, block.id, data)
      const target = textTargets(scoped).find(target => (target.cellBlockId ?? target.blockId) === block.id)
      if (!target) throw new RangeError('代码块插入定位失败')
      selection.value = collapsedSelection({ ...target, offset: 0 })
    }
    if (!publishEditingCandidate(next, scoped)) return
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
  }
}
function applyLatexFormula(data: { latex: string; mode: LatexFormulaMode }) {
  if (documentInputComposing.value || selection.value.tableSelection) { notify('请先在段落或单元格文字中定位公式插入位置'); return }
  if (legacyStoryActive.value) activateStoryBlockEditing()
  if (!canEditContextSelection.value) return
  if (!canDocumentPermission('edit', permissionTargetsForSelection(editingModel.value, selection.value))) return
  const next = cloneDocument(model.value), scoped = editingCandidate(next)
  try {
    if (activeLatexFormula.value) {
      if (!updateLatexFormula(scoped, activeLatexFormula.value.runId, data.latex)) return
    } else {
      const point = insertLatexFormula(scoped, selection.value, data.latex, data.mode)
      selection.value = collapsedSelection(point)
    }
    if (!publishEditingCandidate(next, scoped)) return
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
  }
}
function insertNote(kind: DocumentNoteKind) {
  if (documentInputComposing.value || storyLocation.value || editingHost.value || selection.value.tableSelection) { notify('请先在正文或单元格文字中定位脚注或尾注位置'); return }
  if (!canDocumentPermission('edit', permissionTargetsForSelection(model.value, selection.value))) return
  const next = cloneDocument(model.value)
  try {
    selection.value = insertDocumentNote(next, selection.value, kind).selection
    publish(next)
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) { notify(error instanceof Error ? error.message : String(error)) }
}
function applyNotes(kind: DocumentNoteKind, values: Record<string, string>) {
  const next = cloneDocument(model.value)
  const notes = kind === 'footnote' ? next.footnotes ?? {} : next.endnotes ?? {}
  let changed = false
  try {
    for (const [id, text] of Object.entries(values)) if (notes[id] && notes[id].text !== text) {
      const host: DocumentAuxiliaryBlockHost = { kind: 'note', noteKind: kind, noteId: id }
      const target = textTargets(documentModelForBlockHost(next, host))[0]
      if (!target || !canDocumentPermission('edit', permissionTargetsForSelection(documentModelForBlockHost(next, host), collapsedSelection({ ...target, offset: 0 })))) return
      updateDocumentNote(next, kind, id, text); changed = true
    }
    if (changed) publish(next)
  } catch (error) { notify(error instanceof Error ? error.message : String(error)) }
}
function removeNote(kind: DocumentNoteKind, id: string) {
  if (!canDocumentPermission('edit', [{ scope: 'entity', id: `${kind}:${id}` }])) return
  const next = cloneDocument(model.value)
  if (!deleteDocumentNote(next, kind, id)) return
  if (editingHost.value?.kind === 'note' && editingHost.value.noteKind === kind && editingHost.value.noteId === id) closeStory(false)
  publish(next)
}
function navigateToNote(kind: DocumentNoteKind, id: string) {
  const note = (kind === 'footnote' ? model.value.footnotes : model.value.endnotes)?.[id]
  const point = note && noteReferencePoint(model.value, note)
  if (!point) return
  closeStory(false)
  onSelection(collapsedSelection(point))
  void nextTick(() => canvasRef.value?.focus())
}
function openNoteCanvas(kind: DocumentNoteKind, id: string) {
  if (documentInputComposing.value) return
  const host: DocumentAuxiliaryBlockHost = { kind: 'note', noteKind: kind, noteId: id }
  try {
    const scoped = documentModelForBlockHost(model.value, host), target = textTargets(scoped)[0]
    if (!target) throw new RangeError('注释没有可编辑段落。')
    canvasRef.value?.cancelInput(false)
    if (!returnBodySelection.value) returnBodySelection.value = sanitizeDocumentInputSelection(model.value, selection.value)
    if (!editingHost.value && !storyLocation.value) previousStoryTab.value = activeTab.value
    clearInputMarks(); clipboardRequest++; imageRequest++
    storyLocation.value = undefined
    editingHost.value = host
    selection.value = collapsedSelection({ ...target, offset: 0 })
    notesPanelOpen.value = false
    activeTab.value = 'home'
    searchScopeSelection.value = undefined; searchInSelection.value = false
    props.documentApi?.setInputSelection(selection.value, undefined, host)
    scheduleInputFocus(() => canvasRef.value?.focus())
  } catch (error) { notify(error instanceof Error ? error.message : String(error)) }
}
function leaveCodeBlock(blockId: string) {
  if (!canDocumentPermission('edit', permissionTargetsForSelection(editingModel.value, selection.value))) return
  const next = cloneDocument(model.value), scoped = editingCandidate(next)
  const paragraph = exitCodeBlock(scoped, blockId)
  if (!paragraph) return
  const target = textTargets(scoped).find(target => (target.cellBlockId ?? target.blockId) === paragraph.id)
  if (!target) return
  selection.value = collapsedSelection({ ...target, offset: 0 })
  publishEditingCandidate(next, scoped)
}
function onSelection(value: DocumentSelection) {
  activeContentControlId.value = undefined
  standalonePendingMarks.value = {}
  selection.value = value
  visibleBlockId.value = value.focus.blockId
  if (outlineNavigationTarget.value !== value.focus.blockId) outlineNavigationTarget.value = undefined
  props.documentApi?.setInputSelection(value, undefined, editingHost.value)
}
function contentControlRequestCurrent(request: DocumentContentControlSelectionRequest, requireSession = false): boolean {
  return documentContentControlInteractionCurrent(request, {
    source: model.value, host: editingHost.value, session: captureInputSession(), requireSession: requireSession && !!props.documentApi,
    disabled: !inputStateActive || documentInputComposing.value || legacyStoryActive.value,
  })
}
function onContentControlSelection(request: DocumentContentControlSelectionRequest) {
  if (!contentControlRequestCurrent(request)) return
  onSelection(cloneDocumentSelection(request.selection))
  activeContentControlId.value = request.controlId
}
function onContentControlCommand(request: DocumentContentControlCommandRequest) {
  const api = props.documentApi
  if (!api || request.command.params.controlId !== request.controlId || !contentControlRequestCurrent(request, true)) return
  const state = contentControlValueCommandState(model.value, request.command, { host: editingHost.value, permission: permissionContext.value })
  if (!state.available) { notify(state.reason ?? '无法修改此内容控件。'); return }
  const host = editingHost.value && { ...editingHost.value }, sourceId = model.value.id
  try {
    onSelection(cloneDocumentSelection(request.selection))
    const session = captureInputSession()
    if (!session?.valid) return
    const changed = api.executeContentControlCommand(request.command.id, { ...request.command.params, host,
      expectedRevision: session.modelRevision, inputGeneration: session.generation })
    if (!inputStateActive || api !== props.documentApi || sourceId !== model.value.id || !sameDocumentBlockHost(host, editingHost.value)) return
    const latest = api.getInputState()
    // General history publication invalidates the captured session without
    // changing its generation. An observer selection/edit must keep its caret.
    if (api.getModelRevision() !== session.modelRevision + (changed ? 1 : 0) || latest?.generation !== session.generation) return
    const control = editingModel.value.contentControls?.find(control => control.id === request.controlId)
    const target = control && textTargets(editingModel.value).find(target => (target.cellBlockId ?? target.blockId) === control.blockId)
    if (control && target) {
      onSelection({ anchor: { ...target, offset: control.startOffset }, focus: { ...target, offset: control.endOffset } })
      activeContentControlId.value = control.id
    }
  } catch (error) { notify(error instanceof Error ? error.message : String(error)) }
}
function toggleFormat(field: 'bold' | 'italic' | 'underline', source?: DocumentModel, story?: DocumentStoryRange) {
  if (source || story) {
    const api = props.documentApi, scope = storyActionVersion.value
    const current = () => {
      const range = currentStoryRange.value
      return inputStateActive && source === model.value && api === props.documentApi && (!api || api.id === model.value.id) && !!story && !!range
        && range.sectionId === story.sectionId && range.story === story.story && range.start === story.start && range.end === story.end
    }
    if (!source || !story || !current()) return
    const input = captureInputSession(source)
    if (!current() || scope !== storyActionVersion.value || api && (!input?.valid || input.location.kind !== 'story'
      || input.location.range.sectionId !== story.sectionId || input.location.range.story !== story.story
      || input.location.range.start !== story.start || input.location.range.end !== story.end)) {
      if (api) syncCurrentInput(api)
      return
    }
  }
  execute({ id: 'doc.command.format-text', ...(story ? { story } : {}), patch: { [field]: currentMarks.value[field] !== true } })
}
function undo() {
  if (!inputStateActive || props.documentApi && props.documentApi.id !== model.value.id) return
  emit('undo', model.value.id)
  scheduleInputFocus(() => canvasRef.value?.focus())
}
function redo() {
  if (!inputStateActive || props.documentApi && props.documentApi.id !== model.value.id) return
  emit('redo', model.value.id)
  scheduleInputFocus(() => canvasRef.value?.focus())
}
function changePageSetup() {
  applyPageSetup(documentOrientedPage(model.value.page, orientation.value, paperSize.value))
}
function applyPageSetup(page: DocumentPageSetup, mode: DocumentPageMode = pageMode.value) {
  if (!canDocumentPermission('edit')) return
  const error = documentPageSettingsError(page)
  if (error) { notify(error); return }
  const next = cloneDocument(model.value)
  next.page = { ...page }
  next.metadata = { ...next.metadata, documentFlavor: mode === 'modern' ? 2 : 0 }
  paperSize.value = documentPaperSize(page)
  orientation.value = page.width > page.height ? 'landscape' : 'portrait'
  publish(next)
  void nextTick(() => canvasRef.value?.focus())
}
function applyHeaderFooter(patch: SectionHeaderFooterPatch, input?: DocumentStoryInputCommit, sectionId = headerFooterSectionId.value) {
  const feedback: { selectionCurrent: boolean; modelRevision?: number } = { selectionCurrent: true }
  const accepted = execute({ id: 'doc.command.set-section-header-footer', sectionId, patch, ...(input ? { input } : {}) }, false, feedback)
  return { accepted, selectionCurrent: feedback.selectionCurrent, modelRevision: feedback.modelRevision }
}
function applyParagraphSettings(patch: Partial<DocumentParagraphStyle>) {
  execute({ id: 'doc.command.format-paragraph', patch }, false)
}
function setParagraphSettingsPanelOpen(open: boolean) {
  paragraphSettingsPanelOpen.value = open
  if (open) commentsPanelOpen.value = false
  else void nextTick(() => canvasRef.value?.focus())
}
function resizeDocumentDock(sizes: number[]) {
  if (documentDockOpen.value && sizes[1] !== undefined && Number.isFinite(sizes[1])) documentDockWidth.value = Math.max(280, Math.min(800, sizes[1]))
}
function applySectionSettings(sectionId: string, patch: Partial<ReturnType<typeof documentSectionSetup>>) {
  execute({ id: 'doc.command.set-section-setup', sectionId, patch }, false)
}
function selectSection(sectionId: string) {
  if (!documentSectionIds(model.value).includes(sectionId)) return
  closeStory(false)
  const targets = textTargets(model.value).filter(target => sectionIdForBlock(model.value, target.blockId) === sectionId)
  const target = targets.at(-1)
  if (target) {
    const offset = runText(runsForTarget(model.value, target)).length
    onSelection(collapsedSelection({ ...target, offset }))
    canvasRef.value?.scrollToBlock(target.blockId)
  }
}
function insertContentControl(type: DocumentContentControlType) {
  if (!props.documentApi) return
  if (legacyStoryActive.value) activateStoryBlockEditing()
  if (!canEditContextSelection.value || documentInputComposing.value) return
  if (selection.value.tableSelection) { notify('请先在单元格正文中选择内容控件的插入范围'); return }
  const anchor = selection.value.anchor
  const focus = selection.value.focus
  const blockId = focus.cellBlockId ?? focus.blockId
  if ((anchor.cellBlockId ?? anchor.blockId) !== blockId) {
    notify('内容控件必须位于同一段落或单元格内')
    return
  }
  try {
    const text = '在此处输入'
    const typeDefaults: Partial<Record<DocumentContentControlType, Record<string, unknown>>> = {
      richText: { placeholder: text, showingPlaceholder: true },
      text: { placeholder: text, showingPlaceholder: true, multiLine: false },
      checkbox: { checked: false, checkbox: { checked: false, ...CONTENT_CONTROL_CHECKBOX_STYLES.Cross }, metadata: { checkboxStyle: 'Cross' } },
      date: { placeholder: '选择日期', showingPlaceholder: true, dateFormat: CONTENT_CONTROL_DATE_FORMATS[0], dateCalendar: 'gregorian', dateStorage: 'dateTime' },
      dropDownList: { placeholder: '选择一项', showingPlaceholder: true, items: [{ displayText: '选择一项', value: '选择一项' }] },
      comboBox: { placeholder: '选择或输入一项', showingPlaceholder: true, items: [{ displayText: '选择一项', value: '选择一项' }] },
      picture: { placeholder: '选择图片', showingPlaceholder: true },
      repeatingSection: { placement: 'block', repeatingSection: { doNotAllowInsertDelete: false } },
      group: {},
      docPartObj: { docPart: { gallery: 'Quick Parts', category: 'General', unique: false } },
    }
    if (type === 'repeatingSectionItem') {
      notify('重复节项由重复节控件生成，不能单独插入')
      return
    }
    const insertedControl = props.documentApi.addContentControl({
      ...(typeDefaults[type] as Partial<DocumentContentControl>),
      type,
      blockId,
      startOffset: Math.min(anchor.offset, focus.offset),
      endOffset: Math.max(anchor.offset, focus.offset),
      placement: type === 'repeatingSection' ? 'block' : focus.cellBlockId ? 'cell' : 'inline',
    })
    activeContentControlId.value = insertedControl.getId()
    if (type === 'repeatingSection') {
      const section = insertedControl.getInfo()
      activeContentControlId.value = model.value.contentControls?.find(control => control.type === 'repeatingSectionItem'
        && control.parentId === section.id)?.id
    }
    contentControlPanelOpen.value = false
    activeTab.value = 'content-control'
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
  }
}
function insertPageField(story: DocumentStoryKey, field: 'PAGE' | 'NUMPAGES') {
  const sectionId = headerFooterSectionId.value, area = story.toLowerCase().includes('header') ? 'header' : 'footer'
  const preferredPage = storyLocation.value?.page ?? activePageIndex.value
  const pageIndex = findDocumentStoryPage(layout.value, { area, sectionId, story, preferredPage })
    ?? findDocumentStoryPage(layout.value, { area, sectionId, preferredPage })
  if (pageIndex === undefined) { notify('当前节尚未完成排版，暂时无法插入页码。'); return }
  try {
    const page = layout.value.pages[pageIndex]!
    execute({ id: 'doc.command.insert-page-field', sectionId, story, ...createDocumentNumericPageField(field, page.pageNumber, layout.value.pages.length) }, false)
  } catch (error) { notify(error instanceof Error ? error.message : String(error)) }
}
function updateTitle() { emit('update:fileName', documentTitle.value); publish(cloneDocument(model.value)) }
function saveDocument() { emit('save', { fileName: documentTitle.value, html: documentHtml(model.value), model: model.value }) }
function toggleCommentsPanel() {
  commentsPanelOpen.value = !commentsPanelOpen.value
  if (commentsPanelOpen.value) paragraphSettingsPanelOpen.value = false
  else { activeCommentId.value = ''; commentDraftSelection.value = undefined; commentDraftReference.value = ''; void nextTick(() => canvasRef.value?.focus()) }
}
function startAddComment() {
  if (!canCreateComment.value) {
    const key = !canComment.value || !selectionCollapsed(selection.value) ? 'doc.comments.permission-denied' : 'doc.comments.select-text'
    const fallback = key === 'doc.comments.permission-denied' ? 'You do not have permission to comment' : 'Select text to add a comment'
    notify(createRibbonTranslator(officeLocale.value, 'doc')(key, fallback))
    return
  }
  commentDraftSelection.value = cloneDocumentSelection(selection.value)
  commentDraftHost.value = editingHost.value && { ...editingHost.value }
  commentDraftReference.value = selectedPlainText(editingModel.value, selection.value)
  activeCommentId.value = ''
  paragraphSettingsPanelOpen.value = false
  commentsPanelOpen.value = true
}
function commentActor() {
  const userId = permissionContext.value.userId
  return { ...permissionContext.value, userId, displayName: userId }
}
function commitCommentMutation(mutate: (next: DocumentModel) => void) {
  const next = cloneDocument(model.value)
  try {
    mutate(next)
    publish(next)
    return true
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
    return false
  }
}
function createComment(text: string) {
  const range = commentDraftSelection.value
  const host = commentDraftHost.value
  if (!range || !canComment.value || !sameDocumentBlockHost(host, editingHost.value)) return false
  let referenceCurrent = false
  try { referenceCurrent = selectedPlainText(editingModel.value, range) === commentDraftReference.value && documentCommentAnchorsForSelection(editingModel.value, range).length > 0 }
  catch (error) { if (!(error instanceof RangeError || error instanceof TypeError)) throw error }
  if (!referenceCurrent) {
    notify(createRibbonTranslator(officeLocale.value, 'doc')('doc.comments.selection-changed', 'The selected text changed. Select it again before commenting'))
    commentDraftSelection.value = undefined
    commentDraftReference.value = ''
    return false
  }
  let commentId = ''
  const committed = commitCommentMutation(next => {
    const scoped = host ? documentModelForBlockHost(next, host, true) : next
    commentId = createDocumentThreadComment(scoped, range, text, commentActor()).id
    if (host) applyDocumentBlockHostModel(next, host, scoped)
  })
  if (committed) {
    commentDraftSelection.value = undefined
    commentDraftHost.value = undefined
    commentDraftReference.value = ''
    activeCommentId.value = commentId
    selection.value = collapsedSelection(range.focus)
    void nextTick(() => canvasRef.value?.focus())
  }
  return committed
}
function replyToComment(commentId: string, text: string, mentions: readonly { userId: string; label: string }[] = []) {
  if (!canDocumentPermission('comment')) return false
  return commitCommentMutation(next => { replyToDocumentThreadComment(next, commentId, text, commentActor(), new Date().toISOString(), mentions) })
}
function editComment(commentId: string, messageId: string, text: string) {
  if (!canDocumentPermission('comment')) return false
  return commitCommentMutation(next => { editDocumentCommentMessage(next, commentId, messageId, text, commentActor()) })
}
function removeComment(commentId: string) {
  if (!deletableCommentIds.value.includes(commentId) || !canDocumentPermission('comment')) return
  commitCommentMutation(next => { deleteDocumentThreadComment(next, commentId, commentActor()) })
  if (!model.value.comments?.[commentId]) activeCommentId.value = ''
}
function removeCommentMessage(commentId: string, messageId: string) {
  if (!canDocumentPermission('comment')) return
  commitCommentMutation(next => { deleteDocumentCommentMessage(next, commentId, messageId, commentActor()) })
  if (!model.value.comments?.[commentId]) activeCommentId.value = ''
}
function resolveComment(commentId: string, resolved: boolean) {
  if (!canDocumentPermission('comment')) return
  commitCommentMutation(next => { setDocumentThreadCommentResolved(next, commentId, resolved, commentActor()) })
  if (resolved && activeCommentId.value === commentId) activeCommentId.value = ''
}
function selectComment(commentId: string) {
  if (!commentId) { activeCommentId.value = ''; return }
  let comment = model.value.comments?.[commentId]
  if (!comment) return
  const host = comment.anchor[0] && documentBlockHostForId(model.value, comment.anchor[0].blockId)
  if (!sameDocumentBlockHost(host, editingHost.value)) {
    if (!host) closeStory(false)
    else if (host.kind === 'note') openNoteCanvas(host.noteKind, host.noteId)
    else {
      const area = host.story.toLowerCase().includes('header') ? 'header' : 'footer'
      const page = findDocumentStoryPage(layout.value, { area, sectionId: host.sectionId, story: host.story, preferredPage: activePageIndex.value })
      if (page === undefined) return
      openStory({ page, area })
    }
    comment = editingModel.value.comments?.[commentId]
    if (!comment) return
  }
  activeCommentId.value = commentId
  const firstAnchor = comment.anchor[0]
  const lastAnchor = comment.anchor.at(-1)
  const targets = textTargets(editingModel.value)
  const firstTarget = firstAnchor && targets.find(target => (target.cellBlockId ?? target.blockId) === firstAnchor.blockId)
  const lastTarget = lastAnchor && targets.find(target => (target.cellBlockId ?? target.blockId) === lastAnchor.blockId)
  const firstBlock = firstTarget && blockAtPoint(editingModel.value, firstTarget)
  const lastBlock = lastTarget && blockAtPoint(editingModel.value, lastTarget)
  if (!firstAnchor || !lastAnchor || !firstTarget || !lastTarget || !firstBlock || !lastBlock) {
    notify(createRibbonTranslator(officeLocale.value, 'doc')('doc.comments.anchor-unavailable', 'Text anchor unavailable'))
    return
  }
  // Imported threads and command updates may leave anchors stale. Clamp before navigating.
  const startOffset = Math.max(0, Math.min(firstAnchor.startOffset, runText(firstBlock.runs).length))
  const endOffset = Math.max(0, Math.min(lastAnchor.endOffset, runText(lastBlock.runs).length))
  const fallback = { anchor: { ...firstTarget, offset: startOffset }, focus: { ...lastTarget, offset: endOffset } }
  const tableSelection = documentCommentTableSelection(editingModel.value, comment, fallback)
  onSelection(tableSelection ?? (comment.metadata?.officeTableCellAnchor ? { anchor: fallback.anchor, focus: { ...firstTarget, offset: Math.min(firstAnchor.endOffset, runText(firstBlock.runs).length) } } : fallback))
  canvasRef.value?.scrollToBlock(firstTarget.blockId)
  canvasRef.value?.focus()
}
function importFile(file: File): false { emit('import-file', file); return false }
function movePage(direction: -1 | 1) { setActivePage(Math.max(0, Math.min(pages.value.length - 1, activePageIndex.value + direction))) }
function setActivePage(index: number) { activePageIndex.value = index; canvasRef.value?.scrollToPage(index) }
function jumpToHeading(id: string) { closeStory(false); const point = { blockId: id, offset: 0 }; outlineNavigationTarget.value = id; onSelection(collapsedSelection(point)); canvasRef.value?.scrollToBlock(id); canvasRef.value?.focus() }
function onDocumentLayout(value: DocumentLayout, source: DocumentModel) {
  layout.value = value
  layoutSource.value = source
}
function onViewportChange(location: DocumentViewportLocation) {
  const target = outlineNavigationTarget.value
  if (target) {
    const heading = layout.value.headings.find((candidate) => candidate.id === target)
    // A short page may not scroll at all; keep the explicitly selected heading
    // while it is visible, including the programmatic scroll emitted by navigation.
    if (heading && heading.y >= location.top && heading.y < location.bottom) { visibleBlockId.value = target; return }
    if (location.reason !== 'scroll') return
    outlineNavigationTarget.value = undefined
  }
  visibleBlockId.value = location.firstVisibleBlockId
}
function navigateOutline(item: DocumentOutlineItem) { if (item.blockId) jumpToHeading(item.blockId); else setActivePage(0) }
function openDocumentContextMenu(payload: { x: number; y: number; source?: DocumentModel; story?: DocumentStoryRange; host?: DocumentAuxiliaryBlockHost }) {
  if (payload.source && payload.source !== model.value) return
  if (!sameDocumentBlockHost(payload.host, editingHost.value)) return
  const story = payload.story, current = currentStoryRange.value
  if (story ? !current || current.sectionId !== story.sectionId || current.story !== story.story || current.start !== story.start || current.end !== story.end : legacyStoryActive.value) return
  if (!contextMenuActions.value.length) return
  const bounds = workbenchRef.value?.getBoundingClientRect()
  if (!bounds) return
  documentContextMenu.value = {
    left: Math.max(4, Math.min(payload.x - bounds.left, bounds.width - 238)),
    top: Math.max(4, Math.min(payload.y - bounds.top, bounds.height - 300)),
    fieldRequest: { source: model.value, requestScope: storyActionVersion.value,
      ...(editingHost.value ? { host: { ...editingHost.value } } : {}),
      ...(story ? { story: { ...story } } : { bodySelection: cloneDocumentSelection(selection.value) }) },
  }
}
function runContextMenuAction(action: { id: string; commandId?: string; params?: unknown }) {
  if (action.commandId === 'doc.command.update-selected-fields') {
    const request = documentContextMenu.value?.fieldRequest, native = canvasRef.value?.storyRange()
    documentContextMenu.value = undefined
    if (!request || request.requestScope !== storyActionVersion.value
      || !sameDocumentBlockHost(request.host, editingHost.value)
      || request.story && native && (native.start !== request.story.start || native.end !== request.story.end)) return
    refreshSelectedFields(request.source, request.story, request.bodySelection)
    return
  }
  onRibbonExecute(action.commandId ?? action.id, action.params)
  documentContextMenu.value = undefined
  if (!tableProperties.value?.open) void nextTick(() => canvasRef.value?.focus())
}
function selectDocumentContextMenuItem(info: { key: string | number }) {
  const find = (items: readonly { id: string; commandId?: string; params?: unknown; disabled?: boolean; children?: RibbonItem[] }[]): { id: string; commandId?: string; params?: unknown; disabled?: boolean } | undefined => {
    for (const item of items) { if (item.id === String(info.key)) return item; const nested = item.children && find(item.children); if (nested) return nested }
    return undefined
  }
  const action = find(contextMenuActions.value)
  if (action && !action.disabled) runContextMenuAction(action)
}
function openTableProperties(params?: unknown) {
  if (documentInputComposing.value || legacyStoryActive.value || !params || typeof params !== 'object' || Array.isArray(params)) return
  const parsed = documentCommandFromRibbon('doc.command.table-set-background', { ...params, color: null })
  if (!parsed || !isDocumentTableCommand(parsed)) return
  const tableId = parsed.tableId ?? toolbarState.value.table?.blockId, block = tableId && findDocumentBlock(editingModel.value, tableId), table = block && tableForBlock(block)
  if (!tableId || !table) return
  const requestedKind = 'kind' in params ? params.kind : 'table'
  if (requestedKind !== 'table' && requestedKind !== 'row' && requestedKind !== 'column' && requestedKind !== 'cell' && requestedKind !== 'header'
    && requestedKind !== 'type' && requestedKind !== 'move' && requestedKind !== 'background' && requestedKind !== 'border') return
  const type = 'type' in params ? params.type : undefined
  tableProperties.value = { open: true, tableId, range: parsed.range ?? documentTableWholeRange(documentTableGrid(table)), kind: requestedKind,
    ...(editingHost.value ? { host: { ...editingHost.value } } : {}),
    ...(type === 'text' || type === 'number' || type === 'date' || type === 'checkbox' || type === 'dropdown' ? { type } : {}) }
}
function applyTableProperties(command: DocumentTableCommand) {
  if (!tableProperties.value || !sameDocumentBlockHost(tableProperties.value.host, editingHost.value)) return
  if (execute(command, false) && tableProperties.value) tableProperties.value = { ...tableProperties.value, open: false }
}
function closeTableProperties() { scheduleInputFocus(() => canvasRef.value?.focus()) }
function insertTableFromDialog(size: { rows: number; columns: number; options?: DocumentTableInsertOptions }) {
  execute({ ...size.options, id: 'doc.command.insert-table', rows: size.rows, cols: size.columns })
}
function openTocInsertDialog(mode: 'automatic' | 'custom') {
  if (!canInsertToc.value) return
  if (!buildDocumentTocBlocks(model.value, DEFAULT_DOCUMENT_TOC_OPTIONS).length) {
    notify(createRibbonTranslator(officeLocale.value, 'doc')('document.toc.no-headings', 'There are no headings to include in a table of contents.'))
    return
  }
  if (mode === 'automatic') {
    const translate = createRibbonTranslator(officeLocale.value, 'doc')
    const blocks = buildDocumentTocBlocks(
      model.value,
      DEFAULT_DOCUMENT_TOC_OPTIONS,
      currentTocPageNumbers.value,
      undefined,
      [],
      translate('document.toc.contents-title', 'Contents'),
    )
    execute({ id: 'doc.command.paste-blocks', blocks })
    return
  }
  tocDialogOptions.value = { ...DEFAULT_DOCUMENT_TOC_OPTIONS }
  tocDialog.value = { open: true, mode: 'insert', insertMode: 'custom', updateMode: 'pageNumbersOnly' }
}
function openTocUpdateDialog(tocId: string) {
  if (!canEditSelectedToc.value || selectedTocId.value !== tocId) return
  tocDialogOptions.value = selectedTocOptions.value
  tocDialog.value = { open: true, mode: 'update', insertMode: 'custom', tocId, updateMode: selectedTocActionStates.value.pageNumbersOnly.available ? 'pageNumbersOnly' : 'all' }
}
function applyTocDialog(payload: { options: DocumentTocOptions; updateMode?: DocumentTocUpdateMode }) {
  if (tocDialog.value.mode === 'insert') {
    if (!canInsertToc.value) return
    const translate = createRibbonTranslator(officeLocale.value, 'doc')
    const blocks = buildDocumentTocBlocks(
      model.value,
      payload.options,
      currentTocPageNumbers.value,
      undefined,
      [],
      translate('document.toc.contents-title', 'Contents'),
    )
    if (!blocks.length) {
      notify(translate('document.toc.no-headings', 'There are no headings to include in a table of contents.'))
      return
    }
    execute({ id: 'doc.command.paste-blocks', blocks })
    return
  }
  const tocId = tocDialog.value.tocId
  const updateMode = payload.updateMode ?? 'pageNumbersOnly'
  if (!tocId || selectedTocId.value !== tocId || !selectedTocActionStates.value[updateMode].available) return
  execute({
    id: 'doc.command.update-table-of-contents',
    tocId,
    mode: updateMode,
    pageNumbers: currentTocPageNumbers.value,
  })
}
function dismissTableContextMenu(event?: Event) {
  if (event?.target instanceof HTMLElement && event.target.closest('.document-context-menu')) return
  documentContextMenu.value = undefined
}
function onWorkbenchKeydown(event: KeyboardEvent) {
  if (event.key === 'Escape' && painter.value && !event.isComposing) { event.preventDefault(); painter.value = undefined; canvasRef.value?.focus(); return }
  if (event.key === 'Escape' && documentContextMenu.value) { event.preventDefault(); documentContextMenu.value = undefined; canvasRef.value?.focus() }
  else if (event.key === 'Escape' && editingHost.value?.kind === 'note' && !event.isComposing) { event.preventDefault(); closeStory() }
}
function insertImage() { pictureReplacementControlId.value = undefined; modalKind.value = 'image'; modalValue.value = ''; modalError.value = ''; modalOpen.value = true }
function insertLink() { modalKind.value = 'link'; modalValue.value = currentMarks.value.href || ''; modalError.value = ''; modalOpen.value = true }
function confirmModal() {
  const value = safeUrl(modalValue.value)
  if (!value) { modalError.value = '请输入有效的 http 或 https 地址'; return }
  if (modalKind.value === 'link') execute({ id: 'doc.command.format-text', patch: { href: value, underline: true, color: '#2563b6' } })
  else if (pictureReplacementControlId.value) replaceSelectedPicture({ src: value, width: 320, height: 0, alt: 'Image' })
  else execute({ id: 'doc.command.insert-image', image: { src: value, width: 320, height: 0, alt: '图片' } })
  modalOpen.value = false
}
function beforeImageUpload(file: File): false { void onImageFile(file); return false }
async function onImageFile(file: File) {
  const request = ++imageRequest
  const base = model.value
  const host = editingHost.value && { ...editingHost.value }
  const range = cloneDocumentSelection(selection.value)
  const source = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(reader.error); reader.readAsDataURL(file) }).catch(() => '')
  if (request !== imageRequest || !modalOpen.value || modalKind.value !== 'image') return
  if (!source) { notify('读取图片失败'); return }
  const image = new Image()
  image.onload = () => {
    if (request !== imageRequest || !modalOpen.value || modalKind.value !== 'image') return
    if (model.value !== base || !sameDocumentBlockHost(host, editingHost.value)) { notify('文档或内容区已更新，请重新插入图片'); return }
    if (!image.naturalWidth || !image.naturalHeight) { notify('图片尺寸无效'); return }
    selection.value = range
    const imageInfo = { src: source, alt: file.name, width: Math.min(500, image.naturalWidth), height: Math.min(500, image.naturalWidth) * image.naturalHeight / image.naturalWidth }
    if (pictureReplacementControlId.value) replaceSelectedPicture(imageInfo)
    else execute({ id: 'doc.command.insert-image', image: imageInfo })
    modalOpen.value = false
  }
  image.onerror = () => { if (request === imageRequest && modalOpen.value && modalKind.value === 'image') notify('无法读取该图片') }
  image.src = source
}
function replaceSelectedPicture(image: { src: string; alt: string; width: number; height: number }) {
  const controlId = pictureReplacementControlId.value
  const control = controlId && props.documentApi?.getContentControl(controlId)
  if (!control || control.getType() !== 'picture') return
  try {
    control.replacePicture(image)
    const updated = control.getInfo()
    const target = textTargets(editingModel.value).find(target => (target.cellBlockId ?? target.blockId) === updated.blockId)
    if (target) onSelection({ anchor: { ...target, offset: updated.startOffset }, focus: { ...target, offset: updated.endOffset } })
    const pictureBlockId = updated.metadata?.pictureBlockId
    if (typeof pictureBlockId === 'string') canvasRef.value?.scrollToBlock(pictureBlockId)
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
  }
  activeContentControlId.value = controlId
  pictureReplacementControlId.value = undefined
}
function currentSearchMatchIndex() {
  return searchMatches.value.findIndex((match) => JSON.stringify(pointForMatch(match)) === JSON.stringify(selection.value))
}
function navigateSearch(direction: -1 | 1) {
  const matches = searchMatches.value
  if (!matches.length) return
  const selectedIndex = currentSearchMatchIndex()
  let index = selectedIndex + direction
  if (selectedIndex < 0) {
    const targets = textTargets(editingModel.value)
    const targetIndex = targets.findIndex((target) => target.blockId === selection.value.focus.blockId && target.row === selection.value.focus.row && target.col === selection.value.focus.col && target.cellBlockId === selection.value.focus.cellBlockId)
    const isEligible = (match: typeof matches[number]) => {
      const matchTargetIndex = targets.findIndex((target) => target.blockId === match.target.blockId && target.row === match.target.row && target.col === match.target.col && target.cellBlockId === match.target.cellBlockId)
      const sameTextTarget = match.target.blockId === selection.value.focus.blockId && match.target.row === selection.value.focus.row && match.target.col === selection.value.focus.col && match.target.cellBlockId === selection.value.focus.cellBlockId
      return direction > 0
        ? matchTargetIndex > targetIndex || (sameTextTarget && match.start >= selection.value.focus.offset)
        : matchTargetIndex < targetIndex || (sameTextTarget && match.end <= selection.value.focus.offset)
    }
    if (direction > 0) index = matches.findIndex(isEligible)
    else {
      for (let candidate = matches.length - 1; candidate >= 0; candidate--) {
        if (isEligible(matches[candidate]!)) { index = candidate; break }
      }
    }
    if (index < 0) index = direction > 0 ? 0 : matches.length - 1
  }
  index = (index + matches.length) % matches.length
  const match = matches[index]!
  onSelection(pointForMatch(match))
  canvasRef.value?.scrollToBlock(match.target.blockId)
  canvasRef.value?.focus()
}
function findNext() { navigateSearch(1) }
function findPrevious() { navigateSearch(-1) }
function findLabel(key: string, fallback: string) { return createRibbonTranslator(officeLocale.value, 'doc')(`document.search.${key}`, fallback) }
function searchPermissionTargets(matches: typeof searchMatches.value) {
  const targets = matches.flatMap((match) => permissionTargetsForSelection(editingModel.value, pointForMatch(match)))
  return [...new Map(targets.map((target) => [`${target.scope}:${target.id}`, target])).values()]
}
function replaceCurrent() {
  const match = searchMatches.value[currentSearchMatchIndex()]
  if (!match) { findNext(); return }
  if (!canDocumentPermission('edit', searchPermissionTargets([match]))) return
  try {
    const next = cloneDocument(model.value), scoped = editingCandidate(next)
    const range = replaceDocumentMatches(scoped, [match], replaceText.value, permissionContext.value)
    if (!range) return
    selection.value = range
    if (!publishEditingCandidate(next, scoped)) return
    searchInSelection.value = false
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
  }
}
function replaceAll() {
  const matches = searchMatches.value
  if (!matches.length || !canDocumentPermission('edit', searchPermissionTargets(matches))) return
  try {
    const next = cloneDocument(model.value), scoped = editingCandidate(next)
    const range = replaceDocumentMatches(scoped, matches, replaceText.value, permissionContext.value)
    if (!range) return
    selection.value = range
    if (!publishEditingCandidate(next, scoped)) return
    searchInSelection.value = false
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
  }
}
async function copySelection(cut = false) {
  if (legacyStoryActive.value) return
  const targets = permissionTargetsForSelection(editingModel.value, selection.value)
  if (!canDocumentPermission('copy', targets) || (cut && !canDocumentPermission('edit', targets))) return
  const text = selectedPlainText(editingModel.value, selection.value)
  if (!text && !selection.value.tableSelection) return
  const base = model.value
  const host = editingHost.value && { ...editingHost.value }
  const range = JSON.stringify(selection.value)
  const request = ++clipboardRequest
  try {
    const data: DocumentClipboardData = selection.value.tableSelection ? documentTableClipboardData(editingModel.value, selection.value.tableSelection) : { text }
    const rich = await writeDocumentClipboard(navigator.clipboard, data)
    if (cut) {
      if (request !== clipboardRequest || model.value !== base || !sameDocumentBlockHost(host, editingHost.value) || JSON.stringify(selection.value) !== range) { notify('文档或选区已变化，内容已复制，请重新剪切'); return }
      if (data.html && !rich) { notify('已复制文本；请使用 Ctrl+X 保留表格格式并剪切'); canvasRef.value?.focus(); return }
      execute({ id: 'doc.command.insert-text', text: '' })
    } else if (data.html && !rich) notify('系统剪贴板仅支持文本；可使用 Ctrl+C 复制表格格式')
  } catch { notify(cut ? '无法访问剪贴板，内容未删除，请使用 Ctrl+X 剪切' : '无法访问剪贴板，请使用 Ctrl+C 复制') }
  canvasRef.value?.focus()
}
async function pasteSelection(plainText = false) {
  if (legacyStoryActive.value) activateStoryBlockEditing()
  if (!canEditContextSelection.value) return
  const base = model.value
  const host = editingHost.value && { ...editingHost.value }
  const range = JSON.stringify(selection.value)
  const request = ++clipboardRequest
  try {
    if (!navigator.clipboard) throw new Error('Clipboard is unavailable')
    const data = await readDocumentClipboard(navigator.clipboard, plainText)
    if (request !== clipboardRequest || model.value !== base || !sameDocumentBlockHost(host, editingHost.value) || JSON.stringify(selection.value) !== range) { notify('文档或选区已变化，请重新粘贴'); return }
    const command = documentTablePasteCommand(editingModel.value, selection.value, data, plainText)
    if (command) execute(command)
  } catch (error) { notify(error instanceof Error ? error.message : '无法读取剪贴板，请在文档中使用 Ctrl+V 或 Ctrl+Shift+V 粘贴') }
  canvasRef.value?.focus()
}
async function pasteSpecialSelection(mode: DocumentPasteMode) {
  if (legacyStoryActive.value) activateStoryBlockEditing()
  if (!canEditContextSelection.value) return
  const base = model.value
  const host = editingHost.value && { ...editingHost.value }
  const range = JSON.stringify(selection.value)
  const request = ++clipboardRequest
  try {
    if (!navigator.clipboard) throw new Error('Clipboard is unavailable')
    const data = await readDocumentClipboard(navigator.clipboard, mode === 'text')
    if (request !== clipboardRequest || model.value !== base || !sameDocumentBlockHost(host, editingHost.value) || JSON.stringify(selection.value) !== range) { notify('文档或选区已变化，请重新粘贴'); return }
    if (selection.value.tableSelection) {
      const tableCommand = documentTablePasteCommand(editingModel.value, selection.value, data, mode === 'text', mode === 'destination' ? 'destination' : 'source')
      if (tableCommand) execute(tableCommand)
      canvasRef.value?.focus()
      return
    }
    const paragraph = blockAtPoint(editingModel.value, selection.value.focus)
    if (!paragraph) { notify('无法确定目标段落格式，请重新选择插入位置。'); return }
    const command = documentPasteSpecialCommand(data, mode, {
      paragraphStyle: { ...paragraph.style },
      textStyle: { ...marksAtPoint(editingModel.value, selection.value.focus), ...pendingMarks.value },
    })
    if (command) execute(command)
  } catch (error) { notify(error instanceof Error ? error.message : '无法读取剪贴板，请在文档中使用 Ctrl+V 或 Ctrl+Shift+V 粘贴') }
  canvasRef.value?.focus()
}
function insertCharacter(text: string) { execute({ id: 'doc.command.insert-text', text }) }
function savePermissionRules(rules: DocumentPermissionRule[]) {
  if (!canManagePermissionRules.value) return
  execute({ id: 'doc.command.set-document-permission-rules', rules })
}
function setZoom(value: number) { zoom.value = clampDocumentZoom(value) }
function printDocument() {
  if (!canDocumentPermission('print')) return
  void printDocumentModel(model.value).catch((error: unknown) => notify(error instanceof Error ? error.message : String(error)))
}
function canDocumentPermission(action: 'edit' | 'copy' | 'comment' | 'print' | 'export', targets: Array<{ scope: 'section' | 'paragraph' | 'entity'; id: string }> = []): boolean {
  const decision = resolveDocumentPermission(model.value.permissionRules ?? [], {
    ...permissionContext.value,
  }, action, targets)
  if (!decision.allowed) notify(`当前用户没有${({ edit: '编辑', copy: '复制', comment: '评论', print: '打印', export: '导出' })[action]}权限。`)
  return decision.allowed
}
function syncFullscreen() { isFullscreen.value = document.fullscreenElement === workbenchRef.value }
async function toggleFullscreen() {
  const workbench = workbenchRef.value
  if (!workbench) return
  try {
    if (document.fullscreenElement === workbench) await document.exitFullscreen()
    else await workbench.requestFullscreen()
  } catch (error) {
    notify(error instanceof Error ? error.message : '无法切换全屏')
  }
}
function notify(message: string) { notice.value = message; emit('notify', message) }
const canEditStory = computed(() => (!storyLocation.value || !!storyPage.value) && resolveDocumentPermission(model.value.permissionRules ?? [], permissionContext.value, 'edit', [
  { scope: 'section', id: headerFooterSectionId.value },
  ...(storyKey.value ? [{ scope: 'entity' as const, id: `${headerFooterSectionId.value}:${storyKey.value}` }] : []),
  ...(storyLocation.value ? [{ scope: 'entity' as const, id: `${headerFooterSectionId.value}:${storyLocation.value.area}` }] : []),
]).allowed)
const canEditHeaderFooterSection = computed(() => resolveDocumentPermission(model.value.permissionRules ?? [], permissionContext.value, 'edit', [{ scope: 'section', id: headerFooterSectionId.value }]).allowed)
const editableHeaderFooterStories = computed(() => DOCUMENT_STORY_KEYS.filter(story => resolveDocumentPermission(model.value.permissionRules ?? [], permissionContext.value, 'edit', [
  { scope: 'section', id: headerFooterSectionId.value }, { scope: 'entity', id: `${headerFooterSectionId.value}:${story}` },
  { scope: 'entity', id: `${headerFooterSectionId.value}:${story.toLowerCase().includes('header') ? 'header' : 'footer'}` },
]).allowed))
function openStory(location: StoryLocation, point?: DocumentPoint) {
  const page = layout.value.pageHeaderFooters[location.page]
  if (pageMode.value !== 'traditional' || !page || page.isParityBlank) return
  painter.value = undefined
  clipboardRequest++
  imageRequest++
  if (!storyLocation.value) previousStoryTab.value = activeTab.value
  if (!returnBodySelection.value) returnBodySelection.value = sanitizeDocumentInputSelection(model.value, selection.value)
  if (storyLocation.value?.page !== location.page || storyLocation.value?.area !== location.area) storySelection.value = { start: 0, end: 0 }
  clearInputMarks()
  editingHost.value = undefined
  selection.value = sanitizeDocumentInputSelection(model.value, returnBodySelection.value)
  storyLocation.value = location
  activeTab.value = 'header-footer'
  const key = location.area === 'header' ? page.headerStory : page.footerStory
  if (key && resolveSectionHeaderFooter(model.value, page.sectionId).storyContent?.[key]?.blocks) {
    activateStoryBlockEditing(point ? collapsedSelection(point) : undefined)
  } else props.documentApi?.setInputSelection(selection.value, currentStoryRange.value)
  scheduleInputFocus(() => {
    if (storyLocation.value?.page === location.page && storyLocation.value.area === location.area) canvasRef.value?.revealStory(location)
  })
}
function closeStory(refocus = true) {
  if (!storyLocation.value && !editingHost.value) return
  const restore = returnBodySelection.value
  clearInputMarks()
  storyLocation.value = undefined
  editingHost.value = undefined
  returnBodySelection.value = undefined
  selection.value = sanitizeDocumentInputSelection(model.value, restore ?? selection.value)
  tableProperties.value = undefined
  documentContextMenu.value = undefined
  clipboardRequest++; imageRequest++
  searchScopeSelection.value = undefined; searchInSelection.value = false
  if (activeTab.value === 'header-footer' || activeTab.value === 'table' || activeTab.value === 'content-control') activeTab.value = previousStoryTab.value
  const api = props.documentApi
  if (inputStateActive && api?.getInputState()) api.setInputSelection(selection.value)
  if (refocus) scheduleInputFocus(() => { if (!storyLocation.value) canvasRef.value?.focus() })
}
function restoreStoryAfterOptions(location: StoryLocation, sectionId: string, range: { start: number; end: number }, modelRevision?: number) {
  const epoch = ++inputStateEpoch, documentId = model.value.id, source = model.value, api = props.documentApi
  const host = editingHost.value && { ...editingHost.value }, previousSelection = cloneDocumentSelection(selection.value)
  storyOptionsPendingEpoch = epoch
  void nextTick(() => {
    try {
      if (!inputStateActive || epoch !== inputStateEpoch || documentId !== model.value.id || source !== model.value || api !== props.documentApi
        || api && modelRevision !== undefined && api.getModelRevision() !== modelRevision) return
      const page = findDocumentStoryPage(layout.value, { area: location.area, sectionId, preferredPage: location.page })
      if (page === undefined) { closeStory(false); return }
      storyLocation.value = { page, area: location.area }
      storySelection.value = clampDocumentTextRange(storyValue.value ?? '', range)
      if (host?.kind === 'story') {
        const entry = layout.value.pageHeaderFooters[page], key = location.area === 'header' ? entry?.headerStory : entry?.footerStory
        if (!key) { closeStory(false); return }
        editingHost.value = { kind: 'story', sectionId, story: key }
        const scoped = documentModelForBlockHost(model.value, editingHost.value)
        const target = textTargets(scoped)[0]
        if (!target) { closeStory(false); return }
        selection.value = sanitizeDocumentInputSelection(scoped, sameDocumentBlockHost(host, editingHost.value) ? previousSelection : collapsedSelection({ ...target, offset: 0 }))
        api?.setInputSelection(selection.value, undefined, editingHost.value)
      } else {
        const story = currentStoryRange.value
        if (story) api?.setInputSelection(selection.value, story)
      }
      scheduleInputFocus(() => canvasRef.value?.revealStory({ page, area: location.area }))
    } finally { if (storyOptionsPendingEpoch === epoch) storyOptionsPendingEpoch = undefined }
  })
}
watch([pageMode, () => model.value.id, storyPage], ([mode, id, page], [, oldId, previousPage]) => {
  const location = storyLocation.value
  if (!location) return
  if (mode !== 'traditional' || id !== oldId) { closeStory(false); return }
  if (storyOptionsPendingEpoch === inputStateEpoch) return
  if (previousPage && (!page || page.sectionId !== previousPage.sectionId)) {
    const story = location.area === 'header' ? previousPage.headerStory : previousPage.footerStory
    const next = findDocumentStoryPage(layout.value, { area: location.area, sectionId: previousPage.sectionId, story, preferredPage: location.page })
    if (next === undefined) { closeStory(false); return }
    canvasRef.value?.cancelInput(false)
    storyLocation.value = { ...location, page: next }
    scheduleInputFocus(() => canvasRef.value?.revealStory({ ...location, page: next }))
  } else if (!page) closeStory(false)
})
watch(() => {
  const range = currentStoryRange.value
  return range ? `${model.value.id}\u0000${range.sectionId}\u0000${range.story}` : ''
}, (scope, previous) => {
  if (scope !== previous && !applyingInputState) standalonePendingMarks.value = {}
}, { flush: 'sync' })
watch(editingHost, (host, previous) => {
  if (sameDocumentBlockHost(host, previous)) return
  if (!applyingInputState) standalonePendingMarks.value = {}
  searchScopeSelection.value = undefined; searchInSelection.value = false
  documentContextMenu.value = undefined
  if (tableProperties.value && !sameDocumentBlockHost(tableProperties.value.host, host)) tableProperties.value = undefined
}, { flush: 'sync' })
function onStorySelection(range: { start: number; end: number }, preserveInputStyle = false) {
  storySelection.value = range
  if (preserveInputStyle) return
  standalonePendingMarks.value = {}
  const story = currentStoryRange.value
  if (story) props.documentApi?.setInputSelection(selection.value, story)
}
function captureInputSession(sourceModel?: DocumentModel): DocumentInputSession | undefined {
  const api = props.documentApi
  if (!inputStateActive || !api || api.id !== model.value.id) return undefined
  let state = api.getInputState()
  const locationCurrent = (candidate: DocumentInputSession | undefined) => {
    const story = currentStoryRange.value
    if (!candidate) return false
    if (candidate.location.kind === 'story') return !!story && candidate.location.range.sectionId === story.sectionId
      && candidate.location.range.story === story.story && candidate.location.range.start === story.start && candidate.location.range.end === story.end
    return !story && sameDocumentBlockHost(candidate.location.host, editingHost.value)
      && JSON.stringify(candidate.location.selection) === JSON.stringify(selection.value)
  }
  // Only the beginning of a real input may renew an absent or invalid session.
  // Observer reads stay passive, and an older Canvas model cannot authorize it.
  if (sourceModel !== undefined) {
    if (sourceModel !== model.value) return state ? { ...state, valid: false } : undefined
    try {
      if (JSON.stringify(api.getModel()) !== JSON.stringify(sourceModel)) return state ? { ...state, valid: false } : undefined
      if (!state?.valid || state.modelRevision !== api.getModelRevision() || !locationCurrent(state)) {
        api.setInputSelection(selection.value, currentStoryRange.value, editingHost.value)
        if (!inputStateActive || api !== props.documentApi || api.id !== model.value.id) return undefined
        state = api.getInputState()
        if (sourceModel !== model.value || JSON.stringify(api.getModel()) !== JSON.stringify(sourceModel)) return state ? { ...state, valid: false } : undefined
      }
    } catch (error) {
      notify(error instanceof Error ? error.message : String(error))
      return state ? { ...state, valid: false } : undefined
    }
  }
  return state ? { ...state, valid: state.valid && state.modelRevision === api.getModelRevision() && locationCurrent(state) } : undefined
}
function captureStoryInputCommit(caret: number, source = model.value): DocumentStoryInputCommit | undefined {
  const state = captureInputSession(source), story = currentStoryRange.value
  if (!state?.valid || state.location.kind !== 'story' || !story || state.location.range.sectionId !== story.sectionId || state.location.range.story !== story.story) return undefined
  return { range: { ...state.location.range }, caret, modelRevision: state.modelRevision, generation: state.generation }
}
function advanceStoryInput(caret: number, marks: RunPatch, input?: DocumentStoryInputCommit, modelRevision = input ? input.modelRevision + 1 : undefined): boolean {
  const story = currentStoryRange.value
  if (!story) return false
  const api = props.documentApi
  if (api) {
    if (!input || input.range.sectionId !== story.sectionId || input.range.story !== story.story) { syncCurrentInput(api); return false }
    const target = { ...input.range, start: caret, end: caret }
    const advanced = api.advanceInputMarks(selection.value, target), state = captureInputSession()
    if (!advanced || api !== props.documentApi || !state?.valid || state.modelRevision !== modelRevision || state.generation !== input.generation + 1
      || state.location.kind !== 'story' || state.location.range.sectionId !== target.sectionId || state.location.range.story !== target.story
      || state.location.range.start !== caret || state.location.range.end !== caret) { syncCurrentInput(api); return false }
    storySelection.value = { start: caret, end: caret }
    return true
  }
  const range = clampDocumentTextRange(storyValue.value ?? '', { start: caret, end: caret })
  storySelection.value = range
  standalonePendingMarks.value = { ...marks }
  return true
}
function editStory(value: string, caret: number, marks: RunPatch = pendingMarks.value, input?: DocumentStoryInputCommit, sourceRange?: { start: number; end: number }) {
  const range = input?.range ?? sourceRange
  if (!storyKey.value || value === storyValue.value && (!range || range.start === range.end)) return
  if (storyLinkedToPrevious.value || !canEditStory.value) { canvasRef.value?.focusStory(); return }
  try {
    if (props.documentApi && !input) throw new Error('页眉页脚输入会话已更改，请重新输入。')
    const result = applyHeaderFooter(editStoryBuffer(storySection.value, storyKey.value, value, caret, model.value, marks, range), input)
    if (!result.selectionCurrent) return
    if (result.accepted) {
      advanceStoryInput(caret, marks, input, result.modelRevision)
      return
    }
  }
  catch (error) { notify(error instanceof Error ? error.message : String(error)) }
  const api = props.documentApi
  if (api) { syncCurrentInput(api); return }
  if (range) storySelection.value = clampDocumentTextRange(storyValue.value ?? '', range)
  scheduleInputFocus(() => canvasRef.value?.focusStory(currentStorySelection.value.start, currentStorySelection.value.end))
}
function handleStoryAction(raw: unknown) {
  const action = parseHeaderFooterAction(raw), currentLocation = storyLocation.value, key = storyKey.value
  if (!inputStateActive || !currentLocation || !key || !action || !headerFooterRequestScopeMatches({ requestScope: storyActionVersion.value }, action)) return
  if (editingHost.value?.kind === 'story') { handleBlockStoryAction(action, currentLocation, key); return }
  const location = { ...currentLocation }, source = model.value, api = props.documentApi, sectionId = headerFooterSectionId.value, captureScope = storyActionVersion.value
  const targetCurrent = () => inputStateActive && api === props.documentApi && source.id === model.value.id
    && (!api || api.id === model.value.id) && sectionId === headerFooterSectionId.value && key === storyKey.value
    && location.page === storyLocation.value?.page && location.area === storyLocation.value?.area
  const sourceCurrent = () => targetCurrent() && source === model.value
  const rangeCurrent = (start: number, end: number) => {
    const range = currentStoryRange.value
    return !!range && range.start === start && range.end === end
  }
  const recoverCurrentInput = () => {
    if (!inputStateActive) return
    const currentApi = props.documentApi
    if (currentApi) syncCurrentInput(currentApi)
    else scheduleInputFocus(() => { if (storyLocation.value) canvasRef.value?.focusStory(currentStorySelection.value.start, currentStorySelection.value.end) })
  }
  if (!targetCurrent()) return
  if (action.action === 'navigate') { openStory({ page: location.page, area: action.area }); return }
  if (headerFooterActionDisabled({ disabled: !canEditStory.value, sectionDisabled: !canEditHeaderFooterSection.value, contentDisabled: storyLinkedToPrevious.value,
    refreshDisabled: storyComposing.value || !canRefreshFields.value }, action)) return
  if (action.action === 'refresh-fields') {
    const range = canvasRef.value?.storyRange() ?? currentStorySelection.value
    if (range.start !== currentStorySelection.value.start || range.end !== currentStorySelection.value.end) onStorySelection(range)
    if (!sourceCurrent() || captureScope !== storyActionVersion.value || !rangeCurrent(range.start, range.end)) { recoverCurrentInput(); return }
    const feedback: { selectionCurrent: boolean; modelRevision?: number } = { selectionCurrent: true }
    const accepted = execute({ id: 'doc.command.refresh-header-footer-fields', locale: officeLocale.value.replaceAll('_', '-'), timestamp: Date.now() }, false, feedback)
    if (!accepted || !feedback.selectionCurrent || !targetCurrent() || api && feedback.modelRevision !== undefined && feedback.modelRevision !== api.getModelRevision()) { recoverCurrentInput(); return }
    scheduleInputFocus(() => canvasRef.value?.focusStory(range.start, range.end))
    return
  }
  if (action.action === 'set-options') {
    const { marginHeader, marginFooter, linkToPrevious, ...options } = action.patch
    const range = { ...currentStorySelection.value }
    if (marginHeader !== undefined || marginFooter !== undefined) {
      const feedback: { selectionCurrent: boolean; modelRevision?: number } = { selectionCurrent: true }
      const geometry = { ...(marginHeader === undefined ? {} : { marginHeader }), ...(marginFooter === undefined ? {} : { marginFooter }) }
      if (!execute({ id: 'doc.command.set-section-setup', sectionId, patch: geometry }, false, feedback) || !feedback.selectionCurrent || !targetCurrent()
        || api && feedback.modelRevision !== undefined && feedback.modelRevision !== api.getModelRevision()) { recoverCurrentInput(); return }
    }
    const patch: SectionHeaderFooterPatch = { ...options, ...(linkToPrevious === undefined ? {} : { storyLinkToPrevious: { [key]: linkToPrevious } }) }
    if (Object.keys(patch).length) {
      const result = applyHeaderFooter(patch, undefined, sectionId)
      if (result.accepted && result.selectionCurrent && (options.differentFirstPage !== undefined || options.differentOddEven !== undefined)) restoreStoryAfterOptions(location, sectionId, range, result.modelRevision)
    }
  } else {
    const section = resolveSectionHeaderFooter(source, sectionId)
    if (sectionStoryLinkedToPrevious(section, key)) return
    const now = new Date(), locale = officeLocale.value.replaceAll('_', '-')
    if (action.action === 'insert-field' || action.action === 'insert-date-time') {
      const { start, end } = canvasRef.value?.storyRange() ?? currentStorySelection.value
      if (start !== currentStorySelection.value.start || end !== currentStorySelection.value.end) onStorySelection({ start, end })
      if (!sourceCurrent() || captureScope !== storyActionVersion.value || !rangeCurrent(start, end)) { recoverCurrentInput(); return }
      const marks = { ...pendingMarks.value }
      try {
        const date = action.action === 'insert-date-time' ? storyDateField(action.format, locale, now) : undefined
        const page = layout.value.pages[location.page]!
        const insert = action.action === 'insert-field' ? createDocumentNumericPageField(action.field, page.pageNumber, layout.value.pages.length)
          : action.automatic ? date! : date!.cachedResult
        const caret = start + (typeof insert === 'string' ? insert.length : 1)
        const input = captureStoryInputCommit(caret, source)
        if (!sourceCurrent() || captureScope !== storyActionVersion.value || !rangeCurrent(start, end)) { recoverCurrentInput(); return }
        if (api && (!input || input.range.start !== start || input.range.end !== end)) throw new Error('页眉页脚输入会话已更改，请重新输入。')
        const result = applyHeaderFooter(replaceStoryRange(section, key, start, end, insert, source, marks), input, sectionId)
        if (!result.selectionCurrent) return
        if (result.accepted) {
          if (advanceStoryInput(caret, marks, input, result.modelRevision)) scheduleInputFocus(() => canvasRef.value?.focusStory(caret))
          return
        }
      } catch (error) { notify(error instanceof Error ? error.message : String(error)) }
      const currentApi = props.documentApi
      if (currentApi) syncCurrentInput(currentApi)
      else scheduleInputFocus(() => canvasRef.value?.focusStory(start, end))
      return
    }
  }
  if (targetCurrent()) canvasRef.value?.focusStory()
}
function handleBlockStoryAction(action: NonNullable<ReturnType<typeof parseHeaderFooterAction>>, location: StoryLocation, key: DocumentStoryKey) {
  if (headerFooterActionDisabled({ disabled: !canEditStory.value, sectionDisabled: !canEditHeaderFooterSection.value, contentDisabled: storyLinkedToPrevious.value,
    refreshDisabled: documentInputComposing.value || !canRefreshFields.value }, action)) return
  const sectionId = headerFooterSectionId.value
  if (action.action === 'refresh-fields') {
    execute({ id: 'doc.command.refresh-header-footer-fields', locale: officeLocale.value.replaceAll('_', '-'), timestamp: Date.now() })
  } else if (action.action === 'set-options') {
    const { marginHeader, marginFooter, linkToPrevious, ...options } = action.patch
    if (marginHeader !== undefined || marginFooter !== undefined) {
      const geometry = { ...(marginHeader === undefined ? {} : { marginHeader }), ...(marginFooter === undefined ? {} : { marginFooter }) }
      if (!execute({ id: 'doc.command.set-section-setup', sectionId, patch: geometry }, false)) return
    }
    const patch: SectionHeaderFooterPatch = { ...options, ...(linkToPrevious === undefined ? {} : { storyLinkToPrevious: { [key]: linkToPrevious } }) }
    if (Object.keys(patch).length) {
      const result = applyHeaderFooter(patch, undefined, sectionId)
      if (result.accepted && result.selectionCurrent && (options.differentFirstPage !== undefined || options.differentOddEven !== undefined)) {
        restoreStoryAfterOptions(location, sectionId, storySelection.value, result.modelRevision)
      }
    }
  } else {
    const page = layout.value.pages[location.page]
    if (!page) return
    const date = action.action === 'insert-date-time' ? storyDateField(action.format, officeLocale.value.replaceAll('_', '-'), new Date()) : undefined
    if (action.action === 'insert-date-time' && !action.automatic) execute({ id: 'doc.command.insert-text', text: date!.cachedResult })
    else if (action.action === 'insert-field' || action.action === 'insert-date-time') {
      const field = action.action === 'insert-field' ? createDocumentNumericPageField(action.field, page.pageNumber, layout.value.pages.length) : date!
      execute({ id: 'doc.command.insert-page-field', sectionId, story: key, ...field, textStyle: { ...pendingMarks.value, ...field.textStyle } })
    }
  }
  scheduleInputFocus(() => canvasRef.value?.focus())
}
function refreshSelectedFields(source = model.value, story?: DocumentStoryRange, bodySelection?: DocumentSelection) {
  if (!inputStateActive || source !== model.value || documentInputComposing.value) return
  const current = currentStoryRange.value, api = props.documentApi
  if (story ? !current || current.sectionId !== story.sectionId || current.story !== story.story || current.start !== story.start || current.end !== story.end : legacyStoryActive.value) return
  if (bodySelection && (!sameTarget(bodySelection.anchor, selection.value.anchor) || !sameTarget(bodySelection.focus, selection.value.focus)
    || bodySelection.anchor.offset !== selection.value.anchor.offset || bodySelection.focus.offset !== selection.value.focus.offset
    || !sameDocumentTableViewSelection(bodySelection.tableSelection, selection.value.tableSelection))) return
  const feedback: { selectionCurrent: boolean; modelRevision?: number } = { selectionCurrent: true }
  if (!execute({ id: 'doc.command.update-selected-fields', locale: officeLocale.value.replaceAll('_', '-'), timestamp: Date.now(), stories: story ? [{ ...story }] : [],
    ...(!story ? { bodySelection: bodySelection ?? cloneDocumentSelection(selection.value) } : {}) }, false, feedback)) return
  if (!feedback.selectionCurrent || api !== props.documentApi || api && feedback.modelRevision !== undefined && feedback.modelRevision !== api.getModelRevision()) return
  scheduleInputFocus(() => story ? canvasRef.value?.focusStory(story.start, story.end) : canvasRef.value?.focus())
}
function activatePainter(continuous: boolean) {
  if (painter.value) {
    painter.value = continuous ? { ...painter.value, continuous } : undefined
  } else if (!legacyStoryActive.value && canEditContextSelection.value) {
    const format = captureDocumentFormat(editingModel.value, selection.value)
    if (format) painter.value = { format, continuous }
  }
  canvasRef.value?.focus()
}
function applyPainter(format?: DocumentPainterFormat) {
  if (!format || legacyStoryActive.value || !documentPainterTarget(editingModel.value, selection.value)) return
  if (execute({ id: 'doc.command.apply-format-painter', format }) && !painter.value?.continuous) painter.value = undefined
}
function onRibbonExecute(id: string, params?: unknown) {
  if (id === 'doc.operation.open-table-properties') { openTableProperties(params); return }
  if (id === 'doc.operation.toggle-table-floating-toolbar') { tableFloatingToolbarOpen.value = !tableFloatingToolbarOpen.value; return }
  if (id === 'ui.operation.activate-format-painter' || id === 'ui.operation.continuous-format-painter') { activatePainter(id === 'ui.operation.continuous-format-painter'); return }
  if (id === 'doc.operation.header-footer-control') { handleStoryAction(params); return }
  if (id === 'doc.operation.close-header-footer') { closeStory(); return }
  if (id === 'core.command.undo') { if (props.canUndo) undo(); return }
  if (id === 'core.command.redo') { if (props.canRedo) redo(); return }
  const command = documentCommandFromRibbon(id, params)
  if (command) { execute(command); return }
  switch (id) {
    case 'doc.operation.insert-table-of-contents':
      if (params && typeof params === 'object' && 'mode' in params && (params.mode === 'automatic' || params.mode === 'custom')) openTocInsertDialog(params.mode)
      break
    case 'doc.operation.update-table-of-contents':
      if (params && typeof params === 'object' && 'tocId' in params && typeof params.tocId === 'string') openTocUpdateDialog(params.tocId)
      break
    case 'docs.operation.toggle-comment-panel': toggleCommentsPanel(); break
    case 'docs.operation.start-add-comment': startAddComment(); break
    case 'doc.operation.insert-footnote': insertNote('footnote'); break
    case 'doc.operation.insert-endnote': insertNote('endnote'); break
    case 'doc.operation.open-notes': notesPanelOpen.value = true; break
    case 'doc.operation.open-permissions': permissionPanelOpen.value = true; break
    case 'doc.operation.insert-code-block': openCodeBlockPanel(); break
    case 'doc.operation.open-latex': latexPanelOpen.value = true; break
    case 'doc.command.copy': void copySelection(); break
    case 'doc.command.cut': void copySelection(true); break
    case 'doc.command.paste': void pasteSelection(); break
    case 'doc.command.paste-text': void pasteSelection(true); break
    case 'doc.command.paste-special': {
      const value = params && typeof params === 'object' ? ('value' in params ? params.value : 'mode' in params ? params.mode : undefined) : undefined
      if (value === 'source' || value === 'destination' || value === 'text') void pasteSpecialSelection(value)
      break
    }
    case 'doc.command.print': printDocument(); break
    case 'doc.command.refresh-header-footer-fields': execute({ id, locale: officeLocale.value.replaceAll('_', '-'), timestamp: Date.now() }); break
    case 'doc.command.update-selected-fields': refreshSelectedFields(model.value, currentStoryRange.value); break
    case 'doc.operation.open-table': tableInsertOpen.value = true; break
    case 'doc.operation.open-image': insertImage(); break
    case 'doc.operation.open-link': insertLink(); break
    case 'doc.operation.insert-content-control':
      if (typeof params === 'string' && ['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox', 'picture', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj'].includes(params)) insertContentControl(params as DocumentContentControlType)
      else insertContentControl('richText')
      break
    case 'doc.command.open-selected-sdt-properties':
      if (selectedContentControlId.value) contentControlPanelOpen.value = true
      break
    case 'doc.operation.select-parent-sdt': {
      const parentId = selectedContentControl.value?.parentId
      const parent = parentId && model.value.contentControls?.find(control => control.id === parentId)
      if (parent) {
        activeContentControlId.value = parent.id
        onSelection({ anchor: { blockId: parent.blockId, offset: parent.startOffset }, focus: { blockId: parent.blockId, offset: parent.endOffset } })
        activeContentControlId.value = parent.id
        canvasRef.value?.scrollToBlock(parent.blockId)
      }
      break
    }
    case 'doc.command.reset-selected-sdt-placeholder':
      if (selectedContentControl.value && props.documentApi) {
        try { props.documentApi.getContentControl(selectedContentControl.value.id)?.resetPlaceholder() }
        catch (error) { notify(error instanceof Error ? error.message : String(error)) }
      }
      break
    case 'doc.command.change-selected-sdt-picture':
      if (selectedContentControl.value?.type === 'picture') {
        pictureReplacementControlId.value = selectedContentControl.value.id
        modalKind.value = 'image'
        modalValue.value = ''
        modalError.value = ''
        modalOpen.value = true
      }
      break
    case 'doc.command.duplicate-selected-repeating-sdt-item': {
      const item = selectedContentControl.value
      if (item?.type === 'repeatingSectionItem' && props.documentApi) {
        const start = item.endOffset
        try {
          if (props.documentApi.getContentControl(item.id)?.duplicateRepeatingItem()) {
            const duplicate = model.value.contentControls?.find(control => control.type === 'repeatingSectionItem'
              && control.parentId === item.parentId && control.blockId === item.blockId && control.startOffset === start)
            if (duplicate) {
              activeContentControlId.value = duplicate.id
              onSelection({ anchor: { blockId: duplicate.blockId, offset: duplicate.startOffset }, focus: { blockId: duplicate.blockId, offset: duplicate.endOffset } })
              activeContentControlId.value = duplicate.id
            }
          }
        } catch (error) { notify(error instanceof Error ? error.message : String(error)) }
      }
      break
    }
    case 'doc.command.delete-selected-repeating-sdt-item': {
      const item = selectedContentControl.value
      if (item?.type === 'repeatingSectionItem' && props.documentApi) {
        try {
          props.documentApi.getContentControl(item.id)?.deleteRepeatingItem()
          onSelection({ anchor: { blockId: item.blockId, offset: item.startOffset }, focus: { blockId: item.blockId, offset: item.startOffset } })
        } catch (error) { notify(error instanceof Error ? error.message : String(error)) }
      }
      break
    }
    case 'doc.command.remove-selected-sdt':
      if (selectedContentControlId.value && props.documentApi) {
        try {
          props.documentApi.getContentControl(selectedContentControlId.value)?.remove()
        } catch (error) {
          notify(error instanceof Error ? error.message : String(error))
        }
      }
      break
    case 'doc.command.set-selected-sdt-checkbox-style':
      if (selectedContentControl.value && props.documentApi && params && typeof params === 'object' && 'value' in params) {
        const styles = {
          Cross: { checkedState: { value: '2612', font: 'Segoe UI Symbol' }, uncheckedState: { value: '2610', font: 'Segoe UI Symbol' } },
          Check: { checkedState: { value: '2611', font: 'Segoe UI Symbol' }, uncheckedState: { value: '2610', font: 'Segoe UI Symbol' } },
          Minimal: { checkedState: { value: '2713', font: 'Segoe UI Symbol' }, uncheckedState: { value: '25A1', font: 'Segoe UI Symbol' } },
          Round: { checkedState: { value: '25CF', font: 'Segoe UI Symbol' }, uncheckedState: { value: '25CB', font: 'Segoe UI Symbol' } },
        } as const
        const name = params.value
        if (typeof name === 'string' && name in styles && selectedContentControl.value.type === 'checkbox') {
          const style = styles[name as keyof typeof styles]
          props.documentApi.getContentControl(selectedContentControl.value.id)?.update({
            checkbox: { checked: selectedContentControl.value.checkbox?.checked ?? false, ...style },
            metadata: { ...selectedContentControl.value.metadata, checkboxStyle: name },
          })
        }
      }
      break
    case 'doc.command.set-selected-sdt-appearance':
    case 'doc.command.set-selected-sdt-lock':
    case 'doc.command.set-selected-sdt-date-format':
      if (selectedContentControl.value && props.documentApi && params && typeof params === 'object' && 'value' in params && typeof params.value === 'string') {
        const value = params.value
        const control = props.documentApi.getContentControl(selectedContentControl.value.id)
        if (id === 'doc.command.set-selected-sdt-appearance' && ['boundingBox', 'tags', 'hidden'].includes(value)) control?.update({ appearance: value })
        if (id === 'doc.command.set-selected-sdt-lock' && ['unlocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked'].includes(value)) control?.update({ lock: value as DocumentContentControl['lock'] })
        if (id === 'doc.command.set-selected-sdt-date-format' && selectedContentControl.value.type === 'date') control?.update({ dateFormat: value })
      }
      break
    case 'doc.command.toggle-selected-sdt-multiline':
      if (selectedContentControl.value?.type === 'text' && props.documentApi) props.documentApi.getContentControl(selectedContentControl.value.id)?.update({ multiLine: !selectedContentControl.value.multiLine })
      break
    case 'doc.command.toggle-selected-sdt-temporary':
      if ((selectedContentControl.value?.type === 'text' || selectedContentControl.value?.type === 'richText') && props.documentApi) props.documentApi.getContentControl(selectedContentControl.value.id)?.update({ temporary: !selectedContentControl.value.temporary })
      break
    case 'doc.command.insert-character':
      if (typeof params === 'string') insertCharacter(params)
      break
    case 'doc.operation.toggle-outline': showOutline.value = !showOutline.value; break
    case 'doc.operation.toggle-focus': isFocusMode.value = !isFocusMode.value; void nextTick(() => canvasRef.value?.focus()); break
    case 'doc.operation.toggle-fullscreen': void toggleFullscreen(); break
    case 'doc.operation.set-zoom': if (typeof params === 'number') { setZoom(params); void nextTick(() => canvasRef.value?.focus()) }; break
    case 'doc.operation.toggle-find': showFind.value = !showFind.value; if (showFind.value) void nextTick(() => workbenchRef.value?.querySelector<HTMLInputElement>('[data-document-search-input] input')?.focus()); break
    case 'doc.operation.open-page-settings': pageSettingsOpen.value = true; break
    case 'doc.operation.open-header-footer': {
      const preferredPage = layout.value.lines.find(line => line.target.blockId === selection.value.focus.blockId && line.start <= selection.value.focus.offset && line.end >= selection.value.focus.offset)?.page ?? activePageIndex.value
      const page = findDocumentStoryPage(layout.value, { area: 'header', preferredPage })
      if (page !== undefined) openStory({ page, area: 'header' })
      break
    }
    case 'sidebar.operation.doc-paragraph-setting-panel': paragraphSettingsPanelOpen.value = true; commentsPanelOpen.value = false; break
    case 'sidebar.operation.doc-section-setting-panel': sectionSettingsPanelOpen.value = true; break
    case 'doc.command.set-paper-size': if (params === 'Custom') pageSettingsOpen.value = true; else if (params === 'A4' || params === 'Letter') { paperSize.value = params; changePageSetup() }; break
    case 'doc.command.set-orientation': if (params === 'portrait' || params === 'landscape') { orientation.value = params; changePageSetup() }; break
  }
}
defineExpose({ save: saveDocument, focus: () => canvasRef.value?.focus() })
</script>

<template>
  <ConfigProvider :direction="officeDirection" :theme="WORKBENCH_CONTROL_THEME">
  <section v-bind="$attrs" ref="workbenchRef" class="document-workbench" :class="{ 'focus-mode': isFocusMode }" :style="workbenchStyle" aria-label="文档工作台" @pointerdown="dismissTableContextMenu" @keydown="onWorkbenchKeydown">
    <header v-if="showHeader && chrome.header && !isFocusMode" class="document-header">
      <FileText class="document-mark" :size="20" />
      <Input v-model:value="documentTitle" class="document-name" aria-label="文档名称" variant="borderless" @change="updateTitle" @press-enter="($event.target as HTMLInputElement).blur()" />
      <Tooltip title="打印文档"><Button type="text" aria-label="打印文档" @click="printDocument"><Printer :size="16" /></Button></Tooltip>
      <Button @click="saveDocument"><Save :size="15" />保存</Button>
      <Upload accept=".doc,.docx,.odt,.txt" :show-upload-list="false" :before-upload="importFile">
        <Button type="primary"><FileUp :size="15" />打开文档</Button>
      </Upload>
    </header>
    <WorkbenchRibbon v-if="chrome.toolbar && !isFocusMode" v-model="activeTab" class="document-ribbon" :tabs="ribbonTabs" :variant="props.ribbonVariant" label="文档功能区" @execute="onRibbonExecute">
      <template #tab-extra>
        <div v-if="showFind" class="find-replace-toolbar">
          <div class="find-row">
            <Search :size="14" />
            <Input v-model:value="searchText" data-document-search-input variant="borderless" :placeholder="findLabel('find-placeholder', 'Find')" :aria-label="findLabel('find-placeholder', 'Find')" @press-enter="findNext" @keydown.esc="showFind = false; canvasRef?.focus()" />
            <span class="match-count">{{ searchMatches.length ? `${currentSearchMatchIndex() + 1} / ${searchMatches.length}` : '0 / 0' }}</span>
            <Tooltip :title="findLabel('previous', 'Previous match')"><Button type="text" size="small" :aria-label="findLabel('previous', 'Previous match')" @mousedown.prevent @click="findPrevious"><ChevronLeft :size="14" /></Button></Tooltip>
            <Tooltip :title="findLabel('next', 'Next match')"><Button type="text" size="small" :aria-label="findLabel('next', 'Next match')" @mousedown.prevent @click="findNext"><ChevronRight :size="14" /></Button></Tooltip>
            <Tooltip :title="findLabel('close', 'Close find')"><Button type="text" size="small" :aria-label="findLabel('close', 'Close find')" @click="showFind = false; canvasRef?.focus()"><X :size="14" /></Button></Tooltip>
          </div>
          <div class="replace-row">
            <Input v-model:value="replaceText" variant="borderless" :placeholder="findLabel('replace-placeholder', 'Replace with')" :aria-label="findLabel('replace-placeholder', 'Replace with')" @press-enter="replaceCurrent" />
            <Button size="small" @click="replaceCurrent">{{ findLabel('replace', 'Replace') }}</Button>
            <Button size="small" @click="replaceAll">{{ findLabel('replace-all', 'Replace all') }}</Button>
            <Checkbox v-model:checked="searchCaseSensitive">{{ findLabel('case-sensitive', 'Match case') }}</Checkbox>
            <Checkbox v-model:checked="searchWholeWord">{{ findLabel('whole-word', 'Whole words') }}</Checkbox>
            <Checkbox v-model:checked="searchInSelection" :disabled="selectionCollapsed(selection)">{{ findLabel('selection-only', 'Selection only') }}</Checkbox>
          </div>
        </div>
      </template>
    </WorkbenchRibbon>
    <div v-if="editingHost?.kind === 'note'" class="document-note-editing" data-document-note-canvas>
      <span>{{ createRibbonTranslator(officeLocale, 'doc')(editingHost.noteKind === 'footnote' ? 'doc.notes.footnote-content' : 'doc.notes.endnote-content', editingHost.noteKind === 'footnote' ? '脚注内容' : '尾注内容') }}</span>
      <Button size="small" @click="closeStory()">{{ createRibbonTranslator(officeLocale, 'doc')('doc.notes.finish-editing', '返回正文') }}</Button>
    </div>
    <main class="editor-workspace">
      <DocumentOutline v-if="chrome.leftSidebar && showOutline && !isFocusMode && editingHost?.kind !== 'note'" :model="model" :active-id="activeOutlineId" :mode="outlineMode" :max-width="256" @navigate="navigateOutline" />
      <Splitter class="document-editor-splitter" :styles="documentSplitterStyles" :destroy-on-hidden="false" @resize="resizeDocumentDock">
        <SplitterPanel class="document-canvas-panel" :min="160" :resizable="documentDockOpen">
          <DocumentCanvasViewport ref="canvasRef" :model="model" :host="editingHost" :selection="selection" :zoom="zoom" :search-matches="searchMatches"
            :active-content-control-id="activeContentControlId" @content-control-selection="onContentControlSelection" @content-control-command="onContentControlCommand"
            :permission-context="permissionContext" :table-floating-toolbar="tableFloatingToolbarOpen" @operation="onRibbonExecute"
            @selection-committed="applyPainter(painter?.format)"
            :stories-enabled="pageMode === 'traditional'" :story="storyLocation" :story-value="storyValue" :story-selection="currentStorySelection" :story-readonly="!canEditStory || storyLinkedToPrevious" :typing-marks="pendingMarks" :capture-input-session="props.documentApi ? captureInputSession : undefined"
            @open-story="openStory" @close-story="closeStory()" @edit-story="editStory" @update:story-selection="onStorySelection" @story-composing="storyComposing = $event" @command="execute" @refresh-fields="refreshSelectedFields" @code-block-exit="leaveCodeBlock" @toggle-format="toggleFormat" @update:selection="onSelection" @navigate-to-heading="jumpToHeading" @context-menu="openDocumentContextMenu" @layout="onDocumentLayout" @active-page="activePageIndex = $event" @viewport-change="onViewportChange" @undo="undo" @redo="redo" @save="saveDocument" @notify="notify" />
        </SplitterPanel>
        <SplitterPanel :size="documentDockOpen ? documentDockWidth : 0" :min="documentDockOpen ? 280 : 0" :max="800" :resizable="documentDockOpen" :destroy-on-hidden="false">
          <aside v-if="paragraphSettingsPanelOpen" class="document-settings-dock" :aria-label="createRibbonTranslator(officeLocale, 'doc')('doc.paragraph-settings.title', 'Paragraph settings')">
            <DocumentParagraphSettingsPanel :open="paragraphSettingsPanelOpen" :style="paragraphSettingsStyle" :mixed="toolbarState.mixedParagraph" :font-size="paragraphSettingsFontSize"
              :line-metrics="paragraphSettingsLineMetrics" :selection-key="paragraphSettingsSelectionKey" :read-only="!canEditContextSelection || legacyStoryActive" @apply="applyParagraphSettings" @update:open="setParagraphSettingsPanelOpen" />
          </aside>
          <aside v-else-if="commentsPanelOpen" class="document-comments-dock" :aria-label="createRibbonTranslator(officeLocale, 'doc')('doc.comments.title', 'Comments')">
            <DocumentCommentsPanel
              :comments="model.comments ?? {}"
              :current-user-id="permissionContext.userId"
              :can-comment="canComment"
              :can-create="!!commentDraftSelection && canComment"
              :can-start-create="canCreateComment"
              :deletable-comment-ids="deletableCommentIds"
              :selected-comment-id="activeCommentId"
              :on-create-comment="createComment"
              :on-reply-comment="replyToComment"
              :on-edit-comment="editComment"
              @delete="removeComment"
              @delete-message="removeCommentMessage"
              @resolve="resolveComment"
              @select="selectComment"
              @add="startAddComment"
              @close="toggleCommentsPanel"
            />
          </aside>
        </SplitterPanel>
      </Splitter>
    </main>
    <DocumentFooter v-if="chrome.footer" :active-page="activePageIndex" :page-count="pages.length" :character-count="wordCount" :selected-characters="selectedCharacters" :zoom="zoom" :modified="model.modified" :focus-mode="isFocusMode" @navigate="movePage" @zoom="setZoom" @exit-focus="isFocusMode = false" />
    <div v-if="notice" class="document-notice" role="status"><span>{{ notice }}</span><Tooltip title="关闭提示"><Button type="text" aria-label="关闭提示" @click="notice = ''"><X :size="14" /></Button></Tooltip></div>
    <div v-if="documentContextMenu" class="document-context-menu" :aria-label="createRibbonTranslator(officeLocale, 'doc')('doc.context.menu-label', 'Document context menu')" :style="{ left: `${documentContextMenu.left}px`, top: `${documentContextMenu.top}px` }">
      <Menu mode="vertical" :selectable="false" :items="documentContextMenuItems" @click="selectDocumentContextMenuItem" />
    </div>
    <Modal v-model:open="modalOpen" :title="modalTitle" ok-text="确定" cancel-text="取消" :width="440" destroy-on-hidden @ok="confirmModal">
      <Form layout="vertical" class="insert-form" @submit.prevent>
        <FormItem :label="modalKind === 'image' ? '图片地址' : '链接地址'"><Input v-model:value="modalValue" placeholder="https://" allow-clear @press-enter="confirmModal" /></FormItem>
        <Upload v-if="modalKind === 'image'" accept="image/png,image/jpeg,image/gif,image/webp" :show-upload-list="false" :before-upload="beforeImageUpload">
          <Button type="primary"><FileUp :size="15" />选择本地图片</Button>
        </Upload>
      </Form>
      <p v-if="modalError" role="alert" class="insert-error">{{ modalError }}</p>
    </Modal>
    <DocumentContentControlPanel v-model:open="contentControlPanelOpen" :selected-control-id="selectedContentControlId" :document-api="documentApi" :read-only="!canEditContextSelection || legacyStoryActive" @error="notify" />
    <DocumentTableInsertDialog v-model:open="tableInsertOpen" @insert="insertTableFromDialog" />
    <DocumentTablePropertiesDialog v-if="tableProperties" :open="tableProperties.open" :model="editingModel" :selection="selection" :permission="permissionContext"
      :table-id="tableProperties.tableId" :range="tableProperties.range" :kind="tableProperties.kind" :type="tableProperties.type" :composing="documentInputComposing"
      :layout="layoutSource === model ? editingLayout : undefined"
      @update:open="open => { if (tableProperties) tableProperties = { ...tableProperties, open } }" @submit="applyTableProperties" @closed="closeTableProperties" />
    <DocumentPageSettings v-model:open="pageSettingsOpen" :page="model.page" :mode="pageMode" @apply="applyPageSetup" />
    <DocumentHeaderFooterPanel v-model:open="headerFooterPanelOpen" :model="model" :section-id="headerFooterSectionId" :setup="activeSectionSetup" :read-only="!canEditHeaderFooterSection" :editable-stories="editableHeaderFooterStories" @apply="applyHeaderFooter" @position="applySectionSettings(headerFooterSectionId, $event)" @insert-page-field="insertPageField" @closed="canvasRef?.focus()" />
    <DocumentSectionSettingsPanel v-model:open="sectionSettingsPanelOpen" :section-id="headerFooterSectionId" :setup="activeSectionSetup" :section-options="sectionOptions" @apply="applySectionSettings" @select-section="selectSection" />
    <DocumentCodeBlockPanel v-model:open="codeBlockPanelOpen" :initial="activeCodeBlockData" @apply="applyCodeBlock" />
    <DocumentLatexPanel v-model:open="latexPanelOpen" :initial="activeLatexFormula?.formula" @apply="applyLatexFormula" />
    <DocumentNotesPanel v-model:open="notesPanelOpen" :footnotes="model.footnotes ?? {}" :endnotes="model.endnotes ?? {}" @apply="applyNotes" @remove="removeNote" @navigate="navigateToNote" @edit="openNoteCanvas" />
    <DocumentPermissionPanel :open="permissionPanelOpen" :rules="model.permissionRules ?? []" :targets="permissionTargets" :read-only="!canManagePermissionRules" @update:open="permissionPanelOpen = $event" @save="savePermissionRules" />
    <DocumentTableOfContentsDialog
      v-model:open="tocDialog.open"
      :mode="tocDialog.mode"
      :options="tocDialogOptions"
      :update-mode="tocDialog.updateMode"
      :update-states="selectedTocActionStates"
      @apply="applyTocDialog"
    />
  </section>
  </ConfigProvider>
</template>

<style scoped>
.document-workbench { position: relative; min-width: 0; height: 100%; min-height: 0; flex: 1; display: flex; flex-direction: column; overflow: hidden; color: var(--document-text); background: var(--document-background); font: 13px/1.45 "Microsoft YaHei", "Segoe UI", sans-serif; }
.document-header { flex: 0 0 44px; display: flex; align-items: center; gap: 8px; min-width: 0; padding: 0 12px; border-bottom: 1px solid var(--document-secondary-border); background: var(--document-surface); }
.document-mark { flex-shrink: 0; color: var(--document-primary); }
.document-name { min-width: 60px; flex: 1; }
.document-header :deep(.ant-btn) { display: inline-flex; align-items: center; justify-content: center; gap: 6px; height: 28px; }
.document-ribbon { flex-shrink: 0; }
.editor-workspace { position: relative; min-width: 0; min-height: 0; flex: 1; display: flex; overflow: hidden; }
.document-editor-splitter { flex: 1; min-width: 0; min-height: 0; height: 100%; }
.document-editor-splitter :deep(.document-canvas-panel) { display: flex; min-height: 0; }
.document-editor-splitter :deep(.document-canvas-viewport) { background: var(--document-background); }
.document-settings-dock, .document-comments-dock { z-index: 30; width: 100%; height: 100%; min-height: 0; border-inline-start: 1px solid var(--document-secondary-border); background: var(--document-surface); }
.find-replace-toolbar { position: absolute; z-index: 30; top: calc(100% - 1px); inset-inline-end: 8px; width: min(620px, calc(100vw - 24px)); display: flex; flex-direction: column; gap: 6px; padding: 8px; border: 1px solid var(--document-border); border-radius: 4px; background: var(--document-surface); box-shadow: var(--document-shadow); }
.find-row, .replace-row { min-width: 0; display: flex; align-items: center; gap: 4px; }
.find-row { color: var(--document-secondary-text); }
.find-row :deep(.ant-input), .replace-row :deep(.ant-input) { min-width: 0; padding: 0 4px; color: var(--document-text); background: transparent; font: inherit; }
.find-row :deep(.ant-input) { flex: 1; }
.replace-row :deep(.ant-input) { width: 145px; }
.find-row :deep(.ant-btn) { width: 24px; height: 24px; padding: 0; display: grid; place-items: center; color: var(--document-secondary-text); }
.match-count { flex: 0 0 auto; min-width: 42px; color: var(--document-secondary-text); font-variant-numeric: tabular-nums; text-align: center; }
.replace-row :deep(.ant-checkbox-wrapper) { flex: 0 0 auto; white-space: nowrap; font-size: 12px; }
.document-notice { position: absolute; z-index: 40; bottom: 44px; inset-inline: 12px; display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 8px 12px; border: 1px solid var(--document-info-border); border-radius: 4px; color: var(--document-info-text); background: var(--document-info-background); font-size: 12px; box-shadow: var(--document-shadow); }
.document-notice :deep(.ant-btn) { display: grid; place-items: center; padding: 0; color: inherit; }
.document-context-menu { position: absolute; z-index: 60; width: 230px; overflow: auto; border-radius: 4px; }
.document-context-menu :deep(.ant-menu) { border-inline-end: 0; box-shadow: var(--document-shadow); }
.document-note-editing { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 6px 12px; border-block-end: 1px solid var(--document-border); color: var(--document-text); background: var(--document-surface); }
.insert-form { display: flex; flex-direction: column; gap: 16px; padding: 12px 0; }
.insert-form :deep(.ant-form-item) { margin-bottom: 0; }
.insert-error { color: var(--document-error); }
@media (max-width: 650px) { .find-replace-toolbar { inset-inline-end: 4px; width: calc(100vw - 16px); } .replace-row { flex-wrap: wrap; } .replace-row :deep(.ant-input) { flex: 1 1 100px; } .document-header { padding: 0 6px; } }
@media print { .document-header, .document-ribbon, .document-notice, :deep(.document-outline), :deep(.document-statusbar) { display: none !important; } .document-workbench, .editor-workspace { height: auto; overflow: visible; background: #fff; } }
</style>
