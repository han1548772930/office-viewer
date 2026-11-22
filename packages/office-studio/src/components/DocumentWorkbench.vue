<script setup lang="ts">
import { ChevronLeft, ChevronRight, FileText, FileUp, Printer, Save, Search, X } from 'lucide-vue-next'
import { Button, Checkbox, ConfigProvider, Form, FormItem, Input, Menu, Modal, Tooltip, Upload } from 'antdv-next'
import type { MenuItemType } from 'antdv-next'
import { computed, inject, nextTick, onDeactivated, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue'
import { OfficeFontListKey } from '../platform/ribbon/ribbon-fonts'
import { selectionFontPatch } from '../features/document/model/document-font'
import type { DocumentContentControl, DocumentContentControlType, DocumentModel, DocumentPageSetup, DocumentParagraphStyle, DocumentPermissionRule } from '../domain/model'
import type { FDocument } from '../api/facade'
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
import WorkbenchRibbon from '../platform/ribbon/WorkbenchRibbon.vue'
import type { RibbonVariant } from '../platform/ribbon/ribbon-schema'
import { useWorkbenchDisplay, workbenchVisibility } from '../platform/workbench-display'
import { activeOutlineIdForBlock, buildDocumentOutline, documentOutlineEnabled, documentOutlinePlacement, type DocumentOutlineItem } from '../features/document/services/document-outline'
import { createDocumentRibbon, documentCommandFromRibbon, documentTableContextMenuItems, documentTocContextMenuItems } from './document/document-ribbon'
import { clampDocumentZoom, documentCharacterCount, documentSelectionCharacterCount, documentToolbarState } from './document/document-toolbar-state'
import { documentOrientedPage, documentPageSettingsError, documentPaperSize, type DocumentPageMode, type DocumentPaperSize } from './document/document-page-settings'
import { blockAtPoint, executeDocumentCommand, type DocumentCommand } from '../features/document/commands/document-commands'
import { emptyLayout, type DocumentLayout } from '../features/document/layout/document-layout'
import type { DocumentViewportLocation } from '../features/document/layout/document-viewport'
import { blocksFromHtml, clampPoint, cloneDocument, collapsedSelection, documentHtml, marksAtPoint, normalizeDocument, runText, runsForTarget, safeUrl, selectedPlainText, selectionCollapsed, textTargets, type DocumentSelection, type RunPatch } from '../features/document/model/document-model'
import { printDocumentModel } from '../features/document/services/document-print'
import { DOCUMENT_PASTE_OPTIONS, documentPasteCommand, documentPasteSpecialCommand, readDocumentClipboard, type DocumentPasteMode } from '../features/document/services/document-clipboard'
import { documentSectionIds, resolveSectionHeaderFooter, sectionIdForBlock, type SectionHeaderFooterPatch } from '../features/document/services/section-header-footer'
import { editStoryBuffer, replaceStoryRange, storyBuffer, storyDateField, type StoryLocation } from '../features/document/services/story-editor'
import type { HeaderFooterAction } from './document/DocumentHeaderFooterRibbonControl.vue'
import { documentSectionSetup } from '../features/document/services/document-settings.ts'
import { selectedDocumentContentControl } from '../features/document/services/content-controls'
import { codeBlockData, editCodeBlockText, exitCodeBlock, insertCodeBlock, updateCodeBlock, type DocumentCodeBlockData } from '../features/document/code-blocks/code-blocks.ts'
import { getLatexFormula, insertLatexFormula, updateLatexFormula, type LatexFormulaMode } from '../features/document/latex/latex-formula.ts'
import { deleteDocumentNote, insertDocumentNote, noteReferencePoint, updateDocumentNote, type DocumentNoteKind } from '../features/document/notes/document-notes.ts'
import { createDocumentThreadComment, deleteDocumentCommentMessage, deleteDocumentThreadComment, editDocumentCommentMessage, replyToDocumentThreadComment, setDocumentThreadCommentResolved } from '../features/document/comments/thread-comments.ts'
import { canManageDocumentPermissions, permissionTargetsForSelection, resolveDocumentPermission } from '../features/document/permissions/document-permissions.ts'
import { documentPermissionTargetOptions } from '../features/document/permissions/permission-rules.ts'
import { findDocumentMatches, pointForMatch, replaceDocumentMatches } from '../features/document/services/document-search.ts'
import { buildDocumentTocBlocks, DEFAULT_DOCUMENT_TOC_OPTIONS, documentTocIdForSelection, documentTocIdAtBlock, documentTocOptionsAtBlock, documentTocPageNumbers, supportsDocumentTableOfContents, type DocumentTocOptions, type DocumentTocUpdateMode } from '../features/document/toc/document-toc.ts'
import { createRibbonTranslator, provideWorkbenchLocales } from '../i18n/office-i18n'
import { captureDocumentFormat, clearDocumentFormat, documentPainterTarget, type DocumentPainterFormat } from '../features/document/services/document-format-painter'

const props = withDefaults(defineProps<{ modelValue?: DocumentModel | string; fileName?: string; canUndo?: boolean; canRedo?: boolean; documentApi?: FDocument; showHeader?: boolean; showOutline?: boolean; permissionUserId?: string; permissionAdminUserId?: string; permissionDefaultAllow?: boolean; ribbonVariant?: RibbonVariant }>(), { showHeader: false, showOutline: undefined, permissionUserId: 'local-user', permissionDefaultAllow: true, ribbonVariant: 'classic' })
const officeLocale = provideWorkbenchLocales(() => props.documentApi?.regionalSettings)
const emit = defineEmits<{
  'update:modelValue': [value: DocumentModel]
  'update:fileName': [value: string]
  'document-change': [value: DocumentModel]
  'import-file': [file: File]
  save: [payload: { fileName: string; html: string; model: DocumentModel }]
  undo: []
  redo: []
  notify: [message: string]
}>()
const model = shallowRef<DocumentModel>(normalizeDocument(createBlankDocument()))
const firstTarget = textTargets(model.value)[0]!
const selection = shallowRef<DocumentSelection>(collapsedSelection({ ...firstTarget, offset: 0 }))
const layout = shallowRef<DocumentLayout>(emptyLayout())
const canvasRef = ref<InstanceType<typeof DocumentCanvasViewport> | null>(null)
const activeTab = ref('home')
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
    ? { anchor: { ...selection.value.anchor }, focus: { ...selection.value.focus } }
    : undefined
})
const searchMatches = computed(() => showFind.value ? findDocumentMatches(model.value, searchText.value, {
  caseSensitive: searchCaseSensitive.value,
  wholeWord: searchWholeWord.value,
  locale: officeLocale.value,
  selection: searchInSelection.value ? searchScopeSelection.value : undefined,
}) : [])
const activePageIndex = ref(0)
const pendingMarks = shallowRef<RunPatch>({})
const modalOpen = ref(false)
const modalKind = ref<'image' | 'link'>('image')
const modalValue = ref('')
const modalError = ref('')
const tableInsertOpen = ref(false)
const notice = ref('')
const documentContextMenu = ref<{ left: number; top: number }>()
const contentControlPanelOpen = ref(false)
const activeContentControlId = ref<string>()
const pictureReplacementControlId = ref<string>()
const selectedContentControlId = computed(() => {
  if (activeContentControlId.value && model.value.contentControls?.some(control => control.id === activeContentControlId.value)) return activeContentControlId.value
  const anchor = selection.value.anchor
  const focus = selection.value.focus
  const blockId = focus.cellBlockId ?? focus.blockId
  if ((anchor.cellBlockId ?? anchor.blockId) !== blockId) return undefined
  const start = Math.min(anchor.offset, focus.offset)
  const end = Math.max(anchor.offset, focus.offset)
  return selectedDocumentContentControl(model.value, blockId, start, end)?.id
})
const selectedContentControl = computed(() => selectedContentControlId.value
  ? model.value.contentControls?.find(control => control.id === selectedContentControlId.value)
  : undefined)
const pageSettingsOpen = ref(false)
const headerFooterPanelOpen = ref(false)
const storyLocation = ref<StoryLocation>()
const previousStoryTab = ref('home')
const storyPage = computed(() => storyLocation.value && layout.value.pageHeaderFooters[storyLocation.value.page])
const headerFooterSectionId = computed(() => storyPage.value?.sectionId ?? sectionIdForBlock(model.value, selection.value.focus.blockId))
const storyKey = computed(() => storyLocation.value?.area === 'footer' ? storyPage.value?.footerStory : storyPage.value?.headerStory)
const storySection = computed(() => resolveSectionHeaderFooter(model.value, headerFooterSectionId.value))
const storyValue = computed(() => storyKey.value && storyBuffer(storySection.value, storyKey.value).value)
const paragraphSettingsPanelOpen = ref(false)
const sectionSettingsPanelOpen = ref(false)
const activeParagraphBlock = computed(() => blockAtPoint(model.value, selection.value.focus))
const paragraphSettingsStyle = computed<DocumentParagraphStyle>(() => ({
  align: activeParagraphBlock.value?.style.align ?? 'left',
  indent: activeParagraphBlock.value?.style.indent ?? 0,
  lineHeight: activeParagraphBlock.value?.style.lineHeight ?? 1.5,
  spaceBefore: activeParagraphBlock.value?.style.spaceBefore ?? 0,
  spaceAfter: activeParagraphBlock.value?.style.spaceAfter ?? 8,
  ...activeParagraphBlock.value?.style,
}))
const paragraphSettingsFontSize = computed(() => activeParagraphBlock.value?.runs[0]?.fontSize ?? 11)
const activeSectionSetup = computed(() => documentSectionSetup(model.value, headerFooterSectionId.value))
const sectionOptions = computed(() => documentSectionIds(model.value).map((value, index) => ({ value, label: `#${index + 1}` })))
const pageMode = computed<DocumentPageMode>(() => documentOutlineEnabled(model.value) ? 'modern' : 'traditional')
const codeBlockPanelOpen = ref(false)
const latexPanelOpen = ref(false)
const notesPanelOpen = ref(false)
const permissionPanelOpen = ref(false)
const commentsPanelOpen = ref(false)
const activeCommentId = ref('')
const commentDraftSelection = shallowRef<DocumentSelection>()
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
const canCreateComment = computed(() => canComment.value && !selectionCollapsed(selection.value)
  && resolveDocumentPermission(model.value.permissionRules ?? [], permissionContext.value, 'comment', permissionTargetsForSelection(model.value, selection.value)).allowed)
const deletableCommentIds = computed(() => Object.values(model.value.comments ?? {})
  .filter(comment => canComment.value && comment.messages[0]?.authorId === permissionContext.value.userId)
  .map(comment => comment.id))
const activeCodeBlock = computed(() => model.value.blocks.find((block) => block.id === selection.value.focus.blockId))
const activeCodeBlockData = computed(() => activeCodeBlock.value && codeBlockData(activeCodeBlock.value))
const activeLatexFormula = computed(() => {
  const block = activeCodeBlock.value
  if (!block || selection.value.focus.row != null || selection.value.anchor.blockId !== selection.value.focus.blockId) return undefined
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
const selectedCharacters = computed(() => documentSelectionCharacterCount(model.value, selection.value))
const toolbarState = computed(() => documentToolbarState(model.value, selection.value, pendingMarks.value))
const tableContextActions = computed(() => documentTableContextMenuItems(createRibbonTranslator(officeLocale.value, 'doc')))
const selectedTocId = computed(() => supportsDocumentTableOfContents(model.value)
  ? documentTocIdForSelection(model.value, selection.value)
  : undefined)
const selectedTocOptions = computed(() => documentTocOptionsAtBlock(blockAtPoint(model.value, selection.value.focus)) ?? { ...DEFAULT_DOCUMENT_TOC_OPTIONS })
const previousTabBeforeToc = ref<string>()
const canEditContextSelection = computed(() => resolveDocumentPermission(
  model.value.permissionRules ?? [], permissionContext.value, 'edit', permissionTargetsForSelection(model.value, selection.value),
).allowed)
const canEditSelectedToc = computed(() => selectedTocId.value != null && resolveDocumentPermission(
  model.value.permissionRules ?? [], permissionContext.value, 'edit', permissionTargetsForSelection(model.value, selection.value),
).allowed)
const canInsertToc = computed(() => supportsDocumentTableOfContents(model.value)
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
  const pasteSpecialActions = DOCUMENT_PASTE_OPTIONS.map(({ value: mode, label }) => ({
    id: `doc-context-paste-special-${mode}`,
    title: translate(`doc.context.paste-special.${mode}`, label),
    commandId: 'doc.command.paste-special',
    params: { value: mode },
    disabled: !canEditContextSelection.value || typeof navigator === 'undefined' || !navigator.clipboard
      || (mode === 'destination' && !blockAtPoint(model.value, selection.value.focus)),
    group: 'paste-special' as const,
  }))
  const quickActions = [
    { id: 'doc-context-copy', title: translate('doc.context.copy', '复制'), commandId: 'doc.command.copy', params: undefined, disabled: !toolbarState.value.hasSelection || !canDocumentPermission('copy'), group: 'quick' as const },
    { id: 'doc-context-cut', title: translate('doc.context.cut', '剪切'), commandId: 'doc.command.cut', params: undefined, disabled: !toolbarState.value.hasSelection || !canDocumentPermission('copy') || !canEditContextSelection.value, group: 'quick' as const },
    { id: 'doc-context-paste', title: translate('doc.context.paste', '粘贴'), commandId: 'doc.command.paste', params: undefined, disabled: !canEditContextSelection.value || typeof navigator === 'undefined' || !navigator.clipboard, group: 'quick' as const },
    { id: 'doc-context-delete', title: translate('doc.context.delete', '删除'), commandId: 'doc.command.delete', params: { direction: -1 }, disabled: !toolbarState.value.hasSelection || !canEditContextSelection.value, group: 'format' as const },
  ]
  const paragraphSettings = {
    id: 'doc-context-paragraph-settings', title: translate('doc.context.paragraph-settings', 'Paragraph settings'),
    commandId: 'sidebar.operation.doc-paragraph-setting-panel', params: undefined,
    disabled: !canEditContextSelection.value, group: 'layout' as const,
  }
  const layoutActions = [paragraphSettings]
  if (supportsDocumentTableOfContents(model.value) && selection.value.focus.row == null) layoutActions.push({
    id: 'doc-context-section-settings', title: translate('doc.context.section-settings', 'Section settings'),
    commandId: 'sidebar.operation.doc-section-setting-panel', params: undefined,
    disabled: !canEditContextSelection.value, group: 'layout' as const,
  })
  const tableActions = toolbarState.value.table
    ? tableContextActions.value.map(action => ({ ...action, disabled: !!action.disabled || !canEditContextSelection.value, group: 'table' as const }))
    : []
  const tocActions = selectedTocId.value
    ? documentTocContextMenuItems(selectedTocId.value, !canEditSelectedToc.value, translate).map(action => ({ ...action, group: 'layout' as const }))
    : []
  return [...quickActions, ...pasteSpecialActions, ...layoutActions, ...tocActions, ...tableActions]
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
    ...(layoutItems.length ? [{ type: 'divider' as const }, ...layoutItems.map(action => ({ key: action.id, label: action.title, disabled: action.disabled }))] : []),
    ...(tableItems.length ? [{ type: 'divider' as const }, ...tableItems.map(action => ({ key: action.id, label: action.title, disabled: action.disabled }))] : []),
  ]
})
const currentMarks = computed(() => toolbarState.value.marks)
const officeFonts = inject(OfficeFontListKey, undefined)
const painter = shallowRef<{ format: DocumentPainterFormat; continuous: boolean }>()
const modalTitle = computed(() => ({ image: '插入图片', link: '插入链接' })[modalKind.value])
const ribbonTabs = computed(() => {
  const control = selectedContentControl.value
  return createDocumentRibbon({
    fontList: officeFonts?.value,
    painterActive: !!painter.value, painterUnavailable: !canEditContextSelection.value || !!storyLocation.value,
    mode: pageMode.value,
    editingDisabled: !canEditContextSelection.value || !!storyLocation.value,
    editingHeaderFooter: !!storyLocation.value,
    headerFooter: storyLocation.value && {
      visible: true, area: storyLocation.value.area, disabled: !canEditStory.value,
      contentDisabled: storySection.value.linkToPrevious, canLinkToPrevious: headerFooterSectionId.value !== 'first',
      linkedToPrevious: storySection.value.linkToPrevious, differentFirstPage: storySection.value.differentFirstPage,
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
      checkboxStyle: typeof control.metadata?.checkboxStyle === 'string' ? control.metadata.checkboxStyle : undefined,
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
    tableOfContents: { available: supportsDocumentTableOfContents(model.value), insertDisabled: !canInsertToc.value, selectedId: selectedTocId.value, editDisabled: !canEditSelectedToc.value },
  }, createRibbonTranslator(officeLocale.value, 'doc'))
})
let resizeObserver: ResizeObserver | undefined
let imageRequest = 0
let clipboardRequest = 0
onMounted(() => {
  document.addEventListener('fullscreenchange', syncFullscreen)
  if (!workbenchRef.value) return
  resizeObserver = new ResizeObserver(([entry]) => { if (entry) workbenchWidth.value = entry.contentRect.width })
  resizeObserver.observe(workbenchRef.value)
})
onUnmounted(() => { document.removeEventListener('fullscreenchange', syncFullscreen); resizeObserver?.disconnect(); imageRequest++; clipboardRequest++ })
onDeactivated(() => {
  painter.value = undefined
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
watch(() => toolbarState.value.table?.blockId, (id, previous) => { if (id && id !== previous) activeTab.value = 'table' })
watch(selectedContentControlId, (id) => {
  if (id) activeTab.value = 'content-control'
  else if (activeTab.value === 'content-control') activeTab.value = 'home'
})

watch(() => props.fileName, (value) => { if (value) documentTitle.value = value })
watch(() => props.modelValue, (value) => {
  if (value == null) return
  // A parent history undo/redo can legitimately return to a previously
  // emitted snapshot. Skip only when the local model is already at that
  // snapshot; comparing against lastEmitted alone would leave the canvas on
  // the undo state and make redo appear to do nothing.
  if (typeof value !== 'string' && JSON.stringify(value) === JSON.stringify(model.value)) return
  const next = typeof value === 'string' ? { ...createBlankDocument(), blocks: blocksFromHtml(value) } : value
  if (typeof value === 'string' || next.id !== model.value.id) painter.value = undefined
  model.value = normalizeDocument(next)
  documentTitle.value = next.name
  selection.value = { anchor: clampPoint(model.value, selection.value.anchor), focus: clampPoint(model.value, selection.value.focus) }
  paperSize.value = documentPaperSize(next.page)
  orientation.value = next.page.width > next.page.height ? 'landscape' : 'portrait'
}, { immediate: true })
function publish(next: DocumentModel) {
  next.name = documentTitle.value
  next.modified = true
  next.wordCount = textTargets(next).reduce((count, target) => count + runText(runsForTarget(next, target)).replace(/\s/g, '').length, 0)
  model.value = next
  emit('update:modelValue', next)
  emit('document-change', next)
}
function execute(command: DocumentCommand, restoreFocus = true) {
  const targets = 'sectionId' in command ? [{ scope: 'section' as const, id: command.sectionId }] : permissionTargetsForSelection(model.value, selection.value)
  if (command.id !== 'doc.command.set-document-permission-rules' && !canDocumentPermission('edit', targets)) return
  if (command.id === 'doc.command.insert-text') {
    const next = cloneDocument(model.value)
    const codeSelection = editCodeBlockText(next, selection.value, { text: command.text })
    if (codeSelection) { selection.value = codeSelection; publish(next); return }
  } else if (command.id === 'doc.command.delete') {
    const next = cloneDocument(model.value)
    const codeSelection = editCodeBlockText(next, selection.value, { deleteDirection: command.direction })
    if (codeSelection) { selection.value = codeSelection; if (JSON.stringify(next) !== JSON.stringify(model.value)) publish(next); return }
  }
  if (command.id === 'doc.command.format-text' && selectionCollapsed(selection.value)) {
    pendingMarks.value = { ...pendingMarks.value, ...selectionFontPatch(model.value, selection.value, command.patch) }
    canvasRef.value?.focus()
    return
  }
  try {
    const next = cloneDocument(model.value)
    const nextCommand = command.id === 'doc.command.insert-text' ? { ...command, marks: { ...pendingMarks.value, ...command.marks } } : command
    selection.value = executeDocumentCommand(next, selection.value, nextCommand, permissionContext.value)
    if (outlineNavigationTarget.value !== selection.value.focus.blockId) outlineNavigationTarget.value = undefined
    publish(next)
    if (restoreFocus) void nextTick(() => canvasRef.value?.focus())
    return true
  } catch (error) {
    pendingMarks.value = {}
    notify(error instanceof Error ? error.message : String(error))
    if (restoreFocus) void nextTick(() => canvasRef.value?.focus())
  }
}
function openCodeBlockPanel() { codeBlockPanelOpen.value = true }
function applyCodeBlock(data: DocumentCodeBlockData) {
  if (!canDocumentPermission('edit', permissionTargetsForSelection(model.value, selection.value))) return
  const next = cloneDocument(model.value)
  const existing = next.blocks.find((block) => block.id === selection.value.focus.blockId)
  if (existing && codeBlockData(existing)) {
    updateCodeBlock(next, existing.id, data)
  } else {
    const block = insertCodeBlock(next, selection.value.focus.blockId, data.source, data.language)
    updateCodeBlock(next, block.id, data)
    selection.value = collapsedSelection({ blockId: block.id, offset: 0 })
  }
  publish(next)
  void nextTick(() => canvasRef.value?.focus())
}
function applyLatexFormula(data: { latex: string; mode: LatexFormulaMode }) {
  if (!canDocumentPermission('edit', permissionTargetsForSelection(model.value, selection.value))) return
  const next = cloneDocument(model.value)
  try {
    if (activeLatexFormula.value) {
      if (!updateLatexFormula(next, activeLatexFormula.value.runId, data.latex)) return
    } else {
      const point = insertLatexFormula(next, selection.value, data.latex, data.mode)
      selection.value = collapsedSelection(point)
    }
    publish(next)
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
  }
}
function insertNote(kind: DocumentNoteKind) {
  if (!canDocumentPermission('edit', permissionTargetsForSelection(model.value, selection.value))) return
  const next = cloneDocument(model.value)
  try {
    selection.value = insertDocumentNote(next, selection.value, kind).selection
    publish(next)
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) { notify(error instanceof Error ? error.message : String(error)) }
}
function applyNotes(kind: DocumentNoteKind, values: Record<string, string>) {
  if (!canDocumentPermission('edit', permissionTargetsForSelection(model.value, selection.value))) return
  const next = cloneDocument(model.value)
  const notes = kind === 'footnote' ? next.footnotes ?? {} : next.endnotes ?? {}
  let changed = false
  for (const [id, text] of Object.entries(values)) if (notes[id] && notes[id].text !== text) { updateDocumentNote(next, kind, id, text); changed = true }
  if (changed) publish(next)
}
function removeNote(kind: DocumentNoteKind, id: string) {
  if (!canDocumentPermission('edit', permissionTargetsForSelection(model.value, selection.value))) return
  const next = cloneDocument(model.value)
  if (!deleteDocumentNote(next, kind, id)) return
  publish(next)
}
function navigateToNote(kind: DocumentNoteKind, id: string) {
  const note = (kind === 'footnote' ? model.value.footnotes : model.value.endnotes)?.[id]
  const point = note && noteReferencePoint(model.value, note)
  if (!point) return
  selection.value = collapsedSelection(point)
  void nextTick(() => canvasRef.value?.focus())
}
function leaveCodeBlock(blockId: string) {
  if (!canDocumentPermission('edit', [{ scope: 'paragraph', id: blockId }])) return
  const next = cloneDocument(model.value)
  const paragraph = exitCodeBlock(next, blockId)
  if (!paragraph) return
  selection.value = collapsedSelection({ blockId: paragraph.id, offset: 0 })
  publish(next)
}
function onSelection(value: DocumentSelection) { activeContentControlId.value = undefined; selection.value = value; visibleBlockId.value = value.focus.blockId; if (outlineNavigationTarget.value !== value.focus.blockId) outlineNavigationTarget.value = undefined; pendingMarks.value = {} }
function toggleFormat(field: 'bold' | 'italic' | 'underline') { execute({ id: 'doc.command.format-text', patch: { [field]: currentMarks.value[field] !== true } }) }
function undo() { pendingMarks.value = {}; emit('undo'); void nextTick(() => canvasRef.value?.focus()) }
function redo() { pendingMarks.value = {}; emit('redo'); void nextTick(() => canvasRef.value?.focus()) }
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
function applyHeaderFooter(patch: SectionHeaderFooterPatch) {
  execute({ id: 'doc.command.set-section-header-footer', sectionId: headerFooterSectionId.value, patch }, false)
}
function applyParagraphSettings(patch: Partial<DocumentParagraphStyle>) {
  execute({ id: 'doc.command.format-paragraph', patch }, false)
}
function applySectionSettings(sectionId: string, patch: Partial<ReturnType<typeof documentSectionSetup>>) {
  execute({ id: 'doc.command.set-section-setup', sectionId, patch }, false)
}
function selectSection(sectionId: string) {
  if (!documentSectionIds(model.value).includes(sectionId)) return
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
  const anchor = selection.value.anchor
  const focus = selection.value.focus
  const blockId = focus.cellBlockId ?? focus.blockId
  if ((anchor.cellBlockId ?? anchor.blockId) !== blockId) {
    notify('内容控件必须位于同一段落或单元格内')
    return
  }
  try {
    const text = '在此处输入'
    const checkboxStyles = {
      Cross: { checkedState: { value: '2612', font: 'Segoe UI Symbol' }, uncheckedState: { value: '2610', font: 'Segoe UI Symbol' } },
      Check: { checkedState: { value: '2611', font: 'Segoe UI Symbol' }, uncheckedState: { value: '2610', font: 'Segoe UI Symbol' } },
      Minimal: { checkedState: { value: '2713', font: 'Segoe UI Symbol' }, uncheckedState: { value: '25A1', font: 'Segoe UI Symbol' } },
      Round: { checkedState: { value: '25CF', font: 'Segoe UI Symbol' }, uncheckedState: { value: '25CB', font: 'Segoe UI Symbol' } },
    } as const
    const typeDefaults: Partial<Record<DocumentContentControlType, Record<string, unknown>>> = {
      richText: { placeholder: text, showingPlaceholder: true },
      text: { placeholder: text, showingPlaceholder: true, multiLine: false },
      checkbox: { checked: false, checkbox: { checked: false, ...checkboxStyles.Cross }, metadata: { checkboxStyle: 'Cross' } },
      date: { placeholder: '选择日期', showingPlaceholder: true, dateFormat: 'yyyy-MM-dd', dateCalendar: 'gregorian', dateStorage: 'dateTime' },
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
function insertPageField(story: 'header' | 'footer', field: 'PAGE' | 'NUMPAGES') {
  execute({ id: 'doc.command.insert-page-field', sectionId: headerFooterSectionId.value, story, field }, false)
}
function updateTitle() { emit('update:fileName', documentTitle.value); publish(cloneDocument(model.value)) }
function saveDocument() { emit('save', { fileName: documentTitle.value, html: documentHtml(model.value), model: model.value }) }
function toggleCommentsPanel() {
  commentsPanelOpen.value = !commentsPanelOpen.value
  if (!commentsPanelOpen.value) { activeCommentId.value = ''; commentDraftSelection.value = undefined; commentDraftReference.value = '' }
}
function startAddComment() {
  if (!canCreateComment.value) {
    const key = !canComment.value || !selectionCollapsed(selection.value) ? 'doc.comments.permission-denied' : 'doc.comments.select-text'
    const fallback = key === 'doc.comments.permission-denied' ? 'You do not have permission to comment' : 'Select text to add a comment'
    notify(createRibbonTranslator(officeLocale.value, 'doc')(key, fallback))
    return
  }
  commentDraftSelection.value = { anchor: { ...selection.value.anchor }, focus: { ...selection.value.focus } }
  commentDraftReference.value = selectedPlainText(model.value, selection.value)
  activeCommentId.value = ''
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
  if (!range || !canComment.value) return false
  if (selectedPlainText(model.value, range) !== commentDraftReference.value) {
    notify(createRibbonTranslator(officeLocale.value, 'doc')('doc.comments.selection-changed', 'The selected text changed. Select it again before commenting'))
    commentDraftSelection.value = undefined
    commentDraftReference.value = ''
    return false
  }
  let commentId = ''
  const committed = commitCommentMutation(next => { commentId = createDocumentThreadComment(next, range, text, commentActor()).id })
  if (committed) {
    commentDraftSelection.value = undefined
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
  const comment = model.value.comments?.[commentId]
  if (!comment) return
  activeCommentId.value = commentId
  const firstAnchor = comment.anchor[0]
  const lastAnchor = comment.anchor.at(-1)
  const targets = textTargets(model.value)
  const firstTarget = firstAnchor && targets.find(target => (target.cellBlockId ?? target.blockId) === firstAnchor.blockId)
  const lastTarget = lastAnchor && targets.find(target => (target.cellBlockId ?? target.blockId) === lastAnchor.blockId)
  const firstBlock = firstTarget && blockAtPoint(model.value, firstTarget)
  const lastBlock = lastTarget && blockAtPoint(model.value, lastTarget)
  if (!firstAnchor || !lastAnchor || !firstTarget || !lastTarget || !firstBlock || !lastBlock) {
    notify(createRibbonTranslator(officeLocale.value, 'doc')('doc.comments.anchor-unavailable', 'Text anchor unavailable'))
    return
  }
  // Imported threads and command updates may leave anchors stale. Clamp before navigating.
  const startOffset = Math.max(0, Math.min(firstAnchor.startOffset, runText(firstBlock.runs).length))
  const endOffset = Math.max(0, Math.min(lastAnchor.endOffset, runText(lastBlock.runs).length))
  onSelection({ anchor: { ...firstTarget, offset: startOffset }, focus: { ...lastTarget, offset: endOffset } })
  canvasRef.value?.scrollToBlock(firstTarget.blockId)
  canvasRef.value?.focus()
}
function importFile(file: File): false { emit('import-file', file); return false }
function movePage(direction: -1 | 1) { setActivePage(Math.max(0, Math.min(pages.value.length - 1, activePageIndex.value + direction))) }
function setActivePage(index: number) { activePageIndex.value = index; canvasRef.value?.scrollToPage(index) }
function jumpToHeading(id: string) { const point = { blockId: id, offset: 0 }; outlineNavigationTarget.value = id; onSelection(collapsedSelection(point)); canvasRef.value?.scrollToBlock(id); canvasRef.value?.focus() }
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
function openDocumentContextMenu(payload: { x: number; y: number }) {
  const bounds = workbenchRef.value?.getBoundingClientRect()
  if (!bounds) return
  documentContextMenu.value = {
    left: Math.max(4, Math.min(payload.x - bounds.left, bounds.width - 238)),
    top: Math.max(4, Math.min(payload.y - bounds.top, bounds.height - 300)),
  }
}
function runContextMenuAction(action: { id: string; commandId?: string; params?: unknown }) {
  onRibbonExecute(action.commandId ?? action.id, action.params)
  documentContextMenu.value = undefined
  void nextTick(() => canvasRef.value?.focus())
}
function selectDocumentContextMenuItem(info: { key: string | number }) {
  const action = contextMenuActions.value.find(item => item.id === String(info.key))
  if (action && !action.disabled) runContextMenuAction(action)
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
      documentTocPageNumbers(layout.value.headings),
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
  tocDialog.value = { open: true, mode: 'update', insertMode: 'custom', tocId, updateMode: 'pageNumbersOnly' }
}
function applyTocDialog(payload: { options: DocumentTocOptions; updateMode?: DocumentTocUpdateMode }) {
  if (tocDialog.value.mode === 'insert') {
    if (!canInsertToc.value) return
    const translate = createRibbonTranslator(officeLocale.value, 'doc')
    const blocks = buildDocumentTocBlocks(
      model.value,
      payload.options,
      documentTocPageNumbers(layout.value.headings),
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
  if (!tocId || selectedTocId.value !== tocId || !canEditSelectedToc.value) return
  execute({
    id: 'doc.command.update-table-of-contents',
    tocId,
    mode: payload.updateMode ?? 'pageNumbersOnly',
    pageNumbers: documentTocPageNumbers(layout.value.headings),
  })
}
function dismissTableContextMenu(event?: Event) {
  if (event?.target instanceof HTMLElement && event.target.closest('.document-context-menu')) return
  documentContextMenu.value = undefined
}
function onWorkbenchKeydown(event: KeyboardEvent) {
  if (event.key === 'Escape' && painter.value && !event.isComposing) { event.preventDefault(); painter.value = undefined; canvasRef.value?.focus(); return }
  if (event.key === 'Escape' && documentContextMenu.value) { event.preventDefault(); documentContextMenu.value = undefined; canvasRef.value?.focus() }
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
  const range: DocumentSelection = { anchor: { ...selection.value.anchor }, focus: { ...selection.value.focus } }
  const source = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(reader.error); reader.readAsDataURL(file) }).catch(() => '')
  if (request !== imageRequest || !modalOpen.value || modalKind.value !== 'image') return
  if (!source) { notify('读取图片失败'); return }
  const image = new Image()
  image.onload = () => {
    if (request !== imageRequest || !modalOpen.value || modalKind.value !== 'image') return
    if (model.value !== base) { notify('文档已更新，请重新插入图片'); return }
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
    onSelection({ anchor: { blockId: updated.blockId, offset: updated.startOffset }, focus: { blockId: updated.blockId, offset: updated.endOffset } })
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
    const targets = textTargets(model.value)
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
  const targets = matches.flatMap((match) => permissionTargetsForSelection(model.value, pointForMatch(match)))
  return [...new Map(targets.map((target) => [`${target.scope}:${target.id}`, target])).values()]
}
function replaceCurrent() {
  const match = searchMatches.value[currentSearchMatchIndex()]
  if (!match) { findNext(); return }
  if (!canDocumentPermission('edit', searchPermissionTargets([match]))) return
  try {
    const next = cloneDocument(model.value)
    const range = replaceDocumentMatches(next, [match], replaceText.value, { documentId: next.id ?? '', userId: props.documentApi?.getCurrentUserId() ?? props.permissionUserId, defaultAllow: props.permissionDefaultAllow })
    if (!range) return
    selection.value = range
    publish(next)
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
    const next = cloneDocument(model.value)
    const range = replaceDocumentMatches(next, matches, replaceText.value, { documentId: next.id ?? '', userId: props.documentApi?.getCurrentUserId() ?? props.permissionUserId, defaultAllow: props.permissionDefaultAllow })
    if (!range) return
    selection.value = range
    publish(next)
    searchInSelection.value = false
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) {
    notify(error instanceof Error ? error.message : String(error))
  }
}
async function copySelection(cut = false) {
  if (storyLocation.value) return
  const targets = permissionTargetsForSelection(model.value, selection.value)
  if (!canDocumentPermission('copy') || (cut && !canDocumentPermission('edit', targets))) return
  const text = selectedPlainText(model.value, selection.value)
  if (!text) return
  const base = model.value
  const range = JSON.stringify(selection.value)
  const request = ++clipboardRequest
  try {
    await navigator.clipboard.writeText(text)
    if (cut) {
      if (request !== clipboardRequest || model.value !== base || JSON.stringify(selection.value) !== range) { notify('文档或选区已变化，内容已复制，请重新剪切'); return }
      execute({ id: 'doc.command.insert-text', text: '' })
    }
  } catch { notify(cut ? '无法访问剪贴板，内容未删除，请使用 Ctrl+X 剪切' : '无法访问剪贴板，请使用 Ctrl+C 复制') }
  canvasRef.value?.focus()
}
async function pasteSelection(plainText = false) {
  if (storyLocation.value || !canEditContextSelection.value) return
  const base = model.value
  const range = JSON.stringify(selection.value)
  const request = ++clipboardRequest
  try {
    if (!navigator.clipboard) throw new Error('Clipboard is unavailable')
    const data = await readDocumentClipboard(navigator.clipboard, plainText)
    if (request !== clipboardRequest || model.value !== base || JSON.stringify(selection.value) !== range) { notify('文档或选区已变化，请重新粘贴'); return }
    const command = documentPasteCommand(data, plainText)
    if (command) execute(command)
  } catch { notify('无法读取剪贴板，请在文档中使用 Ctrl+V 或 Ctrl+Shift+V 粘贴') }
  canvasRef.value?.focus()
}
async function pasteSpecialSelection(mode: DocumentPasteMode) {
  if (storyLocation.value || !canEditContextSelection.value) return
  const base = model.value
  const range = JSON.stringify(selection.value)
  const request = ++clipboardRequest
  try {
    if (!navigator.clipboard) throw new Error('Clipboard is unavailable')
    const data = await readDocumentClipboard(navigator.clipboard)
    if (request !== clipboardRequest || model.value !== base || JSON.stringify(selection.value) !== range) { notify('文档或选区已变化，请重新粘贴'); return }
    const paragraph = blockAtPoint(model.value, selection.value.focus)
    if (!paragraph) { notify('无法确定目标段落格式，请重新选择插入位置。'); return }
    const command = documentPasteSpecialCommand(data, mode, {
      paragraphStyle: { ...paragraph.style },
      textStyle: { ...marksAtPoint(model.value, selection.value.focus), ...pendingMarks.value },
    })
    if (command) execute(command)
  } catch { notify('无法读取剪贴板，请在文档中使用 Ctrl+V 或 Ctrl+Shift+V 粘贴') }
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
const canEditStory = computed(() => resolveDocumentPermission(model.value.permissionRules ?? [], permissionContext.value, 'edit', [{ scope: 'section', id: headerFooterSectionId.value }]).allowed)
function openStory(location: StoryLocation) {
  painter.value = undefined
  if (pageMode.value !== 'traditional' || !layout.value.pageHeaderFooters[location.page]) return
  clipboardRequest++
  if (!storyLocation.value) previousStoryTab.value = activeTab.value
  storyLocation.value = location
  activeTab.value = 'header-footer'
  void nextTick(() => canvasRef.value?.revealStory(location))
}
function closeStory(refocus = true) {
  if (!storyLocation.value) return
  storyLocation.value = undefined
  if (activeTab.value === 'header-footer') activeTab.value = previousStoryTab.value
  if (refocus) void nextTick(() => canvasRef.value?.focus())
}
watch(activeTab, tab => { if (storyLocation.value && tab !== 'header-footer') closeStory(false) })
watch([pageMode, () => model.value.id, storyPage], ([mode, id, page], [, oldId]) => {
  if (storyLocation.value && (mode !== 'traditional' || id !== oldId || !page)) closeStory(false)
})
function editStory(value: string, caret: number) {
  if (!storyKey.value || storySection.value.linkToPrevious || !canEditStory.value || value === storyValue.value) return
  applyHeaderFooter(editStoryBuffer(storySection.value, storyKey.value, value, caret))
}
function handleStoryAction(raw: unknown) {
  const location = storyLocation.value, key = storyKey.value
  if (!location || !key || typeof raw !== 'string') return
  let action: HeaderFooterAction
  try { action = JSON.parse(raw) } catch { return }
  if (action.action === 'navigate') { openStory({ page: location.page, area: action.area }); return }
  if (!canEditStory.value) return
  if (action.action === 'set-options') {
    const { marginHeader, marginFooter, ...options } = action.patch
    if (marginHeader !== undefined) applySectionSettings(headerFooterSectionId.value, { marginHeader })
    if (marginFooter !== undefined) applySectionSettings(headerFooterSectionId.value, { marginFooter })
    if (Object.keys(options).length) applyHeaderFooter(options)
  } else {
    const section = storySection.value
    if (section.linkToPrevious) return
    const now = new Date(), locale = officeLocale.value.replaceAll('_', '-')
    if (action.action === 'refresh-fields') {
      const dates = (['short', 'long', 'time'] as const).map(format => storyDateField(format, locale, now))
      const fields = section.pageFields?.map(field => {
        const date = dates.find(d => d.instruction === field.instruction)
        return date ? { ...field, cachedResult: date.cachedResult } : field
      })
      if (fields && fields.some((field, i) => field.cachedResult !== section.pageFields?.[i]?.cachedResult)) applyHeaderFooter({ pageFields: fields })
    } else if (action.action === 'insert-field' || action.action === 'insert-date-time') {
      const date = action.action === 'insert-date-time' ? storyDateField(action.format, locale, now) : undefined
      const insert = action.action === 'insert-field' ? { field: action.field } : action.automatic ? date! : date!.cachedResult
      const { start, end } = canvasRef.value?.storyRange() ?? { start: 0, end: 0 }
      applyHeaderFooter(replaceStoryRange(section, key, start, end, insert))
      canvasRef.value?.focusStory(start + (typeof insert === 'string' ? insert.length : 1))
      return
    }
  }
  canvasRef.value?.focusStory()
}
function activatePainter(continuous: boolean) {
  if (painter.value) {
    painter.value = continuous ? { ...painter.value, continuous } : undefined
  } else if (!storyLocation.value && canEditContextSelection.value) {
    const format = captureDocumentFormat(model.value, selection.value)
    if (format) painter.value = { format, continuous }
  }
  canvasRef.value?.focus()
}
function applyPainter(format?: DocumentPainterFormat) {
  if (!format || storyLocation.value || !documentPainterTarget(model.value, selection.value)) return
  if (execute({ id: 'doc.command.apply-format-painter', format }) && !painter.value?.continuous) painter.value = undefined
}
function onRibbonExecute(id: string, params?: unknown) {
  if (id === 'ui.operation.activate-format-painter' || id === 'ui.operation.continuous-format-painter') { activatePainter(id === 'ui.operation.continuous-format-painter'); return }
  if (id === 'ui.command.clear-formatting') { painter.value = undefined; applyPainter(clearDocumentFormat()); return }
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
    case 'doc.operation.open-header-footer': openStory({ page: layout.value.lines.find(line => line.target.blockId === selection.value.focus.blockId && line.start <= selection.value.focus.offset && line.end >= selection.value.focus.offset)?.page ?? activePageIndex.value, area: 'header' }); break
    case 'sidebar.operation.doc-paragraph-setting-panel': paragraphSettingsPanelOpen.value = true; break
    case 'sidebar.operation.doc-section-setting-panel': sectionSettingsPanelOpen.value = true; break
    case 'doc.command.set-paper-size': if (params === 'Custom') pageSettingsOpen.value = true; else if (params === 'A4' || params === 'Letter') { paperSize.value = params; changePageSetup() }; break
    case 'doc.command.set-orientation': if (params === 'portrait' || params === 'landscape') { orientation.value = params; changePageSetup() }; break
  }
}
defineExpose({ save: saveDocument, focus: () => canvasRef.value?.focus() })
</script>

<template>
  <ConfigProvider>
  <section ref="workbenchRef" class="document-workbench" :class="{ 'focus-mode': isFocusMode }" aria-label="文档工作台" @pointerdown="dismissTableContextMenu" @keydown="onWorkbenchKeydown">
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
    <main class="editor-workspace">
      <DocumentOutline v-if="chrome.leftSidebar && showOutline && !isFocusMode" :model="model" :active-id="activeOutlineId" :mode="outlineMode" :max-width="256" @navigate="navigateOutline" />
      <DocumentCanvasViewport ref="canvasRef" :model="model" :selection="selection" :zoom="zoom" :search-matches="searchMatches"
        @selection-committed="applyPainter(painter?.format)"
        :stories-enabled="pageMode === 'traditional'" :story="storyLocation" :story-value="storyValue" :story-readonly="!canEditStory || storySection.linkToPrevious"
        @open-story="openStory" @close-story="closeStory()" @edit-story="editStory" @command="execute" @code-block-exit="leaveCodeBlock" @toggle-format="toggleFormat" @update:selection="onSelection" @navigate-to-heading="jumpToHeading" @context-menu="openDocumentContextMenu" @layout="layout = $event" @active-page="activePageIndex = $event" @viewport-change="onViewportChange" @undo="undo" @redo="redo" @save="saveDocument" @notify="notify" />
      <aside v-if="commentsPanelOpen" class="document-comments-dock" :aria-label="createRibbonTranslator(officeLocale, 'doc')('doc.comments.title', 'Comments')">
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
    <DocumentContentControlPanel v-model:open="contentControlPanelOpen" :selected-control-id="selectedContentControlId" :document-api="documentApi" @error="notify" />
    <DocumentTableInsertDialog v-model:open="tableInsertOpen" @insert="size => execute({ id: 'doc.command.insert-table', rows: size.rows, cols: size.columns })" />
    <DocumentPageSettings v-model:open="pageSettingsOpen" :page="model.page" :mode="pageMode" @apply="applyPageSetup" />
    <DocumentHeaderFooterPanel v-model:open="headerFooterPanelOpen" :model="model" :section-id="headerFooterSectionId" :setup="activeSectionSetup" @apply="applyHeaderFooter" @position="applySectionSettings(headerFooterSectionId, $event)" @insert-page-field="insertPageField" @closed="canvasRef?.focus()" />
    <DocumentParagraphSettingsPanel v-model:open="paragraphSettingsPanelOpen" :style="paragraphSettingsStyle" :font-size="paragraphSettingsFontSize" @apply="applyParagraphSettings" />
    <DocumentSectionSettingsPanel v-model:open="sectionSettingsPanelOpen" :section-id="headerFooterSectionId" :setup="activeSectionSetup" :section-options="sectionOptions" @apply="applySectionSettings" @select-section="selectSection" />
    <DocumentCodeBlockPanel v-model:open="codeBlockPanelOpen" :initial="activeCodeBlockData" @apply="applyCodeBlock" />
    <DocumentLatexPanel v-model:open="latexPanelOpen" :initial="activeLatexFormula?.formula" @apply="applyLatexFormula" />
    <DocumentNotesPanel v-model:open="notesPanelOpen" :footnotes="model.footnotes ?? {}" :endnotes="model.endnotes ?? {}" @apply="applyNotes" @remove="removeNote" @navigate="navigateToNote" />
    <DocumentPermissionPanel :open="permissionPanelOpen" :rules="model.permissionRules ?? []" :targets="permissionTargets" :read-only="!canManagePermissionRules" @update:open="permissionPanelOpen = $event" @save="savePermissionRules" />
    <DocumentTableOfContentsDialog
      v-model:open="tocDialog.open"
      :mode="tocDialog.mode"
      :options="tocDialogOptions"
      :update-mode="tocDialog.updateMode"
      @apply="applyTocDialog"
    />
  </section>
  </ConfigProvider>
</template>

<style scoped>
.document-workbench { position: relative; min-width: 0; height: 100%; min-height: 0; display: flex; flex-direction: column; overflow: hidden; color: #30343b; background: #f3f5f7; font: 13px/1.45 "Microsoft YaHei", "Segoe UI", sans-serif; }
.document-header { flex: 0 0 44px; display: flex; align-items: center; gap: 8px; min-width: 0; padding: 0 12px; border-bottom: 1px solid #e4e7eb; background: #fff; }
.document-mark { flex-shrink: 0; color: #3182ce; }
.document-name { min-width: 60px; flex: 1; }
.document-header :deep(.ant-btn) { display: inline-flex; align-items: center; justify-content: center; gap: 6px; height: 28px; }
.document-ribbon { flex-shrink: 0; }
.editor-workspace { position: relative; min-width: 0; min-height: 0; flex: 1; display: flex; overflow: hidden; }
.document-comments-dock { z-index: 30; width: 320px; min-width: 280px; flex: 0 0 min(320px, 38vw); border-inline-start: 1px solid #d9dde4; background: #fff; }
.find-replace-toolbar { position: absolute; z-index: 30; top: calc(100% - 1px); right: 8px; width: min(620px, calc(100vw - 24px)); display: flex; flex-direction: column; gap: 6px; padding: 8px; border: 1px solid #d8dde5; border-radius: 4px; background: #fff; box-shadow: 0 3px 10px #18283b20; }
.find-row, .replace-row { min-width: 0; display: flex; align-items: center; gap: 4px; }
.find-row { color: #77818d; }
.find-row :deep(.ant-input), .replace-row :deep(.ant-input) { min-width: 0; padding: 0 4px; color: #424a54; background: transparent; font: inherit; }
.find-row :deep(.ant-input) { flex: 1; }
.replace-row :deep(.ant-input) { width: 145px; }
.find-row :deep(.ant-btn) { width: 24px; height: 24px; padding: 0; display: grid; place-items: center; color: #5b6672; }
.match-count { flex: 0 0 auto; min-width: 42px; color: #6e7783; font-variant-numeric: tabular-nums; text-align: center; }
.replace-row :deep(.ant-checkbox-wrapper) { flex: 0 0 auto; white-space: nowrap; font-size: 12px; }
.document-notice { position: absolute; z-index: 40; bottom: 44px; inset-inline: 12px; display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 8px 12px; border: 1px solid #c9dcf4; border-radius: 4px; color: #4b6482; background: #eef5ff; font-size: 12px; box-shadow: 0 2px 8px #18375912; }
.document-notice :deep(.ant-btn) { display: grid; place-items: center; padding: 0; color: inherit; }
.document-context-menu { position: absolute; z-index: 60; width: 230px; overflow: auto; border-radius: 4px; }
.document-context-menu :deep(.ant-menu) { border-inline-end: 0; box-shadow: 0 4px 14px #18283b24; }
.insert-form { display: flex; flex-direction: column; gap: 16px; padding: 12px 0; }
.insert-form :deep(.ant-form-item) { margin-bottom: 0; }
.insert-error { color: #c14444; }
@media (max-width: 650px) { .find-replace-toolbar { right: 4px; width: calc(100vw - 16px); } .replace-row { flex-wrap: wrap; } .replace-row :deep(.ant-input) { flex: 1 1 100px; } .document-header { padding: 0 6px; } }
@media print { .document-header, .document-ribbon, .document-notice, :deep(.document-outline), :deep(.document-statusbar) { display: none !important; } .document-workbench, .editor-workspace { height: auto; overflow: visible; background: #fff; } }
</style>
