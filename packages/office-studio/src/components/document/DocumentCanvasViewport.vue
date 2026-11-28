<script setup lang="ts">
import { computed, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue'
import { Alert, Button, Input, Popover, theme } from 'antdv-next'
import type { DocumentModel } from '../../domain/model'
import type { RenderSurface } from '../../rendering/contracts'
import { createRenderSurface } from '../../rendering/createRenderSurface'
import { blockAtPoint, documentTabCommand, graphemeOffsets, type DocumentCommand } from '../../features/document/commands/document-commands'
import { applyDocumentEditorCommand } from '../../api/document-editor-commands'
import { caretRect, createTextMeasurer, documentPhysicalLineNavigation, documentStoryCaretRect, documentStoryLayoutLines, emptyLayout, hitTestDocument, hitTestDocumentStory, layoutDocument, lineForPoint, type DocumentLayout } from '../../features/document/layout/document-layout'
import { cloneDocument, marksAtPoint, orderedSelection, runText, runsForTarget, sameTarget, selectionCollapsed, textTargets, type DocumentPoint, type DocumentSelection, type RunPatch, type TextTarget } from '../../features/document/model/document-model'
import { documentSelectionForContextMenu } from '../../features/document/model/document-selection'
import type { DocumentClipboardData, DocumentPasteMode } from '../../features/document/services/document-clipboard'
import { sceneFromDocument, sceneOffset, type DocumentViewport } from '../../features/document/rendering/document-scene'
import { documentRectScrollOffset, documentViewportLocation, zoomedDocumentScrollOffset, type DocumentScrollOffset, type DocumentViewportLocation } from '../../features/document/layout/document-viewport'
import { codeBlockData } from '../../features/document/code-blocks/code-blocks.ts'
import { documentStructuredBlockAtSelection } from '../../features/document/commands/document-structure-commands'
import { clampDocumentCodeScrollOffsets, documentCodeCaretScrollOffset, documentCodeScrollRegions, projectDocumentCodeScroll } from '../../features/document/layout/document-code-scroll'
import { captureDocumentStructuredCanvasTarget, documentStructuredCanvasTargetCurrent, type DocumentStructuredCanvasTarget, type DocumentStructuredCanvasCommandRequest, type DocumentStructuredCanvasOperationRequest } from '../../features/document/controllers/document-structured-canvas'
import { documentLatexCanvasRequestCurrent, type DocumentLatexCanvasEditRequest } from '../../features/document/controllers/document-latex-canvas'
import { documentLatexDragAnchor, hitTestDocumentLatex, type DocumentLatexHit } from '../../features/document/layout/document-latex-hit'
import DocumentStructuredOverlay from './DocumentStructuredOverlay.vue'
import { documentTocHeadingTarget } from '../../features/document/toc/document-toc'
import type { DocumentTocLayoutPublication } from '../../features/document/toc/document-toc-pagination'
import { documentHyperlinkForSelection } from '../../features/document/services/document-hyperlinks'
import { documentSegmentIdForHost } from '../../api/document-segments'
import type { StoryLocation } from '../../features/document/services/story-editor'
import { documentFormattingWordRange } from '../../features/document/services/document-format-painter'
import type { DocumentInputSession, DocumentStoryInputCommit } from '../../features/document/model/document-input-session'
import { documentInputLocation, sameDocumentInputLocation } from '../../features/document/model/document-input-session'
import type { DocumentStoryRange } from '../../features/document/model/document-story'
import { useDocumentStoryInput } from './use-document-story-input'
import { cloneDocumentSelection, sameDocumentTableViewSelection } from '../../features/document/model/document-table-selection'
import { documentTableRangeBetweenHits, documentTableSelectionAtHit, documentTableSelectionFocusCell, documentTableVisibleRect, hitTestDocumentTable, type DocumentTableDragPointer, type DocumentTableGesturePreview, type DocumentTableHit } from '../../features/document/layout/document-table-interaction'
import { navigateDocumentTableSelection, type DocumentTableKeyboardCursor } from '../../features/document/model/document-table-navigation'
import { documentTableCommandState } from '../../features/document/commands/table-commands'
import { assertDocumentTextSelectionEditable } from '../../features/document/services/document-text-formatting'
import { permissionTargetsForSelection, resolveDocumentPermission, type DocumentPermissionContext } from '../../features/document/permissions/document-permissions'
import DocumentTableOverlay from './DocumentTableOverlay.vue'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../../features/document/model/document-block-host'
import { documentModelForBlockHost } from '../../features/document/services/document-block-host'
import { documentLayoutForBlockHost } from '../../features/document/layout/document-auxiliary-layout'
import { layoutDocumentBlockHostInspection } from '../../features/document/layout/document-block-host-inspection'
import DocumentContentControlOverlay from './DocumentContentControlOverlay.vue'
import { documentContentControlGeometry } from '../../features/document/layout/content-control-geometry'
import { appendDocumentContentControlScene } from '../../features/document/rendering/content-control-scene'
import { selectedDocumentContentControlForSelection } from '../../features/document/services/content-controls'
import { contentControlValueCommandState } from '../../features/document/commands/content-control-command-state'
import type { DocumentContentControlCommandRequest, DocumentContentControlPopupMode, DocumentContentControlSelectionRequest } from '../../features/document/commands/content-control-command-types'
import { captureDocumentContentControlFocusToken, documentContentControlFocusCurrent, documentContentControlFocusTokenMatchesRequest, type DocumentContentControlFocusToken } from '../../features/document/controllers/content-control-interaction'
import { documentContentControlCommandIsTextInput, documentContentControlCommandUsesWholeCellInput, prepareDocumentContentControlTextInputSelection } from '../../features/document/services/content-control-editing'
import { documentContentControlContainsSelection } from '../../features/document/model/content-control-table-rows'
import { hitTestDocumentNoteDecoration, hitTestDocumentNoteGlyph, hitTestDocumentNoteRegion, type DocumentNoteGlyphHit } from '../../features/document/layout/document-note-hit'
import { documentNoteCanvasRequestCurrent, documentNotePreview, type DocumentNoteCanvasNavigationRequest, type DocumentNotePreview } from '../../features/document/controllers/document-note-canvas'
import { documentImageEntries } from '../../features/document/images/document-image-model'
import { assertDocumentImageEditable, type DocumentImageCommand } from '../../features/document/commands/document-image-commands'
import { documentImageBatchMembers, assertDocumentImageBatchMembers, documentImageBatchBounds, type DocumentImageAlign } from '../../features/document/images/document-image-batch'
import type { DocumentImageArrange } from '../../features/document/images/document-image-model'
import { permissionTargetsForImage } from '../../features/document/permissions/document-permissions'
import { captureDocumentImageCanvasTarget, documentImageCanvasTargetCurrent, documentImageCanvasCommandMatchesTarget, documentImageAtPoint, documentImageResizeCorner, documentImageAtRotationHandle, documentImageGestureResult, documentImageBatchGestureResult,
  type DocumentImageCanvasTarget, type DocumentImageCanvasCommandRequest, type DocumentImageCanvasOperationRequest, type DocumentImageGesture, type DocumentImageGestureResult, type DocumentImageBatchGesture, type DocumentImageBatchGestureResult } from '../../features/document/controllers/document-image-canvas'
import DocumentImageOverlay from './DocumentImageOverlay.vue'
import DocumentHeadingCollapseOverlay from './DocumentHeadingCollapseOverlay.vue'
import { normalizeDocumentHeadingSelection, visibleDocumentTextTargets } from '../../features/document/headings/heading-collapse'

const props = defineProps<{ model: DocumentModel; selection: DocumentSelection; zoom: number; selectedImageId?: string; selectedImageIds?: readonly string[]; host?: DocumentAuxiliaryBlockHost; inspectionHost?: DocumentAuxiliaryBlockHost; activeContentControlId?: string; activeContentControlPopupMode?: DocumentContentControlPopupMode; permissionContext?: DocumentPermissionContext; tableFloatingToolbar?: boolean; storiesEnabled?: boolean; story?: StoryLocation; storyValue?: string; storySelection?: { start: number; end: number }; storyReadonly?: boolean; typingMarks?: RunPatch; captureInputSession?: (sourceModel?: DocumentModel) => DocumentInputSession | undefined; searchMatches?: Array<{ target: TextTarget; start: number; end: number }> }>()
const emit = defineEmits<{
  command: [command: DocumentCommand]
  'clipboard-copy': [transfer: DataTransfer, cut: boolean]
  'clipboard-paste': [data: DocumentClipboardData, mode: DocumentPasteMode]
  'structured-command': [request: DocumentStructuredCanvasCommandRequest]
  'structured-operation': [request: DocumentStructuredCanvasOperationRequest]
  'latex-edit': [request: DocumentLatexCanvasEditRequest]
  'image-selection': [target: DocumentImageCanvasTarget | undefined]
  'image-command': [request: DocumentImageCanvasCommandRequest]
  'image-operation': [request: DocumentImageCanvasOperationRequest]
  operation: [id: string, params?: unknown]
  'open-story': [location: StoryLocation, point?: DocumentPoint]
  'close-story': []
  'edit-story': [value: string, caret: number, marks?: RunPatch, input?: DocumentStoryInputCommit, sourceRange?: { start: number; end: number }]
  'update:story-selection': [range: { start: number; end: number }, preserveInputStyle?: boolean]
  'story-composing': [value: boolean]
  'code-block-exit': [blockId: string]
  'toggle-format': [field: 'bold' | 'italic' | 'underline' | 'strike' | 'subscript' | 'superscript', source?: DocumentModel, story?: DocumentStoryRange]
  'refresh-fields': [source: DocumentModel, story?: DocumentStoryRange, bodySelection?: DocumentSelection]
  'update:selection': [selection: DocumentSelection]
  'selection-committed': []
  'content-control-selection': [request: DocumentContentControlSelectionRequest]
  'content-control-command': [request: DocumentContentControlCommandRequest]
  'navigate-to-heading': [headingId: string]
  'note-navigation': [request: DocumentNoteCanvasNavigationRequest]
  'context-menu': [payload: { x: number; y: number; point: DocumentPoint; source?: DocumentModel; story?: DocumentStoryRange; host?: DocumentAuxiliaryBlockHost }]
  layout: [layout: DocumentLayout, source: DocumentModel]
  'layout-state': [publication: DocumentTocLayoutPublication]
  'active-page': [page: number]
  'viewport-change': [location: DocumentViewportLocation]
  undo: []
  redo: []
  save: []
  notify: [message: string]
}>()
const { token } = theme.useToken()
const scrollRef = ref<HTMLDivElement | null>(null)
const hostRef = ref<HTMLDivElement | null>(null)
interface DocumentTextAreaBridgeRef { blur: () => void; nativeElement?: HTMLElement; $el?: Node }
const inputRef = ref<DocumentTextAreaBridgeRef | null>(null)
function bridgeTextArea(bridge: DocumentTextAreaBridgeRef | null | undefined): HTMLTextAreaElement | null {
  // Prefer antdv-next's exposed nativeElement. A component root may be a
  // wrapper or a fragment comment, so $el alone is not a stable input handle.
  for (const element of [bridge?.nativeElement, bridge?.$el]) {
    if (element instanceof HTMLTextAreaElement) return element
    if (element instanceof HTMLElement) {
      const input = element.querySelector('textarea')
      if (input) return input
    }
  }
  return null
}
function inputElement(): HTMLTextAreaElement | null { return bridgeTextArea(inputRef.value) }
const sourceLayout = shallowRef<DocumentLayout>(emptyLayout())
const codeScrollOffsets = shallowRef<Record<string, number>>({})
const layout = computed(() => projectDocumentCodeScroll(sourceLayout.value, codeScrollOffsets.value))
const structuredOverlay = ref<{ cancelInteraction: () => void }>()
const inspectionActive = computed(() => !!props.inspectionHost && sameDocumentBlockHost(props.inspectionHost, props.host))
const tableOverlay = ref<{ cancelGesture: () => void; cancelInteraction: () => void; continueGesture: (pointer: DocumentTableDragPointer) => void }>()
const contentControlOverlay = ref<{ cancelInteraction: () => void }>()
const contentControlActive = ref(true)
const hoverTable = shallowRef<DocumentTableHit>()
const noteHover = shallowRef<{ source: DocumentModel; glyph: DocumentNoteGlyphHit; preview: DocumentNotePreview }>()
let noteHoverHide: ReturnType<typeof setTimeout> | undefined
const tablePreview = shallowRef<DocumentTableGesturePreview>()
const storyRef = ref<DocumentTextAreaBridgeRef>()
function storyElement() { return bridgeTextArea(storyRef.value) ?? undefined }
function storyBeforeInput(event: InputEvent) {
  if (!storyEventIsCurrent(event) || props.storyReadonly) { event.preventDefault(); return }
  captureStoryInput(event)
  if (event.inputType === 'historyUndo' || event.inputType === 'historyRedo') { event.preventDefault(); emit(event.inputType === 'historyUndo' ? 'undo' : 'redo') }
}
const paragraphShortcutCommands: Partial<Record<string, DocumentCommand>> = {
  KeyL: { id: 'doc.command.format-paragraph', patch: { align: 'left' } },
  KeyE: { id: 'doc.command.format-paragraph', patch: { align: 'center' } },
  KeyR: { id: 'doc.command.format-paragraph', patch: { align: 'right' } },
  KeyJ: { id: 'doc.command.format-paragraph', patch: { align: 'justify' } },
  Digit7: { id: 'doc.command.set-block-type', type: 'numbered' },
  Digit8: { id: 'doc.command.set-block-type', type: 'bullet' },
}
function storyKeys(event: KeyboardEvent) {
  if (!storyEventIsCurrent(event)) return
  if (event.isComposing || event.keyCode === 229) return
  const key = event.key.toLowerCase(), mod = event.ctrlKey || event.metaKey
  if (key === 'f9' && !mod && !event.altKey && !event.shiftKey) {
    if (storyComposing.value || document.activeElement !== storyElement()) return
    event.preventDefault(); event.stopPropagation()
    const source = props.model, location = props.story, page = location && layout.value.pageHeaderFooters[location.page]
    const storyKey = location?.area === 'header' ? page?.headerStory : page?.footerStory, range = storyRange()
    if (!page || page.isParityBlank || !storyKey) return
    syncStorySelection(event)
    if (!storyEventIsCurrent(event) || source !== props.model) return
    emit('refresh-fields', source, { sectionId: page.sectionId, story: storyKey, ...range })
    return
  }
  const paragraphShortcut = mod && event.shiftKey && !event.altKey ? paragraphShortcutCommands[event.code] : undefined
  if (paragraphShortcut) {
    if (props.storyReadonly || storyComposing.value || selectedImageIds.value.length || document.activeElement !== storyElement()) return
    const source = props.model
    syncStorySelection(event)
    if (!storyEventIsCurrent(event) || source !== props.model) return
    event.preventDefault(); event.stopPropagation()
    emit('command', paragraphShortcut)
    return
  }
  const verticalShortcut = mod && !event.altKey && !event.shiftKey && (event.code === 'Comma' || event.code === 'Period')
  const strikeShortcut = mod && !event.altKey && event.shiftKey && event.code === 'KeyX'
  if (mod && ['b', 'i', 'u'].includes(key) || verticalShortcut || strikeShortcut) {
    if ((verticalShortcut || strikeShortcut) && (selectedImageIds.value.length || document.activeElement !== storyElement())) return
    event.preventDefault(); event.stopPropagation()
    if (props.storyReadonly || storyComposing.value) return
    const source = props.model, location = props.story, page = location && layout.value.pageHeaderFooters[location.page]
    const storyKey = location?.area === 'header' ? page?.headerStory : page?.footerStory, range = storyRange()
    if (!page || page.isParityBlank || !storyKey) return
    syncStorySelection(event)
    if (!storyEventIsCurrent(event) || source !== props.model) return
    if (strikeShortcut) {
      emit('toggle-format', 'strike', source, { sectionId: page.sectionId, story: storyKey, ...range })
      return
    }
    if (verticalShortcut) {
      emit('toggle-format', event.code === 'Comma' ? 'subscript' : 'superscript', source, { sectionId: page.sectionId, story: storyKey, ...range })
      return
    }
    emit('toggle-format', key === 'b' ? 'bold' : key === 'i' ? 'italic' : 'underline', source, { sectionId: page.sectionId, story: storyKey, ...range })
    return
  }
  if (key !== 'escape' && (!mod || !['s', 'z', 'y'].includes(key))) return
  event.preventDefault(); event.stopPropagation()
  if (key === 'escape') { emit('close-story'); return }
  if (key === 's') { emit('save'); return }
  emit(key === 'y' || event.shiftKey ? 'redo' : 'undo')
}
const preview = shallowRef<DocumentModel | null>(null)
const compositionSelection = shallowRef<DocumentSelection | null>(null)
const composing = ref(false)
const bodyBridgeVersion = ref(0)
const focused = ref(false)
const caretVisible = ref(true)
const viewport = ref<DocumentViewport>({ width: 1, height: 1, scrollLeft: 0, scrollTop: 0, zoom: props.zoom })
const backend = ref('初始化画布')
const error = ref('')
const images = new Map<string, HTMLImageElement>()
let surface: RenderSurface | null = null
let resizeObserver: ResizeObserver | undefined
let dprQuery: MediaQueryList | undefined
let lastSurfaceSize: { width: number; height: number; dpr: number } | undefined
let frame = 0
let blink: ReturnType<typeof setInterval> | undefined
let dragAnchor: DocumentPoint | undefined
let latexDragSelection: DocumentSelection | undefined
let tableDragAnchor: DocumentTableHit | undefined
let tableKeyboardCursor: DocumentTableKeyboardCursor | undefined
let dragSelectionSource: DocumentModel | undefined
let dragScrollFrame = 0
let dragScrollPointer: DocumentTableDragPointer | undefined
let dragScrollSource: 'selection' | 'overlay' | 'image' | undefined
let storyDragAnchor: { location: StoryLocation; offset: number } | undefined
let pointerId: number | undefined
let suppressCompositionInput = false
let pendingText = ''
let compositionModel: DocumentModel | undefined
let compositionRange: DocumentSelection | undefined
let compositionInput: HTMLTextAreaElement | null | undefined
let compositionMarks: RunPatch = {}
let compositionSession: DocumentInputSession | undefined
let compositionHost: DocumentAuxiliaryBlockHost | undefined
let compositionControlId: string | undefined
let blockedBodyInput: HTMLTextAreaElement | null | undefined
let disposed = false
let active = true
let nextPasteMode: DocumentPasteMode = 'source'
let viewportRestoreVersion = 0
let pendingContentControlFocus: DocumentContentControlFocusToken | undefined
let measurer = createTextMeasurer()
let layoutGeneration = 0
let fontVersion = 0
let fontSet: FontFaceSet | undefined
let fontsReady = false
function publishLayoutState(complete = false) {
  emit('layout-state', { source: displayModel.value, layoutGeneration, fontVersion,
    complete: complete && active && !disposed && !preview.value && !composing.value && !storyComposing.value
      && !inspectionActive.value && fontsReady && fontSet?.status !== 'loading',
    layout: sourceLayout.value, engine: 'main-thread' })
}
function fontsLoading() { fontVersion++; fontsReady = false; publishLayoutState() }
function fontsChanged() {
  fontVersion++; fontsReady = fontSet?.status !== 'loading'
  // Widths cached for fallback fonts cannot serve a new font generation.
  measurer = createTextMeasurer(hostRef.value?.ownerDocument)
  if (!disposed && active) relayout()
}
const { inputValue: storyInputValue, composing: storyComposing, bridgeVersion: storyBridgeVersion, storyRange, storyEventIsCurrent, syncStorySelection, focusStory, storyInputChanged,
  captureStoryInput, storyCompositionStart, storyCompositionEnd, storyFocused, cancelStoryInput } = useDocumentStoryInput(props, {
  element: storyElement, layout: () => layout.value, active: () => active && !disposed,
  edit: (value, caret, marks, input, sourceRange) => emit('edit-story', value, caret, marks, input, sourceRange), select: (range, preserveInputStyle) => emit('update:story-selection', range, preserveInputStyle),
  composing: value => emit('story-composing', value), preview: model => { preview.value = model; relayout() },
  draw: () => { caretVisible.value = true; requestDraw() }, notify: message => emit('notify', message),
})
function bodyFocusChanged(event: FocusEvent) {
  if (!bodyEventIsCurrent(event)) return
  focused.value = event.type === 'focus'
  if (!focused.value) nextPasteMode = 'source'
  requestDraw()
}
function storyFocusChanged(event: FocusEvent) {
  if (!storyEventIsCurrent(event)) return
  focused.value = event.type === 'focus'
  if (focused.value) storyFocused(event)
  else requestDraw()
}
const displayModel = computed(() => preview.value ?? props.model)
const editingModel = computed(() => props.host ? documentModelForBlockHost(props.model, props.host) : props.model)
const displayEditingModel = computed(() => props.host ? documentModelForBlockHost(displayModel.value, props.host) : displayModel.value)
const editingLayout = computed(() => documentLayoutForBlockHost(layout.value, props.host))
const sourceEditingLayout = computed(() => documentLayoutForBlockHost(sourceLayout.value, props.host))
const selectedImageIds = computed(() => props.selectedImageIds ?? (props.selectedImageId ? [props.selectedImageId] : []))
const selectedImagePage = ref<number>()
const selectedImage = computed(() => documentImageEntries(editingModel.value).find(entry => entry.imageId === selectedImageIds.value.at(-1)))
const selectedImageMembers = computed(() => {
  if (!selectedImageIds.value.length) return []
  try { return documentImageBatchMembers(editingModel.value, editingLayout.value, selectedImageIds.value, selectedImagePage.value) } catch { return [] }
})
const selectedImageRectangle = computed(() => selectedImageMembers.value.find(member => member.imageId === selectedImageIds.value.at(-1))?.rectangle)
const selectedImageFrame = computed(() => selectedImageMembers.value.length > 1 ? documentImageBatchBounds(selectedImageMembers.value) : selectedImageRectangle.value)
const imageGeometryState = computed(() => {
  if (selectedImageMembers.value.length !== selectedImageIds.value.length || !selectedImageMembers.value.length) return { available: false, reason: '请等待图片排版完成。' }
  try { assertDocumentImageBatchMembers(editingModel.value, selectedImageMembers.value); return { available: true } }
  catch (cause) { return { available: false, reason: cause instanceof Error ? cause.message : String(cause) } }
})
const imageArrangeAvailable = computed(() => {
  const entries = documentImageEntries(editingModel.value), first = entries.find(entry => entry.imageId === selectedImageIds.value[0])
  return !!first && selectedImageIds.value.every(id => entries.some(entry => entry.imageId === id && entry.container === first.container))
})
const imageGesture = shallowRef<DocumentImageGesture | DocumentImageBatchGesture>(), imageGesturePreview = shallowRef<DocumentImageGestureResult | DocumentImageBatchGestureResult>()
const imageContextPoint = shallowRef<{ x: number; y: number }>()
let imageSelectionRequest = 0, imagePointerModifiers = { altKey: false, shiftKey: false }
function imageTarget(imageId = selectedImageIds.value.at(-1), ids: readonly string[] = selectedImageIds.value): DocumentImageCanvasTarget | undefined {
  return imageId ? captureDocumentImageCanvasTarget(props.model, imageId, props.selection, props.host, props.captureInputSession?.(props.model), ids) : undefined
}
function imageEditable(imageId = selectedImageIds.value.at(-1)): boolean {
  if (!imageId || !active || disposed || composing.value || storyComposing.value || legacyStoryActive.value || props.host?.kind === 'story' && props.storyReadonly) return false
  const context = props.permissionContext ?? { documentId: props.model.id ?? '', userId: 'local-user' }
  const targets = permissionTargetsForImage(editingModel.value, imageId), rules = props.model.permissionRules ?? []
  if (!resolveDocumentPermission(rules, context, 'edit', targets).allowed || !resolveDocumentPermission(rules, context, 'view', targets).allowed) return false
  try { assertDocumentImageEditable(editingModel.value, imageId); return true } catch { return false }
}
const imageReadOnly = computed(() => !selectedImageIds.value.length || selectedImageIds.value.some(id => !imageEditable(id)))
function imageTargetCurrent(target: DocumentImageCanvasTarget, requireEditable = true): boolean {
  return documentImageCanvasTargetCurrent(target, { source: props.model, model: editingModel.value, selection: props.selection, host: props.host,
    imageIds: selectedImageIds.value, session: props.captureInputSession?.(), requireSession: !!props.captureInputSession,
    disabled: !active || disposed || composing.value || storyComposing.value || legacyStoryActive.value
      || requireEditable && (target.imageIds ?? [target.imageId]).some(id => !imageEditable(id)) })
}
function imageOperation(mode: 'edit' | 'crop' | 'replace') {
  if (selectedImageIds.value.length !== 1) return
  const target = imageTarget()
  if (target && imageTargetCurrent(target)) { imageContextPoint.value = undefined; emit('image-operation', { target, mode }) }
}
function imageCommand(command: DocumentImageCommand) {
  const target = imageTarget()
  if (target && imageTargetCurrent(target) && documentImageCanvasCommandMatchesTarget({ target, command })) {
    imageContextPoint.value = undefined; cancelImageGesture(); emit('image-command', { target, command })
  }
}
function removeCanvasImages() {
  const ids = selectedImageIds.value
  if (ids.length === 1) imageCommand({ id: 'doc.command.remove-image', imageId: ids[0]! })
  else if (ids.length) imageCommand({ id: 'doc.command.batch-images', commands: ids.map(imageId => ({ id: 'doc.command.remove-image', imageId })) })
}
function resetCanvasImage() {
  const id = selectedImageIds.value.length === 1 ? selectedImageIds.value[0] : undefined
  if (id) imageCommand({ id: 'doc.command.reset-image-size', imageId: id })
}
function arrangeCanvasImages(arrange: DocumentImageArrange) {
  if (!imageArrangeAvailable.value) return
  const ids = selectedImageIds.value
  if (ids.length === 1) imageCommand({ id: 'doc.command.arrange-image', imageId: ids[0]!, arrange })
  else imageCommand({ id: 'doc.command.arrange-images', imageIds: [...ids], arrange })
}
function alignCanvasImages(align: DocumentImageAlign) {
  if (imageGeometryState.value.available && selectedImageIds.value.length > 1) imageCommand({ id: 'doc.command.align-images', members: selectedImageMembers.value, align })
}
function closeImageContext() {
  const target = imageTarget(); imageContextPoint.value = undefined
  if (target) void nextTick(() => { if (imageTargetCurrent(target, false)) focus() })
}
function selectCanvasImage(imageId: string, page: number, after?: (target: DocumentImageCanvasTarget) => void, ids: readonly string[] = [imageId], preserveSelection = false) {
  const source = props.model, host = props.host && { ...props.host }, entry = documentImageEntries(editingModel.value).find(entry => entry.imageId === imageId)
  if (!entry || !ids.includes(imageId)) return
  const request = ++imageSelectionRequest
  const before = props.captureInputSession?.(source), snapshot = props.captureInputSession ? undefined : JSON.stringify(source)
  const selection: DocumentSelection = !preserveSelection && ids.length === 1 && entry.inline ? {
    anchor: { ...entry.inline.target, offset: entry.inline.startOffset, pageAffinity: page },
    focus: { ...entry.inline.target, offset: entry.inline.endOffset, pageAffinity: page },
  } : cloneDocumentSelection(props.selection)
  if (JSON.stringify(selection) !== JSON.stringify(props.selection)) select(selection)
  focus()
  void nextTick(() => {
    const current = props.captureInputSession?.(source)
    if (request !== imageSelectionRequest || !active || disposed || composing.value || storyComposing.value || source !== props.model || !sameDocumentBlockHost(host, props.host)
      || JSON.stringify(selection) !== JSON.stringify(props.selection)
      || before && current?.modelRevision !== before.modelRevision || snapshot !== undefined && snapshot !== JSON.stringify(source)) return
    const target = imageTarget(imageId, ids)
    if (!target) return
    selectedImagePage.value = page
    emit('image-selection', target); requestDraw()
    if (after) void nextTick(() => { if (request === imageSelectionRequest && imageTargetCurrent(target, false)) after(target) })
  })
}
function cancelImageGesture() {
  if (imageGesture.value?.pointerId === pointerId) cancelSelectionDrag()
  if (dragScrollSource === 'image') cancelDragAutoscroll()
  imageGesture.value = undefined; imageGesturePreview.value = undefined; requestDraw()
}
function moveImageGesture(pointer: Pick<PointerEvent, 'pointerId' | 'clientX' | 'clientY' | 'altKey' | 'shiftKey'>) {
  const gesture = imageGesture.value, coordinates = coordinatesFromEvent(pointer)
  if (!gesture || !coordinates || pointer.pointerId !== gesture.pointerId) return
  if (!imageTargetCurrent(gesture.target) || gesture.zoom !== props.zoom) { cancelImageGesture(); return }
  imagePointerModifiers = { altKey: pointer.altKey, shiftKey: pointer.shiftKey }
  if (Math.hypot(coordinates.x - gesture.start.x, coordinates.y - gesture.start.y) * gesture.zoom / 100 < 2) { imageGesturePreview.value = undefined; requestDraw(); return }
  try {
    imageGesturePreview.value = 'members' in gesture ? documentImageBatchGestureResult(gesture, coordinates.x, coordinates.y, pointer.shiftKey)
      : documentImageGestureResult(gesture, coordinates.x, coordinates.y, pointer.altKey, pointer.shiftKey)
  } catch (cause) { cancelImageGesture(); emit('notify', cause instanceof Error ? cause.message : String(cause)); return }
  requestDraw()
}
watch(() => [props.model, props.host, props.zoom], () => { imageSelectionRequest++; cancelImageGesture(); imageContextPoint.value = undefined })
watch(() => JSON.stringify(selectedImageIds.value), () => {
  if (imageGesture.value && !imageTargetCurrent(imageGesture.value.target)) cancelImageGesture()
  requestDraw()
})
watch(() => [JSON.stringify(props.selection), props.captureInputSession?.()?.generation, composing.value, storyComposing.value], () => {
  if (imageGesture.value && !imageTargetCurrent(imageGesture.value.target)) cancelImageGesture()
})
const legacyStoryActive = computed(() => !!props.story && !props.host)
function clearImageSelection() {
  if (!selectedImageIds.value.length) return
  imageSelectionRequest++; cancelImageGesture(); imageContextPoint.value = undefined; emit('image-selection', undefined)
}
function dispatchCommand(command: DocumentCommand) {
  if (documentContentControlCommandIsTextInput(command)) clearImageSelection()
  const input = documentContentControlCommandIsTextInput(command) && !documentContentControlCommandUsesWholeCellInput(command, props.selection)
    ? { ...command, activeContentControlId: command.activeContentControlId ?? explicitInputControlId() } : command
  emit('command', props.host ? { ...input, host: { ...props.host } } : input)
}
function structuredTargetCurrent(target: DocumentStructuredCanvasTarget): boolean {
  return documentStructuredCanvasTargetCurrent(target, { source: props.model, selection: props.selection, host: props.host,
    session: props.captureInputSession?.(), requireSession: !!props.captureInputSession,
    disabled: !active || disposed || composing.value || storyComposing.value || legacyStoryActive.value || props.host?.kind === 'story' && props.storyReadonly })
}
function dispatchStructuredCommand(request: DocumentStructuredCanvasCommandRequest) { if (structuredTargetCurrent(request)) emit('structured-command', request) }
function dispatchStructuredOperation(request: DocumentStructuredCanvasOperationRequest) { if (structuredTargetCurrent(request)) emit('structured-operation', request) }
function focusStructured(target: DocumentStructuredCanvasTarget) { if (structuredTargetCurrent(target)) focus() }
function requestLatexEdit(hit: DocumentLatexHit) {
  const source = props.model, host = props.host && { ...props.host }, before = props.captureInputSession?.(source)
  const snapshot = props.captureInputSession ? undefined : JSON.stringify(source)
  select(hit.selection); focus()
  // Selection publication advances the input session. Capture the new session
  // after Vue has delivered it, while retaining the pre-selection content revision.
  void nextTick(() => {
    if (source !== props.model || !sameDocumentBlockHost(host, props.host)) return
    const session = props.captureInputSession?.(source)
    if (before && session?.modelRevision !== before.modelRevision || snapshot !== undefined && snapshot !== JSON.stringify(source)) return
    const request: DocumentLatexCanvasEditRequest = {
      ...captureDocumentStructuredCanvasTarget(source, hit.selection, host, session), rangeId: hit.item.rangeId,
    }
    if (documentLatexCanvasRequestCurrent(request, { source: props.model, selection: props.selection, host: props.host, session,
      requireSession: !!props.captureInputSession, disabled: !active || disposed || composing.value || storyComposing.value || legacyStoryActive.value
        || props.host?.kind === 'story' && props.storyReadonly })) emit('latex-edit', request)
  })
}
function scrollCode(request: { source: DocumentModel; key: string; value: number }) {
  if (!active || disposed || composing.value || storyComposing.value || request.source !== props.model || !Number.isFinite(request.value)) return
  const region = documentCodeScrollRegions(sourceLayout.value).find(region => region.key === request.key)
  if (!region) return
  codeScrollOffsets.value = { ...codeScrollOffsets.value, [request.key]: Math.max(0, Math.min(region.maxScroll, request.value)) }
}
const displaySelection = computed(() => compositionSelection.value ?? props.selection)
const contentControlGeometries = computed(() => documentContentControlGeometry(displayEditingModel.value, editingLayout.value))
const activeContentControl = computed(() => {
  return selectedDocumentContentControlForSelection(displayEditingModel.value, displaySelection.value, props.activeContentControlId)
})
function explicitInputControlId(): string | undefined {
  const selection = props.selection
  if (selection.tableSelection || !sameTarget(selection.anchor, selection.focus)) return undefined
  const control = editingModel.value.contentControls?.find(control => control.id === props.activeContentControlId)
  if (!control || !documentContentControlContainsSelection(editingModel.value, control, selection)) return undefined
  return control.id
}
function contentControlRequestCurrent(request: DocumentContentControlSelectionRequest): boolean {
  if (!active || disposed || composing.value || storyComposing.value || legacyStoryActive.value
    || props.host?.kind === 'story' && props.storyReadonly || request.source !== props.model || !sameDocumentBlockHost(request.host, props.host)) return false
  const control = editingModel.value.contentControls?.find(control => control.id === request.controlId)
  if (!control || !documentContentControlContainsSelection(editingModel.value, control, request.selection)) return false
  const captured = request.inputSession, current = captured && props.captureInputSession?.(request.source)
  return !captured || !!current && current.valid && captured.valid && captured.generation === current.generation
    && captured.modelRevision === current.modelRevision && sameDocumentInputLocation(captured.location, current.location)
}
function selectContentControl(request: DocumentContentControlSelectionRequest) {
  if (!contentControlRequestCurrent(request)) return
  pendingContentControlFocus = undefined
  emit('content-control-selection', request)
  caretVisible.value = true; requestDraw()
}
function dispatchContentControlCommand(request: DocumentContentControlCommandRequest) {
  pendingContentControlFocus = undefined
  if (!contentControlRequestCurrent(request) || request.command.params.controlId !== request.controlId
    || props.captureInputSession && !request.inputSession?.valid) return
  const state = contentControlValueCommandState(props.model, request.command, { host: props.host, permission: props.permissionContext })
  if (!state.available) { emit('notify', state.reason ?? '无法修改内容控件。'); return }
  emit('content-control-command', request)
  // Read passively after the shared command has restored its checked selection.
  // Passing an old source model here must not renew a stale input session.
  pendingContentControlFocus = captureDocumentContentControlFocusToken(request, props.captureInputSession?.())
}
/** Ribbon value actions share the overlay's checked command and delayed focus path. */
function executeContentControlCommand(request: DocumentContentControlCommandRequest) {
  dispatchContentControlCommand(request)
  void nextTick(() => focusContentControl(request))
}
function focusContentControl(request: DocumentContentControlSelectionRequest) {
  if (!active || disposed || composing.value || storyComposing.value || legacyStoryActive.value
    || props.host?.kind === 'story' && props.storyReadonly || props.model.id !== request.source.id || !sameDocumentBlockHost(request.host, props.host)) return
  if (request.mode !== undefined && request.mode !== (props.activeContentControlPopupMode ?? 'value')) return
  const pending = pendingContentControlFocus
  if (pending && documentContentControlFocusTokenMatchesRequest(pending, request)) {
    pendingContentControlFocus = undefined
    if (!documentContentControlFocusCurrent(pending, { source: props.model, host: props.host,
      session: props.captureInputSession?.(), selection: props.selection })) return
  } else if (!contentControlRequestCurrent(request)) return
  focus()
}
const storyStyle = computed(() => {
  const target = props.story, entry = target && layout.value.pageHeaderFooters[target.page]
  if (!entry || entry.isParityBlank || !target) return { display: 'none' }
  const k = props.zoom / 100, off = sceneOffset(layout.value, viewport.value)
  const caret = documentStoryCaretRect(documentStoryLayoutLines(layout.value, target), props.storySelection?.end ?? 0)
  return { left: `${off.x + (caret?.x ?? entry.x) * k}px`, top: `${off.y + (caret?.y ?? (target.area === 'header' ? entry.headerTop : entry.footerBottom)) * k}px`, fontSize: `${Math.max(12, (caret?.height ?? 12) * k)}px` }
})
const spacerStyle = computed(() => ({ width: `${Math.max(viewport.value.width, layout.value.width * props.zoom / 100)}px`, height: `${layout.value.height * props.zoom / 100}px` }))
const inputStyle = computed(() => {
  const caret = caretRect(editingLayout.value, displaySelection.value.focus)
  const view = displaySelection.value.tableSelection
  const selectedCell = view && documentTableSelectionFocusCell(editingLayout.value, displaySelection.value, tableKeyboardCursor?.focus)
  const offset = sceneOffset(layout.value, viewport.value)
  return {
    left: `${offset.x + (selectedCell?.x ?? caret?.x ?? 0) * props.zoom / 100}px`,
    top: `${offset.y + (selectedCell?.y ?? caret?.y ?? 0) * props.zoom / 100}px`,
    fontSize: `${Math.max(12, (caret?.height ?? 18) * props.zoom / 100)}px`,
  }
})
const noteHoverStyle = computed(() => {
  const hover = noteHover.value
  if (!hover) return { display: 'none' }
  const offset = sceneOffset(layout.value, viewport.value), scale = props.zoom / 100
  return { left: `${offset.x + hover.glyph.rect.x * scale}px`, top: `${offset.y + hover.glyph.rect.y * scale}px`,
    width: `${Math.max(1, hover.glyph.rect.width * scale)}px`, height: `${Math.max(1, hover.glyph.rect.height * scale)}px` }
})

watch(() => props.model, () => { clearNoteHover(); hoverTable.value = undefined; cancelSelectionDrag(); tableKeyboardCursor = undefined; if (composing.value && !bodyCompositionCurrent()) cancelBodyComposition(); if (!composing.value && !storyComposing.value) preview.value = null; relayout() }, { deep: false })
watch(() => props.story, story => { if (story && composing.value) cancelBodyComposition(); hoverTable.value = undefined; cancelSelectionDrag(); tableKeyboardCursor = undefined; requestDraw() })
watch(() => [props.host, props.inspectionHost], () => {
  clearNoteHover(); cancelInput(false); preview.value = null; hoverTable.value = undefined; relayout()
  const source = props.model, host = props.host && { ...props.host }, session = props.captureInputSession?.(), restoreVersion = viewportRestoreVersion
  void nextTick(() => {
    const current = props.captureInputSession?.()
    if (disposed || !active || source !== props.model || !sameDocumentBlockHost(host, props.host)
      || session && (!current?.valid || current.generation !== session.generation || current.modelRevision !== session.modelRevision)) return
    focus()
    if (restoreVersion === viewportRestoreVersion) ensureCaretVisible()
  })
}, { deep: true })
watch(() => props.selection, () => {
  clearNoteHover()
  const view = props.selection.tableSelection, range = tableKeyboardCursor?.range
  if (tableKeyboardCursor && (!view || view.tableId !== tableKeyboardCursor.tableId || view.range.startRow !== range!.startRow
    || view.range.endRow !== range!.endRow || view.range.startColumn !== range!.startColumn || view.range.endColumn !== range!.endColumn)) tableKeyboardCursor = undefined
  if (composing.value && !bodyCompositionCurrent()) cancelBodyComposition()
  const restoreVersion = viewportRestoreVersion
  caretVisible.value = true; requestDraw(); void nextTick(() => { if (restoreVersion === viewportRestoreVersion) ensureCaretVisible() })
}, { deep: true })
watch(() => props.captureInputSession?.(), () => {
  clearNoteHover()
  if (composing.value && !bodyCompositionCurrent()) cancelBodyComposition(false)
}, { flush: 'sync' })
watch([composing, storyComposing], ([body, story]) => {
  if (body || story) { clearNoteHover(); publishLayoutState() }
  else if (!disposed && active) publishLayoutState(true)
}, { flush: 'sync' })
watch(() => props.searchMatches, requestDraw, { deep: true })
watch(codeScrollOffsets, requestDraw)
watch(() => props.model.id, () => { codeScrollOffsets.value = {} })
watch(() => token.value.colorPrimary, requestDraw)
watch(() => [token.value.colorBorder, token.value.colorText, token.value.colorError], requestDraw)
watch(() => props.activeContentControlId, () => {
  if (composing.value && !bodyCompositionCurrent()) cancelBodyComposition(false)
  requestDraw()
}, { flush: 'sync' })
watch(tablePreview, requestDraw)
watch(() => props.zoom, (zoom) => {
  clearNoteHover()
  cancelSelectionDrag()
  tableOverlay.value?.cancelInteraction()
  const offset = zoomedDocumentScrollOffset(layout.value, viewport.value, zoom)
  const restoreVersion = viewportRestoreVersion
  viewport.value = { ...viewport.value, zoom }
  requestDraw()
  void nextTick(() => {
    if (disposed || props.zoom !== zoom || !scrollRef.value || restoreVersion !== viewportRestoreVersion) return
    scrollRef.value.scrollLeft = offset.left
    scrollRef.value.scrollTop = offset.top
    syncViewport('zoom')
  })
})

onActivated(() => { active = true; contentControlActive.value = true; void nextTick(() => { resize(); relayout() }) })
onDeactivated(() => {
  active = false
  publishLayoutState()
  nextPasteMode = 'source'
  cancelImageGesture(); imageContextPoint.value = undefined
  clearNoteHover()
  pendingContentControlFocus = undefined
  contentControlActive.value = false
  contentControlOverlay.value?.cancelInteraction()
  structuredOverlay.value?.cancelInteraction()
  cancelBodyComposition()
  cancelStoryInput()
  focused.value = false
  inputRef.value?.blur()
  storyElement()?.blur()
  if (pointerId != null && scrollRef.value?.hasPointerCapture(pointerId)) scrollRef.value.releasePointerCapture(pointerId)
  pointerId = undefined
  dragAnchor = undefined
  tableDragAnchor = undefined
  storyDragAnchor = undefined
  cancelDragAutoscroll()
  tableOverlay.value?.cancelInteraction()
  cancelAnimationFrame(frame)
  frame = 0
  surface?.resize(0, 0, window.devicePixelRatio || 1)
  lastSurfaceSize = undefined
})

onMounted(async () => {
  if (!hostRef.value || !scrollRef.value) return
  fontSet = hostRef.value.ownerDocument.fonts
  measurer = createTextMeasurer(hostRef.value.ownerDocument)
  fontsReady = !fontSet
  fontSet?.addEventListener('loading', fontsLoading)
  fontSet?.addEventListener('loadingdone', fontsChanged)
  fontSet?.addEventListener('loadingerror', fontsChanged)
  void fontSet?.ready.then(() => { if (!disposed) fontsChanged() })
  resize()
  relayout()
  try {
    const nextSurface = await createRenderSurface(hostRef.value, { preferredBackend: 'webgpu', onBackendChange: (value) => { backend.value = value; requestDraw() } })
    if (disposed) { nextSurface.dispose(); return }
    surface = nextSurface
    backend.value = nextSurface.backend
    resizeObserver = new ResizeObserver(resize)
    resizeObserver.observe(scrollRef.value)
    watchDevicePixelRatio()
    window.addEventListener('resize', onViewportResize)
    window.visualViewport?.addEventListener('resize', onViewportResize)
    blink = setInterval(() => { caretVisible.value = !caretVisible.value; requestDraw() }, 530)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
})

onUnmounted(() => {
  disposed = true
  publishLayoutState()
  fontSet?.removeEventListener('loading', fontsLoading)
  fontSet?.removeEventListener('loadingdone', fontsChanged)
  fontSet?.removeEventListener('loadingerror', fontsChanged)
  cancelImageGesture()
  structuredOverlay.value?.cancelInteraction()
  clearNoteHover()
  pendingContentControlFocus = undefined
  cancelDragAutoscroll()
  cancelBodyComposition()
  cancelStoryInput()
  resizeObserver?.disconnect()
  dprQuery?.removeEventListener('change', onDevicePixelRatioChange)
  window.removeEventListener('resize', onViewportResize)
  window.visualViewport?.removeEventListener('resize', onViewportResize)
  if (blink) clearInterval(blink)
  cancelAnimationFrame(frame)
  surface?.dispose()
  images.forEach((image) => { image.onload = null; image.onerror = null; image.src = '' })
})

function resize() {
  if (!scrollRef.value) return
  const width = scrollRef.value.clientWidth
  const height = scrollRef.value.clientHeight
  const dpr = Math.max(1, window.devicePixelRatio || 1)
  viewport.value = { ...viewport.value, width, height, scrollLeft: scrollRef.value.scrollLeft, scrollTop: scrollRef.value.scrollTop }
  if (surface && (lastSurfaceSize?.width !== width || lastSurfaceSize.height !== height || lastSurfaceSize.dpr !== dpr)) {
    surface.resize(width, height, dpr)
    lastSurfaceSize = { width, height, dpr }
  }
  publishViewport('resize')
  requestDraw()
}

function watchDevicePixelRatio() {
  dprQuery?.removeEventListener('change', onDevicePixelRatioChange)
  dprQuery = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`)
  dprQuery.addEventListener('change', onDevicePixelRatioChange, { once: true })
}

function onDevicePixelRatioChange() {
  watchDevicePixelRatio()
  resize()
}

function onViewportResize() {
  watchDevicePixelRatio()
  resize()
}

function relayout() {
  if (disposed) return
  layoutGeneration++
  publishLayoutState()
  sourceLayout.value = inspectionActive.value ? layoutDocumentBlockHostInspection(displayModel.value, props.inspectionHost!, measurer)
    : layoutDocument(displayModel.value, measurer)
  codeScrollOffsets.value = clampDocumentCodeScrollOffsets(sourceLayout.value, codeScrollOffsets.value)
  const layoutImages = [...layout.value.images, ...(layout.value.auxiliaryContent ?? []).flatMap(content => content.images),
    ...(layout.value.noteDecorations ?? []).flatMap(content => content.images)]
  const sources = new Set(layoutImages.map((image) => image.src))
  for (const [src, image] of images) {
    if (sources.has(src)) continue
    image.onload = null
    image.onerror = null
    image.src = ''
    images.delete(src)
  }
  for (const image of layoutImages) {
    if (!image.src || images.has(image.src)) continue
    const source = new Image()
    source.crossOrigin = image.src.startsWith('data:') || image.src.startsWith('blob:') ? null : 'anonymous'
    source.onload = () => { if (!disposed && images.get(image.src) === source) requestDraw() }
    source.onerror = () => { if (!disposed && images.get(image.src) === source) { emit('notify', `图片无法载入：${image.alt || image.src.slice(0, 80)}`); requestDraw() } }
    images.set(image.src, source)
    source.src = image.src
  }
  emit('layout', sourceLayout.value, displayModel.value)
  publishLayoutState(true)
  publishViewport('layout')
  void nextTick(() => { if (!disposed) syncViewport('layout') })
  requestDraw()
}

function requestDraw() {
  if (frame || disposed || !active) return
  frame = requestAnimationFrame(() => {
    frame = 0
    const scene = sceneFromDocument(displayModel.value, layout.value, viewport.value, { selection: displaySelection.value, focused: focused.value, caretVisible: caretVisible.value, composing: composing.value || storyComposing.value, images,
      imageSelection: imageGesturePreview.value?.rectangle ?? selectedImageFrame.value, imageHandleColor: token.value.colorBgElevated,
      imageHandles: !imageReadOnly.value && imageGeometryState.value.available,
      ...(selectedImageMembers.value.length > 1 ? { imageSelections: selectedImageMembers.value.map(member => {
        const result = imageGesturePreview.value
        return result && 'updates' in result ? result.updates.find(update => update.imageId === member.imageId)?.rectangle ?? member.rectangle : member.rectangle
      }) } : {}),
      ...(imageGesture.value && imageGesturePreview.value ? { imagePreviews: 'updates' in imageGesturePreview.value
        ? imageGesturePreview.value.updates.map(update => ({ imageId: update.imageId, page: selectedImageMembers.value.find(member => member.imageId === update.imageId)?.rectangle.page,
          rectangle: update.rectangle, wrappingStyle: update.patch.wrappingStyle }))
        : [{ imageId: imageGesture.value.target.imageId, page: selectedImageRectangle.value?.page,
          rectangle: imageGesturePreview.value.rectangle, wrappingStyle: imageGesturePreview.value.patch.wrappingStyle }] } : {}),
      searchMatches: props.searchMatches, tocPrimaryColor: token.value.colorPrimary, formulaErrorColor: token.value.colorError, tableGesture: tablePreview.value,
      host: props.host, selectionModel: displayEditingModel.value, selectionLayout: editingLayout.value,
      ...(legacyStoryActive.value && props.story ? { story: { location: props.story, selection: props.storySelection ?? storyRange() } } : {}) })
    if (!legacyStoryActive.value) appendDocumentContentControlScene(scene, contentControlGeometries.value, layout.value, viewport.value,
      { activeControlId: activeContentControl.value?.id, primary: token.value.colorPrimary, border: token.value.colorBorder })
    surface?.render(scene)
  })
}

function onScroll() {
  clearNoteHover()
  syncViewport('scroll')
}

function syncViewport(reason: NonNullable<DocumentViewportLocation['reason']>) {
  if (!scrollRef.value) return
  viewport.value = { ...viewport.value, scrollLeft: scrollRef.value.scrollLeft, scrollTop: scrollRef.value.scrollTop }
  publishViewport(reason)
  requestDraw()
}

function publishViewport(reason: NonNullable<DocumentViewportLocation['reason']>) {
  if (!active || disposed || layout.value.inspectionHost) return
  const location = documentViewportLocation(layout.value, viewport.value)
  if (location.pageIndex != null) emit('active-page', location.pageIndex)
  emit('viewport-change', { ...location, reason })
}

function coordinatesFromEvent(event: Pick<MouseEvent, 'clientX' | 'clientY'>): { x: number; y: number } | undefined {
  const element = scrollRef.value
  if (!element) return undefined
  const bounds = element.getBoundingClientRect()
  const offset = sceneOffset(layout.value, viewport.value)
  return { x: (event.clientX - bounds.left - offset.x) / (props.zoom / 100), y: (event.clientY - bounds.top - offset.y) / (props.zoom / 100) }
}

function pointFromEvent(event: Pick<MouseEvent, 'clientX' | 'clientY'>): DocumentPoint | undefined {
  const coordinates = coordinatesFromEvent(event)
  return coordinates ? hitTestDocument(editingLayout.value, coordinates.x, coordinates.y) : undefined
}

function select(selection: DocumentSelection) {
  if (!props.host && !legacyStoryActive.value) selection = normalizeDocumentHeadingSelection(editingModel.value, selection)
  emit('update:selection', selection)
  if (!legacyStoryActive.value) {
    const control = selectedDocumentContentControlForSelection(editingModel.value, selection)
    if (control && !control.metadata?.officeStoryLockSource) emit('content-control-selection', { controlId: control.id, selection, source: props.model,
      host: props.host ? { ...props.host } : undefined, inputSession: props.captureInputSession?.(props.model) })
  }
  caretVisible.value = true
  requestDraw()
}

function focus() {
  if (!active || disposed) return
  if (legacyStoryActive.value) { focusStory(); return }
  // antdv-next 1.8.2 TextArea.focus() does not forward focus options. Keep the
  // library bridge, but focus its native textarea so returning to a story does
  // not scroll the document before its captured viewport is restored.
  inputElement()?.focus({ preventScroll: true })
  focused.value = true; requestDraw()
}

function keepNoteHover() { if (noteHoverHide) clearTimeout(noteHoverHide); noteHoverHide = undefined }
function clearNoteHover() { keepNoteHover(); noteHover.value = undefined }
function deferNoteHoverHide() { keepNoteHover(); noteHoverHide = setTimeout(clearNoteHover, 150) }
function updateNoteHover(coordinates?: { x: number; y: number }) {
  if (!active || disposed || composing.value || storyComposing.value || pointerId !== undefined) { clearNoteHover(); return }
  const glyph = coordinates && hitTestDocumentNoteGlyph(layout.value, coordinates.x, coordinates.y)
  if (!glyph || glyph.role !== 'reference') { if (noteHover.value) deferNoteHoverHide(); return }
  keepNoteHover()
  const source = props.model, content = documentNotePreview(source, glyph.noteKind, glyph.noteId, layout.value.noteLabels, props.permissionContext)
  if (!content) { clearNoteHover(); return }
  noteHover.value = { source, glyph, preview: content }
}
function navigateCanvasNote(target: Omit<Extract<DocumentNoteCanvasNavigationRequest, { type: 'note' }>, 'source' | 'originHost' | 'inputSession'>
  | { type: 'body'; point: DocumentPoint }, source = props.model): boolean {
  const request: DocumentNoteCanvasNavigationRequest = { ...target, source, originHost: props.host && { ...props.host },
    inputSession: props.captureInputSession?.(source) }
  if (!documentNoteCanvasRequestCurrent(request, { source: props.model, host: props.host, session: props.captureInputSession?.(),
    requireSession: !!props.captureInputSession, disabled: !active || disposed || composing.value || storyComposing.value, permission: props.permissionContext })) return false
  clearNoteHover(); emit('note-navigation', request)
  return true
}
function navigateNotePreview() {
  const hover = noteHover.value
  if (hover) navigateCanvasNote({ type: 'note', kind: hover.preview.kind, noteId: hover.preview.noteId, toReference: false }, hover.source)
}
watch([() => props.permissionContext, () => props.model.permissionRules], clearNoteHover, { deep: true })

function storyAt(pos: { x: number; y: number } | undefined): StoryLocation | undefined {
  if (!pos || !props.storiesEnabled || inspectionActive.value || props.host?.kind === 'note') return
  const page = layout.value.pages.find(p => pos.x >= p.x && pos.x <= p.x + p.width && pos.y >= p.y && pos.y <= p.y + p.height)
  const entry = page && layout.value.pageHeaderFooters[page.index]
  if (!page || !entry || page.isParityBlank) return
  if (entry.headerBlockLayout && pos.y >= entry.headerBlockLayout.y && pos.y <= entry.headerBlockLayout.y + entry.headerBlockLayout.height) return { page: page.index, area: 'header' }
  if (entry.footerBlockLayout && pos.y >= entry.footerBlockLayout.y && pos.y <= entry.footerBlockLayout.y + entry.footerBlockLayout.height) return { page: page.index, area: 'footer' }
  if (pos.y < page.content.y || entry.headerLayoutLines.some(line => pos.y >= line.y && pos.y < line.y + line.height)) return { page: page.index, area: 'header' }
  if (pos.y >= page.content.y + page.content.height || entry.footerLayoutLines.some(line => pos.y >= line.y && pos.y < line.y + line.height)) return { page: page.index, area: 'footer' }
}

function pointerDown(event: PointerEvent) {
  if (!active || disposed || event.button !== 0 || composing.value || storyComposing.value) return
  nextPasteMode = 'source'
  tableKeyboardCursor = undefined
  const coordinates = coordinatesFromEvent(event)
  const noteGlyph = coordinates && hitTestDocumentNoteGlyph(layout.value, coordinates.x, coordinates.y)
  if (noteGlyph) {
    event.preventDefault()
    navigateCanvasNote({ type: 'note', kind: noteGlyph.noteKind, noteId: noteGlyph.noteId, toReference: noteGlyph.role === 'bodyLabel' })
    return
  }
  const noteRegion = coordinates && hitTestDocumentNoteRegion(layout.value, coordinates.x, coordinates.y)
  if (noteRegion && !sameDocumentBlockHost(noteRegion.host, props.host)) {
    event.preventDefault()
    navigateCanvasNote({ type: 'note', kind: noteRegion.host.noteKind, noteId: noteRegion.host.noteId, toReference: false, point: noteRegion.point })
    return
  }
  if (coordinates && hitTestDocumentNoteDecoration(layout.value, coordinates.x, coordinates.y)) { event.preventDefault(); return }
  if (!inspectionActive.value && props.host?.kind === 'note' && !noteRegion) {
    const point = coordinates && hitTestDocument(layout.value, coordinates.x, coordinates.y)
    if (point) { event.preventDefault(); navigateCanvasNote({ type: 'body', point }) }
    return
  }
  const story = storyAt(coordinates)
  const storyPage = story && layout.value.pageHeaderFooters[story.page], activeStory = story && props.host?.kind === 'story'
    && props.host.sectionId === storyPage?.sectionId && props.host.story === (story.area === 'header' ? storyPage?.headerStory : storyPage?.footerStory)
  if (story && !activeStory) {
    event.preventDefault()
    if (props.story?.page === story.page && props.story.area === story.area && coordinates) {
      const offset = hitTestDocumentStory(documentStoryLayoutLines(layout.value, story), coordinates.x, coordinates.y)
      storyDragAnchor = { location: story, offset: event.shiftKey ? props.storySelection?.start ?? offset : offset }
      pointerId = event.pointerId
      dragSelectionSource = props.model
      scrollRef.value?.setPointerCapture(event.pointerId)
      focusStory(Math.min(storyDragAnchor.offset, offset), Math.max(storyDragAnchor.offset, offset))
    }
    return
  }
  if (props.story && !props.host) emit('close-story')
  if (!inspectionActive.value && props.host?.kind === 'story' && !story) { emit('close-story'); return }
  const frame = imageGeometryState.value.available && !imageReadOnly.value ? selectedImageFrame.value : undefined
  const corner = coordinates && frame && documentImageResizeCorner(frame, coordinates.x, coordinates.y, props.zoom)
  const rotate = !!(coordinates && frame && documentImageAtRotationHandle(frame, coordinates.x, coordinates.y, props.zoom))
  const imageHit = corner || rotate ? selectedImageRectangle.value : coordinates && documentImageAtPoint(editingModel.value, editingLayout.value, coordinates.x, coordinates.y, event.altKey)
  if (imageHit && coordinates) {
    const image = documentImageEntries(editingModel.value).find(entry => entry.imageId === imageHit.blockId)?.image
    if (!image) return
    event.preventDefault(); cancelImageGesture(); cancelSelectionDrag(); imageContextPoint.value = undefined
    const previous = selectedImageIds.value, toggle = !corner && !rotate && (event.ctrlKey || event.metaKey || event.shiftKey)
    const ids = toggle ? previous.includes(imageHit.blockId) ? previous.filter(id => id !== imageHit.blockId) : [...previous, imageHit.blockId]
      : previous.includes(imageHit.blockId) ? [...previous] : [imageHit.blockId]
    if (ids.length > 1000) { emit('notify', '一次图片操作最多选择 1000 张图片。'); return }
    if (!ids.length) { imageSelectionRequest++; emit('image-selection', undefined); focus(); return }
    const primary = ids.at(-1)!, preserveSelection = ids.length > 1 || previous.includes(imageHit.blockId)
    if (!toggle && ids.every(id => imageEditable(id))) {
      pointerId = event.pointerId; dragSelectionSource = props.model; scrollRef.value?.setPointerCapture(event.pointerId)
    }
    selectCanvasImage(primary, imageHit.page, target => {
      if (toggle || pointerId !== event.pointerId || dragSelectionSource !== props.model || !imageTargetCurrent(target)) return
      if (ids.length > 1) {
        try {
          const members = documentImageBatchMembers(editingModel.value, editingLayout.value, ids, imageHit.page)
          assertDocumentImageBatchMembers(editingModel.value, members)
          imageGesture.value = { target, members, rectangle: documentImageBatchBounds(members), corner: corner || undefined, rotate,
            start: coordinates, pointerId: event.pointerId, zoom: props.zoom }
        } catch (cause) { cancelSelectionDrag(); emit('notify', cause instanceof Error ? cause.message : String(cause)) }
        return
      }
      const page = layout.value.pages[imageHit.page]
      imageGesture.value = { target, image: { ...image }, rectangle: { ...imageHit }, corner: corner || undefined, rotate,
        start: coordinates, pageOrigin: { x: page?.x ?? 0, y: page?.y ?? 0 }, pointerId: event.pointerId, zoom: props.zoom }
    }, ids, preserveSelection)
    requestDraw(); return
  }
  imageSelectionRequest++; emit('image-selection', undefined); imageContextPoint.value = undefined
  const tableHit = coordinates && hitTestDocumentTable(editingLayout.value, coordinates.x, coordinates.y)
  hoverTable.value = tableHit
  const point = pointFromEvent(event)
  const line = point && lineForPoint(editingLayout.value, point)
  if (tableHit && (!point || line?.tableCellId !== tableHit.cell.cellId)) {
    event.preventDefault()
    select(documentTableSelectionAtHit(editingModel.value, tableHit, props.selection))
    tableDragAnchor = tableHit; pointerId = event.pointerId
    dragSelectionSource = props.model
    scrollRef.value?.setPointerCapture(event.pointerId)
    focus()
    return
  }
  if (!point) return
  event.preventDefault()
  if (event.ctrlKey || event.metaKey) {
    const headingId = documentTocHeadingTarget(editingModel.value, blockAtPoint(editingModel.value, point))
    if (headingId) {
      emit('navigate-to-heading', headingId)
      return
    }
    const link = documentHyperlinkForSelection(editingModel.value, { anchor: point, focus: point })
    if (link) {
      emit('operation', 'doc.operation.click-hyper-link', { unitId: props.model.id, linkId: link.id,
        ...(props.host ? { segmentId: documentSegmentIdForHost(props.model, props.host) } : {}), source: props.model })
      return
    }
  }
  const latexHit = coordinates && hitTestDocumentLatex(editingModel.value, editingLayout.value, coordinates.x, coordinates.y)
  latexDragSelection = latexHit && !event.shiftKey ? latexHit.selection : undefined
  dragAnchor = event.shiftKey ? { ...props.selection.anchor } : point
  if (latexDragSelection) dragAnchor = latexDragSelection.anchor
  tableDragAnchor = tableHit
  pointerId = event.pointerId
  dragSelectionSource = props.model
  scrollRef.value?.setPointerCapture(event.pointerId)
  select(latexDragSelection ?? { anchor: dragAnchor,
    focus: latexHit ? documentLatexDragAnchor(editingModel.value, latexHit.selection, dragAnchor) : point })
  focus()
}

function contextMenu(event: MouseEvent) {
  if (!active || disposed) return
  const source = props.model, coordinates = coordinatesFromEvent(event), location = legacyStoryActive.value ? props.story : undefined
  const noteRegion = coordinates && hitTestDocumentNoteRegion(layout.value, coordinates.x, coordinates.y)
  if (noteRegion && !sameDocumentBlockHost(noteRegion.host, props.host)) {
    event.preventDefault()
    if (noteRegion.point && navigateCanvasNote({ type: 'note', kind: noteRegion.host.noteKind, noteId: noteRegion.host.noteId,
      toReference: false, point: noteRegion.point }, source)) {
      emit('context-menu', { x: event.clientX, y: event.clientY, point: noteRegion.point, source, host: noteRegion.host })
    }
    return
  }
  if (coordinates && hitTestDocumentNoteDecoration(layout.value, coordinates.x, coordinates.y)) { event.preventDefault(); return }
  if (!inspectionActive.value && props.host?.kind === 'note' && !noteRegion) {
    event.preventDefault()
    const point = coordinates && hitTestDocument(layout.value, coordinates.x, coordinates.y)
    if (point && navigateCanvasNote({ type: 'body', point }, source)) emit('context-menu', { x: event.clientX, y: event.clientY, point, source })
    return
  }
  if (location) {
    event.preventDefault()
    const target = storyAt(coordinates), page = layout.value.pageHeaderFooters[location.page]
    const storyKey = location.area === 'header' ? page?.headerStory : page?.footerStory
    if (storyComposing.value || !coordinates || !target || target.page !== location.page || target.area !== location.area || !page || page.isParityBlank || !storyKey) return
    const lines = documentStoryLayoutLines(layout.value, location), offset = hitTestDocumentStory(lines, coordinates.x, coordinates.y)
    const range = storyRange(), element = storyElement()
    if (!element || !storyEventIsCurrent()) return
    let selected = range
    if (range.start === range.end || offset < range.start || offset > range.end) {
      const line = lines.find(candidate => coordinates.y >= candidate.y && coordinates.y < candidate.y + candidate.height)
      const glyph = line?.glyphs.find(candidate => coordinates.x >= candidate.x && coordinates.x < candidate.x + candidate.width)
      selected = glyph && (props.storyValue ?? '')[glyph.offset] === '\uFFFC' ? { start: glyph.offset, end: glyph.end } : { start: offset, end: offset }
      element.setSelectionRange(selected.start, selected.end)
    }
    syncStorySelection()
    if (source !== props.model || !storyEventIsCurrent()) return
    const actual = storyRange()
    if (actual.start !== selected.start || actual.end !== selected.end) return
    emit('context-menu', { x: event.clientX, y: event.clientY, point: props.selection.focus, source,
      story: { sectionId: page.sectionId, story: storyKey, ...selected } })
    caretVisible.value = true
    requestDraw()
    return
  }
  const imageHit = coordinates && documentImageAtPoint(editingModel.value, editingLayout.value, coordinates.x, coordinates.y, event.altKey)
  if (imageHit && !composing.value && !storyComposing.value) {
    const root = scrollRef.value?.getBoundingClientRect()
    if (!root) return
    event.preventDefault(); cancelImageGesture(); cancelSelectionDrag()
    const ids = selectedImageIds.value.includes(imageHit.blockId) ? [...selectedImageIds.value] : [imageHit.blockId]
    selectCanvasImage(ids.at(-1)!, imageHit.page, () => {
      imageContextPoint.value = { x: Math.max(8, Math.min(event.clientX - root.left, viewport.value.width - 210)),
        y: Math.max(8, Math.min(event.clientY - root.top, viewport.value.height - 240)) }
    }, ids, selectedImageIds.value.includes(imageHit.blockId))
    requestDraw(); return
  }
  imageSelectionRequest++; cancelImageGesture(); emit('image-selection', undefined); imageContextPoint.value = undefined
  const tableHit = coordinates && hitTestDocumentTable(editingLayout.value, coordinates.x, coordinates.y)
  const point = pointFromEvent(event)
  if (tableHit && (!point || lineForPoint(editingLayout.value, point)?.tableCellId !== tableHit.cell.cellId)) {
    event.preventDefault()
    const selection = props.selection.tableSelection?.tableId === tableHit.fragment.tableId
      && tableHit.cell.row! >= props.selection.tableSelection.range.startRow && tableHit.cell.row! <= props.selection.tableSelection.range.endRow
      && tableHit.cell.columnIndex! >= props.selection.tableSelection.range.startColumn && tableHit.cell.columnIndex! <= props.selection.tableSelection.range.endColumn
      ? props.selection : documentTableSelectionAtHit(editingModel.value, tableHit, props.selection)
    if (selection !== props.selection) select(selection)
    if (source === props.model) emit('context-menu', { x: event.clientX, y: event.clientY, point: selection.focus, source, host: props.host })
    return
  }
  if (!point) return
  event.preventDefault()
  const latexHit = coordinates && hitTestDocumentLatex(editingModel.value, editingLayout.value, coordinates.x, coordinates.y)
  const selection = latexHit?.selection ?? documentSelectionForContextMenu(editingModel.value, props.selection, point)
  if (selection !== props.selection) select(selection)
  if (source !== props.model) return
  emit('context-menu', { x: event.clientX, y: event.clientY, point, source, host: props.host })
  caretVisible.value = true
  requestDraw()
}

function pointerMove(event: PointerEvent) {
  if (imageGesture.value?.pointerId === event.pointerId) { event.preventDefault(); moveImageGesture(event); startDragAutoscroll(event, 'image'); return }
  const coordinates = coordinatesFromEvent(event)
  updateNoteHover(coordinates)
  const tableHit = coordinates && !legacyStoryActive.value ? hitTestDocumentTable(editingLayout.value, coordinates.x, coordinates.y) : undefined
  if (!composing.value && !storyComposing.value) hoverTable.value = tableHit
  if (pointerId !== event.pointerId) return
  if (dragSelectionSource !== props.model) { cancelSelectionDrag(); return }
  updateDragSelection(event)
  startDragAutoscroll(event, 'selection')
}

function updateDragSelection(event: Pick<PointerEvent, 'clientX' | 'clientY' | 'pointerId'>) {
  if (pointerId !== event.pointerId || dragSelectionSource !== props.model) return
  const coordinates = coordinatesFromEvent(event)
  const tableHit = coordinates && !legacyStoryActive.value ? hitTestDocumentTable(editingLayout.value, coordinates.x, coordinates.y, tableDragAnchor?.fragment.tableId) : undefined
  if (storyDragAnchor) {
    const point = coordinatesFromEvent(event)
    if (point) {
      const offset = hitTestDocumentStory(documentStoryLayoutLines(layout.value, storyDragAnchor.location), point.x, point.y)
      focusStory(Math.min(storyDragAnchor.offset, offset), Math.max(storyDragAnchor.offset, offset))
    }
    return
  }
  if (!dragAnchor && !tableDragAnchor) return
  if (tableDragAnchor && tableHit && tableDragAnchor.fragment.tableId === tableHit.fragment.tableId
    && (!dragAnchor || tableDragAnchor.cell.cellId !== tableHit.cell.cellId || props.selection.tableSelection)) {
    const range = documentTableRangeBetweenHits(editingModel.value, tableDragAnchor, tableHit, props.selection)
    if (range) select(range)
    return
  }
  if (!dragAnchor) return
  const point = pointFromEvent(event)
  if (point) {
    const latexHit = coordinates && hitTestDocumentLatex(editingModel.value, editingLayout.value, coordinates.x, coordinates.y)
    const focus = latexHit ? documentLatexDragAnchor(editingModel.value, latexHit.selection, dragAnchor) : point
    const anchor = latexDragSelection ? documentLatexDragAnchor(editingModel.value, latexDragSelection, focus) : dragAnchor
    select({ anchor, focus })
  }
}

function cancelSelectionDrag() {
  if (pointerId !== undefined && scrollRef.value?.hasPointerCapture(pointerId)) scrollRef.value.releasePointerCapture(pointerId)
  pointerId = undefined; dragAnchor = undefined; latexDragSelection = undefined; tableDragAnchor = undefined; storyDragAnchor = undefined; dragSelectionSource = undefined
  if (dragScrollSource === 'selection') cancelDragAutoscroll()
}

function cancelDragAutoscroll() {
  cancelAnimationFrame(dragScrollFrame); dragScrollFrame = 0; dragScrollPointer = undefined; dragScrollSource = undefined
}

function startDragAutoscroll(pointer: DocumentTableDragPointer, source: 'selection' | 'overlay' | 'image') {
  dragScrollPointer = { clientX: pointer.clientX, clientY: pointer.clientY, pointerId: pointer.pointerId, ctrlKey: pointer.ctrlKey, metaKey: pointer.metaKey }
  dragScrollSource = source
  if (!dragScrollFrame) dragScrollFrame = requestAnimationFrame(scrollDrag)
}

function scrollDrag() {
  dragScrollFrame = 0
  const element = scrollRef.value, pointer = dragScrollPointer, source = dragScrollSource
  if (!element || !pointer || !source || !active || disposed) { cancelDragAutoscroll(); return }
  const rect = element.getBoundingClientRect(), edge = 26
  const step = (value: number, start: number, end: number) => value < start + edge ? -Math.min(24, (start + edge - value) * 0.65)
    : value > end - edge ? Math.min(24, (value - end + edge) * 0.65) : 0
  const x = step(pointer.clientX, rect.left, rect.right), y = step(pointer.clientY, rect.top, rect.bottom)
  if (!x && !y) return
  const left = element.scrollLeft, top = element.scrollTop
  element.scrollLeft += x; element.scrollTop += y
  if (left === element.scrollLeft && top === element.scrollTop) return
  syncViewport('scroll')
  if (source === 'selection') updateDragSelection(pointer)
  else if (source === 'image') moveImageGesture({ ...pointer, ...imagePointerModifiers })
  else tableOverlay.value?.continueGesture(pointer)
  if (dragScrollPointer) dragScrollFrame = requestAnimationFrame(scrollDrag)
}

function overlayDragPointer(pointer?: DocumentTableDragPointer) {
  if (pointer) startDragAutoscroll(pointer, 'overlay')
  else if (dragScrollSource === 'overlay') cancelDragAutoscroll()
}

function pointerUp(event: PointerEvent) {
  if (imageGesture.value?.pointerId === event.pointerId) {
    const gesture = imageGesture.value
    if (event.type === 'pointerup') moveImageGesture(event)
    const preview = imageGesturePreview.value
    if (event.type === 'pointerup' && preview && imageTargetCurrent(gesture.target)) {
      if ('members' in gesture && 'updates' in preview && preview.updates.length) emit('image-command', { target: gesture.target,
        command: { id: 'doc.command.transform-images', members: gesture.members, transform: preview.transform } })
      else if (!('members' in gesture) && 'patch' in preview) emit('image-command', { target: gesture.target,
        command: { id: 'doc.command.update-image', imageId: gesture.target.imageId, patch: preview.patch } })
    }
    cancelImageGesture(); cancelDragAutoscroll(); focus(); return
  }
  if (pointerId !== event.pointerId) return
  if (event.type === 'pointerup' && dragSelectionSource === props.model && (dragAnchor || tableDragAnchor)) {
    updateDragSelection(event)
    emit('selection-committed')
  }
  cancelSelectionDrag()
  focus()
}

function doubleClick(event: MouseEvent) {
  if (!active || disposed || composing.value || storyComposing.value) return
  const coordinates = coordinatesFromEvent(event), location = storyAt(coordinates)
  if (coordinates && (hitTestDocumentNoteGlyph(layout.value, coordinates.x, coordinates.y)
    || hitTestDocumentNoteDecoration(layout.value, coordinates.x, coordinates.y))) { event.preventDefault(); return }
  const noteRegion = coordinates && hitTestDocumentNoteRegion(layout.value, coordinates.x, coordinates.y)
  if (noteRegion && !sameDocumentBlockHost(noteRegion.host, props.host)) {
    event.preventDefault(); navigateCanvasNote({ type: 'note', kind: noteRegion.host.noteKind, noteId: noteRegion.host.noteId,
      toReference: false, point: noteRegion.point }); return
  }
  if (!inspectionActive.value && props.host?.kind === 'note' && !noteRegion) {
    const point = coordinates && hitTestDocument(layout.value, coordinates.x, coordinates.y)
    if (point) { event.preventDefault(); navigateCanvasNote({ type: 'body', point }) }
    return
  }
  const storyPage = location && layout.value.pageHeaderFooters[location.page], activeStory = location && props.host?.kind === 'story'
    && props.host.sectionId === storyPage?.sectionId && props.host.story === (location.area === 'header' ? storyPage?.headerStory : storyPage?.footerStory)
  if (location && coordinates && !activeStory) {
    event.preventDefault()
    const blockLayout = location.area === 'header' ? storyPage?.headerBlockLayout : storyPage?.footerBlockLayout
    if (blockLayout) {
      const point = hitTestDocument(documentLayoutForBlockHost(layout.value, blockLayout.host), coordinates.x, coordinates.y)
      emit('open-story', location, point)
      void nextTick(focus)
      return
    }
    const offset = hitTestDocumentStory(documentStoryLayoutLines(layout.value, location), coordinates.x, coordinates.y)
    const word = props.story?.page === location.page && props.story.area === location.area ? documentFormattingWordRange(props.storyValue ?? '', offset) : undefined
    emit('open-story', location)
    focusStory(word?.start ?? offset, word?.end ?? offset, location)
    return
  }
  const latexHit = coordinates && hitTestDocumentLatex(editingModel.value, editingLayout.value, coordinates.x, coordinates.y)
  const imageHit = coordinates && documentImageAtPoint(editingModel.value, editingLayout.value, coordinates.x, coordinates.y, event.altKey)
  if (imageHit) {
    event.preventDefault()
    selectCanvasImage(imageHit.blockId, imageHit.page, target => emit('image-operation', { target, mode: 'edit' }))
    return
  }
  if (latexHit) { event.preventDefault(); requestLatexEdit(latexHit); return }
  const point = pointFromEvent(event)
  if (!point) return
  const text = runText(runsForTarget(editingModel.value, point))
  const words = Array.from(new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text))
  const word = words.find((item) => point.offset >= item.index && point.offset <= item.index + item.segment.length)
  if (word) select({ anchor: { ...point, offset: word.index }, focus: { ...point, offset: word.index + word.segment.length } })
}

function ensureCaretVisible() {
  const element = scrollRef.value
  if (!element || !focused.value || pointerId !== undefined || dragScrollSource || selectedImageIds.value.length) return
  if (!displaySelection.value.tableSelection) {
    const code = documentCodeCaretScrollOffset(sourceEditingLayout.value, displaySelection.value.focus, codeScrollOffsets.value)
    if (code && Math.abs(code.offset - (codeScrollOffsets.value[code.key] ?? 0)) > 1e-6) codeScrollOffsets.value = { ...codeScrollOffsets.value, [code.key]: code.offset }
  }
  const cell = documentTableSelectionFocusCell(editingLayout.value, displaySelection.value, tableKeyboardCursor?.focus)
  const rectangle = cell && documentTableVisibleRect(cell)
  const caret = rectangle ? { ...rectangle, width: Math.min(30, rectangle.width), height: Math.min(30, rectangle.height) } : caretRect(editingLayout.value, displaySelection.value.focus)
  if (!caret) return
  const offset = documentRectScrollOffset(layout.value, viewport.value, caret)
  if (element.scrollLeft !== offset.left) element.scrollLeft = offset.left
  if (element.scrollTop !== offset.top) element.scrollTop = offset.top
  syncViewport('selection')
}

function bodyEventIsCurrent(event: Event): boolean {
  const input = inputElement()
  return active && !disposed && !legacyStoryActive.value && !(props.host?.kind === 'story' && props.storyReadonly) && !!input && input !== blockedBodyInput && event.target === input
}

function beforeInput(event: InputEvent) {
  if (!bodyEventIsCurrent(event)) { event.preventDefault(); return }
  if (composing.value || event.isComposing || event.inputType === 'insertCompositionText') return
  if (suppressCompositionInput && ['insertText', 'insertFromComposition'].includes(event.inputType)) { event.preventDefault(); suppressCompositionInput = false; return }
  if (event.inputType === 'insertText' && event.data != null) { event.preventDefault(); dispatchCommand({ id: 'doc.command.insert-text', text: event.data }) }
  else if (event.inputType === 'insertLineBreak' || event.inputType === 'insertParagraph') { event.preventDefault(); dispatchCommand({ id: 'doc.command.insert-text', text: event.inputType === 'insertLineBreak' ? '\u2028' : '\n' }) }
  else if (event.inputType.startsWith('delete')) {
    event.preventDefault()
    if (selectedImageIds.value.length && selectedImage.value) removeCanvasImages()
    else dispatchCommand({ id: 'doc.command.delete', direction: event.inputType.includes('Forward') ? 1 : -1 })
  }
  else if (event.inputType === 'historyUndo') { event.preventDefault(); emit('undo') }
  else if (event.inputType === 'historyRedo') { event.preventDefault(); emit('redo') }
}

function inputFallback(event: Event) {
  const input = inputElement()
  if (!bodyEventIsCurrent(event) || composing.value || !input?.value) return
  const text = input.value
  input.value = ''
  if (suppressCompositionInput) { suppressCompositionInput = false; return }
  dispatchCommand({ id: 'doc.command.insert-text', text })
}

function compositionStart(event: CompositionEvent) {
  if (!bodyEventIsCurrent(event)) return
  try {
    const context = props.permissionContext ?? { documentId: props.model.id ?? '', userId: 'local-user' }
    const inputSelection = prepareDocumentContentControlTextInputSelection(editingModel.value, props.selection, { controlId: explicitInputControlId() })
    const permission = resolveDocumentPermission(props.model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(editingModel.value, inputSelection))
    if (!permission.allowed) throw new Error('此选区不允许编辑')
    if (props.selection.tableSelection) {
      const state = documentTableCommandState(editingModel.value, props.selection, { id: 'doc.command.table-set-range-text', mode: 'replace', text: '' }, context)
      if (!state.available) throw new Error(state.reason ?? '此表格选区不允许输入')
    } else assertDocumentTextSelectionEditable(editingModel.value, inputSelection)
  } catch (error) {
    const input = inputElement()
    if (input) input.value = ''
    blockedBodyInput = input; bodyBridgeVersion.value++
    emit('notify', error instanceof Error ? error.message : String(error))
    return
  }
  compositionSession = props.captureInputSession?.(props.model)
  if ((props.captureInputSession && !compositionSession?.valid) || !bodyInputSessionCurrent()) {
    compositionSession = undefined
    const input = inputElement()
    if (input) input.value = ''
    blockedBodyInput = input
    bodyBridgeVersion.value++
    return
  }
  if (selectedImageIds.value.length) { imageSelectionRequest++; cancelImageGesture(); imageContextPoint.value = undefined; emit('image-selection', undefined) }
  composing.value = true
  suppressCompositionInput = false
  pendingText = ''
  compositionSelection.value = cloneDocumentSelection(props.selection)
  compositionRange = compositionSelection.value
  compositionModel = props.model
  compositionHost = props.host ? { ...props.host } : undefined
  compositionControlId = explicitInputControlId()
  compositionInput = inputElement()
  compositionMarks = { ...props.typingMarks }
}

function bodyCompositionCurrent() {
  return active && !disposed && !legacyStoryActive.value && sameDocumentBlockHost(compositionHost, props.host) && compositionModel === props.model && !!compositionRange
    && compositionControlId === explicitInputControlId()
    && compositionInput === inputElement()
    && sameTarget(compositionRange.anchor, props.selection.anchor) && compositionRange.anchor.offset === props.selection.anchor.offset
    && sameTarget(compositionRange.focus, props.selection.focus) && compositionRange.focus.offset === props.selection.focus.offset
    && sameDocumentTableViewSelection(compositionRange.tableSelection, props.selection.tableSelection)
    && bodyInputSessionCurrent()
}

function bodyInputSessionCurrent() {
  if (!props.captureInputSession) return true
  const current = props.captureInputSession()
  // A public draft starts only after its current Canvas model has captured a
  // session. The standalone editor has no public-session callback.
  return compositionSession?.valid === true && current?.valid === true
    && compositionSession.modelRevision === current.modelRevision && compositionSession.generation === current.generation
    && current.location.kind === 'body' && sameDocumentInputLocation(current.location, compositionSession.location)
    && sameDocumentInputLocation(current.location, documentInputLocation(props.selection, undefined, props.host))
}

function cancelBodyComposition(refocus = true) {
  const hadComposition = composing.value
  composing.value = false
  compositionSelection.value = null
  compositionRange = undefined
  compositionModel = undefined
  compositionInput = undefined
  compositionMarks = {}
  compositionSession = undefined
  compositionHost = undefined
  compositionControlId = undefined
  pendingText = ''
  suppressCompositionInput = false
  if (hadComposition) { preview.value = null; if (active && !disposed) relayout() }
  const input = inputElement()
  if (input) input.value = ''
  if (hadComposition) {
    blockedBodyInput = input
    bodyBridgeVersion.value++
    if (refocus) void nextTick(() => { if (active && !disposed && !legacyStoryActive.value) inputElement()?.focus({ preventScroll: true }) })
  }
}

function compositionUpdate(event: CompositionEvent) {
  if (!bodyEventIsCurrent(event)) return
  if (!composing.value) return
  if (!bodyCompositionCurrent() || !compositionRange) { cancelBodyComposition(); return }
  pendingText = event.data
  if (!pendingText) {
    preview.value = null; compositionSelection.value = cloneDocumentSelection(compositionRange); relayout(); return
  }
  const model = cloneDocument(props.model)
  try {
    const selection = applyDocumentEditorCommand(model, { unitId: props.model.id ?? '', selection: compositionRange,
      command: { id: 'doc.command.insert-text', text: pendingText, marks: compositionMarks,
        host: props.host, activeContentControlId: compositionControlId } }, props.permissionContext ?? { documentId: props.model.id ?? '', userId: 'local-user' })
    preview.value = model
    compositionSelection.value = selection
    relayout()
  } catch (error) { cancelBodyComposition(); emit('notify', error instanceof Error ? error.message : String(error)) }
}

function compositionEnd(event: CompositionEvent) {
  if (!bodyEventIsCurrent(event)) return
  if (!composing.value || !bodyCompositionCurrent()) { cancelBodyComposition(); return }
  const text = event.data || pendingText
  const marks = compositionMarks
  composing.value = false
  compositionRange = undefined
  compositionModel = undefined
  compositionInput = undefined
  compositionMarks = {}
  compositionSession = undefined
  compositionHost = undefined
  compositionControlId = undefined
  compositionSelection.value = null
  preview.value = null
  const input = inputElement()
  if (input) input.value = ''
  if (text) dispatchCommand({ id: 'doc.command.insert-text', text, marks })
  pendingText = ''
  suppressCompositionInput = true
  queueMicrotask(() => { suppressCompositionInput = false })
  relayout()
}

function move(direction: 'left' | 'right' | 'up' | 'down' | 'home' | 'end', extend: boolean, byWord: boolean) {
  const model = editingModel.value
  if (props.selection.tableSelection) {
    const result = navigateDocumentTableSelection(model, props.selection, direction, extend, byWord, tableKeyboardCursor)
    if (!result) return
    tableKeyboardCursor = result.cursor
    const affinitySelection = { ...result.selection, tableSelection: { ...result.selection.tableSelection!,
      ...(props.selection.tableSelection.pageAffinity === undefined ? {} : { pageAffinity: props.selection.tableSelection.pageAffinity }) } }
    const cell = documentTableSelectionFocusCell(editingLayout.value, affinitySelection, result.cursor.focus)
    const next = cell ? { ...result.selection, tableSelection: { ...result.selection.tableSelection!, pageAffinity: cell.page },
      anchor: { ...result.selection.anchor, pageAffinity: cell.page }, focus: { ...result.selection.focus, pageAffinity: cell.page } } : result.selection
    select(next)
    return
  }
  tableKeyboardCursor = undefined
  const ordered = orderedSelection(model, props.selection)
  let point = { ...props.selection.focus }
  if (!extend && !selectionCollapsed(props.selection) && (direction === 'left' || direction === 'right')) point = { ...(direction === 'left' ? ordered.start : ordered.end) }
  else if (direction === 'home' || direction === 'end') {
    if (byWord) {
      const targets = visibleDocumentTextTargets(model)
      const target = direction === 'home' ? targets[0] : targets.at(-1)
      if (target) point = { ...target, offset: direction === 'home' ? 0 : runText(runsForTarget(model, target)).length }
    } else {
      point = documentPhysicalLineNavigation(editingLayout.value, point, direction) ?? point
    }
  } else if (direction === 'left' || direction === 'right') {
    const text = runText(runsForTarget(model, point))
    const offsets = byWord ? [...Array.from(new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text), (part) => part.index), text.length] : graphemeOffsets(text)
    const offset = direction === 'left' ? offsets.filter((value) => value < point.offset).at(-1) : offsets.find((value) => value > point.offset)
    if (offset != null) point.offset = offset
    else {
      const targets = visibleDocumentTextTargets(model)
      const index = targets.findIndex((target) => sameTarget(target, point))
      const target = targets[index + (direction === 'left' ? -1 : 1)]
      if (target) point = { ...target, offset: direction === 'left' ? runText(runsForTarget(model, target)).length : 0 }
    }
  } else {
    point = documentPhysicalLineNavigation(editingLayout.value, point, direction) ?? point
  }
  select({ anchor: extend ? props.selection.anchor : point, focus: point })
}

function keydown(event: KeyboardEvent) {
  if (!bodyEventIsCurrent(event)) return
  if (event.isComposing || composing.value || event.keyCode === 229) return
  const command = event.ctrlKey || event.metaKey
  const key = event.key.toLowerCase()
  if (!['Control', 'Meta', 'Shift', 'Alt'].includes(event.key)) nextPasteMode = command && event.shiftKey && key === 'v' ? 'text' : 'source'
  if (selectedImageIds.value.length && selectedImage.value) {
    if (event.key === 'Escape') {
      event.preventDefault(); imageSelectionRequest++
      const dragging = !!imageGesture.value; cancelImageGesture(); imageContextPoint.value = undefined
      if (!dragging) emit('image-selection', undefined)
      return
    }
    if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); removeCanvasImages(); return }
    if (event.key === 'Enter') { event.preventDefault(); imageOperation('edit'); return }
    const delta = ({ ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] } as Record<string, [number, number]>)[event.key]
    if (selectedImageIds.value.length > 1 && delta && !command) {
      event.preventDefault()
      if (!imageReadOnly.value && imageGeometryState.value.available) {
        const step = event.shiftKey ? 10 : 1
        imageCommand({ id: 'doc.command.transform-images', members: selectedImageMembers.value, transform: { x: delta[0] * step, y: delta[1] * step } })
      }
      return
    }
    if ((selectedImageIds.value.length === 1 && (selectedImage.value.image.wrappingStyle ?? 'inline') === 'inline' || selectedImageIds.value.length > 1)
      && (delta || event.key === 'Home' || event.key === 'End')) {
      event.preventDefault(); cancelImageGesture(); emit('image-selection', undefined); imageContextPoint.value = undefined
      const direction = ({ ArrowLeft: 'left', ArrowRight: 'right', ArrowUp: 'up', ArrowDown: 'down', Home: 'home', End: 'end' } as const)[event.key as 'ArrowLeft']
      move(direction, event.shiftKey, command); return
    }
    if (delta && !command && selectedImageRectangle.value && imageEditable()) {
      event.preventDefault(); const target = imageTarget()!, rectangle = selectedImageRectangle.value, page = layout.value.pages[rectangle.page], step = event.shiftKey ? 10 : 1
      const result = documentImageGestureResult({ target, image: selectedImage.value.image, rectangle, start: { x: 0, y: 0 },
        pageOrigin: { x: page?.x ?? 0, y: page?.y ?? 0 }, pointerId: -1, zoom: props.zoom }, delta[0] * step, delta[1] * step)
      imageCommand({ id: 'doc.command.update-image', imageId: target.imageId, patch: result.patch }); return
    }
  }
  const paragraphShortcut = command && event.shiftKey && !event.altKey ? paragraphShortcutCommands[event.code] : undefined
  if (paragraphShortcut) {
    if (selectedImageIds.value.length || document.activeElement !== inputElement()) return
    event.preventDefault(); event.stopPropagation()
    dispatchCommand(paragraphShortcut)
    return
  }
  if (command && !event.altKey && (event.shiftKey ? event.code === 'KeyX' : event.code === 'Comma' || event.code === 'Period')) {
    if (selectedImageIds.value.length || document.activeElement !== inputElement()) return
    event.preventDefault(); event.stopPropagation()
    if (event.code === 'KeyX') { emit('toggle-format', 'strike'); return }
    emit('toggle-format', event.code === 'Comma' ? 'subscript' : 'superscript')
    return
  }
  if (event.key === 'Escape') { tableOverlay.value?.cancelInteraction(); cancelSelectionDrag(); tablePreview.value = undefined }
  if (key === 'f9' && !command && !event.altKey && !event.shiftKey) {
    if (document.activeElement !== inputElement()) return
    event.preventDefault(); event.stopPropagation()
    emit('refresh-fields', props.model, undefined, cloneDocumentSelection(props.selection))
    return
  }
  const activeBlock = blockAtPoint(editingModel.value, props.selection.focus)
  const codeBlock = documentStructuredBlockAtSelection(editingModel.value, props.selection, 'code') ?? activeBlock
  if (!props.selection.tableSelection && codeBlock && codeBlockData(codeBlock)) {
    if (command && event.key === 'Enter') {
      event.preventDefault()
      emit('code-block-exit', codeBlock.id)
      return
    }
    if (event.key === 'Tab' && !command) {
      event.preventDefault()
      dispatchCommand({ id: 'doc.command.indent-code-block', blockId: codeBlock.id, direction: event.shiftKey ? -1 : 1 })
      return
    }
  }
  if (event.key === ' ' && !command && activeBlock?.type === 'checklist' && props.selection.focus.offset === 0 && selectionCollapsed(props.selection)) {
    event.preventDefault()
    dispatchCommand({ id: 'doc.command.toggle-checklist-item', blockId: activeBlock.id })
  } else if (command && key === 'a') {
    event.preventDefault()
    imageSelectionRequest++; cancelImageGesture(); emit('image-selection', undefined); imageContextPoint.value = undefined
    const targets = visibleDocumentTextTargets(editingModel.value)
    const first = targets[0]
    const last = targets.at(-1)
    if (first && last) select({ anchor: { ...first, offset: 0 }, focus: { ...last, offset: runText(runsForTarget(editingModel.value, last)).length } })
  } else if (command && (key === 'z' || key === 'y')) {
    event.preventDefault()
    if (key === 'y' || event.shiftKey) emit('redo')
    else emit('undo')
  }
  else if (command && key === 's') { event.preventDefault(); event.stopPropagation(); emit('save') }
  else if (command && ['b', 'i', 'u'].includes(key)) {
    event.preventDefault()
    const field = key === 'b' ? 'bold' : key === 'i' ? 'italic' : 'underline'
    emit('toggle-format', field)
  } else if (event.key === 'Enter') {
    event.preventDefault()
    dispatchCommand(command && !props.host ? { id: 'doc.command.insert-page-break' } : { id: 'doc.command.insert-text', text: event.shiftKey ? '\u2028' : '\n' })
  } else if (event.key === 'Backspace' || event.key === 'Delete') { event.preventDefault(); dispatchCommand({ id: 'doc.command.delete', direction: event.key === 'Backspace' ? -1 : 1 }) }
  else if (event.key === 'Tab') {
    event.preventDefault()
    const tab = documentTabCommand(editingModel.value, props.selection, event.shiftKey)
    if (tab) dispatchCommand(tab)
  } else {
    const direction = ({ ArrowLeft: 'left', ArrowRight: 'right', ArrowUp: 'up', ArrowDown: 'down', Home: 'home', End: 'end' } as const)[event.key as 'ArrowLeft']
    if (direction) { event.preventDefault(); move(direction, event.shiftKey, command) }
  }
}

function copy(event: ClipboardEvent, cut = false) {
  if (!bodyEventIsCurrent(event) || composing.value || !selectedImage.value && selectionCollapsed(props.selection)) return
  event.preventDefault()
  if (event.clipboardData) emit('clipboard-copy', event.clipboardData, cut)
}
function paste(event: ClipboardEvent) {
  if (!bodyEventIsCurrent(event)) return
  event.preventDefault()
  const mode = nextPasteMode
  nextPasteMode = 'source'
  if (!composing.value && event.clipboardData) emit('clipboard-paste', {
    html: event.clipboardData.getData('text/html'),
    text: event.clipboardData.getData('text/plain'),
  }, mode)
}

function scrollToPage(index: number) {
  const page = layout.value.pages[index]
  if (page && scrollRef.value) scrollRef.value.scrollTo({ top: Math.max(0, page.y * props.zoom / 100 - 16), behavior: 'smooth' })
}
function scrollToBlock(id: string) {
  const line = editingLayout.value.lines.find((candidate) => (candidate.target.cellBlockId ?? candidate.target.blockId) === id)
  if (line && scrollRef.value) scrollRef.value.scrollTo({ top: Math.max(0, line.y * props.zoom / 100 - 28), behavior: 'smooth' })
}
function captureScrollOffset(): DocumentScrollOffset | undefined {
  const element = scrollRef.value
  return element ? { left: element.scrollLeft, top: element.scrollTop } : undefined
}
function restoreScrollOffset(offset: DocumentScrollOffset) {
  const element = scrollRef.value
  if (!element || !active || disposed || !Number.isFinite(offset.left) || !Number.isFinite(offset.top)) return
  // A captured body viewport takes precedence over caret/zoom reveals queued
  // by the same host transition. The next user selection captures this version.
  viewportRestoreVersion++
  element.scrollLeft = Math.max(0, offset.left); element.scrollTop = Math.max(0, offset.top)
  syncViewport('selection')
}
function revealStory(target: StoryLocation) {
  const entry = layout.value.pageHeaderFooters[target.page], el = scrollRef.value
  if (!entry || entry.isParityBlank || !el) return
  const y = (target.area === 'header' ? entry.headerTop : entry.footerBottom) * props.zoom / 100
  if (y < el.scrollTop + 30 || y > el.scrollTop + el.clientHeight - 30) el.scrollTop = Math.max(0, y - el.clientHeight / 2)
  focus()
}
/** SDT navigation closes value UI without canceling the document's text/IME session. */
function cancelContentControlPopup() { pendingContentControlFocus = undefined; contentControlOverlay.value?.cancelInteraction() }
function cancelInput(refocus = true) { imageSelectionRequest++; nextPasteMode = 'source'; clearNoteHover(); cancelImageGesture(); imageContextPoint.value = undefined; structuredOverlay.value?.cancelInteraction(); cancelContentControlPopup(); tableOverlay.value?.cancelInteraction(); cancelSelectionDrag(); tableKeyboardCursor = undefined; cancelBodyComposition(refocus); cancelStoryInput(refocus) }
defineExpose({ focus, focusStory, revealStory, storyRange, cancelInput, cancelContentControlPopup, clearImageSelection, executeContentControlCommand, scrollToPage, scrollToBlock, captureScrollOffset, restoreScrollOffset, backend, getLayout: () => layout.value, isBodyComposing: () => composing.value })
</script>

<template>
  <div class="document-canvas-viewport">
    <div ref="hostRef" class="document-surface" aria-hidden="true" />
    <div ref="scrollRef" class="document-scroll" role="textbox" aria-label="文档画布编辑区" aria-multiline="true" tabindex="-1" @scroll="onScroll" @pointerdown="pointerDown" @pointermove="pointerMove" @pointerleave="() => { hoverTable = undefined; deferNoteHoverHide() }" @pointerup="pointerUp" @pointercancel="pointerUp" @dblclick="doubleClick" @contextmenu="contextMenu">
      <div class="document-scroll-spacer" :style="spacerStyle" />
    </div>
    <DocumentImageOverlay :image="selectedImage?.image" :rectangle="imageGesturePreview?.rectangle ?? selectedImageFrame" :selected-count="selectedImageIds.length"
      :layout="layout" :viewport="viewport" :disabled="imageReadOnly" :geometry-disabled="!imageGeometryState.available" :geometry-reason="imageGeometryState.reason" :arrange-disabled="!imageArrangeAvailable"
      :inactive="legacyStoryActive || !contentControlActive || !!imageGesture" :context-point="imageContextPoint" @close-context="closeImageContext" @focus="closeImageContext" @edit="imageOperation"
      @remove="removeCanvasImages" @reset="resetCanvasImage"
      @arrange="arrangeCanvasImages" @align="alignCanvasImages" />
    <DocumentTableOverlay ref="tableOverlay" :model="displayEditingModel" :layout="editingLayout" :selection="displaySelection" :viewport="viewport" :permission="permissionContext" :hover="hoverTable"
      :floating-toolbar="tableFloatingToolbar" :disabled="composing || storyComposing || (host?.kind === 'story' && storyReadonly)" :story-active="legacyStoryActive" @selection="select" @command="command => dispatchCommand(command)" @operation="(id, params) => emit('operation', id, params)" @preview="value => tablePreview = value" @drag-pointer="overlayDragPointer" @focus="focus" @notify="value => emit('notify', value)" />
    <DocumentHeadingCollapseOverlay v-if="!host && !inspectionActive && !legacyStoryActive" :model="model" :layout="layout" :viewport="viewport"
      :permission="permissionContext" :capture-input-session="captureInputSession" :disabled="!contentControlActive || composing || storyComposing"
      @command="dispatchCommand" />
    <DocumentStructuredOverlay ref="structuredOverlay" :source="model" :model="displayEditingModel" :layout="editingLayout" :source-layout="sourceEditingLayout" :selection="displaySelection" :viewport="viewport"
      :host="host" :permission="permissionContext" :capture-input-session="captureInputSession" :code-scroll-offsets="codeScrollOffsets"
      :inactive="legacyStoryActive || !contentControlActive" :disabled="composing || storyComposing || (host?.kind === 'story' && storyReadonly)"
      @command="dispatchStructuredCommand" @operation="dispatchStructuredOperation" @code-scroll="scrollCode" @focus="focusStructured" />
    <DocumentContentControlOverlay ref="contentControlOverlay" :source="model" :model="displayEditingModel" :layout="layout" :geometries="contentControlGeometries"
      :selection="displaySelection" :viewport="viewport" :active-control-id="activeContentControl?.id" :popup-mode="activeContentControlPopupMode" :host="host" :permission="permissionContext"
      :capture-input-session="captureInputSession" :inactive="legacyStoryActive || !contentControlActive" :disabled="composing || storyComposing || (host?.kind === 'story' && storyReadonly)"
      @selection="selectContentControl" @command="dispatchContentControlCommand" @focus="focusContentControl" @notify="message => emit('notify', message)" />
    <Popover :open="!!noteHover" placement="bottomLeft" :title="noteHover ? `${noteHover.preview.kind === 'footnote' ? '脚注' : '尾注'} ${noteHover.preview.label}` : ''">
      <template #content>
        <div v-if="noteHover" class="document-note-preview" data-document-note-preview :style="{ color: token.colorText }" @pointerenter="keepNoteHover" @pointerleave="deferNoteHoverHide">
          <p>{{ noteHover.preview.text || '空注释' }}</p>
          <span v-if="noteHover.preview.truncated" :style="{ color: token.colorTextSecondary }">…</span>
          <Button type="link" size="small" @click="navigateNotePreview">转到注释</Button>
        </div>
      </template>
      <span class="document-note-preview-anchor" :style="noteHoverStyle" aria-hidden="true" />
    </Popover>
    <Input.TextArea :key="bodyBridgeVersion" ref="inputRef" class="document-input-bridge" :style="inputStyle" aria-label="文档文字输入" autocapitalize="off" autocomplete="off" spellcheck="false" variant="borderless" @beforeinput="beforeInput" @input="inputFallback" @keydown="keydown" @compositionstart="compositionStart" @compositionupdate="compositionUpdate" @compositionend="compositionEnd" @focus="bodyFocusChanged" @blur="bodyFocusChanged" @copy="copy($event)" @cut="copy($event, true)" @paste="paste" />
    <Input.TextArea v-if="legacyStoryActive && story" :key="`${model.id}:${story.page}:${story.area}:${storyBridgeVersion}`" ref="storyRef" class="document-story-bridge" :style="storyStyle" :aria-label="story.area === 'header' ? '编辑页眉' : '编辑页脚'" :value="storyInputValue" :readonly="storyReadonly" :rows="1" :change-on-composing="true" variant="borderless" @keydown="storyKeys" @keyup="syncStorySelection" @select="syncStorySelection" @beforeinput="storyBeforeInput" @compositionstart="storyCompositionStart" @compositionend="storyCompositionEnd" @change="storyInputChanged" @focus="storyFocusChanged" @blur="storyFocusChanged" />
    <Alert v-if="error" class="canvas-error" type="error" show-icon :message="error" role="alert" />
    <span class="backend-announcement" aria-live="polite">{{ backend }}</span>
  </div>
</template>

<style scoped>
.document-canvas-viewport { position: relative; min-width: 0; min-height: 0; flex: 1; overflow: hidden; background: var(--document-background, #eef1f5); }
.document-surface { position: absolute; inset: 0; pointer-events: none; }
.document-scroll { position: absolute; inset: 0; overflow: auto; overscroll-behavior: contain; touch-action: none; direction: ltr; cursor: text; outline: none; }
.document-scroll-spacer { pointer-events: none; }
.document-note-preview-anchor { position: absolute; pointer-events: none; }
.document-note-preview { max-width: 320px; }
.document-note-preview p { display: -webkit-box; -webkit-box-orient: vertical; -webkit-line-clamp: 6; max-height: 9em; overflow: hidden; margin: 0 0 6px; white-space: pre-wrap; overflow-wrap: anywhere; }
.document-story-bridge { position: absolute; z-index: 3; width: 1px; height: 1px; min-height: 1px; padding: 0; border: 0; outline: 0; opacity: 0.01; resize: none; overflow: hidden; caret-color: transparent; color: transparent; background: transparent; white-space: pre; pointer-events: none; }
.document-story-bridge :deep(textarea) { width: 1px; height: 1px; min-height: 1px; padding: 0; border: 0; opacity: 0.01; resize: none; overflow: hidden; caret-color: transparent; color: transparent; background: transparent; white-space: pre; pointer-events: none; }
.document-input-bridge { position: absolute; z-index: 2; width: 1px; height: 1px; min-height: 1px; padding: 0; border: 0; outline: 0; opacity: 0.01; resize: none; overflow: hidden; caret-color: transparent; color: transparent; background: transparent; white-space: pre; }
.document-input-bridge :deep(textarea) { position: absolute; inset: 0; width: 1px; height: 1px; min-height: 1px; padding: 0; border: 0; outline: 0; opacity: 0.01; resize: none; overflow: hidden; caret-color: transparent; color: transparent; background: transparent; white-space: pre; }
.backend-announcement { position: absolute; width: 1px; height: 1px; overflow: hidden; clip-path: inset(50%); }
.canvas-error { position: absolute; top: 20px; left: 20px; right: 20px; }
</style>
