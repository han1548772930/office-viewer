<script setup lang="ts">
import { computed, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue'
import type { DocumentModel } from '../../domain/model'
import type { RenderSurface } from '../../rendering/contracts'
import { createRenderSurface } from '../../rendering/createRenderSurface'
import { blockAtPoint, executeDocumentCommand, graphemeOffsets, type DocumentCommand } from '../../features/document/commands/document-commands'
import { caretRect, createTextMeasurer, emptyLayout, hitTestDocument, layoutDocument, lineForPoint, type DocumentLayout } from '../../features/document/layout/document-layout'
import { cloneDocument, marksAtPoint, orderedSelection, runText, runsForTarget, sameTarget, selectedPlainText, selectionCollapsed, textTargets, type DocumentPoint, type DocumentSelection } from '../../features/document/model/document-model'
import { documentPasteCommand } from '../../features/document/services/document-clipboard'
import { sceneFromDocument, sceneOffset, type DocumentViewport } from '../../features/document/rendering/document-scene'
import { documentRectScrollOffset, documentViewportLocation, zoomedDocumentScrollOffset, type DocumentViewportLocation } from '../../features/document/layout/document-viewport'

const props = defineProps<{ model: DocumentModel; selection: DocumentSelection; zoom: number }>()
const emit = defineEmits<{
  command: [command: DocumentCommand]
  'toggle-format': [field: 'bold' | 'italic' | 'underline']
  'update:selection': [selection: DocumentSelection]
  layout: [layout: DocumentLayout]
  'active-page': [page: number]
  'viewport-change': [location: DocumentViewportLocation]
  undo: []
  redo: []
  save: []
  notify: [message: string]
}>()
const scrollRef = ref<HTMLDivElement | null>(null)
const hostRef = ref<HTMLDivElement | null>(null)
const inputRef = ref<HTMLTextAreaElement | null>(null)
const layout = shallowRef<DocumentLayout>(emptyLayout())
const preview = shallowRef<DocumentModel | null>(null)
const compositionSelection = shallowRef<DocumentSelection | null>(null)
const composing = ref(false)
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
let pointerId: number | undefined
let suppressCompositionInput = false
let pendingText = ''
let disposed = false
let active = true
const measurer = createTextMeasurer()
const displayModel = computed(() => preview.value ?? props.model)
const displaySelection = computed(() => compositionSelection.value ?? props.selection)
const spacerStyle = computed(() => ({ width: `${Math.max(viewport.value.width, layout.value.width * props.zoom / 100)}px`, height: `${layout.value.height * props.zoom / 100}px` }))
const inputStyle = computed(() => {
  const caret = caretRect(layout.value, displaySelection.value.focus)
  const offset = sceneOffset(layout.value, viewport.value)
  return {
    left: `${offset.x + (caret?.x ?? 0) * props.zoom / 100}px`,
    top: `${offset.y + (caret?.y ?? 0) * props.zoom / 100}px`,
    fontSize: `${Math.max(12, (caret?.height ?? 18) * props.zoom / 100)}px`,
  }
})

watch(() => props.model, () => { if (!composing.value) preview.value = null; relayout() }, { deep: false })
watch(() => props.selection, () => { caretVisible.value = true; requestDraw(); void nextTick(ensureCaretVisible) }, { deep: true })
watch(() => props.zoom, (zoom) => {
  const offset = zoomedDocumentScrollOffset(layout.value, viewport.value, zoom)
  viewport.value = { ...viewport.value, zoom }
  requestDraw()
  void nextTick(() => {
    if (disposed || props.zoom !== zoom || !scrollRef.value) return
    scrollRef.value.scrollLeft = offset.left
    scrollRef.value.scrollTop = offset.top
    syncViewport('zoom')
  })
})

onActivated(() => { active = true; void nextTick(resize) })
onDeactivated(() => {
  active = false
  focused.value = false
  inputRef.value?.blur()
  if (pointerId != null && scrollRef.value?.hasPointerCapture(pointerId)) scrollRef.value.releasePointerCapture(pointerId)
  pointerId = undefined
  dragAnchor = undefined
  cancelAnimationFrame(frame)
  frame = 0
  surface?.resize(0, 0, window.devicePixelRatio || 1)
  lastSurfaceSize = undefined
})

onMounted(async () => {
  if (!hostRef.value || !scrollRef.value) return
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
    resize()
    relayout()
    document.fonts?.ready.then(() => { if (!disposed) relayout() })
    blink = setInterval(() => { caretVisible.value = !caretVisible.value; requestDraw() }, 530)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
})

onUnmounted(() => {
  disposed = true
  resizeObserver?.disconnect()
  dprQuery?.removeEventListener('change', onDevicePixelRatioChange)
  window.removeEventListener('resize', onViewportResize)
  window.visualViewport?.removeEventListener('resize', onViewportResize)
  if (blink) clearInterval(blink)
  cancelAnimationFrame(frame)
  surface?.dispose()
  images.forEach((image) => { image.onload = null; image.onerror = null })
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
  layout.value = layoutDocument(displayModel.value, measurer)
  const sources = new Set(layout.value.images.map((image) => image.src))
  for (const [src, image] of images) {
    if (sources.has(src)) continue
    image.onload = null
    image.onerror = null
    images.delete(src)
  }
  for (const image of layout.value.images) {
    if (!image.src || images.has(image.src)) continue
    const source = new Image()
    source.crossOrigin = image.src.startsWith('data:') || image.src.startsWith('blob:') ? null : 'anonymous'
    source.onload = () => { if (!disposed && images.get(image.src) === source) requestDraw() }
    source.onerror = () => { if (!disposed && images.get(image.src) === source) { emit('notify', `图片无法载入：${image.alt || image.src.slice(0, 80)}`); requestDraw() } }
    images.set(image.src, source)
    source.src = image.src
  }
  emit('layout', layout.value)
  publishViewport('layout')
  void nextTick(() => { if (!disposed) syncViewport('layout') })
  requestDraw()
}

function requestDraw() {
  if (frame || disposed || !active) return
  frame = requestAnimationFrame(() => {
    frame = 0
    surface?.render(sceneFromDocument(displayModel.value, layout.value, viewport.value, { selection: displaySelection.value, focused: focused.value, caretVisible: caretVisible.value, composing: composing.value, images }))
  })
}

function onScroll() {
  syncViewport('scroll')
}

function syncViewport(reason: NonNullable<DocumentViewportLocation['reason']>) {
  if (!scrollRef.value) return
  viewport.value = { ...viewport.value, scrollLeft: scrollRef.value.scrollLeft, scrollTop: scrollRef.value.scrollTop }
  publishViewport(reason)
  requestDraw()
}

function publishViewport(reason: NonNullable<DocumentViewportLocation['reason']>) {
  if (!active || disposed) return
  const location = documentViewportLocation(layout.value, viewport.value)
  if (location.pageIndex != null) emit('active-page', location.pageIndex)
  emit('viewport-change', { ...location, reason })
}

function pointFromEvent(event: PointerEvent | MouseEvent): DocumentPoint | undefined {
  const element = scrollRef.value
  if (!element) return undefined
  const bounds = element.getBoundingClientRect()
  const offset = sceneOffset(layout.value, viewport.value)
  return hitTestDocument(layout.value, (event.clientX - bounds.left - offset.x) / (props.zoom / 100), (event.clientY - bounds.top - offset.y) / (props.zoom / 100))
}

function select(selection: DocumentSelection) {
  emit('update:selection', selection)
  caretVisible.value = true
  requestDraw()
}

function focus() { inputRef.value?.focus({ preventScroll: true }); focused.value = true; requestDraw() }

function pointerDown(event: PointerEvent) {
  if (event.button !== 0 || composing.value) return
  const point = pointFromEvent(event)
  if (!point) return
  event.preventDefault()
  if (event.ctrlKey || event.metaKey) {
    const href = marksAtPoint(props.model, point).href
    if (href) { window.open(href, '_blank', 'noopener,noreferrer'); return }
  }
  dragAnchor = event.shiftKey ? { ...props.selection.anchor } : point
  pointerId = event.pointerId
  scrollRef.value?.setPointerCapture(event.pointerId)
  select({ anchor: dragAnchor, focus: point })
  focus()
}

function pointerMove(event: PointerEvent) {
  if (pointerId !== event.pointerId || !dragAnchor) return
  const point = pointFromEvent(event)
  if (point) select({ anchor: dragAnchor, focus: point })
  const element = scrollRef.value
  if (element) {
    const rect = element.getBoundingClientRect()
    if (event.clientY < rect.top + 18) element.scrollTop -= 24
    if (event.clientY > rect.bottom - 18) element.scrollTop += 24
  }
}

function pointerUp(event: PointerEvent) {
  if (pointerId !== event.pointerId) return
  if (scrollRef.value?.hasPointerCapture(event.pointerId)) scrollRef.value.releasePointerCapture(event.pointerId)
  pointerId = undefined
  dragAnchor = undefined
  focus()
}

function doubleClick(event: MouseEvent) {
  const point = pointFromEvent(event)
  if (!point) return
  const text = runText(runsForTarget(props.model, point))
  const words = Array.from(new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text))
  const word = words.find((item) => point.offset >= item.index && point.offset <= item.index + item.segment.length)
  if (word) select({ anchor: { ...point, offset: word.index }, focus: { ...point, offset: word.index + word.segment.length } })
}

function ensureCaretVisible() {
  const element = scrollRef.value
  const caret = caretRect(layout.value, displaySelection.value.focus)
  if (!element || !caret || !focused.value) return
  const offset = documentRectScrollOffset(layout.value, viewport.value, caret)
  if (element.scrollLeft !== offset.left) element.scrollLeft = offset.left
  if (element.scrollTop !== offset.top) element.scrollTop = offset.top
  syncViewport('selection')
}

function beforeInput(event: InputEvent) {
  if (composing.value || event.isComposing || event.inputType === 'insertCompositionText') return
  if (suppressCompositionInput && ['insertText', 'insertFromComposition'].includes(event.inputType)) { event.preventDefault(); suppressCompositionInput = false; return }
  if (event.inputType === 'insertText' && event.data != null) { event.preventDefault(); emit('command', { id: 'doc.command.insert-text', text: event.data }) }
  else if (event.inputType === 'insertLineBreak' || event.inputType === 'insertParagraph') { event.preventDefault(); emit('command', { id: 'doc.command.insert-text', text: event.inputType === 'insertLineBreak' ? '\u2028' : '\n' }) }
  else if (event.inputType.startsWith('delete')) { event.preventDefault(); emit('command', { id: 'doc.command.delete', direction: event.inputType.includes('Forward') ? 1 : -1 }) }
  else if (event.inputType === 'historyUndo') { event.preventDefault(); emit('undo') }
  else if (event.inputType === 'historyRedo') { event.preventDefault(); emit('redo') }
}

function inputFallback() {
  if (composing.value || !inputRef.value?.value) return
  const text = inputRef.value.value
  inputRef.value.value = ''
  if (suppressCompositionInput) { suppressCompositionInput = false; return }
  emit('command', { id: 'doc.command.insert-text', text })
}

function compositionStart() {
  composing.value = true
  suppressCompositionInput = false
  pendingText = ''
  compositionSelection.value = { anchor: { ...props.selection.anchor }, focus: { ...props.selection.focus } }
}

function compositionUpdate(event: CompositionEvent) {
  if (!composing.value) return
  pendingText = event.data
  const model = cloneDocument(props.model)
  const selection = executeDocumentCommand(model, props.selection, { id: 'doc.command.insert-text', text: pendingText })
  preview.value = model
  compositionSelection.value = selection
  relayout()
}

function compositionEnd(event: CompositionEvent) {
  const text = event.data || pendingText
  composing.value = false
  compositionSelection.value = null
  preview.value = null
  if (inputRef.value) inputRef.value.value = ''
  if (text) emit('command', { id: 'doc.command.insert-text', text })
  pendingText = ''
  suppressCompositionInput = true
  queueMicrotask(() => { suppressCompositionInput = false })
  relayout()
}

function move(direction: 'left' | 'right' | 'up' | 'down' | 'home' | 'end', extend: boolean, byWord: boolean) {
  const model = props.model
  const ordered = orderedSelection(model, props.selection)
  let point = { ...props.selection.focus }
  if (!extend && !selectionCollapsed(props.selection) && (direction === 'left' || direction === 'right')) point = { ...(direction === 'left' ? ordered.start : ordered.end) }
  else if (direction === 'home' || direction === 'end') {
    if (byWord) {
      const targets = textTargets(model)
      const target = direction === 'home' ? targets[0] : targets.at(-1)
      if (target) point = { ...target, offset: direction === 'home' ? 0 : runText(runsForTarget(model, target)).length }
    } else {
      const line = lineForPoint(layout.value, point)
      if (line) point.offset = direction === 'home' ? line.start : line.end
    }
  } else if (direction === 'left' || direction === 'right') {
    const text = runText(runsForTarget(model, point))
    const offsets = byWord ? [...Array.from(new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text), (part) => part.index), text.length] : graphemeOffsets(text)
    const offset = direction === 'left' ? offsets.filter((value) => value < point.offset).at(-1) : offsets.find((value) => value > point.offset)
    if (offset != null) point.offset = offset
    else {
      const targets = textTargets(model)
      const index = targets.findIndex((target) => sameTarget(target, point))
      const target = targets[index + (direction === 'left' ? -1 : 1)]
      if (target) point = { ...target, offset: direction === 'left' ? runText(runsForTarget(model, target)).length : 0 }
    }
  } else {
    const line = lineForPoint(layout.value, point)
    const caret = caretRect(layout.value, point)
    if (line && caret) {
      const index = layout.value.lines.indexOf(line)
      const nextLine = layout.value.lines[index + (direction === 'up' ? -1 : 1)]
      if (nextLine) point = hitTestDocument({ ...layout.value, lines: [nextLine] }, caret.x, nextLine.y + nextLine.height / 2) ?? point
    }
  }
  select({ anchor: extend ? props.selection.anchor : point, focus: point })
}

function keydown(event: KeyboardEvent) {
  if (event.isComposing || composing.value || event.keyCode === 229) return
  const command = event.ctrlKey || event.metaKey
  const key = event.key.toLowerCase()
  if (command && key === 'a') {
    event.preventDefault()
    const targets = textTargets(props.model)
    const first = targets[0]
    const last = targets.at(-1)
    if (first && last) select({ anchor: { ...first, offset: 0 }, focus: { ...last, offset: runText(runsForTarget(props.model, last)).length } })
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
    emit('command', command ? { id: 'doc.command.insert-page-break' } : { id: 'doc.command.insert-text', text: event.shiftKey ? '\u2028' : '\n' })
  } else if (event.key === 'Backspace' || event.key === 'Delete') { event.preventDefault(); emit('command', { id: 'doc.command.delete', direction: event.key === 'Backspace' ? -1 : 1 }) }
  else if (event.key === 'Tab') {
    event.preventDefault()
    if (props.selection.focus.row != null) emit('command', { id: 'doc.command.table-tab', direction: event.shiftKey ? -1 : 1 })
    else {
      const block = blockAtPoint(props.model, props.selection.focus)
      emit('command', { id: 'doc.command.format-paragraph', patch: { indent: Math.max(0, (block?.style.indent ?? 0) + (event.shiftKey ? -1 : 1)) } })
    }
  } else {
    const direction = ({ ArrowLeft: 'left', ArrowRight: 'right', ArrowUp: 'up', ArrowDown: 'down', Home: 'home', End: 'end' } as const)[event.key as 'ArrowLeft']
    if (direction) { event.preventDefault(); move(direction, event.shiftKey, command) }
  }
}

function copy(event: ClipboardEvent) {
  if (selectionCollapsed(props.selection)) return
  event.preventDefault()
  event.clipboardData?.setData('text/plain', selectedPlainText(props.model, props.selection))
}

function cut(event: ClipboardEvent) { copy(event); if (!selectionCollapsed(props.selection)) emit('command', { id: 'doc.command.insert-text', text: '' }) }
function paste(event: ClipboardEvent) {
  event.preventDefault()
  const command = documentPasteCommand({ html: event.clipboardData?.getData('text/html'), text: event.clipboardData?.getData('text/plain') ?? '' })
  if (command) emit('command', command)
}

function scrollToPage(index: number) {
  const page = layout.value.pages[index]
  if (page && scrollRef.value) scrollRef.value.scrollTo({ top: Math.max(0, page.y * props.zoom / 100 - 16), behavior: 'smooth' })
}
function scrollToBlock(id: string) {
  const line = layout.value.lines.find((candidate) => candidate.target.blockId === id)
  if (line && scrollRef.value) scrollRef.value.scrollTo({ top: Math.max(0, line.y * props.zoom / 100 - 28), behavior: 'smooth' })
}
defineExpose({ focus, scrollToPage, scrollToBlock, backend, getLayout: () => layout.value })
</script>

<template>
  <div class="document-canvas-viewport">
    <div ref="hostRef" class="document-surface" aria-hidden="true" />
    <div ref="scrollRef" class="document-scroll" role="textbox" aria-label="文档画布编辑区" aria-multiline="true" tabindex="-1" @scroll="onScroll" @pointerdown="pointerDown" @pointermove="pointerMove" @pointerup="pointerUp" @pointercancel="pointerUp" @dblclick="doubleClick">
      <div class="document-scroll-spacer" :style="spacerStyle" />
    </div>
    <textarea ref="inputRef" class="document-input-bridge" :style="inputStyle" aria-label="文档文字输入" autocapitalize="off" autocomplete="off" spellcheck="false" @beforeinput="beforeInput" @input="inputFallback" @keydown="keydown" @compositionstart="compositionStart" @compositionupdate="compositionUpdate" @compositionend="compositionEnd" @focus="focused = true; requestDraw()" @blur="focused = false; requestDraw()" @copy="copy" @cut="cut" @paste="paste" />
    <span v-if="error" class="canvas-error" role="alert">{{ error }}</span>
    <span class="backend-announcement" aria-live="polite">{{ backend }}</span>
  </div>
</template>

<style scoped>
.document-canvas-viewport { position: relative; min-width: 0; min-height: 0; flex: 1; overflow: hidden; background: #eef1f5; }
.document-surface { position: absolute; inset: 0; pointer-events: none; }
.document-scroll { position: absolute; inset: 0; overflow: auto; overscroll-behavior: contain; touch-action: none; cursor: text; outline: none; }
.document-scroll-spacer { pointer-events: none; }
.document-input-bridge { position: absolute; z-index: 2; width: 1px; height: 1px; min-height: 1px; padding: 0; border: 0; outline: 0; opacity: 0.01; resize: none; overflow: hidden; caret-color: transparent; color: transparent; background: transparent; white-space: pre; }
.backend-announcement { position: absolute; width: 1px; height: 1px; overflow: hidden; clip-path: inset(50%); }
.canvas-error { position: absolute; top: 20px; left: 20px; right: 20px; padding: 12px; border: 1px solid #efb3b3; border-radius: 4px; color: #a02f2f; background: #fff3f3; }
</style>
