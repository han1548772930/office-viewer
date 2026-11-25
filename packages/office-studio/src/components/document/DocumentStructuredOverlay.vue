<script setup lang="ts">
import { computed, onDeactivated, onUnmounted, ref, shallowRef, watch } from 'vue'
import { Slider, theme } from 'antdv-next'
import type { DocumentModel } from '../../domain/model'
import { documentBlockAncestry, findDocumentBlock, runText, sameTarget, selectionCollapsed, type DocumentSelection } from '../../features/document/model/document-model'
import { isDocumentTextBlock } from '../../features/document/model/document-structural-blocks'
import { documentCodeFontFamily } from '../../features/document/code-blocks/document-code-fonts'
import { documentStructuredBlockAtSelection } from '../../features/document/commands/document-structure-commands'
import type { DocumentInputSession } from '../../features/document/model/document-input-session'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../../features/document/model/document-block-host'
import { resolveDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../../features/document/permissions/document-permissions'
import { lineForPoint } from '../../features/document/layout/document-layout'
import type { DocumentLayout, LayoutContainerDecoration, LayoutRect } from '../../features/document/layout/document-layout-types'
import { sceneOffset, type DocumentViewport } from '../../features/document/rendering/document-scene'
import { documentCodeScrollRegions, type DocumentCodeScrollOffsets, type DocumentCodeScrollRegion } from '../../features/document/layout/document-code-scroll'
import { documentTableVisibleRect } from '../../features/document/layout/document-table-interaction'
import { captureDocumentStructuredCanvasTarget, documentStructuredCanvasTargetCurrent, resizeDocumentColumnBoundary,
  type DocumentStructuredCanvasTarget, type DocumentStructuredCanvasCommandRequest, type DocumentStructuredCanvasOperationRequest,
  type DocumentColumnResizeResult } from '../../features/document/controllers/document-structured-canvas'
import type { DocumentStructureRibbonContext, DocumentStructuredParagraphMenuContext } from './document-structure-ribbon'
import DocumentStructuredPopMenus from './DocumentStructuredPopMenus.vue'
import { selectedLatexFormula } from '../../features/document/latex/latex-formula'
import { documentLatexSelectionState } from '../../features/document/commands/document-latex-commands'
import { documentLatexContextItems, documentLatexMountedMenuItem } from './document-latex-menu'
import { useI18n } from 'vue-i18n'

const props = defineProps<{ source: DocumentModel; model: DocumentModel; layout: DocumentLayout; sourceLayout: DocumentLayout; selection: DocumentSelection; viewport: DocumentViewport;
  host?: DocumentAuxiliaryBlockHost; permission?: DocumentPermissionContext; captureInputSession?: (source?: DocumentModel) => DocumentInputSession | undefined;
  codeScrollOffsets: DocumentCodeScrollOffsets; disabled?: boolean; inactive?: boolean }>()
const emit = defineEmits<{ command: [request: DocumentStructuredCanvasCommandRequest]; operation: [request: DocumentStructuredCanvasOperationRequest];
  'code-scroll': [request: { source: DocumentModel; key: string; value: number }]; focus: [target: DocumentStructuredCanvasTarget] }>()
const { token } = theme.useToken()
const { t } = useI18n({ useScope: 'global' })
const root = ref<HTMLDivElement>()
const popMenus = ref<{ cancelInteraction: (refocus?: boolean) => void; hasInteraction: () => boolean }>()
const offset = computed(() => sceneOffset(props.layout, props.viewport)), scale = computed(() => props.viewport.zoom / 100)
const activeBlock = computed(() => props.inactive ? undefined : documentStructuredBlockAtSelection(props.model, props.selection))
const activeTitle = computed(() => ({ quote: '引用块', callout: '标注块', code: '代码块', columns: '多栏组' } as Record<string, string>)[activeBlock.value?.type ?? ''] ?? '高级块')
const columnBlock = computed(() => props.inactive ? undefined : documentStructuredBlockAtSelection(props.model, props.selection, 'columns'))
const focusId = computed(() => props.selection.tableSelection?.tableId ?? props.selection.focus.cellBlockId ?? props.selection.focus.blockId)
const columnId = computed(() => columnBlock.value?.columns?.columns.find(column => !!documentBlockAncestry({ blocks: column.blocks }, focusId.value))?.id)
const context = computed(() => props.permission ?? { documentId: props.source.id ?? '', userId: 'local-user' })
const disabled = computed(() => !!props.disabled || !!props.inactive || !resolveDocumentPermission(props.source.permissionRules ?? [], context.value, 'edit',
  permissionTargetsForSelection(props.model, props.selection)).allowed)
const inputBusy = computed(() => !!props.disabled || !!props.inactive)
const style = computed(() => ({ '--structured-bg': token.value.colorBgElevated, '--structured-border': token.value.colorBorder,
  '--structured-primary': token.value.colorPrimary, '--structured-text': token.value.colorText, '--structured-shadow': token.value.boxShadowSecondary }))
const visible = (rect: LayoutRect) => offset.value.x + (rect.x + rect.width) * scale.value > 0 && offset.value.x + rect.x * scale.value < props.viewport.width
  && offset.value.y + (rect.y + rect.height) * scale.value > 0 && offset.value.y + rect.y * scale.value < props.viewport.height
const containers = computed(() => (props.layout.containers ?? []).filter(container => sameDocumentBlockHost(container.host, props.host) && visible(container)))
const activeContainer = computed(() => containers.value.find(container => container.blockId === activeBlock.value?.id && container.page === props.selection.focus.pageAffinity)
  ?? containers.value.find(container => container.blockId === activeBlock.value?.id))
const toolbarStyle = computed(() => {
  const rectangle = activeContainer.value
  if (!rectangle) return { display: 'none' }
  return { left: `${Math.max(8, Math.min(offset.value.x + rectangle.x * scale.value, props.viewport.width - 290))}px`,
    top: `${Math.max(8, offset.value.y + rectangle.y * scale.value - 34)}px`, maxWidth: `${Math.max(120, props.viewport.width - 16)}px` }
})
const paragraphLine = computed(() => props.inactive || props.selection.tableSelection ? undefined : lineForPoint(props.layout, props.selection.focus))
const latexState = computed(() => documentLatexSelectionState(props.model, props.selection, context.value, disabled.value))
const floatLatex = computed(() => {
  const line = paragraphLine.value
  if (!line || !visible(line) || props.inactive || props.selection.tableSelection || selectionCollapsed(props.selection)
    || activeBlock.value || !sameTarget(props.selection.anchor, props.selection.focus) || selectedLatexFormula(props.model, props.selection)
    || documentBlockAncestry(props.model, focusId.value)?.some(block => !!block.metadata.univerBodyBlockRangeSource)) return undefined
  return { item: documentLatexMountedMenuItem('float-text-style', latexState.value, t), style: props.viewport.width <= 640
    ? { left: '8px', bottom: '8px', maxWidth: 'calc(100% - 16px)' }
    : { left: `${Math.max(8, Math.min(offset.value.x + line.x * scale.value, props.viewport.width - 148))}px`,
      top: `${Math.max(8, offset.value.y + line.y * scale.value - 34)}px`, maxWidth: `${Math.max(100, props.viewport.width - 16)}px` } }
})
const mobileLatex = computed(() => {
  const formula = props.viewport.width <= 640 && !props.inactive ? selectedLatexFormula(props.model, props.selection) : undefined
  if (!formula) return undefined
  return { items: documentLatexContextItems(formula, latexState.value, t).filter(item => item.id !== 'doc-context-latex-text'),
    style: { left: '8px', bottom: '8px', maxWidth: 'calc(100% - 16px)' } }
})
const paragraphStyle = computed(() => {
  const line = paragraphLine.value
  if (!line || !visible(line)) return { display: 'none' }
  const gutter = line.codeBlock?.gutterX ?? line.x
  return { left: `${Math.max(4, offset.value.x + gutter * scale.value - 30)}px`, top: `${offset.value.y + line.y * scale.value}px` }
})
function capture(): DocumentStructuredCanvasTarget {
  return captureDocumentStructuredCanvasTarget(props.source, props.selection, props.host, props.captureInputSession?.(props.source))
}
function current(target: DocumentStructuredCanvasTarget): boolean {
  return documentStructuredCanvasTargetCurrent(target, { source: props.source, selection: props.selection, host: props.host,
    session: props.captureInputSession?.(), requireSession: !!props.captureInputSession, disabled: disabled.value })
}
const activeMenu = computed<(DocumentStructureRibbonContext & { title: string }) | undefined>(() => {
  const block = activeBlock.value
  return activeContainer.value && block && (block.type === 'quote' || block.type === 'callout' || block.type === 'code' || block.type === 'columns')
    ? { blockId: block.id, type: block.type, title: activeTitle.value, disabled: disabled.value, code: block.code,
      fontFamily: block.code ? documentCodeFontFamily(block) : undefined, columns: block.columns, columnId: block.type === 'columns' ? columnId.value : undefined } : undefined
})
const belowInColumns = (id: string) => !!documentBlockAncestry(props.model, id)?.slice(0, -1).some(block => block.type === 'columns')
const paragraphContext = computed<DocumentStructuredParagraphMenuContext>(() => {
  const block = findDocumentBlock(props.model, focusId.value), active = activeBlock.value, afterBlockId = active?.id ?? focusId.value
  return { paragraphId: focusId.value, afterBlockId, disabled: disabled.value, inColumns: !!columnBlock.value, belowInColumns: belowInColumns(afterBlockId),
    activeBlockId: active?.id, activeType: activeMenu.value?.type,
    codeAllowed: active?.children ? active.children.every(isDocumentTextBlock) : !!block && isDocumentTextBlock(block),
    columns: columnBlock.value?.columns, columnId: columnId.value, columnGroupId: columnBlock.value?.id }
})
const paragraphControl = computed(() => {
  const line = paragraphLine.value, block = findDocumentBlock(props.model, focusId.value)
  return line && visible(line) && selectionCollapsed(props.selection) && block && isDocumentTextBlock(block)
    ? { key: `paragraph:${focusId.value}`, style: paragraphStyle.value, context: paragraphContext.value, empty: runText(block.runs).length === 0 } : undefined
})
const contentInsert = computed(() => {
  if (!paragraphControl.value) return undefined
  const active = activeBlock.value, block = findDocumentBlock(props.model, focusId.value)
  const candidates = active ? containers.value.filter(container => container.blockId === active.id && container.last)
    : props.layout.lines.filter(line => (line.target.cellBlockId ?? line.target.blockId) === focusId.value && line.end === runText(block?.runs ?? []).length && visible(line))
  const anchor = candidates.find(value => value.page === props.selection.focus.pageAffinity) ?? candidates[0]
  if (!anchor) return undefined
  return { key: `content:${paragraphContext.value.afterBlockId}`, context: paragraphContext.value, style: {
    left: `${Math.max(4, offset.value.x + anchor.x * scale.value - 28)}px`, top: `${offset.value.y + (anchor.y + anchor.height) * scale.value + 2}px` } }
})
const tableControls = computed(() => {
  if (props.inactive) return []
  const tableId = props.selection.tableSelection?.tableId ?? documentBlockAncestry(props.model, focusId.value)?.reverse().find(block => block.type === 'table')?.id
  return (props.layout.tableFragments ?? []).filter(fragment => {
    const rect = documentTableVisibleRect(fragment)
    return fragment.tableId === tableId && !fragment.continuedAfter && sameDocumentBlockHost(fragment.host, props.host) && rect.width > 0 && rect.height > 0 && visible(rect)
  })
    .map(fragment => { const rect = documentTableVisibleRect(fragment); return { key: `${fragment.tableId}:${fragment.page}:${fragment.sectionFragment}:${fragment.y}`, context: {
      paragraphId: focusId.value, afterBlockId: fragment.tableId, disabled: disabled.value, inColumns: belowInColumns(fragment.tableId), belowInColumns: belowInColumns(fragment.tableId),
    }, style: { left: `${Math.max(4, Math.min(props.viewport.width - 30, offset.value.x + (rect.x + rect.width) * scale.value - 24))}px`,
      top: `${offset.value.y + (rect.y + rect.height) * scale.value + 4}px` } } })
})

const columnFrames = computed(() => containers.value.filter(container => container.blockId === columnBlock.value?.id && container.columnGroup && !container.columnGroup.stacked))
function clipped(rect: LayoutContainerDecoration): LayoutRect {
  if (!rect.clip) return rect
  const x = Math.max(rect.x, rect.clip.x), y = Math.max(rect.y, rect.clip.y)
  return { x, y, width: Math.max(0, Math.min(rect.x + rect.width, rect.clip.x + rect.clip.width) - x),
    height: Math.max(0, Math.min(rect.y + rect.height, rect.clip.y + rect.clip.height) - y) }
}
const handles = computed(() => columnFrames.value.flatMap(container => container.columnGroup!.widths.slice(0, -1).map((_, index) => ({ container, index,
  x: container.x + container.columnGroup!.widths.slice(0, index + 1).reduce((sum, width) => sum + width, 0) + container.columnGroup!.gap * (index + 0.5) })))
  .filter(handle => { const rect = clipped(handle.container); return rect.width > 0 && rect.height > 0 && handle.x >= rect.x && handle.x <= rect.x + rect.width }))
function handleStyle(handle: typeof handles.value[number]) {
  const rect = clipped(handle.container)
  return { left: `${offset.value.x + handle.x * scale.value - 5}px`, top: `${offset.value.y + rect.y * scale.value}px`, height: `${rect.height * scale.value}px` }
}
interface Gesture { target: DocumentStructuredCanvasTarget; container: LayoutContainerDecoration; index: number; startX: number;
  pointerId: number; element: HTMLElement; preview?: DocumentColumnResizeResult; moved: boolean }
const gesture = shallowRef<Gesture>()
function xAt(event: Pick<PointerEvent, 'clientX'>): number | undefined {
  const rect = root.value?.getBoundingClientRect()
  return rect ? (event.clientX - rect.left - offset.value.x) / scale.value : undefined
}
function startGesture(container: LayoutContainerDecoration, index: number, event: PointerEvent) {
  if (event.button !== 0 || disabled.value || !(event.currentTarget instanceof HTMLElement) || !columnBlock.value?.columns) return
  const target = capture(), x = xAt(event)
  if (!current(target) || x === undefined || !resizeDocumentColumnBoundary(columnBlock.value.columns, container, index, 0)) return
  event.preventDefault(); event.stopPropagation(); cancelInteraction()
  event.currentTarget.setPointerCapture(event.pointerId)
  gesture.value = { target, container, index, startX: x, pointerId: event.pointerId, element: event.currentTarget, moved: false }
  window.addEventListener('pointermove', moveGesture)
  window.addEventListener('pointerup', finishGesture)
  window.addEventListener('pointercancel', cancelGesture)
}
function moveGesture(event: PointerEvent) {
  const drag = gesture.value, x = xAt(event)
  if (!drag || drag.pointerId !== event.pointerId || x === undefined) return
  event.preventDefault(); event.stopPropagation()
  if (!current(drag.target)) { cancelGesture(); return }
  const group = findDocumentBlock(props.model, drag.container.blockId)?.columns
  if (!group) { cancelGesture(); return }
  gesture.value = { ...drag, preview: resizeDocumentColumnBoundary(group, drag.container, drag.index, x - drag.startX), moved: drag.moved || Math.abs(x - drag.startX) >= 1 }
}
function finishGesture(event: PointerEvent) {
  moveGesture(event)
  const drag = gesture.value
  if (!drag || drag.pointerId !== event.pointerId) return
  event.preventDefault(); event.stopPropagation(); cancelGesture()
  if (!drag.moved || !drag.preview || !current(drag.target)) return
  emit('command', { ...drag.target, command: { id: 'doc.command.update-column-group', blockId: drag.container.blockId, widthRatios: drag.preview.widthRatios } })
}
function cancelGesture() {
  const drag = gesture.value; gesture.value = undefined
  window.removeEventListener('pointermove', moveGesture); window.removeEventListener('pointerup', finishGesture); window.removeEventListener('pointercancel', cancelGesture)
  if (drag?.element.hasPointerCapture(drag.pointerId)) drag.element.releasePointerCapture(drag.pointerId)
}
function keyboardResize(container: LayoutContainerDecoration, index: number, event: KeyboardEvent) {
  if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight' || disabled.value || !columnBlock.value?.columns) return
  event.preventDefault(); event.stopPropagation()
  const target = capture(), result = resizeDocumentColumnBoundary(columnBlock.value.columns, container, index, (event.key === 'ArrowLeft' ? -1 : 1) * (event.shiftKey ? 24 : 8))
  if (result && current(target)) emit('command', { ...target, command: { id: 'doc.command.update-column-group', blockId: container.blockId, widthRatios: result.widthRatios } })
}
const guideStyle = computed(() => {
  const drag = gesture.value
  return drag?.preview ? { left: `${offset.value.x + drag.preview.guideX * scale.value}px`, top: `${offset.value.y + drag.container.y * scale.value}px`, height: `${drag.container.height * scale.value}px` } : { display: 'none' }
})

const codeRegions = computed(() => props.inactive ? [] : documentCodeScrollRegions(props.sourceLayout).filter(region => region.maxScroll > 0 && sameDocumentBlockHost(region.container.host, props.host) && visible(region.container)))
const scrollTarget = shallowRef<DocumentStructuredCanvasTarget>()
function captureScroll() { scrollTarget.value = capture() }
function codeStyle(region: DocumentCodeScrollRegion) {
  return { left: `${offset.value.x + region.clip.x * scale.value}px`, top: `${offset.value.y + (region.container.y + region.container.height - 20) * scale.value}px`,
    width: `${Math.max(1, region.clip.width * scale.value)}px` }
}
function scrollCode(region: DocumentCodeScrollRegion, value: number | number[]) {
  const target = scrollTarget.value
  if (typeof value !== 'number' || !target || props.disabled || props.inactive || target.source !== props.source || !sameDocumentBlockHost(target.host, props.host)) return
  emit('code-scroll', { source: target.source, key: region.key, value: Math.min(region.maxScroll, Math.max(0, value)) })
}
function finishCodeScroll() { const target = scrollTarget.value; scrollTarget.value = undefined; if (target && current(target)) emit('focus', target) }
function escape(event: KeyboardEvent) { if (event.key === 'Escape' && (gesture.value || popMenus.value?.hasInteraction())) { event.preventDefault(); event.stopPropagation(); popMenus.value?.cancelInteraction(true); cancelInteraction() } }
function cancelInteraction() { popMenus.value?.cancelInteraction(); scrollTarget.value = undefined; cancelGesture() }
watch(() => [props.source, props.model, JSON.stringify(props.selection), JSON.stringify(props.host), props.disabled, props.inactive, props.viewport.zoom, props.viewport.width, props.viewport.height], cancelInteraction, { flush: 'sync' })
onDeactivated(cancelInteraction)
onUnmounted(() => { cancelInteraction(); window.removeEventListener('keydown', escape, true) })
watch(root, value => { if (value) window.addEventListener('keydown', escape, true) }, { flush: 'post' })
defineExpose({ cancelInteraction })
</script>

<template>
  <div ref="root" class="structured-overlay" :style="style" @pointerdown.stop @dblclick.stop @contextmenu.stop>
    <DocumentStructuredPopMenus ref="popMenus" :active="activeMenu" :toolbar-style="toolbarStyle" :paragraph="paragraphControl" :content-insert="contentInsert" :tables="tableControls"
      :latex-state="latexState" :float-latex="floatLatex" :mobile-latex="mobileLatex"
      :mobile="viewport.width <= 640" :disabled="disabled" :capture-target="capture" :target-current="current"
      @command="request => emit('command', request)" @operation="request => emit('operation', request)" @focus="target => emit('focus', target)" />
    <!-- Canvas hit/resize handles are editor kernel geometry; menus and forms use antdv-next. -->
    <div v-for="handle in handles" :key="`${handle.container.blockId}:${handle.container.page}:${handle.container.sectionFragment}:${handle.index}`"
      class="structured-column-handle" :class="{ 'structured-column-handle--disabled': disabled }" :style="handleStyle(handle)"
      :data-document-column-group="handle.container.blockId" :data-document-column-boundary="handle.index"
      role="separator" aria-orientation="vertical" :aria-label="`调整第 ${handle.index + 1}、${handle.index + 2} 栏宽度`" :aria-disabled="disabled" :tabindex="disabled ? -1 : 0"
      @pointerdown="event => startGesture(handle.container, handle.index, event)" @keydown="event => keyboardResize(handle.container, handle.index, event)" />
    <div v-if="gesture?.preview" class="structured-column-guide" :style="guideStyle" aria-hidden="true" />
    <div v-for="region in codeRegions" :key="`${region.key}:${region.container.page}:${region.container.sectionFragment}:${region.container.y}`"
      class="structured-code-scroll" :style="codeStyle(region)" data-document-code-scroll :data-document-code-block="region.container.blockId" @pointerdown="captureScroll" @focusin="captureScroll">
      <Slider :value="Math.min(region.maxScroll, codeScrollOffsets[region.key] ?? 0)" :min="0" :max="region.maxScroll" :step="1" :disabled="inputBusy"
        aria-label="代码块水平滚动" aria-label-for-handle="代码块水平滚动" :tooltip="{ open: false }" @update:value="value => scrollCode(region, value)" @change-complete="finishCodeScroll" />
    </div>
  </div>
</template>

<style scoped>
.structured-overlay { position: absolute; inset: 0; z-index: 4; pointer-events: none; color: var(--structured-text); }
.structured-column-handle { position: absolute; width: 10px; cursor: col-resize; pointer-events: auto; touch-action: none; }
.structured-column-handle::after { content: ''; position: absolute; left: 4px; width: 1px; height: 100%; background: var(--structured-border); }
.structured-column-handle:hover::after, .structured-column-handle:focus-visible::after { width: 2px; background: var(--structured-primary); }
.structured-column-handle:focus-visible { outline: 1px solid var(--structured-primary); }
.structured-column-handle--disabled { pointer-events: none; }
.structured-column-guide { position: absolute; width: 2px; background: var(--structured-primary); }
.structured-code-scroll { position: absolute; pointer-events: auto; }
.structured-code-scroll :deep(.ant-slider) { margin: 0 5px; }
</style>
