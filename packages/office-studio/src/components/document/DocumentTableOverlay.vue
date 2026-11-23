<script setup lang="ts">
import { computed, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import { Button, Checkbox, Divider, Select, Tooltip, theme } from 'antdv-next'
import { AlignCenter, AlignLeft, AlignRight, Bold, GripHorizontal, GripVertical, Italic, Move, PaintBucket, Palette, Plus, RotateCcw, Strikethrough, Table2, Underline } from 'lucide-vue-next'
import type { DocumentModel } from '../../domain/model'
import { blockAtPoint, collapsedSelection, documentBlockAncestry, findDocumentBlock, marksAtPoint, runText, tableForBlock, type DocumentSelection } from '../../features/document/model/document-model'
import { documentTableGrid, documentTableWholeRange, type DocumentTableCellRange } from '../../features/document/model/document-table-grid'
import { documentTableCellBorder, documentTableCellFill, documentTableCellMargin, documentTableRowProperties } from '../../features/document/model/document-table-style'
import { documentTableSelectedTextTargets } from '../../features/document/model/document-table-selection'
import { documentTableCommandState, documentTableSelection } from '../../features/document/commands/table-commands'
import type { DocumentTableCommand } from '../../features/document/commands/table-command-types'
import type { DocumentCommand } from '../../features/document/commands/document-commands'
import { documentTextFormattingState } from '../../features/document/services/document-text-formatting'
import { documentClearFormattingState } from '../../features/document/services/document-clear-formatting'
import { resolveDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../../features/document/permissions/document-permissions'
import { hitTestDocument, lineForPoint } from '../../features/document/layout/document-layout'
import type { DocumentLayout, LayoutTableFragment } from '../../features/document/layout/document-layout-types'
import { documentTableAxisSelection, documentTableFragmentColumns, documentTableFragmentKey, documentTableFragmentRows, documentTableLogicalHit, documentTableVisibleRect, type DocumentTableDragPointer, type DocumentTableGesturePreview, type DocumentTableHit } from '../../features/document/layout/document-table-interaction'
import { sceneOffset, type DocumentViewport } from '../../features/document/rendering/document-scene'
import { DOCUMENT_TABLE_BORDER_LABELS } from './document-table-ribbon'
import DocumentColorControl from './DocumentColorControl.vue'
import DocumentTableColorControl from './DocumentTableColorControl.vue'
import type { DocumentColorParams } from '../../features/document/services/document-colors'
import { documentTableCellBlocks } from '../../features/document/model/document-table-continuations'

const props = defineProps<{ model: DocumentModel; layout: DocumentLayout; selection: DocumentSelection; viewport: DocumentViewport;
  permission?: DocumentPermissionContext; hover?: DocumentTableHit; floatingToolbar?: boolean; disabled?: boolean; storyActive?: boolean }>()
const emit = defineEmits<{ selection: [selection: DocumentSelection]; command: [command: DocumentCommand]; operation: [id: string, params?: unknown];
  preview: [preview?: DocumentTableGesturePreview]; 'drag-pointer': [pointer?: DocumentTableDragPointer]; focus: []; notify: [message: string] }>()
const { token } = theme.useToken()
const root = ref<HTMLDivElement>()
const popup = ref<string>()
const context = computed(() => props.permission ?? { documentId: props.model.id ?? '', userId: 'local-user' })
const selected = computed(() => props.storyActive ? undefined : documentTableSelection(props.model, props.selection))
const tableId = computed(() => selected.value?.blockId ?? (!props.storyActive ? props.hover?.fragment.tableId : undefined))
const source = computed(() => { const block = tableId.value && findDocumentBlock(props.model, tableId.value); return block && tableForBlock(block) })
const offset = computed(() => sceneOffset(props.layout, props.viewport))
const scale = computed(() => props.viewport.zoom / 100)
const overlayStyle = computed(() => ({ '--table-overlay-primary': token.value.colorPrimary, '--table-overlay-background': token.value.colorBgElevated,
  '--table-overlay-border': token.value.colorBorder, '--table-overlay-shadow': token.value.boxShadowSecondary, '--table-overlay-text': token.value.colorText }))
const visible = (rectangle: { x: number; y: number; width: number; height: number }) => {
  const k = scale.value, off = offset.value
  return off.x + (rectangle.x + rectangle.width) * k > 0 && off.x + rectangle.x * k < props.viewport.width
    && off.y + (rectangle.y + rectangle.height) * k > 0 && off.y + rectangle.y * k < props.viewport.height
}
const fragments = computed(() => (props.layout.tableFragments ?? []).filter(fragment => fragment.tableId === tableId.value && visible(documentTableVisibleRect(fragment))))
const selectedRange = computed<DocumentTableCellRange | undefined>(() => selected.value ? { startRow: selected.value.startRow, endRow: selected.value.endRow,
  startColumn: selected.value.startColumn, endColumn: selected.value.endColumn } : undefined)
const target = computed(() => tableId.value && source.value ? { tableId: tableId.value, range: selectedRange.value ?? documentTableWholeRange(documentTableGrid(source.value)) } : undefined)
const anchorFragment = computed(() => fragments.value.find(fragment => fragment.page === props.selection.tableSelection?.pageAffinity)
  ?? fragments.value.find(fragment => fragment.page === props.selection.focus.pageAffinity) ?? fragments.value[0])
const toolbarStyle = computed(() => {
  const fragment = anchorFragment.value
  if (!fragment) return { display: 'none' }
  const rectangle = documentTableVisibleRect(fragment), x = offset.value.x + rectangle.x * scale.value, y = offset.value.y + rectangle.y * scale.value
  return { left: `${Math.max(8, Math.min(x, props.viewport.width - 330))}px`, top: `${Math.max(8, Math.min(y - 96, props.viewport.height - 90))}px`, maxWidth: `${Math.max(160, props.viewport.width - 16)}px` }
})
function at(x: number, y: number, centered = true) {
  return { left: `${offset.value.x + x * scale.value}px`, top: `${offset.value.y + y * scale.value}px`, transform: centered ? 'translate(-50%, -50%)' : undefined }
}
function point(event: Pick<PointerEvent, 'clientX' | 'clientY'>): { x: number; y: number } | undefined {
  const bounds = root.value?.getBoundingClientRect()
  return bounds ? { x: (event.clientX - bounds.left - offset.value.x) / scale.value, y: (event.clientY - bounds.top - offset.value.y) / scale.value } : undefined
}
function state(command: DocumentTableCommand) {
  return props.disabled || props.storyActive ? { available: false, reason: '请先完成当前输入' } : documentTableCommandState(props.model, props.selection, command, context.value)
}
function commit(command: DocumentCommand) {
  if (props.disabled || props.storyActive) return
  if (command.id.startsWith('doc.command.table-') && command.id !== 'doc.command.table-tab') {
    const availability = state(command as DocumentTableCommand)
    if (!availability.available) { if (availability.reason) emit('notify', availability.reason); return }
  }
  popup.value = undefined
  emit('command', command)
  emit('focus')
}
function axisSelection(fragment: LayoutTableFragment, kind: 'table' | 'row' | 'column', index: number | undefined, event: MouseEvent) {
  if (props.disabled || props.storyActive) return
  popup.value = undefined
  emit('selection', documentTableAxisSelection(props.model, fragment.tableId, kind, index, fragment.page, props.selection, event.shiftKey))
  emit('focus')
}
function rowCommand(fragment: LayoutTableFragment, row: number, position: 'before' | 'after'): DocumentTableCommand {
  return { id: 'doc.command.table-insert-row', tableId: fragment.tableId, row, position, range: { startRow: row, endRow: row, startColumn: 0, endColumn: fragment.columnWidths.length - 1 } }
}
function columnCommand(fragment: LayoutTableFragment, column: number, position: 'before' | 'after'): DocumentTableCommand {
  const rows = fragment.rows.length
  return { id: 'doc.command.table-insert-column', tableId: fragment.tableId, column, position, range: { startRow: 0, endRow: rows - 1, startColumn: column, endColumn: column } }
}
function prepareGesture(kind: Gesture['kind'], fragment: LayoutTableFragment, index?: number) {
  const block = findDocumentBlock(props.model, fragment.tableId), table = block && tableForBlock(block)
  if (!table) return undefined
  const grid = documentTableGrid(table), axis = kind === 'row-move' ? 'row' : kind === 'column-move' ? 'column' : 'table'
  const selectedAxis = documentTableAxisSelection(props.model, fragment.tableId, axis, index, fragment.page, props.selection)
  const existing = props.selection.tableSelection
  const keepSelection = existing?.tableId === fragment.tableId && index !== undefined
    && (kind === 'row-move' && existing.kind === 'row' && index >= existing.range.startRow && index <= existing.range.endRow
      || kind === 'column-move' && existing.kind === 'column' && index >= existing.range.startColumn && index <= existing.range.endColumn)
  const range = keepSelection ? { ...existing!.range } : selectedAxis.tableSelection?.range ?? documentTableWholeRange(grid)
  const row = fragment.rows.find(row => row.row === index)
  const size = kind === 'row-size' ? row && (row.contentHeight ?? row.height) : kind === 'column-size' ? fragment.columnWidths[index!] : undefined
  const target = { tableId: fragment.tableId, range }
  const probe: DocumentTableCommand = kind === 'row-size' ? { id: 'doc.command.table-resize-row', ...target, row: index, height: size! }
    : kind === 'column-size' ? { id: 'doc.command.table-resize-column', ...target, column: index, width: size!, widths: [...fragment.columnWidths] }
    : kind === 'row-move' ? { id: 'doc.command.table-move-rows', ...target, fromRow: range.startRow, count: range.endRow - range.startRow + 1, toRow: range.startRow, position: 'before' }
    : kind === 'column-move' ? { id: 'doc.command.table-move-columns', ...target, fromColumn: range.startColumn, count: range.endColumn - range.startColumn + 1, toColumn: range.startColumn, position: 'before' }
    : { id: 'doc.command.table-move', ...target, deltaX: 0, deltaY: 0 }
  return { range, size, probe, selection: { ...selectedAxis, tableSelection: { ...selectedAxis.tableSelection!, range } } }
}
function canGesture(kind: Gesture['kind'], fragment: LayoutTableFragment, index?: number): boolean {
  const prepared = prepareGesture(kind, fragment, index)
  return !!prepared && state(prepared.probe).available
}
const rowControls = computed(() => fragments.value.map(fragment => ({ fragment, key: documentTableFragmentKey(fragment), rect: documentTableVisibleRect(fragment),
  rows: documentTableFragmentRows(fragment).filter(row => visible({ x: fragment.x - 40 / scale.value, y: row.visibleTop, width: 48 / scale.value, height: row.visibleBottom - row.visibleTop })).map(row => ({ ...row,
    insertBefore: rowCommand(fragment, row.row, 'before'), insertAfter: rowCommand(fragment, row.row, 'after') })) })))
const columnControls = computed(() => fragments.value.map(fragment => ({ fragment, key: documentTableFragmentKey(fragment), rect: documentTableVisibleRect(fragment),
  columns: documentTableFragmentColumns(fragment).filter(column => column.start < documentTableVisibleRect(fragment).x + documentTableVisibleRect(fragment).width
    && column.end > documentTableVisibleRect(fragment).x && visible({ x: column.start, y: fragment.y - 32 / scale.value, width: column.width, height: 40 / scale.value })).map(column => ({ ...column,
      insertBefore: columnCommand(fragment, column.column, 'before'), insertAfter: columnCommand(fragment, column.column, 'after') })) })))

interface Gesture {
  kind: 'row-size' | 'column-size' | 'table-move' | 'row-move' | 'column-move'; source: DocumentModel; fragment: LayoutTableFragment;
  pointerId: number; element: HTMLElement; start: { x: number; y: number }; range: DocumentTableCellRange; index?: number; size?: number;
  command?: DocumentTableCommand; moved: boolean; reason?: string
}
const gesture = shallowRef<Gesture>()
function startGesture(kind: Gesture['kind'], fragment: LayoutTableFragment, index: number | undefined, event: PointerEvent) {
  if (event.button !== 0 || props.disabled || props.storyActive || !(event.currentTarget instanceof HTMLElement)) return
  const start = point(event), prepared = prepareGesture(kind, fragment, index)
  if (!start || !prepared) return
  event.preventDefault(); event.stopPropagation()
  const source = props.model, { range, size, selection, probe } = prepared
  const availability = state(probe)
  if (!availability.available) { if (availability.reason) emit('notify', availability.reason); return }
  popup.value = undefined
  if (kind === 'table-move' || kind === 'row-move' || kind === 'column-move') emit('selection', { ...selection, tableSelection: { ...selection.tableSelection!, range } })
  if (source !== props.model) return
  event.currentTarget.setPointerCapture(event.pointerId)
  gesture.value = { kind, source, fragment, index, size, range, pointerId: event.pointerId, element: event.currentTarget, start, moved: false }
  emit('drag-pointer', event)
}
function moveGesture(event: PointerEvent) {
  if (!gesture.value || gesture.value.pointerId !== event.pointerId) return
  event.preventDefault(); event.stopPropagation()
  emit('drag-pointer', event)
  continueGesture(event)
}
function continueGesture(event: DocumentTableDragPointer) {
  const drag = gesture.value, position = point(event)
  if (!drag || drag.pointerId !== event.pointerId || !position) return
  if (drag.source !== props.model || props.disabled || props.storyActive) { cancelGesture(); return }
  const dx = position.x - drag.start.x, dy = position.y - drag.start.y, table = tableForBlock(findDocumentBlock(props.model, drag.fragment.tableId)!)!
  let command: DocumentTableCommand | undefined, preview: DocumentTableGesturePreview = { available: false }
  const target = { tableId: drag.fragment.tableId, range: drag.range }, fragment = drag.fragment
  if (drag.kind === 'row-size') {
    const height = Math.max(1, drag.size! + dy), row = fragment.rows.find(row => row.row === drag.index)!, properties = documentTableRowProperties(table, drag.index!)
    command = { id: 'doc.command.table-resize-row', ...target, row: drag.index, height, heightRule: properties.heightRule === 'auto' ? 'atLeast' : properties.heightRule, allowSplit: properties.allowSplit }
    preview.guide = { x1: fragment.x, y1: row.y + row.height + height - drag.size!, x2: fragment.x + fragment.width, y2: row.y + row.height + height - drag.size! }
  } else if (drag.kind === 'column-size') {
    const width = Math.max(1, drag.size! + dx), column = documentTableFragmentColumns(fragment)[drag.index!]!
    command = { id: 'doc.command.table-resize-column', ...target, column: drag.index, width, widths: [...fragment.columnWidths] }
    preview.guide = { x1: column.start + width, y1: fragment.y, x2: column.start + width, y2: fragment.y + fragment.height }
  } else if (drag.kind === 'table-move') {
    preview.rectangle = { x: fragment.x + dx, y: fragment.y + dy, width: fragment.width, height: fragment.height }
    if (event.ctrlKey || event.metaKey) command = { id: 'doc.command.table-move', ...target, deltaX: dx, deltaY: dy }
    else {
      const caret = hitTestDocument(props.layout, position.x, position.y), id = caret?.cellBlockId ?? caret?.blockId
      if (caret && id && !documentBlockAncestry(props.model, id)?.some(block => block.id === fragment.tableId)) {
        const line = lineForPoint(props.layout, caret)
        command = { id: 'doc.command.table-move', ...target, targetBlockId: id, position: line && position.y < line.y + line.height / 2 ? 'before' : 'after' }
        if (line) preview.guide = { x1: line.x, y1: command.position === 'before' ? line.y : line.y + line.height, x2: line.x + line.width, y2: command.position === 'before' ? line.y : line.y + line.height }
      }
    }
  } else {
    const destination = (props.layout.tableFragments ?? []).filter(candidate => candidate.tableId === fragment.tableId)
      .sort((left, right) => {
        const distance = (candidate: LayoutTableFragment) => { const rect = documentTableVisibleRect(candidate); return Math.max(rect.y - position.y, position.y - rect.y - rect.height, 0) }
        return distance(left) - distance(right)
      }).map(candidate => {
        const rect = documentTableVisibleRect(candidate), x = drag.kind === 'row-move' ? Math.max(rect.x + 0.01, Math.min(position.x, rect.x + rect.width - 0.01)) : position.x
        const y = drag.kind === 'column-move' ? Math.max(rect.y + 0.01, Math.min(position.y, rect.y + rect.height - 0.01)) : position.y
        return { fragment: candidate, position: documentTableLogicalHit(candidate, x, y) }
      }).find(candidate => candidate.position)
    if (destination?.position) {
      if (drag.kind === 'row-move') {
        const row = documentTableFragmentRows(destination.fragment).find(row => row.row === destination.position!.row)!
        const before = position.y < (row.visibleTop + row.visibleBottom) / 2
        command = { id: 'doc.command.table-move-rows', ...target, fromRow: drag.range.startRow, count: drag.range.endRow - drag.range.startRow + 1, toRow: row.row, position: before ? 'before' : 'after' }
        preview.guide = { x1: destination.fragment.x, y1: before ? row.visibleTop : row.visibleBottom, x2: destination.fragment.x + destination.fragment.width, y2: before ? row.visibleTop : row.visibleBottom }
      } else {
        const column = documentTableFragmentColumns(destination.fragment)[destination.position.column]!, before = position.x < (column.start + column.end) / 2
        command = { id: 'doc.command.table-move-columns', ...target, fromColumn: drag.range.startColumn, count: drag.range.endColumn - drag.range.startColumn + 1, toColumn: column.column, position: before ? 'before' : 'after' }
        preview.guide = { x1: before ? column.start : column.end, y1: destination.fragment.y, x2: before ? column.start : column.end, y2: destination.fragment.y + destination.fragment.height }
      }
    }
  }
  const availability = command ? state(command) : { available: false, reason: '请拖到表格外的段落，或按 Ctrl 拖动调整表格偏移' }
  preview.available = availability.available
  gesture.value = { ...drag, command, moved: drag.moved || Math.hypot(dx, dy) * scale.value > 3, reason: availability.reason }
  emit('preview', preview)
}
function endGesture(event: PointerEvent) {
  const drag = gesture.value
  if (!drag || drag.pointerId !== event.pointerId) return
  if (event.type === 'pointerup') moveGesture(event)
  const current = gesture.value
  cancelGesture()
  if (event.type === 'pointerup' && current?.source === props.model && current.moved) {
    if (current.command && state(current.command).available) commit(current.command)
    else if (current.reason) emit('notify', current.reason)
  }
  emit('focus')
}
function cancelGesture() {
  const current = gesture.value
  gesture.value = undefined
  if (current?.element.hasPointerCapture(current.pointerId)) current.element.releasePointerCapture(current.pointerId)
  emit('preview', undefined)
  emit('drag-pointer', undefined)
}
function cancelInteraction() {
  cancelGesture()
  popup.value = undefined
}
function escape(event: KeyboardEvent) {
  if (event.key !== 'Escape' || event.isComposing) return
  if (gesture.value || popup.value) { event.preventDefault(); event.stopPropagation(); cancelInteraction(); emit('focus') }
}
watch(() => [props.model, props.permission, props.disabled, props.storyActive, props.viewport.zoom] as const, cancelInteraction)
watch(() => [tableId.value, JSON.stringify(selectedRange.value)] as const, () => { popup.value = undefined })
onBeforeUnmount(cancelInteraction)
defineExpose({ cancelGesture, cancelInteraction, continueGesture })

const checklistControls = computed(() => props.layout.lines.filter(line => line.checklist && visible(line)).flatMap(line => {
  const rectangle = { x: line.x - 23, y: line.y + (line.height - 16) / 2, width: 16, height: 16 }
  const clip = line.clip && documentTableVisibleRect({ ...rectangle, clip: line.clip })
  if (clip && (clip.width < 16 || clip.height < 16)) return []
  const selection = collapsedSelection({ ...line.target, offset: 0, pageAffinity: line.page })
  const allowed = !props.disabled && !props.storyActive && resolveDocumentPermission(props.model.permissionRules ?? [], context.value, 'edit', permissionTargetsForSelection(props.model, selection)).allowed
    && documentTextFormattingState(props.model, selection, context.value).available
  return [{ key: `${line.checklist!.blockId}:${line.page}:${line.y}`, rectangle, line, selection, disabled: !allowed }]
}))
function toggleChecklist(control: typeof checklistControls.value[number]) {
  if (control.disabled) return
  const source = props.model, block = blockAtPoint(source, control.selection.focus)
  if (!block || block.id !== control.line.checklist?.blockId || block.type !== 'checklist') return
  emit('selection', control.selection)
  if (source === props.model) commit({ id: 'doc.command.toggle-checklist-item', blockId: block.id })
}
const typedControls = computed(() => !props.storyActive && source.value && tableId.value ? props.layout.cells.filter(cell => cell.tableId === tableId.value && cell.cellId && visible(documentTableVisibleRect(cell))).flatMap(cell => {
  const table = source.value!, grid = documentTableGrid(table), owner = grid.cellAt(cell.row!, cell.columnIndex!)!, config = table.columnTypes?.[owner.column]
  const blocks = documentTableCellBlocks(owner.cell)
  if (!config || config.type !== 'checkbox' && config.type !== 'dropdown' || blocks.length !== 1 || blocks.some(block => ['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type))) return []
  if (owner.row < Math.max(table.headerRowCount ?? 0, table.titleRow ? 1 : 0)) return []
  const margin = documentTableCellMargin(table, owner.cell), text = blocks.map(block => runText(block.runs)).join('\n')
  // A small editor action sits beside Canvas text; the domain value still renders on Canvas.
  const actionWidth = (config.type === 'checkbox' ? 18 : 26) / scale.value
  const rectangle = { x: cell.x + cell.width - margin.right - actionWidth, y: cell.y + margin.top, width: actionWidth, height: Math.min(26 / scale.value, Math.max(0, cell.height - margin.top - margin.bottom)) }
  const clipped = documentTableVisibleRect({ ...rectangle, ...(cell.clip ? { clip: cell.clip } : {}) })
  if (clipped.width < Math.min(24, rectangle.width) || clipped.height < Math.min(18 / scale.value, rectangle.height) || rectangle.height <= 0) return []
  const range = { startRow: owner.row, endRow: owner.row + owner.rowSpan - 1, startColumn: owner.column, endColumn: owner.column + owner.colSpan - 1 }
  const options = config.dropdownOptions ?? [], option = options.find(option => option.id === owner.cell.style.selectedDropdownOptionId && option.label === text)
    ?? (options.filter(option => option.label === text).length === 1 ? options.find(option => option.label === text) : undefined)
  const choiceOptions = options.map(option => ({ value: option.id, label: option.label, style: option.color ? { color: option.color } : undefined,
    disabled: !state({ id: 'doc.command.table-set-cell-text', tableId: tableId.value!, range, row: owner.row, column: owner.column, text: option.label, optionId: option.id }).available }))
  return [{ key: `${cell.cellId}:${cell.page}:${cell.y}`, cell, rectangle, text, config, range, option, choiceOptions,
    checked: /^(?:true|1|yes|checked|☑|✓)$/i.test(text.trim()), command: { id: 'doc.command.table-set-cell-text', tableId: tableId.value!, range, row: owner.row, column: owner.column, text } as DocumentTableCommand }]
}) : [])
function typedValue(control: typeof typedControls.value[number], value: unknown) {
  if (props.disabled) return
  const option = control.config.type === 'dropdown' && typeof value === 'string' ? control.config.dropdownOptions?.find(option => option.id === value) : undefined
  if (control.config.type === 'dropdown' && !option) return
  if (control.config.type === 'checkbox' && typeof value !== 'boolean') return
  commit({ id: 'doc.command.table-set-cell-text', tableId: control.cell.tableId, range: control.range, row: control.cell.row!, column: control.cell.columnIndex!,
    text: option ? option.label : String(value), ...(option ? { optionId: option.id } : {}) })
}
const textTargets = computed(() => props.selection.tableSelection ? documentTableSelectedTextTargets(props.model, props.selection.tableSelection)
  : selected.value ? [props.selection.focus] : [])
const firstTextPoint = computed(() => textTargets.value[0] ? { ...textTargets.value[0]!, offset: props.selection.tableSelection ? 1 : props.selection.focus.offset } : undefined)
const marks = computed(() => firstTextPoint.value ? marksAtPoint(props.model, firstTextPoint.value) : {})
const textState = computed(() => !props.disabled && selected.value ? documentTextFormattingState(props.model, props.selection, context.value) : { available: false })
const resetState = computed(() => !props.disabled && selected.value ? documentClearFormattingState(props.model, props.selection, context.value) : { available: false })
const border = computed(() => source.value && selectedRange.value ? documentTableCellBorder(documentTableGrid(source.value), documentTableGrid(source.value).cellAt(selectedRange.value.startRow, selectedRange.value.startColumn)!, 'top') : undefined)
const background = computed(() => source.value && selectedRange.value ? documentTableCellFill(source.value, documentTableGrid(source.value).cellAt(selectedRange.value.startRow, selectedRange.value.startColumn)!.cell) : undefined)
const vertical = computed(() => source.value && selectedRange.value ? documentTableGrid(source.value).cellAt(selectedRange.value.startRow, selectedRange.value.startColumn)?.cell.style.verticalAlign ?? 'top' : 'top')
const textButtons = [{ field: 'bold', label: '粗体', icon: Bold }, { field: 'italic', label: '斜体', icon: Italic }, { field: 'underline', label: '下划线', icon: Underline }, { field: 'strike', label: '删除线', icon: Strikethrough }] as const
const alignButtons = [{ value: 'left', label: '左对齐', icon: AlignLeft }, { value: 'center', label: '居中', icon: AlignCenter }, { value: 'right', label: '右对齐', icon: AlignRight }] as const
const namedStyles = [{ value: 'normal', label: '正文' }, { value: 'title', label: '标题' }, { value: 'subtitle', label: '副标题' },
  { value: 'heading1', label: '标题 1' }, { value: 'heading2', label: '标题 2' }, { value: 'heading3', label: '标题 3' }, { value: 'heading4', label: '标题 4' }, { value: 'heading5', label: '标题 5' }]
function textColor(params: DocumentColorParams) { if (textState.value.available) commit({ id: 'doc.command.format-text', ...params }) }
function tableColor(id: 'doc.command.table-set-background' | 'doc.command.table-set-border-color', params: { tableId: string; range: DocumentTableCellRange; color: string | null }) { commit({ id, ...params }) }
function setTableValue(kind: 'vertical' | 'preset' | 'width' | 'style', value: unknown) {
  if (!target.value) return
  const targetValue = target.value
  if (kind === 'vertical' && (value === 'top' || value === 'center' || value === 'bottom')) commit({ id: 'doc.command.table-set-vertical-align', ...targetValue, value })
  if (kind === 'preset' && typeof value === 'string' && Object.hasOwn(DOCUMENT_TABLE_BORDER_LABELS, value)) commit({ id: 'doc.command.table-set-border', ...targetValue, preset: value as keyof typeof DOCUMENT_TABLE_BORDER_LABELS })
  if (kind === 'width' && typeof value === 'number') commit({ id: 'doc.command.table-set-border-width', ...targetValue, width: value * 4 / 3 })
  if (kind === 'style' && (value === 'solid' || value === 'dot' || value === 'dash')) commit({ id: 'doc.command.table-set-border-style', ...targetValue, style: value })
}
function namedStyle(value: unknown) {
  if (typeof value === 'string' && namedStyles.some(style => style.value === value)) commit({ id: 'doc.command.set-paragraph-named-style', namedStyle: value as 'normal' | 'title' | 'subtitle' | 'heading1' | 'heading2' | 'heading3' | 'heading4' | 'heading5' })
}
</script>

<template>
  <div ref="root" class="document-table-overlay" :style="overlayStyle" @keydown="escape" @pointermove="moveGesture" @pointerup="endGesture" @pointercancel="endGesture" @lostpointercapture="cancelGesture">
    <template v-for="control in rowControls" :key="control.key">
      <Tooltip title="选择整个表格"><Button class="table-overlay-handle table-select" size="small" :style="at(control.rect.x - 14 / scale, control.rect.y - 14 / scale)" :disabled="disabled" aria-label="选择整个表格" @pointerdown.stop @click="axisSelection(control.fragment, 'table', undefined, $event)"><Table2 :size="13" /></Button></Tooltip>
      <Tooltip title="拖动移动表格；Ctrl 拖动调整偏移"><Button class="table-overlay-handle table-move" size="small" :style="at(control.rect.x + 14 / scale, control.rect.y - 14 / scale)" :disabled="!canGesture('table-move', control.fragment)" aria-label="移动表格" @pointerdown="startGesture('table-move', control.fragment, undefined, $event)"><Move :size="13" /></Button></Tooltip>
      <template v-for="row in control.rows" :key="row.id">
        <Tooltip :title="`选择第 ${row.row + 1} 行`"><Button class="table-overlay-handle table-row-select" size="small" :style="at(control.rect.x - 14 / scale, (row.visibleTop + row.visibleBottom) / 2)" :disabled="disabled" :aria-label="`选择第 ${row.row + 1} 行`" @pointerdown.stop @click="axisSelection(control.fragment, 'row', row.row, $event)">{{ row.row + 1 }}</Button></Tooltip>
        <Tooltip :title="`移动第 ${row.row + 1} 行`"><Button class="table-overlay-handle table-axis-move" size="small" :style="at(control.rect.x - 35 / scale, (row.visibleTop + row.visibleBottom) / 2)" :disabled="!canGesture('row-move', control.fragment, row.row)" :aria-label="`移动第 ${row.row + 1} 行`" @pointerdown="startGesture('row-move', control.fragment, row.row, $event)"><GripVertical :size="12" /></Button></Tooltip>
        <Tooltip v-if="row.y >= control.rect.y" :title="`在第 ${row.row + 1} 行上方插入`"><Button class="table-overlay-handle table-insert-dot" size="small" :style="at(control.rect.x - 14 / scale, row.y)" :disabled="!state(row.insertBefore).available" :aria-label="`在第 ${row.row + 1} 行上方插入`" @pointerdown.stop @click="commit(row.insertBefore)"><Plus :size="12" /></Button></Tooltip>
        <template v-if="row.ownsBottom">
          <Tooltip :title="`调整第 ${row.row + 1} 行高度`"><Button class="table-overlay-resize table-row-resize" type="text" :style="{ ...at(control.rect.x, row.y + row.height, false), width: `${control.rect.width * scale}px` }" :disabled="!canGesture('row-size', control.fragment, row.row)" :aria-label="`调整第 ${row.row + 1} 行高度`" @pointerdown="startGesture('row-size', control.fragment, row.row, $event)" /></Tooltip>
          <Tooltip v-if="row.row === control.fragment.rows.length - 1" title="在表格末尾插入行"><Button class="table-overlay-handle table-insert-dot" size="small" :style="at(control.rect.x - 14 / scale, row.y + row.height)" :disabled="!state(row.insertAfter).available" aria-label="在表格末尾插入行" @pointerdown.stop @click="commit(row.insertAfter)"><Plus :size="12" /></Button></Tooltip>
        </template>
      </template>
    </template>
    <template v-for="control in columnControls" :key="control.key">
      <template v-for="column in control.columns" :key="column.column">
        <Tooltip :title="`选择第 ${column.column + 1} 列`"><Button class="table-overlay-handle table-column-select" size="small" :style="at((column.start + column.end) / 2, control.rect.y - 14 / scale)" :disabled="disabled" :aria-label="`选择第 ${column.column + 1} 列`" @pointerdown.stop @click="axisSelection(control.fragment, 'column', column.column, $event)">{{ column.column + 1 }}</Button></Tooltip>
        <Tooltip :title="`移动第 ${column.column + 1} 列`"><Button class="table-overlay-handle table-axis-move" size="small" :style="at((column.start + column.end) / 2, control.rect.y - 35 / scale)" :disabled="!canGesture('column-move', control.fragment, column.column)" :aria-label="`移动第 ${column.column + 1} 列`" @pointerdown="startGesture('column-move', control.fragment, column.column, $event)"><GripHorizontal :size="12" /></Button></Tooltip>
        <Tooltip v-if="column.start >= control.rect.x" :title="`在第 ${column.column + 1} 列左侧插入`"><Button class="table-overlay-handle table-insert-dot" size="small" :style="at(column.start, control.rect.y - 14 / scale)" :disabled="!state(column.insertBefore).available" :aria-label="`在第 ${column.column + 1} 列左侧插入`" @pointerdown.stop @click="commit(column.insertBefore)"><Plus :size="12" /></Button></Tooltip>
        <Tooltip v-if="column.end <= control.rect.x + control.rect.width" :title="`调整第 ${column.column + 1} 列宽度`"><Button class="table-overlay-resize table-column-resize" type="text" :style="{ ...at(column.end, control.rect.y, false), height: `${control.rect.height * scale}px` }" :disabled="!canGesture('column-size', control.fragment, column.column)" :aria-label="`调整第 ${column.column + 1} 列宽度`" @pointerdown="startGesture('column-size', control.fragment, column.column, $event)" /></Tooltip>
        <Tooltip v-if="column.column === control.fragment.columnWidths.length - 1" title="在表格末尾插入列"><Button class="table-overlay-handle table-insert-dot" size="small" :style="at(column.end, control.rect.y - 14 / scale)" :disabled="!state(column.insertAfter).available" aria-label="在表格末尾插入列" @pointerdown.stop @click="commit(column.insertAfter)"><Plus :size="12" /></Button></Tooltip>
      </template>
    </template>
    <div v-for="control in checklistControls" :key="control.key" class="document-checklist-control" :style="{ ...at(control.rectangle.x, control.rectangle.y, false), transform: `scale(${scale})` }" @pointerdown.stop @dblclick.stop>
      <Checkbox :checked="control.line.checklist!.checked" :disabled="control.disabled" aria-label="切换清单状态" @change="toggleChecklist(control)" />
    </div>
    <div v-for="control in typedControls" :key="control.key" class="document-typed-cell-control" :style="{ ...at(control.rectangle.x, control.rectangle.y, false), width: `${control.rectangle.width * scale}px`, maxHeight: `${control.rectangle.height * scale}px` }" @pointerdown.stop @dblclick.stop>
      <Checkbox v-if="control.config.type === 'checkbox'" :checked="control.checked" :disabled="!state({ ...control.command, id: 'doc.command.table-set-cell-text', row: control.cell.row!, column: control.cell.columnIndex!, text: control.checked ? 'false' : 'true' }).available" :aria-label="`第 ${control.cell.row! + 1} 行复选框`" @change="typedValue(control, !control.checked)" />
      <Select v-else class="table-cell-choice-action" variant="borderless" size="small" :value="control.option?.id" :options="control.choiceOptions" :open="popup === control.key" :disabled="disabled || !control.choiceOptions.some(option => !option.disabled)" :aria-label="`第 ${control.cell.row! + 1} 行下拉选项，当前 ${control.text || '空值'}`" @open-change="value => popup = value ? control.key : undefined" @change="value => typedValue(control, value)" />
    </div>
    <div v-if="floatingToolbar && selected && target && anchorFragment" class="document-table-floating-toolbar" :style="toolbarStyle" role="toolbar" aria-label="表格浮动工具栏" @pointerdown.stop @dblclick.stop @contextmenu.prevent.stop>
      <div class="table-floating-row">
        <Tooltip v-for="button in textButtons" :key="button.field" :title="button.label"><Button type="text" size="small" :disabled="!textState.available" :aria-label="button.label" :aria-pressed="marks[button.field] === true" @pointerdown.prevent @click="commit({ id: 'doc.command.format-text', patch: { [button.field]: marks[button.field] !== true } })"><component :is="button.icon" :size="14" /></Button></Tooltip>
        <Select size="small" class="table-floating-style" placeholder="文字样式" :options="namedStyles" :disabled="!textState.available" aria-label="文字样式" @change="namedStyle" />
        <DocumentColorControl title="文字颜色" field="color" :icon="Palette" :default-color="token.colorText" :selected-color="marks.color" :disabled="!textState.available" :open="popup === 'text-color'" @update:open="value => popup = value ? 'text-color' : undefined" @insert="textColor" />
        <Tooltip title="清除格式"><Button type="text" size="small" :disabled="!resetState.available" aria-label="清除格式" @pointerdown.prevent @click="commit({ id: 'doc.command.clear-formatting' })"><RotateCcw :size="14" /></Button></Tooltip>
      </div>
      <Divider class="table-floating-divider" />
      <div class="table-floating-row">
        <Tooltip v-for="button in alignButtons" :key="button.value" :title="button.label"><Button type="text" size="small" :disabled="!state({ id: 'doc.command.table-set-text-align', ...target, value: button.value }).available" :aria-label="button.label" @pointerdown.prevent @click="commit({ id: 'doc.command.table-set-text-align', ...target, value: button.value })"><component :is="button.icon" :size="14" /></Button></Tooltip>
        <Select size="small" :value="vertical" class="table-floating-vertical" :disabled="!state({ id: 'doc.command.table-set-vertical-align', ...target, value: vertical }).available" :options="[{ value: 'top', label: '顶端对齐' }, { value: 'center', label: '垂直居中' }, { value: 'bottom', label: '底端对齐' }]" aria-label="单元格垂直对齐" @change="value => setTableValue('vertical', value)" />
        <DocumentTableColorControl title="背景色" field="background" :icon="PaintBucket" default-color="#ffffff" :selected-color="background" :table-id="target.tableId" :range="target.range" :disabled="!state({ id: 'doc.command.table-set-background', ...target, color: background ?? '#ffffff' }).available" :open="popup === 'background'" @update:open="value => popup = value ? 'background' : undefined" @insert="params => tableColor('doc.command.table-set-background', params)" />
        <DocumentTableColorControl title="边框颜色" field="border" :icon="Palette" default-color="#aab3bf" :selected-color="border?.color" :table-id="target.tableId" :range="target.range" :disabled="!state({ id: 'doc.command.table-set-border-color', ...target, color: border?.color ?? '#aab3bf' }).available" :open="popup === 'border-color'" @update:open="value => popup = value ? 'border-color' : undefined" @insert="params => tableColor('doc.command.table-set-border-color', params)" />
      </div>
      <div class="table-floating-row">
        <Select size="small" class="table-floating-border" placeholder="边框" :options="Object.entries(DOCUMENT_TABLE_BORDER_LABELS).map(([value, label]) => ({ value, label, disabled: !state({ id: 'doc.command.table-set-border', ...target, preset: value as keyof typeof DOCUMENT_TABLE_BORDER_LABELS }).available }))" :disabled="disabled" aria-label="表格边框" @change="value => setTableValue('preset', value)" />
        <Select size="small" class="table-floating-width" :value="(border?.width ?? 1) * 0.75" :options="[0, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map(value => ({ value, label: `${value} pt`, disabled: !state({ id: 'doc.command.table-set-border-width', ...target, width: value * 4 / 3 }).available }))" :disabled="disabled" aria-label="边框宽度" @change="value => setTableValue('width', value)" />
        <Select size="small" class="table-floating-border-style" :value="border?.style ?? 'solid'" :options="[{ value: 'solid', label: '实线' }, { value: 'dot', label: '点线' }, { value: 'dash', label: '虚线' }]" :disabled="!state({ id: 'doc.command.table-set-border-style', ...target, style: border?.style ?? 'solid' }).available" aria-label="边框样式" @change="value => setTableValue('style', value)" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.document-table-overlay { position: absolute; inset: 0; z-index: 4; pointer-events: none; color: var(--table-overlay-text); }
.table-overlay-handle { position: absolute; z-index: 2; width: 20px; min-width: 20px; height: 20px; padding: 0; display: grid; place-items: center; pointer-events: auto; font-size: 10px; background: var(--table-overlay-background); border-color: var(--table-overlay-border); }
.table-overlay-handle:hover { color: var(--table-overlay-primary); border-color: var(--table-overlay-primary); }
.table-move, .table-axis-move { cursor: grab; touch-action: none; }
.table-insert-dot { width: 16px; min-width: 16px; height: 16px; z-index: 4; border-radius: 50%; }
.table-overlay-resize { position: absolute; z-index: 1; min-width: 0; padding: 0; pointer-events: auto; border: 0; border-radius: 0; background: transparent; touch-action: none; }
.table-row-resize { height: 6px; min-height: 6px; margin-top: -3px; cursor: row-resize; }
.table-column-resize { width: 6px; height: auto; margin-left: -3px; cursor: col-resize; }
.table-overlay-resize:hover, .table-overlay-resize:focus-visible { background: color-mix(in srgb, var(--table-overlay-primary) 24%, transparent); }
.document-checklist-control { position: absolute; pointer-events: auto; transform-origin: top left; line-height: 16px; }
.document-typed-cell-control { position: absolute; overflow: hidden; pointer-events: auto; background: var(--table-overlay-background); }
.document-typed-cell-control :deep(.ant-select) { width: 100%; }
.table-cell-choice-action :deep(.ant-select-selection-item), .table-cell-choice-action :deep(.ant-select-selection-placeholder) { opacity: 0; }
.document-table-floating-toolbar { position: absolute; z-index: 6; pointer-events: auto; padding: 5px; border: 1px solid var(--table-overlay-border); border-radius: 7px; background: var(--table-overlay-background); box-shadow: var(--table-overlay-shadow); }
.table-floating-row { display: flex; align-items: center; gap: 3px; flex-wrap: wrap; }
.table-floating-row + .table-floating-row { margin-top: 4px; }
.table-floating-divider { margin: 4px 0; }
.table-floating-style, .table-floating-vertical { width: 106px; }
.table-floating-border { width: 120px; }
.table-floating-width { width: 76px; }
.table-floating-border-style { width: 70px; }
</style>
