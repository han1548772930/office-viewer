<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { Alert, Button, Input, Tooltip, theme } from 'antdv-next'
import type { InputTextAreaRef } from 'antdv-next'
import { Check, ChevronDown, ChevronUp, X } from 'lucide-vue-next'
import { sheetCellInputActionForKey } from '../../features/sheet-workbench/sheet-cell-edit-input.ts'
import { replaceSheetCellInputSelection } from '../../features/sheet-workbench/sheet-cell-edit-input.ts'
import { createSheetCellNativeInputBridge, type SheetCellInputBridgeContext } from '../../features/sheet-workbench/sheet-cell-native-input.ts'
import type { SheetCellNativeInputEvent } from '../../features/sheet-workbench/sheet-cell-input-trace.ts'
import { sheetFontFormatCommandForShortcut } from '../../features/sheet-workbench/sheet-font-format-command.ts'

const props = defineProps<{ address: string; value: string; editing: boolean; editSessionId?: string; inputContext?: () => SheetCellInputBridgeContext; composing?: boolean; error?: string }>()
const emit = defineEmits<{
  (event: 'update:value', value: string): void
  (event: 'begin', selection?: { start: number; end: number }): void
  (event: 'input-event', value: SheetCellNativeInputEvent): void
  (event: 'format-command', commandId: string): void
  (event: 'format-keydown', value: KeyboardEvent): void
  (event: 'confirm', navigation?: { key: 'Enter' | 'Tab'; shift: boolean }): void
  (event: 'cancel'): void
  (event: 'navigate', address: string): void
  (event: 'focus-grid'): void
  (event: 'composition-change', composing: boolean): void
}>()

const addressDraft = ref(props.address)
const addressComposing = ref(false), formulaComposing = ref(false)
const expanded = ref(false)
const root = ref<HTMLElement | null>(null)
const editor = ref<InputTextAreaRef | null>(null)
const nativeEditorKey = ref(0)
// The synchronous callback also owns "no session" while Vue is still flushing the previous props.
const inputSessionId = () => props.inputContext ? props.inputContext().sessionId : props.editSessionId
const nativeInput = createSheetCellNativeInputBridge({ host: 'formula', sessionId: inputSessionId,
  value: () => props.inputContext?.().value ?? props.value, contains: element => Boolean(root.value?.contains(element) && element.closest('.formula-input')),
  emit: event => emit('input-event', event) })
const { token } = theme.useToken()
const colors = computed(() => ({
  '--formula-surface': token.value.colorBgContainer,
  '--formula-divider': token.value.colorBorderSecondary,
  '--formula-accent': token.value.colorPrimary,
  '--formula-muted': token.value.colorTextSecondary,
}))
watch(() => props.address, address => { addressDraft.value = address })
watch(() => props.editing, editing => { if (!editing) changeFormulaComposition(false) })
watch(() => props.editSessionId, (sessionId, previous) => {
  // Preserve a pointer/IME-chosen caret while idle focus synchronously opens the first session.
  if (previous || !sessionId) { nativeEditorKey.value++; changeFormulaComposition(false) }
}, { flush: 'sync' })

function changeFormulaComposition(composing: boolean): void {
  formulaComposing.value = composing
  emit('composition-change', composing)
}
function onFormulaComposition(event: CompositionEvent, composing: boolean): void {
  const target = event.target
  if (!(target instanceof HTMLTextAreaElement) || !target.closest('.formula-input')) return
  nativeInput.capture(event, composing ? 'compositionstart' : 'compositionend')
  changeFormulaComposition(composing)
}
function onFormulaFocus(event: FocusEvent): void {
  const target = event.target
  if (!(target instanceof HTMLTextAreaElement) || !root.value?.contains(target)) return
  emit('begin', { start: target.selectionStart ?? target.value.length, end: target.selectionEnd ?? target.value.length })
  nativeInput.capture(event, 'selection')
}
function onFormulaValue(value: string): void {
  if (!inputSessionId()) emit('update:value', value)
}

function onKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented || !(event.target instanceof HTMLElement) || !root.value?.contains(event.target)) return
  emit('format-keydown', event)
  if (event.defaultPrevented) return
  const fontCommand = sheetFontFormatCommandForShortcut(event)
  if (fontCommand) { event.preventDefault(); emit('format-command', fontCommand); return }
  const action = sheetCellInputActionForKey(event, formulaComposing.value || props.composing)
  if (action === 'cancel') {
    event.preventDefault()
    emit('cancel')
  } else if (action === 'commit') {
    event.preventDefault(); emit('confirm')
  } else if (action === 'line-break' && event.target instanceof HTMLTextAreaElement) {
    event.preventDefault()
    const element = event.target, selection = { start: element.selectionStart, end: element.selectionEnd }
    const replacement = inputSessionId() ? nativeInput.replace(element, '\n', selection, 'insertLineBreak', element.value)
      : replaceSheetCellInputSelection(element.value, '\n', selection)
    if (!replacement) return
    if (!inputSessionId()) onFormulaValue(replacement.value)
    const sessionId = inputSessionId()
    void nextTick(() => { if (inputSessionId() === sessionId) focus(replacement.selection) })
  } else if (action !== 'native' && action !== 'line-break') {
    event.preventDefault()
    emit('confirm', { key: action === 'next-column' || action === 'previous-column' ? 'Tab' : 'Enter', shift: action === 'previous-column' || action === 'previous-row' })
  }
}

function onAddressKeydown(event: KeyboardEvent): void {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || addressComposing.value) return
  if (event.key === 'Enter') {
    event.preventDefault()
    emit('navigate', addressDraft.value)
  } else if (event.key === 'Escape') {
    event.preventDefault()
    addressDraft.value = props.address
    emit('focus-grid')
  }
}

function toggleExpanded() {
  const sessionId = inputSessionId()
  expanded.value = !expanded.value
  void nextTick(() => { if (props.editing && inputSessionId() === sessionId) editor.value?.focus() })
}

function startFormula() {
  const address = props.address
  emit('begin')
  const sessionId = inputSessionId()
  void nextTick(() => {
    if (address !== props.address || !props.editing || inputSessionId() !== sessionId) return
    const element = formulaElement()
    if (!props.value && element) {
      if (inputSessionId()) nativeInput.replace(element, '=', { start: 0, end: 0 }, 'insertText', '')
      else emit('update:value', '=')
    }
    editor.value?.focus()
  })
}

function formulaElement(): HTMLTextAreaElement | undefined {
  const native = editor.value?.nativeElement
  const element = native instanceof HTMLTextAreaElement ? native : native?.querySelector('textarea')
  return element instanceof HTMLTextAreaElement ? element : undefined
}
function focus(selection?: { start: number; end: number }) {
  editor.value?.focus({ preventScroll: true })
  const element = formulaElement()
  if (element && selection && !formulaComposing.value && !props.composing) element.setSelectionRange(selection.start, selection.end)
  element?.dispatchEvent(new Event('select', { bubbles: true }))
}
defineExpose({ focus })
</script>

<template>
  <div ref="root" class="formula-toolbar" :class="{ expanded }" :style="colors" data-testid="sheet-formula-bar" @keydown.stop @keydown.capture="nativeInput.capture($event, 'selection')" @keyup.capture="nativeInput.capture($event, 'selection')" @beforeinput.capture="nativeInput.capture($event, 'beforeinput')" @input.capture="nativeInput.capture($event, 'input')" @compositionstart.capture="onFormulaComposition($event, true)" @compositionend.capture="onFormulaComposition($event, false)" @focusin.capture="nativeInput.capture($event, 'selection')" @select.capture="nativeInput.capture($event, 'selection')">
    <Input
      v-model:value="addressDraft"
      class="name-box"
      aria-label="名称框"
      :aria-invalid="Boolean(error)"
      :title="error || '输入单元格或范围，例如 A1 或 A1:C8'"
      :status="error ? 'error' : undefined"
      variant="borderless"
      @keydown="onAddressKeydown"
      @compositionstart="addressComposing = true"
      @compositionend="addressComposing = false"
      @focus="($event.target as HTMLInputElement).select()"
    />
    <div class="formula-actions">
      <Tooltip title="取消编辑（Esc）"><Button type="text" size="small" aria-label="取消编辑" :disabled="!editing" @mousedown.prevent @click="emit('cancel')"><X :size="15" /></Button></Tooltip>
      <Tooltip title="确认编辑（Enter）"><Button type="text" size="small" aria-label="确认编辑" :disabled="!editing || formulaComposing || composing" @mousedown.prevent @click="emit('confirm')"><Check :size="15" /></Button></Tooltip>
      <Tooltip title="编辑公式"><Button type="text" size="small" aria-label="编辑公式" @mousedown.prevent @click="startFormula"><span class="formula-symbol">fx</span></Button></Tooltip>
    </div>
    <Input.TextArea
      :key="nativeEditorKey"
      ref="editor"
      class="formula-input"
      aria-label="公式编辑"
      :rows="expanded ? 3 : 1"
      :change-on-composing="true"
      variant="borderless"
      spellcheck="false"
      @focus="onFormulaFocus"
      @update:value="onFormulaValue"
      @keydown="onKeydown"
      :value="value"
    />
    <Tooltip :title="expanded ? '折叠公式栏' : '展开公式栏'">
      <Button class="formula-expand" type="text" size="small" :aria-label="expanded ? '折叠公式栏' : '展开公式栏'" :aria-expanded="expanded" @mousedown.prevent @click="toggleExpanded"><component :is="expanded ? ChevronUp : ChevronDown" :size="14" /></Button>
    </Tooltip>
    <Alert v-if="error" class="formula-error" type="error" :title="error" show-icon />
  </div>
</template>

<style scoped>
.formula-toolbar { position: relative; display: flex; box-sizing: border-box; flex: 0 0 28px; min-width: 0; height: 28px; align-items: flex-start; border-bottom: 1px solid var(--formula-divider); background: var(--formula-surface); }
.formula-toolbar.expanded { flex-basis: 80px; height: 80px; }
.formula-toolbar .name-box { flex: 0 0 100px; box-sizing: border-box; width: 100px; height: 27px; border-radius: 0; border-inline-end: 1px solid var(--formula-divider); padding: 2px 9px; font-size: 12px; }
.formula-actions { display: flex; flex: 0 0 88px; height: 27px; align-items: center; justify-content: center; border-inline-end: 1px solid var(--formula-divider); }
.formula-actions :deep(.ant-btn) { display: inline-flex; width: 27px; height: 24px; align-items: center; justify-content: center; padding: 0; }
.formula-symbol { color: var(--formula-muted); font-family: Georgia, serif; font-size: 16px; font-style: italic; }
.formula-input { display: block; box-sizing: border-box; flex: 1; min-width: 0; height: 100%; }
.formula-toolbar textarea.formula-input, .formula-input :deep(textarea) { box-sizing: border-box; height: 100%; resize: none; border: 0; border-radius: 0; padding: 4px 10px; background: transparent; font-family: Consolas, "Microsoft YaHei", monospace; font-size: 12px; line-height: 19px; white-space: pre-wrap; overflow-y: auto; }
.formula-toolbar textarea.formula-input:focus, .formula-input :deep(textarea:focus) { box-shadow: inset 0 0 0 1px var(--formula-accent); }
.formula-expand { display: flex; flex: 0 0 26px; width: 26px; height: 100%; align-items: center; justify-content: center; padding: 0; }
.formula-error { position: absolute; z-index: 8; top: 100%; inset-inline-start: 0; max-width: min(360px, 90vw); }
</style>
