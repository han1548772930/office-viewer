<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { Alert, Button, Input, Tooltip, theme } from 'antdv-next'
import type { InputTextAreaRef } from 'antdv-next'
import { Check, ChevronDown, ChevronUp, X } from 'lucide-vue-next'
import { sheetCellInputActionForKey } from '../../features/sheet-workbench/sheet-cell-edit-input.ts'

const props = defineProps<{ address: string; value: string; editing: boolean; composing?: boolean; error?: string }>()
const emit = defineEmits<{
  (event: 'update:value', value: string): void
  (event: 'begin'): void
  (event: 'confirm', navigation?: { key: 'Enter' | 'Tab'; shift: boolean }): void
  (event: 'cancel'): void
  (event: 'navigate', address: string): void
  (event: 'focus-grid'): void
  (event: 'composition-change', composing: boolean): void
}>()

const addressDraft = ref(props.address)
const addressComposing = ref(false), formulaComposing = ref(false)
const expanded = ref(false)
const editor = ref<InputTextAreaRef | null>(null)
const { token } = theme.useToken()
const colors = computed(() => ({
  '--formula-surface': token.value.colorBgContainer,
  '--formula-divider': token.value.colorBorderSecondary,
  '--formula-accent': token.value.colorPrimary,
  '--formula-muted': token.value.colorTextSecondary,
}))
watch(() => props.address, address => { addressDraft.value = address })
watch(() => props.editing, editing => { if (!editing) changeFormulaComposition(false) })

function changeFormulaComposition(composing: boolean): void {
  formulaComposing.value = composing
  emit('composition-change', composing)
}

function onKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented) return
  const action = sheetCellInputActionForKey(event, formulaComposing.value || props.composing)
  if (action === 'cancel') {
    event.preventDefault()
    emit('cancel')
  } else if (action === 'commit') {
    event.preventDefault(); emit('confirm')
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
  expanded.value = !expanded.value
  void nextTick(() => { if (props.editing) editor.value?.focus() })
}

function startFormula() {
  emit('begin')
  if (!props.value) emit('update:value', '=')
  void nextTick(() => editor.value?.focus())
}

function focus(selection?: { start: number; end: number }) {
  editor.value?.focus({ preventScroll: true })
  const native = editor.value?.nativeElement
  const element = native instanceof HTMLTextAreaElement ? native : native?.querySelector('textarea')
  if (element instanceof HTMLTextAreaElement && selection && !formulaComposing.value && !props.composing) element.setSelectionRange(selection.start, selection.end)
}
defineExpose({ focus })
</script>

<template>
  <div class="formula-toolbar" :class="{ expanded }" :style="colors" data-testid="sheet-formula-bar" @keydown.stop>
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
      ref="editor"
      class="formula-input"
      aria-label="公式编辑"
      :rows="expanded ? 3 : 1"
      :change-on-composing="true"
      variant="borderless"
      spellcheck="false"
      @focus="emit('begin')"
      @update:value="emit('update:value', $event)"
      @keydown="onKeydown"
      @compositionstart="changeFormulaComposition(true)"
      @compositionend="changeFormulaComposition(false)"
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
.formula-expand { display: flex; flex: 0 0 26px; width: 26px; height: 27px; align-items: center; justify-content: center; padding: 0; }
.formula-error { position: absolute; z-index: 8; top: 100%; inset-inline-start: 0; max-width: min(360px, 90vw); }
</style>
