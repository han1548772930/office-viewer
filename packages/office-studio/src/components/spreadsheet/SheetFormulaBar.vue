<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { Button, Input, Tooltip, theme } from 'antdv-next'
import type { InputTextAreaRef } from 'antdv-next'
import { Check, ChevronDown, ChevronUp, X } from 'lucide-vue-next'

const props = defineProps<{ address: string; value: string; editing: boolean; error?: string }>()
const emit = defineEmits<{
  (event: 'update:value', value: string): void
  (event: 'begin'): void
  (event: 'confirm', navigation?: { key: 'Enter' | 'Tab'; shift: boolean }): void
  (event: 'cancel'): void
  (event: 'navigate', address: string): void
  (event: 'focus-grid'): void
}>()

const addressDraft = ref(props.address)
const expanded = ref(false)
const editor = ref<InputTextAreaRef | null>(null)
const { token } = theme.useToken()
const colors = computed(() => ({
  '--formula-surface': token.value.colorBgContainer,
  '--formula-divider': token.value.colorBorderSecondary,
  '--formula-accent': token.value.colorPrimary,
  '--formula-muted': token.value.colorTextSecondary,
  '--formula-error-bg': token.value.colorErrorBg,
  '--formula-error-border': token.value.colorErrorBorder,
  '--formula-error-text': token.value.colorErrorText,
}))
watch(() => props.address, address => { addressDraft.value = address })

function onKeydown(event: KeyboardEvent) {
  if (event.isComposing) return
  if (event.key === 'Escape') {
    event.preventDefault()
    emit('cancel')
  } else if ((event.key === 'Enter' && !event.altKey) || event.key === 'Tab') {
    event.preventDefault()
    emit('confirm', { key: event.key, shift: event.shiftKey })
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

function focus() { editor.value?.focus() }
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
      @keydown.enter.prevent.stop="emit('navigate', addressDraft)"
      @keydown.esc.prevent.stop="addressDraft = address; emit('focus-grid')"
      @focus="($event.target as HTMLInputElement).select()"
    />
    <div class="formula-actions">
      <Tooltip title="取消编辑（Esc）"><Button type="text" size="small" aria-label="取消编辑" :disabled="!editing" @mousedown.prevent @click="emit('cancel')"><X :size="15" /></Button></Tooltip>
      <Tooltip title="确认编辑（Enter）"><Button type="text" size="small" aria-label="确认编辑" :disabled="!editing" @mousedown.prevent @click="emit('confirm')"><Check :size="15" /></Button></Tooltip>
      <Tooltip title="编辑公式"><Button type="text" size="small" aria-label="编辑公式" @mousedown.prevent @click="startFormula"><span class="formula-symbol">fx</span></Button></Tooltip>
    </div>
    <Input.TextArea
      ref="editor"
      class="formula-input"
      aria-label="公式编辑"
      :rows="expanded ? 3 : 1"
      variant="borderless"
      spellcheck="false"
      @focus="emit('begin')"
      @update:value="emit('update:value', $event)"
      @keydown="onKeydown"
      :value="value"
    />
    <Tooltip :title="expanded ? '折叠公式栏' : '展开公式栏'">
      <Button class="formula-expand" type="text" size="small" :aria-label="expanded ? '折叠公式栏' : '展开公式栏'" :aria-expanded="expanded" @mousedown.prevent @click="toggleExpanded"><component :is="expanded ? ChevronUp : ChevronDown" :size="14" /></Button>
    </Tooltip>
    <span v-if="error" class="formula-error" role="alert">{{ error }}</span>
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
.formula-error { position: absolute; z-index: 8; top: 100%; inset-inline-start: 0; max-width: min(360px, 90vw); border: 1px solid var(--formula-error-border); padding: 5px 9px; color: var(--formula-error-text); background: var(--formula-error-bg); font-size: 12px; }
</style>
