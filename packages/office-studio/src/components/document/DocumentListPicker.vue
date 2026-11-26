<script setup lang="ts">
import { computed, ref } from 'vue'
import { Button, Form, InputNumber } from 'antdv-next'
import { listMarker, listPickerPresets, type ListKind, type ListPreset } from '../../features/document/model/document-lists'

const props = defineProps<{ kind: ListKind; active?: ListPreset; title: string; canRestart?: boolean; canContinue?: boolean; disabled?: boolean }>()
const emit = defineEmits<{ insert: [value: ListPreset]; execute: [commandId: string, params?: Record<string, unknown>] }>()
const numbering = ref<{ start: number | null }>({ start: 1 })
const composing = ref(false)
function restart() {
  if (props.disabled) return
  const start = numbering.value.start
  if (!composing.value && props.canRestart && typeof start === 'number' && Number.isSafeInteger(start) && start >= 0 && start <= 2147483647) emit('execute', 'doc.command.restart-list-numbering', { start })
}
function onNumberingKeydown(event: KeyboardEvent) {
  if ((event.key === 'Enter' || event.key === 'Escape') && (composing.value || event.isComposing || event.keyCode === 229)) {
    event.preventDefault()
    event.stopPropagation()
  }
}
const cards = computed(() => listPickerPresets[props.kind].map(preset => ({
  preset, preview: [0, 1, 2, 0].map(level => ({ level, glyph: listMarker(preset, level) })),
})))
</script>

<template>
  <div class="doc-list-picker" role="group" :aria-label="title" @keydown.capture="onNumberingKeydown">
    <Button v-for="card in cards" :key="card.preset" class="doc-list-card" :disabled="disabled" :aria-pressed="active === card.preset"
      :aria-label="`${title}: ${card.preview.slice(0, 3).map(row => row.glyph).join(' ')}`" :data-list-preset="card.preset" @mousedown.prevent @click="emit('insert', card.preset)">
      <span class="doc-list-preview" aria-hidden="true">
        <span v-for="(row, index) in card.preview" :key="index" class="doc-list-row" :style="{ paddingInlineStart: `${row.level * 8}px` }">
          <span>{{ row.glyph }}</span><i />
        </span>
      </span>
    </Button>
    <Form v-if="kind === 'numbered'" class="doc-numbering-actions" :model="numbering" :disabled="disabled" layout="vertical" @finish="restart">
      <Form.Item label="起始编号" name="start" :rules="[{ required: true, type: 'integer', min: 0, max: 2147483647 }]">
        <InputNumber v-model:value="numbering.start" :min="0" :max="2147483647" :precision="0" :disabled="disabled || !canRestart" aria-label="起始编号"
          @compositionstart="composing = true" @compositionend="composing = false" />
      </Form.Item>
      <Button html-type="submit" :disabled="disabled || !canRestart || composing">重新编号</Button>
      <Button :disabled="disabled || !canContinue" @mousedown.prevent @click="emit('execute', 'doc.command.continue-list-numbering')">继续前一编号</Button>
    </Form>
  </div>
</template>

<style scoped>
.doc-list-picker { display: grid; grid-template-columns: repeat(3, minmax(0, 72px)); gap: 8px; padding: 6px; }
.doc-list-card { width: 100%; height: 80px; padding: 0; overflow: hidden; border-radius: 4px; background: var(--ribbon-surface); color: var(--ribbon-text); }
.doc-list-card[aria-pressed="true"], .doc-list-card:not(:disabled):hover { border-color: var(--ribbon-primary); }
.doc-list-preview { box-sizing: border-box; width: 100%; height: 100%; padding: 8px; display: grid; grid-template-rows: repeat(4, 1fr); gap: 4px; }
.doc-list-row { display: flex; align-items: center; gap: 4px; min-width: 0; font-size: 10px; line-height: 1; }
.doc-list-row > span {
  width: 12px;
  flex: 0 0 12px;
  overflow: hidden;
  text-align: center;
}
.doc-list-row i { min-width: 2px; flex: 1; height: 4px; border-radius: 2px; background: var(--ribbon-border); }
.doc-numbering-actions { grid-column: 1 / -1; display: flex; flex-direction: column; gap: 6px; padding-block-start: 6px; border-block-start: 1px solid var(--ribbon-border); }
.doc-numbering-actions :deep(.ant-form-item) { margin-block-end: 0; }
.doc-numbering-actions :deep(.ant-input-number) { width: 100%; }
</style>
