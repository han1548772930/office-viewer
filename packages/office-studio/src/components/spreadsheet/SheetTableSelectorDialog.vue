<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Alert, Button, Form, FormItem, Input, Modal } from 'antdv-next'
import type { CellRange } from '../../domain/model.ts'
import { formatValidationRange } from '../../features/sheet-validation/validation-panel-state.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { useRibbonTheme } from '../../platform/ribbon/ribbon-theme.ts'

const props = defineProps<{ open: boolean; range: CellRange; error: string; canSubmit: boolean; submit(range: string): boolean }>()
const emit = defineEmits<{ cancel: []; closed: [] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const style = useRibbonTheme(), draft = ref(''), composing = ref(false)
watch(() => props.open, open => { if (open) { draft.value = formatValidationRange(props.range); composing.value = false } }, { immediate: true })
const disabled = computed(() => !props.canSubmit || composing.value || !draft.value.trim())
function confirm(): void { if (!disabled.value) props.submit(draft.value.trim()) }
function cancel(): void { if (!composing.value) emit('cancel') }
</script>

<template>
  <Modal :open="open" :width="300" :title="t('sheet.table.create', '创建表格')" :footer="null" :keyboard="!composing" :style="style"
    :focus-trigger-after-close="false" :after-close="() => emit('closed')" @cancel="cancel">
    <Form layout="vertical" class="sheet-table-selector" @submit.prevent="confirm">
      <FormItem :label="t('sheet.table.range', '表格范围')" name="range">
        <Input v-model:value="draft" :auto-focus="true" :disabled="!canSubmit" :aria-label="t('sheet.table.range', '表格范围')" @compositionstart="composing = true" @compositionend="composing = false" />
      </FormItem>
      <Alert v-if="error" type="error" show-icon :title="error" />
      <div class="sheet-table-selector-actions">
        <Button :disabled="composing" @click="cancel">{{ t('sheet.table.cancel', '取消') }}</Button>
        <Button type="primary" html-type="submit" :disabled="disabled">{{ t('sheet.table.confirm', '确认') }}</Button>
      </div>
    </Form>
  </Modal>
</template>

<style scoped>
.sheet-table-selector-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
</style>
