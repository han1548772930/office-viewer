<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Button, Input, Modal } from 'antdv-next'
import { previewCustomNumberFormat, validateCustomNumberFormat } from '../../features/sheet-workbench/custom-number-format.ts'
import type { CellScalar } from '../../domain/model.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; currentPattern: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [pattern: string] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)

const pattern = ref('')
const sample = ref('1234.567')
const error = ref('')
const example = computed(() => {
  try {
    const result = validateCustomNumberFormat(pattern.value)
    const numeric = Number(sample.value)
    const value: CellScalar = sample.value.trim() && Number.isFinite(numeric) ? numeric : sample.value
    error.value = ''
    return previewCustomNumberFormat(result.pattern, value)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('sheet.number-format.invalid', '格式代码无效。')
    return undefined
  }
})

watch(() => [props.open, props.currentPattern] as const, ([open, current]) => {
  if (open) {
    pattern.value = current || 'General'
    sample.value = '1234.567'
    error.value = ''
  }
})

function apply() {
  try {
    const { pattern: value } = validateCustomNumberFormat(pattern.value)
    emit('apply', value)
    emit('update:open', false)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('sheet.number-format.invalid', '格式代码无效。')
  }
}
</script>

<template>
  <Modal :open="open" :title="t('sheet.number-format.title', '自定义数字格式')" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.cancel', '取消')" :width="480" @ok="apply" @cancel="emit('update:open', false)">
    <div class="number-format-panel">
      <label>{{ t('sheet.number-format.pattern', '格式代码') }}<Input v-model:value="pattern" :aria-label="t('sheet.number-format.pattern-aria', '自定义数字格式代码')" :maxlength="255" /></label>
      <label>{{ t('sheet.number-format.sample', '示例值') }}<Input v-model:value="sample" :aria-label="t('sheet.number-format.sample-aria', '自定义格式示例值')" /></label>
      <div class="number-format-preview" aria-live="polite">
        <span>{{ t('sheet.number-format.preview', '预览') }}</span><strong>{{ example ?? t('sheet.number-format.unavailable', '无法预览') }}</strong>
      </div>
      <p v-if="error" class="number-format-error" role="alert">{{ error }}</p>
      <div class="number-format-presets" :aria-label="t('sheet.number-format.presets', '常用格式')">
        <Button size="small" @click="pattern = '#,##0.00'">{{ t('sheet.number-format.number', '数字') }}</Button>
        <Button size="small" @click="pattern = 'yyyy-mm-dd'">{{ t('sheet.number-format.date', '日期') }}</Button>
        <Button size="small" @click="pattern = 'hh:mm:ss'">{{ t('sheet.number-format.time', '时间') }}</Button>
        <Button size="small" @click="pattern = '0.00%'">{{ t('sheet.number-format.percent', '百分比') }}</Button>
        <Button size="small" @click="pattern = '@'">{{ t('sheet.number-format.text', '文本') }}</Button>
      </div>
    </div>
  </Modal>
</template>

<style scoped>
.number-format-panel { display: grid; gap: 14px; }
.number-format-panel label { display: grid; gap: 6px; color: #475467; font-size: 13px; }
.number-format-preview { display: flex; min-height: 36px; align-items: center; justify-content: space-between; gap: 12px; padding: 0 10px; border: 1px solid #e4e7ec; border-radius: 4px; color: #667085; font-size: 12px; }
.number-format-preview strong { overflow-wrap: anywhere; color: #1d2939; font-size: 14px; font-weight: 500; }
.number-format-error { margin: -6px 0 0; color: #b42318; font-size: 12px; }
.number-format-presets { display: flex; flex-wrap: wrap; gap: 6px; }
</style>
