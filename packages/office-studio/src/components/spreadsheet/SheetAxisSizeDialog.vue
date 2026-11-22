<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Form, FormItem, InputNumber, Modal } from 'antdv-next'
import type { SheetAxisSizeRequest } from './use-sheet-axis-dimensions.ts'

const props = defineProps<{ request: SheetAxisSizeRequest | null; error: string }>()
const emit = defineEmits<{ confirm: [value: number]; cancel: []; closed: [] }>()
const value = ref<number | null>(null), input = ref<InstanceType<typeof InputNumber> | null>(null), invalid = ref(''), composing = ref(false)
const maximum = computed(() => props.request?.axis === 'row' ? 2000 : undefined)
const errorMessage = computed(() => invalid.value || props.error)
watch(() => props.request, request => { value.value = request?.value ?? null; invalid.value = ''; composing.value = false }, { immediate: true })
function changeValue(next: number | string | null): void { value.value = typeof next === 'number' ? next : null; invalid.value = '' }
function confirm(event?: KeyboardEvent): void {
  if (composing.value || event?.isComposing) return
  if (value.value === null || !Number.isFinite(value.value) || value.value < 0 || maximum.value !== undefined && value.value > maximum.value) {
    invalid.value = maximum.value === undefined ? '请输入有限的非负尺寸。' : `请输入 0 到 ${maximum.value} 之间的尺寸。`; return
  }
  emit('confirm', value.value)
}
function afterOpenChange(open: boolean): void { if (open) input.value?.focus() }
</script>

<template>
  <Modal :open="Boolean(request)" :title="request?.axis === 'row' ? '行高' : '列宽'" ok-text="确定" cancel-text="取消" :width="360"
    :mask="{ closable: false }" :focus-trigger-after-close="false" :after-open-change="afterOpenChange"
    @ok="confirm()" @cancel="emit('cancel')" @after-close="emit('closed')">
    <Form layout="vertical" :model="{ size: value }" @submit.prevent>
      <FormItem name="size" :label="request?.axis === 'row' ? '行高（像素）' : '列宽（像素）'" :validate-status="errorMessage ? 'error' : undefined" :help="errorMessage || undefined">
        <InputNumber ref="input" :value="value" :min="0" :max="maximum" :step="1" :aria-label="request?.axis === 'row' ? '行高' : '列宽'"
          @update:value="changeValue" @press-enter="confirm" @compositionstart="composing = true" @compositionend="composing = false" />
      </FormItem>
    </Form>
  </Modal>
</template>
