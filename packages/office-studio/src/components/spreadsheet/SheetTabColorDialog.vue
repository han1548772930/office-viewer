<script setup lang="ts">
import { ref, watch } from 'vue'
import { ColorPicker, Form, FormItem, Modal } from 'antdv-next'

const props = defineProps<{ request: { sheetId: string; name: string; color?: string } | null }>()
const emit = defineEmits<{
  (event: 'confirm', color: string): void
  (event: 'cancel'): void
  (event: 'closed'): void
}>()
const color = ref<string | undefined>()
watch(() => props.request, request => { color.value = request?.color ?? '#4472c4' }, { immediate: true })
function confirm() { if (props.request) emit('confirm', color.value ?? '') }
</script>

<template>
  <Modal :open="Boolean(request)" :title="`标签颜色 · ${request?.name ?? ''}`" ok-text="确定" cancel-text="取消" :width="360" @ok="confirm" @cancel="emit('cancel')" @after-close="emit('closed')">
    <Form layout="vertical" :model="{ color }" @finish="confirm">
      <FormItem label="标签颜色" name="color">
        <ColorPicker v-model:value="color" value-format="hex" allow-clear show-text aria-label="工作表标签颜色" />
      </FormItem>
    </Form>
  </Modal>
</template>
