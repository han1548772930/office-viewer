<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Alert, Button, Form, FormItem, Modal, Upload, type UploadFile } from 'antdv-next'
import { ImagePlus } from 'lucide-vue-next'
import type { SheetImagePickerRequest } from './use-sheet-images.ts'
import { SHEET_IMAGE_ACCEPT, SHEET_IMAGE_MAX_BATCH_BYTES, SHEET_IMAGE_MAX_FILE_BYTES } from '../../features/sheet-images/sheet-image-model.ts'

const props = defineProps<{ request: SheetImagePickerRequest | null; busy: boolean; error: string; progress: string }>()
const emit = defineEmits<{ confirm: [files: File[]]; cancel: []; closed: [] }>()
const fileList = ref<UploadFile[]>([])
const files = computed(() => fileList.value.flatMap(item => item instanceof File ? [item] : item.originFileObj ? [item.originFileObj] : []))
const totalBytes = computed(() => files.value.reduce((total, file) => total + file.size, 0))
const fileError = computed(() => files.value.some(file => !file.size || file.size > SHEET_IMAGE_MAX_FILE_BYTES) ? '每张图片应非空且不超过 5 MiB。'
  : totalBytes.value > SHEET_IMAGE_MAX_BATCH_BYTES ? '本次图片总大小超过 32 MiB。' : '')
const title = computed(() => props.request?.kind === 'background' ? '设置工作表背景图' : props.request?.kind === 'cell' ? '插入单元格图片' : '插入浮动图片')
watch(() => props.request, () => { fileList.value = [] }, { flush: 'sync' })
</script>

<template>
  <Modal :open="Boolean(request)" :title="title" :width="480" :confirm-loading="busy" :focus-trigger-after-close="false"
    :ok-button-props="{ disabled: !files.length || Boolean(fileError) }" ok-text="插入" cancel-text="取消" destroy-on-hidden
    :mask="{ closable: false }" @ok="emit('confirm', files)" @cancel="emit('cancel')" @after-close="emit('closed')">
    <Form layout="vertical" :preserve="false" @submit.prevent>
      <FormItem label="图片文件" :validate-status="fileError ? 'error' : undefined" :help="fileError || 'PNG、JPEG、GIF、BMP、WebP；每张最多 5 MiB，本次总计最多 32 MiB。'">
        <Upload v-model:file-list="fileList" :accept="SHEET_IMAGE_ACCEPT" :multiple="request?.multiple" :max-count="request?.multiple ? 1000 : 1"
          :before-upload="() => false" :disabled="busy" list-type="text">
          <Button :disabled="busy"><ImagePlus :size="16" />选择图片</Button>
        </Upload>
      </FormItem>
      <p v-if="request?.kind === 'cell'">图片将替换目标单元格内容。</p>
      <p v-if="request?.kind === 'background'">背景图按图片分辨率在工作表中重复平铺。</p>
      <p v-if="progress" role="status">正在读取 {{ progress }}</p>
      <Alert v-if="error" type="error" :title="error" show-icon />
    </Form>
  </Modal>
</template>
