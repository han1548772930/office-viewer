<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Alert, Button, Checkbox, Form, FormItem, Input, InputNumber, Modal, Select, Slider, Upload, type UploadFile } from 'antdv-next'
import { ImagePlus } from 'lucide-vue-next'
import type { SheetImageEditorRequest } from './use-sheet-images.ts'
import type { SheetImageCommandInput } from '../../features/sheet-images/sheet-image-commands.ts'
import { applySheetImageEditorCropRatio, createSheetImageEditorDraft, resizeSheetImageEditorDraft, sheetImagePatchFromDraft, type SheetImageEditorDraft } from '../../features/sheet-images/sheet-image-editor-state.ts'
import { SHEET_IMAGE_ACCEPT } from '../../features/sheet-images/sheet-image-model.ts'
import { SHEET_IMAGE_CROP_PRESETS, type SheetImageCropType } from '../../features/sheet-images/sheet-image-crop.ts'
import SheetImagePreview from './SheetImagePreview.vue'
const props = defineProps<{ request: SheetImageEditorRequest | null; busy: boolean; error: string }>()
const emit = defineEmits<{ confirm: [patch: SheetImageCommandInput['patch'], files: File[]]; cancel: []; closed: [] }>()
const draft = ref<SheetImageEditorDraft>(), fileList = ref<UploadFile[]>([]), draftError = ref('')
const cropPreset = ref<SheetImageCropType>('0')
const anchorOptions = [{ value: '0', label: '随单元格移动，保持大小' }, { value: '1', label: '随单元格移动和缩放' }, { value: '2', label: '固定位置和大小' }]
const files = computed(() => fileList.value.flatMap(item => item instanceof File ? [item] : item.originFileObj ? [item.originFileObj] : []))
watch(() => props.request, request => {
  draft.value = request ? createSheetImageEditorDraft(request.image, request.row, request.column, request.axes) : undefined
  fileList.value = []; draftError.value = ''; cropPreset.value = '0'
}, { immediate: true, flush: 'sync' })
function resize(axis: 'width' | 'height', value: unknown): void {
  if (!draft.value) return
  try { resizeSheetImageEditorDraft(draft.value, axis, value); draftError.value = '' }
  catch (reason: unknown) { draftError.value = reason instanceof Error ? reason.message : '图片尺寸无效。' }
}
function applyCropPreset(): void {
  const ratio = SHEET_IMAGE_CROP_PRESETS.find(preset => preset.value === cropPreset.value)?.ratio
  if (!draft.value || ratio === undefined) return
  try { applySheetImageEditorCropRatio(draft.value, ratio); draftError.value = '' }
  catch (reason: unknown) { draftError.value = reason instanceof Error ? reason.message : '无法应用裁剪比例。' }
}
function confirm(): void {
  if (!draft.value || props.busy) return
  try { const patch = sheetImagePatchFromDraft(draft.value); draftError.value = ''; emit('confirm', patch, files.value) }
  catch (reason: unknown) { draftError.value = reason instanceof Error ? reason.message : '图片属性无效。' }
}
</script>
<template>
  <Modal :open="Boolean(request)" title="图片属性" :width="560" :confirm-loading="busy" :focus-trigger-after-close="false"
    :mask="{ closable: false }" ok-text="应用" cancel-text="取消" destroy-on-hidden @ok="confirm" @cancel="emit('cancel')" @after-close="emit('closed')">
    <Form v-if="draft && request" layout="vertical" :model="draft" :preserve="false" :disabled="busy" @submit.prevent="confirm">
      <SheetImagePreview :resource="request.image.resource" :frame="draft.floating ? { width: draft.width, height: draft.height } : undefined"
        :crop="draft.floating ? draft.crop : undefined" :rotation="draft.rotation" :flip-x="draft.flipX" :flip-y="draft.flipY" :opacity="draft.opacity / 100" />
      <FormItem v-if="draft.floating" label="单元格锚点"><Select v-model:value="draft.anchorType" :options="anchorOptions" aria-label="图片锚点类型" /></FormItem>
      <div v-if="draft.floating" class="image-fields">
        <template v-if="draft.anchorType !== '2'">
          <FormItem label="锚点行号"><InputNumber v-model:value="draft.row" :min="1" :max="1048576" :precision="0" aria-label="图片锚点行号" /></FormItem>
          <FormItem label="锚点列号"><InputNumber v-model:value="draft.column" :min="1" :max="16384" :precision="0" aria-label="图片锚点列号" /></FormItem>
          <FormItem label="水平偏移（像素）"><InputNumber v-model:value="draft.offsetX" aria-label="图片水平偏移" /></FormItem>
          <FormItem label="垂直偏移（像素）"><InputNumber v-model:value="draft.offsetY" aria-label="图片垂直偏移" /></FormItem>
        </template>
        <template v-else>
          <FormItem label="水平位置（像素）"><InputNumber v-model:value="draft.left" aria-label="图片水平位置" /></FormItem>
          <FormItem label="垂直位置（像素）"><InputNumber v-model:value="draft.top" aria-label="图片垂直位置" /></FormItem>
        </template>
        <FormItem label="宽度（像素）"><InputNumber :value="draft.width" :min="1" :max="100000" :precision="2" aria-label="图片宽度" @change="resize('width', $event)" /></FormItem>
        <FormItem label="高度（像素）"><InputNumber :value="draft.height" :min="1" :max="100000" :precision="2" aria-label="图片高度" @change="resize('height', $event)" /></FormItem>
      </div>
      <FormItem v-if="draft.floating" label="缩放"><Checkbox v-model:checked="draft.keepRatio" @change="draft.ratio = draft.width / draft.height">锁定纵横比</Checkbox></FormItem>
      <FormItem v-if="draft.floating" label="不透明度"><Slider v-model:value="draft.opacity" :min="0" :max="100" aria-label="图片不透明度" /></FormItem>
      <div v-if="draft.floating" class="image-fields">
        <FormItem label="旋转（度）"><InputNumber v-model:value="draft.rotation" :step="15" aria-label="图片旋转角度" /></FormItem>
        <FormItem label="翻转"><Checkbox v-model:checked="draft.flipX">水平翻转</Checkbox><Checkbox v-model:checked="draft.flipY">垂直翻转</Checkbox></FormItem>
      </div>
      <FormItem v-if="draft.floating" label="裁剪比例">
        <div class="crop-preset"><Select v-model:value="cropPreset" :options="SHEET_IMAGE_CROP_PRESETS" aria-label="图片裁剪比例" />
          <Button :disabled="cropPreset === '0' || !request.axes" @click="applyCropPreset">应用比例</Button>
          <Button @click="draft.crop = { top: 0, left: 0, bottom: 0, right: 0 }">重置裁剪</Button>
        </div>
      </FormItem>
      <div v-if="draft.floating" class="image-fields">
        <FormItem label="顶部裁剪（像素）"><InputNumber v-model:value="draft.crop.top" :min="-100000" :max="100000" aria-label="图片顶部裁剪" /></FormItem>
        <FormItem label="左侧裁剪（像素）"><InputNumber v-model:value="draft.crop.left" :min="-100000" :max="100000" aria-label="图片左侧裁剪" /></FormItem>
        <FormItem label="底部裁剪（像素）"><InputNumber v-model:value="draft.crop.bottom" :min="-100000" :max="100000" aria-label="图片底部裁剪" /></FormItem>
        <FormItem label="右侧裁剪（像素）"><InputNumber v-model:value="draft.crop.right" :min="-100000" :max="100000" aria-label="图片右侧裁剪" /></FormItem>
      </div>
      <FormItem label="标题"><Input v-model:value="draft.title" :maxlength="4096" aria-label="图片标题" /></FormItem>
      <FormItem label="替代文本"><Input.TextArea v-model:value="draft.altText" :maxlength="4096" :rows="2" aria-label="图片替代文本" /></FormItem>
      <FormItem label="替换图片">
        <Upload v-model:file-list="fileList" :accept="SHEET_IMAGE_ACCEPT" :max-count="1" :before-upload="() => false" :disabled="busy" list-type="text">
          <Button :disabled="busy"><ImagePlus :size="16" />选择替换图片</Button>
        </Upload>
      </FormItem>
      <Alert v-if="draftError || error" type="error" :title="draftError || error" show-icon />
    </Form>
  </Modal>
</template>
<style scoped>
.image-fields { display: grid; grid-template-columns: 1fr 1fr; column-gap: 16px; }
.image-fields :deep(.ant-input-number) { width: 100%; }
.crop-preset { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
.crop-preset :deep(.ant-select) { min-width: 112px; }
@media (max-width: 480px) { .image-fields { grid-template-columns: 1fr; } }
</style>
