<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { Alert, Button, Drawer, Form, FormItem, Input, InputNumber, Modal, Select, Switch, Upload } from 'antdv-next'
import { FileUp } from 'lucide-vue-next'
import type { DocumentImage } from '../../features/document/model/document-model'
import type { DocumentImagePatch, DocumentImagePosition, DocumentImageWrappingStyle } from '../../features/document/images/document-image-model'
import type { DocumentImageEditorMode } from '../../api/document-image-editors'

const props = defineProps<{ open: boolean; mode: DocumentImageEditorMode; image?: DocumentImage; editorKey: number; busy?: boolean; readOnly?: boolean; error?: string }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; source: [source: string | File, alt: string]; apply: [patch: DocumentImagePatch]; closed: []; composing: [value: boolean] }>()
const picking = computed(() => props.mode === 'insert' || props.mode === 'replace')
const title = computed(() => ({ insert: '插入图片', edit: '图片格式', crop: '裁剪图片', replace: '替换图片' })[props.mode])
const file = ref<File>(), source = ref(''), alt = ref(''), width = ref(320), height = ref(200), angle = ref(0), keepRatio = ref(true)
const flipX = ref(false), flipY = ref(false)
const composing = ref(false)
const sourceField = ref<{ focus(): void }>(), altField = ref<{ focus(): void }>()
function composition(value: boolean) { composing.value = value; emit('composing', value) }
function focusFields(open: boolean) { if (open) void nextTick(() => { if (props.open) (picking.value ? sourceField.value : altField.value)?.focus() }) }
const wrappingStyle = ref<DocumentImageWrappingStyle>('inline'), wrapSide = ref<'both' | 'left' | 'right' | 'largest'>('both')
const horizontal = ref<DocumentImagePosition>({ relativeFrom: 1, posOffset: 0 }), vertical = ref<DocumentImagePosition>({ relativeFrom: 1, posOffset: 0 })
const hMode = ref<'offset' | 'align' | 'percent'>('offset'), vMode = ref<'offset' | 'align' | 'percent'>('offset')
const crop = ref({ top: 0, right: 0, bottom: 0, left: 0 }), distances = ref({ top: 0, right: 0, bottom: 0, left: 0 })
const fields = [{ key: 'top', label: '上' }, { key: 'right', label: '右' }, { key: 'bottom', label: '下' }, { key: 'left', label: '左' }] as const
const wrappingOptions = [{ value: 'inline', label: '嵌入型' }, { value: 'wrapSquare', label: '四周型' }, { value: 'wrapTopAndBottom', label: '上下型' },
  { value: 'behindText', label: '衬于文字下方' }, { value: 'inFrontOfText', label: '浮于文字上方' }]
const hReference = ['页面', '栏', '字符', '页边距', '内侧页边距', '外侧页边距', '左侧页边距', '右侧页边距'].map((label, value) => ({ label, value }))
const vReference = ['页面', '段落', '行', '页边距', '上侧页边距', '下侧页边距', '内侧页边距', '外侧页边距'].map((label, value) => ({ label, value }))
const modeOptions = [{ value: 'offset', label: '绝对偏移' }, { value: 'align', label: '对齐' }, { value: 'percent', label: '相对百分比' }]
const hAlign = ['居中', '内侧', '左', '外侧', '右', '两端', '分散'].map((label, value) => ({ label, value }))
const vAlign = ['下', '居中', '内侧', '外侧', '上'].map((label, value) => ({ label, value }))
function initialize() {
  composition(false)
  const image = props.image
  file.value = undefined; source.value = ''; alt.value = image?.alt ?? ''; width.value = image?.width ?? 320; height.value = image?.height ?? 200
  angle.value = image?.angle ?? 0; keepRatio.value = image?.keepAspectRatio !== false; wrappingStyle.value = image?.wrappingStyle ?? 'inline'; wrapSide.value = image?.wrapSide ?? 'both'
  flipX.value = image?.flipX ?? false; flipY.value = image?.flipY ?? false
  horizontal.value = { ...image?.positionH, relativeFrom: image?.positionH?.relativeFrom ?? 1 }; vertical.value = { ...image?.positionV, relativeFrom: image?.positionV?.relativeFrom ?? 1 }
  hMode.value = horizontal.value.align !== undefined ? 'align' : horizontal.value.percent !== undefined ? 'percent' : 'offset'
  vMode.value = vertical.value.align !== undefined ? 'align' : vertical.value.percent !== undefined ? 'percent' : 'offset'
  crop.value = { top: (image?.crop?.top ?? 0) * 100, right: (image?.crop?.right ?? 0) * 100, bottom: (image?.crop?.bottom ?? 0) * 100, left: (image?.crop?.left ?? 0) * 100 }
  distances.value = { top: image?.distances?.top ?? 0, right: image?.distances?.right ?? 0, bottom: image?.distances?.bottom ?? 0, left: image?.distances?.left ?? 0 }
}
watch(() => props.editorKey, initialize, { immediate: true })
function chooseFile(value: File): false { file.value = value; if (!alt.value) alt.value = value.name; return false }
function resizeWidth(value: number | null) { if (value === null) return; if (keepRatio.value) height.value *= value / width.value; width.value = value }
function resizeHeight(value: number | null) { if (value === null) return; if (keepRatio.value) width.value *= value / height.value; height.value = value }
function positioning(value: DocumentImagePosition, mode: typeof hMode.value): DocumentImagePosition {
  return { relativeFrom: value.relativeFrom, ...(mode === 'align' ? { align: value.align ?? 0 } : mode === 'percent' ? { percent: value.percent ?? 0 } : { posOffset: value.posOffset ?? 0 }) }
}
function submit() {
  if (props.busy || props.readOnly || composing.value) return
  if (picking.value) { if (file.value || source.value.trim()) emit('source', file.value ?? source.value.trim(), alt.value); return }
  emit('apply', { alt: alt.value, width: width.value, height: height.value, angle: angle.value, flipX: flipX.value, flipY: flipY.value, keepAspectRatio: keepRatio.value,
    wrappingStyle: wrappingStyle.value, wrapSide: wrapSide.value, positionH: positioning(horizontal.value, hMode.value), positionV: positioning(vertical.value, vMode.value),
    crop: { top: crop.value.top / 100, right: crop.value.right / 100, bottom: crop.value.bottom / 100, left: crop.value.left / 100 }, distances: { ...distances.value } })
}
</script>

<template>
  <Modal v-if="picking" :open="open" :title="title" ok-text="确定" cancel-text="取消" :confirm-loading="busy" :focus-trigger-after-close="false" :keyboard="!composing" :closable="!composing" :mask-closable="!composing" :cancel-button-props="{ disabled: composing }" destroy-on-hidden
    :after-open-change="focusFields" :auto-focus-button="null" :after-close="() => emit('closed')" :ok-button-props="{ disabled: readOnly || composing || (!file && !source.trim()) }" @update:open="value => emit('update:open', value)" @ok="submit">
    <Form layout="vertical" :disabled="readOnly || busy" @submit.prevent="submit" @compositionstart="composition(true)" @compositionend="composition(false)">
      <FormItem label="本地图片"><Upload accept="image/png,image/jpeg,image/gif,image/webp,image/bmp" :show-upload-list="false" :before-upload="chooseFile" :disabled="readOnly || busy">
        <Button :disabled="readOnly || busy"><FileUp :size="15" />选择图片</Button>
      </Upload><span v-if="file" class="selected-file">{{ file.name }} <Button type="link" size="small" :disabled="readOnly || busy || composing" @click="file = undefined">移除</Button></span></FormItem>
      <FormItem label="图片地址"><Input ref="sourceField" v-model:value="source" aria-label="图片地址" :disabled="readOnly || busy || !!file" placeholder="https://" allow-clear @press-enter="submit" @compositionstart="composition(true)" @compositionend="composition(false)" /></FormItem>
      <FormItem label="替代文字"><Input v-model:value="alt" aria-label="替代文字" :maxlength="32768" @compositionstart="composition(true)" @compositionend="composition(false)" /></FormItem>
    </Form>
    <Alert v-if="error" :message="error" type="error" show-icon role="alert" />
  </Modal>
  <Drawer v-else :open="open" :title="title" :size="420" :keyboard="!composing" :closable="!composing" :mask-closable="!composing" destroy-on-hidden :after-open-change="value => { if (value) focusFields(value); else emit('closed') }"
    @update:open="value => emit('update:open', value)">
    <Form layout="vertical" :disabled="readOnly || busy" @submit.prevent="submit" @compositionstart="composition(true)" @compositionend="composition(false)" class="document-image-fields">
      <FormItem label="替代文字"><Input.TextArea ref="altField" v-model:value="alt" aria-label="替代文字" :maxlength="32768" :auto-size="{ minRows: 2, maxRows: 4 }" @compositionstart="composition(true)" @compositionend="composition(false)" /></FormItem>
      <div class="field-grid">
        <FormItem label="宽度（px）"><InputNumber :value="width" aria-label="宽度（px）" :min="0.1" :max="100000" :precision="2" @update:value="resizeWidth" /></FormItem>
        <FormItem label="高度（px）"><InputNumber :value="height" aria-label="高度（px）" :min="0.1" :max="100000" :precision="2" @update:value="resizeHeight" /></FormItem>
        <FormItem label="锁定纵横比"><Switch v-model:checked="keepRatio" aria-label="锁定纵横比" /></FormItem>
        <FormItem label="旋转（度）"><InputNumber v-model:value="angle" aria-label="旋转（度）" :min="-360" :max="360" :precision="2" /></FormItem>
        <FormItem label="水平翻转"><Switch v-model:checked="flipX" aria-label="水平翻转" /></FormItem>
        <FormItem label="垂直翻转"><Switch v-model:checked="flipY" aria-label="垂直翻转" /></FormItem>
      </div>
      <FormItem label="文字环绕"><Select v-model:value="wrappingStyle" aria-label="文字环绕" :options="wrappingOptions" /></FormItem>
      <template v-if="wrappingStyle !== 'inline'">
        <FormItem v-if="wrappingStyle === 'wrapSquare'" label="文字位置"><Select v-model:value="wrapSide" aria-label="文字位置" :options="[{ value: 'both', label: '两侧' }, { value: 'left', label: '仅左侧' }, { value: 'right', label: '仅右侧' }, { value: 'largest', label: '宽度较大的一侧' }]" /></FormItem>
        <div class="field-grid">
          <FormItem label="水平基准"><Select v-model:value="horizontal.relativeFrom" aria-label="水平基准" :options="hReference" /></FormItem>
          <FormItem label="水平定位"><Select v-model:value="hMode" aria-label="水平定位" :options="modeOptions" /></FormItem>
          <FormItem v-if="hMode === 'align'" label="水平对齐"><Select v-model:value="horizontal.align" aria-label="水平对齐" :options="hAlign" /></FormItem>
          <FormItem v-else-if="hMode === 'percent'" label="水平位置（%）"><InputNumber v-model:value="horizontal.percent" aria-label="水平位置（%）" :min="-1000" :max="1000" /></FormItem>
          <FormItem v-else label="水平偏移（px）"><InputNumber v-model:value="horizontal.posOffset" aria-label="水平偏移（px）" :min="-100000" :max="100000" /></FormItem>
          <FormItem label="垂直基准"><Select v-model:value="vertical.relativeFrom" aria-label="垂直基准" :options="vReference" /></FormItem>
          <FormItem label="垂直定位"><Select v-model:value="vMode" aria-label="垂直定位" :options="modeOptions" /></FormItem>
          <FormItem v-if="vMode === 'align'" label="垂直对齐"><Select v-model:value="vertical.align" aria-label="垂直对齐" :options="vAlign" /></FormItem>
          <FormItem v-else-if="vMode === 'percent'" label="垂直位置（%）"><InputNumber v-model:value="vertical.percent" aria-label="垂直位置（%）" :min="-1000" :max="1000" /></FormItem>
          <FormItem v-else label="垂直偏移（px）"><InputNumber v-model:value="vertical.posOffset" aria-label="垂直偏移（px）" :min="-100000" :max="100000" /></FormItem>
        </div>
        <div class="field-grid"><FormItem v-for="field in fields" :key="field.key" :label="`${field.label}侧文字距离（px）`"><InputNumber v-model:value="distances[field.key]" :aria-label="`${field.label}侧文字距离（px）`" :min="0" :max="100000" :precision="2" /></FormItem></div>
      </template>
      <div class="field-grid"><FormItem v-for="field in fields" :key="field.key" :label="`${field.label}侧裁剪（%）`"><InputNumber v-model:value="crop[field.key]" :aria-label="`${field.label}侧裁剪（%）`" :min="-100000" :max="99.99" :precision="2" /></FormItem></div>
      <Button @click="crop = { top: 0, right: 0, bottom: 0, left: 0 }" :disabled="readOnly || busy || composing">清除裁剪</Button>
    </Form>
    <Alert v-if="error" :message="error" type="error" show-icon role="alert" />
    <template #footer><div class="panel-actions"><Button :disabled="composing" @click="emit('update:open', false)">取消</Button><Button type="primary" :loading="busy" :disabled="readOnly || composing" @click="submit">应用</Button></div></template>
  </Drawer>
</template>

<style scoped>
.field-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 0 12px; }
.field-grid :deep(.ant-input-number), .field-grid :deep(.ant-select) { width: 100%; }
.panel-actions { display: flex; justify-content: flex-end; gap: 8px; }
.selected-file { display: flex; align-items: center; gap: 8px; margin-top: 8px; overflow-wrap: anywhere; }
.document-image-fields :deep(.ant-form-item) { margin-bottom: 16px; }
</style>
