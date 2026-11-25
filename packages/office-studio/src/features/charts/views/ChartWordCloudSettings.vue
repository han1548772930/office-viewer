<script setup lang="ts">
import { computed, onBeforeUnmount, ref, watch } from 'vue'
import { Alert, Button, FormItem, Input, InputNumber, Select, Switch, Upload } from 'antdv-next'
import type { ChartWordCloudOptions } from '../../../domain/model'
import type { ChartConfig } from '../chart-engine'
import { cloneChartDraftConfig } from '../chart-settings-draft'
import { validateChartOptions } from '../chart-option-validation'
import { readChartImageFile } from '../chart-image-resources'
import { CHART_WORD_CLOUD_MASK_CHOICES } from '../chart-word-cloud-mask'
import { CHART_WORD_CLOUD_DEFAULTS, ChartWordCloudShapes } from '../chart-word-cloud'

const props = defineProps<{ config: ChartConfig; active?: boolean; sessionKey?: number | string; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig]; 'resource-loading': [value: boolean] }>()
const options = computed(() => props.config.options?.wordCloud ?? {})
const maskId = computed(() => options.value.maskImageId ?? options.value.maskImage)
const builtinIds = new Set<string>(CHART_WORD_CLOUD_MASK_CHOICES.map(([id]) => id))
const sourceMode = ref<'shape' | 'builtin' | 'resource'>(modeFromConfig())
const loading = ref(false), error = ref('')
let generation = 0, owner: AbortController | undefined
const shapeNames: Record<string, string> = { cloud: '云形', rect: '矩形', circle: '圆形', cardioid: '心迹线', heart: '心形',
  diamond: '菱形', 'triangle-forward': '右三角', triangle: '三角形', pentagon: '五边形', star: '星形' }
const shapeChoices = Object.values(ChartWordCloudShapes).map(value => ({ value, label: shapeNames[value] ?? value }))
const maskChoices = CHART_WORD_CLOUD_MASK_CHOICES.map(([value, label]) => ({ value, label }))
const sourceChoices = [{ value: 'shape', label: '几何形状' }, { value: 'builtin', label: '内置遮罩' }, { value: 'resource', label: '图片资源' }]
const unavailable = computed(() => props.disabled || props.active === false)
const minFontSize = computed(() => options.value.minFontSize ?? CHART_WORD_CLOUD_DEFAULTS.minFontSize)
const maxFontSize = computed(() => options.value.maxFontSize ?? CHART_WORD_CLOUD_DEFAULTS.maxFontSize)
function modeFromConfig(): 'shape' | 'builtin' | 'resource' {
  const id = props.config.options?.wordCloud?.maskImageId ?? props.config.options?.wordCloud?.maskImage
  return id ? builtinIds.has(id) ? 'builtin' : 'resource' : 'shape'
}
function setLoading(value: boolean): void { loading.value = value; emit('resource-loading', value) }
function cancelRead(): void { generation++; owner?.abort(); owner = undefined; setLoading(false) }
watch(() => [props.config, props.active, props.disabled, props.sessionKey], () => {
  cancelRead()
  if (maskId.value) sourceMode.value = modeFromConfig()
}, { flush: 'sync' })
watch(() => props.sessionKey, () => { sourceMode.value = modeFromConfig(); error.value = '' }, { flush: 'sync' })
onBeforeUnmount(cancelRead)
function edit(patch: Partial<ChartWordCloudOptions>): void {
  if (unavailable.value) return
  error.value = ''
  try {
    const config = cloneChartDraftConfig(props.config)
    config.options = { ...config.options, wordCloud: { ...config.options?.wordCloud, ...patch } }
    validateChartOptions(config.options)
    emit('update:config', config)
  } catch (cause) { error.value = cause instanceof Error ? cause.message : '无法更新词云设置。' }
}
function clearMask(): void {
  cancelRead(); sourceMode.value = 'shape'
  edit({ maskImageId: undefined, maskImage: undefined, maskImageResource: undefined })
}
function sourceChanged(value: unknown): void {
  if (unavailable.value || value !== 'shape' && value !== 'builtin' && value !== 'resource') return
  cancelRead(); sourceMode.value = value
  if (value === 'shape') clearMask()
  else if (value === 'builtin') edit({ maskImageId: builtinIds.has(maskId.value ?? '') ? maskId.value : 'Cloud', maskImage: undefined, maskImageResource: undefined })
  else if (builtinIds.has(maskId.value ?? '')) edit({ maskImageId: undefined, maskImage: undefined, maskImageResource: undefined })
}
function maskChanged(value: unknown): void {
  if (typeof value !== 'string' && value !== undefined) return
  edit({ maskImageId: value?.trim() ? value : undefined, maskImage: undefined, maskImageResource: undefined })
}
function shapeChanged(value: unknown): void {
  if (Object.values(ChartWordCloudShapes).some(shape => shape === value)) edit({ shape: value as ChartWordCloudOptions['shape'] })
}
function numberChanged(key: 'minFontSize' | 'maxFontSize' | 'rotation', value: unknown): void {
  if (value === null || value === undefined) edit({ [key]: undefined })
  else if (typeof value === 'number' && Number.isFinite(value)) edit({ [key]: value })
}
function repeatChanged(value: unknown): void { if (typeof value === 'boolean') edit({ repeat: value }) }
/** beforeUpload=false keeps antdv-next out of the network. The host publishes
 * this detached payload and its reference together through the existing Apply. */
async function beforeUpload(file: File): Promise<false> {
  if (unavailable.value) return false
  cancelRead(); error.value = ''
  const current = ++generation, config = props.config, session = props.sessionKey, controller = new AbortController()
  owner = controller; setLoading(true)
  try {
    const resource = await readChartImageFile(file, controller.signal)
    if (controller.signal.aborted || generation !== current || props.config !== config || props.sessionKey !== session || unavailable.value) return false
    sourceMode.value = 'resource'
    edit({ maskImageId: resource.id, maskImage: undefined, maskImageResource: resource })
  } catch (cause) {
    if (!controller.signal.aborted && generation === current) error.value = cause instanceof Error ? cause.message : '词云图片读取失败。'
  } finally { if (owner === controller) { owner = undefined; setLoading(false) } }
  return false
}
</script>

<template>
  <div class="chart-word-cloud-settings">
    <FormItem label="轮廓来源"><Select :value="sourceMode" :options="sourceChoices" aria-label="词云轮廓来源" :disabled="unavailable" @update:value="sourceChanged" /></FormItem>
    <FormItem v-if="sourceMode === 'shape'" label="词云形状">
      <Select :value="options.shape ?? CHART_WORD_CLOUD_DEFAULTS.shape" :options="shapeChoices" aria-label="词云形状" :disabled="unavailable" @update:value="shapeChanged" />
      <Button type="link" :disabled="unavailable || options.shape === undefined" @click="edit({ shape: undefined })">恢复默认形状</Button>
    </FormItem>
    <FormItem v-if="sourceMode === 'builtin'" label="遮罩图形"><Select :value="maskId" :options="maskChoices" aria-label="词云内置遮罩" :disabled="unavailable" @update:value="maskChanged" /></FormItem>
    <template v-if="sourceMode === 'resource'">
      <FormItem label="图片资源 ID 或地址" extra="可上传图片，或填写已注册的资源 ID、HTTP(S) 地址及明确的相对路径。">
        <Input :value="maskId ?? ''" aria-label="词云遮罩图片资源" :disabled="unavailable || loading" @update:value="maskChanged" />
      </FormItem>
      <FormItem label="上传遮罩图片" extra="单张图片不超过 8 MiB；点击应用后保存到当前工作簿。">
        <Upload accept="image/png,image/jpeg,image/gif,image/webp,image/avif,image/bmp,image/svg+xml,image/x-icon,image/vnd.microsoft.icon" :before-upload="beforeUpload" :show-upload-list="false" :disabled="unavailable || loading">
          <Button :loading="loading" :disabled="unavailable || loading" aria-label="上传词云遮罩图片">选择图片</Button>
        </Upload>
        <span v-if="options.maskImageResource?.name" class="chart-mask-name">{{ options.maskImageResource.name }}</span>
      </FormItem>
    </template>
    <Button v-if="sourceMode !== 'shape'" :disabled="unavailable || (!maskId && !loading)" @click="clearMask">清除遮罩并恢复形状</Button>
    <FormItem label="重复填充词语"><Switch :checked="options.repeat ?? false" aria-label="词云重复填充" :disabled="unavailable" @update:checked="repeatChanged" />
      <Button type="link" :disabled="unavailable || options.repeat === undefined" @click="edit({ repeat: undefined })">恢复默认重复设置</Button>
    </FormItem>
    <div class="chart-word-cloud-sizes">
      <FormItem label="最小字号"><InputNumber :value="minFontSize" :min="1" :max="maxFontSize" aria-label="词云最小字号" :disabled="unavailable" @update:value="value => numberChanged('minFontSize', value)" /></FormItem>
      <FormItem label="最大字号"><InputNumber :value="maxFontSize" :min="minFontSize" :max="256" aria-label="词云最大字号" :disabled="unavailable" @update:value="value => numberChanged('maxFontSize', value)" /></FormItem>
    </div>
    <FormItem label="词语旋转角" extra="清空后按默认角度集合排列词语。"><InputNumber :value="options.rotation" :min="-90" :max="90" aria-label="词云旋转角" :disabled="unavailable" @update:value="value => numberChanged('rotation', value)" /></FormItem>
    <Button :disabled="unavailable || options.minFontSize === undefined && options.maxFontSize === undefined && options.rotation === undefined" @click="edit({ minFontSize: undefined, maxFontSize: undefined, rotation: undefined })">恢复字号与旋转默认值</Button>
    <Alert v-if="error" type="error" :message="error" show-icon />
  </div>
</template>

<style scoped>
.chart-word-cloud-settings { display: grid; gap: 12px; margin-top: 12px; min-width: 0; }
.chart-word-cloud-sizes { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 12px; }
.chart-word-cloud-sizes :deep(.ant-input-number) { width: 100%; }
.chart-mask-name { display: inline-block; margin-inline-start: 8px; overflow-wrap: anywhere; }
</style>
