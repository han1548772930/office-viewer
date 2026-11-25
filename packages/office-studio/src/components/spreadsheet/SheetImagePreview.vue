<script setup lang="ts">
import { nextTick, onMounted, onUnmounted, ref, watch } from 'vue'
import { Alert, theme } from 'antdv-next'
import type { SheetImageResource, SheetImageSourceRect } from '../../domain/sheet-images.ts'
import { decodeSheetImageSource, prepareSheetImageResources } from '../../features/sheet-images/sheet-image-io.ts'
import { sheetImageNormalizedCrop, sheetImageRotatedBounds } from '../../features/sheet-images/sheet-image-placement.ts'
import { createRenderSurface } from '../../rendering/createRenderSurface.ts'
import type { RenderSurface } from '../../rendering/contracts.ts'
const props = defineProps<{ resource: SheetImageResource; frame?: { width: number; height: number }; crop?: SheetImageSourceRect;
  rotation?: number; flipX?: boolean; flipY?: boolean; opacity?: number }>()
const host = ref<HTMLDivElement>(), error = ref(''), { token } = theme.useToken()
let controller: AbortController | undefined
let surface: RenderSurface | undefined, decoded: HTMLImageElement | undefined, disposed = false
function setPreviewSemantics(): void {
  const canvas = surface?.canvas ?? host.value?.querySelector('canvas[data-render-surface]')
  if (!canvas) return
  canvas.removeAttribute('aria-hidden'); canvas.setAttribute('role', 'img'); canvas.setAttribute('aria-label', props.resource.name)
}
function paint(): void {
  if (!surface || disposed) return
  const size = props.frame ?? { width: props.resource.width, height: props.resource.height },
    bounds = sheetImageRotatedBounds({ left: 0, top: 0, ...size }, props.rotation), scale = Math.min(240 / bounds.width, 160 / bounds.height, 1)
  if (![size.width, size.height, scale].every(Number.isFinite) || size.width <= 0 || size.height <= 0) return
  const width = Math.max(1, bounds.width * scale), height = Math.max(1, bounds.height * scale)
  if (host.value) { host.value.style.width = `${width}px`; host.value.style.height = `${height}px` }
  surface.resize(width, height, globalThis.devicePixelRatio || 1)
  setPreviewSemantics()
  surface.render({ width, height, background: token.value.colorFillAlter, commands: decoded ? [{ kind: 'image', source: decoded,
    x: (width - size.width * scale) / 2, y: (height - size.height * scale) / 2, width: size.width * scale, height: size.height * scale,
    sourceRect: sheetImageNormalizedCrop(props.crop, size), rotation: props.rotation, flipX: props.flipX, flipY: props.flipY, opacity: props.opacity }] : [] })
}
onMounted(async () => {
  if (!host.value) return
  try {
    const created = await createRenderSurface(host.value, { onBackendChange: setPreviewSemantics })
    if (disposed) { created.dispose(); return }
    surface = created; paint()
  } catch (reason: unknown) { if (!disposed) error.value = reason instanceof Error ? reason.message : '无法创建图片预览。' }
})
watch(() => props.resource, async resource => {
  controller?.abort(); const task = controller = new AbortController(); error.value = ''
  if (decoded) decoded.src = ''; decoded = undefined; paint()
  let image: HTMLImageElement | undefined
  try {
    const actual = (await prepareSheetImageResources([resource], { signal: task.signal }))[0]!
    image = await decodeSheetImageSource(actual.source, task.signal)
    await nextTick()
    if (task.signal.aborted || controller !== task || disposed) return
    decoded = image; image = undefined; paint()
  } catch (reason: unknown) { if (!task.signal.aborted && controller === task) error.value = reason instanceof Error ? reason.message : '无法显示图片预览。' }
  finally { if (image) image.src = '' }
}, { immediate: true })
watch([() => props.frame, () => props.crop, () => props.rotation, () => props.flipX, () => props.flipY, () => props.opacity, token], paint, { deep: true })
onUnmounted(() => { disposed = true; controller?.abort(); if (decoded) decoded.src = ''; decoded = undefined; surface?.dispose(); surface = undefined })
</script>
<template>
  <div class="sheet-image-preview" :style="{ background: token.colorFillAlter, borderColor: token.colorBorderSecondary }">
    <div ref="host" class="sheet-image-preview-surface" />
  </div>
  <Alert v-if="error" type="warning" :title="error" show-icon />
</template>
<style scoped>
.sheet-image-preview { display: flex; min-height: 80px; max-height: 180px; align-items: center; justify-content: center; border: 1px solid; padding: 8px; margin-bottom: 12px; }
.sheet-image-preview-surface { position: relative; flex: none; max-width: 100%; }
.sheet-image-preview :deep(canvas) { max-width: 100%; max-height: 160px; }
</style>
