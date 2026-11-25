<script setup lang="ts">
import { nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { theme } from 'antdv-next'
import { useI18n } from 'vue-i18n'
import { layoutLatexFormula } from '../../features/document/latex/latex-layout'

const props = withDefaults(defineProps<{ latex: string; label: string; width?: number; height?: number; fontSize?: number }>(), { width: 560, height: 112, fontSize: 20 })
const { token } = theme.useToken(), { t } = useI18n({ useScope: 'global' })
const canvas = ref<HTMLCanvasElement>()
let generation = 0, disposed = false
function draw() {
  const surface = canvas.value, context = surface?.getContext('2d')
  if (disposed || !surface || !context) return
  const width = props.width, height = props.height, ratio = Math.max(1, window.devicePixelRatio || 1)
  surface.width = Math.round(width * ratio); surface.height = Math.round(height * ratio)
  context.setTransform(ratio, 0, 0, ratio, 0, 0)
  context.clearRect(0, 0, width, height)
  context.fillStyle = token.value.colorBgContainer; context.fillRect(0, 0, width, height)
  try {
    const result = layoutLatexFormula(props.latex, { fontSize: props.fontSize, measure: (text, font) => { context.font = font; return context.measureText(text).width } })
    const margin = Math.min(12, height / 8), scale = Math.min(1, (width - 2 * margin) / Math.max(result.width, 1), (height - 2 * margin) / Math.max(result.height, 1))
    const x = (width / scale - result.width) / 2, baseline = (height / scale - result.height) / 2 + result.baseline
    context.scale(scale, scale); context.fillStyle = token.value.colorText; context.strokeStyle = token.value.colorText
    for (const primitive of result.primitives) {
      if (primitive.kind === 'text') { context.font = primitive.font; context.fillText(primitive.text, x + primitive.x, baseline + primitive.y) }
      else {
        context.beginPath(); context.lineWidth = primitive.width
        context.moveTo(x + primitive.x1, baseline + primitive.y1); context.lineTo(x + primitive.x2, baseline + primitive.y2); context.stroke()
      }
    }
  } catch {
    context.fillStyle = token.value.colorError; context.font = '12px sans-serif'
    context.fillText(t('document.latex.invalid'), 8, Math.min(height - 4, 24), width - 16)
  }
}
watch(() => [props.latex, props.width, props.height, props.fontSize, token.value.colorText, token.value.colorBgContainer, token.value.colorError], async () => {
  const current = ++generation
  await nextTick()
  if (current === generation) draw()
}, { immediate: true })
onMounted(() => {
  draw()
  window.addEventListener('resize', draw)
  if (document.fonts) { document.fonts.addEventListener('loadingdone', draw); void document.fonts.ready.then(() => { if (!disposed) draw() }) }
})
onBeforeUnmount(() => { disposed = true; generation++; window.removeEventListener('resize', draw); document.fonts?.removeEventListener('loadingdone', draw) })
</script>

<template><canvas ref="canvas" class="latex-preview-canvas" :width="width" :height="height" role="img" :aria-label="label" :style="{ aspectRatio: width + '/' + height }" /></template>
<style scoped>.latex-preview-canvas { display: block; width: 100%; height: auto; max-width: 100%; }</style>
