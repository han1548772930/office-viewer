<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { Input, Modal, Select } from 'antdv-next'
import { layoutLatexFormula, type LatexFormula, type LatexFormulaMode } from '../../features/document/latex/latex-formula.ts'

const props = defineProps<{ open: boolean; initial?: LatexFormula }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [value: { latex: string; mode: LatexFormulaMode }] }>()
const { t } = useI18n({ useScope: 'global' })
const source = ref('')
const mode = ref<LatexFormulaMode>('inline')
const error = ref('')
const previewCanvas = ref<HTMLCanvasElement | null>(null)
const modeOptions = computed(() => [
  { label: t('document.latex.inline'), value: 'inline' },
  { label: t('document.latex.block'), value: 'block' },
])
const valid = computed(() => {
  try { layoutLatexFormula(source.value); return source.value.trim().length > 0 } catch { return false }
})

watch(() => [props.open, props.initial] as const, ([open, initial]) => {
  if (!open) return
  source.value = initial?.latex ?? ''
  mode.value = initial?.mode ?? 'inline'
  error.value = ''
}, { immediate: true })

watch(() => [source.value, props.open] as const, async ([, open]) => {
  if (!open) return
  await nextTick()
  drawPreview()
}, { immediate: true })

function drawPreview() {
  const canvas = previewCanvas.value
  const context = canvas?.getContext('2d')
  if (!canvas || !context) return
  context.clearRect(0, 0, canvas.width, canvas.height)
  try {
    const formula = layoutLatexFormula(source.value, { fontSize: 20, measure: (text, font) => {
      context.font = font
      return context.measureText(text).width
    } })
    const originX = 16
    const baseline = (canvas.height - formula.height) / 2 + formula.baseline
    context.strokeStyle = '#344054'
    context.fillStyle = '#344054'
    context.lineWidth = 1
    for (const primitive of formula.primitives) {
      if (primitive.kind === 'text') {
        context.font = primitive.font
        context.fillText(primitive.text, originX + primitive.x, baseline + primitive.y)
      } else {
        context.beginPath()
        context.moveTo(originX + primitive.x1, baseline + primitive.y1)
        context.lineTo(originX + primitive.x2, baseline + primitive.y2)
        context.lineWidth = primitive.width
        context.stroke()
      }
    }
  } catch {
    context.fillStyle = '#b42318'
    context.font = '13px sans-serif'
    context.fillText(t('document.latex.invalid'), 16, 30)
  }
}

function apply() {
  try {
    layoutLatexFormula(source.value)
    emit('apply', { latex: source.value.trim(), mode: mode.value })
    emit('update:open', false)
  } catch (cause) {
    error.value = cause instanceof Error ? cause.message : t('document.latex.invalid')
  }
}
</script>

<template>
  <Modal :open="open" :title="t(initial ? 'document.latex.editTitle' : 'document.latex.title')" :ok-text="t('common.apply')" :cancel-text="t('common.cancel')" :width="600" :ok-button-props="{ disabled: !valid }" destroy-on-hidden @ok="apply" @cancel="emit('update:open', false)">
    <div class="latex-panel">
      <label class="latex-panel__mode">
        <span>{{ t('document.latex.mode') }}</span>
        <Select v-model:value="mode" :options="modeOptions" :disabled="!!initial" :aria-label="t('document.latex.mode')" />
      </label>
      <label class="latex-panel__source">
        <span>{{ t('document.latex.source') }}</span>
        <Input.TextArea v-model:value="source" :aria-label="t('document.latex.source')" :auto-size="{ minRows: 3, maxRows: 8 }" spellcheck="false" />
      </label>
      <div class="latex-panel__preview">
        <span>{{ t('document.latex.preview') }}</span>
        <canvas ref="previewCanvas" width="560" height="88" role="img" :aria-label="t('document.latex.preview')" />
      </div>
      <p v-if="error" class="latex-panel__error" role="alert">{{ error }}</p>
    </div>
  </Modal>
</template>

<style scoped>
.latex-panel { display: grid; gap: 14px; }
.latex-panel__mode, .latex-panel__source, .latex-panel__preview { display: grid; gap: 6px; color: #465365; font-size: 12px; }
.latex-panel__mode { grid-template-columns: 120px minmax(0, 1fr); align-items: center; }
.latex-panel__source :deep(textarea) { font: 13px/1.5 "IBM Plex Mono", Consolas, monospace; }
.latex-panel__preview canvas { display: block; width: 100%; height: 88px; border: 1px solid #d9dfe7; background: #fff; }
.latex-panel__error { margin: 0; color: #b42318; font-size: 12px; overflow-wrap: anywhere; }
</style>
