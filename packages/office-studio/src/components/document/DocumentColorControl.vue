<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, onMounted, reactive, ref, useId, watch, type Component, type VNode } from 'vue'
import { Button, ColorPicker, Divider, Form, FormItem, Tooltip } from 'antdv-next'
import { ChevronDown } from 'lucide-vue-next'
import { OfficeDirectionKey, OfficeTranslatorKey } from '../../i18n/office-i18n'
import { DOCUMENT_COLOR_PRESETS, documentColorDraftIsValid, documentColorIsAutomatic, documentColorParams, documentColorPreview, type DocumentColorField, type DocumentColorInputFormat, type DocumentColorParams } from '../../features/document/services/document-colors'
import { useRibbonTheme } from '../../platform/ribbon/ribbon-theme'

const props = withDefaults(defineProps<{
  title: string; field: DocumentColorField; icon: Component; defaultColor: string; lastColor?: string;
  selectedColor?: string; mixed?: boolean; open?: boolean; disabled?: boolean; measurement?: boolean;
  grid?: boolean; overflow?: boolean; iconSize?: number;
}>(), { mixed: false, open: false, disabled: false, measurement: false, grid: false, overflow: false, iconSize: 16 })
const emit = defineEmits<{ insert: [params: DocumentColorParams]; 'update:open': [value: boolean] }>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const direction = inject(OfficeDirectionKey, ref<'ltr' | 'rtl'>('ltr'))
const themeStyle = useRibbonTheme()
const popupId = useId()
const root = ref<HTMLElement>()
const palette = ref<HTMLElement>()
const arrow = ref<HTMLElement>()
const more = ref<HTMLElement>()
const unavailable = computed(() => props.disabled || props.measurement)
const mainColor = computed(() => props.lastColor ?? props.defaultColor)
const selectedPreview = computed(() => documentColorPreview(props.field, props.selectedColor, props.defaultColor))
const automatic = computed(() => !props.mixed && documentColorIsAutomatic(props.field, props.selectedColor))
const automaticTitle = computed(() => props.field === 'highlight' ? translate('document.color.noBackground', '无背景') : translate('document.color.automatic', '自动文字颜色'))
const draft = ref(selectedPreview.value)
const custom = ref(false)
const format = ref<DocumentColorInputFormat>('hex')
const draftForm = reactive({ color: selectedPreview.value })
const draftFields = ref<string[]>([])
const draftTouched = ref(false)
const invalidDraft = computed(() => draftTouched.value && !documentColorDraftIsValid(format.value, draftFields.value))
const optionsTitle = computed(() => `${props.title}${translate('common.options', '选项')}`)
const valueTitle = computed(() => translate('document.color.value', '颜色值'))
const formatTitle = computed(() => translate('document.color.format', '颜色格式'))
const inputId = `${popupId}-value`
const PanelNode = ({ node }: { node: VNode }) => node

watch(() => [props.open, selectedPreview.value, props.mixed] as const, async ([open], previous) => {
  resetDraft()
  if (!open || !previous[0]) { custom.value = false; format.value = 'hex' }
  if (open && custom.value) { await nextTick(); refreshDraftFields() }
})
watch(format, async () => { draftTouched.value = false; await nextTick(); refreshDraftFields() })
watch(unavailable, value => { if (value) setOpen(false) })

function resetDraft() {
  draft.value = selectedPreview.value
  draftForm.color = selectedPreview.value
  draftFields.value = []
  draftTouched.value = false
}
function setOpen(value: boolean) {
  if (value && unavailable.value) return
  if (!value) custom.value = false
  emit('update:open', value)
}
function apply(params: DocumentColorParams | undefined) {
  if (!params || unavailable.value) return
  setOpen(false)
  emit('insert', params)
}
function presetActive(color: string): boolean {
  return !props.mixed && !automatic.value && documentColorParams(props.field, selectedPreview.value)?.patch[props.field] === color
}
function draftInputs(): HTMLInputElement[] {
  return Array.from(palette.value?.querySelectorAll<HTMLInputElement>('.document-color-custom input:not([role="combobox"])') ?? [])
}
function refreshDraftFields() {
  if (!custom.value) return
  const inputs = draftInputs()
  const labels = format.value === 'hex' ? ['HEX'] : format.value === 'rgb' ? ['R', 'G', 'B'] : ['H', 'S', 'B']
  inputs.forEach((input, index) => {
    if (index === 0) input.id = inputId
    input.setAttribute('aria-label', labels[index] ?? valueTitle.value)
  })
  palette.value?.querySelector<HTMLInputElement>('[role="combobox"]')?.setAttribute('aria-label', formatTitle.value)
  draftFields.value = inputs.map(input => input.value)
  if (format.value === 'hex') {
    const value = draftFields.value[0]?.trim() ?? ''
    const params = documentColorParams(props.field, value.startsWith('#') ? value : `#${value}`)
    const color = params?.patch[props.field]
    if (typeof color === 'string') draftForm.color = color
  }
}
function updateDraft(value: unknown) {
  if (typeof value !== 'string') return
  draft.value = value
  draftForm.color = value
  void nextTick(refreshDraftFields)
}
function onDraftInput(event: Event) {
  if (!view || !(event.target instanceof view.HTMLInputElement) || event.target.getAttribute('role') === 'combobox') return
  draftTouched.value = true
  void nextTick(refreshDraftFields)
}
async function openCustom() {
  if (unavailable.value) return
  resetDraft()
  format.value = 'hex'
  custom.value = true
  await nextTick()
  refreshDraftFields()
  draftInputs()[0]?.focus({ preventScroll: true })
}
function confirmCustom() {
  if (unavailable.value || !custom.value) return
  refreshDraftFields()
  draftTouched.value = true
  if (invalidDraft.value) { draftInputs()[0]?.focus({ preventScroll: true }); return }
  apply(documentColorParams(props.field, draftForm.color))
}
async function cancelCustom() {
  resetDraft()
  custom.value = false
  await nextTick()
  more.value?.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true })
}
async function openPalette(last = false) {
  if (unavailable.value) return
  setOpen(true)
  await nextTick()
  const controls = palette.value?.querySelectorAll<HTMLElement>('button:not(:disabled), input:not(:disabled), [tabindex="0"]')
  ;(last ? controls?.[controls.length - 1] : controls?.[0])?.focus({ preventScroll: true })
}
function closePalette() {
  setOpen(false)
  arrow.value?.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true })
}
function onPanelKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return
  if (custom.value) event.stopPropagation()
  if (event.key === 'Escape') {
    event.preventDefault(); event.stopPropagation()
    if (custom.value) void cancelCustom()
    else closePalette()
  }
}
function onTriggerKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return
  if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp' && event.key !== 'Escape') return
  event.preventDefault(); event.stopPropagation()
  if (event.key === 'Escape') closePalette()
  else void openPalette(event.key === 'ArrowUp')
}

// Handle custom-stage keys before document-level popup dismissal; keep IME defaults intact.
let view: (Window & typeof globalThis) | null = null
function onWindowKeydown(event: KeyboardEvent) {
  if (!props.open || !view || !(event.target instanceof view.HTMLElement) || !palette.value?.contains(event.target)) return
  const input = event.target instanceof view.HTMLInputElement
  if (event.isComposing || event.keyCode === 229) {
    if (input && (event.key === 'Enter' || event.key === 'Escape')) event.stopPropagation()
    return
  }
  if (event.key === 'Escape') {
    event.preventDefault(); event.stopPropagation()
    if (custom.value) void cancelCustom()
    else closePalette()
  } else if (custom.value && input && event.target.getAttribute('role') !== 'combobox' && event.key === 'Enter') {
    event.preventDefault(); event.stopPropagation(); confirmCustom()
  }
}
onMounted(() => { view = root.value?.ownerDocument.defaultView ?? null; view?.addEventListener('keydown', onWindowKeydown, true) })
onBeforeUnmount(() => { view?.removeEventListener('keydown', onWindowKeydown, true) })
</script>

<template>
  <span ref="root" class="document-color-control" :class="{ 'is-grid': grid, 'is-overflow': overflow }" :style="themeStyle">
    <Tooltip :title="measurement ? undefined : title" :open="open || measurement ? false : undefined" :mouse-enter-delay="0.6">
      <Button type="text" size="small" class="document-color-main" :disabled="unavailable" :aria-label="title" @mousedown.prevent @click="apply({ patch: { [field]: mainColor } })">
        <span class="document-color-icon" aria-hidden="true"><component :is="icon" :size="iconSize" /><span class="document-color-bar" :style="{ backgroundColor: documentColorPreview(field, mainColor, defaultColor) }" /></span>
        <span v-if="overflow">{{ title }}</span>
      </Button>
    </Tooltip>
    <ColorPicker :value="draft" :open="open" value-format="hex" v-model:format="format" :disabled-alpha="true" :disabled="unavailable" :placement="direction === 'rtl' ? 'bottomRight' : 'bottomLeft'" :arrow="false" :destroy-on-hidden="true" @update:value="updateDraft" @update:open="setOpen">
      <span ref="arrow" class="document-color-arrow">
        <Tooltip :title="measurement ? undefined : optionsTitle" :open="open || measurement ? false : undefined" :mouse-enter-delay="0.6">
          <Button type="text" size="small" :disabled="unavailable" :aria-label="optionsTitle" aria-haspopup="dialog" :aria-expanded="open" :aria-controls="popupId" @mousedown.prevent @keydown="onTriggerKeydown"><ChevronDown :size="12" /></Button>
        </Tooltip>
      </span>
      <template #panelRender="{ panel }">
        <div :id="popupId" ref="palette" class="document-color-panel" role="dialog" :aria-label="optionsTitle" :dir="direction" :style="themeStyle" @keydown="onPanelKeydown">
          <span v-if="mixed" class="document-color-state">{{ translate('common.mixed', '混合') }}</span>
          <span v-else-if="automatic" class="document-color-state">{{ automaticTitle }}</span>
          <Form v-if="custom" :model="draftForm" layout="vertical" class="document-color-custom" @input.capture="onDraftInput" @finish="confirmCustom">
            <FormItem name="color" :label="valueTitle" :html-for="inputId" :validate-status="invalidDraft ? 'error' : undefined" :help="invalidDraft ? translate('document.color.invalid', '请输入有效的颜色值') : undefined">
              <PanelNode :node="panel" />
            </FormItem>
            <div class="document-color-actions is-custom">
              <Button size="small" :disabled="unavailable" @click="cancelCustom">{{ translate('document.color.cancel', '取消') }}</Button>
              <Button type="primary" size="small" :disabled="unavailable" @click="confirmCustom">{{ translate('document.color.confirm', '确定') }}</Button>
            </div>
          </Form>
          <template v-else>
            <div class="document-color-presets" role="group" :aria-label="translate('document.color.standard', '标准颜色')">
              <div v-for="(row, rowIndex) in DOCUMENT_COLOR_PRESETS" :key="rowIndex" class="document-color-preset-row">
                <Button v-for="color in row" :key="color" type="text" size="small" class="document-color-preset" :disabled="unavailable" :aria-label="color.toUpperCase()" :aria-pressed="presetActive(color)" :title="color.toUpperCase()" @click="apply(documentColorParams(field, color))">
                  <span class="document-color-swatch" :style="{ backgroundColor: color }" aria-hidden="true" />
                </Button>
              </div>
            </div>
            <span ref="more" class="document-color-more">
              <Button type="text" size="small" :disabled="unavailable" @click="openCustom">
                {{ translate('document.color.more', '更多颜色') }}
                <span class="document-color-swatch" :style="{ backgroundColor: selectedPreview }" aria-hidden="true" />
              </Button>
            </span>
            <Divider style="margin: 8px 0" />
            <div class="document-color-actions">
              <Button size="small" :disabled="unavailable" @click="apply(documentColorParams(field, defaultColor))">{{ translate('document.color.reset', '重置颜色') }}</Button>
              <Button size="small" :disabled="unavailable" :aria-pressed="automatic" @click="apply(documentColorParams(field, null))">{{ automaticTitle }}</Button>
            </div>
          </template>
        </div>
      </template>
    </ColorPicker>
  </span>
</template>

<style scoped>
.document-color-control { display: inline-flex; align-items: center; height: var(--ribbon-selector-height, 28px); }
.document-color-control.is-grid { width: 100%; height: 100%; }
.document-color-main { display: inline-flex; align-items: center; justify-content: center; min-width: 24px; height: 24px; padding: 0 4px; color: var(--ribbon-text); }
.document-color-control.is-grid .document-color-main { flex: 1 1 auto; min-width: 20px; height: 100%; }
.document-color-icon { position: relative; display: inline-flex; flex: 0 0 auto; }
.document-color-bar { position: absolute; inset-inline: 0; bottom: -2px; height: 3px; border-radius: 1px; }
.document-color-main:disabled .document-color-bar { opacity: 0.35; }
.document-color-arrow { display: inline-flex; height: 28px; }
.document-color-arrow :deep(button) { min-width: 16px; width: 16px; height: 28px; padding: 0; border-inline-start: 1px solid var(--ribbon-border); border-start-start-radius: 0; border-end-start-radius: 0; color: var(--ribbon-text); }
.document-color-control.is-grid .document-color-arrow, .document-color-control.is-grid .document-color-arrow :deep(button) { height: 100%; border: 0; }
.document-color-control.is-overflow { width: 100%; }
.document-color-control.is-overflow .document-color-main { justify-content: flex-start; gap: 8px; flex: 1 1 auto; }
.document-color-panel { box-sizing: border-box; width: 288px; max-width: calc(100vw - 32px); max-height: 70vh; overflow: auto; color: var(--ribbon-text); }
.document-color-actions { display: flex; gap: 8px; }
.document-color-actions.is-custom { justify-content: flex-end; }
.document-color-custom :deep(.ant-form-item) { margin-bottom: 12px; }
.document-color-preset-row { display: grid; grid-template-columns: repeat(9, minmax(0, 1fr)); }
.document-color-preset { display: flex; align-items: center; justify-content: center; width: 32px; max-width: 100%; height: 32px; padding: 4px; }
.document-color-swatch { box-sizing: border-box; display: inline-block; width: 24px; height: 24px; border: 1px solid var(--ribbon-border); border-radius: 4px; }
.document-color-preset[aria-pressed="true"] { background: var(--ribbon-active-background); }
.document-color-preset[aria-pressed="true"] .document-color-swatch { outline: 2px solid var(--ribbon-primary); outline-offset: 1px; }
.document-color-more { display: block; margin-top: 8px; }
.document-color-more :deep(button) { display: flex; justify-content: space-between; width: 100%; height: 32px; padding-inline: 8px; }
.document-color-more .document-color-swatch { width: 20px; height: 20px; }
.document-color-state { display: block; margin-bottom: 8px; color: var(--ribbon-text-secondary); }
</style>
