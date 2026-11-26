<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, onMounted, reactive, ref, useId, watch, type Component, type VNode } from 'vue'
import { Button, ColorPicker, Divider, Form, FormItem, Tooltip, theme } from 'antdv-next'
import { ChevronDown } from 'lucide-vue-next'
import { OfficeDirectionKey, OfficeTranslatorKey } from '../../i18n/office-i18n'
import { SHEET_COLOR_PRESETS, sheetColorDraftValue, sheetColorHex, sheetOpaqueColorHex, type SheetColorField, type SheetColorInputFormat } from '../../features/sheet-workbench/sheet-colors'
import { useRibbonTheme } from '../../platform/ribbon/ribbon-theme'
import { RibbonPopupScopeKey, RibbonEditorFocusScopeKey } from '../../platform/ribbon/ribbon-popup-scope'

const props = withDefaults(defineProps<{
  title: string; field: SheetColorField | 'border'; commandId: string; icon?: Component;
  defaultColor?: string; lastColor?: string; selectedColor?: string | null; mixed?: boolean;
  resetCommandId?: string; resetTitle?: string; open?: boolean; disabled?: boolean;
  measurement?: boolean; grid?: boolean; overflow?: boolean; iconSize?: number; split?: boolean;
}>(), { mixed: false, disabled: false, measurement: false, grid: false, overflow: false, iconSize: 16, split: true })
const emit = defineEmits<{ execute: [commandId: string, params?: string]; 'update:open': [value: boolean] }>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const direction = inject(OfficeDirectionKey, ref<'ltr' | 'rtl'>('ltr'))
const themeStyle = useRibbonTheme()
const { token } = theme.useToken()
const defaultColor = computed(() => sheetOpaqueColorHex(props.defaultColor ?? (props.field === 'background' ? token.value.colorPrimary : token.value.colorText)) ?? '#000000')
const mainColor = computed(() => sheetColorHex(props.lastColor) ?? defaultColor.value)
const selectedPreview = computed(() => sheetColorHex(props.selectedColor) ?? props.selectedColor ?? defaultColor.value)
const automatic = computed(() => !props.mixed && props.selectedColor == null)
const automaticTitle = computed(() => props.field === 'background' ? translate('sheet.color.no-fill', '无填充')
  : props.field === 'color' ? translate('sheet.color.automatic', '自动文字颜色') : translate('sheet.color.unset-border', '未设置边框颜色'))
const optionsTitle = computed(() => props.split ? translate('sheet.color.options', '{title}选项').replace('{title}', props.title) : props.title)
const popupId = useId(), inputId = `${popupId}-value`
const popupScope = inject(RibbonPopupScopeKey, undefined)
const editorFocusScope = inject(RibbonEditorFocusScopeKey, undefined)
const root = ref<HTMLElement>(), palette = ref<HTMLElement>(), trigger = ref<HTMLElement>(), more = ref<HTMLElement>()
const localOpen = ref(false), custom = ref(false), composing = ref(false)
const unavailable = computed(() => props.disabled || props.measurement)
const popupOpen = computed(() => props.open ?? (popupScope ? popupScope.activeId.value === popupId : localOpen.value))
const draft = ref(selectedPreview.value), format = ref<SheetColorInputFormat>('hex')
const draftForm = reactive({ color: selectedPreview.value })
const draftFields = ref<string[]>([]), draftTouched = ref(false)
const invalidDraft = computed(() => draftTouched.value && sheetColorDraftValue(format.value, draftFields.value) === undefined)
const valueTitle = computed(() => translate('sheet.color.value', '颜色值'))
const formatTitle = computed(() => translate('sheet.color.format', '颜色格式'))
const PanelNode = ({ node }: { node: VNode }) => node
// Nested palettes stay within the owning popup's outside-click and disposal boundary.
const popupContainer = (triggerNode: HTMLElement) => props.open === undefined && popupScope
  ? root.value ?? triggerNode.parentElement ?? triggerNode.ownerDocument.body : triggerNode.ownerDocument.body

function resetDraft() {
  draft.value = selectedPreview.value
  draftForm.color = selectedPreview.value
  draftFields.value = []; draftTouched.value = false
}
function setOpen(value: boolean) {
  if (value && unavailable.value) return
  localOpen.value = value
  if (props.open === undefined && popupScope) {
    if (value) popupScope.activeId.value = popupId
    else if (popupScope.activeId.value === popupId) popupScope.close()
  }
  if (!value) { custom.value = false; composing.value = false }
  emit('update:open', value)
}
watch(() => [popupOpen.value, selectedPreview.value, props.mixed] as const, async ([open], previous) => {
  resetDraft()
  if (!open || !previous[0]) { custom.value = false; composing.value = false; format.value = 'hex' }
  if (open && custom.value) { await nextTick(); refreshDraftFields() }
})
watch(format, async () => { draftTouched.value = false; await nextTick(); refreshDraftFields() })
watch(unavailable, value => { if (value) setOpen(false) })

function apply(value: unknown) {
  const color = sheetColorHex(value)
  if (!color || unavailable.value || composing.value) return
  closePalette(); emit('execute', props.commandId, color)
}
function resetColor() {
  if (!props.resetCommandId || unavailable.value || composing.value) return
  closePalette(); emit('execute', props.resetCommandId)
}
function presetActive(color: string): boolean {
  return !props.mixed && !automatic.value && sheetColorHex(props.selectedColor) === color
}
function draftInputs(): HTMLInputElement[] {
  return Array.from(palette.value?.querySelectorAll<HTMLInputElement>('.sheet-color-custom input:not([role="combobox"])') ?? [])
}
function refreshDraftFields(force = false) {
  if (!custom.value) return
  // ColorPicker 1.8.2 exposes no raw channel-draft or input refs. These are its own
  // Input/InputNumber controls; the narrow query supplies labels and validates partial edits.
  const inputs = draftInputs(), labels = format.value === 'hex' ? ['HEX'] : format.value === 'rgb' ? ['R', 'G', 'B'] : ['H', 'S', 'B']
  inputs.forEach((input, index) => {
    if (index === 0) input.id = inputId
    input.setAttribute('aria-label', labels[index] ?? valueTitle.value)
  })
  palette.value?.querySelector<HTMLInputElement>('[role="combobox"]')?.setAttribute('aria-label', formatTitle.value)
  if (force || !draftTouched.value) draftFields.value = inputs.map(input => input.value)
  const color = sheetColorDraftValue(format.value, draftFields.value)
  if (color) draftForm.color = color
}
function updateDraft(value: unknown) {
  if (typeof value !== 'string') return
  draft.value = value; draftForm.color = value
  void nextTick(refreshDraftFields)
}
function onDraftInput(event: Event) {
  if (!view || !(event.target instanceof view.HTMLInputElement) || event.target.getAttribute('role') === 'combobox') return
  draftTouched.value = true
  refreshDraftFields(true)
}
function onCustomPointerDown(event: PointerEvent) {
  if (!view || !(event.target instanceof view.Element)) return
  if (event.target.closest('.ant-color-picker-select, .ant-color-picker-slider, [role="slider"]')) {
    draftTouched.value = false
    void nextTick(refreshDraftFields)
  }
}
function onCompositionEnd() { composing.value = false; void nextTick(() => refreshDraftFields(true)) }
async function openCustom() {
  if (unavailable.value) return
  resetDraft(); format.value = 'hex'; custom.value = true
  await nextTick(); refreshDraftFields(); draftInputs()[0]?.focus({ preventScroll: true })
}
function confirmCustom() {
  if (unavailable.value || composing.value || !custom.value) return
  refreshDraftFields(); draftTouched.value = true
  const color = sheetColorDraftValue(format.value, draftFields.value)
  if (!color) { draftInputs()[0]?.focus({ preventScroll: true }); return }
  apply(color)
}
async function cancelCustom() {
  if (composing.value) return
  resetDraft(); custom.value = false
  await nextTick(); more.value?.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true })
}
async function openPalette(last = false) {
  if (unavailable.value) return
  setOpen(true); await nextTick()
  const controls = palette.value?.querySelectorAll<HTMLElement>('button:not(:disabled), input:not(:disabled), [tabindex="0"]')
  ;(last ? controls?.[controls.length - 1] : controls?.[0])?.focus({ preventScroll: true })
}
function closePalette() {
  setOpen(false); trigger.value?.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true })
}
function onPanelKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || composing.value) return
  if (custom.value) event.stopPropagation()
  if (event.key === 'Escape') {
    event.preventDefault(); event.stopPropagation()
    if (custom.value) void cancelCustom()
    else closePalette()
  }
}
function onTriggerKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || composing.value) return
  if (!['ArrowDown', 'ArrowUp', 'Escape'].includes(event.key)) return
  event.preventDefault(); event.stopPropagation()
  if (event.key === 'Escape') closePalette()
  else void openPalette(event.key === 'ArrowUp')
}
let view: (Window & typeof globalThis) | null = null
function onWindowKeydown(event: KeyboardEvent) {
  if (!popupOpen.value || !view || !(event.target instanceof view.HTMLElement) || !palette.value?.contains(event.target)) return
  const input = event.target instanceof view.HTMLInputElement
  if (event.isComposing || event.keyCode === 229 || composing.value) {
    if (input && (event.key === 'Enter' || event.key === 'Escape')) event.stopPropagation()
    return
  }
  if (custom.value && event.target.closest('[role="slider"]') && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'PageUp', 'PageDown'].includes(event.key)) {
    draftTouched.value = false
    void nextTick(refreshDraftFields)
  }
  // Window capture precedes the document popup-dismiss listener, including nested border panels.
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
  <span ref="root" class="sheet-color-control" :class="{ 'is-grid': grid, 'is-overflow': overflow, 'is-select': !split }" :style="themeStyle">
    <Tooltip v-if="split" :title="measurement ? undefined : title" :open="popupOpen || measurement ? false : undefined" :mouse-enter-delay="0.6">
      <Button type="text" size="small" class="sheet-color-main" :styles="{ root: { border: 0 } }" :disabled="unavailable || composing" :aria-label="title" @mousedown.prevent @click="apply(mainColor)">
        <span class="sheet-color-icon" aria-hidden="true"><component :is="icon" v-if="icon" :size="iconSize" /><span class="sheet-color-bar" :style="{ backgroundColor: mainColor }" /></span>
        <span v-if="overflow" class="sheet-color-main-label">{{ title }}</span>
      </Button>
    </Tooltip>
    <ColorPicker :value="draft" :open="popupOpen" value-format="hex" v-model:format="format" :disabled-alpha="true" :disabled="unavailable" :placement="direction === 'rtl' ? 'bottomRight' : 'bottomLeft'" :get-popup-container="popupContainer" :arrow="false" :destroy-on-hidden="true" @update:value="updateDraft" @update:open="setOpen">
      <span ref="trigger" class="sheet-color-arrow">
        <Tooltip :title="measurement ? undefined : optionsTitle" :open="popupOpen || measurement ? false : undefined" :mouse-enter-delay="0.6">
          <Button type="text" size="small" :styles="{ root: { border: 0 } }" :disabled="unavailable || composing" :aria-label="optionsTitle" aria-haspopup="dialog" :aria-expanded="popupOpen" :aria-controls="popupId" @mousedown.prevent @keydown="onTriggerKeydown">
            <span v-if="!split" class="sheet-color-swatch" :style="{ backgroundColor: mainColor }" aria-hidden="true" /><ChevronDown :size="12" />
          </Button>
        </Tooltip>
      </span>
      <template #panelRender="{ panel }">
        <div :id="popupId" ref="palette" class="sheet-color-panel" :data-ribbon-editor-owner="editorFocusScope?.ownerId" role="dialog" :aria-label="optionsTitle" :dir="direction" :style="themeStyle" @keydown="onPanelKeydown" @compositionstart.capture="composing = true" @compositionend.capture="onCompositionEnd">
          <span v-if="mixed" class="sheet-color-state">{{ translate('sheet.color.mixed', '混合') }}</span>
          <span v-else-if="automatic" class="sheet-color-state">{{ automaticTitle }}</span>
          <Form v-if="custom" :model="draftForm" layout="vertical" class="sheet-color-custom" @input.capture="onDraftInput" @pointerdown.capture="onCustomPointerDown" @finish="confirmCustom">
            <FormItem name="color" :label="valueTitle" :html-for="inputId" :validate-status="invalidDraft ? 'error' : undefined" :help="invalidDraft ? translate('sheet.color.invalid', '请输入有效的颜色值') : undefined">
              <PanelNode :node="panel" />
            </FormItem>
            <div class="sheet-color-actions is-custom">
              <Button size="small" :disabled="unavailable || composing" @mousedown.prevent @click="cancelCustom">{{ translate('sheet.color.cancel', '取消') }}</Button>
              <Button type="primary" size="small" :disabled="unavailable || composing" @mousedown.prevent @click="confirmCustom">{{ translate('sheet.color.confirm', '确定') }}</Button>
            </div>
          </Form>
          <template v-else>
            <div class="sheet-color-presets" role="group" :aria-label="translate('sheet.color.standard', '标准颜色')">
              <div v-for="(row, rowIndex) in SHEET_COLOR_PRESETS" :key="rowIndex" class="sheet-color-preset-row">
                <Tooltip v-for="color in row" :key="color" :title="color.toUpperCase()" :trigger="['hover', 'focus']">
                  <Button type="text" size="small" class="sheet-color-preset" :disabled="unavailable" :aria-label="color.toUpperCase()" :aria-pressed="presetActive(color)" @mousedown.prevent @click="apply(color)">
                    <span class="sheet-color-swatch" :style="{ backgroundColor: color }" aria-hidden="true" />
                  </Button>
                </Tooltip>
              </div>
            </div>
            <span ref="more" class="sheet-color-more"><Button type="text" size="small" :disabled="unavailable" @click="openCustom">
              {{ translate('sheet.color.more', '更多颜色') }}
              <span v-if="!mixed && !automatic" class="sheet-color-swatch" :style="{ backgroundColor: selectedPreview }" aria-hidden="true" />
            </Button></span>
            <template v-if="resetCommandId">
              <Divider style="margin: 8px 0" />
              <Button size="small" class="sheet-color-reset" :disabled="unavailable" :aria-pressed="automatic" @mousedown.prevent @click="resetColor">{{ resetTitle ?? automaticTitle }}</Button>
            </template>
          </template>
        </div>
      </template>
    </ColorPicker>
  </span>
</template>

<style scoped>
.sheet-color-control { display: inline-flex; box-sizing: border-box; min-width: calc(var(--ribbon-icon-size, 16px) + 24px); align-items: center; height: var(--ribbon-selector-height, 28px); }
.sheet-color-control.is-grid { width: 100%; min-width: calc(var(--ribbon-icon-size, 16px) + 20px); height: 100%; }
.sheet-color-main { display: inline-flex; box-sizing: border-box; align-items: center; justify-content: center; min-width: calc(var(--ribbon-icon-size, 16px) + 8px); height: 100%; padding: 0 4px; color: var(--ribbon-text); border: 0; }
.sheet-color-control.is-grid .sheet-color-main { flex: 1 1 auto; min-width: calc(var(--ribbon-icon-size, 16px) + 4px); height: 100%; padding-inline: 2px; }
.sheet-color-icon { position: relative; display: inline-flex; flex: 0 0 auto; align-items: center; justify-content: center; }
.sheet-color-icon :deep(svg) { display: block; }
.sheet-color-bar { position: absolute; inset-inline: 0; bottom: -2px; height: 3px; border-radius: 1px; }
.sheet-color-main:disabled .sheet-color-bar { opacity: 0.35; }
.sheet-color-arrow { display: inline-flex; flex: 0 0 16px; height: 100%; }
.sheet-color-arrow :deep(button) { position: relative; display: inline-flex; box-sizing: border-box; align-items: center; justify-content: center; min-width: 16px; width: 16px; height: 100%; padding: 0; border: 0; border-start-start-radius: 0; border-end-start-radius: 0; color: var(--ribbon-text); }
.sheet-color-arrow :deep(button)::before { position: absolute; inset-block: 0; inset-inline-start: 0; width: 1px; background: var(--ribbon-border); content: ''; pointer-events: none; }
.sheet-color-main:disabled, .sheet-color-arrow :deep(button:disabled) { color: var(--ribbon-disabled); }
.sheet-color-arrow :deep(svg) { display: block; width: 12px; height: 12px; flex: 0 0 12px; }
.sheet-color-control.is-grid .sheet-color-arrow, .sheet-color-control.is-grid .sheet-color-arrow :deep(button) { height: 100%; border: 0; }
.sheet-color-control.is-grid .sheet-color-arrow :deep(button)::before, .sheet-color-control.is-select .sheet-color-arrow :deep(button)::before { display: none; }
.sheet-color-control.is-overflow { width: 100%; min-width: 0; height: 28px; }
.sheet-color-control.is-overflow .sheet-color-main { justify-content: flex-start; gap: 8px; flex: 1 1 auto; }
.sheet-color-main-label { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sheet-color-control.is-select { width: 100%; }
.sheet-color-control.is-select .sheet-color-arrow { width: 100%; flex: 1 1 auto; }
.sheet-color-control.is-select .sheet-color-arrow :deep(button) { display: flex; justify-content: space-between; gap: 8px; width: 100%; padding: 0 4px; border: 0; }
.sheet-color-control.is-select .sheet-color-swatch { width: 20px; height: 20px; }
.sheet-color-main:not(:disabled):hover, .sheet-color-arrow :deep(button:not(:disabled):hover) { background: var(--ribbon-hover-background); }
.sheet-color-panel { box-sizing: border-box; width: 288px; max-width: calc(100vw - 32px); max-height: 70vh; overflow: auto; color: var(--ribbon-text); }
.sheet-color-actions { display: flex; justify-content: flex-end; gap: 8px; }
.sheet-color-custom :deep(.ant-form-item) { margin-bottom: 12px; }
.sheet-color-preset-row { display: grid; grid-template-columns: repeat(9, minmax(0, 1fr)); }
.sheet-color-preset { display: flex; box-sizing: border-box; align-items: center; justify-content: center; width: 32px; max-width: 100%; height: 32px; padding: 3px; }
.sheet-color-swatch { box-sizing: border-box; display: inline-block; width: 24px; height: 24px; border: 1px solid var(--ribbon-border); border-radius: 4px; }
.sheet-color-preset .sheet-color-swatch { max-width: 100%; height: auto; aspect-ratio: 1; }
.sheet-color-preset[aria-pressed="true"] { background: var(--ribbon-active-background); }
.sheet-color-preset[aria-pressed="true"] .sheet-color-swatch { outline: 2px solid var(--ribbon-primary); outline-offset: 1px; }
.sheet-color-more { display: block; margin-top: 8px; }
.sheet-color-more :deep(button) { display: flex; justify-content: space-between; width: 100%; height: 32px; padding-inline: 8px; }
.sheet-color-more .sheet-color-swatch { width: 20px; height: 20px; }
.sheet-color-state { display: block; margin-bottom: 8px; color: var(--ribbon-text-secondary); }
</style>
