<script setup lang="ts">
import { computed, inject, ref } from 'vue'
import { Button, ColorPicker, FormItem, Input, InputNumber, Select, theme } from 'antdv-next'
import { SHEET_RANGE_THEME_SECTIONS, type SheetRangeThemeSection, type SheetRangeThemeStyle } from '../../domain/model.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { normalizeSheetBorderStyle, sheetBorderStyleLabels } from '../../features/sheet-workbench/sheet-cell-border.ts'
import { resetSheetRangeThemeEditorSection, setSheetRangeThemeEditorBorder, setSheetRangeThemeEditorValue,
  sheetRangeThemeEditorBorder, sheetRangeThemeEditorValue, SHEET_RANGE_THEME_BORDER_EDGES, SHEET_RANGE_THEME_EDITOR_FIELDS,
  SHEET_RANGE_THEME_SECTION_LABELS, type SheetRangeThemeBorderEdge } from '../../features/sheet-workbench/sheet-range-theme-editor-state.ts'

const props = defineProps<{ modelValue: SheetRangeThemeStyle }>()
const emit = defineEmits<{ 'update:modelValue': [value: SheetRangeThemeStyle]; error: [message: string] }>()
const tr = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const section = ref<SheetRangeThemeSection>('wholeStyle'), borderEdge = ref<SheetRangeThemeBorderEdge>('top')
type Field = typeof SHEET_RANGE_THEME_EDITOR_FIELDS[number]
const sections = computed(() => SHEET_RANGE_THEME_SECTIONS.map(value => ({ value,
  label: tr(`sheet.range-theme.section.${value}`, SHEET_RANGE_THEME_SECTION_LABELS[value]) })))
const inherit = computed(() => tr('sheet.range-theme.inherit', '继承'))
const toggleOptions = computed(() => [
  { value: 'inherit', label: inherit.value }, { value: 'on', label: tr('sheet.range-theme.on', '开启') },
  { value: 'off', label: tr('sheet.range-theme.off', '关闭') },
])
const colorOptions = computed(() => [
  { value: 'inherit', label: inherit.value }, { value: 'none', label: tr('sheet.range-theme.no-color', '无颜色') },
  { value: 'custom', label: tr('sheet.range-theme.choose-color', '指定颜色') },
])
const borderEdges = computed(() => SHEET_RANGE_THEME_BORDER_EDGES.map(edge => ({ value: edge.value,
  label: tr(`sheet.range-theme.border.${edge.value}`, edge.label) })))
const borderStyles = computed(() => [
  { value: 'inherit', label: inherit.value }, { value: 'none', label: tr('sheet.range-theme.no-border', '无边框') },
  ...Object.entries(sheetBorderStyleLabels).map(([value, label]) => ({ value, label: tr(`borderStyle.${value}`, label) })),
  { value: 'slant-dash-dot', label: tr('sheet.range-theme.slant-dash-dot', '斜点划线') },
])
const currentBorder = computed(() => sheetRangeThemeEditorBorder(props.modelValue, section.value, borderEdge.value))
const borderObject = computed(() => currentBorder.value && typeof currentBorder.value === 'object'
  ? currentBorder.value as Record<string, unknown> : undefined)
const borderStyle = computed(() => {
  if (currentBorder.value === null) return 'none'
  const value = borderObject.value?.style ?? borderObject.value?.s
  if (value === undefined) return 'inherit'
  try { return normalizeSheetBorderStyle(value) } catch { return 'inherit' }
})
const borderColor = computed(() => typeof borderObject.value?.color === 'string' ? borderObject.value.color : undefined)
function label(field: Field): string { return tr(`sheet.range-theme.field.${field.property}`, field.label) }
function value(field: Field): unknown { return sheetRangeThemeEditorValue(props.modelValue, section.value, field.property) }
function textValue(field: Field): string { const current = value(field); return typeof current === 'string' ? current : '' }
function numberValue(field: Field): number | undefined { const current = value(field); return typeof current === 'number' ? current : undefined }
function toggleValue(field: Field): string { const current = value(field); return current === true ? 'on' : current === false ? 'off' : 'inherit' }
function colorMode(field: Field): string { const current = value(field); return current === null ? 'none' : typeof current === 'string' ? 'custom' : 'inherit' }
function options(field: Field) { return (field.options ?? []).map(option => ({ value: option.value,
  label: tr(`sheet.range-theme.option.${field.property}.${option.value}`, option.label) })) }
function update(run: () => SheetRangeThemeStyle): void {
  try { emit('update:modelValue', run()); emit('error', '') }
  catch (error: unknown) { emit('error', error instanceof Error ? error.message : tr('sheet.range-theme.invalid-style', '范围主题样式无效。')) }
}
function setField(field: Field, next: unknown): void { update(() => setSheetRangeThemeEditorValue(props.modelValue, section.value, field.property, next)) }
function setToggle(field: Field, next: unknown): void { setField(field, next === 'on' ? true : next === 'off' ? false : undefined) }
function setColorMode(field: Field, next: unknown): void {
  setField(field, next === 'inherit' ? undefined : next === 'none' ? null : typeof value(field) === 'string' ? value(field) : token.value.colorPrimary)
}
function setBorder(property: 'style' | 'color', next: unknown): void {
  update(() => setSheetRangeThemeEditorBorder(props.modelValue, section.value, borderEdge.value, property,
    property === 'style' && next === 'inherit' ? undefined : next))
}
</script>

<template>
  <div data-sheet-range-theme-style-editor>
    <div :style="{ display: 'flex', alignItems: 'end', gap: `${token.marginSM}px` }">
      <FormItem name="section" :label="tr('sheet.range-theme.section', '样式区段')" :style="{ flex: 1 }">
        <Select v-model:value="section" :options="sections" :aria-label="tr('sheet.range-theme.section', '样式区段')" />
      </FormItem>
      <FormItem><Button @click="update(() => resetSheetRangeThemeEditorSection(modelValue, section))">{{ tr('sheet.range-theme.clear-section', '清空区段') }}</Button></FormItem>
    </div>
    <p :style="{ color: token.colorTextSecondary }">{{ tr('sheet.range-theme.inherit-description', '继承会移除当前区段的该项设置；无颜色、关闭和无边框会覆盖较低区段的对应设置。') }}</p>
    <div :style="{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: `0 ${token.marginSM}px` }">
      <FormItem v-for="field in SHEET_RANGE_THEME_EDITOR_FIELDS" :key="`${section}:${field.property}`" :name="[section, ...field.property.split('.')]" :label="label(field)">
        <Input v-if="field.kind === 'text'" :value="textValue(field)" :aria-label="label(field)" :placeholder="inherit" allow-clear @update:value="setField(field, $event)" />
        <InputNumber v-else-if="field.kind === 'number'" :value="numberValue(field)" :min="field.min" :max="field.max" :aria-label="label(field)" :placeholder="inherit" :style="{ width: '100%' }" @update:value="setField(field, $event)" />
        <Select v-else-if="field.kind === 'toggle'" :value="toggleValue(field)" :options="toggleOptions" :aria-label="label(field)" @update:value="setToggle(field, $event)" />
        <Select v-else-if="field.kind === 'enum'" :value="textValue(field) || undefined" :options="options(field)" :placeholder="inherit" allow-clear :aria-label="label(field)" @update:value="setField(field, $event)" />
        <div v-else :style="{ display: 'flex', gap: `${token.marginXS}px` }">
          <Select :value="colorMode(field)" :options="colorOptions" :aria-label="`${label(field)} ${tr('sheet.range-theme.mode', '模式')}`" :style="{ flex: 1 }" @update:value="setColorMode(field, $event)" />
          <ColorPicker v-if="colorMode(field) === 'custom'" :value="textValue(field)" value-format="rgb" format="hex" show-text allow-clear :aria-label="label(field)" @update:value="setField(field, $event)" @clear="setField(field, null)" />
        </div>
      </FormItem>
    </div>
    <div :style="{ display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: `${token.marginSM}px` }" data-sheet-range-theme-border-fields>
      <FormItem name="borderEdge" :label="tr('sheet.range-theme.border-edge', '边框位置')"><Select v-model:value="borderEdge" :options="borderEdges" :aria-label="tr('sheet.range-theme.border-edge', '边框位置')" /></FormItem>
      <FormItem :name="[section, 'border', borderEdge, 'style']" :label="tr('sheet.range-theme.border-style', '边框线型')"><Select :value="borderStyle" :options="borderStyles" :aria-label="tr('sheet.range-theme.border-style', '边框线型')" @update:value="setBorder('style', $event)" /></FormItem>
      <FormItem :name="[section, 'border', borderEdge, 'color']" :label="tr('sheet.range-theme.border-color', '边框颜色')"><ColorPicker :value="borderColor" value-format="rgb" format="hex" show-text allow-clear :disabled="!borderObject" :aria-label="tr('sheet.range-theme.border-color', '边框颜色')" @update:value="setBorder('color', $event)" @clear="setBorder('color', undefined)" /></FormItem>
    </div>
  </div>
</template>
