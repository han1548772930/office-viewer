<script setup lang="ts">
import { computed, inject, shallowRef } from 'vue'
import { Checkbox, theme } from 'antdv-next'
import type { SheetCheckboxCellPosition } from '../../rendering/spreadsheet/validation-checkbox-layout.ts'
import type { SheetCheckboxCellEntry } from './use-sheet-validation-checkbox.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ position: SheetCheckboxCellPosition; entry: SheetCheckboxCellEntry; selectedKey: string }>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const pressedEntry = shallowRef<SheetCheckboxCellEntry>()
const clipStyle = computed(() => {
  const { clip, clips } = props.position
  const path = clips.map(part => `M${part.x - clip.x} ${part.y - clip.y}h${part.width}v${part.height}h${-part.width}Z`).join(' ')
  return { left: `${clip.x}px`, top: `${clip.y}px`, width: `${clip.width}px`, height: `${clip.height}px`,
    clipPath: clips.length > 1 ? `path("${path}")` : undefined }
})
const checkboxStyle = computed(() => {
  const size = token.value.controlInteractiveSize
  return { left: `${props.position.rect.x - props.position.clip.x}px`, top: `${props.position.rect.y - props.position.clip.y}px`,
    width: `${size}px`, height: `${size}px`, transform: `scale(${props.position.rect.width / size})`, transformOrigin: '0 0' }
})
function onPointerdown(event: PointerEvent): void { event.stopPropagation(); pressedEntry.value = event.button === 0 ? props.entry : undefined }
function onChange(event: { target: { checked: boolean } }): void {
  const entry = pressedEntry.value ?? props.entry
  pressedEntry.value = undefined
  entry.activate(event.target.checked)
}
function onKeydown(event: KeyboardEvent): void {
  if (event.isComposing || event.keyCode === 229) return
  if (event.key === 'Enter') { event.preventDefault(); event.stopPropagation(); props.entry.activate() }
  else if (event.key === ' ') { event.stopPropagation(); pressedEntry.value = props.entry }
}
</script>

<template>
  <span class="sheet-checkbox-clip" :style="clipStyle" :data-sheet-checkbox-cell="entry.key">
    <Checkbox class="sheet-checkbox-control" :style="checkboxStyle" :checked="entry.checked" :disabled="entry.disabled"
      :tab-index="selectedKey === entry.key ? 0 : -1" :data-ribbon-editor-owner="entry.inputOwner"
      :aria-label="`${entry.key} ${translate('sheet.validation.checkbox', '复选框')}`" :skip-group="true"
      @pointerdown="onPointerdown" @pointercancel="pressedEntry = undefined" @blur="pressedEntry = undefined"
      @dblclick.stop @keydown="onKeydown" @change="onChange" />
  </span>
</template>

<style scoped>
.sheet-checkbox-clip { position: absolute; z-index: 3; overflow: hidden; pointer-events: none; }
.sheet-checkbox-control { position: absolute; box-sizing: border-box; margin: 0; padding: 0; line-height: 1; pointer-events: auto; }
</style>
