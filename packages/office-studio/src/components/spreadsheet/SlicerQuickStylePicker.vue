<script setup lang="ts">
import { Button, Tooltip, theme } from 'antdv-next'
import { computed } from 'vue'
import type { SlicerType } from '../../domain/model'
import { listSlicerStyleNames, type SlicerStyleOwner } from '../../features/slicer/slicer-style-catalog'
import { slicerPalette } from '../../features/slicer/slicer-renderer'

const props = defineProps<{ slicerId: string; kind: SlicerType; style?: string; styleOwner?: SlicerStyleOwner; expectedRevision: number; disabled?: boolean }>()
const emit = defineEmits<{ insert: [input: { slicerId: string; expectedRevision: number; patch: { style: string | undefined } }] }>()
const { token } = theme.useToken()
const styles = computed(() => ['', ...listSlicerStyleNames(props.styleOwner ?? {}, props.kind)])
function preview(style: string) {
  const palette = slicerPalette(style || undefined, props.kind === 'timeline', { background: token.value.colorBgContainer, text: token.value.colorText,
    secondaryFill: token.value.colorFillSecondary, mutedText: token.value.colorTextTertiary }, props.styleOwner)
  return { '--preview-header': palette.header, '--preview-item': palette.item, '--preview-border': palette.itemBorder }
}
function choose(style: string): void {
  if (!props.disabled) emit('insert', { slicerId: props.slicerId, expectedRevision: props.expectedRevision, patch: { style: style || undefined } })
}
</script>

<template>
  <div class="slicer-quick-styles" role="group" aria-label="切片器快速样式" :style="{ '--style-border': token.colorBorder, '--style-selected': token.colorPrimary, '--style-surface': token.colorBgElevated }">
    <Tooltip v-for="style in styles" :key="style || 'default'" :title="style || '文档默认值'">
      <Button html-type="button" class="slicer-quick-style" :class="{ selected: (props.style ?? '') === style }" :style="preview(style)" :aria-label="style || '文档默认值'" :aria-pressed="(props.style ?? '') === style" :disabled="disabled" @click="choose(style)">
        <span v-if="style" class="slicer-quick-sample" aria-hidden="true"><i /><i /><i /></span><span v-else>Aa</span>
      </Button>
    </Tooltip>
  </div>
</template>

<style scoped>
.slicer-quick-styles { width: min(320px, calc(100vw - 48px)); max-height: 70vh; overflow: auto; display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; padding: 3px; }
.slicer-quick-style { height: 66px; padding: 8px; border: 1px solid var(--style-border); background: var(--style-surface); }
.slicer-quick-style.selected { outline: 2px solid var(--style-selected); outline-offset: 1px; }
.slicer-quick-sample { width: 100%; height: 100%; display: grid; grid-template-rows: repeat(3, 1fr); gap: 3px; padding: 4px; background: var(--preview-header); border: 1px solid var(--preview-border); }
.slicer-quick-sample i { display: block; background: var(--preview-item); border: 1px solid var(--preview-border); }
</style>
