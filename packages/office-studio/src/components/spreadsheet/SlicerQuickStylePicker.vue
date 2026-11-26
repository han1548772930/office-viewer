<script setup lang="ts">
import { Button, Tooltip, theme } from 'antdv-next'
import { computed } from 'vue'
import type { SlicerType } from '../../domain/model'
import { getSlicerStyle, listSlicerStyleNames, type SlicerStyleOwner } from '../../features/slicer/slicer-style-catalog'
import { resolveSlicerRenderProfile, slicerRenderRegionCss } from '../../features/slicer/slicer-style-rendering'

const props = defineProps<{ slicerId: string; kind: SlicerType; style?: string; styleOwner?: SlicerStyleOwner; expectedRevision: number; disabled?: boolean }>()
const emit = defineEmits<{ insert: [input: { slicerId: string; expectedRevision: number; patch: { style: string | undefined } }] }>()
const { token } = theme.useToken()
const activeStyle = computed(() => getSlicerStyle(props.styleOwner ?? {}, props.kind, props.style).name)
const styles = computed(() => ['', ...listSlicerStyleNames(props.styleOwner ?? {}, props.kind)].map(name => ({
  name,
  profile: resolveSlicerRenderProfile(props.styleOwner ?? {}, props.kind, name || undefined, {
    background: token.value.colorBgContainer, text: token.value.colorText,
    secondaryFill: token.value.colorFillSecondary, mutedText: token.value.colorTextTertiary,
  }),
})))
function choose(style: string): void {
  if (!props.disabled) emit('insert', { slicerId: props.slicerId, expectedRevision: props.expectedRevision, patch: { style: style || undefined } })
}
</script>

<template>
  <div class="slicer-quick-styles" role="group" :aria-label="kind === 'timeline' ? '时间线快速样式' : '切片器快速样式'" :style="{ '--style-border': token.colorBorder, '--style-selected': token.colorPrimary, '--style-surface': token.colorBgElevated }">
    <Tooltip v-for="style in styles" :key="style.name || 'default'" :title="style.name || '文档默认值'" :trigger="['hover', 'focus']">
      <Button html-type="button" class="slicer-quick-style" :class="{ selected: activeStyle === style.name }" :aria-label="style.name || '文档默认值'" :aria-pressed="activeStyle === style.name" :disabled="disabled" @click="choose(style.name)">
        <span v-if="style.profile.type === 'slicer'" class="slicer-quick-sample" aria-hidden="true" :style="slicerRenderRegionCss(style.profile.whole)">
          <span v-for="(region, index) in [
            style.profile.header,
            style.profile.items.selectedItemWithData,
            style.profile.items.unselectedItemWithData,
            style.profile.items.unselectedItemWithData,
          ]" :key="index" class="sample-row" :class="{ 'sample-header': index === 0 }" :style="slicerRenderRegionCss(region)">
            <i />
          </span>
        </span>
        <span v-else class="slicer-quick-sample timeline-quick-sample" aria-hidden="true" :style="slicerRenderRegionCss(style.profile.whole)">
          <span class="sample-row sample-header" :style="slicerRenderRegionCss(style.profile.header)"><i /></span>
          <span class="sample-row sample-label" :style="slicerRenderRegionCss(style.profile.elements.selectionLabel)"><i /></span>
          <span class="sample-periods">
            <span v-for="index in 5" :key="index" :style="slicerRenderRegionCss(index % 2 === 1 ? style.profile.elements.periodLabel1 : style.profile.elements.periodLabel2)"><i /></span>
          </span>
          <span class="sample-blocks">
            <span v-for="index in 5" :key="index" :style="slicerRenderRegionCss(index >= 2 && index <= 4 ? style.profile.elements.selectedTimeBlock : style.profile.elements.unselectedTimeBlock)" />
          </span>
        </span>
      </Button>
    </Tooltip>
  </div>
</template>

<style scoped>
.slicer-quick-styles { box-sizing: border-box; width: min(320px, calc(100vw - 48px)); max-height: 70vh; overflow: auto; display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; padding: 3px; }
.slicer-quick-style { box-sizing: border-box; width: 100%; min-width: 0; height: 72px; padding: 4px; overflow: hidden; border: 1px solid var(--style-border); background: var(--style-surface); }
.slicer-quick-style.selected { outline: 2px solid var(--style-selected); outline-offset: 1px; }
.slicer-quick-sample { box-sizing: border-box; width: 100%; height: 100%; min-width: 0; display: flex; flex-direction: column; gap: 2px; padding: 4px; overflow: hidden; border-radius: 2px; }
.sample-row { box-sizing: border-box; display: flex; flex: 1; min-height: 0; padding-inline: 4px; border-radius: 2px; }
.sample-header { flex: 0 0 10px; }
.slicer-quick-sample i { display: block; width: 50%; height: 1px; background: currentColor; }
.timeline-quick-sample { gap: 4px; }
.sample-label i { width: 25%; }
.sample-periods, .sample-blocks { display: flex; flex: 0 0 8px; gap: 1px; }
.sample-blocks { flex-basis: 12px; }
.sample-periods > span, .sample-blocks > span { box-sizing: border-box; display: flex; flex: 1; min-width: 0; }
</style>
