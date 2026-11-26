<script setup lang="ts">
import { computed, h, inject, ref, useId, type CSSProperties, type VNode } from 'vue'
import { Alert, Button, ConfigProvider, Dropdown, Slider, Spin, Tooltip, Typography } from 'antdv-next'
import { ChevronDown, FilterX, ListChecks, Minus } from 'lucide-vue-next'
import type { TimelineLevel } from '../../domain/model.ts'
import type { RenderRect } from '../../rendering/contracts.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n'
import { slicerControlTabIndex, slicerControlVisible, type SlicerControlPaneLayout } from '../../features/slicer/slicer-control-overlay.ts'
import { slicerPalette } from '../../features/slicer/slicer-renderer.ts'
import { mixSlicerColor, resolveSlicerRenderProfile, slicerItemRenderRegion, slicerRenderRegionCss,
  type SlicerRenderProfile, type SlicerRenderRegion } from '../../features/slicer/slicer-style-rendering.ts'
import { slicerRegionDecorationBackground } from '../../features/slicer/slicer-style-canvas.ts'
import { sheetTextGraphemes } from '../../features/sheet-workbench/sheet-cell-text-layout.ts'

const props = defineProps<{
  pane: SlicerControlPaneLayout
  panes: readonly SlicerControlPaneLayout[]
  colors: ReturnType<typeof slicerPalette>
  profile?: SlicerRenderProfile
  disabled: boolean
  menuOpen: boolean
  multiSelected: boolean
  focusedItem?: number
  focusedPeriod?: number
}>()
const emit = defineEmits<{
  'native-pointer': [event: PointerEvent]
  'activate': [id: string]
  'keyboard-action': [payload: { id: string; action: 'clear' | 'multiSelect' | 'item' | 'period'; event: MouseEvent; value?: string; index?: number }]
  'level-pointer': [payload: { id: string; event: PointerEvent; pane: SlicerControlPaneLayout }]
  'level-toggle': [payload: { id: string; event: MouseEvent }]
  'level-change': [payload: { id: string; level: TimelineLevel }]
  'level-close': [id: string]
  'scroll-pointer': [payload: { id: string; event: PointerEvent; pane: SlicerControlPaneLayout }]
  'scroll-change': [payload: { id: string; value: number }]
  'scroll-complete': [payload: { id: string; value: number }]
  'scroll-leave': [id: string]
  'cancel': [id: string]
  'retry': [id: string]
}>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback?: string) => fallback ?? _key)
const popupClass = `sheet-slicer-level-popup-${useId().replace(/[^a-zA-Z0-9_-]/g, '-')}`
const layout = computed(() => props.pane.layout), model = computed(() => layout.value.model)
const filtered = computed(() => model.value.type === 'timeline'
  ? model.value.timelineDateRange !== undefined || model.value.timelineDateCondition !== undefined || model.value.timelineRange !== undefined || model.value.selectedItems !== undefined
  : model.value.selectedItems !== undefined)
const dataRect = computed<RenderRect>(() => ({ x: layout.value.rect.x + 6, y: layout.value.rect.y + layout.value.header.height + 6,
  width: Math.max(0, layout.value.rect.width - 12), height: Math.max(0, layout.value.rect.height - layout.value.header.height - 12) }))
const profile = computed(() => props.profile ?? resolveSlicerRenderProfile(layout.value.styleOwner ?? {}, model.value.type, model.value.style,
  { background: props.colors.background, text: props.colors.text, secondaryFill: props.colors.item, mutedText: props.colors.scrollThumb }))
const hoveredItem = ref<number>()
const levels = computed<Array<{ key: TimelineLevel; label: string }>>(() => [
  { key: 'year', label: translate('sheet.timeline.year', '年') }, { key: 'quarter', label: translate('sheet.timeline.quarter', '季度') },
  { key: 'month', label: translate('sheet.timeline.month', '月') }, { key: 'day', label: translate('sheet.timeline.day', '日') },
])
const menu = computed(() => ({ items: levels.value.map(level => ({ ...level, disabled: props.disabled })), selectedKeys: [model.value.level ?? 'month'],
  onClick: (info: { key: string | number }) => {
    const level = levels.value.find(level => level.key === info.key)
    if (level && !props.disabled) emit('level-change', { id: layout.value.id, level: level.key })
  } }))
const controlTheme = computed(() => ({ token: {
  colorPrimary: props.colors.selected, colorBgContainer: props.colors.item, colorBgElevated: props.colors.background,
  colorText: props.colors.text, colorBorder: props.colors.itemBorder,
  controlHeightSM: 20, borderRadiusSM: 4,
}, components: { Button: { primaryColor: props.colors.selectedText } } }))
const headerControlTheme = computed(() => ({ token: { colorText: props.colors.headerText, colorBgContainer: props.colors.header },
  components: { Button: { textTextColor: props.colors.headerText, textTextHoverColor: props.colors.headerText,
    textTextActiveColor: props.colors.headerText } } }))
const outerStyle = computed<CSSProperties>(() => ({ position: 'absolute', left: `${props.pane.pane.clip.x}px`, top: `${props.pane.pane.clip.y}px`,
  width: `${props.pane.pane.clip.width}px`, height: `${props.pane.pane.clip.height}px`, clipPath: props.pane.clipPath }))
const innerStyle = computed<CSSProperties>(() => {
  const matrix = props.pane.pane.frame.matrix, clip = props.pane.pane.clip
  return { width: `${layout.value.geometry.width}px`, height: `${layout.value.geometry.height}px`,
    transform: `matrix(${matrix.a},${matrix.b},${matrix.c},${matrix.d},${matrix.e - clip.x},${matrix.f - clip.y})`, transformOrigin: '0 0' }
})
function style(rect: RenderRect): CSSProperties {
  return { position: 'absolute', left: `${rect.x - layout.value.rect.x}px`, top: `${rect.y - layout.value.rect.y}px`,
    width: `${rect.width}px`, height: `${rect.height}px`, minHeight: 0 }
}
function regionStyle(rect: RenderRect, region: SlicerRenderRegion, text = '', fontSize = 14, padding = 8): CSSProperties {
  return { ...style(rect), ...slicerRenderRegionCss(region), backgroundImage: slicerRegionDecorationBackground(region, rect.width, rect.height, text, fontSize, padding),
    backgroundRepeat: 'no-repeat', textDecorationLine: 'none' }
}
function itemRegion(item: SlicerControlPaneLayout['layout']['items'][number]): SlicerRenderRegion {
  const region = slicerItemRenderRegion(profile.value, { selected: item.selected, hasData: item.hasData !== false, hovered: hoveredItem.value === item.index })
  return item.partial ? { ...region, fill: mixSlicerColor(region.fill, props.colors.selected, .24) } : region
}
function itemLabel(item: SlicerControlPaneLayout['layout']['items'][number]): string { return item.value || translate('sheet.slicer.blank', '（空白）') }
function labelRuns(item: SlicerControlPaneLayout['layout']['items'][number]): string[] {
  const text = itemLabel(item), alignment = itemRegion(item).horizontal
  return alignment === 4 ? text.split(/(\s+)/u).filter(Boolean) : alignment >= 5 ? sheetTextGraphemes(text).map(glyph => glyph.text) : [text]
}
function periodStyle(rect: RenderRect, selected: boolean): CSSProperties {
  const region = slicerItemRenderRegion(profile.value, { selected, hasData: true, hovered: false })
  return { ...regionStyle(rect, region), borderRadius: 0 }
}
const levelStyle = computed(() => regionStyle(layout.value.level, profile.value.type === 'timeline' ? profile.value.elements.timeLevel : profile.value.whole,
  levels.value.find(level => level.key === (model.value.level ?? 'month'))?.label ?? '', 12, 4))
const visible = (rect: RenderRect) => slicerControlVisible(props.pane, rect)
const tabIndex = (rect: RenderRect) => slicerControlTabIndex(props.panes, props.pane, rect)
const periodRect = (rect: RenderRect): RenderRect => ({ x: rect.x, y: layout.value.track?.y ?? rect.y, width: rect.width,
  height: layout.value.track?.height ?? rect.height })
function activate() { if (!props.disabled) emit('activate', layout.value.id) }
function keyboard(event: MouseEvent, action: 'clear' | 'multiSelect' | 'item' | 'period', value?: string, index?: number) {
  if (!props.disabled && event.detail === 0) emit('keyboard-action', { id: layout.value.id, action, event, value, index })
}
function scroll(value: number | number[], complete: boolean) {
  if (props.disabled || typeof value !== 'number' || !Number.isFinite(value)) return
  if (complete) emit('scroll-complete', { id: layout.value.id, value })
  else emit('scroll-change', { id: layout.value.id, value })
}
function leaveScroll(event: FocusEvent): void {
  const target = event.currentTarget
  if (!(target instanceof HTMLElement) || event.relatedTarget instanceof Node && target.contains(event.relatedTarget)) return
  emit('scroll-leave', layout.value.id)
}
function popup(menuNode: VNode): VNode {
  return h('div', { 'data-slicer-popup': layout.value.id, onKeydownCapture: (event: KeyboardEvent) => {
    if (event.key === 'Escape' && !event.isComposing && event.keyCode !== 229) {
      event.preventDefault(); event.stopPropagation(); emit('cancel', layout.value.id)
    }
  } }, [menuNode])
}
function keydown(event: KeyboardEvent) {
  if (event.key === 'Escape' && !event.isComposing && event.keyCode !== 229) {
    event.preventDefault(); emit('cancel', layout.value.id)
  }
}
</script>

<template>
  <ConfigProvider :theme="controlTheme">
    <div class="slicer-control-pane" :style="outerStyle" :data-slicer-id="layout.id" :data-slicer-kind="model.type" :data-slicer-pane="pane.key"
      @keydown.stop="keydown" @dblclick.stop @contextmenu.stop>
      <div class="slicer-object-controls" :style="innerStyle">
        <div v-if="layout.dataStatus?.error && visible(dataRect)" class="slicer-data-feedback" :style="style(dataRect)" @pointerdown.stop @keydown.stop>
          <Alert type="error" :message="layout.dataStatus.error" />
          <Button size="small" :disabled="layout.dataStatus.pending" @click.stop="emit('retry', layout.id)">{{ translate('sheet.slicer.retry', '重新加载') }}</Button>
        </div>
        <div v-else-if="layout.dataStatus?.pending && !layout.items.length && !layout.periods.length && visible(dataRect)" class="slicer-data-feedback" :style="style(dataRect)">
          <Spin size="small" />
          <Typography.Text type="secondary" role="status">{{ translate('sheet.slicer.loading', '正在加载') }}</Typography.Text>
        </div>
        <ConfigProvider :theme="headerControlTheme">
        <Tooltip v-if="layout.header.height > 0 && visible(layout.clear)" :title="translate('sheet.slicer.clear', '清除筛选')">
          <Button class="slicer-header-control" data-slicer-control="clear" :style="style(layout.clear)" size="small" type="text"
            :disabled="disabled || !filtered"
            :tabindex="tabIndex(layout.clear)" :aria-hidden="tabIndex(layout.clear) === -1 ? true : undefined"
            :aria-label="translate('sheet.slicer.clear', '清除筛选')" @focus="activate" @pointerdown.stop="emit('native-pointer', $event)"
            @click="keyboard($event, 'clear')"><FilterX :size="14" aria-hidden="true" /></Button>
        </Tooltip>
        <Tooltip v-if="model.type === 'slicer' && layout.header.height > 0 && visible(layout.multiSelect)" :title="translate('sheet.slicer.multi-select', '多选')">
          <Button class="slicer-header-control" data-slicer-control="multiSelect" :style="style(layout.multiSelect)" size="small"
            :type="multiSelected ? 'primary' : 'text'" :disabled="disabled" :tabindex="tabIndex(layout.multiSelect)"
            :aria-hidden="tabIndex(layout.multiSelect) === -1 ? true : undefined" :aria-pressed="multiSelected"
            :aria-label="translate('sheet.slicer.multi-select', '多选')" @focus="activate" @pointerdown.stop="emit('native-pointer', $event)"
            @click="keyboard($event, 'multiSelect')"><ListChecks :size="14" aria-hidden="true" /></Button>
        </Tooltip>
        </ConfigProvider>
        <template v-if="model.type === 'slicer'">
          <Button v-for="item in layout.items.filter(item => visible(item.rect))" :key="item.index" class="slicer-item-control"
            data-slicer-control="item" :data-slicer-index="item.index" :style="regionStyle(item.rect, itemRegion(item), itemLabel(item))" size="small"
            :type="item.selected ? 'primary' : 'default'" :disabled="disabled" :tabindex="tabIndex(item.rect)"
            :aria-hidden="tabIndex(item.rect) === -1 ? true : undefined" :aria-pressed="item.partial ? 'mixed' : item.selected"
            :data-slicer-selection="item.partial ? 'partial' : item.selected ? 'checked' : 'unchecked'"
            :data-slicer-has-data="item.hasData !== false" :data-slicer-style-name="profile.reference.name"
            :aria-label="item.partial ? `${item.value || translate('sheet.slicer.blank', '（空白）')}，${translate('sheet.slicer.partial', '部分选中')}` : undefined"
            :class="{ 'slicer-content-focused': focusedItem === item.index, 'slicer-item-partial': item.partial }" @focus="activate"
            @pointerenter="hoveredItem = item.index" @pointerleave="hoveredItem = undefined"
            @pointerdown.stop="emit('native-pointer', $event)" @click="keyboard($event, 'item', item.value, item.index)">
            <span class="slicer-item-label" :class="{ 'slicer-item-label-distributed': itemRegion(item).horizontal >= 4 }">
              <span v-for="(run, index) in labelRuns(item)" :key="index">{{ run }}</span>
            </span>
            <Minus v-if="item.partial" :size="12" aria-hidden="true" class="slicer-item-partial-icon" />
          </Button>
        </template>
        <template v-else>
          <Dropdown v-if="model.showTimeLevel !== false && visible(layout.level)" :trigger="[]" :open="menuOpen" :menu="menu"
            :classes="{ root: popupClass }" :popup-render="popup" auto-focus destroy-on-hidden @update:open="open => { if (!open) emit('level-close', layout.id) }">
            <Button class="slicer-level-control" data-slicer-control="level" :style="levelStyle" size="small" :disabled="disabled"
              :tabindex="tabIndex(layout.level)" :aria-hidden="tabIndex(layout.level) === -1 ? true : undefined"
              :aria-label="translate('sheet.timeline.level', '时间粒度')" :aria-expanded="menuOpen" aria-haspopup="menu" @focus="activate"
              @pointerdown.prevent.stop="emit('level-pointer', { id: layout.id, event: $event, pane })"
              @click="emit('level-toggle', { id: layout.id, event: $event })">
              {{ levels.find(level => level.key === (model.level ?? 'month'))?.label }}<ChevronDown :size="12" aria-hidden="true" />
            </Button>
          </Dropdown>
          <Button v-for="period in layout.periods.filter(period => visible(periodRect(period.rect)))" :key="period.period.key"
            class="slicer-period-control" data-slicer-control="period" :data-slicer-index="period.period.key" :style="periodStyle(periodRect(period.rect), period.selected)"
            size="small" :type="period.selected ? 'primary' : 'default'" :disabled="disabled || layout.timeline?.interactive === false" :tabindex="tabIndex(periodRect(period.rect))"
            :aria-hidden="tabIndex(periodRect(period.rect)) === -1 ? true : undefined" :aria-pressed="period.selected"
            :aria-label="`${period.period.groupLabel} ${period.period.label}`" :class="{ 'slicer-content-focused': focusedPeriod === period.period.key }"
            @focus="activate" @pointerdown.stop="emit('native-pointer', $event)" @click="keyboard($event, 'period', undefined, period.period.key)" />
        </template>
        <div v-if="layout.scrollbar && visible(layout.scrollbar)" class="slicer-scroll-control" data-slicer-control="scroll" :style="style(layout.scrollbar)"
          @pointerdown.capture.prevent.stop="emit('scroll-pointer', { id: layout.id, event: $event, pane })" @focusin="activate" @focusout="leaveScroll">
          <Slider :value="layout.scrollOffset" :min="0" :max="layout.maxScrollOffset" :step="model.type === 'timeline' ? 1 : Math.max(1, model.columnCount ?? 1)"
            :vertical="model.type === 'slicer'" :reverse="model.type === 'slicer'" :disabled="disabled || layout.maxScrollOffset <= 0"
            :tab-index="tabIndex(layout.scrollbar)" :aria-hidden="tabIndex(layout.scrollbar) === -1 ? true : undefined"
            :tooltip="{ open: false }" :style="{ height: model.type === 'slicer' ? '100%' : undefined, margin: model.type === 'slicer' ? '0 4px' : '4px 0' }"
            :aria-label-for-handle="model.type === 'timeline' ? translate('sheet.timeline.scroll', '滚动时间线') : translate('sheet.slicer.scroll', '滚动切片器')"
            @change="scroll($event, false)" @change-complete="scroll($event, true)" />
        </div>
      </div>
    </div>
  </ConfigProvider>
</template>

<style scoped>
.slicer-control-pane { pointer-events: none; overflow: hidden; z-index: 2; }
.slicer-object-controls { position: absolute; left: 0; top: 0; overflow: hidden; pointer-events: none; }
.slicer-object-controls :deep(button), .slicer-scroll-control { pointer-events: auto; }
.slicer-header-control, .slicer-level-control { padding: 0; }
.slicer-item-control { justify-content: flex-start; padding: 0 8px; font-size: 14px; }
.slicer-item-control :deep(span) { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.slicer-item-label { display: inline-block; max-width: 100%; }
.slicer-item-label-distributed { display: flex; flex: 1; justify-content: space-between; width: 100%; }
.slicer-item-partial { border-color: var(--ant-color-primary, var(--sheet-canvas-primary)); background: var(--ant-color-primary-bg, var(--sheet-canvas-background)); }
.slicer-item-partial-icon { flex: none; margin-left: auto; color: var(--ant-color-primary, var(--sheet-canvas-primary)); }
.slicer-level-control { display: flex; gap: 3px; justify-content: space-between; padding: 0 4px; }
.slicer-period-control { border-radius: 0; padding: 0; }
.slicer-content-focused { outline: 2px solid var(--sheet-canvas-primary); outline-offset: -2px; }
.slicer-scroll-control { position: absolute; }
.slicer-data-feedback { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 6px; overflow: auto; }
.slicer-data-feedback :deep(.ant-alert) { padding: 4px 6px; font-size: 12px; }
</style>
