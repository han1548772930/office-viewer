<script setup lang="ts">
import { computed, inject, nextTick, onMounted, onUnmounted, ref, watch } from 'vue'
import { Button, Dropdown, Slider, Tooltip, theme } from 'antdv-next'
import { ChevronDown, ChevronLeft, ChevronRight, List, Lock, MoreHorizontal, Plus, ZoomIn, ZoomOut } from 'lucide-vue-next'
import type { ExcelSheetView } from '../../services/excelViewAdapter'
import { formatSelectionStatistic, selectSelectionStatistic, visibleSheetMenuEntries } from '../../features/sheet-workbench/sheet-workbench'
import type { calculateSelectionStatistics, SelectionStatisticType } from '../../features/sheet-workbench/sheet-workbench'
import { createSheetTabMenu } from './sheet-tab-menu.ts'
import { OfficeTranslatorKey, useOfficeRegion } from '../../i18n/office-i18n.ts'
import { sheetTabDropOrder } from '../../features/sheet-workbench/sheet-tab-drag-sort.ts'
import type { WorksheetLifecycleCommandInput } from '../../features/sheet-workbench/sheet-worksheet-lifecycle.ts'
import SheetTabColorDialog from './SheetTabColorDialog.vue'

const props = defineProps<{ sheets: ExcelSheetView[]; sheetOrder?: readonly string[]; workbookId?: string; revision?: number; historyGeneration?: number; activeSheetId: string; zoom: number; statistics: ReturnType<typeof calculateSelectionStatistics>; numberFormat?: string }>()
const emit = defineEmits<{
  (event: 'activate', id: string): void
  (event: 'add'): void
  (event: 'rename', id: string): void
  (event: 'copy', id: string): void
  (event: 'move', id: string, order: number): void
  (event: 'tab-color', id: string, color: string, context?: WorksheetLifecycleCommandInput): void
  (event: 'protect-sheet', id: string): void
  (event: 'unprotect-sheet', id: string): void
  (event: 'delete', id: string): void
  (event: 'hide', id: string): void
  (event: 'show', id: string): void
  (event: 'update:zoom', value: number): void
}>()
const tabs = ref<HTMLElement | null>(null)
const { token } = theme.useToken()
const footerTheme = computed(() => ({
  '--sheet-footer-bg': token.value.colorBgContainer, '--sheet-footer-text': token.value.colorTextSecondary,
  '--sheet-footer-border': token.value.colorBorderSecondary, '--sheet-tab-active-bg': token.value.controlItemBgActive,
  '--sheet-tab-active-text': token.value.colorPrimary, '--sheet-tab-hover-bg': token.value.colorFillTertiary,
  '--sheet-tab-focus': token.value.colorPrimaryBorder,
}))
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const preferredStatistic = ref<SelectionStatisticType>('SUM')
const visibleSheets = computed(() => props.sheets.filter(sheet => sheet.visible !== false))
const sheetOrder = computed(() => props.sheetOrder ?? props.sheets.map(sheet => sheet.id))
const draggingTab = ref<{ id: string; order: string } | null>(null)
const tabDrop = ref<{ id: string; after: boolean } | null>(null)
const tabColorRequest = ref<{ sheetId: string; name: string; color?: string; context: WorksheetLifecycleCommandInput } | null>(null)

function openTabColor(id: string) {
  const sheet = props.sheets.find(candidate => candidate.id === id)
  if (!sheet) return
  tabColorRequest.value = { sheetId: id, name: sheet.name, color: sheet.tabColor,
    context: { unitId: props.workbookId, expectedRevision: props.revision, expectedHistoryGeneration: props.historyGeneration,
      ...(props.workbookId ? { host: { unitId: props.workbookId, sheetId: props.activeSheetId } } : {}) } }
}
function confirmTabColor(color: string) {
  const request = tabColorRequest.value
  if (!request || request.context.unitId !== props.workbookId) return
  const sheet = props.sheets.find(candidate => candidate.id === request.sheetId)
  if (!sheet || sheet.tabColor !== request.color || request.context.host && request.context.host.sheetId !== props.activeSheetId) return
  emit('tab-color', request.sheetId, color, request.context)
  tabColorRequest.value = null
}
function restoreTabFocus() {
  void nextTick(() => tabs.value?.querySelector<HTMLButtonElement>('[aria-selected="true"]')?.focus())
}
watch(() => [props.workbookId, props.activeSheetId, props.revision, props.sheets], () => {
  const request = tabColorRequest.value
  if (request && (request.context.unitId !== props.workbookId || request.context.expectedRevision !== props.revision
    || request.context.host && request.context.host.sheetId !== props.activeSheetId
    || props.sheets.find(sheet => sheet.id === request.sheetId)?.tabColor !== request.color)) tabColorRequest.value = null
})

function startTabDrag(event: DragEvent, id: string) {
  draggingTab.value = { id, order: JSON.stringify(sheetOrder.value) }
  tabDrop.value = null
  if (event.dataTransfer) { event.dataTransfer.effectAllowed = 'move'; event.dataTransfer.setData('text/plain', id) }
}
function overTabDrag(event: DragEvent, id: string) {
  if (!draggingTab.value) return
  event.preventDefault()
  if (event.dataTransfer) event.dataTransfer.dropEffect = 'move'
  const element = event.currentTarget
  if (!(element instanceof HTMLElement)) return
  const bounds = element.getBoundingClientRect()
  const rtl = getComputedStyle(element).direction === 'rtl'
  tabDrop.value = { id, after: rtl ? event.clientX < bounds.left + bounds.width / 2 : event.clientX > bounds.left + bounds.width / 2 }
  if (tabs.value) {
    const viewport = tabs.value.getBoundingClientRect()
    if (event.clientX < viewport.left + 24) tabs.value.scrollBy({ left: -16 })
    else if (event.clientX > viewport.right - 24) tabs.value.scrollBy({ left: 16 })
  }
}
function endTabDrag() { draggingTab.value = null; tabDrop.value = null }
function dropTab(event: DragEvent, id: string) {
  const drag = draggingTab.value, drop = tabDrop.value
  if (!drag) return
  event.preventDefault()
  endTabDrag()
  if (drag.order !== JSON.stringify(sheetOrder.value) || !drop || drop.id !== id) return
  const order = sheetTabDropOrder(sheetOrder.value, visibleSheets.value.map(sheet => sheet.id), drag.id, id, drop.after)
  if (order !== undefined && order >= 0 && order !== sheetOrder.value.indexOf(drag.id)) emit('move', drag.id, order)
}
const menuEntries = computed(() => visibleSheetMenuEntries(props.sheets, props.activeSheetId))
const statistic = computed(() => selectSelectionStatistic(props.statistics, preferredStatistic.value))
const region = useOfficeRegion()
function statisticCaption(item: NonNullable<typeof statistic.value>) { return formatSelectionStatistic(item, { locale: region.value, numberFormat: props.numberFormat }) }
const statisticsMenu = computed(() => ({
  items: props.statistics.items.map(item => ({ key: item.id, label: statisticCaption(item) })),
  selectedKeys: statistic.value ? [statistic.value.id] : [],
  onClick: ({ key }: { key: string | number }) => {
    const item = props.statistics.items.find(candidate => candidate.id === key)
    if (item) preferredStatistic.value = item.id
  },
}))
const allSheetsMenu = computed(() => ({
  items: menuEntries.value.map(sheet => ({ key: sheet.id, label: `${sheet.name}${sheet.hidden ? '（隐藏，点击显示）' : ''}` })),
  selectedKeys: [props.activeSheetId],
  onClick: ({ key }: { key: string | number }) => {
    const sheet = menuEntries.value.find(entry => entry.id === String(key))
    if (sheet?.hidden) emit('show', sheet.id)
    else if (sheet) emit('activate', sheet.id)
  },
}))

function sheetMenu(id: string) {
  const targetSheet = props.sheets.find(sheet => sheet.id === id)
  return createSheetTabMenu(id, visibleSheets.value.length, Boolean(targetSheet?.worksheetProtection), translate, {
    rename: sheetId => emit('rename', sheetId),
    copy: sheetId => emit('copy', sheetId),
    protect: sheetId => emit('protect-sheet', sheetId),
    unprotect: sheetId => emit('unprotect-sheet', sheetId),
    setColor: (sheetId, color) => emit('tab-color', sheetId, color),
    customColor: openTabColor,
    hide: sheetId => emit('hide', sheetId),
    delete: sheetId => emit('delete', sheetId),
  })
}

function onTabKeydown(event: KeyboardEvent, index: number) {
  const sheets = visibleSheets.value
  let next = index
  if (event.key === 'ArrowRight') next = (index + 1) % sheets.length
  else if (event.key === 'ArrowLeft') next = (index - 1 + sheets.length) % sheets.length
  else if (event.key === 'Home') next = 0
  else if (event.key === 'End') next = sheets.length - 1
  else return
  event.preventDefault()
  event.stopPropagation()
  const sheet = sheets[next]
  if (!sheet) return
  emit('activate', sheet.id)
  void nextTick(() => tabs.value?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus())
}

const canScrollLeft = ref(false)
const canScrollRight = ref(false)
let observer: ResizeObserver | undefined
function measureTabs() {
  const element = tabs.value
  if (!element) return
  canScrollLeft.value = element.scrollLeft > 1
  canScrollRight.value = element.scrollLeft + element.clientWidth < element.scrollWidth - 1
}
function scrollTabs(direction: -1 | 1) { tabs.value?.scrollBy({ left: direction * 160, behavior: 'smooth' }) }
watch(() => [props.activeSheetId, props.sheets.map(sheet => `${sheet.id}:${sheet.name}:${sheet.visible}`).join('|')], () => {
  void nextTick(() => {
    const active = tabs.value?.querySelector<HTMLButtonElement>('[aria-selected="true"]')
    if (active && tabs.value) {
      if (active.offsetLeft < tabs.value.scrollLeft) tabs.value.scrollLeft = active.offsetLeft
      else if (active.offsetLeft + active.offsetWidth > tabs.value.scrollLeft + tabs.value.clientWidth) tabs.value.scrollLeft = active.offsetLeft + active.offsetWidth - tabs.value.clientWidth
    }
    measureTabs()
  })
})
onMounted(() => { if (tabs.value) { observer = new ResizeObserver(measureTabs); observer.observe(tabs.value); measureTabs() } })
onUnmounted(() => { observer?.disconnect(); endTabDrag() })
</script>

<template>
  <footer class="sheet-bar" :style="footerTheme" data-testid="sheet-footer" @keydown.stop>
    <div class="sheet-navigation">
      <Tooltip title="添加工作表"><Button type="text" size="small" aria-label="添加工作表" @click="emit('add')"><Plus :size="17" /></Button></Tooltip>
      <Dropdown :trigger="['click']" placement="topLeft" :menu="allSheetsMenu"><Button type="text" size="small" aria-label="所有工作表"><List :size="16" /></Button></Dropdown>
      <div ref="tabs" class="sheet-tabs" role="tablist" aria-label="工作表" @scroll="measureTabs">
        <Dropdown v-for="(sheet, index) in visibleSheets" :key="sheet.id" :trigger="['contextmenu']" placement="topLeft" :menu="sheetMenu(sheet.id)">
          <Button class="sheet-tab" :class="{ active: activeSheetId === sheet.id, dragging: draggingTab?.id === sheet.id, 'drop-before': tabDrop?.id === sheet.id && !tabDrop.after, 'drop-after': tabDrop?.id === sheet.id && tabDrop.after }" :style="sheet.tabColor ? { '--sheet-tab-color': sheet.tabColor } : undefined" type="text" html-type="button" role="tab" :aria-label="sheet.name" :aria-selected="activeSheetId === sheet.id" :tabindex="activeSheetId === sheet.id ? 0 : -1" :data-sheet-id="sheet.id" draggable="true" @dragstart="startTabDrag($event, sheet.id)" @dragover="overTabDrag($event, sheet.id)" @drop="dropTab($event, sheet.id)" @dragend="endTabDrag" @click="emit('activate', sheet.id)" @dblclick="emit('rename', sheet.id)" @keydown="onTabKeydown($event, index)"><Lock v-if="sheet.worksheetProtection || sheet.rangeProtectionRules?.length" :size="12" aria-label="工作表包含保护规则" /><span>{{ sheet.name }}</span></Button>
        </Dropdown>
      </div>
      <div v-if="canScrollLeft || canScrollRight" class="sheet-scroll-controls">
        <Button type="text" size="small" aria-label="向左滚动工作表" :disabled="!canScrollLeft" @click="scrollTabs(-1)"><ChevronLeft :size="14" /></Button>
        <Button type="text" size="small" aria-label="向右滚动工作表" :disabled="!canScrollRight" @click="scrollTabs(1)"><ChevronRight :size="14" /></Button>
      </div>
    </div>
    <Dropdown v-if="statistic" :trigger="['click']" placement="topRight" :menu="statisticsMenu">
      <Button class="selection-statistic" type="text" size="small" aria-label="选区统计" :title="statisticCaption(statistic)"><span class="statistic-caption">{{ statisticCaption(statistic) }}</span><span class="statistic-compact" aria-hidden="true">Σ</span><ChevronDown :size="12" /></Button>
    </Dropdown>
    <Dropdown :trigger="['click']" placement="topRight" :menu="sheetMenu(activeSheetId)"><Button class="sheet-menu" type="text" size="small" aria-label="工作表操作"><MoreHorizontal :size="16" /></Button></Dropdown>
    <div class="sheet-zoom">
      <Tooltip title="缩小"><Button type="text" size="small" aria-label="缩小工作表" :disabled="zoom <= 10" @click="emit('update:zoom', Math.max(10, zoom - 10))"><ZoomOut :size="15" /></Button></Tooltip>
      <Slider class="zoom-slider" aria-label="工作表缩放" :min="10" :max="400" :step="10" :value="zoom" :tooltip="{ formatter: (value?: number) => `${value}%` }" @change="value => { if (typeof value === 'number') emit('update:zoom', value) }" />
      <Tooltip title="恢复 100%"><Button type="text" size="small" class="zoom-value" aria-label="恢复工作表缩放" @click="emit('update:zoom', 100)">{{ zoom }}%</Button></Tooltip>
      <Tooltip title="放大"><Button type="text" size="small" aria-label="放大工作表" :disabled="zoom >= 400" @click="emit('update:zoom', Math.min(400, zoom + 10))"><ZoomIn :size="15" /></Button></Tooltip>
    </div>
    <SheetTabColorDialog :request="tabColorRequest" @confirm="confirmTabColor" @cancel="tabColorRequest = null" @closed="restoreTabFocus" />
  </footer>
</template>

<style scoped>
.sheet-bar { display: flex; box-sizing: border-box; flex: 0 0 36px; min-width: 0; height: 36px; align-items: center; gap: 4px; border-top: 1px solid var(--sheet-footer-border); padding: 0 8px; color: var(--sheet-footer-text); background: var(--sheet-footer-bg); }
.sheet-navigation { display: flex; flex: 1; min-width: 0; height: 100%; align-items: center; }
.sheet-bar :deep(.ant-btn) { display: inline-flex; flex-shrink: 0; height: 28px; align-items: center; justify-content: center; padding: 0 6px; color: var(--sheet-footer-text); font-size: 12px; }
.sheet-tabs { position: relative; display: flex; min-width: 0; height: 100%; flex: 1; align-items: stretch; overflow-x: auto; scrollbar-width: none; }
.sheet-tabs::-webkit-scrollbar { display: none; }
.sheet-tab.ant-btn { display: flex; flex-shrink: 0; height: 100%; align-items: center; border: 0; border-radius: 0; border-right: 1px solid transparent; border-left: 1px solid transparent; padding: 0 15px; color: var(--sheet-footer-text); background: transparent; box-shadow: inset 0 -2px 0 var(--sheet-tab-color, transparent); cursor: pointer; font: inherit; font-size: 12px; white-space: nowrap; }
.sheet-tab.ant-btn:hover { background: var(--sheet-tab-hover-bg); }
.sheet-tab.ant-btn.active { color: var(--sheet-tab-active-text); background: var(--sheet-tab-active-bg); box-shadow: inset 0 -2px 0 var(--sheet-tab-color, var(--sheet-tab-active-text)); font-weight: 600; }
.sheet-tab.ant-btn.dragging { opacity: 0.6; }
.sheet-tab.ant-btn.drop-before { border-inline-start-color: var(--sheet-tab-active-text); }
.sheet-tab.ant-btn.drop-after { border-inline-end-color: var(--sheet-tab-active-text); }
.sheet-tab:focus-visible { outline: 2px solid var(--sheet-tab-focus); outline-offset: -3px; }
.sheet-scroll-controls, .sheet-zoom { display: flex; flex-shrink: 0; align-items: center; }
.selection-statistic { max-width: 230px; }
.statistic-caption { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.statistic-compact { display: none; }
.sheet-menu { border-right: 1px solid var(--sheet-footer-border); border-radius: 0; }
.zoom-slider { width: 88px; margin: 0 8px; }
.zoom-value { min-width: 43px; font-variant-numeric: tabular-nums; }
@media (max-width: 900px) { .zoom-slider { display: none; } }
@media (max-width: 600px) { .statistic-caption { display: none; } .statistic-compact { display: inline; } .sheet-bar { gap: 1px; padding: 0 3px; } }
</style>
