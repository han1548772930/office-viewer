<script setup lang="ts">
import { computed, h, inject, nextTick, onBeforeUnmount, ref, watch, type CSSProperties } from 'vue'
import { Alert, Button, Checkbox, Form, FormItem, Input, Listy, Popover, Progress, Spin, Tag, Tooltip, theme, type ListyRef } from 'antdv-next'
import { CalendarDays, ChevronDown, Settings } from 'lucide-vue-next'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import SheetDateValidationPicker from './SheetDateValidationPicker.vue'
import type { SheetValidationCellEntry, SheetValidationCellPopup } from './use-sheet-validation-cell.ts'
import type { SheetValidationListOption } from '../../features/sheet-validation/validation-list-input.ts'
import { useSheetValidationListQuery } from './use-sheet-validation-list-query.ts'
import { createSheetValidationListNavigation, sheetValidationListNavigationIndex } from './sheet-validation-list-navigation.ts'
import { sheetValidationListHighlightRange } from '../../features/sheet-validation/validation-list-highlight.ts'

const props = defineProps<{ entry: SheetValidationCellEntry; popup?: SheetValidationCellPopup; position: CSSProperties; region: string; currentYear: number }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback), { token } = theme.useToken()
const query = ref(''), stableQuery = ref(''), composing = ref(false), activeIndex = ref(0)
const list = ref<HTMLElement>(), listControl = ref<ListyRef>(), popupRoot = ref<HTMLElement>()
const owner = computed(() => props.popup?.hostKey === props.entry.hostKey ? props.popup : undefined)
const emptyOptions: readonly SheetValidationListOption[] = Object.freeze([])
const { matches: items, busy: queryBusy, progress: queryProgress, error: queryError, errorCode: queryErrorCode, refresh: retryQuery } = useSheetValidationListQuery({
  state: () => {
    const popup = owner.value, context = popup?.queryContext()
    return { owner: popup, options: popup?.kind === 'list' ? popup.options : emptyOptions, query: stableQuery.value,
      active: popup?.kind === 'list' && popup.visible && !!context, unitId: context?.unitId ?? '', baseRevision: context?.baseRevision ?? 0,
      current: () => owner.value === popup && popup?.current() === true }
  },
})
const queryProgressLabel = computed(() => queryProgress.value?.phase === 'querying'
  ? t('sheet.validation.list.query.querying', '正在搜索列表项') : t('sheet.validation.list.query.preparing', '正在准备列表项'))
const queryErrorLabel = computed(() => queryErrorCode.value === 'COMPUTE_MEMORY_LIMIT'
  ? t('sheet.validation.list.query.memory', '列表查询超过当前内存预算。') : t('sheet.validation.list.query.failed', '列表查询未能完成，请重试。'))
const listHeight = computed(() => Math.min(240, Math.max(1, items.value.length) * (token.value.fontHeight + token.value.paddingSM * 2)))
const frames = new Map<number, () => void>()
let activeComponent = true
const navigation = createSheetValidationListNavigation({
  state: () => {
    const popup = owner.value, context = popup?.queryContext()
    return { owner: popup, items: items.value, query: stableQuery.value, unitId: context?.unitId ?? '', baseRevision: context?.baseRevision ?? 0,
      ready: !!popup?.visible && !popup.disabled && !composing.value && !queryBusy.value && !queryError.value && query.value === stableQuery.value,
      current: () => owner.value === popup && popup?.current() === true }
  },
  activate: index => { activeIndex.value = index },
  scrollTo: (item, align) => listControl.value?.scrollTo({ key: item.value, align }),
  focusAt: (index, current) => {
    const input = list.value?.querySelector<HTMLInputElement>(`[data-validation-list-index="${index}"] input[type="checkbox"]:not(:disabled)`)
    if (!input || !current()) return false
    input.focus(); return input.ownerDocument.activeElement === input
  },
  nextFrame: async () => {
    await nextTick()
    if (!activeComponent) return
    await new Promise<void>(resolve => {
      const id = requestAnimationFrame(() => { frames.delete(id); resolve() })
      frames.set(id, resolve)
    })
  },
})
const callbacks = computed(() => {
  const popup = owner.value, entry = props.entry
  return { open: (value: boolean) => { if (!composing.value) { if (popup) popup.setOpen(value); else if (value) entry.open() } },
    after: (value: boolean) => popup?.afterOpenChange(value),
    close: () => { if (!composing.value) popup?.setOpen(false) },
    clear: () => { if (!composing.value && !popup?.disabled) popup?.chooseList([]) },
    composition: (value: boolean) => { if (owner.value === popup) { composing.value = value; popup?.setComposing(value) } },
    compositionStart: () => { if (owner.value === popup) { composing.value = true; popup?.setComposing(true) } },
    compositionEnd: (event: CompositionEvent) => {
      if (owner.value !== popup) return
      if (event.target instanceof HTMLInputElement) query.value = event.target.value
      composing.value = false; popup?.setComposing(false); stableQuery.value = query.value
    },
    retry: () => { if (owner.value === popup && popup?.current() && !composing.value) void retryQuery() },
    keydown: (event: KeyboardEvent) => { if (owner.value === popup) keydown(event) },
  }
})
watch(query, value => { if (!composing.value) stableQuery.value = value }, { flush: 'sync' })
watch(() => owner.value?.ownerKey, () => { query.value = ''; stableQuery.value = ''; composing.value = false }, { flush: 'sync' })
watch(() => [owner.value, items.value, stableQuery.value, query.value, composing.value, queryBusy.value, owner.value?.visible,
  owner.value?.queryContext()?.baseRevision] as const, (values, previous) => {
  navigation.invalidate()
  if (previous && values[0] === previous[0] && values[1] === previous[1] && values[2] === previous[2]) return
  activeIndex.value = 0
  const generation = navigation.generation, popup = owner.value
  void nextTick(() => {
    if (popup?.visible && popup.current() && owner.value === popup && generation === navigation.generation) listControl.value?.scrollTo(0)
  })
}, { flush: 'sync' })
watch(() => [owner.value?.ownerKey, owner.value?.visible] as const, async () => {
  const popup = owner.value
  if (!popup?.visible) return
  await nextTick()
  if (activeComponent && popup.current() && owner.value === popup && popup.visible && !composing.value) popupRoot.value?.querySelector<HTMLInputElement>('input:not(:disabled)')?.focus()
}, { immediate: true })
onBeforeUnmount(() => {
  activeComponent = false
  navigation.dispose()
  for (const [id, resolve] of frames) { cancelAnimationFrame(id); resolve() }
  frames.clear()
  const popup = owner.value; if (popup?.visible) { popup.setOpen(false); popup.afterOpenChange(false) }
})
const title = computed(() => t(props.entry.kind === 'date' ? 'sheet.validation.date.open' : 'sheet.validation.list.open', props.entry.kind === 'date' ? '选择日期或时间' : '选择列表项'))
function rowKey(item: SheetValidationListOption): string { return item.value }
function itemLabel(item: SheetValidationListOption) {
  const range = sheetValidationListHighlightRange(item.label, stableQuery.value)
  if (!range) return item.label
  return [item.label.slice(0, range.start), h('span', { 'data-validation-list-highlight': '', style: { fontWeight: token.value.fontWeightStrong } },
    item.label.slice(range.start, range.end)), item.label.slice(range.end)]
}
function itemRender(item: SheetValidationListOption, index: number) {
  const popup = owner.value, source = items.value, filter = stableQuery.value, context = popup?.queryContext()
  const current = () => {
    const latest = popup?.queryContext()
    return !!popup && !!context && owner.value === popup && popup.visible && !popup.disabled && !composing.value && !queryBusy.value
      && !queryError.value && source === items.value && filter === stableQuery.value && query.value === filter
      && latest?.unitId === context.unitId && latest?.baseRevision === context.baseRevision && popup.current()
      && owner.value === popup && source === items.value && filter === stableQuery.value && !composing.value
  }
  return h('div', { class: 'sheet-validation-list-row', 'data-validation-list-index': index }, [h(Checkbox, {
    class: 'sheet-validation-list-option', checked: popup?.selected.includes(item.value) === true,
    disabled: !current(), tabIndex: index === activeIndex.value ? 0 : -1,
    onFocus: () => { if (current()) activeIndex.value = index },
    onChange: (event: { target: { checked: boolean } }) => {
      if (!current() || !popup) return
      const selected = popup.selected.filter(value => value !== item.value)
      if (!current()) return
      popup.chooseList(event.target.checked ? popup.multiple ? [...selected, item.value] : [item.value] : selected)
    },
  }, () => item.color ? h(Tag, { color: item.color }, () => itemLabel(item)) : h('span', itemLabel(item)))])
}
function keydown(event: KeyboardEvent): void {
  if (composing.value || event.isComposing || event.keyCode === 229) { event.stopPropagation(); return }
  const popup = owner.value
  if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); popup?.setOpen(false); return }
  if (!(event.target instanceof HTMLInputElement) || !popup || popup.disabled || queryBusy.value || queryError.value) return
  const checkbox = event.target.type === 'checkbox'
  if (!checkbox && !['ArrowUp', 'ArrowDown'].includes(event.key)) return
  const row = event.target.closest<HTMLElement>('[data-validation-list-index]'), index = row ? Number(row.dataset.validationListIndex) : -1
  if (checkbox && (!Number.isSafeInteger(index) || !items.value[index])) return
  const next = sheetValidationListNavigationIndex(event.key, checkbox ? activeIndex.value : -1, items.value.length)
  if (next === undefined) return
  event.preventDefault(); event.stopPropagation()
  void navigation.focus(next, event.key === 'Home' ? 'top' : event.key === 'End' ? 'bottom' : 'auto')
}
</script>

<template>
  <Popover :open="owner?.visible === true" :trigger="['click']" placement="bottomRight" :destroy-on-hidden="true" :after-open-change="callbacks.after"
    :classes="{ root: 'sheet-validation-cell-popover' }" :styles="{ body: { padding: '12px' } }" @open-change="callbacks.open">
    <Tooltip v-if="entry.triggerVisible !== false" :title="title" :open="owner?.visible ? false : undefined">
      <Button class="sheet-validation-cell-trigger" :style="position" size="small" type="text" :disabled="entry.disabled" :aria-label="`${title} ${entry.key}`"
        :aria-expanded="owner?.visible === true" aria-haspopup="dialog" :data-ribbon-editor-owner="entry.inputOwner"
        @pointerdown.stop @dblclick.stop @contextmenu.prevent.stop @keydown.stop>
        <CalendarDays v-if="entry.kind === 'date'" :size="14" aria-hidden="true" /><ChevronDown v-else :size="14" aria-hidden="true" />
      </Button>
    </Tooltip>
    <span v-else class="sheet-validation-hidden-trigger" :style="position" aria-hidden="true" />
    <template #content>
      <div v-if="owner" ref="popupRoot" :key="owner.ownerKey" :data-sheet-validation-popup-owner="owner.ownerKey" :data-ribbon-editor-owner="owner.inputOwner"
        class="sheet-validation-popup" role="dialog" :aria-label="`${title} ${owner.key}`" :style="{ '--validation-border': token.colorBorderSecondary }"
        @pointerdown.stop @dblclick.stop @contextmenu.prevent.stop @keydown.stop="callbacks.keydown">
        <SheetDateValidationPicker v-if="owner.kind === 'date' && owner.date" :state="owner.date" :disabled="owner.disabled" :message="owner.message"
          :region="owner.region" :current-year="owner.currentYear" :submit="owner.submitDate" @cancel="callbacks.close" @composition="callbacks.composition" />
        <Form v-else layout="vertical" class="sheet-validation-list-form" @submit.prevent
          @compositionstart.capture="callbacks.compositionStart" @compositionend.capture="callbacks.compositionEnd">
          <FormItem :label="t('sheet.validation.list.search', '搜索列表项')">
            <Input v-model:value="query" allow-clear :aria-label="t('sheet.validation.list.search', '搜索列表项')" />
          </FormItem>
          <div v-if="queryBusy" class="sheet-validation-query-progress" role="status" aria-live="polite">
            <Spin size="small" /><span>{{ queryProgressLabel }}</span>
            <span v-if="queryProgress">{{ queryProgress.processed }} / {{ queryProgress.total }}</span>
            <Progress v-if="queryProgress" :percent="Math.round(queryProgress.ratio * 100)" :show-info="false" size="small" />
          </div>
          <div ref="list" class="sheet-validation-list-options" :aria-label="t('sheet.validation.list.options', '列表选项')" :aria-busy="queryBusy">
            <Listy v-if="!queryBusy && !queryError && items.length" ref="listControl" :items="items" :row-key="rowKey" :height="listHeight" :virtual="true" :item-render="itemRender"
              :styles="{ item: { paddingBlock: '4px', paddingInline: '4px' } }" />
          </div>
          <Alert v-if="queryError" type="error" :title="queryErrorLabel" show-icon>
            <template #action><Button size="small" :disabled="composing" @click="callbacks.retry">{{ t('sheet.validation.list.query.retry', '重试') }}</Button></template>
          </Alert>
          <Alert v-else-if="!queryBusy && !items.length" type="info" :title="t('sheet.validation.list.empty', '没有匹配的列表项。')" />
          <Alert v-if="owner.message" type="error" :title="owner.message" show-icon />
          <div class="sheet-validation-list-actions">
            <Button :disabled="owner.disabled || composing || !owner.selected.length" @click="callbacks.clear">{{ t('sheet.validation.list.clear', '清除选择') }}</Button>
            <Button :disabled="composing" @click="callbacks.close">{{ t('common.close', '关闭') }}</Button>
          </div>
        </Form>
        <Button v-if="owner.canEditRule" type="text" size="small" class="sheet-validation-edit-rule" :disabled="composing" @click="owner.editRule">
          <Settings :size="13" aria-hidden="true" />{{ t('sheet.validation.edit-title', '编辑数据验证') }}
        </Button>
      </div>
    </template>
  </Popover>
</template>

<style scoped>
.sheet-validation-cell-trigger { position: absolute; z-index: 12; min-width: 18px; width: 20px; height: 20px; padding: 1px; background: var(--sheet-canvas-background); color: var(--sheet-canvas-secondary); }
.sheet-validation-hidden-trigger { position: absolute; pointer-events: none; }
.sheet-validation-list-form { width: 288px; max-width: calc(100vw - 48px); }
.sheet-validation-list-form :deep(.ant-form-item) { margin-block-end: 8px; }
.sheet-validation-list-options { min-height: 0; max-height: 240px; padding-block: 4px; }
.sheet-validation-list-options :deep(.sheet-validation-list-row) { min-width: 0; }
.sheet-validation-list-options :deep(.sheet-validation-list-option) { width: 100%; margin-inline-start: 0; }
.sheet-validation-list-options :deep(.ant-checkbox-label) { min-width: 0; white-space: normal; overflow-wrap: anywhere; }
.sheet-validation-list-options :deep(.ant-tag) { white-space: normal; margin-inline-end: 0; }
.sheet-validation-query-progress { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; padding-block: 8px; }
.sheet-validation-query-progress :deep(.ant-progress) { width: 100%; margin: 0; }
.sheet-validation-list-actions { display: flex; justify-content: space-between; gap: 8px; border-block-start: 1px solid var(--validation-border); margin-block-start: 10px; padding-block-start: 10px; }
.sheet-validation-edit-rule { margin-block-start: 8px; }
</style>
