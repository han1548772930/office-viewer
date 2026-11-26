<script setup lang="ts">
import { Alert, Button, Form, FormItem, Modal, Select, theme } from 'antdv-next'
import { computed, nextTick, onScopeDispose, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { SheetSlicerCacheOptions } from '../../domain/model.ts'
import { slicerPanelOwnerCurrent, type SlicerPanelOwner } from './slicer-panel-owner.ts'
import { SLICER_CACHE_OPTIONS_DEFAULT } from './slicer-cache-options-panel-state.ts'
import { useSlicerCacheOptions } from './use-slicer-cache-options.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; slicerId: string }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; changed: [slicerId: string] }>()
const { token } = theme.useToken()
const appearance = computed(() => ({ '--slicer-cache-text': token.value.colorText, '--slicer-cache-muted': token.value.colorTextSecondary }))
let focusGeneration = 0, returnFocus: { element: HTMLElement; owner: SlicerPanelOwner } | undefined
const controls = useSlicerCacheOptions({
  get studio() { return props.studio }, get workbookId() { return props.workbookId }, get sheetId() { return props.sheetId },
  get slicerId() { return props.slicerId }, get open() { return props.open }, changed: id => emit('changed', id),
  close: owner => {
    const captured = returnFocus, generation = ++focusGeneration
    returnFocus = undefined; emit('update:open', false)
    if (captured && owner) void nextTick(() => {
      if (focusGeneration !== generation || props.open || props.studio !== owner.studio || props.workbookId !== owner.target.workbookId
        || props.sheetId !== owner.target.sheetId || captured.owner.studio !== owner.studio || captured.owner.userId !== owner.userId
        || captured.owner.target.workbookId !== owner.target.workbookId || captured.owner.target.sheetId !== owner.target.sheetId
        || !slicerPanelOwnerCurrent(owner, true) || !captured.element.isConnected) return
      captured.element.focus({ preventScroll: true })
    })
  },
})
watch([() => props.open, () => props.studio, () => props.workbookId, () => props.sheetId, () => props.slicerId], (current, previous) => {
  if (current[0] || !previous || current.slice(1).some((value, index) => value !== previous[index + 1])) focusGeneration += 1
  returnFocus = undefined
  if (props.open && controls.snapshot.value && document.activeElement instanceof HTMLElement) returnFocus = { element: document.activeElement, owner: controls.snapshot.value.owner }
}, { immediate: true, flush: 'sync' })
onScopeDispose(() => { focusGeneration += 1; returnFocus = undefined })
const disabled = computed(() => controls.busy.value || !controls.canConfigure.value)
const sortOptions = [{ value: SLICER_CACHE_OPTIONS_DEFAULT, label: '文档默认值（升序）' }, { value: 'ascending', label: '升序' }, { value: 'descending', label: '降序' }]
const booleanOptions = [{ value: SLICER_CACHE_OPTIONS_DEFAULT, label: '文档默认值（关闭）' }, { value: 'true', label: '开启' }, { value: 'false', label: '关闭' }]
const crossOptions = [{ value: SLICER_CACHE_OPTIONS_DEFAULT, label: '文档默认值（显示无数据项）' }, { value: 'none', label: '不标记无数据项' },
  { value: 'showItemsWithDataAtTop', label: '有数据项置顶' }, { value: 'showItemsWithNoData', label: '显示无数据项' }]
function value(key: keyof SheetSlicerCacheOptions): string { const current = controls.draft.value[key]; return current === undefined ? SLICER_CACHE_OPTIONS_DEFAULT : String(current) }
</script>

<template>
  <Modal :open="open" title="排序与无数据项" :root-style="appearance" :width="440" :footer="null" :mask="{ closable: false }" :focus-trigger-after-close="false" destroy-on-hidden @cancel="controls.close">
    <Form class="slicer-cache-options" :style="appearance" layout="vertical" @submit.prevent="controls.apply">
      <p class="slicer-cache-note">这些设置会同步到使用同一来源字段的切片器，保留当前筛选和报表连接。</p>
      <Alert v-if="controls.stale.value" title="排序设置已过期，请重新读取后再保存。" type="warning" show-icon />
      <Alert v-else-if="controls.snapshot.value && !controls.canConfigure.value" title="当前用户没有修改这些设置的权限。" type="info" show-icon />
      <FormItem label="成员排序"><Select :value="value('sortOrder')" :options="sortOptions" aria-label="切片器成员排序" :disabled="disabled" @change="controls.set('sortOrder', $event)" /></FormItem>
      <FormItem label="月份和星期自定义列表"><Select :value="value('customListSort')" :options="booleanOptions" aria-label="切片器自定义列表排序" :disabled="disabled" @change="controls.set('customListSort', $event)" /></FormItem>
      <FormItem label="无数据项显示"><Select :value="value('crossFilter')" :options="crossOptions" aria-label="切片器无数据项显示" :disabled="disabled" @change="controls.set('crossFilter', $event)" /></FormItem>
      <FormItem label="隐藏无数据项"><Select :value="value('hideItemsWithNoData')" :options="booleanOptions" aria-label="切片器隐藏无数据项" :disabled="disabled || controls.draft.value.crossFilter === 'none'" @change="controls.set('hideItemsWithNoData', $event)" /></FormItem>
      <p v-if="controls.draft.value.crossFilter === 'none'" class="slicer-cache-note">不标记无数据项时，所有成员均可用；已有隐藏设置会保留。</p>
      <Alert v-if="controls.error.value" :title="controls.error.value" type="error" show-icon />
      <div class="slicer-cache-actions">
        <Button html-type="button" :disabled="controls.busy.value" @click="controls.read">重新读取</Button>
        <Button html-type="button" :disabled="disabled" @click="controls.reset">文档默认值</Button><span />
        <Button html-type="button" @click="controls.close">取消</Button>
        <Button html-type="button" type="primary" :loading="controls.busy.value" :disabled="disabled || !controls.dirty.value" @click="controls.apply">确定</Button>
      </div>
    </Form>
  </Modal>
</template>

<style scoped>
.slicer-cache-options { color: var(--slicer-cache-text); }
.slicer-cache-note { color: var(--slicer-cache-muted); font-size: 12px; line-height: 1.6; margin: 0 0 12px; }
.slicer-cache-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.slicer-cache-actions > span { flex: 1; }
</style>
