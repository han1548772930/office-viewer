<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Empty, Modal, Select } from 'antdv-next'
import { OfficeTranslatorKey } from '../../i18n/office-i18n'
import { collectFeatureSearchResults, type FeatureSearchResult } from './feature-search'
import type { RibbonGroup, RibbonTab } from './ribbon-schema'
import { RibbonEditorFocusScopeKey } from './ribbon-popup-scope'

const props = defineProps<{
  open: boolean
  tabs: readonly RibbonTab[]
  contextGroups?: readonly RibbonGroup[]
}>()
const emit = defineEmits<{ 'update:open': [open: boolean]; execute: [result: FeatureSearchResult] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const editor = inject(RibbonEditorFocusScopeKey, undefined)
const control = ref<{ focus: () => void }>()
const panel = ref<HTMLElement>()
const resultsOpen = ref(false)
let previousFocus: HTMLElement | undefined
watch(() => props.open, open => {
  resultsOpen.value = open
  if (!open) return
  previousFocus = document.activeElement instanceof HTMLElement ? document.activeElement : undefined
}, { flush: 'sync' })
const options = computed(() => !props.open ? [] : collectFeatureSearchResults(props.tabs, t('feature-search.ribbon', '功能区'), {
  title: t('feature-search.context-menu', '右键菜单'), groups: props.contextGroups ?? [],
})
  .map(result => ({ ...result, value: result.id, label: result.title })))
function restoreFocus() {
  const element = previousFocus
  previousFocus = undefined
  if (element?.isConnected && (document.activeElement === document.body || panel.value?.contains(document.activeElement))) {
    element.focus({ preventScroll: true })
  }
}
function popupContainer(trigger: HTMLElement) { return panel.value ?? trigger }
function select(value: unknown) {
  const result = options.value.find(option => option.value === value)
  if (!result) return
  emit('update:open', false)
  restoreFocus()
  emit('execute', result)
}
</script>

<template>
  <Modal :open="open" :title="t('feature-search.title', '搜索功能')" :width="560" :footer="null"
    :closable="{ 'aria-label': t('common.close', '关闭') }" :focusable="{ focusTriggerAfterClose: false }" :after-close="restoreFocus"
    destroy-on-hidden @after-open-change="value => { if (value && open) control?.focus() }" @cancel="emit('update:open', false)">
    <div ref="panel" class="feature-search-panel" :data-ribbon-editor-owner="editor?.ownerId">
      <Select ref="control" v-model:open="resultsOpen" :get-popup-container="popupContainer" :aria-label="t('feature-search.title', '搜索功能')" :options="options"
        :placeholder="t('feature-search.placeholder', '搜索命令或功能')" :show-search="{ optionFilterProp: 'keywords' }"
        :default-active-first-option="true" class="feature-search-select" @select="select">
        <template #optionRender="{ option }">
          <div class="feature-search-option">
            <span :data-ribbon-editor-owner="editor?.ownerId">{{ option.data.title }}</span><small>{{ option.data.path }}</small>
          </div>
        </template>
        <template #notFoundContent><Empty :description="t('feature-search.empty', '没有匹配的功能')" /></template>
      </Select>
    </div>
  </Modal>
</template>

<style scoped>
.feature-search-panel { position: relative; }
.feature-search-select { inline-size: 100%; }
.feature-search-option { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
.feature-search-option > small { max-inline-size: 55%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ant-color-text-secondary, currentColor); }
</style>
