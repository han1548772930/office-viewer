<script setup lang="ts">
import { computed, inject, ref } from 'vue'
import { Drawer, Typography } from 'antdv-next'
import { OfficeDirectionKey, OfficeTranslatorKey } from '../../i18n/office-i18n'
import { visibleRibbonTabs, type RibbonTab } from './ribbon-schema'
import { RibbonEditorFocusScopeKey } from './ribbon-popup-scope'
import { OPEN_FEATURE_SEARCH, TOGGLE_SHORTCUT_PANEL } from './workbench-operations'

const props = defineProps<{
  tabs: readonly RibbonTab[]
  shortcuts?: readonly { title: string; shortcut: string }[]
}>()
const open = defineModel<boolean>('open', { required: true })
defineEmits<{ keydown: [event: KeyboardEvent] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const direction = inject(OfficeDirectionKey, ref<'ltr' | 'rtl'>('ltr'))
const editor = inject(RibbonEditorFocusScopeKey, undefined)
const mac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/u.test(navigator.platform)
const sections = computed(() => {
  if (!open.value) return []
  const global = [
    { title: t('feature-search.title', '搜索功能'), shortcut: 'Ctrl+Shift+P' },
    { title: t('shortcut-panel.title', '快捷键面板'), shortcut: 'Ctrl+\\' },
  ]
  const editorItems = visibleRibbonTabs(props.tabs).flatMap(tab => tab.groups.flatMap(group =>
    group.items.flatMap(item => item.shortcut && item.id !== OPEN_FEATURE_SEARCH && item.id !== TOGGLE_SHORTCUT_PANEL
      ? [{ title: item.title, shortcut: item.shortcut }] : []),
  ))
  const editorShortcuts = [...new Map([...editorItems, ...(props.shortcuts ?? [])]
    .map(item => [JSON.stringify([item.title, item.shortcut]), item])).values()]
  return [
    { title: t('shortcut-panel.global', '全局快捷键'), items: global },
    { title: t('shortcut-panel.editor', '常用编辑'), items: editorShortcuts },
  ]
})
</script>

<template>
  <Drawer v-model:open="open" :title="t('shortcut-panel.title', '快捷键面板')" :size="360" :mask="false"
    :placement="direction === 'rtl' ? 'left' : 'right'" :styles="{ wrapper: { maxWidth: '100vw' } }"
    data-office-shortcut-panel :data-ribbon-editor-owner="editor?.ownerId" destroy-on-hidden @keydown="$emit('keydown', $event)">
    <div :data-ribbon-editor-owner="editor?.ownerId">
      <section v-for="section in sections" :key="section.title" class="shortcut-section">
        <Typography.Title :level="5">{{ section.title }}</Typography.Title>
        <dl>
          <div v-for="item in section.items" :key="item.title + item.shortcut" class="shortcut-row">
            <dt>{{ item.title }}</dt>
            <dd><Typography.Text keyboard dir="ltr">{{ mac ? item.shortcut.replace('Ctrl+', '⌘+') : item.shortcut }}</Typography.Text></dd>
          </div>
        </dl>
      </section>
    </div>
  </Drawer>
</template>

<style scoped>
.shortcut-section + .shortcut-section { margin-block-start: 20px; }
.shortcut-section dl, .shortcut-row dd { margin: 0; }
.shortcut-row { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-block-size: 34px; }
</style>
