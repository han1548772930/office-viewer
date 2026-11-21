<script setup lang="ts">
import { Button, Dropdown, Select, Tooltip, type MenuProps } from 'antdv-next'
import { computed, nextTick, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import { BookOpenText, ChevronDown, FileSpreadsheet, FileText, FolderOpen, Save } from 'lucide-vue-next'
import type { ExportTask, WorkspaceKind } from './workspace-file-operations'
import { supportedOfficeLocales, type OfficeLocale } from '@office-studio/office-studio'

export type WorkspaceHeaderAction = 'new-workbook' | 'new-document' | 'open' | 'save' | 'export' | 'rename' | 'diagnostics' | 'cancel-import' | 'cancel-export'
const props = defineProps<{
  activeKind: WorkspaceKind
  fileName: string
  modified: boolean
  parserStatus: 'idle' | 'loading' | 'ready' | 'error'
  parserMessage: string
  exportTask: ExportTask | null
  diagnosticCount: number
  locale: OfficeLocale
}>()
const emit = defineEmits<{
  action: [action: WorkspaceHeaderAction]
  switch: [kind: WorkspaceKind]
  locale: [locale: OfficeLocale]
}>()
const tabs = ref<HTMLElement | null>(null)
const fileMenuOpen = ref(false)
const { t } = useI18n()
const localeOptions = supportedOfficeLocales.map(value => ({
  value,
  label: new Intl.DisplayNames([value], { type: 'language' }).of(value) ?? value,
}))
const menuItems = computed<MenuProps['items']>(() => [
  { key: 'new-workbook', label: t('workspace.new-workbook') },
  { key: 'new-document', label: t('workspace.new-document') },
  { type: 'divider' },
  { key: 'open', label: t('workspace.open-office-file') },
  { key: 'rename', label: t('workspace.rename') },
  { key: 'save', label: t('workspace.save-office-file'), disabled: Boolean(props.exportTask) },
  { key: 'export', label: t('workspace.export-office-file'), disabled: Boolean(props.exportTask) },
  ...(props.parserStatus === 'loading' ? [{ key: 'cancel-import', label: t('workspace.cancel-import') }] : []),
  ...(props.exportTask ? [{ key: 'cancel-export', label: t('workspace.cancel-export') }] : []),
  { type: 'divider' },
  { key: 'diagnostics', label: `${t('workspace.diagnostics')}${props.diagnosticCount ? ` (${props.diagnosticCount})` : ''}` },
])
const statusText = computed(() => props.exportTask?.message || (props.parserStatus === 'loading' ? props.parserMessage : props.modified ? t('workspace.unsaved') : t('workspace.saved')))

function onMenuClick(info: { key: string | number }): void {
  const key = String(info.key)
  if (key === 'new-workbook' || key === 'new-document' || key === 'open' || key === 'save' || key === 'export' || key === 'rename' || key === 'diagnostics' || key === 'cancel-import' || key === 'cancel-export') {
    fileMenuOpen.value = false
    emit('action', key)
  }
}

async function onTabsKeydown(event: KeyboardEvent): Promise<void> {
  let kind: WorkspaceKind
  if (event.key === 'Home') kind = 'workbook'
  else if (event.key === 'End') kind = 'document'
  else if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') kind = props.activeKind === 'workbook' ? 'document' : 'workbook'
  else return
  event.preventDefault()
  emit('switch', kind)
  await nextTick()
  tabs.value?.querySelector<HTMLElement>(`[data-workspace="${kind}"]`)?.focus()
}
</script>

<template>
  <header class="app-header workspace-header" :aria-label="t('workspace.app-label')">
    <a class="brand" href="#" :aria-label="t('workspace.brand-label')" @click.prevent="emit('switch', 'workbook')">
      <span class="brand-mark"><BookOpenText :size="18" /></span>
      <span class="brand-name">Office Studio</span>
    </a>
    <Dropdown v-model:open="fileMenuOpen" :menu="{ items: menuItems }" :trigger="['click']" :auto-focus="true" placement="bottomLeft" @menu-click="onMenuClick">
      <Button type="text" class="workspace-file-menu" :aria-label="t('workspace.file-menu-label')" aria-haspopup="menu" :aria-expanded="fileMenuOpen">{{ t('workspace.file-menu') }}<ChevronDown :size="13" /></Button>
    </Dropdown>
    <nav ref="tabs" class="workspace-tabs" :aria-label="t('workspace.tabs-label')" @keydown="onTabsKeydown">
      <div class="workspace-tablist" role="tablist" :aria-label="t('workspace.document-types')">
        <Button id="workbook-tab" type="text" role="tab" :aria-selected="activeKind === 'workbook'" aria-controls="workbook-workspace" :tabindex="activeKind === 'workbook' ? 0 : -1" data-workspace="workbook" :class="['workspace-tab', { active: activeKind === 'workbook' }]" @click="emit('switch', 'workbook')"><FileSpreadsheet :size="16" /><span>{{ t('workspace.workbook') }}</span></Button>
        <Button id="document-tab" type="text" role="tab" :aria-selected="activeKind === 'document'" aria-controls="document-workspace" :tabindex="activeKind === 'document' ? 0 : -1" data-workspace="document" :class="['workspace-tab', { active: activeKind === 'document' }]" @click="emit('switch', 'document')"><FileText :size="16" /><span>{{ t('workspace.document') }}</span></Button>
      </div>
    </nav>
    <Select
      :value="locale"
      class="workspace-locale"
      :aria-label="t('app.language')"
      :options="localeOptions"
      :popup-match-select-width="false"
      @change="emit('locale', $event as OfficeLocale)"
    />
    <div class="file-context">
      <Tooltip :title="t('workspace.rename-current')"><Button type="text" class="file-context-name" :title="fileName" :aria-label="t('workspace.rename-current')" @click="emit('action', 'rename')">{{ fileName }}</Button></Tooltip>
      <span :class="['file-context-status', { modified, busy: exportTask || parserStatus === 'loading' }]" role="status" :title="statusText">{{ statusText }}</span>
    </div>
    <div class="header-actions">
      <Tooltip :title="t('workspace.open-shortcut')"><Button :title="t('workspace.open-office-file')" :aria-label="t('workspace.open-office-file')" @click="emit('action', 'open')"><FolderOpen :size="16" /><span>{{ t('workspace.open') }}</span></Button></Tooltip>
      <Tooltip :title="exportTask ? `${exportTask.fileName}: ${exportTask.message}` : t('workspace.save-shortcut')"><Button type="primary" :title="t('workspace.save-office-file')" :aria-label="t('workspace.save-office-file')" :loading="Boolean(exportTask)" :disabled="Boolean(exportTask)" @mousedown.prevent @click="emit('action', 'save')"><Save v-if="!exportTask" :size="16" /><span>{{ exportTask ? t('workspace.processing') : t('workspace.save') }}</span></Button></Tooltip>
    </div>
  </header>
</template>
