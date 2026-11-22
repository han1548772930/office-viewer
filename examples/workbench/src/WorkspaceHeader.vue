<script setup lang="ts">
import { Button, Dropdown, Tooltip, type MenuProps } from 'antdv-next'
import { computed, h, nextTick, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import { ChevronDown, FolderOpen, Save } from 'lucide-vue-next'
import { theme } from 'antdv-next'
import type { ExportTask, WorkspaceKind } from './workspace-file-operations'
import { isRibbonVariant, RIBBON_VARIANTS, supportedOfficeLocales, WORKBENCH_ACCENTS, WORKBENCH_CHROME, WORKBENCH_ZOOMS, type OfficeLocale, type RibbonVariant, type WorkbenchAccent, type WorkbenchChrome, type WorkbenchZoom } from '@office-studio/office-studio'

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
  ribbonVariant: RibbonVariant
  direction: 'ltr' | 'rtl'
  dark: boolean
  region: OfficeLocale | 'auto'
  accent: WorkbenchAccent
  chrome: WorkbenchChrome
  zoom: WorkbenchZoom
}>()
const emit = defineEmits<{
  action: [action: WorkspaceHeaderAction]
  switch: [kind: WorkspaceKind]
  locale: [locale: OfficeLocale]
  ribbon: [variant: RibbonVariant]
  direction: [direction: 'ltr' | 'rtl']
  dark: [value: boolean]
  region: [value: OfficeLocale | 'auto']
  accent: [value: WorkbenchAccent]
  chrome: [value: WorkbenchChrome]
  zoom: [value: WorkbenchZoom]
  reset: []
}>()
const tabs = ref<HTMLElement | null>(null)
const fileMenuOpen = ref(false)
const settingsOpen = ref(false)
const { t } = useI18n()
const { token } = theme.useToken()
const headerTheme = computed(() => ({ '--shell-accent': token.value.colorPrimary, '--shell-text': token.value.colorText, '--shell-muted': token.value.colorTextSecondary, '--shell-bg': token.value.colorBgContainer, '--shell-hover': token.value.colorFillQuaternary, '--shell-border': token.value.colorBorderSecondary }))
const languageLabels: Record<OfficeLocale, string> = {
  'ar-SA': 'العربية', 'ca-ES': 'Català', 'de-DE': 'Deutsch', 'en-US': 'English', 'es-ES': 'Español', 'fa-IR': 'فارسی',
  'fr-FR': 'Français', 'id-ID': 'Bahasa Indonesia', 'it-IT': 'Italiano', 'ja-JP': '日本語', 'ko-KR': '한국어',
  'pl-PL': 'Polski', 'pt-BR': 'Português (Brasil)', 'ru-RU': 'Русский', 'sk-SK': 'Slovenčina', 'vi-VN': 'Tiếng Việt',
  'zh-CN': '简体中文', 'zh-HK': '繁體中文（香港）', 'zh-TW': '繁體中文',
}
const localeOptions = supportedOfficeLocales.map(value => ({ value, label: languageLabels[value] }))
const regionNames = computed(() => new Intl.DisplayNames([props.locale], { type: 'region' }))
function regionName(tag: OfficeLocale | 'auto') {
  if (tag === 'auto') return t('workspace.follow-language')
  if (tag === 'ca-ES') return t('workspace.catalonia')
  return regionNames.value.of(tag.slice(-2)) ?? tag
}
function setting(label: string, value: string) {
  return h('span', { class: 'workspace-setting-label' }, [h('span', label), h('bdi', { class: 'workspace-setting-value', style: { color: token.value.colorTextSecondary }, title: value }, value)])
}
const settingsItems = computed<MenuProps['items']>(() => [
  { key: 'language', label: setting(t('app.language'), languageLabels[props.locale]), children: localeOptions.map(({ value, label }) => ({ key: `locale:${value}`, label: h('bdi', { lang: value }, label) })) },
  { key: 'region', label: setting(t('workspace.region'), regionName(props.region)), children: (['auto', ...supportedOfficeLocales] as const).map(tag => ({ key: `region:${tag}`, label: regionName(tag) })) },
  { key: 'direction', label: setting(t('workspace.direction'), props.direction.toUpperCase()), children: [{ key: 'ltr', label: 'LTR' }, { key: 'rtl', label: 'RTL' }] },
  { key: 'accent', label: setting(t('workspace.theme'), t(`workspace.theme-${props.accent}`)), children: Object.entries(WORKBENCH_ACCENTS).map(([name, color]) => ({ key: `theme:${name}`, label: t(`workspace.theme-${name}`), icon: h('span', { class: 'workspace-theme-swatch', style: { backgroundColor: color }, 'aria-hidden': true }) })) },
  { key: 'appearance', label: setting(t('workspace.appearance'), t(props.dark ? 'workspace.dark' : 'workspace.light')), children: [{ key: 'light', label: t('workspace.light') }, { key: 'dark', label: t('workspace.dark') }] },
  { key: 'ribbon', label: setting(t('workspace.ribbon'), t(`workspace.${props.ribbonVariant}-ribbon`)), children: RIBBON_VARIANTS.map(key => ({ key, label: t(`workspace.${key}-ribbon`) })) },
  { key: 'chrome', label: setting(t('workspace.chrome'), t(`workspace.${props.chrome}`)), children: WORKBENCH_CHROME.map(mode => ({ key: `chrome:${mode}`, label: t(`workspace.${mode}`) })) },
  { key: 'zoom', label: setting(t('workspace.zoom'), `${props.zoom}%`), children: WORKBENCH_ZOOMS.map(value => ({ key: `zoom:${value}`, label: `${value}%` })) },
  { type: 'divider' }, { key: 'reset', label: t('workspace.reset-preferences') },
])
function onSettingClick(info: { key: string | number }) {
  const key = String(info.key)
  const locale = localeOptions.find(option => `locale:${option.value}` === key)?.value
  const region = supportedOfficeLocales.find(tag => `region:${tag}` === key)
  const accent = (Object.keys(WORKBENCH_ACCENTS) as WorkbenchAccent[]).find(name => `theme:${name}` === key)
  const chrome = WORKBENCH_CHROME.find(mode => `chrome:${mode}` === key)
  const zoom = WORKBENCH_ZOOMS.find(value => `zoom:${value}` === key)
  if (locale) emit('locale', locale)
  else if (region || key === 'region:auto') emit('region', region ?? 'auto')
  else if (accent) emit('accent', accent)
  else if (chrome) emit('chrome', chrome)
  else if (zoom) emit('zoom', zoom)
  else if (key === 'reset') emit('reset')
  else if (key === 'ltr' || key === 'rtl') emit('direction', key)
  else if (isRibbonVariant(key)) emit('ribbon', key)
  else if (key === 'light' || key === 'dark') emit('dark', key === 'dark')
  settingsOpen.value = false
}
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
function onProductClick(event: MouseEvent, kind: WorkspaceKind) {
  if (event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return
  event.preventDefault()
  emit('switch', kind)
}
</script>

<template>
  <header class="app-header workspace-header" :aria-label="t('workspace.app-label')" :style="headerTheme">
    <div class="workspace-primary">
      <span class="workspace-brand" role="img" :aria-label="t('workspace.brand-label')">Office Studio</span>
      <nav ref="tabs" class="workspace-tabs" :aria-label="t('workspace.tabs-label')" @keydown="onTabsKeydown">
        <div class="workspace-tablist" role="tablist" :aria-label="t('workspace.document-types')">
          <Button id="workbook-tab" type="text" role="tab" href="#sheets" :aria-current="activeKind === 'workbook' ? 'page' : undefined" :aria-selected="activeKind === 'workbook'" aria-controls="workbook-workspace" :tabindex="activeKind === 'workbook' ? 0 : -1" data-workspace="workbook" :class="['workspace-tab', { active: activeKind === 'workbook' }]" @click="onProductClick($event, 'workbook')">{{ t('workspace.workbook') }}</Button>
          <Button id="document-tab" type="text" role="tab" href="#docs" :aria-current="activeKind === 'document' ? 'page' : undefined" :aria-selected="activeKind === 'document'" aria-controls="document-workspace" :tabindex="activeKind === 'document' ? 0 : -1" data-workspace="document" :class="['workspace-tab', { active: activeKind === 'document' }]" @click="onProductClick($event, 'document')">{{ t('workspace.document') }}</Button>
        </div>
      </nav>
    </div>
    <Dropdown v-model:open="fileMenuOpen" :menu="{ items: menuItems }" :trigger="['click']" :auto-focus="true" placement="bottomLeft" @menu-click="onMenuClick">
      <Button type="text" class="workspace-file-menu" :aria-label="t('workspace.file-menu-label')" aria-haspopup="menu" :aria-expanded="fileMenuOpen">{{ t('workspace.file-menu') }}<ChevronDown :size="13" /></Button>
    </Dropdown>
    <div class="file-context">
      <Tooltip :title="t('workspace.rename-current')"><Button type="text" class="file-context-name" :title="fileName" :aria-label="t('workspace.rename-current')" @click="emit('action', 'rename')">{{ fileName }}</Button></Tooltip>
      <span :class="['file-context-status', { modified, busy: exportTask || parserStatus === 'loading' }]" role="status" :title="statusText">{{ statusText }}</span>
    </div>
    <div class="header-actions">
      <Tooltip :title="t('workspace.open-shortcut')"><Button :title="t('workspace.open-office-file')" :aria-label="t('workspace.open-office-file')" @click="emit('action', 'open')"><FolderOpen :size="16" /><span>{{ t('workspace.open') }}</span></Button></Tooltip>
      <Tooltip :title="exportTask ? `${exportTask.fileName}: ${exportTask.message}` : t('workspace.save-shortcut')"><Button type="primary" :title="t('workspace.save-office-file')" :aria-label="t('workspace.save-office-file')" :loading="Boolean(exportTask)" :disabled="Boolean(exportTask)" @mousedown.prevent @click="emit('action', 'save')"><Save v-if="!exportTask" :size="16" /><span>{{ exportTask ? t('workspace.processing') : t('workspace.save') }}</span></Button></Tooltip>
    </div>
    <Dropdown v-model:open="settingsOpen" :classes="{ root: 'workspace-settings-menu' }" :trigger="['click']" :placement="direction === 'rtl' ? 'bottomLeft' : 'bottomRight'" :auto-focus="true" :menu="{ items: settingsItems, selectable: true, selectedKeys: [`locale:${locale}`, `region:${region}`, `theme:${accent}`, `chrome:${chrome}`, `zoom:${zoom}`, direction, dark ? 'dark' : 'light', ribbonVariant], triggerSubMenuAction: 'click' }" @menu-click="onSettingClick">
      <Button type="text" class="workspace-settings-trigger" aria-haspopup="menu" :aria-expanded="settingsOpen">{{ t('workspace.settings') }}<ChevronDown :size="13" /></Button>
    </Dropdown>
  </header>
</template>
