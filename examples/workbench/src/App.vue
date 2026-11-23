<script setup lang="ts">
import { Button, ConfigProvider, Form, FormItem, Input, Modal, theme, Upload } from 'antdv-next'
import { computed, onMounted, onUnmounted, provide, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { DocumentWorkbench, ExcelWorkbench, OfficeDirectionKey, OfficeLocaleKey, OfficeRegionKey, OfficeTranslatorKey, WORKBENCH_ACCENTS, WORKBENCH_CONTROL_THEME, WorkbenchDisplayKey, type OfficeLocale } from '@office-studio/office-studio'
import { useDisplayPreferences } from './display-preferences'
import WorkspaceHeader, { type WorkspaceHeaderAction } from './WorkspaceHeader.vue'
import WorkspaceFeedback from './WorkspaceFeedback.vue'
import WorkspaceDiagnostics from './WorkspaceDiagnostics.vue'
import { useWorkspaceController } from './useWorkspaceController'
import type { FileAction } from './workspace-file-operations'

const {
  editor, studio, notice, activeFileName, excelView, documentApi, modified,
  exportTask, diagnosticsOpen, renameOpen, renameValue, renameError, pendingNewKind,
  files, showNotice, undoContent, redoContent, cancelImport, switchWorkspace, requestNew, createNew,
  requestRename, renameFile, onExcelChange, onDocumentChange, saveWorkbook,
  saveDocument, onDragOver, onDrop,
} = useWorkspaceController()

interface WorkbenchHandle { save: () => void; focus?: () => void }
const workbookRef = ref<WorkbenchHandle | null>(null)
const documentRef = ref<WorkbenchHandle | null>(null)
const { locale, t, te } = useI18n()
const { p: preferences, display, reset } = useDisplayPreferences(localStorage)
const regional = studio.regionalSettings
regional.setLocale(locale.value)
function setRegion(value: OfficeLocale | 'auto') {
  if (value === 'auto') regional.followLanguage()
  else regional.setRegion(value)
}
setRegion(preferences.value.region)
watch(regional.regionPreference, value => { preferences.value.region = value }, { flush: 'sync' })
watch(regional.locale, value => { locale.value = value }, { flush: 'sync' })
const ribbonVariant = computed(() => preferences.value.ribbon)
const direction = computed(() => preferences.value.direction)
watch(direction, value => { document.documentElement.dir = value }, { immediate: true })
function resetPreferences() { regional.setLocale('zh-CN'); regional.followLanguage(); reset() }
const importTrigger = ref<{ $el: HTMLElement } | null>(null)
function requestImport(): void { importTrigger.value?.$el.click() }
function importSelectedFile(file: File): false { void files.importFile(file); return false }
provide(OfficeLocaleKey, regional.locale)
provide(OfficeDirectionKey, direction)
provide(OfficeRegionKey, regional.region)
provide(WorkbenchDisplayKey, display)
provide(OfficeTranslatorKey, (key, fallback) => {
  const normalized = key.replace(/^(sheet|doc|core)\./, '')
  return te(key) ? t(key) : te(normalized) ? t(normalized) : fallback
})
watch(locale, value => {
  localStorage.setItem('office-studio.locale', value)
  document.documentElement.lang = value
}, { flush: 'sync', immediate: true })
let requestedAction: FileAction = 'save'

function requestExport(action: FileAction = 'save'): void {
  if (exportTask.value) return
  const workbench = editor.state.activeKind === 'workbook' ? workbookRef.value : documentRef.value
  requestedAction = action
  try {
    if (workbench?.save) workbench.save()
    else void files.exportFile(action)
  } finally { requestedAction = 'save' }
}

function onHeaderAction(action: WorkspaceHeaderAction): void {
  if (action === 'new-workbook') requestNew('workbook')
  else if (action === 'new-document') requestNew('document')
  else if (action === 'open') requestImport()
  else if (action === 'save' || action === 'export') requestExport(action)
  else if (action === 'rename') requestRename()
  else if (action === 'diagnostics') diagnosticsOpen.value = true
  else if (action === 'cancel-import') cancelImport()
  else if (action === 'cancel-export') files.cancelExport()
}

function onGlobalKeydown(event: KeyboardEvent): void {
  if (event.defaultPrevented || event.isComposing || event.altKey || !(event.ctrlKey || event.metaKey)) return
  const key = event.key.toLowerCase()
  const target = event.target
  if (target instanceof HTMLElement && target.closest('[role="dialog"], .ant-modal')) return
  if (key === 'o') { event.preventDefault(); requestImport(); return }
  if (key === 's') { event.preventDefault(); requestExport(); return }
  if (target instanceof HTMLElement && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName))) return
  if (key === 'z' && !event.shiftKey) { event.preventDefault(); undoContent() }
  else if (key === 'y' || (key === 'z' && event.shiftKey)) { event.preventDefault(); redoContent() }
}

onMounted(() => window.addEventListener('keydown', onGlobalKeydown))
onUnmounted(() => window.removeEventListener('keydown', onGlobalKeydown))
</script>

<template>
  <ConfigProvider :direction="direction" :theme="{ ...WORKBENCH_CONTROL_THEME, token: { colorPrimary: WORKBENCH_ACCENTS[preferences.accent] }, algorithm: preferences.dark ? theme.darkAlgorithm : theme.defaultAlgorithm }">
  <div class="app-shell" @dragover="onDragOver" @drop="onDrop">
    <WorkspaceHeader
      :active-kind="editor.state.activeKind" :file-name="activeFileName" :modified="modified"
      :parser-status="editor.state.parserStatus" :parser-message="editor.state.parserMessage"
      :export-task="exportTask" :diagnostic-count="editor.state.diagnostics.length" :locale="locale as OfficeLocale"
      :ribbon-variant="ribbonVariant" :direction="direction" :dark="preferences.dark" :accent="preferences.accent" :region="preferences.region" :chrome="preferences.chrome" :zoom="preferences.zoom"
      @ribbon="preferences.ribbon = $event" @direction="preferences.direction = $event" @dark="preferences.dark = $event" @accent="preferences.accent = $event" @region="setRegion" @chrome="preferences.chrome = $event" @zoom="preferences.zoom = $event" @reset="resetPreferences"
      @action="onHeaderAction" @switch="switchWorkspace" @locale="regional.setLocale($event)"
    />
    <Upload class="file-picker-host" :styles="{ root: { display: 'none' } }" aria-hidden="true" accept=".doc,.docx,.odt,.xls,.xlsx,.xlsb,.ods" :show-upload-list="false" :before-upload="importSelectedFile"><Button ref="importTrigger" class="file-picker-trigger" aria-hidden="true" tabindex="-1" /></Upload>

    <main :id="editor.state.activeKind === 'workbook' ? 'workbook-workspace' : 'document-workspace'" class="workspace" role="tabpanel" :aria-labelledby="editor.state.activeKind === 'workbook' ? 'workbook-tab' : 'document-tab'" :aria-label="editor.state.activeKind === 'workbook' ? t('workspace.workbook-area') : t('workspace.document-area')">
      <WorkspaceFeedback :parser-status="editor.state.parserStatus" :parser-message="editor.state.parserMessage" :notice="notice" @cancel-import="cancelImport" @dismiss-parser-error="editor.state.parserStatus = 'idle'" />
      <KeepAlive :max="2">
        <ExcelWorkbench
          v-if="editor.state.activeKind === 'workbook'" :key="'workbook:' + editor.state.workbook.id" ref="workbookRef"
          :model-value="excelView" :ribbon-variant="ribbonVariant" :studio="studio" :show-header="false" :can-undo="editor.canUndo" :can-redo="editor.canRedo" class="workbench"
          @update:model-value="onExcelChange" @import="files.importFile" @save="(view) => saveWorkbook(view, requestedAction)" @export="(view) => saveWorkbook(view, 'export')"
          @undo="undoContent('workbook', $event?.workbookId)" @redo="redoContent('workbook', $event?.workbookId)" @notify="(payload) => showNotice(payload.type, payload.message)"
        />
        <DocumentWorkbench
          v-else :key="'document:' + editor.state.document.id" ref="documentRef"
          :model-value="editor.state.document" :ribbon-variant="ribbonVariant" :file-name="editor.state.document.name" :document-api="documentApi" :show-header="false"
          :can-undo="editor.canUndo" :can-redo="editor.canRedo" class="workbench"
          @update:model-value="onDocumentChange" @import-file="files.importFile" @save="(payload) => saveDocument(payload, requestedAction)"
          @undo="undoContent('document', $event)" @redo="redoContent('document', $event)" @notify="(message) => showNotice('info', message)"
        />
      </KeepAlive>
    </main>

    <Modal v-model:open="renameOpen" :title="t('workspace.rename-title')" :ok-text="t('workspace.confirm')" :cancel-text="t('workspace.cancel')" @ok="renameFile">
      <Form :model="{ name: renameValue }" layout="vertical" @finish="renameFile">
        <FormItem name="name" :label="t('workspace.file-name')" :validate-status="renameError ? 'error' : undefined" :help="renameError || undefined">
          <Input v-model:value="renameValue" :aria-label="t('workspace.file-name')" :maxlength="255" />
        </FormItem>
      </Form>
    </Modal>
    <Modal :open="pendingNewKind !== null" :title="pendingNewKind === 'workbook' ? t('workspace.new-workbook') : t('workspace.new-document')" :ok-text="t('workspace.continue-create')" :cancel-text="t('workspace.return-editing')" @ok="pendingNewKind && createNew(pendingNewKind)" @cancel="pendingNewKind = null">
      <p>{{ t('workspace.pending-new', { kind: t(pendingNewKind === 'workbook' ? 'workspace.pending-workbook' : 'workspace.pending-document') }) }}</p>
    </Modal>
    <Modal v-model:open="diagnosticsOpen" :title="t('workspace.diagnostics-title')" :footer="null">
      <WorkspaceDiagnostics :diagnostics="editor.state.diagnostics" />
    </Modal>
  </div>
  </ConfigProvider>
</template>
