<script setup lang="ts">
import { Input, Modal } from 'antdv-next'
import { onMounted, onUnmounted, ref } from 'vue'
import { DocumentWorkbench, ExcelWorkbench } from '@office-studio/office-studio'
import WorkspaceHeader, { type WorkspaceHeaderAction } from './WorkspaceHeader.vue'
import { useWorkspaceController } from './useWorkspaceController'
import type { FileAction } from './workspace-file-operations'

const {
  editor, studio, setFileInput, notice, activeFileName, excelView, documentApi, modified,
  exportTask, diagnosticsOpen, renameOpen, renameValue, renameError, pendingNewKind,
  files, showNotice, requestImport, cancelImport, switchWorkspace, requestNew, createNew,
  requestRename, renameFile, onFileChange, onExcelChange, onDocumentChange, saveWorkbook,
  saveDocument, onDragOver, onDrop,
} = useWorkspaceController()

interface WorkbenchHandle { save: () => void; focus?: () => void }
const workbookRef = ref<WorkbenchHandle | null>(null)
const documentRef = ref<WorkbenchHandle | null>(null)
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
  if (key === 'z' && !event.shiftKey) { event.preventDefault(); editor.undo() }
  else if (key === 'y' || (key === 'z' && event.shiftKey)) { event.preventDefault(); editor.redo() }
}

onMounted(() => window.addEventListener('keydown', onGlobalKeydown))
onUnmounted(() => window.removeEventListener('keydown', onGlobalKeydown))
</script>

<template>
  <div class="app-shell" @dragover="onDragOver" @drop="onDrop">
    <WorkspaceHeader
      :active-kind="editor.state.activeKind" :file-name="activeFileName" :modified="modified"
      :parser-status="editor.state.parserStatus" :parser-message="editor.state.parserMessage"
      :export-task="exportTask" :diagnostic-count="editor.state.diagnostics.length"
      @action="onHeaderAction" @switch="switchWorkspace"
    />
    <input :ref="setFileInput" class="file-input" type="file" accept=".doc,.docx,.odt,.xls,.xlsx,.xlsb,.ods" aria-label="选择 Office 文件" @change="onFileChange">

    <main :id="editor.state.activeKind === 'workbook' ? 'workbook-workspace' : 'document-workspace'" class="workspace" role="tabpanel" :aria-labelledby="editor.state.activeKind === 'workbook' ? 'workbook-tab' : 'document-tab'" :aria-label="editor.state.activeKind === 'workbook' ? '表格编辑工作区' : '文档编辑工作区'">
      <div v-if="editor.state.parserStatus === 'loading'" class="workspace-task-strip" role="status" aria-live="polite">
        <span class="workspace-task-spinner" aria-hidden="true" /><span>{{ editor.state.parserMessage }}</span>
        <button type="button" @click="cancelImport">取消导入</button>
      </div>
      <div v-if="editor.state.parserStatus === 'error'" class="parser-error" role="alert">
        <span>{{ editor.state.parserMessage }}</span>
        <button type="button" aria-label="关闭错误提示" @click="editor.state.parserStatus = 'idle'">×</button>
      </div>
      <KeepAlive :max="2">
        <ExcelWorkbench
          v-if="editor.state.activeKind === 'workbook'" :key="'workbook:' + editor.state.workbook.id" ref="workbookRef"
          :model-value="excelView" :studio="studio" :show-header="false" :can-undo="editor.canUndo" :can-redo="editor.canRedo" class="workbench"
          @update:model-value="onExcelChange" @import="files.importFile" @save="(view) => saveWorkbook(view, requestedAction)" @export="(view) => saveWorkbook(view, 'export')"
          @undo="editor.undo" @redo="editor.redo" @notify="(payload) => showNotice(payload.type === 'warning' ? 'info' : payload.type, payload.message)"
        />
        <DocumentWorkbench
          v-else :key="'document:' + editor.state.document.id" ref="documentRef"
          :model-value="editor.state.document" :file-name="editor.state.document.name" :document-api="documentApi" :show-header="false"
          :can-undo="editor.canUndo" :can-redo="editor.canRedo" class="workbench"
          @update:model-value="onDocumentChange" @import-file="files.importFile" @save="(payload) => saveDocument(payload, requestedAction)"
          @undo="editor.undo" @redo="editor.redo" @notify="(message) => showNotice('info', message)"
        />
      </KeepAlive>
    </main>

    <div v-if="notice" :class="['notice', notice.type]" :role="notice.type === 'error' ? 'alert' : 'status'">{{ notice.message }}</div>

    <Modal v-model:open="renameOpen" title="重命名文件" ok-text="确定" cancel-text="取消" @ok="renameFile">
      <Input v-model:value="renameValue" aria-label="文件名称" :maxlength="255" :status="renameError ? 'error' : undefined" @press-enter="renameFile" />
      <p v-if="renameError" class="workspace-dialog-error" role="alert">{{ renameError }}</p>
    </Modal>
    <Modal :open="pendingNewKind !== null" :title="pendingNewKind === 'workbook' ? '新建工作簿' : '新建文档'" ok-text="继续新建" cancel-text="返回编辑" @ok="pendingNewKind && createNew(pendingNewKind)" @cancel="pendingNewKind = null">
      <p>当前{{ pendingNewKind === 'workbook' ? '工作簿' : '文档' }}还有未保存的更改。请先取消并保存需要保留的内容，再新建文件。</p>
    </Modal>
    <Modal v-model:open="diagnosticsOpen" title="解析诊断" :footer="null">
      <p v-if="!editor.state.diagnostics.length" class="workspace-diagnostics-empty">当前没有解析诊断。</p>
      <ul v-else class="workspace-diagnostics">
        <li v-for="(diagnostic, index) in editor.state.diagnostics" :key="diagnostic.code + ':' + index" :data-severity="diagnostic.severity">
          <strong>{{ diagnostic.severity === 'error' ? '错误' : diagnostic.severity === 'warning' ? '提示' : '信息' }} · {{ diagnostic.code }}</strong>
          <p>{{ diagnostic.message }}</p><details v-if="diagnostic.details"><summary>查看详情</summary><pre>{{ diagnostic.details }}</pre></details>
        </li>
      </ul>
    </Modal>
  </div>
</template>
