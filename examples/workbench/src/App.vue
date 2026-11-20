<script setup lang="ts">
import { Button } from 'antdv-next'
import { computed, onMounted, onUnmounted, ref } from 'vue'
import { BookOpenText, Download, FileSpreadsheet, FileText, FolderOpen, Plus, Redo2, Save, Undo2 } from 'lucide-vue-next'
import {
  cloneModel,
  DocumentWorkbench,
  downloadOfficeBytes,
  ExcelWorkbench,
  exportOfficeModel,
  officeDownloadName,
  toExcelWorkbookView,
  useEditorStore,
  type DocumentModel,
  type ExcelWorkbookView,
} from '@office-studio/office-studio'
import { studio, univerAPI } from './runtime'

const editor = useEditorStore()
const fileInput = ref<HTMLInputElement | null>(null)
const notice = ref<{ type: 'success' | 'error' | 'info'; message: string } | null>(null)
let noticeTimer: ReturnType<typeof setTimeout> | undefined
const excelView = computed(() => toExcelWorkbookView(editor.state.workbook))
const activeFileName = computed(() => editor.state.fileName || (editor.state.activeKind === 'workbook' ? editor.state.workbook.name : editor.state.document.name))

function newWorkbook(): void {
  univerAPI.createWorkbook({ name: '未命名工作簿' })
}

function showNotice(type: 'success' | 'error' | 'info', message: string): void {
  notice.value = { type, message }
  if (noticeTimer) clearTimeout(noticeTimer)
  noticeTimer = setTimeout(() => { notice.value = null }, 4000)
}

function requestImport(): void {
  fileInput.value?.click()
}

async function importFile(file?: File): Promise<void> {
  if (!file) return
  try {
    await editor.importFile(file)
    showNotice('success', `${file.name} 已解析并载入`)
  } catch (error: unknown) {
    showNotice('error', error instanceof Error ? error.message : '文件解析失败')
  }
}

function onFileChange(event: Event): void {
  const input = event.target as HTMLInputElement
  void importFile(input.files?.[0])
  input.value = ''
}

function onExcelChange(view: ExcelWorkbookView): void {
  editor.replaceWorkbookFromView(view)
}

function onDocumentChange(document: DocumentModel): void {
  editor.replaceDocumentFromView(document)
  editor.state.fileName = document.name
}

function showExportResult(format: string, action: string, warnings: { message: string }[]): void {
  if (warnings.length === 0) {
    showNotice('success', `${format} ${action}`)
    return
  }
  const detail = warnings.length > 1 ? `另有 ${warnings.length - 1} 项提示` : warnings[0].message
  showNotice('info', `${format} ${action}，有 ${warnings.length} 项内容未能完整保留：${warnings.length > 1 ? `${warnings[0].message}；${detail}` : detail}`)
}

async function exportProject(): Promise<void> {
  const kind = editor.state.activeKind
  const payload = kind === 'workbook' ? editor.state.workbook : editor.state.document
  try {
    const snapshot = cloneModel(payload)
    const outcome = await exportOfficeModel(kind, snapshot)
    downloadOfficeBytes(outcome.bytes, officeDownloadName(payload, kind), kind)
    editor.markSaved()
    showExportResult(kind === 'workbook' ? 'XLSX' : 'DOCX', '文件已导出', outcome.warnings)
  } catch (error: unknown) {
    showNotice('error', error instanceof Error ? error.message : 'Office 文件导出失败')
  }
}

async function onExcelSave(view: ExcelWorkbookView): Promise<void> {
  editor.replaceWorkbookFromView(view)
  const snapshot = cloneModel(editor.state.workbook)
  try {
    const outcome = await exportOfficeModel('workbook', snapshot)
    downloadOfficeBytes(outcome.bytes, officeDownloadName(snapshot, 'workbook'), 'workbook')
    editor.markSaved()
    showExportResult('XLSX 工作簿', '已保存', outcome.warnings)
  } catch (error: unknown) {
    showNotice('error', error instanceof Error ? error.message : 'XLSX 导出失败')
  }
}

async function onDocumentSave(payload: { fileName: string; html: string; model: DocumentModel }): Promise<void> {
  onDocumentChange(payload.model)
  const snapshot = cloneModel(editor.state.document)
  try {
    const outcome = await exportOfficeModel('document', snapshot)
    downloadOfficeBytes(outcome.bytes, officeDownloadName(snapshot, 'document'), 'document')
    editor.markSaved()
    showExportResult('DOCX 文档', '已保存', outcome.warnings)
  } catch (error: unknown) {
    showNotice('error', error instanceof Error ? error.message : 'DOCX 导出失败')
  }
}

function onDrop(event: DragEvent): void {
  const file = event.dataTransfer?.files[0]
  if (file) void importFile(file)
}

function onGlobalKeydown(event: KeyboardEvent): void {
  const command = event.ctrlKey || event.metaKey
  if (!command) return
  if (event.key.toLowerCase() === 'o') {
    event.preventDefault()
    requestImport()
    return
  }
  if (event.key.toLowerCase() === 's') {
    event.preventDefault()
    void exportProject()
    return
  }
  const target = event.target
  const isTextEditing = target instanceof HTMLElement && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName))
  if (isTextEditing) return
  if (event.key.toLowerCase() === 'z' && !event.shiftKey) {
    event.preventDefault()
    editor.undo()
  } else if (event.key.toLowerCase() === 'y' || (event.key.toLowerCase() === 'z' && event.shiftKey)) {
    event.preventDefault()
    editor.redo()
  }
}

onMounted(() => window.addEventListener('keydown', onGlobalKeydown))
onUnmounted(() => {
  window.removeEventListener('keydown', onGlobalKeydown)
  if (noticeTimer) clearTimeout(noticeTimer)
})
</script>

<template>
  <div class="app-shell" @dragover.prevent @drop.prevent="onDrop">
    <header class="app-header">
      <a class="brand" href="#" aria-label="Office Studio 工作台" @click.prevent="editor.setActiveKind('workbook')">
        <span class="brand-mark"><BookOpenText :size="18" /></span>
        <span class="brand-name">Office Studio</span>
      </a>

      <nav class="workspace-tabs" aria-label="工作台">
        <button :class="['workspace-tab', { active: editor.state.activeKind === 'workbook' }]" @click="editor.setActiveKind('workbook')">
          <FileSpreadsheet :size="16" /> 表格
        </button>
        <button :class="['workspace-tab', { active: editor.state.activeKind === 'document' }]" @click="editor.setActiveKind('document')">
          <FileText :size="16" /> 文档
        </button>
      </nav>

      <div class="file-context">
        <span class="file-context-name" :title="activeFileName">{{ activeFileName }}</span>
        <span v-if="editor.state.parserStatus === 'loading'" class="file-context-status">正在解析</span>
        <span v-else-if="editor.state.activeKind === 'workbook' ? editor.state.workbook.modified : editor.state.document.modified" class="file-context-status modified">未保存</span>
      </div>

      <div class="header-actions">
        <Button title="新建工作簿" aria-label="新建工作簿" @click="newWorkbook"><Plus :size="16" /><span>新建</span></Button>
        <Button title="打开 Office 文件" aria-label="打开 Office 文件" @click="requestImport"><FolderOpen :size="16" /><span>打开</span></Button>
        <span class="header-divider" />
        <Button title="撤销 Ctrl+Z" aria-label="撤销" :disabled="!editor.canUndo" @click="editor.undo"><Undo2 :size="16" /></Button>
        <Button title="重做 Ctrl+Y" aria-label="重做" :disabled="!editor.canRedo" @click="editor.redo"><Redo2 :size="16" /></Button>
        <Button title="导出 Office 文件" aria-label="导出 Office 文件" @click="exportProject"><Download :size="16" /><span>导出</span></Button>
        <Button type="primary" title="保存 Office 文件" aria-label="保存 Office 文件" @click="exportProject"><Save :size="16" /><span>保存</span></Button>
      </div>
      <input ref="fileInput" class="file-input" type="file" accept=".doc,.docx,.odt,.xls,.xlsx,.xlsb,.ods" @change="onFileChange">
    </header>

    <main class="workspace" :aria-label="editor.state.activeKind === 'workbook' ? '表格编辑工作区' : '文档编辑工作区'">
      <div v-if="editor.state.parserStatus === 'error'" class="parser-error" role="alert">
        <span>{{ editor.state.parserMessage }}</span>
        <button aria-label="关闭错误提示" @click="editor.state.parserStatus = 'idle'">×</button>
      </div>
      <ExcelWorkbench
        v-if="editor.state.activeKind === 'workbook'"
        :model-value="excelView"
        :studio="studio"
        :can-undo="editor.canUndo"
        :can-redo="editor.canRedo"
        class="workbench"
        @update:model-value="onExcelChange"
        @import="importFile"
        @save="onExcelSave"
        @undo="editor.undo"
        @redo="editor.redo"
        @notify="(payload) => showNotice(payload.type === 'warning' ? 'info' : payload.type, payload.message)"
      />
      <DocumentWorkbench
        v-else
        :model-value="editor.state.document"
        :file-name="editor.state.document.name"
        :can-undo="editor.canUndo"
        :can-redo="editor.canRedo"
        class="workbench"
        @update:model-value="onDocumentChange"
        @import-file="importFile"
        @save="onDocumentSave"
        @undo="editor.undo"
        @redo="editor.redo"
      />
    </main>

    <div v-if="notice" :class="['notice', notice.type]" role="status">{{ notice.message }}</div>
    <footer class="app-footer">
      <span>{{ editor.state.activeKind === 'workbook' ? '工作簿' : '文档' }} · {{ editor.state.parserStatus === 'loading' ? editor.state.parserMessage : '就绪' }}</span>
      <span v-if="editor.state.diagnostics.length">{{ editor.state.diagnostics.length }} 条解析诊断</span>
      <span class="footer-spacer" />
      <span>Rust/WASM · duke-pages · rxls</span>
    </footer>
  </div>
</template>
