<script setup lang="ts">
import { Button } from 'antdv-next'
import { computed, onMounted, onUnmounted, ref } from 'vue'
import { BookOpenText, Download, FileSpreadsheet, FileText, FolderOpen, Plus, Redo2, Save, Undo2 } from 'lucide-vue-next'
import DocumentWorkbench from './components/DocumentWorkbench.vue'
import ExcelWorkbench from './components/ExcelWorkbench.vue'
import type { DocumentModel } from './domain/model'
import { toExcelWorkbookView, type ExcelWorkbookView } from './services/excelViewAdapter'
import { useEditorStore } from './stores/editorStore'

const editor = useEditorStore()
const fileInput = ref<HTMLInputElement | null>(null)
const notice = ref<{ type: 'success' | 'error' | 'info'; message: string } | null>(null)
let noticeTimer: ReturnType<typeof setTimeout> | undefined
const excelView = computed(() => toExcelWorkbookView(editor.state.workbook))
const activeFileName = computed(() => editor.state.fileName || (editor.state.activeKind === 'workbook' ? editor.state.workbook.name : editor.state.document.name))

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

function exportProject(): void {
  const payload = editor.state.activeKind === 'workbook' ? editor.state.workbook : editor.state.document
  const kind = editor.state.activeKind === 'workbook' ? 'workbook' : 'document'
  const blob = new Blob([JSON.stringify({ schemaVersion: 1, kind, data: payload }, null, 2)], { type: 'application/json' })
  const url = URL.createObjectURL(blob)
  const anchor = window.document.createElement('a')
  anchor.href = url
  anchor.download = `${payload.name.replace(/\.[^.]+$/, '')}.office-viewer.json`
  anchor.click()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
  showNotice('success', '编辑数据已导出为 Office Viewer JSON')
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
    exportProject()
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
        <Button title="新建工作簿" aria-label="新建工作簿" @click="editor.newWorkbook"><Plus :size="16" /><span>新建</span></Button>
        <Button title="打开 Office 文件" aria-label="打开 Office 文件" @click="requestImport"><FolderOpen :size="16" /><span>打开</span></Button>
        <span class="header-divider" />
        <Button title="撤销 Ctrl+Z" aria-label="撤销" :disabled="!editor.canUndo.value" @click="editor.undo"><Undo2 :size="16" /></Button>
        <Button title="重做 Ctrl+Y" aria-label="重做" :disabled="!editor.canRedo.value" @click="editor.redo"><Redo2 :size="16" /></Button>
        <Button title="导出编辑数据" aria-label="导出编辑数据" @click="exportProject"><Download :size="16" /><span>导出数据</span></Button>
        <Button type="primary" title="保存编辑数据" aria-label="保存编辑数据" @click="exportProject"><Save :size="16" /><span>保存</span></Button>
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
        class="workbench"
        @update:model-value="onExcelChange"
        @import="importFile"
        @save="onExcelChange"
        @undo="editor.undo"
        @redo="editor.redo"
        @notify="(payload) => showNotice(payload.type === 'warning' ? 'info' : payload.type, payload.message)"
      />
      <DocumentWorkbench
        v-else
        :model-value="editor.state.document"
        :file-name="editor.state.document.name"
        class="workbench"
        @update:model-value="onDocumentChange"
        @import-file="importFile"
        @save="onDocumentChange($event.model)"
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
