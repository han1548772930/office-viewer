import { computed, onMounted, onUnmounted, ref, watch } from 'vue'
import {
  cloneModel,
  downloadOfficeBytes,
  exportOfficeModel,
  officeDownloadName,
  toExcelWorkbookView,
  useEditorStore,
  type DocumentModel,
  type ExcelWorkbookView,
  type WorkbookModel,
} from '@office-studio/office-studio'
import { studio, officeAPI } from './runtime'
import { createWorkspaceFileOperations, workspaceFileNameError, type ExportTask, type FileAction, type NoticeType, type WorkspaceKind } from './workspace-file-operations'

export function useWorkspaceController() {
  const editor = useEditorStore(studio.pinia)
  const notice = ref<{ type: NoticeType; message: string } | null>(null)
  const exportTask = ref<ExportTask | null>(null)
  const diagnosticsOpen = ref(false)
  const renameOpen = ref(false)
  const renameValue = ref('')
  const renameError = ref('')
  let renameTarget: { kind: WorkspaceKind; unitId: string } | undefined
  const pendingNewKind = ref<WorkspaceKind | null>(null)
  let noticeTimer: ReturnType<typeof setTimeout> | undefined

  const activeModel = computed(() => editor.state.activeKind === 'workbook' ? editor.state.workbook : editor.state.document)
  const activeFileName = computed(() => activeModel.value.name)
  const excelView = computed(() => toExcelWorkbookView(editor.state.workbook))
  const documentApi = computed(() => editor.state.document.id ? officeAPI.getDocument(editor.state.document.id) ?? undefined : undefined)
  const modified = computed(() => activeModel.value.modified)

  function showNotice(type: NoticeType, message: string): void {
    notice.value = { type, message }
    if (noticeTimer) clearTimeout(noticeTimer)
    noticeTimer = setTimeout(() => { notice.value = null }, type === 'error' ? 8000 : 5000)
  }

  const files = createWorkspaceFileOperations<WorkbookModel | DocumentModel, File>({
    importFile: (file) => editor.importFile(file),
    cancelImport: () => editor.parser.cancel(),
    capture: (kind = editor.state.activeKind) => {
      const model = cloneModel(kind === 'workbook' ? editor.state.workbook : editor.state.document)
      return { kind, model, fileName: officeDownloadName(model, kind), savepoint: editor.captureSavepoint(kind) }
    },
    write: (snapshot, options) => exportOfficeModel(snapshot.kind, snapshot.model, options),
    download: (bytes, snapshot) => downloadOfficeBytes(bytes, snapshot.fileName, snapshot.kind),
    markSaved: (savepoint) => editor.markSaved(savepoint),
    onExportTask: (task) => { exportTask.value = task },
    notify: showNotice,
  })

  function cancelImport(): void {
    if (editor.state.parserStatus !== 'loading') return
    files.cancelImport()
    showNotice('info', '文件导入已取消')
  }

  function switchWorkspace(kind: WorkspaceKind): void {
    if (kind === editor.state.activeKind) return
    cancelImport()
    officeAPI.setActiveKind(kind)
  }

  function createNew(kind: WorkspaceKind): void {
    cancelImport()
    if (kind === 'workbook') officeAPI.createWorkbook({ name: '未命名工作簿' })
    else officeAPI.createDocument({ name: '未命名文档' })
    pendingNewKind.value = null
  }

  function requestNew(kind: WorkspaceKind): void {
    const replaced = kind === 'workbook' ? editor.state.workbook : editor.state.document
    if (replaced.modified) pendingNewKind.value = kind
    else createNew(kind)
  }

  function requestRename(): void {
    const savepoint = editor.captureSavepoint()
    renameTarget = { kind: savepoint.kind, unitId: savepoint.unitId }
    renameValue.value = activeFileName.value
    renameError.value = ''
    renameOpen.value = true
  }

  function renameFile(): void {
    const name = renameValue.value.trim()
    const error = workspaceFileNameError(name)
    if (error) { renameError.value = error; return }
    try {
      const unit = renameTarget?.kind === 'workbook'
        ? officeAPI.getWorkbook(renameTarget.unitId)
        : renameTarget ? officeAPI.getDocument(renameTarget.unitId) : undefined
      if (!unit) throw new Error('当前文件已切换，请重新打开重命名窗口')
      unit.setName(name)
      renameOpen.value = false
    } catch (error: unknown) {
      renameError.value = error instanceof Error ? error.message : '文件重命名失败'
    }
  }

  function onExcelChange(view: ExcelWorkbookView): void { editor.replaceWorkbookFromView(view) }
  function onDocumentChange(model: DocumentModel): void { editor.replaceDocumentFromView(model) }

  function saveWorkbook(view: ExcelWorkbookView, action: FileAction = 'save'): void {
    onExcelChange(view)
    void files.exportFile(action, 'workbook')
  }

  function saveDocument(payload: { model: DocumentModel }, action: FileAction = 'save'): void {
    onDocumentChange(payload.model)
    void files.exportFile(action, 'document')
  }

  function onDragOver(event: DragEvent): void {
    if (!event.dataTransfer?.types.includes('Files')) return
    event.preventDefault()
    event.dataTransfer.dropEffect = 'copy'
  }

  function onDrop(event: DragEvent): void {
    if (event.defaultPrevented) return
    const file = event.dataTransfer?.files[0]
    if (!file) return
    event.preventDefault()
    void files.importFile(file)
  }

  function onPageHide(event: PageTransitionEvent): void { if (!event.persisted) files.dispose() }
  watch(activeFileName, (name) => { document.title = `${name} · Office Studio` }, { immediate: true })
  onMounted(() => window.addEventListener('pagehide', onPageHide))
  onUnmounted(() => {
    files.dispose()
    window.removeEventListener('pagehide', onPageHide)
    if (noticeTimer) clearTimeout(noticeTimer)
  })

  return {
    editor, studio, notice, activeFileName, excelView, documentApi, modified,
    exportTask, diagnosticsOpen, renameOpen, renameValue, renameError, pendingNewKind,
    files, showNotice, cancelImport, switchWorkspace, requestNew, createNew,
    requestRename, renameFile, onExcelChange, onDocumentChange, saveWorkbook,
    saveDocument, onDragOver, onDrop,
  }
}
