import { defineStore } from 'pinia'
import { computed, reactive, watch } from 'vue'
import { createBlankDocument, createBlankWorkbook, type CellRange, type OfficeKind, type DocumentModel, type EditorDocumentState } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { applyExcelWorkbookView, toExcelWorkbookView, type ExcelWorkbookView } from '../services/excelViewAdapter'
import { equivalentModelData } from '../platform/commands/model-equivalence'
import { useWorkbookStore } from './workbookStore'
import { useDocumentStore } from './documentStore'
import { useParserStore } from './parserStore'
import { useHistoryStore, type HistorySavepoint } from './historyStore'
import { createDocumentSnapshotHistory } from '../features/document/services/document-history-cache.ts'

export interface EditorSavepoint extends HistorySavepoint { readonly kind: OfficeKind }

/** Session facade: orchestration only; models, parser tasks and per-unit history live in separate stores. */
export const useEditorStore = defineStore('editor-session', () => {
  const workbook = useWorkbookStore()
  const document = useDocumentStore()
  const parser = useParserStore()
  const history = useHistoryStore()
  const session = reactive({ activeKind: 'workbook' as OfficeKind, selectedRange: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 } })
  workbook.model.id ??= crypto.randomUUID()
  document.model.id ??= crypto.randomUUID()
  const activeId = computed(() => session.activeKind === 'workbook' ? workbook.model.id! : document.model.id!)
  watch(activeId, (id) => history.activate(id), { immediate: true, flush: 'sync' })
  const state: EditorDocumentState = reactive({
    get activeKind() { return session.activeKind },
    set activeKind(kind: OfficeKind) { session.activeKind = kind },
    get workbook() { return workbook.model },
    set workbook(model) { workbook.replace(model) },
    get document() { return document.model },
    set document(model) { document.replace(model) },
    get fileName() { return session.activeKind === 'workbook' ? workbook.model.name : document.model.name },
    set fileName(name: string) {
      renameActiveFile(name)
    },
    get parserStatus() { return parser.status },
    set parserStatus(value) { parser.status = value },
    get parserMessage() { return parser.message },
    set parserMessage(value: string) { parser.message = value },
    get diagnostics() { return parser.diagnostics },
    set diagnostics(value) { parser.diagnostics = value },
    get selectedRange() { return session.selectedRange },
    set selectedRange(value: CellRange) { session.selectedRange = value },
  })
  function refreshDirty(): void {
    const workbookDirty = history.isDirty(workbook.model.id!)
    const documentDirty = history.isDirty(document.model.id!)
    if (workbook.model.modified !== workbookDirty) workbook.replace({ ...workbook.model, modified: workbookDirty })
    if (document.model.modified !== documentDirty) document.replace({ ...document.model, modified: documentDirty })
  }
  function replaceWorkbookFromView(view: ExcelWorkbookView, numberFormatLocale?: string): boolean {
    const before = cloneModel(workbook.model)
    // Viewport defaults (e.g. 100 visible rows) must not rewrite untouched imported geometry.
    if (equivalentModelData(toExcelWorkbookView(before), view, ['id', 'activeSheetId'])) {
      if (view.activeSheetId && view.activeSheetId !== before.activeSheetId && before.sheets.some((sheet) => sheet.id === view.activeSheetId)) {
        workbook.replace({ ...before, activeSheetId: view.activeSheetId })
      }
      return false
    }
    const after = cloneModel(applyExcelWorkbookView(before, view, numberFormatLocale))
    after.id = before.id
    if (!after.sheets.some((sheet) => sheet.id === after.activeSheetId)) after.activeSheetId = after.sheets[0]?.id ?? ''
    if (equivalentModelData(before, after, ['modified', 'activeSheetId'])) {
      workbook.replace(after)
      return false
    }
    after.modified = true
    let firstApplication = true
    history.execute(before.id!, '编辑工作簿', () => {
      // The initial gesture can activate a newly inserted sheet; a later redo keeps the user's current view.
      replaceWorkbookContent(after, !firstApplication)
      firstApplication = false
    }, () => replaceWorkbookContent(before))
    refreshDirty()
    return true
  }
  function replaceDocumentFromView(model: DocumentModel): boolean {
    const before = cloneModel(document.model)
    const after = cloneModel(model)
    after.id = before.id
    if (equivalentModelData(before, after, ['modified'])) return false
    after.modified = true
    const callbacks = createDocumentSnapshotHistory(before, after, value => document.replace(value))
    history.execute(before.id!, '编辑文档', callbacks.redo, callbacks.undo)
    refreshDirty()
    return true
  }
  function replaceWorkbookContent(model: typeof workbook.model, preserveActiveSheet = true): void {
    const activeSheetId = preserveActiveSheet && workbook.model.id === model.id && model.sheets.some((sheet) => sheet.id === workbook.model.activeSheetId)
      ? workbook.model.activeSheetId : model.activeSheetId
    workbook.replace({ ...model, activeSheetId })
  }
  function renameActiveFile(name: string): void {
    if (session.activeKind === 'workbook') {
      if (workbook.model.name === name) return
      const before = cloneModel(workbook.model)
      const after = { ...before, name, modified: true }
      history.execute(before.id!, '重命名工作簿', () => replaceWorkbookContent(after), () => replaceWorkbookContent(before))
    } else {
      if (document.model.name === name) return
      const before = cloneModel(document.model)
      const after = { ...before, name, modified: true }
      const callbacks = createDocumentSnapshotHistory(before, after, value => document.replace(value))
      history.execute(before.id!, '重命名文档', callbacks.redo, callbacks.undo)
    }
    refreshDirty()
  }
  function undo(): string | undefined { const label = history.undo(); refreshDirty(); return label }
  function redo(): string | undefined { const label = history.redo(); refreshDirty(); return label }
  function newWorkbook(): void {
    parser.cancel()
    const next = createBlankWorkbook()
    next.id = crypto.randomUUID()
    workbook.replace(next)
    session.activeKind = 'workbook'
    history.activate(next.id)
  }
  function newDocument(): void {
    parser.cancel()
    const next = createBlankDocument()
    next.id = crypto.randomUUID()
    document.replace(next)
    session.activeKind = 'document'
    history.activate(next.id)
  }
  async function importFile(file: File): Promise<boolean> {
    const imported = await parser.parse(file)
    if (!imported) return false
    imported.model.id ??= crypto.randomUUID()
    if (imported.kind === 'workbook') workbook.replace(imported.model)
    else document.replace(imported.model)
    session.activeKind = imported.kind
    history.activate(imported.model.id)
    history.clear(imported.model.id)
    refreshDirty()
    return true
  }
  function captureSavepoint(kind: OfficeKind = session.activeKind): EditorSavepoint {
    const model = kind === 'workbook' ? workbook.model : document.model
    return { kind, ...history.captureSavepoint(model.id!) }
  }
  function markSaved(savepoint: EditorSavepoint = captureSavepoint()): boolean {
    const marked = history.markSaved(savepoint.unitId, savepoint)
    if (marked) refreshDirty()
    return marked
  }
  return {
    state, workbook, document, parser, history,
    activeSheet: computed(() => workbook.model.sheets.find((sheet) => sheet.id === workbook.model.activeSheetId) ?? workbook.model.sheets[0]),
    canUndo: computed(() => history.canUndo), canRedo: computed(() => history.canRedo),
    setActiveKind: (kind: OfficeKind) => { session.activeKind = kind },
    setSelection: (range: CellRange) => { session.selectedRange = range },
    replaceWorkbookFromView, replaceDocumentFromView, undo, redo, newWorkbook, newDocument, importFile, captureSavepoint, markSaved,
  }
})
