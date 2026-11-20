import { defineStore } from 'pinia'
import { computed, reactive, watch } from 'vue'
import { createBlankDocument, createBlankWorkbook, type CellRange, type OfficeKind, type DocumentModel, type EditorDocumentState } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { applyExcelWorkbookView, type ExcelWorkbookView } from '../services/excelViewAdapter'
import { useWorkbookStore } from './workbookStore'
import { useDocumentStore } from './documentStore'
import { useParserStore } from './parserStore'
import { useHistoryStore } from './historyStore'

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
  watch(activeId, (id) => history.activate(id), { immediate: true })
  const state: EditorDocumentState = reactive({
    get activeKind() { return session.activeKind },
    set activeKind(kind: OfficeKind) { session.activeKind = kind },
    get workbook() { return workbook.model },
    set workbook(model) { workbook.replace(model) },
    get document() { return document.model },
    set document(model) { document.replace(model) },
    get fileName() { return session.activeKind === 'workbook' ? workbook.model.name : document.model.name },
    set fileName(name: string) {
      if (session.activeKind === 'workbook') workbook.replace({ ...workbook.model, name })
      else document.replace({ ...document.model, name })
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
    workbook.replace({ ...workbook.model, modified: history.isDirty(workbook.model.id!) })
    document.replace({ ...document.model, modified: history.isDirty(document.model.id!) })
  }
  function replaceWorkbookFromView(view: ExcelWorkbookView): void {
    const before = workbook.model
    const after = applyExcelWorkbookView(before, view)
    after.id = before.id
    after.modified = true
    history.execute(before.id!, '编辑工作簿', () => workbook.replace(after), () => workbook.replace(before))
  }
  function replaceDocumentFromView(model: DocumentModel): void {
    const before = document.model
    const after = cloneModel(model)
    after.id = before.id
    after.modified = true
    history.execute(before.id!, '编辑文档', () => document.replace(after), () => document.replace(before))
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
  async function importFile(file: File): Promise<void> {
    const imported = await parser.parse(file)
    if (!imported) return
    imported.model.id ??= crypto.randomUUID()
    if (imported.kind === 'workbook') workbook.replace(imported.model)
    else document.replace(imported.model)
    session.activeKind = imported.kind
    history.activate(imported.model.id)
    history.clear(imported.model.id)
  }
  function markSaved(): void { history.markSaved(activeId.value); refreshDirty() }
  return {
    state, workbook, document, parser, history,
    activeSheet: computed(() => workbook.model.sheets.find((sheet) => sheet.id === workbook.model.activeSheetId) ?? workbook.model.sheets[0]),
    canUndo: computed(() => history.canUndo), canRedo: computed(() => history.canRedo),
    setActiveKind: (kind: OfficeKind) => { session.activeKind = kind },
    setSelection: (range: CellRange) => { session.selectedRange = range },
    replaceWorkbookFromView, replaceDocumentFromView, undo, redo, newWorkbook, newDocument, importFile, markSaved,
  }
})
