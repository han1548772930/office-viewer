import { createPinia } from 'pinia'
import type { Pinia } from 'pinia'
import type { DocumentModel, OfficeKind, PivotFieldArea, PivotFieldFilter, PivotTableOptions, WorkbookModel } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { insertDocumentText, replaceDocumentRange } from '../features/document/services/document-facade'
import { useEditorStore } from '../stores/editorStore'
import { applyCellValue, documentFromInput, workbookFromInput } from './model-utils'
import { EventHub, type ApiEventMap, type ApiRange, type Disposable } from './events'
import {
  createPivotTable,
  refreshPivotTable,
  removePivotFields,
  setPivotField,
  setPivotFieldSetting,
  setPivotLabelFilter,
  setPivotOptions,
  setPivotSort,
  setPivotValueFilter,
  type PivotTableAnchorInput,
  type PivotTableSourceInput,
} from '../features/pivot/services/pivot-engine'

export interface OfficeStudioOptions {
  workbook?: Partial<WorkbookModel> & Record<string, unknown>
  document?: Partial<DocumentModel> & Record<string, unknown>
}

export type OfficeCommandHandler<P = unknown, R = unknown> = (params: P) => R

export class OfficeStudio {
  readonly pinia: Pinia
  readonly editor: ReturnType<typeof useEditorStore>
  readonly events = new EventHub<ApiEventMap>()
  private readonly workbooks = new Map<string, WorkbookModel>()
  private readonly documents = new Map<string, DocumentModel>()
  private readonly commands = new Map<string, OfficeCommandHandler<never, unknown>>()
  private _activeWorkbookId: string
  private _activeDocumentId: string
  private selection: ApiRange | null = null
  private disposed = false

  constructor(options: OfficeStudioOptions = {}) {
    this.pinia = createPinia()
    this.editor = useEditorStore(this.pinia)
    const workbook = workbookFromInput(options.workbook ?? this.editor.workbook.model)
    const document = documentFromInput(options.document ?? this.editor.document.model)
    this._activeWorkbookId = workbook.id!
    this._activeDocumentId = document.id!
    this.workbooks.set(workbook.id!, workbook)
    this.documents.set(document.id!, document)
    this.editor.workbook.replace(workbook)
    this.editor.document.replace(document)
    this.editor.setActiveKind('workbook')
    this.editor.history.activate(workbook.id!)
    this.selection = {
      startRow: 0,
      startColumn: 0,
      endRow: 0,
      endColumn: 0,
      sheetId: workbook.activeSheetId,
      unitId: workbook.id!,
    }
    this.registerBuiltinCommands()
    this.events.emit('LifeCycleChanged', { stage: 'ready' })
  }

  get activeWorkbookId(): string { return this._activeWorkbookId }
  get activeDocumentId(): string { return this._activeDocumentId }
  get isDisposed(): boolean { return this.disposed }

  getWorkbookModel(id: string): WorkbookModel | undefined {
    return id === this._activeWorkbookId ? this.editor.workbook.model : this.workbooks.get(id)
  }

  getDocumentModel(id: string): DocumentModel | undefined {
    return id === this._activeDocumentId ? this.editor.document.model : this.documents.get(id)
  }

  createWorkbook(input: Partial<WorkbookModel> & Record<string, unknown> = {}, makeCurrent = true): WorkbookModel {
    this.assertAlive()
    const model = workbookFromInput(input)
    if (this.workbooks.has(model.id!)) throw new Error(`Workbook id already exists: ${model.id}`)
    this.workbooks.set(model.id!, model)
    if (makeCurrent) this.activateWorkbook(model.id!)
    this.events.emit('WorkbookCreated', { workbookId: model.id! })
    return model
  }

  createDocument(input: Partial<DocumentModel> & Record<string, unknown> = {}, makeCurrent = true): DocumentModel {
    this.assertAlive()
    const model = documentFromInput(input)
    if (this.documents.has(model.id!)) throw new Error(`Document id already exists: ${model.id}`)
    this.documents.set(model.id!, model)
    if (makeCurrent) this.activateDocument(model.id!)
    this.events.emit('DocCreated', { unitId: model.id!, documentId: model.id! })
    return model
  }

  activateWorkbook(id: string): void {
    this.assertAlive()
    const model = this.getWorkbookModel(id)
    if (!model) throw new Error(`Workbook not found: ${id}`)
    if (this._activeWorkbookId !== id) this.workbooks.set(this._activeWorkbookId, this.editor.workbook.model)
    this._activeWorkbookId = id
    this.workbooks.set(id, model)
    this.editor.workbook.replace(model)
    this.editor.setActiveKind('workbook')
    this.editor.history.activate(id)
  }

  activateDocument(id: string): void {
    this.assertAlive()
    const model = this.getDocumentModel(id)
    if (!model) throw new Error(`Document not found: ${id}`)
    if (this._activeDocumentId !== id) this.documents.set(this._activeDocumentId, this.editor.document.model)
    this._activeDocumentId = id
    this.documents.set(id, model)
    this.editor.document.replace(model)
    this.editor.setActiveKind('document')
    this.editor.history.activate(id)
  }

  setActiveKind(kind: OfficeKind): void {
    this.assertAlive()
    this.editor.setActiveKind(kind)
    this.editor.history.activate(kind === 'workbook' ? this._activeWorkbookId : this._activeDocumentId)
  }

  getWorkbook(id: string): WorkbookModel | undefined {
    return this.getWorkbookModel(id)
  }

  getDocument(id: string): DocumentModel | undefined {
    return this.getDocumentModel(id)
  }

  getSelection(): ApiRange | null {
    return this.selection ? { ...this.selection } : null
  }

  setSelection(unitId: string, sheetId: string, range: ApiRange): void {
    this.assertAlive()
    const workbook = this.getWorkbookModel(unitId)
    if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
    if (!workbook.sheets.some((sheet) => sheet.id === sheetId)) throw new Error(`Worksheet not found: ${sheetId}`)
    const next: ApiRange = {
      startRow: range.startRow,
      startColumn: range.startColumn,
      endRow: range.endRow,
      endColumn: range.endColumn,
      sheetId,
      unitId,
    }
    this.selection = next
    this.activateWorkbook(unitId)
    this.editor.setSelection({
      startRow: next.startRow,
      startCol: next.startColumn,
      endRow: next.endRow,
      endCol: next.endColumn,
    })
    this.events.emit('SelectionChanged', { workbookId: unitId, sheetId, range: { ...next } })
  }

  mutateWorkbook(
    id: string,
    commandId: string,
    label: string,
    update: (draft: WorkbookModel) => void,
    params: unknown = {},
  ): boolean {
    const event: ApiEventMap['BeforeCommandExecute'] = {
      id: commandId,
      params,
      unitId: id,
      kind: 'workbook',
    }
    this.events.emit('BeforeCommandExecute', event)
    if (event.cancel) throw new Error(`Command was canceled: ${commandId}`)
    const changed = this.commitWorkbook(id, label, update)
    if (changed) this.events.emit('CommandExecuted', { id: commandId, params, unitId: id, kind: 'workbook' })
    return changed
  }

  mutateDocument(
    id: string,
    commandId: string,
    label: string,
    update: (draft: DocumentModel) => void,
    params: unknown = {},
  ): boolean {
    const event: ApiEventMap['BeforeCommandExecute'] = {
      id: commandId,
      params,
      unitId: id,
      kind: 'document',
    }
    this.events.emit('BeforeCommandExecute', event)
    if (event.cancel) throw new Error(`Command was canceled: ${commandId}`)
    const changed = this.commitDocument(id, label, update)
    if (changed) this.events.emit('CommandExecuted', { id: commandId, params, unitId: id, kind: 'document' })
    return changed
  }

  commitWorkbook(id: string, label: string, update: (draft: WorkbookModel) => void): boolean {
    this.assertAlive()
    const current = this.getWorkbookModel(id)
    if (!current) throw new Error(`Workbook not found: ${id}`)
    const before = cloneModel(current)
    const after = cloneModel(current)
    update(after)
    after.id = id
    if (JSON.stringify(before) === JSON.stringify(after)) return false
    after.modified = true
    this.editor.history.execute(id, label, () => this.replaceWorkbookModel(id, after), () => this.replaceWorkbookModel(id, before))
    return true
  }

  commitDocument(id: string, label: string, update: (draft: DocumentModel) => void): boolean {
    this.assertAlive()
    const current = this.getDocumentModel(id)
    if (!current) throw new Error(`Document not found: ${id}`)
    const before = cloneModel(current)
    const after = cloneModel(current)
    update(after)
    after.id = id
    if (JSON.stringify(before) === JSON.stringify(after)) return false
    after.modified = true
    this.editor.history.execute(id, label, () => this.replaceDocumentModel(id, after), () => this.replaceDocumentModel(id, before))
    return true
  }

  executeCommandSync<P, R>(id: string, params: P): R {
    this.assertAlive()
    const handler = this.commands.get(id)
    if (!handler) throw new Error(`Unknown command: ${id}`)
    const event: ApiEventMap['BeforeCommandExecute'] = {
      id,
      params,
      unitId: commandUnitId(params),
      kind: commandKind(params),
    }
    this.events.emit('BeforeCommandExecute', event)
    if (event.cancel) throw new Error(`Command was canceled: ${id}`)
    const result = (handler as OfficeCommandHandler<P, R>)(params)
    this.events.emit('CommandExecuted', { ...event })
    return result
  }

  executeCommand<P = Record<string, unknown>, R = boolean>(id: string, params?: P): Promise<R> {
    return Promise.resolve().then(() => this.executeCommandSync<P | undefined, R>(id, params))
  }

  registerCommand<P, R>(id: string, handler: OfficeCommandHandler<P, R>): Disposable {
    this.assertAlive()
    if (!id || this.commands.has(id)) throw new Error(`Command already registered: ${id}`)
    this.commands.set(id, handler as OfficeCommandHandler<never, unknown>)
    let active = true
    return { dispose: () => { if (active) { active = false; this.commands.delete(id) } } }
  }

  undo(kind: OfficeKind, id: string): boolean {
    this.assertAlive()
    if (kind === 'workbook') this.activateWorkbook(id)
    else this.activateDocument(id)
    const event: ApiEventMap['BeforeUndo'] = { unitId: id, kind }
    this.events.emit('BeforeUndo', event)
    if (event.cancel) throw new Error('Undo was canceled')
    const label = this.editor.undo()
    if (!label) return false
    this.events.emit('Undo', { unitId: id, kind, label })
    return true
  }

  redo(kind: OfficeKind, id: string): boolean {
    this.assertAlive()
    if (kind === 'workbook') this.activateWorkbook(id)
    else this.activateDocument(id)
    const event: ApiEventMap['BeforeRedo'] = { unitId: id, kind }
    this.events.emit('BeforeRedo', event)
    if (event.cancel) throw new Error('Redo was canceled')
    const label = this.editor.redo()
    if (!label) return false
    this.events.emit('Redo', { unitId: id, kind, label })
    return true
  }

  markSaved(kind: OfficeKind, id: string): void {
    this.assertAlive()
    if (kind === 'workbook') {
      const model = this.getWorkbookModel(id)
      if (!model) throw new Error(`Workbook not found: ${id}`)
      this.editor.history.markSaved(id)
      this.replaceWorkbookModel(id, { ...model, modified: false })
    } else {
      const model = this.getDocumentModel(id)
      if (!model) throw new Error(`Document not found: ${id}`)
      this.editor.history.markSaved(id)
      this.replaceDocumentModel(id, { ...model, modified: false })
    }
  }

  disposeUnit(id: string): boolean {
    this.assertAlive()
    if (this.workbooks.has(id)) {
      this.workbooks.delete(id)
      this.editor.history.clear(id)
      this.events.emit('WorkbookDisposed', { workbookId: id })
      if (id === this._activeWorkbookId) {
        const next = this.createWorkbook({}, false)
        this.activateWorkbook(next.id!)
      }
      return true
    }
    if (this.documents.has(id)) {
      this.documents.delete(id)
      this.editor.history.clear(id)
      this.events.emit('DocDisposed', { unitId: id, documentId: id })
      if (id === this._activeDocumentId) {
        const next = this.createDocument({}, false)
        this.activateDocument(next.id!)
      }
      return true
    }
    return false
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.editor.parser.cancel()
    this.events.emit('LifeCycleChanged', { stage: 'disposed' })
    this.events.clear()
    this.commands.clear()
    this.workbooks.clear()
    this.documents.clear()
  }

  private replaceWorkbookModel(id: string, model: WorkbookModel): void {
    this.workbooks.set(id, model)
    if (id === this._activeWorkbookId) this.editor.workbook.replace(model)
  }

  private replaceDocumentModel(id: string, model: DocumentModel): void {
    this.documents.set(id, model)
    if (id === this._activeDocumentId) this.editor.document.replace(model)
  }

  private registerBuiltinCommands(): void {
    this.registerCommand<{ unitId: string; name: string }, boolean>('sheet.command.set-workbook-name', ({ unitId, name }) => this.commitWorkbook(unitId, 'Rename workbook', (model) => { model.name = name }))
    this.registerCommand<{ unitId: string; sheetId: string; name: string }, boolean>('sheet.command.set-worksheet-name', ({ unitId, sheetId, name }) => this.commitWorkbook(unitId, 'Rename worksheet', (model) => {
      const sheet = model.sheets.find((item) => item.id === sheetId)
      if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
      sheet.name = name
    }))
    this.registerCommand<{ unitId: string; sheetId: string; range: { startRow: number; startColumn: number; endRow: number; endColumn: number }; values: Array<Array<unknown>> }, boolean>('sheet.command.set-range-values', ({ unitId, sheetId, range, values }) => this.commitWorkbook(unitId, 'Set range values', (model) => {
      const sheet = model.sheets.find((item) => item.id === sheetId)
      if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
      for (let row = range.startRow; row <= range.endRow; row++) for (let col = range.startColumn; col <= range.endColumn; col++) {
        const cell = values[row - range.startRow]?.[col - range.startColumn]
        if (cell === undefined) continue
        const key = `${row}:${col}`
        sheet.cells[key] = applyCellValue(sheet.cells[key], row, col, cell as never)
      }
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; source: PivotTableSourceInput; anchor: PivotTableAnchorInput }, boolean>('sheet.command.add-pivot-table', ({ unitId, pivotTableId, source, anchor }) => this.commitPivotChange(unitId, pivotTableId, anchor.subUnitId, 'sheet.command.add-pivot-table', 'Add pivot table', (model) => {
      if (model.pivotTables?.some((pivot) => pivot.id === pivotTableId)) throw new Error(`Pivot table already exists: ${pivotTableId}`)
      createPivotTable(model, source, anchor, pivotTableId)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string }, boolean>('sheet.command.refresh-pivot-table', ({ unitId, pivotTableId }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.refresh-pivot-table', 'Refresh pivot table', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      refreshPivotTable(model, pivot)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string | number; area: PivotFieldArea; index?: number }, boolean>('sheet.command.set-pivot-field', ({ unitId, pivotTableId, fieldId, area, index }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-field', 'Set pivot field', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      setPivotField(model, pivot, fieldId, area, index)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldIds: string[] }, boolean>('sheet.command.remove-pivot-field', ({ unitId, pivotTableId, fieldIds }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.remove-pivot-field', 'Remove pivot field', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      removePivotFields(model, pivot, fieldIds)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; options: Partial<PivotTableOptions> }, boolean>('sheet.command.set-pivot-options', ({ unitId, pivotTableId, options }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-options', 'Set pivot options', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      setPivotOptions(model, pivot, options)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; direction: 'ascending' | 'descending' }, boolean>('sheet.command.set-pivot-label-sort', ({ unitId, pivotTableId, fieldId, direction }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-label-sort', 'Sort pivot labels', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      setPivotSort(model, pivot, fieldId, direction)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; selectedItems: string[] | null }, boolean>('sheet.command.set-pivot-label-filter', ({ unitId, pivotTableId, fieldId, selectedItems }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-label-filter', 'Filter pivot labels', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      setPivotLabelFilter(model, pivot, fieldId, selectedItems)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; setting: { displayName?: string; format?: string; subtotalType?: 'sum' | 'count' | 'average' | 'min' | 'max' } }, boolean>('sheet.command.set-pivot-field-setting', ({ unitId, pivotTableId, fieldId, setting }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-field-setting', 'Set pivot field setting', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      setPivotFieldSetting(model, pivot, fieldId, setting)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; filter: PivotFieldFilter | null }, boolean>('sheet.command.set-pivot-value-filter', ({ unitId, pivotTableId, fieldId, filter }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-value-filter', 'Filter pivot values', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      setPivotValueFilter(model, pivot, fieldId, filter)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; position: 'row' | 'column' }, boolean>('sheet.command.update-pivot-value-position', ({ unitId, pivotTableId, position }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.update-pivot-value-position', 'Update pivot value position', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      setPivotOptions(model, pivot, { valuePosition: position })
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string }, boolean>('sheet.command.remove-pivot-table', ({ unitId, pivotTableId }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.remove-pivot-table', 'Remove pivot table', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      const target = model.sheets.find((sheet) => sheet.id === pivot.anchor.subUnitId)
      if (target) for (const key of pivot.generatedCellKeys ?? []) if (target.cells[key]?.metadata?.pivotTableId === pivot.id) delete target.cells[key]
      model.pivotTables = (model.pivotTables ?? []).filter((candidate) => candidate.id !== pivotTableId)
    }))
    this.registerCommand<{ unitId: string; name: string }, boolean>('doc.command.set-name', ({ unitId, name }) => this.commitDocument(unitId, 'Rename document', (model) => { model.name = name }))
    this.registerCommand<{ unitId: string; index: number; text: string }, boolean>('doc.command.insert-text', ({ unitId, index, text }) => this.commitDocument(unitId, 'Insert text', (model) => Object.assign(model, insertDocumentText(model, index, text))))
    this.registerCommand<{ unitId: string; startOffset: number; endOffset: number; text?: string }, boolean>('doc.command.replace-range', ({ unitId, startOffset, endOffset, text = '' }) => this.commitDocument(unitId, 'Edit text range', (model) => Object.assign(model, replaceDocumentRange(model, startOffset, endOffset, text))))
  }

  private commitPivotChange(
    unitId: string,
    pivotTableId: string,
    subUnitId: string | undefined,
    commandId: string,
    label: string,
    update: (model: WorkbookModel) => void,
  ): boolean {
    const current = this.getWorkbookModel(unitId)
    if (!current) throw new Error(`Workbook not found: ${unitId}`)
    const currentPivot = current.pivotTables?.find((pivot) => pivot.id === pivotTableId)
    const event: ApiEventMap['BeforePivotTableChange'] = {
      unitId,
      subUnitId: subUnitId ?? currentPivot?.anchor.subUnitId ?? '',
      pivotTableId,
      commandId,
    }
    this.events.emit('BeforePivotTableChange', event)
    if (event.cancel) throw new Error(`Pivot table change was canceled: ${commandId}`)
    const changed = this.commitWorkbook(unitId, label, update)
    if (!changed) return false
    const next = this.getWorkbookModel(unitId)?.pivotTables?.find((pivot) => pivot.id === pivotTableId)
    if (next) {
      const payload = { unitId, subUnitId: next.anchor.subUnitId, pivotTableId, revision: next.revision }
      this.events.emit('PivotTableChanged', payload)
      this.events.emit('PivotTableRendered', { ...payload, outputRange: next.outputRange })
    }
    return true
  }

  private assertAlive(): void {
    if (this.disposed) throw new Error('OfficeStudio has been disposed')
  }
}

function commandUnitId(params: unknown): string | undefined {
  return params && typeof params === 'object' && 'unitId' in params && typeof params.unitId === 'string' ? params.unitId : undefined
}

function commandKind(params: unknown): OfficeKind | undefined {
  const unitId = commandUnitId(params)
  return unitId?.startsWith('document-') ? 'document' : unitId ? 'workbook' : undefined
}

function requirePivot(model: WorkbookModel, pivotTableId: string) {
  const pivot = model.pivotTables?.find((candidate) => candidate.id === pivotTableId)
  if (!pivot) throw new Error(`Pivot table not found: ${pivotTableId}`)
  return pivot
}

