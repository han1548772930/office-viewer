import { createPinia } from 'pinia'
import type { Pinia } from 'pinia'
import { watch } from 'vue'
import type { ChartModel, CellRange, DocumentModel, OfficeKind, PivotAggregation, PivotDateGroupType, PivotDateGroupingMode, PivotFieldArea, PivotFieldFilter, PivotManualGroup, PivotNumberGroup, PivotTableOptions, SheetDataValidationRule, SheetDataValidationType, SheetFilterCriteria, SheetRangeProtectionRule, SheetSortCriterion, SlicerTimelineRange, SparklineSpec, WorkbookModel } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { createCell } from '../domain/model'
import { createSparklineSpec } from '../features/sparkline/sparkline-model.ts'
import { equivalentModelData } from '../platform/commands/model-equivalence'
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
  setPivotNumberGroup,
  setPivotDateGroupingMode,
  setPivotDateGroupType,
  setPivotExcelDateGroup,
  setPivotManualGroups,
  type PivotTableAnchorInput,
  type PivotTableSourceInput,
} from '../features/pivot/services/pivot-engine'
import { createChart, updateChart, type ChartConfig } from '../features/charts/chart-engine'
import { CONFIGURE_PIVOT_COMMAND, configurePivotFromPanel, type ConfigurePivotFromPanelInput } from '../features/pivot/services/pivot-panel-state'
import { clearSlicerFilter, createSlicer, removeSlicer, setSlicerItems, setSlicerTimelineRange, updateSlicerPresentation, type SlicerCreateInput, type SlicerPresentationPatch } from '../features/slicer/slicer-engine'
import { recalculateWorkbookFormulas } from '../features/formula/index.ts'
import { clearFilter, refreshAutoFilter, setAutoFilter, setFilterCriteria, sortRange, type SortRangeOptions } from '../features/sheet-sort-filter/sort-filter-engine'
import { CONFIGURE_SHEET_FILTER_COMMAND, configureFilterFromPanel, type ConfigureSheetFilterInput } from '../features/sheet-sort-filter/sheet-data-panel-state'
import { OfficeComputeClient } from '../compute/client.ts'
import { NumericIndexService } from '../features/sheet-compute/numeric-index-service.ts'
import { translateFormulaReferences } from '../features/formula/formula-references.ts'
import { applySheetDataValidationRule, clearSheetDataValidationRanges, clearSheetDataValidationTypes, removeSheetDataValidationRule, replaceSheetDataValidationRule, validateSheetDataValidationChanges } from '../features/sheet-validation/validation-engine.ts'
import { canEditProtectedRanges, removeSheetRangeProtectionRule, setSheetRangeProtectionRule } from '../features/sheet-protection/range-protection.ts'
import { assertDocumentPermission, permissionTargetsForChangedBlocks, permissionTargetsForRange, validateDocumentPermissionRules } from '../features/document/permissions/document-permissions.ts'
import {
  createDocumentContentControl,
  deleteRepeatingSectionItem,
  duplicateRepeatingSectionItem,
  removeDocumentContentControl,
  resetDocumentContentControlPlaceholder,
  selectDocumentContentControlItem,
  setDocumentContentControlChecked,
  setDocumentContentControlValue,
  updateDocumentContentControl,
  type ContentControlInput,
  type ContentControlPatch,
} from '../features/document/services/content-controls'

export interface OfficeStudioOptions {
  workbook?: Partial<WorkbookModel> & Record<string, unknown>
  document?: Partial<DocumentModel> & Record<string, unknown>
  computeClient?: OfficeComputeClient
  numericIndexMinRows?: number
  currentUserId?: string
  /** Local permission-rule manager. Defaults to currentUserId and is not server authorization. */
  permissionAdminUserId?: string
}

export type OfficeCommandHandler<P = unknown, R = unknown> = (params: P) => R

export class OfficeStudio {
  readonly pinia: Pinia
  readonly editor: ReturnType<typeof useEditorStore>
  readonly events = new EventHub<ApiEventMap>()
  readonly numericIndexService: NumericIndexService
  readonly currentUserId: string
  readonly permissionAdminUserId: string
  private readonly workbooks = new Map<string, WorkbookModel>()
  private readonly documents = new Map<string, DocumentModel>()
  private readonly commands = new Map<string, OfficeCommandHandler<never, unknown>>()
  private _activeWorkbookId: string
  private _activeDocumentId: string
  private selection: ApiRange | null = null
  private disposed = false
  private readonly workbookRevisions = new Map<string, number>()
  private readonly computeClient: OfficeComputeClient
  private readonly ownsComputeClient: boolean
  private readonly stopModelSync: Array<() => void> = []

  constructor(options: OfficeStudioOptions = {}) {
    this.pinia = createPinia()
    this.editor = useEditorStore(this.pinia)
    const workbook = workbookFromInput(options.workbook ?? this.editor.workbook.model)
    const document = documentFromInput(options.document ?? this.editor.document.model)
    this._activeWorkbookId = workbook.id!
    this._activeDocumentId = document.id!
    this.currentUserId = options.currentUserId?.trim() || 'local-user'
    this.permissionAdminUserId = options.permissionAdminUserId?.trim() || this.currentUserId
    this.workbooks.set(workbook.id!, workbook)
    this.workbookRevisions.set(workbook.id!, 0)
    this.ownsComputeClient = options.computeClient === undefined
    this.computeClient = options.computeClient ?? new OfficeComputeClient()
    this.numericIndexService = new NumericIndexService(this.computeClient, { minRows: options.numericIndexMinRows })
    this.numericIndexService.setDocumentRevision(workbook.id!, 0)
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
    this.stopModelSync.push(
      watch(() => this.editor.workbook.model, (next, previous) => {
        const id = next.id
        if (!id || this.disposed) return
        if (id !== this._activeWorkbookId) {
          this.workbooks.set(this._activeWorkbookId, cloneModel(previous))
          const created = !this.workbooks.has(id)
          this._activeWorkbookId = id
          this.workbooks.set(id, next)
          this.advanceWorkbookRevision(id)
          this.selection = { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0, sheetId: next.activeSheetId, unitId: id }
          if (created) this.events.emit('WorkbookCreated', { workbookId: id })
        } else {
          this.workbooks.set(id, next)
          if (!equivalentModelData(next, previous, ['modified', 'activeSheetId'])) this.advanceWorkbookRevision(id)
        }
      }, { flush: 'sync' }),
      watch(() => this.editor.document.model, (next, previous) => {
        const id = next.id
        if (!id || this.disposed) return
        if (id !== this._activeDocumentId) {
          this.documents.set(this._activeDocumentId, cloneModel(previous))
          const created = !this.documents.has(id)
          this._activeDocumentId = id
          this.documents.set(id, next)
          if (created) this.events.emit('DocCreated', { unitId: id, documentId: id })
        } else this.documents.set(id, next)
      }, { flush: 'sync' }),
    )
    this.events.emit('LifeCycleChanged', { stage: 'ready' })
  }

  get activeWorkbookId(): string { return this._activeWorkbookId }
  get activeDocumentId(): string { return this._activeDocumentId }
  get isDisposed(): boolean { return this.disposed }

  getWorkbookRevision(id: string): number {
    if (!this.getWorkbookModel(id)) throw new Error(`Workbook not found: ${id}`)
    return this.workbookRevisions.get(id) ?? 0
  }

  getWorkbookModel(id: string): WorkbookModel | undefined {
    if (this.disposed) return undefined
    return id === this._activeWorkbookId ? this.editor.workbook.model : this.workbooks.get(id)
  }

  assertCanEditRange(unitId: string, sheetId: string, range: CellRange): void {
    const workbook = this.getWorkbookModel(unitId)
    if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
    if (!canEditProtectedRanges(requireSheet(workbook, sheetId), [range], this.currentUserId)) {
      throw new Error('该区域受保护，当前用户没有编辑权限。')
    }
  }

  getDocumentModel(id: string): DocumentModel | undefined {
    if (this.disposed) return undefined
    return id === this._activeDocumentId ? this.editor.document.model : this.documents.get(id)
  }

  createWorkbook(input: Partial<WorkbookModel> & Record<string, unknown> = {}, makeCurrent = true): WorkbookModel {
    this.assertAlive()
    const model = workbookFromInput(input)
    if (this.workbooks.has(model.id!)) throw new Error(`Workbook id already exists: ${model.id}`)
    this.workbooks.set(model.id!, model)
    this.workbookRevisions.set(model.id!, 0)
    this.numericIndexService.setDocumentRevision(model.id!, 0)
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
    if (this._activeWorkbookId !== id && this.workbooks.has(this._activeWorkbookId)) this.workbooks.set(this._activeWorkbookId, this.editor.workbook.model)
    this._activeWorkbookId = id
    this.workbooks.set(id, model)
    this.editor.workbook.replace({ ...model, modified: this.editor.history.isDirty(id) })
    this.editor.setActiveKind('workbook')
    this.editor.history.activate(id)
  }

  activateDocument(id: string): void {
    this.assertAlive()
    const model = this.getDocumentModel(id)
    if (!model) throw new Error(`Document not found: ${id}`)
    if (this._activeDocumentId !== id && this.documents.has(this._activeDocumentId)) this.documents.set(this._activeDocumentId, this.editor.document.model)
    this._activeDocumentId = id
    this.documents.set(id, model)
    this.editor.document.replace({ ...model, modified: this.editor.history.isDirty(id) })
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
    const rangeEdit = rangeEditFromParams(params)
    if (rangeEdit) this.assertCanEditRange(id, rangeEdit.sheetId, rangeEdit.range)
    const changed = commandId.includes('.operation.')
      ? this.applyWorkbookOperation(id, update)
      : this.commitWorkbook(id, label, update)
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

  commitWorkbook(id: string, label: string, update: (draft: WorkbookModel) => void, forceRecalculation = false): boolean {
    this.assertAlive()
    const current = this.getWorkbookModel(id)
    if (!current) throw new Error(`Workbook not found: ${id}`)
    const before = cloneModel(current)
    const after = cloneModel(current)
    update(after)
    const recalculated = !forceRecalculation && before.calculationMode === 'manual' && after.calculationMode === 'manual'
      ? after
      : recalculateWorkbookFormulas(after, before).workbook
    for (const sheet of recalculated.sheets) if (sheet.autoFilter) refreshAutoFilter(sheet)
    recalculated.id = id
    if (JSON.stringify(before) === JSON.stringify(recalculated)) return false
    recalculated.modified = true
    let initial = true
    this.editor.history.execute(id, label, () => {
      if (initial) { this.replaceWorkbookModel(id, recalculated); initial = false }
      else this.replaceWorkbookHistorySnapshot(id, recalculated)
    }, () => this.replaceWorkbookHistorySnapshot(id, before))
    return true
  }

  commitDocument(id: string, label: string, update: (draft: DocumentModel) => void): boolean {
    return this.commitDocumentTransaction(id, label, update, false)
  }

  commitDocumentPermissionRules(id: string, rules: DocumentModel['permissionRules']): boolean {
    if (this.currentUserId !== this.permissionAdminUserId) throw new Error('Only the document permission administrator can update permission rules')
    return this.commitDocumentTransaction(id, 'Set document permission rules', (model) => {
      const next = rules ?? []
      validateDocumentPermissionRules(next)
      model.permissionRules = next.map(rule => ({ ...rule, userIds: [...rule.userIds] }))
    }, true)
  }

  private commitDocumentTransaction(id: string, label: string, update: (draft: DocumentModel) => void, permissionAdminCommand: boolean): boolean {
    this.assertAlive()
    const current = this.getDocumentModel(id)
    if (!current) throw new Error(`Document not found: ${id}`)
    const before = cloneModel(current)
    const after = cloneModel(current)
    update(after)
    if (!permissionAdminCommand) {
      assertDocumentPermission(current.permissionRules ?? [], { documentId: id, userId: this.currentUserId }, 'edit', permissionTargetsForChangedBlocks(before, after))
    }
    after.id = id
    if (JSON.stringify(before) === JSON.stringify(after)) return false
    after.modified = true
    this.editor.history.execute(id, label, () => this.replaceDocumentModel(id, after), () => this.replaceDocumentModel(id, before))
    return true
  }

  private commitDocumentWithRangePermission(
    id: string,
    label: string,
    startOffset: number,
    endOffset: number,
    update: (draft: DocumentModel) => void,
  ): boolean {
    const current = this.getDocumentModel(id)
    if (!current) throw new Error(`Document not found: ${id}`)
    assertDocumentPermission(current.permissionRules ?? [], { documentId: id, userId: this.currentUserId }, 'edit', permissionTargetsForRange(current, startOffset, endOffset))
    return this.commitDocument(id, label, update)
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
      this.computeClient.forgetDocument(id)
      this.workbookRevisions.delete(id)
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
    for (const stop of this.stopModelSync) stop()
    this.editor.parser.cancel()
    if (this.ownsComputeClient) this.computeClient.dispose()
    this.events.emit('LifeCycleChanged', { stage: 'disposed' })
    this.events.clear()
    this.commands.clear()
    this.workbooks.clear()
    this.documents.clear()
  }

  private replaceWorkbookModel(id: string, model: WorkbookModel): void {
    this.workbooks.set(id, model)
    if (id === this._activeWorkbookId) this.editor.workbook.replace(model)
    else this.advanceWorkbookRevision(id)
  }

  private applyWorkbookOperation(id: string, update: (draft: WorkbookModel) => void): boolean {
    this.assertAlive()
    const current = this.getWorkbookModel(id)
    if (!current) throw new Error(`Workbook not found: ${id}`)
    const next = cloneModel(current)
    update(next)
    next.id = id
    next.modified = current.modified
    if (JSON.stringify(current) === JSON.stringify(next)) return false
    this.replaceWorkbookModel(id, next)
    return true
  }

  private replaceWorkbookHistorySnapshot(id: string, snapshot: WorkbookModel): void {
    const current = this.getWorkbookModel(id)
    const next = cloneModel(snapshot)
    if (current && next.sheets.some((sheet) => sheet.id === current.activeSheetId)) next.activeSheetId = current.activeSheetId
    this.replaceWorkbookModel(id, next)
  }

  private advanceWorkbookRevision(id: string): void {
    const revision = (this.workbookRevisions.get(id) ?? 0) + 1
    this.workbookRevisions.set(id, revision)
    this.numericIndexService.setDocumentRevision(id, revision)
  }

  private replaceDocumentModel(id: string, model: DocumentModel): void {
    this.documents.set(id, model)
    if (id === this._activeDocumentId) this.editor.document.replace(model)
  }

  private registerBuiltinCommands(): void {
    this.registerCommand<{ unitId: string; mode: 'automatic' | 'manual' }, boolean>('sheet.command.set-calculation-mode', ({ unitId, mode }) => {
      if (mode !== 'automatic' && mode !== 'manual') throw new TypeError('Calculation mode must be automatic or manual')
      return this.commitWorkbook(unitId, 'Set calculation mode', (model) => { model.calculationMode = mode })
    })
    this.registerCommand<{ unitId: string }, boolean>('sheet.command.recalculate-workbook', ({ unitId }) =>
      this.commitWorkbook(unitId, 'Recalculate workbook', () => {}, true),
    )
    this.registerCommand<ConfigureSheetFilterInput, boolean>(CONFIGURE_SHEET_FILTER_COMMAND, (input) => {
      if (input.expectedRevision !== undefined && this.getWorkbookRevision(input.unitId) !== input.expectedRevision) throw new Error('数据已更改，请重新读取数据区域后再应用筛选。')
      return this.commitWorkbook(input.unitId, '配置工作表筛选', (model) => configureFilterFromPanel(model, input))
    })
    this.registerCommand<ConfigurePivotFromPanelInput, boolean>(CONFIGURE_PIVOT_COMMAND, (input) => this.commitPivotChange(
      input.unitId, input.pivotTableId, input.anchor.subUnitId, CONFIGURE_PIVOT_COMMAND,
      input.mode === 'create' ? '插入数据透视表' : '配置数据透视表',
      (draft) => { configurePivotFromPanel(draft, input) },
    ))
    this.registerCommand<{ unitId: string; name: string }, boolean>('sheet.command.set-workbook-name', ({ unitId, name }) => this.commitWorkbook(unitId, 'Rename workbook', (model) => { model.name = name }))
    this.registerCommand<{ unitId: string; sheetId: string; name: string }, boolean>('sheet.command.set-worksheet-name', ({ unitId, sheetId, name }) => this.commitWorkbook(unitId, 'Rename worksheet', (model) => {
      const sheet = model.sheets.find((item) => item.id === sheetId)
      if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
      sheet.name = name
    }))
    this.registerCommand<{ unitId: string; sheetId: string; range: { startRow: number; startColumn: number; endRow: number; endColumn: number }; values: Array<Array<unknown>> }, boolean>('sheet.command.set-range-values', ({ unitId, sheetId, range, values }) => this.commitWorkbook(unitId, 'Set range values', (model) => {
      this.assertCanEditRange(unitId, sheetId, { startRow: range.startRow, startCol: range.startColumn, endRow: range.endRow, endCol: range.endColumn })
      const sheet = model.sheets.find((item) => item.id === sheetId)
      if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
      const positions: Array<{ row: number; column: number }> = []
      for (let row = range.startRow; row <= range.endRow; row++) for (let col = range.startColumn; col <= range.endColumn; col++) {
        const cell = values[row - range.startRow]?.[col - range.startColumn]
        if (cell === undefined) continue
        const key = `${row}:${col}`
        sheet.cells[key] = applyCellValue(sheet.cells[key], row, col, cell as never)
        positions.push({ row, column: col })
      }
      const validation = validateSheetDataValidationChanges(model, sheetId, positions)
      if (validation.blocked) throw new Error(validation.issues[0]?.message ?? '输入值未通过数据验证，数据没有修改。')
    }))
    this.registerCommand<{ unitId: string; sheetId: string; rule: SheetDataValidationRule }, boolean>('sheet.command.set-data-validation', ({ unitId, sheetId, rule }) => this.commitWorkbook(unitId, '设置数据验证', (model) => {
      for (const range of rule.ranges) this.assertCanEditRange(unitId, sheetId, range)
      applySheetDataValidationRule(requireSheet(model, sheetId), rule)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; rule: SheetDataValidationRule }, boolean>('sheet.command.replace-data-validation', ({ unitId, sheetId, rule }) => this.commitWorkbook(unitId, 'Update data validation range', (model) => {
      const sheet = requireSheet(model, sheetId)
      const previousRanges = sheet.dataValidations?.find(existing => existing.id === rule.id)?.ranges ?? []
      for (const range of [...previousRanges, ...rule.ranges]) this.assertCanEditRange(unitId, sheetId, range)
      replaceSheetDataValidationRule(sheet, rule)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; ruleId: string }, boolean>('sheet.command.remove-data-validation', ({ unitId, sheetId, ruleId }) => this.commitWorkbook(unitId, 'Remove data validation rule', (model) => {
      const sheet = requireSheet(model, sheetId)
      const rule = sheet.dataValidations?.find(existing => existing.id === ruleId)
      for (const range of rule?.ranges ?? []) this.assertCanEditRange(unitId, sheetId, range)
      removeSheetDataValidationRule(sheet, ruleId)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; ranges: CellRange[] }, boolean>('sheet.command.clear-data-validation', ({ unitId, sheetId, ranges }) => this.commitWorkbook(unitId, '清除数据验证', (model) => {
      for (const range of ranges) this.assertCanEditRange(unitId, sheetId, range)
      clearSheetDataValidationRanges(requireSheet(model, sheetId), ranges)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; ranges: CellRange[]; types: SheetDataValidationType[] }, boolean>('sheet.command.clear-quick-data-validation', ({ unitId, sheetId, ranges, types }) => this.commitWorkbook(unitId, '清除快速数据验证', (model) => {
      for (const range of ranges) this.assertCanEditRange(unitId, sheetId, range)
      clearSheetDataValidationTypes(requireSheet(model, sheetId), ranges, types)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; range: CellRange | null }, boolean>('sheet.command.set-auto-filter', ({ unitId, sheetId, range }) => this.commitWorkbook(unitId, 'Set auto-filter', (model) => {
      setAutoFilter(requireSheet(model, sheetId), range)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; range: CellRange }, boolean>('sheet.command.toggle-auto-filter', ({ unitId, sheetId, range }) => this.commitWorkbook(unitId, 'Toggle auto-filter', (model) => {
      const sheet = requireSheet(model, sheetId)
      setAutoFilter(sheet, sheet.autoFilter ? null : range)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; column: number; criteria: SheetFilterCriteria | null }, boolean>('sheet.command.set-filter-criteria', ({ unitId, sheetId, column, criteria }) => this.commitWorkbook(unitId, 'Filter worksheet rows', (model) => {
      setFilterCriteria(requireSheet(model, sheetId), column, criteria)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; column?: number }, boolean>('sheet.command.clear-filter', ({ unitId, sheetId, column }) => this.commitWorkbook(unitId, 'Clear worksheet filter', (model) => {
      clearFilter(requireSheet(model, sheetId), column)
    }))
    this.registerCommand<{ unitId: string; sheetId: string }, boolean>('sheet.command.reapply-filter', ({ unitId, sheetId }) => this.commitWorkbook(unitId, 'Reapply worksheet filter', (model) => {
      refreshAutoFilter(requireSheet(model, sheetId))
    }))
    this.registerCommand<{ unitId: string; sheetId: string; range: CellRange; criteria: SheetSortCriterion[]; options: SortRangeOptions; order?: number[]; expectedRevision?: number }, boolean>('sheet.command.sort-range', ({ unitId, sheetId, range, criteria, options, order, expectedRevision }) => {
      if (expectedRevision !== undefined && this.getWorkbookRevision(unitId) !== expectedRevision) throw new Error('Sort result is stale because the workbook changed during calculation')
      this.assertCanEditRange(unitId, sheetId, range)
      return this.commitWorkbook(unitId, 'Sort worksheet range', (model) => sortRange(requireSheet(model, sheetId), range, criteria, order, options, translateFormulaReferences))
    })
    this.registerCommand<{ unitId: string; sheetId: string; rule: SheetRangeProtectionRule }, boolean>('sheet.command.set-range-protection', ({ unitId, sheetId, rule }) => this.commitWorkbook(unitId, 'Set range protection', (model) => {
      const sheet = requireSheet(model, sheetId)
      const existing = sheet.rangeProtectionRules?.find(item => item.id === rule.id)
      if (existing && existing.ownerUserId !== this.currentUserId) throw new Error('只有保护规则创建者可以修改该规则。')
      if (rule.ownerUserId !== this.currentUserId) throw new Error('保护规则创建者必须是当前用户。')
      setSheetRangeProtectionRule(sheet, rule)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; ruleId: string }, boolean>('sheet.command.remove-range-protection', ({ unitId, sheetId, ruleId }) => this.commitWorkbook(unitId, 'Remove range protection', (model) => {
      removeSheetRangeProtectionRule(requireSheet(model, sheetId), ruleId, this.currentUserId)
    }))
    this.registerCommand<{ unitId: string; sheetId: string; patches: Record<string, SparklineSpec | null> }, boolean>('sheet.command.set-sparklines', ({ unitId, sheetId, patches }) => {
      const normalizedPatches = Object.fromEntries(Object.entries(patches).map(([address, sparkline]) => [
        address,
        sparkline === null ? null : createSparklineSpec({ targetRange: address, ...sparkline }),
      ])) as Record<string, SparklineSpec | null>
      const positions = Object.keys(normalizedPatches).map(parseSparklineAddress)
      for (const position of positions) this.assertCanEditRange(unitId, sheetId, {
        startRow: position.row, endRow: position.row, startCol: position.column, endCol: position.column,
      })
      return this.commitWorkbook(unitId, 'Set sparklines', (model) => {
        const sheet = requireSheet(model, sheetId)
        for (const [address, sparkline] of Object.entries(normalizedPatches)) {
          const position = parseSparklineAddress(address)
          const key = `${position.row}:${position.column}`
          const cell = sheet.cells[key] ?? (sparkline ? createCell(position.row, position.column) : undefined)
          if (!cell) continue
          if (sparkline) cell.sparkline = cloneModel(sparkline)
          else delete cell.sparkline
          if (sparkline) sheet.cells[key] = cell
        }
      })
    })
    this.registerCommand<{ unitId: string; sheetId: string; chartId: string; chart: ChartConfig }, boolean>('sheet.command.add-chart', ({ unitId, sheetId, chartId, chart }) => this.commitWorkbook(unitId, 'Insert chart', (model) => {
      if (!model.sheets.some((item) => item.id === sheetId)) throw new Error(`Worksheet not found: ${sheetId}`)
      if (model.charts?.some((item) => item.id === chartId)) throw new Error(`Chart already exists: ${chartId}`)
      const created = createChart(model, sheetId, chart, chartId)
      model.charts ??= []
      model.charts.push(created)
    }))
    this.registerCommand<{ unitId: string; chartId: string; patch: Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>> }, boolean>('sheet.command.update-chart', ({ unitId, chartId, patch }) => this.commitWorkbook(unitId, 'Update chart', (model) => {
      updateChart(model, requireChart(model, chartId), patch)
    }))
    this.registerCommand<{ unitId: string; chartId: string }, boolean>('sheet.command.remove-chart', ({ unitId, chartId }) => this.commitWorkbook(unitId, 'Remove chart', (model) => {
      if (!model.charts?.some((item) => item.id === chartId)) throw new Error(`Chart not found: ${chartId}`)
      model.charts = model.charts.filter((item) => item.id !== chartId)
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
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; setting: { displayName?: string; format?: string; subtotalType?: PivotAggregation } }, boolean>('sheet.command.set-pivot-field-setting', ({ unitId, pivotTableId, fieldId, setting }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-field-setting', 'Set pivot field setting', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      setPivotFieldSetting(model, pivot, fieldId, setting)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; filter: PivotFieldFilter | null }, boolean>('sheet.command.set-pivot-value-filter', ({ unitId, pivotTableId, fieldId, filter }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-value-filter', 'Filter pivot values', (model) => {
      const pivot = requirePivot(model, pivotTableId)
      setPivotValueFilter(model, pivot, fieldId, filter)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; settings: PivotNumberGroup | null }, boolean>('sheet.command.set-pivot-number-group', ({ unitId, pivotTableId, fieldId, settings }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-number-group', 'Group pivot numbers', (model) => {
      setPivotNumberGroup(model, requirePivot(model, pivotTableId), fieldId, settings)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; mode: PivotDateGroupingMode }, boolean>('sheet.command.set-pivot-date-grouping-mode', ({ unitId, pivotTableId, mode }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-date-grouping-mode', 'Set pivot date grouping mode', (model) => {
      setPivotDateGroupingMode(model, requirePivot(model, pivotTableId), mode)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; type: PivotDateGroupType | null }, boolean>('sheet.command.set-pivot-date-group-type', ({ unitId, pivotTableId, fieldId, type }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-date-group-type', 'Group pivot dates', (model) => {
      setPivotDateGroupType(model, requirePivot(model, pivotTableId), fieldId, type)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; type: PivotDateGroupType }, boolean>('sheet.command.set-pivot-excel-date-group', ({ unitId, pivotTableId, fieldId, type }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-excel-date-group', 'Group pivot dates (Excel)', (model) => {
      setPivotExcelDateGroup(model, requirePivot(model, pivotTableId), fieldId, type)
    }))
    this.registerCommand<{ unitId: string; pivotTableId: string; fieldId: string; groups: PivotManualGroup[] }, boolean>('sheet.command.set-pivot-manual-groups', ({ unitId, pivotTableId, fieldId, groups }) => this.commitPivotChange(unitId, pivotTableId, undefined, 'sheet.command.set-pivot-manual-groups', 'Group pivot members', (model) => {
      setPivotManualGroups(model, requirePivot(model, pivotTableId), fieldId, groups)
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
    this.registerCommand<{ unitId: string; slicer: SlicerCreateInput }, boolean>('sheet.command.add-slicer', ({ unitId, slicer }) => this.commitSlicerChange(unitId, slicer.id, slicer.hostSheetId, 'sheet.command.add-slicer', 'Add slicer', (model) => {
      if (slicer.unitId !== unitId) throw new Error('Slicer workbook does not match the command workbook')
      createSlicer(model, slicer)
    }))
    this.registerCommand<{ unitId: string; slicerId: string; selectedItems: string[] | null }, boolean>('sheet.command.set-slicer-items', ({ unitId, slicerId, selectedItems }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.set-slicer-items', 'Filter with slicer', (model) => {
      setSlicerItems(model, requireSlicer(model, slicerId), selectedItems)
    }))
    this.registerCommand<{ unitId: string; slicerId: string; range: SlicerTimelineRange | null }, boolean>('sheet.command.set-timeline-range', ({ unitId, slicerId, range }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.set-timeline-range', 'Filter with timeline', (model) => {
      setSlicerTimelineRange(model, requireSlicer(model, slicerId), range)
    }))
    this.registerCommand<{ unitId: string; slicerId: string; patch: SlicerPresentationPatch }, boolean>('sheet.command.update-slicer-presentation', ({ unitId, slicerId, patch }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.update-slicer-presentation', 'Update slicer presentation', (model) => {
      updateSlicerPresentation(requireSlicer(model, slicerId), patch)
    }))
    this.registerCommand<{ unitId: string; slicerId: string }, boolean>('sheet.command.clear-slicer-filter', ({ unitId, slicerId }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.clear-slicer-filter', 'Clear slicer filter', (model) => {
      clearSlicerFilter(model, requireSlicer(model, slicerId))
    }))
    this.registerCommand<{ unitId: string; slicerId: string }, boolean>('sheet.command.remove-slicer', ({ unitId, slicerId }) => this.commitSlicerChange(unitId, slicerId, undefined, 'sheet.command.remove-slicer', 'Remove slicer', (model) => removeSlicer(model, slicerId)))
    this.registerCommand<{ unitId: string; name: string }, boolean>('doc.command.set-name', ({ unitId, name }) => this.commitDocument(unitId, 'Rename document', (model) => { model.name = name }))
    this.registerCommand<{ unitId: string; rules: DocumentModel['permissionRules'] }, boolean>('doc.command.set-document-permission-rules', ({ unitId, rules }) => this.commitDocumentPermissionRules(unitId, rules))
    this.registerCommand<{ unitId: string; index: number; text: string }, boolean>('doc.command.insert-text', ({ unitId, index, text }) => this.commitDocumentWithRangePermission(unitId, 'Insert text', index, index, (model) => Object.assign(model, insertDocumentText(model, index, text))))
    this.registerCommand<{ unitId: string; startOffset: number; endOffset: number; text?: string }, boolean>('doc.command.replace-range', ({ unitId, startOffset, endOffset, text = '' }) => this.commitDocumentWithRangePermission(unitId, 'Edit text range', startOffset, endOffset, (model) => Object.assign(model, replaceDocumentRange(model, startOffset, endOffset, text))))
    this.registerCommand<{ unitId: string; control: ContentControlInput }, boolean>('doc.command.add-content-control', ({ unitId, control }) => this.commitDocument(unitId, 'Add content control', (model) => {
      createDocumentContentControl(model, control, control.id)
    }))
    this.registerCommand<{ unitId: string; controlId: string; patch: ContentControlPatch }, boolean>('doc.command.update-content-control', ({ unitId, controlId, patch }) => this.commitDocument(unitId, 'Update content control', (model) => {
      updateDocumentContentControl(model, controlId, patch)
    }))
    this.registerCommand<{ unitId: string; controlId: string; value: string }, boolean>('doc.command.set-content-control-value', ({ unitId, controlId, value }) => this.commitDocument(unitId, 'Set content control value', (model) => {
      setDocumentContentControlValue(model, controlId, value)
    }))
    this.registerCommand<{ unitId: string; controlId: string; checked: boolean }, boolean>('doc.command.set-content-control-checked', ({ unitId, controlId, checked }) => this.commitDocument(unitId, 'Set content control checkbox', (model) => {
      setDocumentContentControlChecked(model, controlId, checked)
    }))
    this.registerCommand<{ unitId: string; controlId: string }, boolean>('doc.command.reset-content-control-placeholder', ({ unitId, controlId }) => this.commitDocument(unitId, 'Reset content control placeholder', (model) => {
      resetDocumentContentControlPlaceholder(model, controlId)
    }))
    this.registerCommand<{ unitId: string; controlId: string }, boolean>('doc.command.duplicate-repeating-section-item', ({ unitId, controlId }) => this.commitDocument(unitId, 'Duplicate repeating-section item', (model) => {
      duplicateRepeatingSectionItem(model, controlId)
    }))
    this.registerCommand<{ unitId: string; controlId: string }, boolean>('doc.command.delete-repeating-section-item', ({ unitId, controlId }) => this.commitDocument(unitId, 'Delete repeating-section item', (model) => {
      deleteRepeatingSectionItem(model, controlId)
    }))
    this.registerCommand<{ unitId: string; controlId: string; value: string }, boolean>('doc.command.select-content-control-item', ({ unitId, controlId, value }) => this.commitDocument(unitId, 'Select content control item', (model) => {
      selectDocumentContentControlItem(model, controlId, value)
    }))
    this.registerCommand<{ unitId: string; controlId: string }, boolean>('doc.command.remove-content-control', ({ unitId, controlId }) => this.commitDocument(unitId, 'Remove content control', (model) => {
      removeDocumentContentControl(model, controlId)
    }))
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

  private commitSlicerChange(
    unitId: string,
    slicerId: string,
    hostSheetId: string | undefined,
    commandId: string,
    label: string,
    update: (model: WorkbookModel) => void,
  ): boolean {
    const current = this.getWorkbookModel(unitId)
    if (!current) throw new Error(`Workbook not found: ${unitId}`)
    const currentSlicer = current.slicers?.find((slicer) => slicer.id === slicerId)
    const event: ApiEventMap['BeforeSlicerChange'] = {
      unitId,
      hostSheetId: hostSheetId ?? currentSlicer?.hostSheetId ?? '',
      slicerId,
      commandId,
    }
    this.events.emit('BeforeSlicerChange', event)
    if (event.cancel) throw new Error(`Slicer change was canceled: ${commandId}`)
    const changed = this.commitWorkbook(unitId, label, update)
    if (!changed) return false
    const next = this.getWorkbookModel(unitId)?.slicers?.find((slicer) => slicer.id === slicerId)
    if (commandId === 'sheet.command.add-slicer' && next) {
      this.events.emit('SlicerAdded', { unitId, hostSheetId: next.hostSheetId, slicerId, type: next.type })
    } else if (commandId === 'sheet.command.remove-slicer') {
      this.events.emit('SlicerRemoved', { unitId, hostSheetId: event.hostSheetId, slicerId })
    } else if (next) {
      this.events.emit('SlicerChanged', { unitId, hostSheetId: next.hostSheetId, slicerId, revision: next.revision })
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

function requireSheet(model: WorkbookModel, sheetId: string) {
  const sheet = model.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
  return sheet
}

function parseSparklineAddress(address: string): { row: number; column: number } {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(address)
  if (!match) throw new TypeError(`Invalid sparkline target address: ${address}`)
  let column = 0
  for (const letter of match[1]!.toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, column: column - 1 }
}

function rangeEditFromParams(params: unknown): { sheetId: string; range: CellRange } | undefined {
  if (!params || typeof params !== 'object') return undefined
  const input = params as Record<string, unknown>
  if (typeof input.sheetId !== 'string' || !input.range || typeof input.range !== 'object') return undefined
  const source = input.range as Record<string, unknown>
  const startCol = source.startCol ?? source.startColumn
  const endCol = source.endCol ?? source.endColumn
  const values = [source.startRow, startCol, source.endRow, endCol]
  if (!values.every(value => typeof value === 'number' && Number.isInteger(value))) return undefined
  return {
    sheetId: input.sheetId,
    range: { startRow: source.startRow as number, startCol: startCol as number, endRow: source.endRow as number, endCol: endCol as number },
  }
}

function requireSlicer(model: WorkbookModel, slicerId: string) {
  const slicer = model.slicers?.find((candidate) => candidate.id === slicerId)
  if (!slicer) throw new Error(`Slicer not found: ${slicerId}`)
  return slicer
}

function requireChart(model: WorkbookModel, chartId: string): ChartModel {
  const chart = model.charts?.find((candidate) => candidate.id === chartId)
  if (!chart) throw new Error(`Chart not found: ${chartId}`)
  return chart
}

