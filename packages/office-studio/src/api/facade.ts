import type {
  CellScalar,
  ChartAxisOptions,
  ChartLegendOptions,
  ChartModel,
  ChartSource,
  ChartSourceOrientation,
  ChartType,
  CellRange,
  DocumentContentControl,
  DocumentContentControlType,
  DocumentModel,
  DocumentParagraphStyle,
  PivotAggregation,
  PivotDateGroupType,
  PivotDateGroupingMode,
  PivotFieldArea,
  PivotFieldFilter,
  PivotManualGroup,
  PivotNumberGroup,
  PivotTableModel,
  PivotTableOptions,
  SlicerModel,
  SlicerTimelineRange,
  SpreadsheetCell,
  SpreadsheetSheet,
  SheetDataValidationRule,
  SheetRangeProtectionDraft,
  SheetRangeProtectionRule,
  SheetFilterCriteria,
  SheetSortCriterion,
  WorkbookModel,
} from '../domain/model'
import { createBlankSheet, createCell } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import {
  getCommonDocumentTextStyle,
  getDocumentParagraphStyle,
  getDocumentTextStyleRuns,
  documentBodyText,
  findDocumentParagraphs,
  readDocumentRange,
  saveDocumentSnapshot,
  setDocumentParagraphStyle,
  setDocumentTextStyle,
  type DocumentParagraphIdentity,
  type DocumentParagraphInfo,
  type DocumentTextStyle,
  type DocumentTextStylePatch,
  type DocumentTextStyleRun,
} from '../features/document/services/document-facade'
import {
  contentControlRange,
  contentControlText,
  getDocumentContentControl,
  getDocumentContentControls,
  type ContentControlInput,
  type ContentControlPatch,
} from '../features/document/services/content-controls'
import type { OfficeStudio } from './office-studio'
import { FEventName, type ApiEventMap, type Disposable } from './events'
import { applyCellValue, randomId, type ApiCellValue } from './model-utils'

export type {
  DocumentBodySegment,
  DocumentBodySnapshot,
  DocumentParagraphIdentity,
  DocumentParagraphInfo,
  DocumentParagraphQuery,
  DocumentTextRange,
  DocumentTextStyle,
  DocumentTextStylePatch,
  DocumentTextStyleRun,
} from '../features/document/services/document-facade'
import {
  getPivotSourceFields,
  getPivotTableAtCell,
  pivotDrillDown,
  pivotFieldIds,
  pivotOutputValues,
  type PivotTableAnchorInput,
  type PivotTableSourceInput,
} from '../features/pivot/services/pivot-engine'
import {
  chartLayout,
  chartRenderCommands,
  readChartData,
  type ChartConfig,
} from '../features/charts/chart-engine'
import { getSlicer, getSlicerItems, type SlicerCreateInput } from '../features/slicer/slicer-engine'
import { expandVisibleSortOrder, getFilteredRows, getNumericSortColumns, getSortedRowOrder, type SortRangeOptions } from '../features/sheet-sort-filter/sort-filter-engine'
import { createSheetDataValidationRule, sheetDataValidationOptions, validateSheetDataValidationChanges, type SheetDataValidationDraft, type SheetValidationIssue } from '../features/sheet-validation/validation-engine.ts'
import { createSheetRangeProtectionRule } from '../features/sheet-protection/range-protection.ts'

export interface CreateUnitOptions { makeCurrent?: boolean }
export interface FRangeCoordinates { startRow: number; startColumn: number; endRow: number; endColumn: number }
export type PivotTableSourceInfo = PivotTableSourceInput
export type PivotTableAnchorInfo = PivotTableAnchorInput
export type ChartInfo = ChartConfig
export type ChartPatch = Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>
export type SlicerInfo = Omit<SlicerCreateInput, 'id' | 'unitId' | 'hostSheetId'> & { id?: string }
export interface SortRangeApiOptions extends SortRangeOptions { signal?: AbortSignal; timeoutMs?: number; onProgress?: (progress: import('../compute/protocol.ts').ComputeProgress) => void }
export type DataValidationRuleInput = SheetDataValidationDraft
export type DataValidationRule = SheetDataValidationRule
export type DataValidationError = SheetValidationIssue
export type SortColumnSpec = { column: number; ascending: boolean } | number
export type DocumentContentControlInfo = DocumentContentControl
export type DocumentContentControlCreateInfo = Omit<ContentControlInput, 'id'> & { id?: string }
export type DocumentContentControlPatch = ContentControlPatch

/**
 * Public root facade modelled after Univer's `FUniver.newAPI(...)` entrypoint.
 * The wrapped value is intentionally the headless OfficeStudio runtime so the
 * same API can be used by Vue, tests, workers, or an embedding application.
 */
export class FUniver {
  readonly Event = new FEventName()
  readonly studio: OfficeStudio

  private constructor(studio: OfficeStudio) {
    this.studio = studio
  }

  static newAPI(studio: OfficeStudio): FUniver {
    return new FUniver(studio)
  }

  createWorkbook(data: Partial<WorkbookModel> & Record<string, unknown> = {}, options: CreateUnitOptions = {}): FWorkbook {
    const model = this.studio.createWorkbook(data, options.makeCurrent ?? true)
    return new FWorkbook(this.studio, model.id!)
  }

  getActiveWorkbook(): FWorkbook | null {
    return this.getWorkbook(this.studio.activeWorkbookId)
  }

  getWorkbook(id: string): FWorkbook | null {
    return this.studio.getWorkbookModel(id) ? new FWorkbook(this.studio, id) : null
  }

  createDocument(data: Partial<DocumentModel> & Record<string, unknown> = {}, options: CreateUnitOptions = {}): FDocument {
    const model = this.studio.createDocument(data, options.makeCurrent ?? true)
    return new FDocument(this.studio, model.id!)
  }

  getActiveDocument(): FDocument | null {
    return this.getDocument(this.studio.activeDocumentId)
  }

  getDocument(id: string): FDocument | null {
    return this.studio.getDocumentModel(id) ? new FDocument(this.studio, id) : null
  }

  getActiveSheet(): { workbook: FWorkbook; worksheet: FWorksheet } | null {
    const workbook = this.getActiveWorkbook()
    const worksheet = workbook?.getActiveSheet() ?? null
    return workbook && worksheet ? { workbook, worksheet } : null
  }

  getSheetCommandTarget(params: { unitId?: string; subUnitId?: string; sheetId?: string } = {}): { workbook: FWorkbook; worksheet: FWorksheet; unitId: string; subUnitId: string } | null {
    const workbook = params.unitId ? this.getWorkbook(params.unitId) : this.getActiveWorkbook()
    if (!workbook) return null
    const sheet = params.subUnitId || params.sheetId
    const worksheet = sheet ? workbook.getSheetBySheetId(sheet) : workbook.getActiveSheet()
    return worksheet ? { workbook, worksheet, unitId: workbook.getId(), subUnitId: worksheet.getSheetId() } : null
  }

  getPivotTableByCell(unitId: string, subUnitId: string, row: number, col: number): FPivotTable | null {
    const workbook = this.getWorkbook(unitId)
    return workbook?.getPivotTableByCell(subUnitId, row, col) ?? null
  }

  disposeUnit(unitId: string): boolean { return this.studio.disposeUnit(unitId) }
  addEvent<K extends keyof ApiEventMap>(event: K, listener: (payload: ApiEventMap[K]) => void): Disposable { return this.studio.events.addEvent(event, listener) }
  registerCommand<P, R>(id: string, handler: (params: P) => R): Disposable { return this.studio.registerCommand(id, handler) }
  executeCommand<P = Record<string, unknown>, R = boolean>(id: string, params?: P): Promise<R> { return this.studio.executeCommand(id, params) }
  executeCommandSync<P, R>(id: string, params: P): R { return this.studio.executeCommandSync(id, params) }
  setActiveKind(kind: 'workbook' | 'document'): void { this.studio.setActiveKind(kind) }
  getCurrentLifecycleStage(): 'ready' | 'disposed' { return this.studio.isDisposed ? 'disposed' : 'ready' }
  dispose(): void { this.studio.dispose() }
}

export class FWorkbook {
  private readonly api: OfficeStudio
  readonly id: string

  constructor(api: OfficeStudio, id: string) {
    this.api = api
    this.id = id
  }

  getId(): string { return this.id }
  getName(): string { return this.model().name }
  setName(name: string): this {
    this.api.executeCommandSync('sheet.command.set-workbook-name', { unitId: this.id, name })
    return this
  }
  save(): WorkbookModel {
    const snapshot = cloneModel(this.model())
    snapshot.modified = false
    this.api.markSaved('workbook', this.id)
    return snapshot
  }
  undo(): this { this.api.undo('workbook', this.id); return this }
  redo(): this { this.api.redo('workbook', this.id); return this }
  getActiveSheet(): FWorksheet | null {
    const sheet = this.model().sheets.find((item) => item.id === this.model().activeSheetId)
    return sheet ? new FWorksheet(this.api, this.id, sheet.id) : null
  }
  getSheets(): FWorksheet[] { return this.model().sheets.map((sheet) => new FWorksheet(this.api, this.id, sheet.id)) }
  getSheetBySheetId(sheetId: string): FWorksheet | null {
    return this.model().sheets.some((sheet) => sheet.id === sheetId) ? new FWorksheet(this.api, this.id, sheetId) : null
  }
  getSheetByName(name: string): FWorksheet | null {
    const sheet = this.model().sheets.find((item) => item.name === name)
    return sheet ? new FWorksheet(this.api, this.id, sheet.id) : null
  }
  setActiveSheet(sheet: FWorksheet | string): FWorksheet {
    const id = typeof sheet === 'string' ? sheet : sheet.getSheetId()
    if (!this.model().sheets.some((item) => item.id === id)) throw new Error(`Worksheet not found: ${id}`)
    this.api.mutateWorkbook(this.id, 'sheet.operation.set-active-sheet', 'Set active worksheet', (model) => { model.activeSheetId = id })
    this.api.activateWorkbook(this.id)
    return new FWorksheet(this.api, this.id, id)
  }
  insertSheet(sheetName?: string, options: { index?: number; sheet?: Partial<SpreadsheetSheet> } = {}): FWorksheet {
    const model = this.model()
    const name = sheetName ?? `Sheet${model.sheets.length + 1}`
    if (model.sheets.some((sheet) => sheet.name === name)) throw new Error(`Worksheet name already exists: ${name}`)
    const sheet = { ...createBlankSheet(name), ...(options.sheet ?? {}) }
    sheet.id ||= randomId('sheet')
    sheet.name = name
    const index = options.index ?? model.sheets.length
    if (!Number.isInteger(index) || index < 0 || index > model.sheets.length) throw new RangeError(`Invalid worksheet index: ${index}`)
    this.api.mutateWorkbook(this.id, 'sheet.command.insert-sheet', 'Insert worksheet', (draft) => {
      draft.sheets.splice(index, 0, cloneModel(sheet))
      draft.activeSheetId = sheet.id
    })
    this.api.activateWorkbook(this.id)
    return new FWorksheet(this.api, this.id, sheet.id)
  }
  deleteSheet(sheet: FWorksheet | string): boolean {
    const id = typeof sheet === 'string' ? sheet : sheet.getSheetId()
    if (this.model().sheets.length <= 1) throw new Error('A workbook must keep at least one worksheet')
    let deleted = false
    this.api.mutateWorkbook(this.id, 'sheet.command.delete-sheet', 'Delete worksheet', (model) => {
      const index = model.sheets.findIndex((item) => item.id === id)
      if (index < 0) return
      model.sheets.splice(index, 1)
      if (model.activeSheetId === id) model.activeSheetId = model.sheets[Math.max(0, index - 1)]!.id
      deleted = true
    })
    return deleted
  }
  getNumSheets(): number { return this.model().sheets.length }

  async addPivotTable(
    sourceInfo: PivotTableSourceInfo,
    positionType: PositionTypeEnum | string = PositionTypeEnum.Existing,
    anchorCellInfo: PivotTableAnchorInfo,
  ): Promise<FPivotTable> {
    if (!sourceInfo || !anchorCellInfo) throw new TypeError('Pivot source and anchor are required')
    if (positionType !== PositionTypeEnum.Existing) throw new Error(`Unsupported pivot table position: ${positionType}`)
    const pivotTableId = randomId('pivot')
    this.api.executeCommandSync('sheet.command.add-pivot-table', {
      unitId: this.id,
      pivotTableId,
      source: normalizePivotSourceInfo(sourceInfo, this.id),
      anchor: normalizePivotAnchorInfo(anchorCellInfo, this.id),
    })
    return new FPivotTable(this.api, this.id, pivotTableId)
  }

  getPivotTables(): FPivotTable[] {
    return (this.model().pivotTables ?? []).map((pivot) => new FPivotTable(this.api, this.id, pivot.id))
  }

  getSlicers(): FSheetSlicer[] {
    return (this.model().slicers ?? []).map((slicer) => new FSheetSlicer(this.api, this.id, slicer.id))
  }

  getSlicer(slicerId: string): FSheetSlicer | null {
    return getSlicer(this.model(), slicerId) ? new FSheetSlicer(this.api, this.id, slicerId) : null
  }

  getPivotTableByCell(subUnitId: string, row: number, col: number): FPivotTable | null
  getPivotTableByCell(unitId: string, subUnitId: string, row: number, col: number): FPivotTable | null
  getPivotTableByCell(unitIdOrSheetId: string, subUnitIdOrRow: string | number, rowOrCol: number, col?: number): FPivotTable | null {
    const sheetId = col == null ? unitIdOrSheetId : String(subUnitIdOrRow)
    const row = col == null ? Number(subUnitIdOrRow) : rowOrCol
    const column = col ?? rowOrCol
    if (col != null && unitIdOrSheetId !== this.id) return null
    const pivot = getPivotTableAtCell(this.model(), sheetId, row, column)
    return pivot ? new FPivotTable(this.api, this.id, pivot.id) : null
  }

  private model(): WorkbookModel {
    const model = this.api.getWorkbookModel(this.id)
    if (!model) throw new Error(`Workbook not found: ${this.id}`)
    return model
  }
}

export class FWorksheet {
  private readonly api: OfficeStudio
  readonly unitId: string
  readonly id: string

  constructor(api: OfficeStudio, unitId: string, id: string) {
    this.api = api
    this.unitId = unitId
    this.id = id
  }

  getUnitId(): string { return this.unitId }
  getSheetId(): string { return this.id }
  getSheetName(): string { return this.sheet().name }
  getDataValidations(): DataValidationRule[] { return cloneModel(this.sheet().dataValidations ?? []) }
  getDataValidation(range: FRange | FRangeCoordinates | string): DataValidationRule | null {
    const bounds = normalizeValidationRange(range)
    return cloneModel(this.sheet().dataValidations?.find(rule => rule.ranges.some(item => rangesIntersect(item, bounds))) ?? null)
  }
  setDataValidation(range: FRange | FRangeCoordinates | string, input: DataValidationRuleInput | null): this {
    const bounds = normalizeValidationRange(range)
    if (input === null) this.api.executeCommandSync('sheet.command.clear-data-validation', { unitId: this.unitId, sheetId: this.id, ranges: [bounds] })
    else {
      const rule = createSheetDataValidationRule(input, [bounds])
      this.api.executeCommandSync('sheet.command.set-data-validation', { unitId: this.unitId, sheetId: this.id, rule })
    }
    return this
  }
  getDataValidationErrors(range: FRange | FRangeCoordinates | string): DataValidationError[] {
    const bounds = normalizeValidationRange(range)
    const positions = Array.from({ length: bounds.endRow - bounds.startRow + 1 }, (_, rowOffset) =>
      Array.from({ length: bounds.endCol - bounds.startCol + 1 }, (_, columnOffset) => ({ row: bounds.startRow + rowOffset, column: bounds.startCol + columnOffset })),
    ).flat()
    return validateSheetDataValidationChanges(this.workbook(), this.id, positions).issues
  }
  getRangeProtectionRules(): SheetRangeProtectionRule[] { return cloneModel(this.sheet().rangeProtectionRules ?? []) }
  getRangeProtectionRulesForRange(range: FRange | FRangeCoordinates | string): SheetRangeProtectionRule[] {
    const bounds = normalizeValidationRange(range)
    return cloneModel((this.sheet().rangeProtectionRules ?? []).filter(rule => rule.ranges.some(item => rangesIntersect(item, bounds))))
  }
  setRangeProtection(range: FRange | FRangeCoordinates | string, draft: SheetRangeProtectionDraft): SheetRangeProtectionRule {
    const rule = createSheetRangeProtectionRule(draft, [normalizeValidationRange(range)], this.api.currentUserId)
    this.api.executeCommandSync('sheet.command.set-range-protection', { unitId: this.unitId, sheetId: this.id, rule })
    return cloneModel(rule)
  }
  removeRangeProtection(ruleId: string): boolean {
    return this.api.executeCommandSync('sheet.command.remove-range-protection', { unitId: this.unitId, sheetId: this.id, ruleId })
  }
  getCellValidationOptions(row: number, column: number): string[] {
    const rules = this.sheet().dataValidations?.filter(rule => rule.ranges.some(range => row >= range.startRow && row <= range.endRow && column >= range.startCol && column <= range.endCol)) ?? []
    return rules.flatMap(rule => sheetDataValidationOptions(rule, this.workbook(), this.id))
  }
  setAutoFilter(range: CellRange | FRangeCoordinates | FRange | string | null): this {
    const normalized = range === null ? null : typeof range === 'string'
      ? toCellRange(parseA1Range(range))
      : range instanceof FRange ? toCellRange(range.getRange())
      : 'startCol' in range ? normalizeToCellRange({ startRow: range.startRow, startColumn: range.startCol, endRow: range.endRow, endColumn: range.endCol })
      : normalizeToCellRange(range)
    this.api.executeCommandSync('sheet.command.set-auto-filter', { unitId: this.unitId, sheetId: this.id, range: normalized })
    return this
  }
  getAutoFilter(): { range: CellRange; columns: Record<number, SheetFilterCriteria>; filteredRows: number[] } | null {
    const filter = this.sheet().autoFilter
    return filter ? cloneModel({ ...filter, filteredRows: getFilteredRows(this.sheet()) }) : null
  }
  setFilterCriteria(column: number, criteria: SheetFilterCriteria | null): this {
    this.api.executeCommandSync('sheet.command.set-filter-criteria', { unitId: this.unitId, sheetId: this.id, column, criteria })
    return this
  }
  clearFilter(column?: number): this {
    this.api.executeCommandSync('sheet.command.clear-filter', { unitId: this.unitId, sheetId: this.id, column })
    return this
  }
  getFilteredRows(): number[] { return getFilteredRows(this.sheet()) }
  async sortRange(range: FRange | FRangeCoordinates | string, criteria: readonly SheetSortCriterion[], options: SortRangeApiOptions = {}): Promise<this> {
    const normalized = typeof range === 'string' ? toCellRange(parseA1Range(range)) : range instanceof FRange ? toCellRange(range.getRange()) : normalizeToCellRange(range)
    const sheet = this.sheet()
    const expectedRevision = this.api.getWorkbookRevision(this.unitId)
    const numeric = getNumericSortColumns(sheet, normalized, criteria, options)
    let order: number[] | undefined
    if (numeric && numeric.rowCount > 0) {
      const result = await this.api.numericIndexService.sortRows(numeric, {
        documentId: this.unitId,
        modelRevision: expectedRevision,
        signal: options.signal,
        timeoutMs: options.timeoutMs,
        onProgress: options.onProgress,
      })
      if (result) {
        const dataRowCount = normalized.endRow - normalized.startRow + 1 - (options.hasHeader ? 1 : 0)
        order = expandVisibleSortOrder(dataRowCount, numeric.destinationOffsets, Array.from(result.indices))
      }
    }
    if (this.sheet() !== sheet || this.api.getWorkbookRevision(this.unitId) !== expectedRevision) throw new Error('Sort result is stale because the workbook changed during calculation')
    order ??= getSortedRowOrder(sheet, normalized, criteria, options)
    this.api.executeCommandSync('sheet.command.sort-range', { unitId: this.unitId, sheetId: this.id, range: normalized, criteria: [...criteria], options: { hasHeader: options.hasHeader, locale: options.locale }, order, expectedRevision })
    return this
  }
  getPivotTableByCell(row: number, col: number): FPivotTable | null {
    const pivot = getPivotTableAtCell(this.workbook(), this.id, row, col)
    return pivot ? new FPivotTable(this.api, this.unitId, pivot.id) : null
  }
  newChart(type: ChartType): ChartConfigBuilder { return new ChartConfigBuilder(type) }
  async insertChart(info: ChartInfo): Promise<FChart> {
    const chartId = randomId('chart')
    this.api.executeCommandSync('sheet.command.add-chart', { unitId: this.unitId, sheetId: this.id, chartId, chart: info })
    return new FChart(this.api, this.unitId, chartId)
  }
  getCharts(): FChart[] { return (this.workbook().charts ?? []).filter((chart) => chart.sheetId === this.id).map((chart) => new FChart(this.api, this.unitId, chart.id)) }
  getChartById(chartId: string): FChart | null { return this.workbook().charts?.some((chart) => chart.id === chartId && chart.sheetId === this.id) ? new FChart(this.api, this.unitId, chartId) : null }
  async insertSlicer(info: SlicerInfo): Promise<FSheetSlicer> {
    const slicerId = info.id ?? randomId(info.type === 'timeline' ? 'timeline' : 'slicer')
    const slicer: SlicerCreateInput = { ...info, id: slicerId, unitId: this.unitId, hostSheetId: this.id }
    this.api.executeCommandSync('sheet.command.add-slicer', { unitId: this.unitId, slicer })
    return new FSheetSlicer(this.api, this.unitId, slicerId)
  }
  getSlicers(): FSheetSlicer[] { return (this.workbook().slicers ?? []).filter((slicer) => slicer.hostSheetId === this.id).map((slicer) => new FSheetSlicer(this.api, this.unitId, slicer.id)) }
  getSlicer(slicerId: string): FSheetSlicer | null { return this.workbook().slicers?.some((slicer) => slicer.hostSheetId === this.id && slicer.id === slicerId) ? new FSheetSlicer(this.api, this.unitId, slicerId) : null }
  setName(name: string): this {
    const sheet = this.sheet()
    const oldName = sheet.name
    const before: ApiEventMap['BeforeSheetNameChange'] = { workbookId: this.unitId, sheetId: this.id, oldName, newName: name }
    this.api.events.emit('BeforeSheetNameChange', before)
    if (before.cancel) throw new Error('Worksheet name change was canceled')
    this.api.executeCommandSync('sheet.command.set-worksheet-name', { unitId: this.unitId, sheetId: this.id, name })
    this.api.events.emit('SheetNameChanged', { workbookId: this.unitId, sheetId: this.id, oldName, newName: name })
    return this
  }
  getIndex(): number { return this.workbook().sheets.findIndex((sheet) => sheet.id === this.id) }
  getMaxRows(): number { return this.sheet().rowCount ?? 1_048_576 }
  getMaxColumns(): number { return this.sheet().columnCount ?? 16_384 }
  getRange(row: number, column: number): FRange
  getRange(row: number, column: number, numRows: number): FRange
  getRange(row: number, column: number, numRows: number, numColumns: number): FRange
  getRange(a1Notation: string): FRange
  getRange(range: FRangeCoordinates): FRange
  getRange(rowOrRange: number | string | FRangeCoordinates, column?: number, numRows = 1, numColumns = 1): FRange {
    if (typeof rowOrRange === 'string') return new FRange(this.api, this.unitId, this.id, parseA1Range(rowOrRange))
    if (typeof rowOrRange === 'object') return new FRange(this.api, this.unitId, this.id, normalizeRange(rowOrRange))
    if (column == null) throw new TypeError('A column index is required')
    validateIndex(rowOrRange, 'row')
    validateIndex(column, 'column')
    validateSize(numRows, 'numRows')
    validateSize(numColumns, 'numColumns')
    return new FRange(this.api, this.unitId, this.id, {
      startRow: rowOrRange,
      startColumn: column,
      endRow: rowOrRange + numRows - 1,
      endColumn: column + numColumns - 1,
    })
  }
  getActiveRange(): FRange | null {
    const selection = this.api.getSelection()
    if (!selection || selection.sheetId !== this.id || selection.unitId !== this.unitId) return null
    return new FRange(this.api, this.unitId, this.id, {
      startRow: selection.startRow,
      startColumn: selection.startColumn,
      endRow: selection.endRow,
      endColumn: selection.endColumn,
    })
  }
  setActiveRange(range: FRange): this {
    this.api.setSelection(this.unitId, this.id, range.getRange())
    return this
  }
  getSelection(): FSelection { return new FSelection(this) }
  setFrozenRows(rows: number): this {
    validateIndex(rows, 'rows', true)
    this.api.mutateWorkbook(this.unitId, 'sheet.command.set-freeze', 'Freeze worksheet rows', (model) => { this.sheet(model).frozenRows = rows })
    return this
  }
  setFrozenColumns(columns: number): this {
    validateIndex(columns, 'columns', true)
    this.api.mutateWorkbook(this.unitId, 'sheet.command.set-freeze', 'Freeze worksheet columns', (model) => { this.sheet(model).frozenCols = columns })
    return this
  }
  setFreeze(freeze: { rows?: number; columns?: number }): this {
    const rows = freeze.rows ?? 0
    const columns = freeze.columns ?? 0
    validateIndex(rows, 'rows', true)
    validateIndex(columns, 'columns', true)
    this.api.mutateWorkbook(this.unitId, 'sheet.command.set-freeze', 'Freeze worksheet panes', (model) => {
      const sheet = this.sheet(model)
      sheet.frozenRows = rows
      sheet.frozenCols = columns
    })
    return this
  }
  getFrozenRows(): number { return this.sheet().frozenRows }
  getFrozenColumns(): number { return this.sheet().frozenCols }
  setHiddenGridlines(hidden: boolean): this {
    this.api.mutateWorkbook(this.unitId, 'sheet.command.set-gridlines', 'Set worksheet gridlines', (model) => { this.sheet(model).showGridlines = !hidden })
    return this
  }
  getHiddenGridlines(): boolean { return !this.sheet().showGridlines }
  activate(): this { this.api.activateWorkbook(this.unitId); return this }

  private workbook(): WorkbookModel {
    const model = this.api.getWorkbookModel(this.unitId)
    if (!model) throw new Error(`Workbook not found: ${this.unitId}`)
    return model
  }
  private sheet(model = this.workbook()): SpreadsheetSheet {
    const sheet = model.sheets.find((item) => item.id === this.id)
    if (!sheet) throw new Error(`Worksheet not found: ${this.id}`)
    return sheet
  }
}

export class FSelection {
  private readonly worksheet: FWorksheet

  constructor(worksheet: FWorksheet) {
    this.worksheet = worksheet
  }
  getActiveRange(): FRange | null { return this.worksheet.getActiveRange() }
  getCurrentCell(): FRange | null {
    const range = this.getActiveRange()
    return range ? this.worksheet.getRange(range.getRow(), range.getColumn()) : null
  }
  getActiveRangeList(): FRange[] { const range = this.getActiveRange(); return range ? [range] : [] }
}

export class ChartConfigBuilder {
  private value: ChartConfig

  constructor(type: ChartType) {
    this.value = { type, source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, orientation: 'columns' }, position: { x: 0, y: 0 }, size: { width: 480, height: 300 } }
  }
  setSource(range: string | ChartSource, orientation: ChartSourceOrientation = 'columns'): this {
    const source = typeof range === 'string' ? { range: parseChartA1Range(range), orientation } : range
    this.value = { ...this.value, source: { range: { ...source.range }, orientation: source.orientation } }
    return this
  }
  setPosition(cell: string): this {
    const { row, column } = parseA1Cell(cell)
    const x = column * 96, y = row * 24
    this.value = { ...this.value, position: { x, y } }
    return this
  }
  setAbsolutePosition(x: number, y: number): this { this.value = { ...this.value, position: { x, y } }; return this }
  setSize(width: number, height: number): this { this.value = { ...this.value, size: { width, height } }; return this }
  setCategoryField(index: number): this { this.value = { ...this.value, categoryField: index }; return this }
  setValueFields(fields: number[]): this { this.value = { ...this.value, valueFields: [...fields] }; return this }
  setTitle(title: string): this { this.value = { ...this.value, title }; return this }
  setLegend(legend: Partial<ChartLegendOptions>): this { this.value = { ...this.value, legend: { ...this.value.legend, ...legend } }; return this }
  setXAxis(axis: Partial<ChartAxisOptions>): this { this.value = { ...this.value, xAxis: { ...this.value.xAxis, ...axis } }; return this }
  setYAxis(axis: Partial<ChartAxisOptions>): this { this.value = { ...this.value, yAxis: { ...this.value.yAxis, ...axis } }; return this }
  build(): ChartInfo { return cloneModel(this.value) }
}

export class FChart {
  private readonly api: OfficeStudio
  private readonly unitId: string
  readonly id: string
  constructor(api: OfficeStudio, unitId: string, id: string) { this.api = api; this.unitId = unitId; this.id = id }
  getChartId(): string { return this.id }
  getInfo(): ChartModel { return cloneModel(this.model()) }
  getData(): ReturnType<typeof readChartData> { return readChartData(this.workbook(), this.model()) }
  getLayout(): ReturnType<typeof chartLayout> { return chartLayout(this.workbook(), this.model()) }
  getRenderCommands(): ReturnType<typeof chartRenderCommands> { return chartRenderCommands(this.workbook(), this.model()) }
  async update(info: ChartPatch): Promise<this> { await this.api.executeCommand('sheet.command.update-chart', { unitId: this.unitId, chartId: this.id, patch: info }); return this }
  async setDataSource(source: ChartSource): Promise<this> { return this.update({ source }) }
  async remove(): Promise<boolean> { return this.api.executeCommand('sheet.command.remove-chart', { unitId: this.unitId, chartId: this.id }) }
  private model(): ChartModel {
    const chart = this.workbook().charts?.find((item) => item.id === this.id)
    if (!chart) throw new Error(`Chart not found: ${this.id}`)
    return chart
  }
  private workbook(): WorkbookModel {
    const workbook = this.api.getWorkbookModel(this.unitId)
    if (!workbook) throw new Error(`Workbook not found: ${this.unitId}`)
    return workbook
  }
}

export class FRange {
  private readonly api: OfficeStudio
  readonly unitId: string
  readonly sheetId: string
  private readonly range: FRangeCoordinates

  constructor(
    api: OfficeStudio,
    unitId: string,
    sheetId: string,
    range: FRangeCoordinates
  ) {
    this.api = api
    this.unitId = unitId
    this.sheetId = sheetId
    this.range = range
  }

  getUnitId(): string { return this.unitId }
  getSheetId(): string { return this.sheetId }
  getRange(): FRangeCoordinates { return { ...this.range } }
  getRow(): number { return this.range.startRow }
  getLastRow(): number { return this.range.endRow }
  getColumn(): number { return this.range.startColumn }
  getLastColumn(): number { return this.range.endColumn }
  getHeight(): number { return this.range.endRow - this.range.startRow + 1 }
  getWidth(): number { return this.range.endColumn - this.range.startColumn + 1 }
  getA1Notation(withSheet = false): string {
    const first = cellAddress(this.range.startRow, this.range.startColumn)
    const last = cellAddress(this.range.endRow, this.range.endColumn)
    const ref = first === last ? first : `${first}:${last}`
    const sheetName = this.sheet().name.replaceAll("'", "''")
    return withSheet ? `'${sheetName}'!${ref}` : ref
  }
  getValue(): CellScalar { return this.cellAt(this.range.startRow, this.range.startColumn)?.value ?? null }
  getValues(): CellScalar[][] { return this.mapCells((cell) => cell?.value ?? null) }
  getRawValue(): CellScalar { return this.getValue() }
  getRawValues(): CellScalar[][] { return this.getValues() }
  getDisplayValue(): string { return this.cellAt(this.range.startRow, this.range.startColumn)?.display ?? '' }
  getDisplayValues(): string[][] { return this.mapCells((cell) => cell?.display ?? '') }
  getFormula(): string { return this.cellAt(this.range.startRow, this.range.startColumn)?.formula ?? '' }
  getFormulas(): string[][] { return this.mapCells((cell) => cell?.formula ?? '') }
  getDataValidation(): DataValidationRule | null { return new FWorksheet(this.api, this.unitId, this.sheetId).getDataValidation(this.range) }
  setDataValidation(rule: DataValidationRuleInput | null): this {
    new FWorksheet(this.api, this.unitId, this.sheetId).setDataValidation(this.range, rule)
    return this
  }
  getDataValidationErrors(): DataValidationError[] { return new FWorksheet(this.api, this.unitId, this.sheetId).getDataValidationErrors(this.range) }
  getRangeProtectionRules(): SheetRangeProtectionRule[] { return new FWorksheet(this.api, this.unitId, this.sheetId).getRangeProtectionRulesForRange(this.range) }
  setRangeProtection(draft: SheetRangeProtectionDraft): this {
    new FWorksheet(this.api, this.unitId, this.sheetId).setRangeProtection(this.range, draft)
    return this
  }
  sort(column: SortColumnSpec | SortColumnSpec[], options: SortRangeOptions = {}): this {
    const specs = Array.isArray(column) ? column : [column]
    const criteria: SheetSortCriterion[] = specs.map((spec) => {
      const relative = typeof spec === 'number' ? { column: spec, direction: 'ascending' as const } : { column: spec.column, direction: spec.ascending ? 'ascending' as const : 'descending' as const }
      if (!Number.isInteger(relative.column) || relative.column < 0 || relative.column >= this.getWidth()) throw new RangeError('Sort column must be relative to the range')
      return { column: this.range.startColumn + relative.column, direction: relative.direction }
    })
    const range: CellRange = { startRow: this.range.startRow, startCol: this.range.startColumn, endRow: this.range.endRow, endCol: this.range.endColumn }
    const order = getSortedRowOrder(this.sheet(), range, criteria, options)
    this.api.executeCommandSync('sheet.command.sort-range', { unitId: this.unitId, sheetId: this.sheetId, range, criteria, options, order })
    return this
  }
  getCellData(): SpreadsheetCell | null { const cell = this.cellAt(this.range.startRow, this.range.startColumn); return cell ? cloneModel(cell) : null }
  getCellDatas(): (SpreadsheetCell | null)[][] { return this.mapCells((cell) => cell ? cloneModel(cell) : null) }
  setValue(value: ApiCellValue): this { return this.setValues(this.fill(value)) }
  setValues(values: ApiCellValue[][]): this {
    assertMatrixSize(values, this.getHeight(), this.getWidth())
    const current = this.api.getWorkbookModel(this.unitId)
    if (!current) throw new Error(`Workbook not found: ${this.unitId}`)
    const candidate = cloneModel(current)
    const candidateSheet = candidate.sheets.find(sheet => sheet.id === this.sheetId)
    if (!candidateSheet) throw new Error(`Worksheet not found: ${this.sheetId}`)
    const positions: Array<{ row: number; column: number }> = []
    for (let row = 0; row < this.getHeight(); row++) for (let column = 0; column < this.getWidth(); column++) {
      const targetRow = this.range.startRow + row
      const targetColumn = this.range.startColumn + column
      candidateSheet.cells[`${targetRow}:${targetColumn}`] = applyCellValue(candidateSheet.cells[`${targetRow}:${targetColumn}`], targetRow, targetColumn, values[row]![column]!)
      positions.push({ row: targetRow, column: targetColumn })
    }
    const validation = validateSheetDataValidationChanges(candidate, this.sheetId, positions)
    if (validation.blocked) throw new Error(validation.issues.find(issue => issue.severity === 'stop')?.message ?? 'Input violates data validation')
    this.mutate('sheet.command.set-range-values', 'Set range values', (model) => {
      const sheet = this.sheet(model)
      for (let row = 0; row < this.getHeight(); row++) for (let column = 0; column < this.getWidth(); column++) {
        const targetRow = this.range.startRow + row
        const targetColumn = this.range.startColumn + column
        const key = `${targetRow}:${targetColumn}`
        sheet.cells[key] = applyCellValue(sheet.cells[key], targetRow, targetColumn, values[row]![column]!)
      }
    }, { range: this.range, values })
    this.api.events.emit('SheetValueChanged', { workbookId: this.unitId, sheetId: this.sheetId, ranges: [{ ...this.range, sheetId: this.sheetId, unitId: this.unitId }] })
    return this
  }
  setFormula(formula: string): this { return this.setValue({ f: formula }) }
  setFormulas(formulas: string[][]): this { return this.setValues(formulas.map((row) => row.map((formula) => ({ f: formula || null }))) as ApiCellValue[][]) }
  setFontWeight(weight: 'normal' | 'bold' | null): this { return this.mutateStyle('sheet.command.set-font-weight', 'Set font weight', (cell) => { cell.style.font.bold = weight === 'bold' }) }
  setFontStyle(style: 'normal' | 'italic' | null): this { return this.mutateStyle('sheet.command.set-font-style', 'Set font style', (cell) => { cell.style.font.italic = style === 'italic' }) }
  setFontFamily(family: string | null): this { return this.mutateStyle('sheet.command.set-font-family', 'Set font family', (cell) => { if (family) cell.style.font.family = family }) }
  setFontSize(size: number | null): this {
    if (size !== null && (!Number.isFinite(size) || size <= 0)) throw new RangeError('Font size must be a positive number')
    return this.mutateStyle('sheet.command.set-font-size', 'Set font size', (cell) => { if (size !== null) cell.style.font.size = size })
  }
  setFontColor(color: string | null): this { return this.mutateStyle('sheet.command.set-font-color', 'Set font color', (cell) => { if (color) cell.style.font.color = color }) }
  setBackground(color: string): this { return this.mutateStyle('sheet.command.set-background', 'Set background', (cell) => { cell.style.fill = color }) }
  setHorizontalAlignment(alignment: 'left' | 'center' | 'right' | 'justify'): this { return this.mutateStyle('sheet.command.set-horizontal-alignment', 'Set horizontal alignment', (cell) => { cell.style.horizontal = alignment }) }
  setVerticalAlignment(alignment: 'top' | 'middle' | 'bottom'): this { return this.mutateStyle('sheet.command.set-vertical-alignment', 'Set vertical alignment', (cell) => { cell.style.vertical = alignment }) }
  setWrap(enabled: boolean): this { return this.mutateStyle('sheet.command.set-wrap', 'Set text wrapping', (cell) => { cell.style.wrap = enabled }) }
  getBackground(): string { return this.cellAt(this.range.startRow, this.range.startColumn)?.style.fill ?? '#ffffff' }
  getFontWeight(): 'normal' | 'bold' { return this.cellAt(this.range.startRow, this.range.startColumn)?.style.font.bold ? 'bold' : 'normal' }
  merge(): this {
    if (this.getWidth() === 1 && this.getHeight() === 1) return this
    this.mutate('sheet.command.merge-cells', 'Merge cells', (model) => {
      const sheet = this.sheet(model)
      const range = { startRow: this.range.startRow, startCol: this.range.startColumn, endRow: this.range.endRow, endCol: this.range.endColumn }
      const overlaps = sheet.mergedRanges.some((merged) => !(range.endRow < merged.startRow || range.startRow > merged.endRow || range.endCol < merged.startCol || range.startCol > merged.endCol))
      if (overlaps) throw new Error('Merged ranges cannot overlap')
      sheet.mergedRanges.push(range)
      for (let row = range.startRow; row <= range.endRow; row++) for (let column = range.startCol; column <= range.endCol; column++) {
        if (row !== range.startRow || column !== range.startCol) delete sheet.cells[`${row}:${column}`]
      }
    })
    return this
  }
  clear(options: { contents?: boolean; format?: boolean } = {}): this {
    const clearContents = options.contents ?? true
    const clearFormat = options.format ?? false
    this.mutate('sheet.command.clear-range', 'Clear range', (model) => {
      const sheet = this.sheet(model)
      for (let row = this.range.startRow; row <= this.range.endRow; row++) for (let column = this.range.startColumn; column <= this.range.endColumn; column++) {
        const key = `${row}:${column}`
        const current = sheet.cells[key]
        if (!current) continue
        if (clearContents && clearFormat) delete sheet.cells[key]
        else if (clearContents) { current.value = null; current.formula = null; current.display = '' }
        else if (clearFormat) current.style = createCell(row, column).style
      }
      if (clearContents && clearFormat) sheet.mergedRanges = sheet.mergedRanges.filter((merged) => !rangeContains(this.range, merged))
    })
    return this
  }
  activate(): this { this.api.setSelection(this.unitId, this.sheetId, this.range); return this }

  private mutateStyle(id: string, label: string, change: (cell: SpreadsheetCell) => void): this {
    this.mutate(id, label, (model) => {
      const sheet = this.sheet(model)
      for (let row = this.range.startRow; row <= this.range.endRow; row++) for (let column = this.range.startColumn; column <= this.range.endColumn; column++) {
        const key = `${row}:${column}`
        const cell = sheet.cells[key] ?? createCell(row, column, null)
        change(cell)
        sheet.cells[key] = cell
      }
    })
    return this
  }
  private mutate(id: string, label: string, update: (model: WorkbookModel) => void, params: unknown = { range: this.range }): void {
    this.api.mutateWorkbook(this.unitId, id, label, update, { ...asRecord(params), unitId: this.unitId, sheetId: this.sheetId })
  }
  private mapCells<T>(read: (cell: SpreadsheetCell | undefined) => T): T[][] {
    const sheet = this.sheet()
    return Array.from({ length: this.getHeight() }, (_, row) => Array.from({ length: this.getWidth() }, (_, column) => read(sheet.cells[`${this.range.startRow + row}:${this.range.startColumn + column}`])))
  }
  private cellAt(row: number, column: number): SpreadsheetCell | undefined { return this.sheet().cells[`${row}:${column}`] }
  private sheet(model = this.api.getWorkbookModel(this.unitId)): SpreadsheetSheet {
    if (!model) throw new Error(`Workbook not found: ${this.unitId}`)
    const sheet = model.sheets.find((item) => item.id === this.sheetId)
    if (!sheet) throw new Error(`Worksheet not found: ${this.sheetId}`)
    return sheet
  }
  private fill(value: ApiCellValue): ApiCellValue[][] {
    return Array.from({ length: this.getHeight() }, () => Array.from({ length: this.getWidth() }, () => value))
  }
}

export const PositionTypeEnum = { Existing: 'existing' } as const
export type PositionTypeEnum = typeof PositionTypeEnum[keyof typeof PositionTypeEnum]

export class FPivotTable {
  private readonly api: OfficeStudio
  private readonly unitId: string
  readonly id: string

  constructor(api: OfficeStudio, unitId: string, id: string) {
    this.api = api
    this.unitId = unitId
    this.id = id
  }

  getPivotTableId(): string { return this.id }
  getSourceFieldsInfo(): Array<{ fieldId: string; fieldName: string; sourceColumn: number }> {
    return getPivotSourceFields(this.workbook(), this.model())
  }
  getFieldIdsByArea(area: PivotFieldArea): string[] { return pivotFieldIds(this.model(), area) }
  getData(): CellScalar[][] { return pivotOutputValues(this.workbook(), this.model()) }
  getRange(): FRange | null {
    const range = this.model().outputRange
    return range ? new FRange(this.api, this.unitId, this.model().anchor.subUnitId, {
      startRow: range.startRow,
      startColumn: range.startCol,
      endRow: range.endRow,
      endColumn: range.endCol,
    }) : null
  }
  getOptions(): PivotTableOptions { return cloneModel(this.model().options) }
  async addField(fieldIdOrIndex: string | number, area: PivotFieldArea, index?: number): Promise<this> {
    return this.execute('sheet.command.set-pivot-field', { fieldId: fieldIdOrIndex, area, index })
  }
  async removeField(fieldIds: string | string[]): Promise<this> {
    return this.execute('sheet.command.remove-pivot-field', { fieldIds: Array.isArray(fieldIds) ? fieldIds : [fieldIds] })
  }
  async setOptions(options: Partial<PivotTableOptions>): Promise<this> {
    return this.execute('sheet.command.set-pivot-options', { options })
  }
  async setLabelSort(fieldId: string, options: { type: 'ascending' | 'descending' }): Promise<this> {
    return this.execute('sheet.command.set-pivot-label-sort', { fieldId, direction: options.type })
  }
  async setLabelManualFilter(fieldId: string, selectedItems: string[] | null): Promise<this> {
    return this.execute('sheet.command.set-pivot-label-filter', { fieldId, selectedItems })
  }
  async setFieldSetting(fieldId: string, setting: { displayName?: string; format?: string; subtotalType?: PivotAggregation }): Promise<this> {
    return this.execute('sheet.command.set-pivot-field-setting', { fieldId, setting })
  }
  async setSubtotalType(fieldId: string, subtotalType: PivotAggregation): Promise<this> {
    return this.setFieldSetting(fieldId, { subtotalType })
  }
  async setValueFilter(fieldId: string, filter: PivotFieldFilter | null): Promise<this> {
    return this.execute('sheet.command.set-pivot-value-filter', { fieldId, filter })
  }
  async setNumberGroup(fieldId: string, settings: PivotNumberGroup | null): Promise<this> {
    return this.execute('sheet.command.set-pivot-number-group', { fieldId, settings })
  }
  async setDateGroupingMode(mode: PivotDateGroupingMode): Promise<this> {
    return this.execute('sheet.command.set-pivot-date-grouping-mode', { mode })
  }
  async setDateGroupType(fieldId: string, type: PivotDateGroupType | null): Promise<this> {
    return this.execute('sheet.command.set-pivot-date-group-type', { fieldId, type })
  }
  async setExcelDateGroup(fieldId: string, type: PivotDateGroupType): Promise<this> {
    return this.execute('sheet.command.set-pivot-excel-date-group', { fieldId, type })
  }
  async setManualGroups(fieldId: string, groups: PivotManualGroup[]): Promise<this> {
    return this.execute('sheet.command.set-pivot-manual-groups', { fieldId, groups })
  }
  async updateValuePosition(position: 'row' | 'column', _index = 0): Promise<this> {
    return this.execute('sheet.command.update-pivot-value-position', { position })
  }
  async refresh(): Promise<this> { return this.execute('sheet.command.refresh-pivot-table', {}) }
  async remove(): Promise<void> {
    await this.api.executeCommand('sheet.command.remove-pivot-table', { unitId: this.unitId, pivotTableId: this.id })
  }
  drillDown(rowOffset: number, columnOffset: number): { headers: string[]; rows: CellScalar[][] } {
    return pivotDrillDown(this.workbook(), this.model(), rowOffset, columnOffset)
  }

  private async execute(commandId: string, params: Record<string, unknown>): Promise<this> {
    await this.api.executeCommand(commandId, { unitId: this.unitId, pivotTableId: this.id, ...params })
    return this
  }
  private model(): PivotTableModel {
    const model = this.workbook().pivotTables?.find((pivot) => pivot.id === this.id)
    if (!model) throw new Error(`Pivot table not found: ${this.id}`)
    return model
  }
  private workbook(): WorkbookModel {
    const workbook = this.api.getWorkbookModel(this.unitId)
    if (!workbook) throw new Error(`Workbook not found: ${this.unitId}`)
    return workbook
  }
}

export class FSheetSlicer {
  private readonly api: OfficeStudio
  private readonly unitId: string
  readonly id: string

  constructor(api: OfficeStudio, unitId: string, id: string) {
    this.api = api
    this.unitId = unitId
    this.id = id
  }

  getId(): string { return this.id }
  getInfo(): SlicerModel { return cloneModel(this.model()) }
  getItems(): string[] { return getSlicerItems(this.workbook(), this.model()) }
  async setItems(items: string[] | null): Promise<this> {
    await this.api.executeCommand('sheet.command.set-slicer-items', { unitId: this.unitId, slicerId: this.id, selectedItems: items })
    return this
  }
  async setTimelineRange(range: SlicerTimelineRange | null): Promise<this> {
    await this.api.executeCommand('sheet.command.set-timeline-range', { unitId: this.unitId, slicerId: this.id, range })
    return this
  }
  async clearFilter(): Promise<this> {
    await this.api.executeCommand('sheet.command.clear-slicer-filter', { unitId: this.unitId, slicerId: this.id })
    return this
  }
  async remove(): Promise<boolean> {
    return this.api.executeCommand('sheet.command.remove-slicer', { unitId: this.unitId, slicerId: this.id })
  }
  private model(): SlicerModel {
    const model = getSlicer(this.workbook(), this.id)
    if (!model) throw new Error(`Slicer not found: ${this.id}`)
    return model
  }
  private workbook(): WorkbookModel {
    const workbook = this.api.getWorkbookModel(this.unitId)
    if (!workbook) throw new Error(`Workbook not found: ${this.unitId}`)
    return workbook
  }
}

export class FDocument {
  private readonly api: OfficeStudio
  readonly id: string

  constructor(api: OfficeStudio, id: string) {
    this.api = api
    this.id = id
  }

  getId(): string { return this.id }
  getName(): string { return this.model().name }
  setName(name: string): this { this.api.executeCommandSync('doc.command.set-name', { unitId: this.id, name }); return this }
  save(): DocumentModel {
    const snapshot = saveDocumentSnapshot(this.model())
    this.api.markSaved('document', this.id)
    return snapshot
  }
  undo(): boolean { return this.api.undo('document', this.id) }
  redo(): boolean { return this.api.redo('document', this.id) }
  getText(): string { return documentBodyText(this.model()) }
  getTextRange(startOffset: number, endOffset: number): FDocumentTextRange {
    const range = readDocumentRange(this.model(), startOffset, endOffset)
    return new FDocumentTextRange(this, range.startOffset, range.endOffset)
  }
  insertText(index: number, text: string): boolean {
    return this.api.executeCommandSync('doc.command.insert-text', { unitId: this.id, index, text })
  }
  getParagraphs(): FDocumentParagraph[] { return findDocumentParagraphs(this.model(), {}).map((item) => new FDocumentParagraph(this, item)) }
  findParagraphs(query: string | { text?: string; paragraphId?: string; blockId?: string }): FDocumentParagraph[] {
    return findDocumentParagraphs(this.model(), query).map((item) => new FDocumentParagraph(this, item))
  }
  appendParagraph(text = ''): FDocumentParagraph {
    const start = this.getText().length
    this.insertText(start, `${start ? '\n' : ''}${text}`)
    return this.getParagraphs().at(-1)!
  }
  getModel(): DocumentModel { return this.model() }
  getContentControls(): FDocumentContentControl[] {
    return getDocumentContentControls(this.model()).map((control) => new FDocumentContentControl(this, control.id))
  }
  getContentControl(id: string): FDocumentContentControl | null {
    return getDocumentContentControl(this.model(), id) ? new FDocumentContentControl(this, id) : null
  }
  addContentControl(info: DocumentContentControlCreateInfo): FDocumentContentControl {
    const id = info.id ?? randomId('content-control')
    this.api.executeCommandSync('doc.command.add-content-control', { unitId: this.id, control: { ...info, id } })
    return new FDocumentContentControl(this, id)
  }
  replaceRange(startOffset: number, endOffset: number, text: string): boolean {
    return this.api.executeCommandSync('doc.command.replace-range', { unitId: this.id, startOffset, endOffset, text })
  }
  mutate(label: string, commandId: string, update: (model: DocumentModel) => void): boolean {
    return this.api.mutateDocument(this.id, commandId, label, update)
  }
  executeContentControlCommand<P extends Record<string, unknown>>(commandId: string, params: P): boolean {
    return this.api.executeCommandSync(commandId, { unitId: this.id, ...params })
  }

  private model(): DocumentModel {
    const model = this.api.getDocumentModel(this.id)
    if (!model) throw new Error(`Document not found: ${this.id}`)
    return model
  }
}

export class FDocumentContentControl {
  private readonly document: FDocument
  readonly id: string

  constructor(document: FDocument, id: string) {
    this.document = document
    this.id = id
  }
  getId(): string { return this.id }
  getInfo(): DocumentContentControlInfo {
    const control = getDocumentContentControl(this.document.getModel(), this.id)
    if (!control) throw new Error(`Content control not found: ${this.id}`)
    return control
  }
  getType(): DocumentContentControlType { return this.getInfo().type }
  getText(): string { return contentControlText(this.document.getModel(), this.getInfo()) }
  getRange(): { blockId: string; startOffset: number; endOffset: number } { return contentControlRange(this.getInfo()) }
  setValue(value: string): this { this.document.executeContentControlCommand('doc.command.set-content-control-value', { controlId: this.id, value }); return this }
  setChecked(checked: boolean): this { this.document.executeContentControlCommand('doc.command.set-content-control-checked', { controlId: this.id, checked }); return this }
  selectItem(value: string): this { this.document.executeContentControlCommand('doc.command.select-content-control-item', { controlId: this.id, value }); return this }
  update(patch: DocumentContentControlPatch): this { this.document.executeContentControlCommand('doc.command.update-content-control', { controlId: this.id, patch }); return this }
  remove(): boolean { return this.document.executeContentControlCommand('doc.command.remove-content-control', { controlId: this.id }) }
}

export class FDocumentTextRange {
  private readonly document: FDocument
  readonly startOffset: number
  readonly endOffset: number

  constructor(document: FDocument, startOffset: number, endOffset: number) {
    this.document = document
    this.startOffset = startOffset
    this.endOffset = endOffset
  }
  getRange(): { startOffset: number; endOffset: number } { return { startOffset: this.startOffset, endOffset: this.endOffset } }
  getText(): string { return readDocumentRange(this.document.getModel(), this.startOffset, this.endOffset).text }
  setText(text: string): boolean { return this.document.replaceRange(this.startOffset, this.endOffset, text) }
  getTextStyleRuns(): DocumentTextStyleRun[] {
    return getDocumentTextStyleRuns(this.document.getModel(), this.startOffset, this.endOffset)
  }
  getTextStyle(): DocumentTextStyle {
    return getCommonDocumentTextStyle(this.document.getModel(), this.startOffset, this.endOffset)
  }
  setTextStyle(style: DocumentTextStylePatch): boolean {
    return this.document.mutate('Set document text style', 'doc.command.set-text-style', (model) => {
      setDocumentTextStyle(model, this.startOffset, this.endOffset, style)
    })
  }
}

export class FDocumentParagraph {
  private readonly document: FDocument
  private readonly info: DocumentParagraphInfo

  constructor(document: FDocument, info: DocumentParagraphInfo) {
    this.document = document
    this.info = info
  }
  getId(): string { return this.info.paragraphId }
  getInfo(): DocumentParagraphInfo {
    const current = findDocumentParagraphs(this.document.getModel(), { paragraphId: this.info.paragraphId })[0]
    if (!current) throw new Error(`Document paragraph with id ${this.info.paragraphId} not found`)
    return { ...current }
  }
  getText(): string { return this.getInfo().text }
  getTextRange(): FDocumentTextRange {
    const info = this.getInfo()
    return this.document.getTextRange(info.startOffset, info.endOffset)
  }
  setText(text: string): boolean { return this.getTextRange().setText(text) }
  appendText(text: string): boolean { return this.document.insertText(this.getInfo().endOffset, text) }
  getTextStyle(): DocumentTextStyle { return this.getTextRange().getTextStyle() }
  setTextStyle(style: DocumentTextStylePatch): boolean { return this.getTextRange().setTextStyle(style) }
  getParagraphStyle(): DocumentParagraphStyle {
    const style = getDocumentParagraphStyle(this.document.getModel(), this.identity())
    if (!style) throw new Error(`Document paragraph with id ${this.info.paragraphId} not found`)
    return style
  }
  setParagraphStyle(style: Partial<DocumentParagraphStyle>): boolean {
    return this.document.mutate('Set paragraph style', 'doc.command.set-paragraph-style', (model) => {
      setDocumentParagraphStyle(model, this.identity(), style)
    })
  }

  private identity(): DocumentParagraphIdentity {
    return { paragraphId: this.info.paragraphId }
  }
}

export class FBlob {
  private value: Blob | undefined
  constructor(value?: Blob | Uint8Array | string, contentType = 'text/plain') {
    if (value instanceof Blob) this.value = value
    else if (typeof value === 'string') this.value = new Blob([value], { type: contentType })
    else if (value) this.setBytes(value)
  }
  copyBlob(): FBlob { return new FBlob(this.value) }
  getAs(contentType: string): FBlob { const blob = this.copyBlob(); blob.setContentType(contentType); return blob }
  getDataAsString(charset?: string): Promise<string> {
    if (!this.value) return Promise.resolve('')
    if (!charset) return this.value.text()
    return this.value.arrayBuffer().then((buffer) => new TextDecoder(charset).decode(buffer))
  }
  getBytes(): Promise<Uint8Array> {
    if (!this.value) return Promise.reject(new Error('Blob is undefined or null.'))
    return this.value.arrayBuffer().then((buffer) => new Uint8Array(buffer))
  }
  setBytes(bytes: Uint8Array): this { this.value = new Blob([Uint8Array.from(bytes)]); return this }
  setDataFromString(data: string, contentType = 'text/plain'): this { this.value = new Blob([data], { type: contentType }); return this }
  getContentType(): string | undefined { return this.value?.type }
  setContentType(contentType: string): this { if (this.value) this.value = this.value.slice(0, this.value.size, contentType); return this }
  getBlob(): Blob { if (!this.value) throw new Error('Blob is undefined or null.'); return this.value }
  getSize(): number { return this.value?.size ?? 0 }
}

function normalizeRange(range: FRangeCoordinates): FRangeCoordinates {
  const next = { startRow: range.startRow, startColumn: range.startColumn, endRow: range.endRow, endColumn: range.endColumn }
  validateIndex(next.startRow, 'startRow')
  validateIndex(next.startColumn, 'startColumn')
  validateIndex(next.endRow, 'endRow')
  validateIndex(next.endColumn, 'endColumn')
  if (next.endRow < next.startRow || next.endColumn < next.startColumn) throw new RangeError('Range end must not precede its start')
  return next
}

function normalizePivotSourceInfo(source: PivotTableSourceInfo, workbookId: string): PivotTableSourceInfo {
  if (source.unitId !== workbookId) throw new Error('Pivot source must belong to the target workbook')
  const sheet = source.subUnitId
  if (!sheet || !source.sheetName) throw new TypeError('Pivot source worksheet identity is required')
  const range = source.range
  for (const [name, value] of Object.entries({
    startRow: range.startRow,
    startColumn: range.startColumn,
    endRow: range.endRow,
    endColumn: range.endColumn,
  })) validateIndex(value, `pivot source ${name}`)
  if (range.endRow < range.startRow || range.endColumn < range.startColumn) throw new RangeError('Invalid pivot source range')
  return { ...source, range: { ...range } }
}

function normalizePivotAnchorInfo(anchor: PivotTableAnchorInfo, workbookId: string): PivotTableAnchorInfo {
  if (anchor.unitId !== workbookId) throw new Error('Pivot anchor must belong to the target workbook')
  validateIndex(anchor.row, 'pivot anchor row')
  validateIndex(anchor.col, 'pivot anchor column')
  return { ...anchor }
}

function toCellRange(range: FRangeCoordinates): CellRange {
  const normalized = normalizeToCellRange(range)
  return normalized
}

function normalizeValidationRange(range: FRange | FRangeCoordinates | string): CellRange {
  if (typeof range === 'string') return toCellRange(parseA1Range(range))
  if (range instanceof FRange) return toCellRange(range.getRange())
  return toCellRange(range)
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  return left.startRow <= right.endRow && left.endRow >= right.startRow
    && left.startCol <= right.endCol && left.endCol >= right.startCol
}

function normalizeToCellRange(range: FRangeCoordinates): CellRange {
  const normalized = normalizeRange(range)
  return { startRow: normalized.startRow, startCol: normalized.startColumn, endRow: normalized.endRow, endCol: normalized.endColumn }
}

function parseA1Range(value: string): FRangeCoordinates {
  const refs = value.split('!').at(-1)?.split(':') ?? []
  if (refs.length < 1 || refs.length > 2) throw new TypeError(`Invalid A1 range: ${value}`)
  const first = parseA1Cell(refs[0]!)
  const last = parseA1Cell(refs[1] ?? refs[0]!)
  return normalizeRange({ startRow: first.row, startColumn: first.column, endRow: last.row, endColumn: last.column })
}

function parseChartA1Range(value: string): ChartSource['range'] {
  const range = parseA1Range(value)
  return { startRow: range.startRow, startCol: range.startColumn, endRow: range.endRow, endCol: range.endColumn }
}

function parseA1Cell(value: string): { row: number; column: number } {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(value.trim())
  if (!match) throw new TypeError(`Invalid A1 cell: ${value}`)
  let column = 0
  for (const letter of match[1]!.toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64
  const row = Number(match[2])
  if (row < 1) throw new RangeError(`Invalid A1 row: ${row}`)
  return { row: row - 1, column: column - 1 }
}

function cellAddress(row: number, column: number): string {
  let letters = ''
  for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + ((value - 1) % 26)) + letters
  return `${letters}${row + 1}`
}

function validateIndex(value: number, name: string, allowZero = false): void {
  if (!Number.isInteger(value) || value < (allowZero ? 0 : 0)) throw new RangeError(`${name} must be a non-negative integer`)
}

function validateSize(value: number, name: string): void {
  if (!Number.isInteger(value) || value < 1) throw new RangeError(`${name} must be a positive integer`)
}

function assertMatrixSize(values: unknown[][], rows: number, columns: number): void {
  if (values.length !== rows || values.some((row) => row.length !== columns)) throw new RangeError(`Expected a ${rows}x${columns} value matrix`)
}

function rangeContains(range: FRangeCoordinates, merged: { startRow: number; startCol: number; endRow: number; endCol: number }): boolean {
  return range.startRow <= merged.startRow && range.startColumn <= merged.startCol && range.endRow >= merged.endRow && range.endColumn >= merged.endCol
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' ? value as Record<string, unknown> : {}
}
