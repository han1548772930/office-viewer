import type {
  CellScalar,
  DocumentModel,
  DocumentParagraphStyle,
  PivotAggregation,
  PivotFieldArea,
  PivotFieldFilter,
  PivotTableModel,
  PivotTableOptions,
  SpreadsheetCell,
  SpreadsheetSheet,
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

export interface CreateUnitOptions { makeCurrent?: boolean }
export interface FRangeCoordinates { startRow: number; startColumn: number; endRow: number; endColumn: number }
export type PivotTableSourceInfo = PivotTableSourceInput
export type PivotTableAnchorInfo = PivotTableAnchorInput

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
  getPivotTableByCell(row: number, col: number): FPivotTable | null {
    const pivot = getPivotTableAtCell(this.workbook(), this.id, row, col)
    return pivot ? new FPivotTable(this.api, this.unitId, pivot.id) : null
  }
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
  getCellData(): SpreadsheetCell | null { const cell = this.cellAt(this.range.startRow, this.range.startColumn); return cell ? cloneModel(cell) : null }
  getCellDatas(): (SpreadsheetCell | null)[][] { return this.mapCells((cell) => cell ? cloneModel(cell) : null) }
  setValue(value: ApiCellValue): this { return this.setValues(this.fill(value)) }
  setValues(values: ApiCellValue[][]): this {
    assertMatrixSize(values, this.getHeight(), this.getWidth())
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
  replaceRange(startOffset: number, endOffset: number, text: string): boolean {
    return this.api.executeCommandSync('doc.command.replace-range', { unitId: this.id, startOffset, endOffset, text })
  }
  mutate(label: string, commandId: string, update: (model: DocumentModel) => void): boolean {
    return this.api.mutateDocument(this.id, commandId, label, update)
  }

  private model(): DocumentModel {
    const model = this.api.getDocumentModel(this.id)
    if (!model) throw new Error(`Document not found: ${this.id}`)
    return model
  }
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

function parseA1Range(value: string): FRangeCoordinates {
  const refs = value.split('!').at(-1)?.split(':') ?? []
  if (refs.length < 1 || refs.length > 2) throw new TypeError(`Invalid A1 range: ${value}`)
  const first = parseA1Cell(refs[0]!)
  const last = parseA1Cell(refs[1] ?? refs[0]!)
  return normalizeRange({ startRow: first.row, startColumn: first.column, endRow: last.row, endColumn: last.column })
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
