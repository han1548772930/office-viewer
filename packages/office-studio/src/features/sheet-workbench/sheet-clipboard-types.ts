import type { CellRange, SheetConditionalFormatRule, SheetDataValidationRule, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import type { HistorySavepoint } from '../../platform/history/history-types.ts'
import type { SheetValidationIssue } from '../sheet-validation/validation-engine.ts'
import type { SheetPasteMode } from './sheet-paste.ts'
import type { SheetClipboardPasteType } from './sheet-clipboard-command.ts'
import type { SheetStyleRegionSnapshot } from './sheet-style-region.ts'
import type { SheetImageClipboardEntry } from '../sheet-images/sheet-image-clipboard.ts'

export interface SheetClipboardLocation {
  unitId: string
  sheetId: string
  range: CellRange
}

/** Relative sparse cells and source coordinates are serializable clipboard data, separate from the editor view. */
export interface SheetClipboardSnapshot {
  version: 1
  id: string
  copyType: 'copy' | 'cut'
  formulaOnly: boolean
  rows: number[]
  columns: number[]
  cells: Record<string, SpreadsheetCell>
  /** Captured inherited formatting for every relative blank cell; no live style IDs. */
  styleRegion?: SheetStyleRegionSnapshot
  merges: CellRange[]
  columnWidths: number[]
  conditionalFormats?: SheetConditionalFormatRule[]
  dataValidations?: SheetDataValidationRule[]
  /** Complete selected table/pivot objects and contained chart/filter drawings. */
  workbookResources?: SheetClipboardWorkbookResources
  /** Fully selected floating images and their original worksheet bounds. */
  images?: SheetImageClipboardEntry[]
  /** Header-inclusive, unzoomed source corner used for drawing placement after source closure. */
  sourceOrigin?: { left: number; top: number }
  /** Captured source viewport geometry, used to verify that copied drawings were fully selected. */
  sourceBounds?: { left: number; top: number; width: number; height: number }
  source?: SheetClipboardLocation & {
    sheetName: string
    revision?: number
    historyGeneration?: number
    fingerprint: string
  }
  /** External text preserves destination formatting unless explicit source HTML supplies cell styles. */
  externalFormat?: 'text' | 'html'
}

export type SheetClipboardFailure = 'empty-clipboard' | 'stale-target' | 'stale-source' | 'stale-paste' | 'cross-workbook-cut' | 'validation'
export interface SheetClipboardCopyResult {
  accepted: boolean
  clipboard?: SheetClipboardSnapshot
  imageClipboard?: import('../sheet-images/sheet-image-object-clipboard.ts').SheetImageObjectClipboard
  /** Focused image cut removes the selected object immediately, as in the fixed upstream. */
  changed?: boolean
  plain: string
  html: string
}
export interface SheetClipboardPasteResult {
  accepted: boolean
  changed: boolean
  ranges: CellRange[]
  issues: SheetValidationIssue[]
  reason?: SheetClipboardFailure
  imageIds?: string[]
}

export interface SheetClipboardPasteContext {
  userId: string
  region?: string
  sourceHistoryGeneration?: number
}
export type SheetClipboardWorkbookResources = Pick<WorkbookModel, 'tables' | 'pivotTables' | 'charts' | 'slicers'>
export interface SheetClipboardPastePlan {
  sheetChanges: SpreadsheetSheet[]
  /** Actual changed cell, range-resource and object locations, grouped by sheet for observers. */
  changedRanges?: Record<string, CellRange[]>
  workbookMetadata?: Record<string, unknown>
  metadataChanged: boolean
  workbookResources?: SheetClipboardWorkbookResources
  result: SheetClipboardPasteResult
}
export interface SheetClipboardRecentPaste {
  clipboard: SheetClipboardSnapshot
  target: SheetClipboardLocation
  ranges: CellRange[]
  mode: SheetPasteMode
  originalSheets: SpreadsheetSheet[]
  originalMetadata?: Record<string, unknown>
  metadataChanged: boolean
  originalResources?: SheetClipboardWorkbookResources
  savepoint: HistorySavepoint
  contentRevision: number
}

export interface SheetClipboardCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  range?: CellRange | { startRow: number; startColumn: number; endRow: number; endColumn: number }
  target?: { unitId: string; sheetId?: string; subUnitId?: string; range: NonNullable<SheetClipboardCommandInput['range']> }
  clipboardId?: string
  /** Focused drawing copy uses the same clipboard command and history owner as a cell-range copy. */
  imageIds?: readonly string[]
  text?: string
  html?: string
  textContent?: string
  htmlContent?: string
  /** Fixed Univer paste hook name, or the local mode. */
  value?: string
  mode?: SheetPasteMode
  /** Fixed Univer optional-paste enum key. */
  type?: SheetClipboardPasteType
  expectedRevision?: number
  expectedHistoryGeneration?: number
  /** UI captures this before awaiting the system clipboard; API range operations may omit it. */
  host?: { unitId: string; sheetId: string }
}
