import type { CellScalar } from '../../domain/model.ts'
import { inferSheetCellValueType, isForcedSheetCellValueType, normalizeSheetCellValueType, type SheetCellValueType } from '../../domain/sheet-cell-value.ts'
import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { sheetCellNumberPattern, sheetCellNumberValue } from './sheet-number-presentation.ts'

export interface SheetCellEditContext {
  documentId?: string
  sheetId: string
  key: string
  date1904: boolean
  cell?: ExcelCellView
}

interface SheetCellEditContent {
  formula: string | null
  value: CellScalar
  text: string
  type: SheetCellValueType
  forcedText: boolean
  numberFormat: string
  /** Optional only for compatibility with sessions created before run-level conflict tracking. */
  richText?: ReadonlyArray<Record<string, unknown>> | null
}

export interface SheetCellEditSession {
  readonly sessionId?: string
  readonly documentId?: string
  readonly sheetId: string
  readonly key: string
  readonly date1904: boolean
  readonly region: string
  readonly currentYear: number
  /** F2/formula-bar edits retain run formatting; typing into the grid replaces the prior rich content. */
  readonly mode?: 'edit' | 'overwrite'
  readonly source: Readonly<SheetCellEditContent>
}

export type SheetCellEditConflict = 'document' | 'sheet' | 'cell' | 'date-system' | 'content' | 'number-format' | 'rich-text'

let nextSessionId = 0

/** A sheet override and an inherited workbook setting must have the same meaning in drawing and input. */
export function resolveSheetDate1904(sheet: { date1904?: boolean } | undefined, workbook: { date1904?: boolean }): boolean {
  return sheet?.date1904 ?? workbook.date1904 ?? false
}

function editContent(cell: ExcelCellView | undefined): SheetCellEditContent {
  const numberFormat = sheetCellNumberPattern(cell) || 'General'
  // A formula's value, display and declared cache type can change during recalculation without changing its source.
  if (cell?.formula) return { formula: cell.formula, value: null, text: '', type: 'formula', forcedText: false, numberFormat, richText: null }
  const declaredType = normalizeSheetCellValueType(cell?.dataType)
  const scalar = sheetCellNumberValue(cell)
  const value = scalar === undefined || (cell?.rawValue === undefined && cell?.value === '' && declaredType !== 'text' && declaredType !== 'error') ? null : scalar
  return { formula: null, value, text: cell?.value ?? '', type: declaredType === 'formula' || declaredType === undefined ? inferSheetCellValueType(value) : declaredType,
    forcedText: isForcedSheetCellValueType(cell?.dataType), numberFormat,
    richText: cell?.richText?.length ? cloneModel(cell.richText) : null }
}

/** Capture primitives rather than the live cell, so external mutations cannot rewrite an open draft's baseline. */
export function createSheetCellEditSession(context: SheetCellEditContext,
  options: { region: string; currentYear: number; mode?: 'edit' | 'overwrite' }): SheetCellEditSession {
  return { sessionId: globalThis.crypto?.randomUUID?.() ?? `sheet-input-${Date.now().toString(36)}-${++nextSessionId}`,
    documentId: context.documentId, sheetId: context.sheetId, key: context.key.toUpperCase(), date1904: context.date1904,
    region: options.region, currentYear: options.currentYear, mode: options.mode ?? 'edit', source: editContent(context.cell) }
}

/** Unrelated cells, visual styles and formula caches may change; edited content and interpretation may not. */
export function getSheetCellEditConflict(session: SheetCellEditSession, context: SheetCellEditContext): SheetCellEditConflict | undefined {
  if (context.documentId !== session.documentId) return 'document'
  if (context.sheetId !== session.sheetId) return 'sheet'
  if (context.key.toUpperCase() !== session.key) return 'cell'
  if (context.date1904 !== session.date1904) return 'date-system'
  const content = editContent(context.cell)
  if (content.numberFormat !== session.source.numberFormat) return 'number-format'
  if (content.formula !== session.source.formula || !Object.is(content.value, session.source.value)
    || content.text !== session.source.text || content.type !== session.source.type || content.forcedText !== session.source.forcedText) return 'content'
  if (session.source.richText !== undefined && JSON.stringify(content.richText) !== JSON.stringify(session.source.richText)) return 'rich-text'
  return undefined
}
