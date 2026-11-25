import type { FontStyle } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { sheetCellRichTextMatches } from './sheet-cell-rich-text-edit.ts'

export type SheetRichTextFontField = keyof FontStyle
export type SheetRichTextFontValue = FontStyle[SheetRichTextFontField] | null
const aliases: Record<SheetRichTextFontField, string> = {
  family: 'fontFamily', size: 'fontSize', bold: 'bold', italic: 'italic', underline: 'underline', strike: 'strike', color: 'color',
}
function fontRecord(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : undefined
}

/** Public runs may contain top-level font aliases or a nested font; retain their unknown data. */
export function sheetCellRichTextFontValue(run: Record<string, unknown>, field: SheetRichTextFontField): unknown {
  return run[aliases[field]] ?? fontRecord(run.font)?.[field]
}

export function sheetCellRichTextRunView(cell: ExcelCellView, run: Record<string, unknown>): ExcelCellView {
  const result = { ...cell }
  const family = sheetCellRichTextFontValue(run, 'family'), size = sheetCellRichTextFontValue(run, 'size')
  if (typeof family === 'string' && family.trim()) result.fontFamily = family
  if (typeof size === 'number' && Number.isFinite(size) && size > 0) result.fontSize = size
  for (const field of ['bold', 'italic', 'underline', 'strike'] as const) {
    const value = sheetCellRichTextFontValue(run, field)
    if (typeof value === 'boolean' || value === 0 || value === 1) result[field] = value === true || value === 1
  }
  const color = sheetCellRichTextFontValue(run, 'color')
  if (typeof color === 'string') result.color = color
  if (typeof run.hyperlink === 'string') result.hyperlink = run.hyperlink
  return result
}

/** Selection state and toggles read visible run styles instead of only the cell default. */
export function sheetCellFontValues<K extends SheetRichTextFontField>(cell: ExcelCellView | undefined, field: K,
  selection?: { start: number; end: number }): Array<FontStyle[K]> {
  const inherited: FontStyle = { family: cell?.fontFamily ?? 'Aptos', size: cell?.fontSize ?? 11,
    bold: cell?.bold ?? false, italic: cell?.italic ?? false, underline: cell?.underline ?? false,
    strike: cell?.strike ?? false, color: cell?.color ?? null }
  const fallback = inherited[field]
  if (!cell?.richText?.length || cell.formula || !sheetCellRichTextMatches(cell.richText, String(cell.rawValue ?? cell.value))) return [fallback]
  let offset = 0
  const values: Array<FontStyle[K]> = []
  for (const run of cell.richText) {
    const start = offset; offset += (run.text as string).length
    if (start === offset) continue
    if (selection && (selection.start === selection.end
      ? selection.start === 0 ? start !== 0 : !(start < selection.start && offset >= selection.start)
      : start >= selection.end || offset <= selection.start)) continue
    const view = sheetCellRichTextRunView(cell, run)
    const resolved: FontStyle = { family: view.fontFamily ?? inherited.family, size: view.fontSize ?? inherited.size,
      bold: view.bold ?? inherited.bold, italic: view.italic ?? inherited.italic, underline: view.underline ?? inherited.underline,
      strike: view.strike ?? inherited.strike, color: view.color ?? inherited.color }
    values.push(resolved[field])
  }
  return values.length ? values : [fallback]
}

/** Whole-cell font commands also replace the corresponding explicit run style, as handleStyle() does upstream. */
export function setSheetCellRichTextFont(runs: ReadonlyArray<Record<string, unknown>> | null | undefined,
  field: SheetRichTextFontField, value: SheetRichTextFontValue): { runs: ReadonlyArray<Record<string, unknown>> | null | undefined; changed: boolean } {
  if (!runs?.length) return { runs, changed: false }
  const alias = aliases[field]
  let changed = false
  const next = runs.map(run => {
    const nested = fontRecord(run.font)
    const ownsAlias = Object.prototype.hasOwnProperty.call(run, alias)
    const ownsNested = nested && Object.prototype.hasOwnProperty.call(nested, field)
    if (value === null ? !ownsAlias && !ownsNested
      : ownsAlias ? run[alias] === value && (!ownsNested || nested?.[field] === value) : ownsNested && nested?.[field] === value) return run
    const result = cloneModel(run)
    if (value === null) {
      delete result[alias]
      if (nested) {
        const font = cloneModel(nested); delete font[field]
        if (Object.keys(font).length) result.font = font
        else delete result.font
      }
    } else if (nested || run.font === undefined) {
      result.font = { ...cloneModel(nested ?? {}), [field]: value }
      if (ownsAlias) result[alias] = value
    } else {
      // A malformed imported font extension is preserved; the public alias supplies the authored style.
      result[alias] = value
    }
    changed = true
    return result
  })
  return { runs: changed ? next : runs, changed }
}

/** Clear modeled font properties in one pass while retaining text, links, identities and unknown fields. */
export function clearSheetCellRichTextFonts(runs: ReadonlyArray<Record<string, unknown>> | null | undefined): {
  runs: ReadonlyArray<Record<string, unknown>> | null | undefined; changed: boolean
} {
  if (!runs?.length) return { runs, changed: false }
  const fields = Object.keys(aliases) as SheetRichTextFontField[]
  let changed = false
  const next = runs.map(run => {
    const nested = fontRecord(run.font)
    if (!fields.some(field => Object.hasOwn(run, aliases[field]) || !!nested && Object.hasOwn(nested, field))) return run
    const result = cloneModel(run), font = nested ? cloneModel(nested) : undefined
    for (const field of fields) { delete result[aliases[field]]; if (font) delete font[field] }
    if (font) { if (Object.keys(font).length) result.font = font; else delete result.font }
    changed = true
    return result
  })
  return { runs: changed ? next : runs, changed }
}
