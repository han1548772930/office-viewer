import type { SheetCellStylePatch } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { sheetBorderStyleNumber } from './sheet-cell-border.ts'

export interface SheetPublicColorStyle extends Record<string, unknown> { rgb?: string | null }
export interface SheetPublicTextDecoration extends Record<string, unknown> { s: 0 | 1 }
export interface SheetPublicBorderStyle extends Record<string, unknown> { s?: number; cl?: SheetPublicColorStyle }
/** Fixed Univer short fields; unknown source extensions remain readable. */
export interface SheetPublicStyleData extends Record<string, unknown> {
  ff?: string
  fs?: number
  bl?: 0 | 1
  it?: 0 | 1
  ul?: SheetPublicTextDecoration
  st?: SheetPublicTextDecoration
  cl?: SheetPublicColorStyle
  bg?: SheetPublicColorStyle
  ht?: number
  vt?: number
  tb?: number
  stf?: 0 | 1
  tr?: Record<string, unknown> & { a?: number; v?: 0 | 1 }
  n?: Record<string, unknown> & { pattern: string }
  bd?: Record<string, unknown>
}
/** Nested fields are retained as a local compatibility extension of public reads. */
export type SheetApiStyleData = SheetCellStylePatch & SheetPublicStyleData

const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)
const sourceObject = (value: unknown): Record<string, unknown> => object(value) ? cloneModel(value) : {}
const shortKeys = ['ff', 'fs', 'bl', 'it', 'ul', 'st', 'cl', 'bg', 'ht', 'vt', 'tb', 'stf', 'tr', 'n', 'bd'] as const
const borderNames: Readonly<Record<string, string>> = {
  top: 't', right: 'r', bottom: 'b', left: 'l', diagonalDown: 'tl_br', topLeftBottomCenter: 'tl_bc',
  topLeftMiddleRight: 'tl_mr', diagonalUp: 'bl_tr', middleLeftTopRight: 'ml_tr', bottomCenterTopRight: 'bc_tr',
}
const horizontal = ['general', 'left', 'center', 'right', 'justify', 'justify', 'justify'] as const
const vertical = ['middle', 'top', 'middle', 'bottom'] as const
const wrapping = ['overflow', 'overflow', 'clip', 'wrap'] as const

function preservedEnum(source: unknown, values: readonly string[], value: string, fallback: number): number {
  return typeof source === 'number' && Number.isSafeInteger(source) && source >= 0 && source < values.length && values[source] === value ? source : fallback
}
function publicColor(source: unknown, rgb: string | null): SheetPublicColorStyle {
  return { ...sourceObject(source), rgb }
}
function publicBorders(border: Record<string, unknown>, source: unknown): NonNullable<SheetPublicStyleData['bd']> {
  const original = object(source) ? source : {}, result: NonNullable<SheetPublicStyleData['bd']> = {}
  for (const [name, value] of Object.entries(border)) {
    if (value === undefined) continue
    const key = Object.hasOwn(borderNames, name) ? borderNames[name]! : name
    if (value === null || !object(value) || !Object.hasOwn(borderNames, name)) { result[key] = cloneModel(value); continue }
    const { style, color, ...edge } = value
    const output: SheetPublicBorderStyle = { ...sourceObject(original[key]), ...cloneModel(edge) }
    if (style !== undefined || value.s !== undefined) output.s = sheetBorderStyleNumber(style ?? value.s)
    if (color !== undefined) output.cl = publicColor(value.cl ?? output.cl, color as string | null)
    result[key] = output
  }
  return result
}

/** Only present modeled properties emit short fields; cleared properties cannot reappear from source metadata. */
export function sheetStyleToPublicData(style: SheetCellStylePatch): SheetPublicStyleData {
  const source = sourceObject(style.metadata?.sourceStyle), result: SheetPublicStyleData = cloneModel(source)
  for (const key of shortKeys) delete result[key]
  const font = style.font
  if (font?.family !== undefined) result.ff = font.family
  if (font?.size !== undefined) result.fs = font.size
  if (font?.bold !== undefined) result.bl = font.bold ? 1 : 0
  if (font?.italic !== undefined) result.it = font.italic ? 1 : 0
  if (font?.underline !== undefined) result.ul = { ...sourceObject(source.ul), s: font.underline ? 1 : 0 }
  if (font?.strike !== undefined) result.st = { ...sourceObject(source.st), s: font.strike ? 1 : 0 }
  if (font?.color !== undefined) result.cl = publicColor(source.cl, font.color)
  else if (object(source.cl) && source.cl.th !== undefined) result.cl = sourceObject(source.cl)
  if (style.fill !== undefined) result.bg = publicColor(source.bg, style.fill)
  else if (object(source.bg) && source.bg.th !== undefined) result.bg = sourceObject(source.bg)
  if (style.horizontal !== undefined) result.ht = preservedEnum(source.ht, horizontal, style.horizontal, horizontal.indexOf(style.horizontal))
  if (style.vertical !== undefined) result.vt = preservedEnum(source.vt, vertical, style.vertical, style.vertical === 'top' ? 1 : style.vertical === 'middle' ? 2 : 3)
  const strategy = style.wrapStrategy ?? (style.wrap === undefined ? undefined : style.wrap ? 'wrap' : 'overflow')
  if (strategy !== undefined) result.tb = preservedEnum(source.tb, wrapping, strategy, strategy === 'overflow' ? 1 : strategy === 'clip' ? 2 : 3)
  if (style.shrinkToFit !== undefined) result.stf = style.shrinkToFit ? 1 : 0
  if (style.rotation !== undefined || style.rotationVertical !== undefined) {
    result.tr = sourceObject(source.tr); delete result.tr.a; delete result.tr.v
    if (style.rotation !== undefined) result.tr.a = style.rotation
    if (style.rotationVertical !== undefined) result.tr.v = style.rotationVertical ? 1 : 0
  }
  if (style.numberFormat !== undefined) result.n = { ...sourceObject(source.n), pattern: style.numberFormat }
  if (style.border !== undefined) result.bd = publicBorders(style.border, source.bd)
  return result
}

export function sheetStyleDataWithPublicFields(style: SheetCellStylePatch): SheetApiStyleData {
  const publicData = sheetStyleToPublicData(style), result: SheetApiStyleData = { ...cloneModel(style), ...publicData }
  if (style.font) result.font = { ...sourceObject(publicData.font), ...cloneModel(style.font) }
  return result
}
