import type { CellScalar } from './model.ts'
import { parseExactDecimalInput } from '../shared/decimal-input.ts'

export type SheetCellValueType = 'text' | 'number' | 'boolean' | 'error' | 'blank' | 'formula'
export type SheetCellTypeInput = string | 1 | 2 | 3 | 4 | null | undefined

/** Public scalar types and Office storage aliases share one interpretation. */
export function normalizeSheetCellValueType(type: SheetCellTypeInput): SheetCellValueType | undefined {
  if (typeof type === 'number') return type === 1 || type === 4 ? 'text' : type === 2 ? 'number' : type === 3 ? 'boolean' : undefined
  switch (type?.trim().toLowerCase()) {
    case 's': case 'str': case 'string': case 'text': case 'inlinestr':
    case 'force_string': case 'force-string': case 'forcestring': return 'text'
    case 'n': case 'number': case 'numeric': case 'date': case 'd': return 'number'
    case 'b': case 'bool': case 'boolean': return 'boolean'
    case 'e': case 'error': return 'error'
    case 'blank': case 'empty': return 'blank'
    case 'f': case 'formula': return 'formula'
    default: return undefined
  }
}

export function inferSheetCellValueType(value: CellScalar): SheetCellValueType {
  return value === null ? 'blank' : typeof value === 'string' ? 'text' : typeof value === 'boolean' ? 'boolean' : 'number'
}

export function isForcedSheetCellValueType(type: SheetCellTypeInput): boolean {
  return type === 4 || typeof type === 'string' && /^(?:force_string|force-string|forcestring)$/i.test(type.trim())
}

/** Parse editable/cached text only when its declared type allows it; displays are never inputs. */
export function parseSheetCellText(value: string, type?: SheetCellTypeInput): CellScalar {
  const kind = normalizeSheetCellValueType(type)
  if (kind === 'text' || kind === 'error' || kind === 'blank' || kind === 'formula') return value
  const source = value.trim()
  if (kind === 'boolean') {
    if (/^(?:true|1)$/i.test(source)) return true
    if (/^(?:false|0)$/i.test(source)) return false
    return value
  }
  return parseExactDecimalInput(source) ?? value
}
