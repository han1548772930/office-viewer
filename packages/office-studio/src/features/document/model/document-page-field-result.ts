import type { DocumentPageField } from '../../../domain/model'
import { numberingOrdinal } from './document-numbering-format'

export function createDocumentNumericPageField(field: 'PAGE' | 'NUMPAGES', pageNumber: number, pageCount: number): Pick<DocumentPageField, 'field' | 'instruction' | 'cachedResult'> {
  if (field !== 'PAGE' && field !== 'NUMPAGES') throw new TypeError('页码字段类型无效。')
  const number = field === 'PAGE' ? pageNumber : pageCount
  if (!Number.isSafeInteger(number) || number < (field === 'PAGE' ? 0 : 1)) throw new RangeError('页码字段需要当前文档的有效排版结果。')
  return { field, instruction: field, cachedResult: String(number) }
}

/** Numeric page fields stay live; unsupported instructions retain their original cache. */
export function resolveDocumentPageFieldResult(field: DocumentPageField, pageNumber: number, pageCount?: number): string {
  const cached = field.cachedResult ?? ''
  if (field.field !== 'PAGE' && field.field !== 'NUMPAGES') return cached
  const number = field.field === 'PAGE' ? pageNumber : pageCount
  if (number === undefined || !Number.isSafeInteger(number) || number < 0) return cached
  const instruction = field.instruction
  if (instruction === undefined || !instruction.trim()) return String(number)
  const prefix = /^\s*(PAGE|NUMPAGES)\b/i.exec(instruction)
  if (!prefix || prefix[1]?.toUpperCase() !== field.field) return cached
  let cursor = prefix[0].length, result = String(number)
  for (const match of instruction.matchAll(/\\\*\s*(?:"([^"]*)"|([^\s\\]+))|\\([^\s\\]+)/g)) {
    const start = match.index ?? cursor
    if (start < cursor || instruction.slice(cursor, start).trim() || match[3] !== undefined) return cached
    const format = match[1] ?? match[2] ?? '', key = format.toUpperCase()
    if (key === 'ARABIC') result = String(number)
    else if (key === 'ROMAN' || key === 'ALPHABETIC') {
      const numberFormat = key === 'ROMAN' ? format === 'roman' ? 'lowerRoman' : 'upperRoman'
        : format === 'alphabetic' ? 'lowerLetter' : 'upperLetter'
      const ordinal = numberingOrdinal(number, numberFormat)
      if (!ordinal.recognized) return cached
      result = ordinal.text
    } else if (key === 'UPPER') result = result.toUpperCase()
    else if (key === 'LOWER') result = result.toLowerCase()
    else if (key !== 'MERGEFORMAT' && key !== 'CHARFORMAT') return cached
    cursor = start + match[0].length
  }
  return instruction.slice(cursor).trim() ? cached : result
}
