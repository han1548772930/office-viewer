import { transformSheetAxisReferenceInterval, type SheetAxisEdit } from '../../domain/sheet-axis-structure.ts'

const MAX_ROW = 1_048_575, MAX_COLUMN = 16_383
export interface FormulaReferenceArea {
  kind: 'cell' | 'rectangle' | 'wholeRow' | 'wholeColumn'
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
  absoluteStartRow: boolean
  absoluteEndRow: boolean
  absoluteStartColumn: boolean
  absoluteEndColumn: boolean
  qualifier: string
  endQualifier: string
  sheetName?: string
  endSheetName?: string
  external: boolean
  threeDimensional: boolean
}
interface Qualifier { text: string; sheetName: string; end: number; external: boolean; threeDimensional: boolean }
interface Address { row: number; column: number; absoluteRow: boolean; absoluteColumn: boolean; end: number }
interface AreaSpan { start: number; end: number; area: FormulaReferenceArea }
export interface FormulaReferenceAreaMatch { end: number; area: FormulaReferenceArea }

const identifier = (value: string | undefined): boolean => value !== undefined && /[\p{L}\p{N}_.\\]/u.test(value)
const columnIndex = (letters: string): number => [...letters.toUpperCase()].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1
function columnLabel(column: number): string {
  let result = '', value = column + 1
  while (value > 0) { result = String.fromCharCode(65 + (value - 1) % 26) + result; value = Math.floor((value - 1) / 26) }
  return result
}
function readQualifier(source: string, start: number): Qualifier | undefined {
  let text: string | undefined, end = start
  if (source[start] === "'") {
    end += 1
    while (end < source.length) {
      if (source[end] !== "'") { end += 1; continue }
      if (source[end + 1] === "'") { end += 2; continue }
      end += 1
      if (source[end] === '!') text = source.slice(start + 1, end - 1).replace(/''/g, "'")
      break
    }
  } else {
    const matched = /^(?:\[[^\]\r\n]+\])?[\p{L}\p{N}_\\][\p{L}\p{N}_.\\]*(?::[\p{L}\p{N}_\\][\p{L}\p{N}_.\\]*)?!/u.exec(source.slice(start))
    if (matched) { end = start + matched[0].length - 1; text = matched[0].slice(0, -1) }
  }
  if (text === undefined) return undefined
  const external = text.includes('['), withoutBook = text.slice(text.lastIndexOf(']') + 1)
  return { text: source.slice(start, end + 1), sheetName: text, end: end + 1, external, threeDimensional: withoutBook.includes(':') }
}
function readCell(source: string, start: number): Address | undefined {
  const matched = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(source.slice(start))
  if (!matched) return undefined
  const end = start + matched[0].length, row = Number(matched[4]) - 1, column = columnIndex(matched[2]!)
  if (identifier(source[end]) || source[end] === '$' || source[end] === '!' || /^\s*\(/.test(source.slice(end))
    || row > MAX_ROW || column > MAX_COLUMN) return undefined
  return { row, column, absoluteRow: matched[3] === '$', absoluteColumn: matched[1] === '$', end }
}
function readArea(source: string, start: number): AreaSpan | undefined {
  const qualifier = readQualifier(source, start), addressStart = qualifier?.end ?? start
  const base = { qualifier: qualifier?.text ?? '', endQualifier: '', sheetName: qualifier?.sheetName,
    external: qualifier?.external ?? false, threeDimensional: qualifier?.threeDimensional ?? false }
  const first = readCell(source, addressStart)
  if (first) {
    const separator = /^\s*:\s*/.exec(source.slice(first.end))
    const secondStart = separator ? first.end + separator[0].length : undefined
    const secondQualifier = secondStart === undefined ? undefined : readQualifier(source, secondStart)
    const second = secondStart === undefined ? undefined : readCell(source, secondQualifier?.end ?? secondStart)
    return { start, end: second?.end ?? first.end, area: { ...base, kind: second ? 'rectangle' : 'cell',
      startRow: first.row, endRow: second?.row ?? first.row, startColumn: first.column, endColumn: second?.column ?? first.column,
      absoluteStartRow: first.absoluteRow, absoluteEndRow: second?.absoluteRow ?? first.absoluteRow,
      absoluteStartColumn: first.absoluteColumn, absoluteEndColumn: second?.absoluteColumn ?? first.absoluteColumn,
      endQualifier: secondQualifier?.text ?? '', endSheetName: secondQualifier?.sheetName,
      external: base.external || (secondQualifier?.external ?? false), threeDimensional: base.threeDimensional || (secondQualifier?.threeDimensional ?? false) } }
  }
  const firstAxis = /^(\$?)([A-Za-z]{1,3}|[1-9]\d*)\s*:\s*/.exec(source.slice(addressStart))
  if (!firstAxis) return undefined
  const secondStart = addressStart + firstAxis[0].length, secondQualifier = readQualifier(source, secondStart)
  const secondAxisStart = secondQualifier?.end ?? secondStart
  const secondAxis = /^(\$?)([A-Za-z]{1,3}|[1-9]\d*)/.exec(source.slice(secondAxisStart))
  if (!secondAxis) return undefined
  const end = secondAxisStart + secondAxis[0].length
  if (identifier(source[end]) || source[end] === '$' || /^\s*\(/.test(source.slice(end))) return undefined
  const row = /^\d/.test(firstAxis[2]!), secondRow = /^\d/.test(secondAxis[2]!)
  if (row !== secondRow) return undefined
  const firstIndex = row ? Number(firstAxis[2]) - 1 : columnIndex(firstAxis[2]!)
  const secondIndex = row ? Number(secondAxis[2]) - 1 : columnIndex(secondAxis[2]!)
  if (firstIndex > (row ? MAX_ROW : MAX_COLUMN) || secondIndex > (row ? MAX_ROW : MAX_COLUMN)) return undefined
  return { start, end, area: { ...base, kind: row ? 'wholeRow' : 'wholeColumn',
    startRow: row ? firstIndex : 0, endRow: row ? secondIndex : MAX_ROW,
    startColumn: row ? 0 : firstIndex, endColumn: row ? MAX_COLUMN : secondIndex,
    absoluteStartRow: row && firstAxis[1] === '$', absoluteEndRow: row && secondAxis[1] === '$',
    absoluteStartColumn: !row && firstAxis[1] === '$', absoluteEndColumn: !row && secondAxis[1] === '$',
    endQualifier: secondQualifier?.text ?? '', endSheetName: secondQualifier?.sheetName,
    external: base.external || (secondQualifier?.external ?? false), threeDimensional: base.threeDimensional || (secondQualifier?.threeDimensional ?? false) } }
}
function areaSpans(source: string): AreaSpan[] {
  const spans: AreaSpan[] = []
  let index = 0
  while (index < source.length) {
    if (source[index] === '"') {
      index += 1
      while (index < source.length) {
        if (source[index] !== '"') { index += 1; continue }
        if (source[index + 1] === '"') { index += 2; continue }
        index += 1; break
      }
      continue
    }
    if (!identifier(source[index - 1])) {
      const span = readArea(source, index)
      if (span) { spans.push(span); index = span.end; continue }
    }
    if (source[index] === '[') {
      let nesting = 1; index += 1
      while (index < source.length && nesting) { if (source[index] === '[') nesting++; else if (source[index] === ']') nesting--; index++ }
    } else if (source[index] === "'") {
      index += 1
      while (index < source.length) {
        if (source[index] !== "'") { index++; continue }
        if (source[index + 1] === "'") { index += 2; continue }
        index++; break
      }
    } else index += 1
  }
  return spans
}

/** Shared lexical boundary for formula parsing and reference-only transformations. */
export function readFormulaReferenceArea(source: string, start: number): FormulaReferenceAreaMatch | undefined {
  if (!Number.isSafeInteger(start) || start < 0 || start >= source.length || identifier(source[start - 1])) return undefined
  const span = readArea(source, start)
  return span ? { end: span.end, area: { ...span.area } } : undefined
}

function formatArea(area: FormulaReferenceArea): string {
  const startRow = `${area.absoluteStartRow ? '$' : ''}${area.startRow + 1}`, endRow = `${area.absoluteEndRow ? '$' : ''}${area.endRow + 1}`
  const startColumn = `${area.absoluteStartColumn ? '$' : ''}${columnLabel(area.startColumn)}`, endColumn = `${area.absoluteEndColumn ? '$' : ''}${columnLabel(area.endColumn)}`
  if (area.kind === 'wholeRow') return `${area.qualifier}${startRow}:${area.endQualifier}${endRow}`
  if (area.kind === 'wholeColumn') return `${area.qualifier}${startColumn}:${area.endQualifier}${endColumn}`
  const first = `${area.qualifier}${startColumn}${startRow}`
  return area.kind === 'cell' ? first : `${first}:${area.endQualifier}${endColumn}${endRow}`
}

/** Signed offsets retain recoverable references, including complete rows/columns and qualified rectangles. */
export function formulaReferenceAreaSignature(formula: string, row: number, column: number): string {
  if (typeof formula !== 'string' || !Number.isSafeInteger(row) || !Number.isSafeInteger(column)) throw new TypeError('Invalid formula reference origin')
  const parts: Array<string | FormulaReferenceArea> = []
  let cursor = 0
  for (const span of areaSpans(formula)) {
    const area = span.area
    parts.push(formula.slice(cursor, span.start), {
      ...area,
      startRow: area.kind === 'wholeColumn' || area.absoluteStartRow ? area.startRow : area.startRow - row,
      endRow: area.kind === 'wholeColumn' || area.absoluteEndRow ? area.endRow : area.endRow - row,
      startColumn: area.kind === 'wholeRow' || area.absoluteStartColumn ? area.startColumn : area.startColumn - column,
      endColumn: area.kind === 'wholeRow' || area.absoluteEndColumn ? area.endColumn : area.endColumn - column,
    })
    cursor = span.end
  }
  parts.push(formula.slice(cursor))
  return JSON.stringify(parts)
}

/** Strings, structured column identifiers and literal sheet qualifiers are never transformed. */
export function mapFormulaReferenceAreas(formula: string, transform: (area: Readonly<FormulaReferenceArea>) => FormulaReferenceArea | null): string {
  if (typeof formula !== 'string') throw new TypeError('公式必须是字符串。')
  const patches: Array<{ start: number; end: number; text: string }> = []
  for (const span of areaSpans(formula)) {
    const changed = transform({ ...span.area })
    if (changed !== null && JSON.stringify(changed) === JSON.stringify(span.area)) continue
    if (changed && (![changed.startRow, changed.endRow, changed.startColumn, changed.endColumn].every(Number.isSafeInteger)
      || changed.startRow < 0 || changed.endRow < 0 || changed.startColumn < 0 || changed.endColumn < 0
      || changed.startRow > MAX_ROW || changed.endRow > MAX_ROW || changed.startColumn > MAX_COLUMN || changed.endColumn > MAX_COLUMN)) {
      throw new RangeError('变换后的公式引用超出工作表。')
    }
    patches.push({ start: span.start, end: span.end, text: changed === null ? `${span.area.qualifier}#REF!` : formatArea(changed) })
  }
  let result = formula
  for (const patch of patches.reverse()) result = result.slice(0, patch.start) + patch.text + result.slice(patch.end)
  return result
}

export function translateFormulaReferenceAreas(formula: string, rowDelta: number, columnDelta: number): string {
  if (!Number.isSafeInteger(rowDelta) || !Number.isSafeInteger(columnDelta)) throw new TypeError('公式复制偏移无效。')
  return mapFormulaReferenceAreas(formula, area => {
    const next = { ...area }
    if (area.kind !== 'wholeColumn') { if (!area.absoluteStartRow) next.startRow += rowDelta; if (!area.absoluteEndRow) next.endRow += rowDelta }
    if (area.kind !== 'wholeRow') { if (!area.absoluteStartColumn) next.startColumn += columnDelta; if (!area.absoluteEndColumn) next.endColumn += columnDelta }
    return next.startRow < 0 || next.endRow < 0 || next.startColumn < 0 || next.endColumn < 0
      || next.startRow > MAX_ROW || next.endRow > MAX_ROW || next.startColumn > MAX_COLUMN || next.endColumn > MAX_COLUMN ? null : next
  })
}

export function transformFormulaReferenceAreasForAxisEdit(formula: string, edit: SheetAxisEdit, formulaSheetName: string | undefined, targetSheetName: string): string {
  return mapFormulaReferenceAreas(formula, area => {
    const owner = area.sheetName ?? formulaSheetName, endpointOwner = area.endSheetName ?? owner
    if (area.external || area.threeDimensional || owner?.toLowerCase() !== targetSheetName.toLowerCase()
      || endpointOwner?.toLowerCase() !== targetSheetName.toLowerCase()
      || edit.axis === 'row' && area.kind === 'wholeColumn' || edit.axis === 'column' && area.kind === 'wholeRow') return { ...area }
    const first = edit.axis === 'row' ? area.startRow : area.startColumn, last = edit.axis === 'row' ? area.endRow : area.endColumn
    const interval = transformSheetAxisReferenceInterval({ start: Math.min(first, last), end: Math.max(first, last) }, edit)
    if (!interval) return null
    const start = first <= last ? interval.start : interval.end, end = first <= last ? interval.end : interval.start
    return edit.axis === 'row' ? { ...area, startRow: start, endRow: end } : { ...area, startColumn: start, endColumn: end }
  })
}
