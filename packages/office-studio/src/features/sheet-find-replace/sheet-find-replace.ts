import type { ExcelCellView, ExcelSheetView } from '../../services/excelViewAdapter.ts'

export type SheetFindBy = 'value' | 'formula'
export type SheetFindDirection = 'row' | 'column'
export type SheetFindScope = 'sheet' | 'workbook'

export interface SheetFindQuery {
  text: string
  replacement: string
  findBy: SheetFindBy
  direction: SheetFindDirection
  scope: SheetFindScope
  caseSensitive: boolean
  wholeCell: boolean
}

export interface SheetFindMatch {
  sheetId: string
  key: string
  replaceable: boolean
  formula: boolean
}

export interface SheetFindReplacePatch extends SheetFindMatch {
  cell: ExcelCellView
}

export function findSheetMatches(
  sheets: readonly ExcelSheetView[],
  activeSheetId: string,
  query: SheetFindQuery,
  selectedRanges: readonly { startRow: number; endRow: number; startColumn: number; endColumn: number }[] = [],
  searchInSelection = selectedRanges.some(range => range.startRow !== range.endRow || range.startColumn !== range.endColumn),
): SheetFindMatch[] {
  const text = query.caseSensitive ? query.text : query.text.toLocaleLowerCase()
  if (!text) return []
  if (query.scope === 'sheet' && searchInSelection && !selectedRanges.length) return []

  const searchSheets = query.scope === 'workbook'
    ? sheets.filter(sheet => sheet.visible !== false)
    : sheets.filter(sheet => sheet.id === activeSheetId)
  const selection = query.scope === 'sheet' && searchInSelection ? normalizeRanges(selectedRanges) : []
  const matches: SheetFindMatch[] = []

  for (const sheet of searchSheets) {
    const filteredRows = new Set(sheet.filteredRows ?? [])
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const address = parseKey(key)
      if (!address || filteredRows.has(address.row) || !inRanges(address, selection)) continue
      const formula = Boolean(cell.formula)
      const candidate = query.findBy === 'formula'
        ? formula ? cell.formula ?? '' : ''
        : formula ? cell.cachedValue ?? cell.displayValue ?? '' : rawCellValue(cell)
      if (!matchesText(candidate, text, query.caseSensitive, query.wholeCell)) continue
      matches.push({ sheetId: sheet.id, key, replaceable: query.findBy === 'formula' ? formula : !formula, formula })
    }
  }

  matches.sort((left, right) => {
    const sheetOrder = searchSheets.findIndex(sheet => sheet.id === left.sheetId) - searchSheets.findIndex(sheet => sheet.id === right.sheetId)
    if (sheetOrder !== 0) return sheetOrder
    const a = parseKey(left.key)!
    const b = parseKey(right.key)!
    return query.direction === 'row' ? a.row - b.row || a.column - b.column : a.column - b.column || a.row - b.row
  })
  return matches
}

export function nextSheetMatchIndex(
  matches: readonly SheetFindMatch[],
  sheetId: string,
  key: string,
  direction: 1 | -1,
  findDirection: SheetFindDirection = 'row',
  sheetOrder: readonly string[] = [...new Set(matches.map(match => match.sheetId))],
): number {
  if (!matches.length) return -1
  const current = matches.findIndex(match => match.sheetId === sheetId && match.key === key)
  if (current < 0) {
    const point = parseKey(key)
    if (point) {
      const localMatches = matches.map((match, index) => ({ match, index })).filter(({ match }) => {
        if (match.sheetId !== sheetId) return false
        const address = parseKey(match.key)!
        const order = findDirection === 'row'
          ? address.row - point.row || address.column - point.column
          : address.column - point.column || address.row - point.row
        return order * direction > 0
      })
      if (localMatches.length) return direction === 1 ? localMatches[0]!.index : localMatches.at(-1)!.index
    }

    const currentSheetOrder = sheetOrder.indexOf(sheetId)
    const nextSheet = direction === 1
      ? sheetOrder.slice(currentSheetOrder + 1).find(id => matches.some(match => match.sheetId === id))
      : sheetOrder.slice(0, currentSheetOrder).reverse().find(id => matches.some(match => match.sheetId === id))
    if (nextSheet) {
      return direction === 1
        ? matches.findIndex(match => match.sheetId === nextSheet)
        : matches.map(match => match.sheetId).lastIndexOf(nextSheet)
    }
    return direction === 1 ? 0 : matches.length - 1
  }
  return (current + direction + matches.length) % matches.length
}

export function createSheetReplacePatches(
  sheets: readonly ExcelSheetView[],
  matches: readonly SheetFindMatch[],
  query: SheetFindQuery,
  replaceAll: boolean,
  selectedIndex: number,
): SheetFindReplacePatch[] {
  const targets = replaceAll
    ? matches.filter(match => match.replaceable)
    : selectedIndex >= 0 && matches[selectedIndex]?.replaceable ? [matches[selectedIndex]!] : []
  const patches: SheetFindReplacePatch[] = []
  for (const match of targets) {
    const cell = sheets.find(sheet => sheet.id === match.sheetId)?.cells[match.key]
    if (!cell) continue
    if (query.findBy === 'formula') {
      patches.push({ ...match, cell: { ...cell, formula: replaceText(cell.formula ?? '', query), value: '', cachedValue: undefined, rawValue: undefined, displayValue: undefined } })
      continue
    }
    const source = rawCellValue(cell)
    const replaced = replaceText(source, query)
    patches.push({ ...match, cell: { ...cell, value: replaced, rawValue: replaced, displayValue: undefined } })
  }
  return patches
}

function replaceText(value: string, query: SheetFindQuery): string {
  if (query.wholeCell) return query.replacement
  const source = query.caseSensitive ? value : value.toLocaleLowerCase()
  const needle = query.caseSensitive ? query.text : query.text.toLocaleLowerCase()
  const pieces: string[] = []
  let cursor = 0
  while (cursor < value.length) {
    const index = source.indexOf(needle, cursor)
    if (index < 0) break
    pieces.push(value.slice(cursor, index), query.replacement)
    cursor = index + query.text.length
  }
  pieces.push(value.slice(cursor))
  return pieces.join('')
}

function rawCellValue(cell: ExcelCellView): string {
  if (cell.rawValue !== undefined && cell.rawValue !== null) return String(cell.rawValue)
  return cell.value
}

function matchesText(value: string, text: string, caseSensitive: boolean, wholeCell: boolean): boolean {
  const candidate = caseSensitive ? value : value.toLocaleLowerCase()
  return wholeCell ? candidate === text : candidate.includes(text)
}

function parseKey(key: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/i.exec(key)
  if (!match) return undefined
  let column = 0
  for (const letter of match[1]!.toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, column: column - 1 }
}

function normalizeRanges(ranges: readonly { startRow: number; endRow: number; startColumn: number; endColumn: number }[]) {
  return [...ranges]
}

function inRanges(address: { row: number; column: number }, ranges: ReturnType<typeof normalizeRanges>): boolean {
  return !ranges.length || ranges.some(range => address.row >= range.startRow && address.row <= range.endRow
    && address.column >= range.startColumn && address.column <= range.endColumn)
}
