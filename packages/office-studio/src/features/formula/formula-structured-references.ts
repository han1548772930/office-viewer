import type { CellRange, SheetTableModel, WorkbookModel } from '../../domain/model'
import type { FormulaCellAddress, FormulaStructuredReferenceNode } from './formula-types'

export function readFormulaStructuredSpecifier(source: string, start: number): { specifier: string; end: number } | undefined {
  if (source[start] !== '[') return undefined
  let depth = 0, index = start
  while (index < source.length) {
    const character = source[index++]!
    if (character === "'" && /[\[\]#@']/.test(source[index] ?? '')) { index++; continue }
    if (character === '[') depth++
    if (character === ']' && --depth === 0) return { specifier: source.slice(start, index), end: index }
  }
  return undefined
}
const unescapeColumn = (name: string): string => name.replace(/'([\[\]#@'])/g, '$1')

/** Table metadata owns column identity, while its current range and flags own
 * the header/data/total footprint. Hidden and filtered rows remain references. */
export function resolveFormulaStructuredRange(workbook: WorkbookModel, table: SheetTableModel,
  specifier: string, cell?: Pick<FormulaCellAddress, 'row' | 'column'>): CellRange | undefined {
  if (!workbook.sheets.some(sheet => sheet.id === table.sheetId)) return undefined
  const parts: string[] = []
  let current = '', nesting = 0, thisRow = false, interval = false
  const compound = specifier.startsWith('[[') || specifier.startsWith('[@[')
  const flush = (): void => { if (current.length) parts.push(current); current = '' }
  for (let index = 0; index < specifier.length; index++) {
    const character = specifier[index]!
    if (character === "'" && /[\[\]#@']/.test(specifier[index + 1] ?? '')) { current += character + specifier[++index]!; continue }
    if (character === '[') { nesting++; continue }
    if (character === ']') { nesting--; flush(); continue }
    if (compound && nesting <= 1 && (character === ',' || character === ':')) { flush(); interval ||= character === ':'; continue }
    if (compound && nesting <= 1 && /\s/.test(character)) continue
    if (character === '@' && !current.length) { thisRow = true; continue }
    current += character
  }
  flush()
  const selectors = new Set<string>(), columns: string[] = []
  for (const part of parts) {
    const normalized = part.trim().toLowerCase()
    if (part[0] === '#') selectors.add(normalized)
    else columns.push(unescapeColumn(part))
  }
  const legal = new Set(['#all', '#data', '#headers', '#totals', '#this row'])
  if ([...selectors].some(selector => !legal.has(selector))) return undefined
  thisRow ||= selectors.has('#this row')
  const header = table.range.startRow, total = table.range.endRow
  const startData = header + Number(table.showHeaderRow), endData = total - Number(table.showTotalsRow)
  let startRow = startData, endRow = endData
  if (thisRow) {
    if (!cell || cell.row < startData || cell.row > endData) return undefined
    startRow = endRow = cell.row
  } else if (selectors.has('#all')) { startRow = header; endRow = total }
  else if (selectors.size) {
    if (selectors.has('#headers') && !table.showHeaderRow || selectors.has('#totals') && !table.showTotalsRow) return undefined
    if (selectors.has('#headers') && selectors.has('#totals') && !selectors.has('#data') && endData >= startData) return undefined
    startRow = selectors.has('#headers') ? header : selectors.has('#data') ? startData : total
    endRow = selectors.has('#totals') ? total : selectors.has('#data') ? endData : header
  }
  if (endRow < startRow || columns.length > 2 || columns.length === 2 && !interval) return undefined
  const indices = columns.map(name => table.columns.findIndex(column => column.name.toLowerCase() === name.toLowerCase()))
  if (indices.some(index => index < 0)) return undefined
  return { startRow, endRow, startCol: table.range.startCol + (indices.length ? Math.min(...indices) : 0),
    endCol: table.range.startCol + (indices.length ? Math.max(...indices) : table.columns.length - 1) }
}
export function findFormulaStructuredTable(workbook: WorkbookModel, node: FormulaStructuredReferenceNode,
  sheetId: string, cell?: Pick<FormulaCellAddress, 'row' | 'column'>): SheetTableModel | undefined {
  return node.tableName ? workbook.tables?.find(table => table.name.toLowerCase() === node.tableName!.toLowerCase())
    : cell ? workbook.tables?.find(table => table.sheetId === sheetId && cell.row >= table.range.startRow && cell.row <= table.range.endRow
      && cell.column >= table.range.startCol && cell.column <= table.range.endCol) : undefined
}
