import type { CellRange } from '../../domain/model.ts'

export interface SparklineAddress {
  row: number
  column: number
  absoluteRow: boolean
  absoluteColumn: boolean
}
export interface SparklineReference {
  range: CellRange
  sheetName?: string
  first: SparklineAddress
  last: SparklineAddress
  normalized: string
}

export function sparklineCellAddress(row: number, column: number): string {
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || row > 1_048_575 || column < 0 || column > 16_383) {
    throw new RangeError('迷你图区域超出工作表边界。')
  }
  let label = ''
  for (let index = column; index >= 0; index = Math.floor(index / 26) - 1) label = String.fromCharCode(index % 26 + 65) + label
  return `${label}${row + 1}`
}

function parseAddress(value: string): SparklineAddress | undefined {
  const match = /^(\$?)([A-Z]+)(\$?)([1-9]\d*)$/i.exec(value.trim())
  if (!match) return undefined
  let column = 0
  for (const character of match[2]!.toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64
  const row = Number(match[4]) - 1
  column -= 1
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row > 1_048_575 || column > 16_383) return undefined
  return { row, column, absoluteRow: match[3] === '$', absoluteColumn: match[1] === '$' }
}

function formatAddress(address: SparklineAddress): string {
  const plain = sparklineCellAddress(address.row, address.column)
  const match = /^([A-Z]+)(\d+)$/.exec(plain)!
  return `${address.absoluteColumn ? '$' : ''}${match[1]}${address.absoluteRow ? '$' : ''}${match[2]}`
}

export function sparklineSheetQualifier(name: string): string {
  return /^[\p{L}\p{N}_]+$/u.test(name) ? `${name}!` : `'${name.replaceAll("'", "''")}'!`
}

/** A1 references retain absolute coordinates and sheet spelling; external workbooks are not silently accepted. */
export function parseSparklineReference(input: string): SparklineReference | undefined {
  if (typeof input !== 'string') return undefined
  const text = input.trim().replace(/^=/, '')
  let quoted = false, qualifierEnd = -1
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === "'") {
      if (quoted && text[index + 1] === "'") { index += 1; continue }
      quoted = !quoted
    } else if (character === '!' && !quoted) {
      if (qualifierEnd !== -1) return undefined
      qualifierEnd = index
    }
  }
  if (quoted) return undefined
  let sheetName: string | undefined
  if (qualifierEnd !== -1) {
    const qualifier = text.slice(0, qualifierEnd)
    if (qualifier.startsWith("'") && qualifier.endsWith("'")) {
      const name = qualifier.slice(1, -1)
      if (name.replaceAll("''", '').includes("'")) return undefined
      sheetName = name.replaceAll("''", "'")
    }
    else if (!qualifier.includes("'") && qualifier.trim() === qualifier) sheetName = qualifier
    else return undefined
    if (!sheetName || /[\[\]:\\/?*]/.test(sheetName)) return undefined
  }
  const addresses = text.slice(qualifierEnd + 1).split(':')
  if (!addresses.length || addresses.length > 2) return undefined
  const first = parseAddress(addresses[0] ?? ''), last = parseAddress(addresses[1] ?? addresses[0] ?? '')
  if (!first || !last) return undefined
  const range: CellRange = {
    startRow: Math.min(first.row, last.row), endRow: Math.max(first.row, last.row),
    startCol: Math.min(first.column, last.column), endCol: Math.max(first.column, last.column),
  }
  return { range, ...(sheetName === undefined ? {} : { sheetName }), first, last,
    normalized: `${sheetName === undefined ? '' : sparklineSheetQualifier(sheetName)}${formatAddress(first)}${addresses.length === 2 ? `:${formatAddress(last)}` : ''}` }
}

/** Commas, semicolons and newlines separate areas, except inside quoted worksheet names. */
export function parseSparklineRangeList(input: string): string[] {
  const result: string[] = []
  let quoted = false, start = 0
  for (let index = 0; index < input.length; index += 1) {
    const character = input[index]
    if (character === "'") {
      if (quoted && input[index + 1] === "'") { index += 1; continue }
      quoted = !quoted
    } else if (!quoted && (character === ',' || character === ';' || character === '\n' || character === '\r')) {
      const area = input.slice(start, index).trim()
      if (area) result.push(area)
      start = index + 1
    }
  }
  if (quoted) throw new Error('迷你图工作表名称的引号未闭合。')
  const last = input.slice(start).trim()
  if (last) result.push(last)
  if (!result.length) throw new Error('请填写迷你图数据或放置区域。')
  return result.map(value => {
    const reference = parseSparklineReference(value)
    if (!reference) throw new Error('迷你图区域必须是当前工作簿中的有效 A1 范围。')
    return reference.normalized
  })
}

export function sparklineRangeString(range: CellRange, sheetName?: string): string {
  const first = sparklineCellAddress(range.startRow, range.startCol), last = sparklineCellAddress(range.endRow, range.endCol)
  return `${sheetName === undefined ? '' : sparklineSheetQualifier(sheetName)}${first}${first === last ? '' : `:${last}`}`
}

export function shiftSparklineReference(input: string, rowDelta: number, columnDelta: number): string {
  const reference = parseSparklineReference(input)
  if (!reference || !Number.isSafeInteger(rowDelta) || !Number.isSafeInteger(columnDelta)) throw new Error('迷你图源区域或移动距离无效。')
  const shift = (address: SparklineAddress): SparklineAddress => {
    const next = { ...address, row: address.row + (address.absoluteRow ? 0 : rowDelta), column: address.column + (address.absoluteColumn ? 0 : columnDelta) }
    sparklineCellAddress(next.row, next.column)
    return next
  }
  const first = shift(reference.first), last = shift(reference.last)
  return `${reference.sheetName === undefined ? '' : sparklineSheetQualifier(reference.sheetName)}${formatAddress(first)}${input.includes(':') ? `:${formatAddress(last)}` : ''}`
}
