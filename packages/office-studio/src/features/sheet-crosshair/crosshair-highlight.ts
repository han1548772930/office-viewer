import type { GridRange } from '../../rendering/spreadsheet/layout'

function intersect(left: GridRange, right: GridRange): GridRange | undefined {
  const range = {
    startRow: Math.max(left.startRow, right.startRow),
    endRow: Math.min(left.endRow, right.endRow),
    startColumn: Math.max(left.startColumn, right.startColumn),
    endColumn: Math.min(left.endColumn, right.endColumn),
  }
  return range.startRow <= range.endRow && range.startColumn <= range.endColumn ? range : undefined
}

function subtract(range: GridRange, cut: GridRange): GridRange[] {
  const overlap = intersect(range, cut)
  if (!overlap) return [range]

  return [
    { ...range, endRow: overlap.startRow - 1 },
    { ...range, startRow: overlap.endRow + 1 },
    { startRow: overlap.startRow, endRow: overlap.endRow, startColumn: range.startColumn, endColumn: overlap.startColumn - 1 },
    { startRow: overlap.startRow, endRow: overlap.endRow, startColumn: overlap.endColumn + 1, endColumn: range.endColumn },
  ].filter(item => item.startRow <= item.endRow && item.startColumn <= item.endColumn)
}

function addUncovered(ranges: GridRange[], candidate: GridRange, exclusions: readonly GridRange[]): void {
  let pieces = [candidate]
  for (const exclusion of [...ranges, ...exclusions]) {
    pieces = pieces.flatMap(piece => subtract(piece, exclusion))
    if (!pieces.length) return
  }
  ranges.push(...pieces)
}

/** Creates the unselected row/column cross around ordinary cell selections. */
export function crosshairHighlightRanges(
  selections: readonly GridRange[],
  rowCount: number,
  columnCount: number,
): GridRange[] {
  if (rowCount <= 0 || columnCount <= 0) return []

  const selected = selections.filter(range =>
    range.startRow > 0 || range.endRow < rowCount - 1,
  ).filter(range =>
    range.startColumn > 0 || range.endColumn < columnCount - 1,
  )
  const ranges: GridRange[] = []

  for (const selection of selected) {
    const candidates: GridRange[] = [
      { startRow: selection.startRow, endRow: selection.endRow, startColumn: 0, endColumn: selection.startColumn - 1 },
      { startRow: selection.startRow, endRow: selection.endRow, startColumn: selection.endColumn + 1, endColumn: columnCount - 1 },
      { startRow: 0, endRow: selection.startRow - 1, startColumn: selection.startColumn, endColumn: selection.endColumn },
      { startRow: selection.endRow + 1, endRow: rowCount - 1, startColumn: selection.startColumn, endColumn: selection.endColumn },
    ]
    for (const candidate of candidates) {
      if (candidate.startRow <= candidate.endRow && candidate.startColumn <= candidate.endColumn) {
        addUncovered(ranges, candidate, selected)
      }
    }
  }

  return ranges
}
