export interface SheetRowVisibility {
  rowCount: number
  hiddenRows?: readonly number[]
  filteredRows?: readonly number[]
  rowHeights?: Readonly<Record<number, number>>
}

/** Keep manual and filter reasons distinct in the model; combine them only for viewing. */
export function hiddenSheetRows(sheet: SheetRowVisibility): Set<number> {
  const hidden = new Set([...(sheet.hiddenRows ?? []), ...(sheet.filteredRows ?? [])])
  for (const [key, height] of Object.entries(sheet.rowHeights ?? {})) if (height === 0) hidden.add(Number(key))
  return new Set([...hidden].filter(row => Number.isInteger(row) && row >= 0 && row < sheet.rowCount))
}

export function findVisibleSheetRow(sheet: SheetRowVisibility, start: number, direction: -1 | 1, fallback = true): number | undefined {
  const hidden = hiddenSheetRows(sheet)
  const first = Math.max(0, Math.min(sheet.rowCount - 1, start))
  for (let row = first; row >= 0 && row < sheet.rowCount; row += direction) if (!hidden.has(row)) return row
  if (fallback) for (let row = first - direction; row >= 0 && row < sheet.rowCount; row -= direction) if (!hidden.has(row)) return row
  return undefined
}
