import type { SlicerModel } from '../../domain/model.ts'

/** Local CSS geometry shared by member-window scheduling and Canvas/Ant item placement. */
export function slicerMemberViewport(model: SlicerModel, width: number, height: number, total?: number, startItem = model.startItem ?? 0) {
  const headerHeight = model.showCaption === false ? 0 : Math.min(32, height), rowHeight = model.rowHeight ?? 24,
    columns = Math.max(1, model.columnCount ?? 1), innerHeight = Math.max(0, height - headerHeight - 8)
  let rows = Math.max(1, Math.ceil(innerHeight / (rowHeight + 4)))
  while (rows > 1 && (innerHeight - rowHeight) / (rows - 1) - rowHeight < 0) rows -= 1
  const pitch = rows > 1 ? (innerHeight - rowHeight) / (rows - 1) : rowHeight,
    totalRows = total === undefined ? undefined : Math.ceil(total / columns),
    maxScrollRows = totalRows === undefined ? undefined : Math.max(0, totalRows - rows),
    scrollRow = Math.max(0, Math.min(maxScrollRows ?? Number.POSITIVE_INFINITY, Math.floor(startItem / columns))),
    scrollWidth = totalRows !== undefined && totalRows > rows ? 14 : 0,
    itemWidth = Math.max(0, (width - 8 - scrollWidth - 4 * (columns - 1)) / columns)
  return { start: scrollRow * columns, count: rows * columns, columns, rows, pitch, totalRows: totalRows ?? 0,
    maxScrollOffset: (maxScrollRows ?? 0) * columns, headerHeight, scrollWidth, itemWidth }
}
