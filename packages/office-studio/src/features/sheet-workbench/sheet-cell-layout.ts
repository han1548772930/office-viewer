import type { CellStyle } from '../../domain/model.ts'

export type SheetHorizontalAlignment = CellStyle['horizontal']
export type SheetVerticalAlignment = CellStyle['vertical']
export type SheetWrapStrategy = NonNullable<CellStyle['wrapStrategy']>
export type SheetTextRotation = number | 'v'
export type SheetCellLayoutStyle = Partial<Pick<CellStyle, 'horizontal' | 'vertical' | 'wrap' | 'wrapStrategy' | 'shrinkToFit' | 'rotation' | 'rotationVertical'>>

/** Explicit strategies are authoritative; older snapshots may contain only the wrap boolean. */
export function sheetCellWrapStrategy(style: Pick<SheetCellLayoutStyle, 'wrap' | 'wrapStrategy'> | undefined): SheetWrapStrategy {
  return style?.wrapStrategy ?? (style?.wrap ? 'wrap' : 'overflow')
}
export function sheetCellTextRotation(style: Pick<SheetCellLayoutStyle, 'rotation' | 'rotationVertical'> | undefined): SheetTextRotation {
  return style?.rotationVertical || style?.rotation === 255 ? 'v' : style?.rotation ?? 0
}
export function sheetCellShrinkToFit(style: Pick<SheetCellLayoutStyle, 'wrap' | 'wrapStrategy' | 'shrinkToFit'> | undefined): boolean {
  return style?.shrinkToFit === true && sheetCellWrapStrategy(style) !== 'wrap'
}
