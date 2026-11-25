import type { RenderRect } from '../contracts.ts'

export interface SheetRichTextEditorScroll { left: number; top: number }
export interface SheetRichTextEditorScrollExtent { clip: RenderRect; contentWidth: number; contentHeight: number }

const nonnegative = (value: number) => Number.isFinite(value) ? Math.max(0, value) : 0

/** Rich glyph extents, rather than the uniform-font native input, bound the view scroll. */
export function clampSheetRichTextEditorScroll(extent: SheetRichTextEditorScrollExtent, scroll: SheetRichTextEditorScroll): SheetRichTextEditorScroll {
  return {
    left: Math.min(Math.max(0, nonnegative(extent.contentWidth) - nonnegative(extent.clip.width)), nonnegative(scroll.left)),
    top: Math.min(Math.max(0, nonnegative(extent.contentHeight) - nonnegative(extent.clip.height)), nonnegative(scroll.top)),
  }
}

/** A stationary captured pointer outside the visible editor keeps extending its selection. */
export function stepSheetRichTextEditorDragScroll(extent: SheetRichTextEditorScrollExtent, scroll: SheetRichTextEditorScroll,
  pointer: { x: number; y: number }, elapsedMs: number): SheetRichTextEditorScroll {
  const current = clampSheetRichTextEditorScroll(extent, scroll), clip = extent.clip
  if (!Number.isFinite(pointer.x) || !Number.isFinite(pointer.y) || !Number.isFinite(elapsedMs) || elapsedMs <= 0
    || !Number.isFinite(clip.x) || !Number.isFinite(clip.y) || !Number.isFinite(clip.width) || !Number.isFinite(clip.height)
    || clip.width <= 0 || clip.height <= 0) return current
  const speed = (value: number, start: number, size: number) => {
    const distance = value < start ? value - start : value > start + size ? value - start - size : 0
    return distance === 0 ? 0 : Math.sign(distance) * Math.min(1200, 120 + Math.abs(distance) * 8)
  }
  // A suspended tab or late frame must not apply seconds of scroll in one update.
  const seconds = Math.min(elapsedMs, 64) / 1000
  return clampSheetRichTextEditorScroll(extent, {
    left: current.left + speed(pointer.x, clip.x, clip.width) * seconds,
    top: current.top + speed(pointer.y, clip.y, clip.height) * seconds,
  })
}

