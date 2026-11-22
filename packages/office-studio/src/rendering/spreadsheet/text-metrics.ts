import { sheetTextGraphemes, type SheetTextMeasure } from '../../features/sheet-workbench/sheet-cell-text-layout.ts'

let context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null | undefined

/** Canvas2D text metrics also feed the WebGPU scene; the fallback is used before a canvas is available. */
export const measureSpreadsheetText: SheetTextMeasure = (text, font) => {
  if (context === undefined) context = typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(1, 1).getContext('2d')
    : typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null
  if (context) {
    context.font = font
    const width = context.measureText(text).width
    if (Number.isFinite(width)) return Math.max(0, width)
  }
  const size = Number(/([\d.]+)px/u.exec(font)?.[1] ?? 14.667)
  return sheetTextGraphemes(text).reduce((width, glyph) => width + size * (/^\s+$/u.test(glyph.text) ? 0.3
    : /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Extended_Pictographic}]/u.test(glyph.text) ? 1 : 0.55), 0)
}
