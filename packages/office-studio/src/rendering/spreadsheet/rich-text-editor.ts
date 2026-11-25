import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import type { DrawCommand, RenderRect } from '../contracts.ts'
import { layoutSheetCellRichText } from '../../features/sheet-workbench/sheet-cell-rich-text-layout.ts'
import { sheetTextGraphemes, type SheetTextMeasure } from '../../features/sheet-workbench/sheet-cell-text-layout.ts'
import { sheetCellTextAlignment } from '../../features/sheet-workbench/sheet-cell-alignment.ts'
import { sheetCellWrapStrategy } from '../../features/sheet-workbench/sheet-cell-layout.ts'
import { createSheetCellTextCommands } from './cell-text.ts'
import { measureSpreadsheetText } from './text-metrics.ts'

export interface SheetRichTextEditorLine {
  start: number
  end: number
  top: number
  height: number
  offsets: number[]
  /** Prefix shaping is measured lazily, rather than once for every character. */
  xAt: (offset: number) => number
}
export interface SheetRichTextEditorScene {
  commands: DrawCommand[]
  lines: SheetRichTextEditorLine[]
  clip: RenderRect
  contentWidth: number
  contentHeight: number
  caretAffinity?: SheetRichTextEditorCaretAffinity
}
export type SheetRichTextEditorCaretAffinity = 'upstream' | 'downstream'

/** The antdv input supplies text/IME events; Canvas supplies rich glyphs, selection and pointer positions. */
export function createSheetRichTextEditorScene(cell: ExcelCellView, rect: RenderRect, clip: RenderRect, zoom: number,
  selection: { start: number; end: number }, options: { background: string; color: string; accent: string; selectionColor: string;
    scrollLeft?: number; scrollTop?: number; composing?: boolean; focused?: boolean;
    caretAffinity?: SheetRichTextEditorCaretAffinity; measure?: SheetTextMeasure }): SheetRichTextEditorScene | undefined {
  const rawMeasure = options.measure ?? measureSpreadsheetText
  const measure: SheetTextMeasure = (text, font) => {
    const width = rawMeasure(text, font)
    return Number.isFinite(width) ? Math.max(0, width) : 0
  }
  const view: ExcelCellView = { ...cell, color: cell.color ?? options.color,
    align: cell.align === 'justify' ? 'left' : cell.align, rotation: 0, rotationVertical: false,
    shrinkToFit: false, indent: 0, vertical: 'top' }
  const layout = layoutSheetCellRichText(view, cell.value, { width: Math.max(0, rect.width - 12 * zoom), height: Infinity,
    wrap: sheetCellWrapStrategy(view) === 'wrap', zoom, measure })
  if (!layout) return undefined
  if (layout.width > rect.width - 12 * zoom && (view.align === 'right' || view.align === 'center')) view.align = 'left'
  const align = sheetCellTextAlignment(view), left = rect.x + 6 * zoom - (options.scrollLeft ?? 0), top = rect.y + 3 * zoom - (options.scrollTop ?? 0)
  const lines = layout.lines.map(line => {
    const offset = align === 'center' ? (rect.width - 12 * zoom - line.width) / 2
      : align === 'right' ? rect.width - 12 * zoom - line.width : 0
    const offsets = [line.start]
    const fragments: Array<{ start: number; end: number; text: string; font: string; left: number; width: number }> = []
    let x = left + offset
    for (const fragment of line.fragments) {
      fragments.push({ ...fragment, left: x })
      for (const glyph of sheetTextGraphemes(fragment.text)) {
        const end = glyph.offset + glyph.text.length
        offsets.push(fragment.start + end)
      }
      x += fragment.width
    }
    const cache = new Map<number, number>([[line.start, left + offset], [line.end, x]])
    const xAt = (rawOffset: number) => {
      const offset = Math.max(line.start, Math.min(line.end, rawOffset))
      const known = cache.get(offset)
      if (known !== undefined) return known
      let first = 0, last = fragments.length
      while (first < last) {
        const middle = (first + last) >>> 1
        if (fragments[middle]!.end < offset) first = middle + 1
        else last = middle
      }
      const fragment = fragments[first]
      const position = fragment ? fragment.left + measure(fragment.text.slice(0, offset - fragment.start), fragment.font) : x
      if (cache.size >= 128) cache.clear()
      cache.set(offset, position)
      return position
    }
    return { start: line.start, end: line.end, top: top + line.top, height: line.lineHeight, offsets, xAt }
  })
  const commands: DrawCommand[] = [{ kind: 'rect', ...rect, fill: options.background, clip }]
  for (const line of options.composing ? [] : lines) {
    const start = Math.max(selection.start, line.start), end = Math.min(selection.end, line.end)
    if (start < end || selection.start <= line.end && selection.end > line.end) {
      const x1 = line.xAt(start), x2 = line.xAt(end)
      commands.push({ kind: 'rect', x: x1, y: line.top, width: Math.max(2 * zoom, x2 - x1), height: line.height, fill: options.selectionColor, clip })
    }
  }
  const paintRect = { ...rect, x: rect.x - (options.scrollLeft ?? 0), y: rect.y - (options.scrollTop ?? 0) }
  if (!options.composing) commands.push(...createSheetCellTextCommands({ cell: view, presentation: { text: cell.value }, rect: paintRect, clip, zoom, measure }))
  const scene: SheetRichTextEditorScene = { commands, lines, clip,
    contentWidth: layout.width + 12 * zoom, contentHeight: layout.height + 6 * zoom, caretAffinity: options.caretAffinity }
  if (selection.start === selection.end && !options.composing && options.focused !== false) {
    const caret = sheetRichTextEditorCaret(scene, selection.start)
    if (caret) {
      commands.push({ kind: 'line', x1: caret.x, y1: caret.top, x2: caret.x, y2: caret.top + caret.height, color: options.accent, width: 1, clip })
    }
  }
  return scene
}

export function hitSheetRichTextEditor(scene: SheetRichTextEditorScene, x: number, y: number): number | undefined {
  return hitSheetRichTextEditorPosition(scene, x, y)?.offset
}

/** Wrapped line ends and starts can share one UTF-16 offset while occupying different visual rows. */
export function hitSheetRichTextEditorPosition(scene: SheetRichTextEditorScene, x: number, y: number):
  { offset: number; affinity: SheetRichTextEditorCaretAffinity } | undefined {
  if (!Number.isFinite(x) || !Number.isFinite(y) || !scene.lines.length) return undefined
  const line = scene.lines.reduce((best, line) => {
    const distance = (item: SheetRichTextEditorLine) => y < item.top ? item.top - y : y > item.top + item.height ? y - item.top - item.height : 0
    return distance(line) < distance(best) ? line : best
  })
  let first = 0, last = line.offsets.length - 1
  while (first < last) {
    const middle = (first + last) >>> 1
    const a = line.offsets[middle]!, b = line.offsets[middle + 1]!
    if (x < (line.xAt(a) + line.xAt(b)) / 2) last = middle
    else first = middle + 1
  }
  const offset = line.offsets[first]
  return offset === undefined ? undefined : { offset, affinity: offset === line.end ? 'upstream' : 'downstream' }
}

export function sheetRichTextEditorCaret(scene: SheetRichTextEditorScene, offset: number,
  affinity: SheetRichTextEditorCaretAffinity = scene.caretAffinity ?? 'downstream') {
  let index = -1
  for (let at = affinity === 'upstream' ? 0 : scene.lines.length - 1;
    at >= 0 && at < scene.lines.length; at += affinity === 'upstream' ? 1 : -1) {
    const line = scene.lines[at]!
    if (line.start <= offset && line.end >= offset) { index = at; break }
  }
  const lineIndex = index < 0 ? scene.lines.length - 1 : index, line = scene.lines[lineIndex]
  return line ? { x: line.xAt(offset), top: line.top, height: line.height, lineIndex } : undefined
}

/** Visual lines use the same run metrics as paint; UTF-16 offsets still belong to the antdv input. */
export function moveSheetRichTextEditorCaret(scene: SheetRichTextEditorScene, offset: number,
  key: 'ArrowUp' | 'ArrowDown' | 'Home' | 'End', preferredX?: number):
  { offset: number; preferredX?: number; affinity: SheetRichTextEditorCaretAffinity } | undefined {
  const caret = sheetRichTextEditorCaret(scene, offset)
  if (!caret) return undefined
  const current = scene.lines[caret.lineIndex]!
  if (key === 'Home' || key === 'End') return { offset: key === 'Home' ? current.start : current.end,
    affinity: key === 'Home' ? 'downstream' : 'upstream' }
  const index = Math.max(0, Math.min(scene.lines.length - 1, caret.lineIndex + (key === 'ArrowUp' ? -1 : 1)))
  if (index === caret.lineIndex) return { offset, preferredX: preferredX ?? caret.x, affinity: scene.caretAffinity ?? 'downstream' }
  const next = scene.lines[index]!, x = preferredX ?? caret.x
  const result = hitSheetRichTextEditorPosition({ ...scene, lines: [next] }, x, next.top + next.height / 2)
  return result === undefined ? undefined : { ...result, preferredX: x }
}

export function sheetRichTextEditorWord(text: string, offset: number): { start: number; end: number } {
  if (typeof Intl.Segmenter === 'function') {
    for (const segment of new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text)) {
      if (segment.index <= offset && segment.index + segment.segment.length > offset) return { start: segment.index, end: segment.index + segment.segment.length }
    }
  }
  const glyph = sheetTextGraphemes(text).find(glyph => glyph.offset <= offset && glyph.offset + glyph.text.length > offset)
  return { start: glyph?.offset ?? text.length, end: glyph ? glyph.offset + glyph.text.length : text.length }
}
