import type { DocumentModel } from '../../../domain/model'
import type { DrawCommand, RenderRect, RenderScene } from '../../../rendering/contracts'
import type { DocumentSelection } from '../model/document-model'
import { orderedSelection, sameTarget, selectionCollapsed, targetKey, textTargets } from '../model/document-model'
import { caretRect, type DocumentLayout } from '../layout/document-layout'

export interface DocumentViewport { width: number; height: number; scrollLeft: number; scrollTop: number; zoom: number }
export interface DocumentSceneOptions { selection: DocumentSelection; focused: boolean; caretVisible: boolean; composing: boolean; images: Map<string, HTMLImageElement> }

export function sceneFromDocument(model: DocumentModel, layout: DocumentLayout, viewport: DocumentViewport, options: DocumentSceneOptions): RenderScene {
  const commands: DrawCommand[] = []
  const scale = viewport.zoom / 100
  const { x: pageOffsetX, y: pageOffsetY } = sceneOffset(layout, viewport)
  const x = (value: number) => pageOffsetX + value * scale
  const y = (value: number) => pageOffsetY + value * scale
  const rect = (source: RenderRect): RenderRect => ({ x: x(source.x), y: y(source.y), width: source.width * scale, height: source.height * scale })
  const visible = (source: RenderRect) => y(source.y + source.height) >= -20 && y(source.y) <= viewport.height + 20 && x(source.x + source.width) >= -20 && x(source.x) <= viewport.width + 20
  // Layout objects are ordered by page. Keep a page index set so a long
  // document does not run every cell/image/divider through viewport math on
  // every animation frame. Objects on a visible page still use the precise
  // rectangle test below because a page can be taller than the viewport.
  const visiblePages = new Set<number>()
  for (const page of layout.pages) {
    if (!visible(page)) continue
    visiblePages.add(page.index)
    commands.push({ kind: 'rect', ...rect({ ...page, x: page.x + 3, y: page.y + 4 }), fill: '#dce1e8' })
    commands.push({ kind: 'rect', ...rect(page), fill: '#ffffff', stroke: '#d9dfe7', strokeWidth: 1 })
    commands.push({ kind: 'text', x: x(page.x + page.width / 2), y: y(page.y + page.height - 24), text: String(page.index + 1), font: `${10 * scale}px Arial`, color: '#9ca3af', align: 'center', baseline: 'middle' })
  }
  for (const pageIndex of visiblePages) {
    const content = layout.pageContent[pageIndex]
    if (!content) continue
    for (const cell of content.cells) {
      if (visible(cell)) commands.push({ kind: 'rect', ...rect(cell), fill: cell.fill, stroke: '#aeb8c6', strokeWidth: Math.max(0.5, scale) })
    }
    for (const divider of content.dividers) {
      if (visible(divider)) commands.push({ kind: 'line', x1: x(divider.x), y1: y(divider.y), x2: x(divider.x + divider.width), y2: y(divider.y), color: '#b3bdca', width: Math.max(1, scale) })
    }
    for (const image of content.images) {
      if (!visible(image)) continue
    const source = options.images.get(image.src)
    if (source?.complete && source.naturalWidth) commands.push({ kind: 'image', ...rect(image), source })
    else {
      commands.push({ kind: 'rect', ...rect(image), fill: '#f1f4f8', stroke: '#ccd5e0' })
      commands.push({ kind: 'text', x: x(image.x + 12), y: y(image.y + 22), text: image.alt || '图片', font: `${12 * scale}px Arial`, color: '#738096' })
    }
    }
  }
  const ordered = orderedSelection(model, options.selection)
  const targets = textTargets(model)
  const first = targets.findIndex((target) => sameTarget(target, ordered.start))
  const last = targets.findIndex((target) => sameTarget(target, ordered.end))
  const selected = new Set(targets.slice(first, last + 1).map(targetKey))
  const visibleLines = [...visiblePages].flatMap((pageIndex) => layout.pageContent[pageIndex]?.lines ?? [])
  for (const line of visibleLines) {
    if (!visible(line)) continue
    const clip = rect(layout.pages[line.page]!.content)
    const hasSelection = !selectionCollapsed(options.selection) && selected.has(targetKey(line.target))
    const start = sameTarget(line.target, ordered.start) ? ordered.start.offset : 0
    const end = sameTarget(line.target, ordered.end) ? ordered.end.offset : Number.POSITIVE_INFINITY
    if (line.marker) commands.push({ kind: 'text', x: x(line.x - 17), y: y(line.baseline), text: line.marker, font: `${14 * scale}px Arial`, color: '#445166', clip })
    if (line.checklist) {
      const size = 14 * scale
      const left = x(line.x - 23)
      const top = y(line.y + (line.height - 14) / 2)
      commands.push({ kind: 'rect', x: left, y: top, width: size, height: size, fill: '#ffffff', stroke: line.checklist.checked ? '#28744b' : '#718096', strokeWidth: Math.max(1, scale), clip })
      if (line.checklist.checked) {
        commands.push({ kind: 'line', x1: left + 3 * scale, y1: top + 7 * scale, x2: left + 6 * scale, y2: top + 10 * scale, color: '#28744b', width: Math.max(1.4, scale), clip })
        commands.push({ kind: 'line', x1: left + 6 * scale, y1: top + 10 * scale, x2: left + 12 * scale, y2: top + 3 * scale, color: '#28744b', width: Math.max(1.4, scale), clip })
      }
    }
    for (const glyph of line.glyphs) {
      const baselineOffset = glyph.run.verticalAlign === 'superscript' ? -glyph.run.fontSize * 0.4 : glyph.run.verticalAlign === 'subscript' ? glyph.run.fontSize * 0.2 : 0
      const baseline = line.baseline + baselineOffset
      if (glyph.run.highlight) commands.push({ kind: 'rect', ...rect(glyph), fill: glyph.run.highlight, clip })
      if (hasSelection && glyph.end > start && glyph.offset < end) commands.push({ kind: 'rect', ...rect(glyph), fill: options.focused ? '#bdd8ff' : '#dde6f2', clip })
      const font = glyph.font.replace(/([\d.]+)px/, (_whole, fontSize: string) => `${Number(fontSize) * scale}px`)
      commands.push({ kind: 'text', x: x(glyph.x), y: y(baseline), text: glyph.text === '\t' ? '    ' : glyph.text, font, color: glyph.run.href ? '#2563b6' : glyph.run.color, clip })
      if (glyph.run.underline || glyph.run.href) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline + 2), x2: x(glyph.x + glyph.width), y2: y(baseline + 2), color: glyph.run.color, width: Math.max(0.7, scale), clip })
      if (glyph.run.strike) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline - glyph.run.fontSize * 0.4), x2: x(glyph.x + glyph.width), y2: y(baseline - glyph.run.fontSize * 0.4), color: glyph.run.color, width: Math.max(0.7, scale), clip })
      if (options.composing && sameTarget(line.target, options.selection.focus) && glyph.end <= options.selection.focus.offset) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(line.baseline + 3), x2: x(glyph.x + glyph.width), y2: y(line.baseline + 3), color: '#3478df', width: 1, clip })
    }
    if (hasSelection && !line.glyphs.length && start <= line.start && end >= line.end) commands.push({ kind: 'rect', x: x(line.x), y: y(line.y), width: 5 * scale, height: line.height * scale, fill: '#bdd8ff', clip })
  }
  if (options.focused && options.caretVisible && selectionCollapsed(options.selection)) {
    const caret = caretRect(layout, options.selection.focus)
    if (caret && visible(caret)) commands.push({ kind: 'rect', ...rect(caret), width: Math.max(1, scale), fill: '#24569c' })
  }
  return { width: viewport.width, height: viewport.height, background: '#eef1f5', commands }
}

export function sceneOffset(layout: DocumentLayout, viewport: DocumentViewport): { x: number; y: number } {
  return { x: Math.max(0, (viewport.width - layout.width * viewport.zoom / 100) / 2) - viewport.scrollLeft, y: -viewport.scrollTop }
}
