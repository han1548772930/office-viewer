import type { DocumentModel } from '../../../domain/model'
import type { DrawCommand, RenderRect, RenderScene } from '../../../rendering/contracts'
import type { DocumentSelection } from '../model/document-model'
import { documentTextColor, orderedSelection, sameTarget, selectionCollapsed, targetKey, textTargets, type TextTarget } from '../model/document-model'
import { caretRect, STORY_LINE_HEIGHT, type DocumentLayout } from '../layout/document-layout'
import { codeBlockData, tokenizeCode, type CodeTokenKind } from '../code-blocks/code-blocks.ts'

export interface DocumentViewport { width: number; height: number; scrollLeft: number; scrollTop: number; zoom: number }
export interface DocumentSceneOptions { selection: DocumentSelection; focused: boolean; caretVisible: boolean; composing: boolean; images: Map<string, HTMLImageElement>; searchMatches?: Array<{ target: TextTarget; start: number; end: number }> }

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
    const headerFooter = layout.pageHeaderFooters?.[page.index]
    if (headerFooter) {
      const append = (text: string, top: number) => {
        if (text && y(top + STORY_LINE_HEIGHT) >= 0 && y(top) <= viewport.height) commands.push({ kind: 'text', x: x(headerFooter.x + headerFooter.width / 2), y: y(top), text, font: `${11 * scale}px Arial`, color: '#5b6470', align: 'center', baseline: 'top', clip: rect(page) })
      }
      headerFooter.headerLines.forEach((text, i) => append(text, headerFooter.headerTop + i * STORY_LINE_HEIGHT))
      headerFooter.footerLines.forEach((text, i, lines) => append(text, headerFooter.footerBottom - (lines.length - i) * STORY_LINE_HEIGHT))
    }
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
  const matchesByTarget = new Map<string, Array<{ start: number; end: number }>>()
  for (const match of options.searchMatches ?? []) {
    const key = targetKey(match.target)
    const matches = matchesByTarget.get(key) ?? []
    matches.push({ start: match.start, end: match.end })
    matchesByTarget.set(key, matches)
  }
  for (const matches of matchesByTarget.values()) matches.sort((a, b) => a.start - b.start)
  const visibleLines = [...visiblePages].flatMap((pageIndex) => layout.pageContent[pageIndex]?.lines ?? [])
  const codeBlocks = new Map(model.blocks.map((block) => [block.id, codeBlockData(block)]).filter((entry): entry is [string, NonNullable<ReturnType<typeof codeBlockData>>] => entry[1] !== undefined))
  const codeTokens = new Map([...codeBlocks].map(([blockId, data]) => [blockId, tokenizeCode(data.source, data.language)]))
  for (const line of visibleLines) {
    if (!visible(line)) continue
    const clip = rect(layout.pages[line.page]!.content)
    const hasSelection = !selectionCollapsed(options.selection) && selected.has(targetKey(line.target))
    const lineMatches = matchesByTarget.get(targetKey(line.target)) ?? []
    const start = sameTarget(line.target, ordered.start) ? ordered.start.offset : 0
    const end = sameTarget(line.target, ordered.end) ? ordered.end.offset : Number.POSITIVE_INFINITY
    const codeData = line.codeBlock ? codeBlocks.get(line.target.blockId) : undefined
    if (line.codeBlock) {
      commands.push({ kind: 'rect', x: x(line.x - 8), y: y(line.y), width: (line.width + 16) * scale, height: line.height * scale, fill: '#f1f4f8', clip })
      if (line.codeBlock.lineNumber === 1) commands.push({ kind: 'rect', x: x(line.x - 8), y: y(line.y), width: (line.width + 16) * scale, height: Math.max(2, scale), fill: '#dce3ec', clip })
      commands.push({ kind: 'text', x: x(line.x - 3), y: y(line.baseline), text: String(line.codeBlock.lineNumber), font: `${10 * scale}px Consolas, monospace`, color: '#8a96a5', align: 'right', baseline: 'alphabetic', clip })
    }
    if (line.marker) commands.push({ kind: 'text', x: x(line.x - 6), y: y(line.baseline), text: line.marker, font: `${14 * scale}px Arial`, color: '#445166', align: 'right', clip })
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
      const textColor = documentTextColor(glyph.run.color)
      const baselineOffset = glyph.run.verticalAlign === 'superscript' ? -glyph.run.fontSize * 0.4 : glyph.run.verticalAlign === 'subscript' ? glyph.run.fontSize * 0.2 : 0
      const baseline = line.baseline + baselineOffset
      if (glyph.run.highlight) commands.push({ kind: 'rect', ...rect(glyph), fill: glyph.run.highlight, clip })
      if (overlapsSearchMatch(lineMatches, glyph.offset, glyph.end)) commands.push({ kind: 'rect', ...rect(glyph), fill: '#ffe58f', clip })
      if (hasSelection && glyph.end > start && glyph.offset < end) commands.push({ kind: 'rect', ...rect(glyph), fill: options.focused ? '#bdd8ff' : '#dde6f2', clip })
      if (glyph.latex) {
        for (const primitive of glyph.latex.primitives) {
          if (primitive.kind === 'text') {
            const font = primitive.font.replace(/([\d.]+)px/, (_whole, fontSize: string) => `${Number(fontSize) * scale}px`)
            commands.push({ kind: 'text', x: x(glyph.x + primitive.x), y: y(line.baseline + primitive.y), text: primitive.text, font, color: textColor, clip })
          } else {
            commands.push({ kind: 'line', x1: x(glyph.x + primitive.x1), y1: y(line.baseline + primitive.y1), x2: x(glyph.x + primitive.x2), y2: y(line.baseline + primitive.y2), color: textColor, width: Math.max(0.8, primitive.width * scale), clip })
          }
        }
        continue
      }
      const font = glyph.font.replace(/([\d.]+)px/, (_whole, fontSize: string) => `${Number(fontSize) * scale}px`)
      const token = codeData && codeTokens.get(line.target.blockId)?.find((candidate) => glyph.offset >= candidate.start && glyph.offset < candidate.end)
      commands.push({ kind: 'text', x: x(glyph.x), y: y(baseline), text: glyph.text === '\t' ? '    ' : glyph.text, font, color: token ? codeTokenColor(token.kind) : glyph.run.href ? '#2563b6' : textColor, clip })
      if (glyph.run.underline || glyph.run.href) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline + 2), x2: x(glyph.x + glyph.width), y2: y(baseline + 2), color: textColor, width: Math.max(0.7, scale), clip })
      if (glyph.run.strike) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline - glyph.run.fontSize * 0.4), x2: x(glyph.x + glyph.width), y2: y(baseline - glyph.run.fontSize * 0.4), color: textColor, width: Math.max(0.7, scale), clip })
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

function overlapsSearchMatch(matches: readonly { start: number; end: number }[], start: number, end: number): boolean {
  let low = 0
  let high = matches.length
  while (low < high) {
    const middle = Math.floor((low + high) / 2)
    if (matches[middle]!.end <= start) low = middle + 1
    else high = middle
  }
  return low < matches.length && matches[low]!.start < end
}

function codeTokenColor(kind: CodeTokenKind): string {
  if (kind === 'keyword') return '#7c3aed'
  if (kind === 'string') return '#15803d'
  if (kind === 'number') return '#b45309'
  if (kind === 'comment') return '#718096'
  if (kind === 'operator') return '#2563eb'
  return '#263238'
}

export function sceneOffset(layout: DocumentLayout, viewport: DocumentViewport): { x: number; y: number } {
  return { x: Math.max(0, (viewport.width - layout.width * viewport.zoom / 100) / 2) - viewport.scrollLeft, y: -viewport.scrollTop }
}
