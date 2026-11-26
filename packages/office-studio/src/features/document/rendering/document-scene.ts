import type { DocumentModel } from '../../../domain/model'
import type { DrawCommand, RenderRect, RenderScene } from '../../../rendering/contracts'
import { intersectRect, parseColor } from '../../../rendering/geometry'
import type { DocumentSelection } from '../model/document-model'
import { blockAtPoint, documentTextColor, findDocumentBlock, orderedSelection, sameTarget, selectionCollapsed, tableForBlock, targetKey, textTargets, type TextTarget } from '../model/document-model'
import { caretRect, documentStoryCaretRect, type DocumentLayout, type LayoutStoryLine } from '../layout/document-layout'
import { codeBlockData, codeBlocks as documentCodeBlocks, tokenizeCode, type CodeTokenKind } from '../code-blocks/code-blocks.ts'
import { documentTocHighlightLayout } from '../toc/document-toc-layout'
import { appendDocumentTableCell } from './document-table-scene'
import { documentTableCellsInRange, documentTableGrid, expandDocumentTableRange } from '../model/document-table-grid'
import type { DocumentTableGesturePreview } from '../layout/document-table-interaction'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { appendDocumentContainerDecoration } from './document-container-scene'
import type { LayoutImage } from '../layout/document-layout-types'
import { documentImageFrameCorners, documentImageRotationHandle, type DocumentImageFrame } from '../layout/document-image-layout'

export interface DocumentViewport { width: number; height: number; scrollLeft: number; scrollTop: number; zoom: number }
export interface DocumentImageScenePreview { imageId: string; page?: number; rectangle: DocumentImageFrame; wrappingStyle?: LayoutImage['wrappingStyle'] }
export interface DocumentSceneOptions { selection: DocumentSelection; focused: boolean; caretVisible: boolean; composing: boolean; images: Map<string, HTMLImageElement>; imageSelection?: DocumentImageFrame; imageSelections?: DocumentImageFrame[]; imageHandles?: boolean; imagePreview?: DocumentImageScenePreview; imagePreviews?: DocumentImageScenePreview[]; imageHandleColor?: string; tocPrimaryColor?: string; formulaErrorColor?: string; tableGesture?: DocumentTableGesturePreview; searchMatches?: Array<{ target: TextTarget; start: number; end: number }>; host?: DocumentAuxiliaryBlockHost; selectionModel?: DocumentModel; selectionLayout?: DocumentLayout; story?: { location: { page: number; area: 'header' | 'footer' }; selection: { start: number; end: number } } }

export function sceneFromDocument(model: DocumentModel, layout: DocumentLayout, viewport: DocumentViewport, options: DocumentSceneOptions): RenderScene {
  const commands: DrawCommand[] = []
  const scale = viewport.zoom / 100
  const { x: pageOffsetX, y: pageOffsetY } = sceneOffset(layout, viewport)
  const x = (value: number) => pageOffsetX + value * scale
  const y = (value: number) => pageOffsetY + value * scale
  const rect = (source: RenderRect): RenderRect => ({ x: x(source.x), y: y(source.y), width: source.width * scale, height: source.height * scale })
  const visible = (source: RenderRect) => y(source.y + source.height) >= -20 && y(source.y) <= viewport.height + 20 && x(source.x + source.width) >= -20 && x(source.x) <= viewport.width + 20
  const imagePreviews = new Map([...(options.imagePreview ? [options.imagePreview] : []), ...(options.imagePreviews ?? [])].map(preview => [preview.imageId, preview]))
  const projectedImage = (image: LayoutImage): LayoutImage => {
    const preview = imagePreviews.get(image.blockId)
    return preview && (preview.page === undefined || preview.page === image.page)
      ? { ...image, ...preview.rectangle, wrappingStyle: preview.wrappingStyle ?? image.wrappingStyle } : image
  }
  const appendImage = (image: LayoutImage, source: HTMLImageElement, clip: RenderRect) => {
    const width = image.paintWidth ?? image.width, height = image.paintHeight ?? image.height, crop = image.crop
    commands.push({ kind: 'image', x: x(image.x + (image.width - width) / 2), y: y(image.y + (image.height - height) / 2),
      width: width * scale, height: height * scale, source, clip, rotation: image.angle, flipX: image.flipX, flipY: image.flipY,
      ...(crop ? { sourceRect: { x: crop.left, y: crop.top, width: 1 - crop.left - crop.right, height: 1 - crop.top - crop.bottom } } : {}) })
  }
  const primary = parseColor(options.tocPrimaryColor ?? '#1677ff').slice(0, 3).map(value => Math.round(value * 255)).join(', ')
  const selectionModel = options.selectionModel ?? model, selectionLayout = options.selectionLayout ?? layout
  const tableView = options.story ? undefined : options.selection.tableSelection
  const tableBlock = tableView && findDocumentBlock(selectionModel, tableView.tableId), table = tableBlock && tableForBlock(tableBlock)
  let selectedCells = new Set<string>()
  if (table && tableView) {
    try { const grid = documentTableGrid(table); selectedCells = new Set(documentTableCellsInRange(grid, expandDocumentTableRange(grid, tableView.range)).map(owner => owner.cell.id)) } catch { /* Stale derived input has no body highlight. */ }
  }
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
    if (headerFooter && !headerFooter.isParityBlank) {
      const append = (lines: readonly LayoutStoryLine[], area: 'header' | 'footer') => {
        const active = options.story?.location.page === page.index && options.story.location.area === area
        const selectedStart = active ? Math.min(options.story!.selection.start, options.story!.selection.end) : 0
        const selectedEnd = active ? Math.max(options.story!.selection.start, options.story!.selection.end) : 0
        for (const line of lines) {
          if (!visible(line)) continue
          let group: { text: string; x: number; y: number; font: string; color: string; runId: string } | undefined
          const flush = () => { if (group) { const { runId: _runId, ...text } = group; commands.push({ kind: 'text', ...text, clip: rect(page) }); group = undefined } }
          for (const glyph of line.glyphs) {
            const color = documentTextColor(glyph.run.color)
            const baseline = line.baseline + (glyph.run.verticalAlign === 'superscript' ? -glyph.run.fontSize * 0.4 : glyph.run.verticalAlign === 'subscript' ? glyph.run.fontSize * 0.2 : 0)
            if (glyph.run.highlight) commands.push({ kind: 'rect', ...rect(glyph), fill: glyph.run.highlight, clip: rect(page) })
            if (active && glyph.end > selectedStart && glyph.offset < selectedEnd) commands.push({ kind: 'rect', ...rect(glyph), fill: options.focused ? '#bdd8ff' : '#dde6f2', clip: rect(page) })
            const font = glyph.font.replace(/([\d.]+)px/, (_whole, size: string) => `${Number(size) * scale}px`)
            if (glyph.tab) {
              flush()
              if (glyph.tab.leaderText) commands.push({ kind: 'text', x: x(glyph.x), y: y(baseline), text: glyph.tab.leaderText, font, color, maxWidth: glyph.width * scale, clip: rect(page) })
            } else {
              if (group && (group.runId !== glyph.run.id || group.font !== font || group.color !== color || group.y !== y(baseline))) flush()
              if (group) group.text += glyph.text === '\t' ? '    ' : glyph.text
              else group = { text: glyph.text === '\t' ? '    ' : glyph.text, x: x(glyph.x), y: y(baseline), font, color, runId: glyph.run.id }
            }
            if (glyph.run.underline) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline + 2), x2: x(glyph.x + glyph.width), y2: y(baseline + 2), color, width: Math.max(0.7, scale), clip: rect(page) })
            if (glyph.run.strike) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline - glyph.run.fontSize * 0.4), x2: x(glyph.x + glyph.width), y2: y(baseline - glyph.run.fontSize * 0.4), color, width: Math.max(0.7, scale), clip: rect(page) })
          }
          flush()
        }
        if (active && options.focused && options.caretVisible && selectedStart === selectedEnd) {
          const caret = documentStoryCaretRect(lines, selectedStart)
          if (caret && visible(caret)) commands.push({ kind: 'rect', ...rect(caret), width: Math.max(1, scale), fill: '#24569c', clip: rect(page) })
        }
      }
      if (!headerFooter.headerBlockLayout) append(headerFooter.headerLayoutLines ?? [], 'header')
      if (!headerFooter.footerBlockLayout) append(headerFooter.footerLayoutLines ?? [], 'footer')
    }
  }
  for (const pageIndex of visiblePages) {
    const body = layout.pageContent[pageIndex], auxiliary = [...(layout.auxiliaryContent ?? []), ...(layout.noteDecorations ?? [])].filter(content => content.page === pageIndex)
    if (!body && !auxiliary.length) continue
    const content = { cells: [...(body?.cells ?? []), ...auxiliary.flatMap(content => content.cells)], images: [...(body?.images ?? []), ...auxiliary.flatMap(content => content.images)],
      dividers: [...(body?.dividers ?? []), ...auxiliary.flatMap(content => content.dividers)], columnSeparators: body?.columnSeparators ?? [],
      containers: [...(body?.containers ?? []), ...auxiliary.flatMap(content => content.containers ?? [])] }
    const backgrounds = [...content.cells.map(cell => ({ kind: 'cell' as const, item: cell, layer: (cell.depth ?? 0) * 2 })),
      ...content.containers.map(container => ({ kind: 'container' as const, item: container, layer: container.tableDepth === undefined ? -1 : container.tableDepth * 2 - 1 }))]
      .sort((a, b) => a.layer - b.layer || b.item.width * b.item.height - a.item.width * a.item.height)
    for (const background of backgrounds) {
      if (!visible(background.item)) continue
      if (background.kind === 'cell') appendDocumentTableCell(commands, background.item, rect(background.item), rect(layout.pages[pageIndex]!), scale,
        background.item.clip ? rect(background.item.clip) : undefined)
      else appendDocumentContainerDecoration(commands, background.item, rect(background.item), background.item.clip
        ? intersectRect(rect(layout.pages[pageIndex]!), rect(background.item.clip)) : rect(layout.pages[pageIndex]!), scale)
    }
    for (const cell of content.cells) {
      if (cell.tableId !== tableView?.tableId || !cell.cellId || !selectedCells.has(cell.cellId) || !visible(cell)) continue
      const clip = cell.clip ? intersectRect(rect(layout.pages[pageIndex]!), rect(cell.clip)) : rect(layout.pages[pageIndex]!)
      commands.push({ kind: 'rect', ...rect(cell), fill: `rgba(${primary}, ${options.focused ? 0.16 : 0.10})`, stroke: `rgba(${primary}, 0.65)`, strokeWidth: Math.max(1, scale), clip })
    }
    for (const divider of content.dividers) {
      if (!visible(divider) || divider.height <= 0) continue
      const width = Math.max(1, divider.height * scale), color = documentTextColor(divider.color ?? '#b3bdca')
      const clip = divider.clip ? intersectRect(rect(layout.pages[pageIndex]!), rect(divider.clip)) : rect(layout.pages[pageIndex]!)
      if (!divider.style || divider.style === 'solid') {
        commands.push({ kind: 'line', x1: x(divider.x), y1: y(divider.y), x2: x(divider.x + divider.width), y2: y(divider.y), color, width, clip })
        continue
      }
      const segment = divider.style === 'dot' ? Math.max(1, width) : Math.max(4, width * 3), gap = Math.max(2, width * 2)
      for (let from = 0; from < divider.width * scale; from += segment + gap) {
        commands.push({ kind: 'line', x1: x(divider.x) + from, y1: y(divider.y), x2: x(divider.x) + Math.min(divider.width * scale, from + segment),
          y2: y(divider.y), color, width, clip })
      }
    }
    for (const separator of content.columnSeparators) {
      if (visible(separator)) commands.push({ kind: 'line', x1: x(separator.x), y1: y(separator.y), x2: x(separator.x), y2: y(separator.y + separator.height),
        color: '#b3bdca', width: Math.max(1, scale), clip: rect(layout.pages[pageIndex]!) })
    }
    for (const image of content.images.map(projectedImage).filter(image => image.wrappingStyle !== 'inFrontOfText').sort((a, b) => (a.layer ?? 0) - (b.layer ?? 0))) {
      if (!visible(image)) continue
    const source = options.images.get(image.src)
    const clip = image.clip ? intersectRect(rect(layout.pages[pageIndex]!), rect(image.clip)) : rect(layout.pages[pageIndex]!)
    if (source?.complete && source.naturalWidth) appendImage(image, source, clip)
    else {
      commands.push({ kind: 'rect', ...rect(image), fill: '#f1f4f8', stroke: '#ccd5e0', clip })
      commands.push({ kind: 'text', x: x(image.x + 12), y: y(image.y + 22), text: image.alt || '图片', font: `${12 * scale}px Arial`, color: '#738096', clip })
    }
    }
  }
  const tocHighlight = options.story || options.host || tableView ? undefined : documentTocHighlightLayout(model, layout, options.selection)
  for (const range of tocHighlight?.ranges ?? []) {
    if (!visiblePages.has(range.page) || !visible(range)) continue
    commands.push({ kind: 'rect', ...rect(range), fill: `rgba(${primary}, 0.08)`, clip: rect(layout.pages[range.page]!) })
  }
  const ordered = orderedSelection(selectionModel, options.selection)
  const targets = textTargets(selectionModel)
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
  const visibleLines = [...visiblePages].flatMap((pageIndex) => [...(layout.pageContent[pageIndex]?.lines ?? []),
    ...[...(layout.auxiliaryContent ?? []), ...(layout.noteDecorations ?? [])].filter(content => content.page === pageIndex).flatMap(content => content.lines)])
  const codeBlocks = new Map(documentCodeBlocks(selectionModel).flatMap(block => { const data = codeBlockData(block); return data ? [[block.id, data] as const] : [] }))
  const codeTokens = new Map([...codeBlocks].map(([blockId, data]) => [blockId, tokenizeCode(data.source, data.language)]))
  for (const line of visibleLines) {
    if (!visible(line) && !(line.codeBlock?.gutterClip && visible(line.codeBlock.gutterClip))) continue
    const clip = line.clip ? intersectRect(rect(layout.pages[line.page]!), rect(line.clip)) : rect(layout.pages[line.page]!)
    if (clip.width <= 0 || clip.height <= 0) continue
    const hasSelection = !options.story && !tableView && !selectionCollapsed(options.selection) && selected.has(targetKey(line.target))
    const lineMatches = matchesByTarget.get(targetKey(line.target)) ?? []
    const start = sameTarget(line.target, ordered.start) ? ordered.start.offset : 0
    const end = sameTarget(line.target, ordered.end) ? ordered.end.offset : Number.POSITIVE_INFINITY
    const codeKey = line.codeBlock?.containerId ?? line.target.cellBlockId ?? line.target.blockId
    const codeData = line.codeBlock ? codeBlocks.get(codeKey) : undefined
    if (line.codeBlock && !line.codeBlock.containerId) {
      commands.push({ kind: 'rect', x: x(line.x - 8), y: y(line.y), width: (line.width + 16) * scale, height: line.height * scale, fill: '#f1f4f8', clip })
      if (line.codeBlock.lineNumber === 1) commands.push({ kind: 'rect', x: x(line.x - 8), y: y(line.y), width: (line.width + 16) * scale, height: Math.max(2, scale), fill: '#dce3ec', clip })
      if (codeData?.lineNumbers) commands.push({ kind: 'text', x: x(line.x - 3), y: y(line.baseline), text: String(line.codeBlock.lineNumber), font: `${10 * scale}px Consolas, monospace`, color: '#8a96a5', align: 'right', baseline: 'alphabetic', clip })
    }
    if (line.codeBlock?.containerId && line.codeBlock.showLineNumbers && line.codeBlock.numberVisible) commands.push({ kind: 'text',
      x: x((line.codeBlock.gutterX ?? line.x) - 10), y: y(line.baseline), text: String(line.codeBlock.lineNumber), font: `${10 * scale}px Consolas, monospace`, color: '#8a96a5', align: 'right', baseline: 'alphabetic',
      clip: line.codeBlock.gutterClip ? intersectRect(rect(layout.pages[line.page]!), rect(line.codeBlock.gutterClip)) : clip })
    if (line.marker) {
      const font = line.markerStyle?.font.replace(/([\d.]+)px/, (_whole, fontSize: string) => `${Number(fontSize) * scale}px`) ?? `${14 * scale}px Arial`
      commands.push({ kind: 'text', x: x(line.markerStyle?.x ?? line.x - 6), y: y(line.baseline), text: line.marker, font,
        color: line.markerStyle ? documentTextColor(line.markerStyle.color) : '#445166', align: line.markerStyle ? 'left' : 'right', baseline: 'alphabetic', clip })
    }
    for (const glyph of line.glyphs) {
      const textColor = documentTextColor(glyph.run.color)
      const hyperlink = glyph.run.metadata?.link as { textStyleMode?: string; textColorMode?: string } | undefined
      const hyperlinkStyle = glyph.run.href && hyperlink?.textStyleMode !== 'text'
      const hyperlinkColor = hyperlinkStyle && hyperlink?.textColorMode !== 'text' ? '#2563b6' : textColor
      const baselineOffset = glyph.run.verticalAlign === 'superscript' ? -glyph.run.fontSize * 0.4 : glyph.run.verticalAlign === 'subscript' ? glyph.run.fontSize * 0.2 : 0
      const baseline = line.baseline + baselineOffset
      if (glyph.run.highlight) commands.push({ kind: 'rect', ...rect(glyph), fill: glyph.run.highlight, clip })
      if (overlapsSearchMatch(lineMatches, glyph.offset, glyph.end)) commands.push({ kind: 'rect', ...rect(glyph), fill: '#ffe58f', clip })
      if (hasSelection && glyph.end > start && glyph.offset < end) commands.push({ kind: 'rect', ...rect(glyph), fill: options.focused ? '#bdd8ff' : '#dde6f2', clip })
      if (glyph.latex) {
        const formulaColor = glyph.latexError ? options.formulaErrorColor ?? '#ff4d4f' : textColor
        if (glyph.latexError) commands.push({ kind: 'rect', ...rect(glyph), fill: 'rgba(0,0,0,0)', stroke: formulaColor, strokeWidth: Math.max(1, scale), clip })
        for (const primitive of glyph.latex.primitives) {
          if (primitive.kind === 'text') {
            const font = primitive.font.replace(/([\d.]+)px/, (_whole, fontSize: string) => `${Number(fontSize) * scale}px`)
            commands.push({ kind: 'text', x: x(glyph.x + primitive.x), y: y(baseline + primitive.y), text: primitive.text, font, color: formulaColor, baseline: 'alphabetic', clip })
          } else {
            commands.push({ kind: 'line', x1: x(glyph.x + primitive.x1), y1: y(baseline + primitive.y1), x2: x(glyph.x + primitive.x2), y2: y(baseline + primitive.y2), color: formulaColor, width: Math.max(0.8, primitive.width * scale), clip })
          }
        }
        if (glyph.run.underline || hyperlinkStyle) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline + 2), x2: x(glyph.x + glyph.width), y2: y(baseline + 2), color: formulaColor, width: Math.max(0.7, scale), clip })
        if (glyph.run.strike) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline - glyph.run.fontSize * 0.4), x2: x(glyph.x + glyph.width), y2: y(baseline - glyph.run.fontSize * 0.4), color: formulaColor, width: Math.max(0.7, scale), clip })
        continue
      }
      const font = glyph.font.replace(/([\d.]+)px/, (_whole, fontSize: string) => `${Number(fontSize) * scale}px`)
      const codeOffset = glyph.offset + (line.codeBlock?.sourceOffset ?? 0)
      const token = codeData && codeTokens.get(codeKey)?.find((candidate) => codeOffset >= candidate.start && codeOffset < candidate.end)
      if (glyph.imageAtom) continue
      const text = glyph.tab ? glyph.tab.leaderText : glyph.text === '\t' ? '    ' : glyph.text
      if (text) commands.push({ kind: 'text', x: x(glyph.x), y: y(baseline), text, font, color: token && glyph.run.metadata?.codeSyntaxHighlight !== false ? codeTokenColor(token.kind) : hyperlinkColor,
        ...(glyph.tab ? { maxWidth: glyph.width * scale } : {}), clip })
      if (glyph.run.underline || hyperlinkStyle) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline + 2), x2: x(glyph.x + glyph.width), y2: y(baseline + 2), color: hyperlinkColor, width: Math.max(0.7, scale), clip })
      if (glyph.run.strike) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(baseline - glyph.run.fontSize * 0.4), x2: x(glyph.x + glyph.width), y2: y(baseline - glyph.run.fontSize * 0.4), color: textColor, width: Math.max(0.7, scale), clip })
      if (options.composing && sameTarget(line.target, options.selection.focus) && glyph.end <= options.selection.focus.offset) commands.push({ kind: 'line', x1: x(glyph.x), y1: y(line.baseline + 3), x2: x(glyph.x + glyph.width), y2: y(line.baseline + 3), color: '#3478df', width: 1, clip })
    }
    if (hasSelection && !line.glyphs.length && start <= line.start && end >= line.end) commands.push({ kind: 'rect', x: x(line.x), y: y(line.y), width: 5 * scale, height: line.height * scale, fill: '#bdd8ff', clip })
  }
  const foregroundImages: LayoutImage[] = [...visiblePages].flatMap(page => [...(layout.pageContent[page]?.images ?? []),
    ...[...(layout.auxiliaryContent ?? []), ...(layout.noteDecorations ?? [])].filter(content => content.page === page).flatMap(content => content.images)])
    .map(projectedImage).filter(image => image.wrappingStyle === 'inFrontOfText').sort((a, b) => (a.layer ?? 0) - (b.layer ?? 0))
  for (const image of foregroundImages) {
    if (!visible(image)) continue
    const source = options.images.get(image.src), page = layout.pages[image.page]
    if (!page) continue
    const clip = image.clip ? intersectRect(rect(page), rect(image.clip)) : rect(page)
    if (source?.complete && source.naturalWidth) appendImage(image, source, clip)
    else commands.push({ kind: 'rect', ...rect(image), fill: '#f1f4f8', stroke: '#ccd5e0', clip })
  }
  for (const outline of tocHighlight?.outlines ?? []) {
    if (!visiblePages.has(outline.page) || !visible(outline)) continue
    commands.push({ kind: 'rect', ...rect(outline), fill: 'rgba(0, 0, 0, 0)', stroke: `rgba(${primary}, 0.8)`, strokeWidth: Math.max(1, scale), clip: rect(layout.pages[outline.page]!) })
  }
  if (!options.story && !options.imageSelection && options.focused && options.caretVisible && selectionCollapsed(options.selection)) {
    const caret = caretRect(selectionLayout, options.selection.focus)
    if (caret && visible(caret)) commands.push({ kind: 'rect', ...rect(caret), width: Math.max(1, scale), fill: '#24569c' })
  }
  const gesture = options.tableGesture, guideColor = gesture?.available ? `rgba(${primary}, 0.9)` : '#ff4d4f'
  if (gesture?.rectangle) commands.push({ kind: 'rect', ...rect(gesture.rectangle), fill: `rgba(${primary}, 0.08)`, stroke: guideColor, strokeWidth: Math.max(1, scale) })
  if (gesture?.guide) commands.push({ kind: 'line', x1: x(gesture.guide.x1), y1: y(gesture.guide.y1), x2: x(gesture.guide.x2), y2: y(gesture.guide.y2), color: guideColor, width: Math.max(1, scale) })
  const imageSelectionColor = `rgb(${primary})`
  function appendImageOutline(frame: DocumentImageFrame, handles: boolean) {
    if (!visible(frame)) return
    const corners = documentImageFrameCorners(frame), color = imageSelectionColor
    for (const [index, point] of corners.entries()) {
      const next = corners[(index + 1) % corners.length]!
      commands.push({ kind: 'line', x1: x(point.x), y1: y(point.y), x2: x(next.x), y2: y(next.y), color, width: 1.5 })
      if (handles) commands.push({ kind: 'rect', x: x(point.x) - 4, y: y(point.y) - 4, width: 8, height: 8, fill: options.imageHandleColor ?? '#fff', stroke: color, strokeWidth: 1 })
    }
    if (handles) {
      const rotation = documentImageRotationHandle(frame, viewport.zoom)
      commands.push({ kind: 'line', x1: x(rotation.stemX), y1: y(rotation.stemY), x2: x(rotation.x), y2: y(rotation.y), color, width: 1 })
      commands.push({ kind: 'rect', x: x(rotation.x) - 5, y: y(rotation.y) - 5, width: 10, height: 10, fill: options.imageHandleColor ?? '#fff', stroke: color, strokeWidth: 1 })
    }
  }
  for (const frame of options.imageSelections ?? []) appendImageOutline(frame, false)
  if (options.imageSelection) appendImageOutline(options.imageSelection, options.imageHandles !== false)
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
