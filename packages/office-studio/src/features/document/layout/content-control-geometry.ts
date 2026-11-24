import type { DocumentContentControl, DocumentModel } from '../../../domain/model'
import { sameTarget, textTargets, type DocumentSelection, type TextTarget } from '../model/document-model'
import type { DocumentLayout, LayoutLine, LayoutRect } from './document-layout-types'

export interface LayoutContentControlFragment extends LayoutRect {
  page: number
  startOffset: number
  endOffset: number
  first: boolean
  last: boolean
  clip?: LayoutRect
}

export interface DocumentContentControlGeometry {
  control: DocumentContentControl
  target: TextTarget
  selection: DocumentSelection
  fragments: LayoutContentControlFragment[]
}

export function intersectContentControlRect(left: LayoutRect, right: LayoutRect): LayoutRect | undefined {
  const x = Math.max(left.x, right.x), y = Math.max(left.y, right.y)
  const width = Math.min(left.x + left.width, right.x + right.width) - x
  const height = Math.min(left.y + left.height, right.y + right.height) - y
  return width > 0 && height > 0 ? { x, y, width, height } : undefined
}

export function contentControlFragmentVisibleRect(fragment: LayoutContentControlFragment): LayoutRect | undefined {
  return fragment.clip ? intersectContentControlRect(fragment, fragment.clip) : fragment
}

function xForOffset(line: LayoutLine, offset: number): number {
  const glyph = line.glyphs.find(glyph => glyph.offset <= offset && offset < glyph.end)
  if (glyph) return glyph.x
  const before = [...line.glyphs].reverse().find(glyph => glyph.end <= offset)
  return before ? before.x + before.width : line.x
}

function fragmentForLine(control: DocumentContentControl, line: LayoutLine, layout: DocumentLayout): LayoutContentControlFragment | undefined {
  const startOffset = Math.max(control.startOffset, line.start), endOffset = Math.min(control.endOffset, line.end)
  const glyphs = line.glyphs.filter(glyph => glyph.offset < endOffset && glyph.end > startOffset)
  const left = glyphs.length ? Math.min(...glyphs.map(glyph => glyph.x)) : xForOffset(line, startOffset)
  const right = glyphs.length ? Math.max(...glyphs.map(glyph => glyph.x + glyph.width)) : left
  const page = layout.pages.find(page => page.index === line.page)
  const clip = line.clip && page ? intersectContentControlRect(line.clip, page) : line.clip ?? page
  if (line.clip && page && !clip) return undefined
  return { x: left, y: line.y, width: Math.max(right - left, control.startOffset === control.endOffset ? 12 : 1), height: line.height,
    page: line.page, startOffset, endOffset, first: startOffset === control.startOffset, last: endOffset === control.endOffset,
    ...(clip ? { clip: { x: clip.x, y: clip.y, width: clip.width, height: clip.height } } : {}) }
}

/** Paragraph/cell IDs and UTF-16 offsets drive geometry; DOM text is never an input. */
export function documentContentControlGeometry(model: DocumentModel, layout: DocumentLayout): DocumentContentControlGeometry[] {
  const controls = new Map((model.contentControls ?? []).map(control => [control.id, control]))
  const depths = new Map<string, number>()
  function depth(control: DocumentContentControl): number {
    const known = depths.get(control.id)
    if (known !== undefined) return known
    let result = 0, parent = control.parentId
    const visited = new Set([control.id])
    while (parent && !visited.has(parent)) {
      visited.add(parent)
      const ancestor = controls.get(parent)
      if (!ancestor) break
      result++; parent = ancestor.parentId
    }
    depths.set(control.id, result)
    return result
  }
  const targets = new Map(textTargets(model).map(target => [target.cellBlockId ?? target.blockId, target]))
  const lines = new Map<string, LayoutLine[]>()
  for (const line of layout.lines) {
    const id = line.target.cellBlockId ?? line.target.blockId, items = lines.get(id) ?? []
    items.push(line); lines.set(id, items)
  }
  const result: DocumentContentControlGeometry[] = []
  for (const control of model.contentControls ?? []) {
    if (control.metadata?.officeStoryLockSource) continue
    const target = targets.get(control.blockId)
    if (!target) continue
    const candidates = (lines.get(control.blockId) ?? []).filter(line => sameTarget(line.target, target))
    let matched: LayoutLine[]
    if (control.startOffset === control.endOffset) {
      // At a wrap boundary the following line owns the caret. Repeated headers keep one replica per page.
      const pages = new Set(candidates.map(line => line.page))
      matched = [...pages].flatMap(page => {
        const pageLines = candidates.filter(line => line.page === page)
        const line = pageLines.find(line => line.start === control.startOffset)
          ?? pageLines.find(line => line.start < control.startOffset && control.startOffset < line.end)
          ?? [...pageLines].reverse().find(line => line.end === control.startOffset)
        return line ? [line] : []
      })
    } else matched = candidates.filter(line => line.start < control.endOffset && line.end > control.startOffset)
    let fragments = matched.flatMap(line => {
      const fragment = fragmentForLine(control, line, layout)
      return fragment ? [fragment] : []
    })
    const pictureId = control.type === 'picture' && typeof control.metadata?.pictureBlockId === 'string' ? control.metadata.pictureBlockId : undefined
    if (pictureId) {
      const pictures = layout.images.filter(image => image.blockId === pictureId)
      if (pictures.length) fragments = pictures.flatMap(image => {
        const page = layout.pages.find(page => page.index === image.page)
        const clip = image.clip && page ? intersectContentControlRect(image.clip, page) : image.clip ?? page
        if (image.clip && page && !clip) return []
        return [{ x: image.x, y: image.y, width: image.width, height: image.height, page: image.page,
          startOffset: control.startOffset, endOffset: control.endOffset, first: true, last: true,
          ...(clip ? { clip: { x: clip.x, y: clip.y, width: clip.width, height: clip.height } } : {}) }]
      })
    }
    fragments = fragments.filter(fragment => Number.isFinite(fragment.x) && Number.isFinite(fragment.y) && fragment.width > 0 && fragment.height > 0
      && contentControlFragmentVisibleRect(fragment))
    if (!fragments.length) continue
    const selection = { anchor: { ...target, offset: control.startOffset }, focus: { ...target, offset: control.endOffset } }
    result.push({ control, target, selection, fragments })
  }
  return result.sort((left, right) => (right.control.endOffset - right.control.startOffset) - (left.control.endOffset - left.control.startOffset)
    || depth(left.control) - depth(right.control)
    || left.control.id.localeCompare(right.control.id))
}

export function hitTestDocumentContentControl(geometries: readonly DocumentContentControlGeometry[], x: number, y: number): DocumentContentControlGeometry | undefined {
  return [...geometries].reverse().find(geometry => geometry.fragments.some(fragment => {
    const rect = contentControlFragmentVisibleRect(fragment)
    return rect && x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height
  }))
}

export function contentControlSelectionAtFragment(geometry: DocumentContentControlGeometry, fragment: LayoutContentControlFragment): DocumentSelection {
  return { anchor: { ...geometry.selection.anchor, pageAffinity: fragment.page }, focus: { ...geometry.selection.focus, pageAffinity: fragment.page } }
}
