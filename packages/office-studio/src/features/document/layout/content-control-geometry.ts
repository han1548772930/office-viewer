import type { DocumentContentControl, DocumentModel } from '../../../domain/model'
import { runText, runsForTarget, sameTarget, textTargets, type DocumentSelection, type TextTarget } from '../model/document-model'
import { cloneDocumentSelection } from '../model/document-table-selection'
import { documentContentControlContains, documentContentControlTableRowSelection, resolveDocumentContentControlTableRows,
  type DocumentContentControlTableRows } from '../model/content-control-table-rows'
import { documentContentControlBlocksSelection, documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from '../model/content-control-block-range'
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

function fragmentsForRows(control: DocumentContentControl, rows: DocumentContentControlTableRows, layout: DocumentLayout): LayoutContentControlFragment[] {
  const selected = new Set(rows.rowIds), result: LayoutContentControlFragment[] = []
  for (const table of layout.tableFragments ?? []) {
    if (table.tableId !== rows.block.id) continue
    const metrics = table.rows.filter(row => selected.has(row.id) && rows.table.rows[row.row]?.id === row.id)
    if (!metrics.length) continue
    const y = Math.min(...metrics.map(row => row.y)), bottom = Math.max(...metrics.map(row => row.y + row.height))
    const raw = { x: table.x, y, width: table.width, height: bottom - y }
    let rect = intersectContentControlRect(raw, table)
    const page = layout.pages.find(page => page.index === table.page)
    if (rect && page) rect = intersectContentControlRect(rect, page)
    if (rect && table.clip) rect = intersectContentControlRect(rect, table.clip)
    if (!rect || !Number.isFinite(rect.x) || !Number.isFinite(rect.y) || !Number.isFinite(rect.width) || !Number.isFinite(rect.height)) continue
    const first = metrics.find(row => row.id === rows.rowIds[0]), last = metrics.find(row => row.id === rows.rowIds.at(-1))
    result.push({ ...rect, page: table.page, startOffset: control.startOffset, endOffset: control.endOffset,
      first: !!first && rect.y <= first.y + 0.01, last: !!last && rect.y + rect.height >= last.y + last.height - 0.01,
      clip: { ...rect } })
  }
  return result
}

function fragmentsForBlocks(model: DocumentModel, control: DocumentContentControl, layout: DocumentLayout): LayoutContentControlFragment[] {
  const owned = documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, control.blockRange!))
  const frames = new Map<string, LayoutContentControlFragment>()
  const append = (raw: LayoutRect & { page: number; clip?: LayoutRect; sectionFragment?: number; column?: number }) => {
    const page = layout.pages.find(page => page.index === raw.page)
    let rect = page ? intersectContentControlRect(raw, page) : raw
    if (rect && raw.clip) rect = intersectContentControlRect(rect, raw.clip)
    if (!rect || ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)) return
    const key = JSON.stringify([raw.page, raw.sectionFragment, raw.column, raw.clip])
    const previous = frames.get(key)
    const x = previous ? Math.min(previous.x, rect.x) : rect.x, y = previous ? Math.min(previous.y, rect.y) : rect.y
    const right = previous ? Math.max(previous.x + previous.width, rect.x + rect.width) : rect.x + rect.width
    const bottom = previous ? Math.max(previous.y + previous.height, rect.y + rect.height) : rect.y + rect.height
    frames.set(key, { x, y, width: right - x, height: bottom - y, page: raw.page,
      startOffset: control.startOffset, endOffset: control.endOffset, first: false, last: false,
      ...(raw.clip ? { clip: raw.clip } : {}) })
  }
  for (const line of layout.lines) {
    if (!owned.has(line.target.cellBlockId ?? line.target.blockId)) continue
    const value = runText(runsForTarget(model, line.target))
    const fragment = fragmentForLine({ ...control, startOffset: 0, endOffset: value.length }, line, layout)
    if (fragment) append({ ...fragment, sectionFragment: line.sectionFragment, column: line.column })
  }
  for (const table of layout.tableFragments ?? []) if (owned.has(table.tableId)) append(table)
  for (const image of layout.images) if (owned.has(image.blockId)) append(image)
  const result = [...frames.values()].filter(fragment => fragment.width > 0 && fragment.height > 0 && contentControlFragmentVisibleRect(fragment))
    .sort((left, right) => left.page - right.page || left.y - right.y || left.x - right.x)
  if (result[0]) result[0].first = true
  if (result.at(-1)) result.at(-1)!.last = true
  return result
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
    if (control.blockRange) {
      try {
        const selection = documentContentControlBlocksSelection(model, control.blockRange), fragments = fragmentsForBlocks(model, control, layout)
        if (fragments.length) result.push({ control, target, selection, fragments })
      } catch (error) {
        if (!(error instanceof TypeError || error instanceof RangeError)) throw error
      }
      continue
    }
    if (control.tableRowRange) {
      try {
        const rows = resolveDocumentContentControlTableRows(model, control.tableRowRange)
        const selection = documentContentControlTableRowSelection(model, control.tableRowRange), fragments = fragmentsForRows(control, rows, layout)
        if (fragments.length) result.push({ control, target, selection, fragments })
      } catch (error) {
        if (!(error instanceof TypeError || error instanceof RangeError)) throw error
      }
      continue
    }
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
  return result.sort((left, right) => {
    const leftContains = documentContentControlContains(model, left.control, right.control), rightContains = documentContentControlContains(model, right.control, left.control)
    if (leftContains !== rightContains) return leftContains ? -1 : 1
    return (right.control.blockRange?.blockIds.length ?? 0) - (left.control.blockRange?.blockIds.length ?? 0)
      || (right.control.tableRowRange?.rowIds.length ?? 0) - (left.control.tableRowRange?.rowIds.length ?? 0)
      || (right.control.endOffset - right.control.startOffset) - (left.control.endOffset - left.control.startOffset)
      || depth(left.control) - depth(right.control) || left.control.id.localeCompare(right.control.id)
  })
}

export function hitTestDocumentContentControl(geometries: readonly DocumentContentControlGeometry[], x: number, y: number): DocumentContentControlGeometry | undefined {
  return [...geometries].reverse().find(geometry => geometry.fragments.some(fragment => {
    const rect = contentControlFragmentVisibleRect(fragment)
    return rect && x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height
  }))
}

export function contentControlSelectionAtFragment(geometry: DocumentContentControlGeometry, fragment: LayoutContentControlFragment): DocumentSelection {
  const selection = cloneDocumentSelection(geometry.selection)
  selection.anchor.pageAffinity = fragment.page; selection.focus.pageAffinity = fragment.page
  if (selection.tableSelection) selection.tableSelection.pageAffinity = fragment.page
  return selection
}
