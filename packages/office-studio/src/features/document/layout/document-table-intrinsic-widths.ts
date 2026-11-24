import type { DocumentBlock } from '../../../domain/model'
import { documentTableCellBlocks } from '../model/document-table-continuations'
import { imageForBlock, tableForBlock, type DocumentRun } from '../model/document-model'
import { documentTableGrid } from '../model/document-table-grid'
import { documentTableCellMargin, documentTableColumnWidths } from '../model/document-table-style'
import type { ParagraphListMarker } from '../model/document-lists'
import type { DocumentTextLayoutDisplay, LayoutMeasurer } from './document-layout-types'
import { paragraphGeometry } from './document-paragraph-layout'
import { wrapRuns } from './document-text-layout'

export interface DocumentTableIntrinsicWidth { minimum: number; preferred: number }

/** Uses the same glyph/inline-object metrics as paragraph layout; no model width is rewritten. */
function paragraphWidth(block: DocumentBlock, markers: Map<string, ParagraphListMarker>, measurer: LayoutMeasurer, display?: DocumentTextLayoutDisplay): DocumentTableIntrinsicWidth {
  const nested = tableForBlock(block)
  if (nested) {
    const columns = intrinsicDocumentTableColumns(block, markers, measurer, display)
    const offset = Math.max(0, nested.style?.indent ?? 0)
    if (nested.style?.layout !== 'auto' && nested.columnWidths?.length) {
      const fixed = nested.columnWidths.reduce((sum, width) => sum + width, 0) + offset
      return { minimum: fixed, preferred: fixed }
    }
    return { minimum: columns.reduce((sum, column) => sum + column.minimum, offset),
      preferred: Math.max(nested.style?.width ?? 0, columns.reduce((sum, column) => sum + column.preferred, offset)) }
  }
  const image = imageForBlock(block)
  if (image) { const width = Math.max(1, image.width || 320); return { minimum: width, preferred: width } }
  if (['divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)) return { minimum: 1, preferred: 1 }
  const geometry = paragraphGeometry(block, markers.get(block.id), 1_000_000, measurer)
  const lines = wrapRuns(block.runs as DocumentRun[], Number.POSITIVE_INFINITY, measurer, block.style.wordWrap === true, undefined, undefined,
    { references: display?.references, leading: display?.leadingLabels?.get(block.id) })
  let minimum = 1, preferred = 1
  for (const [index, line] of lines.entries()) {
    const indent = Math.max(0, index === 0 ? geometry.firstIndent : geometry.indent) + Math.max(0, geometry.rightIndent)
    preferred = Math.max(preferred, line.width + indent)
    // The paragraph kernel can break a long word at grapheme boundaries.
    // A formula/image glyph remains indivisible, matching that same fallback.
    for (const glyph of line.glyphs) minimum = Math.max(minimum, glyph.width + indent)
  }
  return { minimum, preferred: Math.max(minimum, preferred) }
}

export function intrinsicDocumentTableColumns(block: DocumentBlock, markers: Map<string, ParagraphListMarker>, measurer: LayoutMeasurer, display?: DocumentTextLayoutDisplay): DocumentTableIntrinsicWidth[] {
  const table = tableForBlock(block)
  if (!table) throw new RangeError('表格模型不存在')
  const grid = documentTableGrid(table), columns = Array.from({ length: grid.columnCount }, () => ({ minimum: 1, preferred: 1 }))
  const constraints = grid.anchors.map(owner => {
    const margin = documentTableCellMargin(table, owner.cell)
    let minimum = 1, preferred = 1
    for (const child of documentTableCellBlocks(owner.cell)) {
      const width = paragraphWidth(child, markers, measurer, display)
      minimum = Math.max(minimum, width.minimum); preferred = Math.max(preferred, width.preferred)
    }
    return { owner, minimum: minimum + margin.left + margin.right, preferred: preferred + margin.left + margin.right }
  }).sort((left, right) => left.owner.colSpan - right.owner.colSpan)
  for (const constraint of constraints) {
    const start = constraint.owner.column, end = start + constraint.owner.colSpan
    for (const field of ['minimum', 'preferred'] as const) {
      let current = 0
      for (let column = start; column < end; column++) current += columns[column]![field]
      const extra = Math.max(0, constraint[field] - current) / constraint.owner.colSpan
      for (let column = start; column < end; column++) columns[column]![field] += extra
    }
    for (let column = start; column < end; column++) columns[column]!.preferred = Math.max(columns[column]!.minimum, columns[column]!.preferred)
  }
  return columns
}

/** Auto layout honors content and colSpan constraints within the requested/page width. */
export function measureDocumentTableColumnWidths(block: DocumentBlock, available: number, markers: Map<string, ParagraphListMarker>, measurer: LayoutMeasurer, display?: DocumentTextLayoutDisplay): number[] {
  const table = tableForBlock(block)
  if (!table) throw new RangeError('表格模型不存在')
  const grid = documentTableGrid(table), fixed = documentTableColumnWidths(table, grid.columnCount, available)
  if (table.style?.layout !== 'auto') return fixed
  const columns = intrinsicDocumentTableColumns(block, markers, measurer, display)
  const minimum = columns.reduce((sum, column) => sum + column.minimum, 0), preferred = columns.reduce((sum, column) => sum + column.preferred, 0)
  const requested = table.style.width ?? preferred
  if (!Number.isFinite(requested) || requested <= 0) throw new RangeError('表格宽度必须大于零')
  const target = Math.max(minimum, Math.min(requested, Math.max(1, available - Math.max(0, table.style.indent ?? 0))))
  if (target <= preferred) {
    const room = preferred - minimum, fraction = room > 0 ? (target - minimum) / room : 0
    return columns.map(column => column.minimum + (column.preferred - column.minimum) * fraction)
  }
  const totalWeight = fixed.reduce((sum, width) => sum + width, 0)
  return columns.map((column, index) => column.preferred + (target - preferred) * fixed[index]! / totalWeight)
}
