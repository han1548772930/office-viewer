import type { DocumentLayout, LayoutLine, LayoutSection } from './document-layout-types'

interface FlowLine { line: LayoutLine; gapBefore: number }

/** Repositions only a complete line fragment with equal-width columns. */
export function balanceDocumentSection(layout: DocumentLayout, section: LayoutSection): boolean {
  const columns = section.columns
  if (columns.length < 2 || section.hasExplicitColumnBreak || columns.at(-1)?.full
    || columns.some(column => Math.abs(column.width - columns[0]!.width) > 0.01)
    || [...layout.cells, ...layout.images, ...layout.dividers].some(item => item.sectionFragment === section.fragment)) return false
  const lines = layout.lines.filter(line => line.sectionFragment === section.fragment)
  if (lines.length < columns.length) return false
  const previousBottom = columns.map(column => column.y)
  const flow: FlowLine[] = lines.map(line => {
    const column = line.column ?? 0
    const gapBefore = Math.max(0, line.y - previousBottom[column]!)
    previousBottom[column] = line.y + line.height
    return { line, gapBefore }
  })
  const lastLine = lines.at(-1)!, lastColumn = columns[lastLine.column ?? 0]!
  const trailingGap = Math.max(0, lastColumn.y + lastColumn.usedHeight - lastLine.y - lastLine.height)
  let remainingHeight = flow.reduce((sum, item) => sum + item.gapBefore + item.line.height, trailingGap), start = 0
  const assignments: Array<{ items: FlowLine[]; height: number }> = []
  for (let column = 0; column < columns.length; column++) {
    const remaining = columns.length - column, maximum = flow.length - start - (remaining - 1)
    const target = remainingHeight / remaining
    let count = 0, height = 0
    while (count < maximum) {
      const item = flow[start + count]!, next = height + item.gapBefore + item.line.height
      if (remaining > 1 && count > 0 && Math.abs(target - height) <= Math.abs(target - next)) break
      height = next; count++
    }
    if (remaining === 1) height += trailingGap
    if (height > columns[column]!.height + 1e-6) return false
    assignments.push({ items: flow.slice(start, start + count), height })
    remainingHeight -= height; start += count
  }
  assignments.forEach(({ items, height }, index) => {
    const column = columns[index]!
    let y = column.y
    for (const { line, gapBefore } of items) {
      y += gapBefore
      const oldColumn = columns[line.column ?? 0]!, dx = column.x - oldColumn.x, dy = y - line.y
      line.x += dx; line.y += dy; line.baseline += dy; line.column = index
      if (line.markerStyle) line.markerStyle.x += dx
      for (const glyph of line.glyphs) { glyph.x += dx; glyph.y += dy }
      y += line.height
      const heading = layout.headings.find(heading => heading.id === line.target.blockId && line.start === 0)
      if (heading) heading.y = line.y
    }
    column.usedHeight = height; column.full = false
  })
  section.usedHeight = Math.max(...columns.map(column => column.usedHeight))
  return true
}
