import type { DocumentModel } from '../../../domain/model'
import { getLatexFormula, getLatexFormulas, type DocumentLatexFormulaTarget } from '../latex/latex-formula'
import { orderedSelection, sameTarget, type DocumentPoint, type DocumentSelection } from '../model/document-model'
import type { DocumentLayout, LayoutGlyph, LayoutLine, LayoutRect } from './document-layout-types'

export interface DocumentLatexHit {
  item: DocumentLatexFormulaTarget
  line: LayoutLine
  glyph: LayoutGlyph
  selection: DocumentSelection
}

function contains(rect: LayoutRect, x: number, y: number): boolean {
  return rect.width > 0 && rect.height > 0 && x >= rect.x && x < rect.x + rect.width && y >= rect.y && y < rect.y + rect.height
}

/** Both halves of an equation hit its complete atom, including persisted errors. */
export function hitTestDocumentLatex(model: DocumentModel, layout: Pick<DocumentLayout, 'lines'>, x: number, y: number): DocumentLatexHit | undefined {
  if (!Number.isFinite(x) || !Number.isFinite(y)) return undefined
  for (const line of layout.lines) {
    if (line.clip && !contains(line.clip, x, y)) continue
    for (const glyph of line.glyphs) {
      if (!getLatexFormula(glyph.run) || !contains(glyph, x, y)) continue
      const item = getLatexFormulas(model).find(item => item.runId === glyph.run.id && sameTarget(item.target, line.target)
        && item.startOffset === glyph.offset && item.endOffset === glyph.end)
      if (!item) continue
      return { item, line, glyph, selection: {
        anchor: { ...item.target, offset: item.startOffset, pageAffinity: line.page },
        focus: { ...item.target, offset: item.endOffset, pageAffinity: line.page },
      } }
    }
  }
  return undefined
}

/** A drag crossing an atom uses the far edge as its anchor. */
export function documentLatexDragAnchor(model: DocumentModel, atom: DocumentSelection, focus: DocumentPoint): DocumentPoint {
  const { start } = orderedSelection(model, { anchor: atom.anchor, focus })
  return sameTarget(start, focus) && start.offset === focus.offset ? atom.focus : atom.anchor
}
