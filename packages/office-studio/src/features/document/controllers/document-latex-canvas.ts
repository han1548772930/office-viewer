import { selectedLatexFormula } from '../latex/latex-formula'
import { documentModelForBlockHost } from '../services/document-block-host'
import { documentStructuredCanvasTargetCurrent, type DocumentStructuredCanvasTarget } from './document-structured-canvas'

export interface DocumentLatexCanvasEditRequest extends DocumentStructuredCanvasTarget { rangeId: string }

export function documentLatexCanvasRequestCurrent(request: DocumentLatexCanvasEditRequest,
  current: Parameters<typeof documentStructuredCanvasTargetCurrent>[1]): boolean {
  if (!documentStructuredCanvasTargetCurrent(request, current)) return false
  try {
    const model = current.host ? documentModelForBlockHost(current.source, current.host) : current.source
    return selectedLatexFormula(model, current.selection)?.rangeId === request.rangeId
  } catch (error) {
    if (error instanceof RangeError || error instanceof TypeError) return false
    throw error
  }
}
