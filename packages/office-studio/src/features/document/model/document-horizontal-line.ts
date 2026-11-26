import type { DocumentBlock } from '../../../domain/model'

export interface DocumentHorizontalLineBorder {
  width: number
  padding: number
  color: string
  style: 'solid' | 'dot' | 'dash'
}

export const DOCUMENT_HORIZONTAL_LINE_BORDER: DocumentHorizontalLineBorder = { width: 1, padding: 5, color: '#CDD0D8', style: 'solid' }

export function documentHorizontalLineBorder(block: DocumentBlock): DocumentHorizontalLineBorder {
  return (block.metadata.documentDividerBorder as DocumentHorizontalLineBorder | undefined) ?? DOCUMENT_HORIZONTAL_LINE_BORDER
}
