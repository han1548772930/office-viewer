import type { DocumentBlock } from '../../../domain/model'
import { imageForBlock, runText, type TextTarget } from '../model/document-model'
import type { ParagraphListMarker } from '../model/document-lists'
import { documentTextFields } from '../model/document-text-fields'
import { appendListMarker, paragraphGeometry } from './document-paragraph-layout'
import { paragraphLineHeight, positionLine, wrapRuns } from './document-text-layout'
import type { DocumentTextLayoutDisplay, LayoutImage, LayoutLine, LayoutMeasurer } from './document-layout-types'

/** The empty input host and its adjacent picture occupy one physical line without losing either block ID. */
export function documentPictureInputAnchors(blocks: readonly DocumentBlock[]): Map<string, DocumentBlock> {
  const result = new Map<string, DocumentBlock>()
  for (let index = 0; index + 1 < blocks.length; index++) {
    const anchor = blocks[index]!, image = blocks[index + 1]!
    if (['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(anchor.type)
      || runText(anchor.runs) !== '' || anchor.runs.some(run => run.metadata?.note || run.metadata?.noteReference || run.metadata?.latexFormula)
      || documentTextFields(anchor).length || image.type !== 'image' || !imageForBlock(image)
      || typeof image.metadata.contentControlId !== 'string' && image.metadata.pictureInputAnchor !== true
      || image.metadata.anchorBlockId !== anchor.id || image.metadata.anchorOffset !== 0) continue
    result.set(image.id, anchor)
  }
  return result
}

/** Keep an actual zero-offset caret/hit target inside the image's existing vertical extent. */
export function layoutDocumentPictureInputAnchor(anchor: DocumentBlock, image: LayoutImage, target: TextTarget,
  available: number, left: number, measurer: LayoutMeasurer, images: LayoutImage[], list?: ParagraphListMarker, display?: DocumentTextLayoutDisplay): LayoutLine {
  const geometry = paragraphGeometry(anchor, list, available, measurer)
  const raw = wrapRuns(anchor.runs, geometry.width, measurer, anchor.style.wordWrap === true, undefined, undefined,
    { references: display?.references, leading: display?.leadingLabels?.get(anchor.id) })[0]
  if (!raw) throw new RangeError('图片输入锚点缺少合法的空文字布局。')
  const height = Math.min(image.height, paragraphLineHeight(Math.max(raw.fontHeight, geometry.markerHeight), anchor.style))
  const marker = list?.marker ?? (anchor.type === 'quote' ? '│' : undefined)
  const line = positionLine(raw, target, image.x, image.y + image.height - height, image.width, height, image.page, 'left', false, marker)
  if (image.tableCellId) line.tableCellId = image.tableCellId
  if (image.clip) line.clip = image.clip
  if (image.sectionId !== undefined) line.sectionId = image.sectionId
  if (image.sectionFragment !== undefined) line.sectionFragment = image.sectionFragment
  if (image.column !== undefined) line.column = image.column
  // An empty input bridge has no font glyph to draw beyond a tiny image.
  line.glyphLineHeight = Math.min(line.glyphLineHeight ?? height, height)
  line.pictureInputAnchorImageId = image.blockId
  appendListMarker(line, anchor, list, geometry, left, images)
  if (anchor.type === 'checklist') line.checklist = { blockId: anchor.id, checked: anchor.checklist?.checked === true }
  return line
}
