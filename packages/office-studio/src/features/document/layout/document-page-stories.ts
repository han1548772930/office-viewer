import type { DocumentModel, DocumentPageField } from '../../../domain/model'
import { sliceRuns } from '../model/document-model'
import { renderSectionStory, resolveSectionHeaderFooter } from '../services/section-header-footer'
import type { EffectiveDocumentSectionSetup } from '../services/document-settings'
import { documentStoryDisplay } from '../services/story-rendering'
import type { DocumentTextLayoutDisplay, LayoutMeasurer, LayoutPage, LayoutPageHeaderFooter, LayoutStoryLine } from './document-layout-types'
import { paragraphLineHeight, positionLine, wrapRuns } from './document-text-layout'
import { layoutDocumentAuxiliaryBlocks } from './document-auxiliary-layout'

export function layoutDocumentPageStories(model: DocumentModel, pages: readonly LayoutPage[], settings: ReadonlyMap<string, EffectiveDocumentSectionSetup>, measurer: LayoutMeasurer,
  noteDisplay?: DocumentTextLayoutDisplay): LayoutPageHeaderFooter[] {
  return pages.map(page => {
    const section = resolveSectionHeaderFooter(model, page.sectionId), setup = settings.get(page.sectionId)!
    const firstPage = section.differentFirstPage && page.sectionPageIndex === 0
    const evenPage = section.differentOddEven && page.pageNumber % 2 === 0
    const headerStory = firstPage ? 'firstPageHeader' : evenPage ? 'evenPageHeader' : 'header'
    const footerStory = firstPage ? 'firstPageFooter' : evenPage ? 'evenPageFooter' : 'footer'
    const headerTop = page.y + setup.marginHeader, footerBottom = page.y + page.height - setup.marginFooter
    const base = { page: page.index, sectionId: page.sectionId, pageNumber: page.pageNumber, pageNumberStart: page.pageNumberStart,
      isParityBlank: page.isParityBlank, headerStory, footerStory, x: page.content.x, y: page.y + 6, width: page.content.width,
      height: Math.max(0, page.content.y - page.y - 12), headerTop, footerBottom }
    if (page.isParityBlank) return { ...base, header: '', footer: '', headerLines: [], footerLines: [], headerLayoutLines: [], footerLayoutLines: [] }
    const storyLines = (story: DocumentPageField['story'], top: number, bottom = false): LayoutStoryLine[] => {
      const display = documentStoryDisplay(section, story, page.pageNumber, pages.length, model)
      const output: LayoutStoryLine[] = []
      let cursor = 0
      for (const paragraph of display.paragraphs) {
        const style = paragraph.style, runs = sliceRuns(display.runs, paragraph.start, paragraph.end)
        if (!runs.length) runs.push(paragraph.emptyRun)
        const left = style.indentStart ?? style.indent * 24, right = style.indentEnd ?? 0
        const first = (style.indentFirstLine ?? 0) - (style.hanging ?? 0)
        const available = (index: number) => Math.max(1, page.content.width - left - right - (index === 0 ? first : 0))
        cursor += style.spaceBefore
        const wrapped = wrapRuns(runs, available, measurer, style.wordWrap === true, 12, undefined, { references: noteDisplay?.references })
        wrapped.forEach((raw, index) => {
          const height = paragraphLineHeight(raw.fontHeight, style)
          const line = positionLine(raw, { blockId: '' }, page.content.x + left + (index === 0 ? first : 0), cursor,
            available(index), height, page.index, style.align, index < wrapped.length - 1)
          line.glyphs = line.glyphs.map(glyph => {
            const range = display.inputRange(paragraph.start + glyph.offset, paragraph.start + glyph.end)
            return { ...glyph, offset: range.start, end: range.end }
          })
          const range = raw.glyphs.length ? display.inputRange(paragraph.start + raw.start, paragraph.start + raw.end)
            : raw.start > 0 ? display.inputRange(paragraph.start + raw.start, paragraph.start + raw.start)
              : { start: paragraph.bufferStart, end: paragraph.start === paragraph.end ? paragraph.bufferEnd : paragraph.bufferStart }
          output.push({ ...line, start: range.start, end: range.end })
          cursor += height
        })
        cursor += style.spaceAfter
      }
      const shift = bottom ? top - cursor : top
      return output.map(line => ({ ...line, y: line.y + shift, baseline: line.baseline + shift,
        glyphs: line.glyphs.map(glyph => ({ ...glyph, y: glyph.y + shift })) }))
    }
    const headerBlockLayout = section.storyContent?.[headerStory]?.blocks !== undefined ? layoutDocumentAuxiliaryBlocks(model, { kind: 'story', sectionId: page.sectionId, story: headerStory },
      { x: page.content.x, y: headerTop, width: page.content.width, height: Math.max(1, page.height - setup.marginHeader - setup.marginFooter) }, page.index, page.pageNumber, pages.length, measurer, false, noteDisplay) : undefined
    const footerBlockLayout = section.storyContent?.[footerStory]?.blocks !== undefined ? layoutDocumentAuxiliaryBlocks(model, { kind: 'story', sectionId: page.sectionId, story: footerStory },
      { x: page.content.x, y: page.y + setup.marginHeader, width: page.content.width, height: Math.max(1, page.height - setup.marginHeader - setup.marginFooter) }, page.index, page.pageNumber, pages.length, measurer, true, noteDisplay) : undefined
    const headerLayoutLines = headerBlockLayout?.lines ?? storyLines(headerStory, headerTop), footerLayoutLines = footerBlockLayout?.lines ?? storyLines(footerStory, footerBottom, true)
    return { ...base, header: renderSectionStory(section, headerStory, page.pageNumber, pages.length), footer: renderSectionStory(section, footerStory, page.pageNumber, pages.length),
      headerLines: headerLayoutLines.map(line => line.glyphs.map(glyph => glyph.text).join('')),
      footerLines: footerLayoutLines.map(line => line.glyphs.map(glyph => glyph.text).join('')), headerLayoutLines, footerLayoutLines,
      ...(headerBlockLayout ? { headerBlockLayout } : {}), ...(footerBlockLayout ? { footerBlockLayout } : {}) }
  })
}
