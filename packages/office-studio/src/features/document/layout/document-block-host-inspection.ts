import type { DocumentModel } from '../../../domain/model'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentModelForBlockHost, resolveDocumentBlockHost } from '../services/document-block-host'
import { documentEffectiveSectionSetup } from '../services/document-settings'
import { appendDocumentBlocks } from './document-block-layout'
import { createDocumentFlowResult, documentFlowContentFragments } from './document-note-layout-common'
import { arrangeDocumentPages } from './document-page-arrangement'
import { DocumentSectionFlow } from './document-section-flow'
import type { DocumentLayout, LayoutMeasurer } from './document-layout-types'

/**
 * Explicit pages for orphan notes and stored separator stories. The same block
 * kernel lays out their real IDs, but all content remains auxiliary: these
 * pages do not allocate a body reference, page number or header/footer story.
 */
export function layoutDocumentBlockHostInspection(model: DocumentModel, host: DocumentAuxiliaryBlockHost,
  measurer: LayoutMeasurer): DocumentLayout {
  const view = resolveDocumentBlockHost(model, host), scoped = documentModelForBlockHost(model, host)
  const setup = documentEffectiveSectionSetup(model, view.sectionId)
  const flow = new DocumentSectionFlow({ ...scoped, page: { ...setup.page },
    sectionSetup: { columnCount: 1, columnGap: 0, columnSeparator: 'none' } })
  const result = createDocumentFlowResult(flow)
  appendDocumentBlocks(scoped.blocks, flow, result, measurer)
  result.auxiliaryContent = documentFlowContentFragments(result, flow, host, view.sectionId, view.sourceHost)
  result.lines = []; result.cells = []; result.images = []; result.dividers = []; result.tableFragments = []
  result.headings = []; result.pageHeaderFooters = []; result.noteDecorations = []
  result.inspectionHost = { ...host }
  arrangeDocumentPages(result)
  return result
}
