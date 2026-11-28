import type { OfficeStudio } from './office-studio'
import { DOCUMENT_HEADER_FOOTER_COMMANDS, DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS, resolveDocumentHeaderFooterRequest } from '../features/document/controllers/document-header-footer-editor'
import { captureDocumentHeaderFooterEditorTarget, requestDocumentHeaderFooterEditor } from './document-header-footer-editors'

/** Fixed desktop commands and local aliases share the real UI owner and content-history boundary. */
export function registerDocumentHeaderFooterCommands(api: OfficeStudio): void {
  for (const commandId of [...Object.values(DOCUMENT_HEADER_FOOTER_COMMANDS), ...Object.values(DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS)]) {
    api.registerCommand<unknown, boolean>(commandId, raw => requestDocumentHeaderFooterEditor(api,
      captureDocumentHeaderFooterEditorTarget(api, resolveDocumentHeaderFooterRequest(commandId, raw))))
  }
}
