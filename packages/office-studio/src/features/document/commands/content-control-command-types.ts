import type { DocumentModel } from '../../../domain/model'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import type { DocumentInputSession } from '../model/document-input-session'
import type { DocumentSelection } from '../model/document-model'

/** UI intentions use the same registered commands as the public content-control facade. */
export type DocumentContentControlValueCommand =
  | { id: 'doc.command.set-content-control-checked'; params: { controlId: string; checked: boolean } }
  | { id: 'doc.command.set-content-control-value'; params: { controlId: string; value: string } }
  | { id: 'doc.command.select-content-control-item'; params: { controlId: string; value: string } }
  | { id: 'doc.command.reset-content-control-placeholder'; params: { controlId: string } }

export type DocumentContentControlRepeatingItemCommand =
  | { id: 'doc.command.duplicate-repeating-section-item'; params: { controlId: string } }
  | { id: 'doc.command.delete-repeating-section-item'; params: { controlId: string } }

/** Transient popup selection mode; it never belongs to the persisted control model. */
export type DocumentContentControlPopupMode = 'value' | 'properties'

export interface DocumentContentControlSelectionRequest {
  controlId: string
  selection: DocumentSelection
  /** The full source document, never the transient story/note view. */
  source: DocumentModel
  /** undefined explicitly addresses the body; a story/note must carry its address. */
  host?: DocumentAuxiliaryBlockHost
  inputSession?: DocumentInputSession
  mode?: DocumentContentControlPopupMode
}

export interface DocumentContentControlCommandRequest extends DocumentContentControlSelectionRequest {
  command: DocumentContentControlValueCommand
}
