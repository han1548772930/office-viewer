import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentImageEntries } from '../images/document-image-model'

/** Text replacement and object commands keep the picture control's resource mirror in one candidate. */
export function synchronizeDocumentPictureControlImages(before: DocumentModel, candidate: DocumentModel): void {
  const controls = (candidate.contentControls ?? []).filter(control => control.type === 'picture'
    && typeof control.metadata?.pictureBlockId === 'string')
  if (!controls.length) return
  const previous = new Map(documentImageEntries(before).map(entry => [entry.imageId, entry.image]))
  const images = new Map(documentImageEntries(candidate).map(entry => [entry.imageId, entry.image]))
  for (const control of controls) {
    const id = control.metadata?.pictureBlockId
    if (typeof id !== 'string') continue
    const image = images.get(id)
    if (JSON.stringify(previous.get(id)) === JSON.stringify(image)) continue
    if (image) control.metadata = { ...control.metadata, pictureImage: cloneModel(image) }
    else {
      const { pictureBlockId: _id, pictureImage: _image, ...metadata } = control.metadata ?? {}
      control.metadata = metadata
      control.value = ''
      control.showingPlaceholder = true
    }
  }
}
