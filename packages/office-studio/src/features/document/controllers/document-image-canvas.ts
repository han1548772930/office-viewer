import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import type { DocumentSelection, DocumentImage } from '../model/document-model'
import { documentImageEntries, type DocumentImagePatch } from '../images/document-image-model'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { sameDocumentInputLocation, type DocumentInputSession } from '../model/document-input-session'
import type { DocumentImageObjectCommand } from '../commands/document-image-commands'
import type { DocumentLayout, LayoutImage } from '../layout/document-layout-types'
import { documentImageFrameCorners, type DocumentImageFrame } from '../layout/document-image-layout'

export interface DocumentImageCanvasTarget { source: DocumentModel; imageId: string; selection: DocumentSelection; host?: DocumentAuxiliaryBlockHost; inputSession?: DocumentInputSession }
export interface DocumentImageCanvasCommandRequest { target: DocumentImageCanvasTarget; command: DocumentImageObjectCommand }
export interface DocumentImageCanvasOperationRequest { target: DocumentImageCanvasTarget; mode: 'edit' | 'crop' | 'replace' }
export interface DocumentImageGesture { target: DocumentImageCanvasTarget; image: DocumentImage; rectangle: LayoutImage; start: { x: number; y: number }; pageOrigin: { x: number; y: number }; corner?: 'nw' | 'ne' | 'sw' | 'se'; pointerId: number; zoom: number }
export interface DocumentImageGestureResult { patch: DocumentImagePatch; rectangle: DocumentImageFrame }

export function captureDocumentImageCanvasTarget(source: DocumentModel, imageId: string, selection: DocumentSelection,
  host?: DocumentAuxiliaryBlockHost, inputSession?: DocumentInputSession): DocumentImageCanvasTarget {
  return { source, imageId, selection: cloneModel(selection), host: cloneModel(host), inputSession: cloneModel(inputSession) }
}
export function documentImageCanvasTargetCurrent(target: DocumentImageCanvasTarget, current: {
  source: DocumentModel; model: DocumentModel; host?: DocumentAuxiliaryBlockHost; selection: DocumentSelection; session?: DocumentInputSession; requireSession?: boolean; disabled?: boolean
}): boolean {
  if (current.disabled || target.source !== current.source || !sameDocumentBlockHost(target.host, current.host)
    || JSON.stringify(target.selection) !== JSON.stringify(current.selection) || !documentImageEntries(current.model).some(entry => entry.imageId === target.imageId)) return false
  const input = target.inputSession, latest = current.session
  return !current.requireSession && !input || !!input?.valid && !!latest?.valid && input.modelRevision === latest.modelRevision
    && input.generation === latest.generation && sameDocumentInputLocation(input.location, latest.location)
}
export function documentImageAtPoint(model: DocumentModel, layout: DocumentLayout, x: number, y: number, includeBehindText = false): LayoutImage | undefined {
  const ids = new Set(documentImageEntries(model).map(entry => entry.imageId))
  const overText = !includeBehindText && layout.lines.some(line => x >= line.x && x <= line.x + line.width && y >= line.y && y <= line.y + line.height)
  const candidates = layout.images.filter(image => image.objectImage && ids.has(image.blockId) && x >= image.x && x <= image.x + image.width
    && y >= image.y && y <= image.y + image.height && (!overText || image.wrappingStyle !== 'behindText')
    && (!image.clip || x >= image.clip.x && x <= image.clip.x + image.clip.width && y >= image.clip.y && y <= image.clip.y + image.clip.height))
    .sort((a, b) => Number(b.wrappingStyle === 'inFrontOfText') - Number(a.wrappingStyle === 'inFrontOfText') || (b.layer ?? 0) - (a.layer ?? 0))
  return candidates.find(image => {
    const angle = -(image.angle ?? 0) * Math.PI / 180, dx = x - image.x - image.width / 2, dy = y - image.y - image.height / 2
    const localX = dx * Math.cos(angle) - dy * Math.sin(angle), localY = dx * Math.sin(angle) + dy * Math.cos(angle)
    return Math.abs(localX) <= (image.paintWidth ?? image.width) / 2 && Math.abs(localY) <= (image.paintHeight ?? image.height) / 2
  })
}
export function documentImageResizeCorner(rectangle: DocumentImageFrame, x: number, y: number, zoom: number): DocumentImageGesture['corner'] {
  const tolerance = 7 / Math.max(0.1, zoom / 100)
  for (const point of documentImageFrameCorners(rectangle)) {
    if (Math.abs(x - point.x) <= tolerance && Math.abs(y - point.y) <= tolerance) return point.corner
  }
}
/** The pointer preview never changes the model; pointer-up commits one object command. */
export function documentImageGestureResult(gesture: DocumentImageGesture, x: number, y: number, unlockRatio = false): DocumentImageGestureResult {
  const dx = x - gesture.start.x, dy = y - gesture.start.y, old = gesture.rectangle
  if (!gesture.corner) return { rectangle: { ...old, x: old.x + dx, y: old.y + dy }, patch: {
    wrappingStyle: gesture.image.wrappingStyle === 'inline' || !gesture.image.wrappingStyle ? 'inFrontOfText' : gesture.image.wrappingStyle,
    positionH: { relativeFrom: 0, posOffset: old.x + dx - gesture.pageOrigin.x },
    positionV: { relativeFrom: 0, posOffset: old.y + dy - gesture.pageOrigin.y },
  } }
  const angle = (old.angle ?? 0) * Math.PI / 180, cosine = Math.cos(angle), sine = Math.sin(angle)
  const localDx = dx * cosine + dy * sine, localDy = -dx * sine + dy * cosine
  const signX = gesture.corner.endsWith('w') ? -1 : 1, signY = gesture.corner.startsWith('n') ? -1 : 1
  const oldWidth = old.paintWidth ?? old.width, oldHeight = old.paintHeight ?? old.height
  let width = Math.max(0.1 * oldWidth / gesture.image.width, oldWidth + signX * localDx)
  let height = Math.max(0.1 * oldHeight / gesture.image.height, oldHeight + signY * localDy)
  if (gesture.image.keepAspectRatio !== false && !unlockRatio) {
    const ratio = Math.max(0.1 / Math.min(gesture.image.width, gesture.image.height),
      Math.min(100_000 / Math.max(gesture.image.width, gesture.image.height), Math.abs(localDx / oldWidth) >= Math.abs(localDy / oldHeight) ? width / oldWidth : height / oldHeight))
    width = oldWidth * ratio; height = oldHeight * ratio
  }
  width = Math.min(width, oldWidth * 100_000 / gesture.image.width); height = Math.min(height, oldHeight * 100_000 / gesture.image.height)
  const centerDx = signX * (width - oldWidth) / 2, centerDy = signY * (height - oldHeight) / 2
  const frameWidth = Math.abs(cosine) * width + Math.abs(sine) * height, frameHeight = Math.abs(sine) * width + Math.abs(cosine) * height
  const rectangle: DocumentImageFrame = { x: old.x + old.width / 2 + centerDx * cosine - centerDy * sine - frameWidth / 2,
    y: old.y + old.height / 2 + centerDx * sine + centerDy * cosine - frameHeight / 2, width: frameWidth, height: frameHeight, paintWidth: width, paintHeight: height, angle: old.angle }
  const patch: DocumentImagePatch = { width: gesture.image.width * width / oldWidth, height: gesture.image.height * height / oldHeight }
  if ((gesture.image.wrappingStyle ?? 'inline') !== 'inline') {
    patch.positionH = { relativeFrom: 0, posOffset: rectangle.x - gesture.pageOrigin.x }
    patch.positionV = { relativeFrom: 0, posOffset: rectangle.y - gesture.pageOrigin.y }
  } else { rectangle.x = old.x; rectangle.y = old.y }
  return { patch, rectangle }
}
