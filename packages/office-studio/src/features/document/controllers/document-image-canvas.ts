import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import type { DocumentSelection, DocumentImage } from '../model/document-model'
import { documentImageEntries, type DocumentImagePatch } from '../images/document-image-model'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { sameDocumentInputLocation, type DocumentInputSession } from '../model/document-input-session'
import type { DocumentImageCommand } from '../commands/document-image-commands'
import type { DocumentLayout, LayoutImage } from '../layout/document-layout-types'
import { documentImageFrameCorners, documentImageRotationHandle, type DocumentImageFrame } from '../layout/document-image-layout'
import { documentImageBatchBounds, transformDocumentImageBatch, type DocumentImageBatchMember, type DocumentImageBatchTransform, type DocumentImageBatchUpdate } from '../images/document-image-batch'

export interface DocumentImageCanvasTarget { source: DocumentModel; imageId: string; imageIds?: string[]; selection: DocumentSelection; host?: DocumentAuxiliaryBlockHost; inputSession?: DocumentInputSession }
export interface DocumentImageCanvasCommandRequest { target: DocumentImageCanvasTarget; command: DocumentImageCommand }
export interface DocumentImageCanvasOperationRequest { target: DocumentImageCanvasTarget; mode: 'edit' | 'crop' | 'replace' }
export interface DocumentImageGesture { target: DocumentImageCanvasTarget; image: DocumentImage; rectangle: LayoutImage; start: { x: number; y: number }; pageOrigin: { x: number; y: number }; corner?: 'nw' | 'ne' | 'sw' | 'se'; rotate?: boolean; pointerId: number; zoom: number }
export interface DocumentImageGestureResult { patch: DocumentImagePatch; rectangle: DocumentImageFrame }
export interface DocumentImageBatchGesture { target: DocumentImageCanvasTarget; members: DocumentImageBatchMember[]; rectangle: DocumentImageFrame; start: { x: number; y: number }; corner?: DocumentImageGesture['corner']; rotate?: boolean; pointerId: number; zoom: number }
export interface DocumentImageBatchGestureResult { transform: DocumentImageBatchTransform; updates: DocumentImageBatchUpdate[]; rectangle: DocumentImageFrame }

export function documentImageCanvasCommandIds(command: DocumentImageCommand): string[] {
  if (command.id === 'doc.command.batch-images') return command.commands.map(item => item.imageId)
  if (command.id === 'doc.command.arrange-images') return [...command.imageIds]
  if (command.id === 'doc.command.align-images' || command.id === 'doc.command.transform-images') return command.members.map(member => member.imageId)
  return [command.imageId]
}
export function documentImageCanvasCommandMatchesTarget(request: DocumentImageCanvasCommandRequest): boolean {
  const ids = request.target.imageIds ?? [request.target.imageId], commandIds = documentImageCanvasCommandIds(request.command)
  return ids.length > 0 && ids.includes(request.target.imageId) && new Set(ids).size === ids.length
    && commandIds.length === ids.length && new Set(commandIds).size === commandIds.length && commandIds.every(id => ids.includes(id))
}

export function captureDocumentImageCanvasTarget(source: DocumentModel, imageId: string, selection: DocumentSelection,
  host?: DocumentAuxiliaryBlockHost, inputSession?: DocumentInputSession, imageIds?: readonly string[]): DocumentImageCanvasTarget {
  return { source, imageId, ...(imageIds ? { imageIds: [...imageIds] } : {}), selection: cloneModel(selection), host: cloneModel(host), inputSession: cloneModel(inputSession) }
}
export function documentImageCanvasTargetCurrent(target: DocumentImageCanvasTarget, current: {
  source: DocumentModel; model: DocumentModel; host?: DocumentAuxiliaryBlockHost; selection: DocumentSelection; imageIds?: readonly string[]; session?: DocumentInputSession; requireSession?: boolean; disabled?: boolean
}): boolean {
  const ids = target.imageIds ?? [target.imageId], entries = documentImageEntries(current.model)
  if (current.disabled || target.source !== current.source || !sameDocumentBlockHost(target.host, current.host)
    || !ids.length || ids.length > 1000 || new Set(ids).size !== ids.length || !ids.includes(target.imageId) || ids.some(id => !entries.some(entry => entry.imageId === id))
    || current.imageIds && JSON.stringify(ids) !== JSON.stringify(current.imageIds)
    || JSON.stringify(target.selection) !== JSON.stringify(current.selection)) return false
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
export function documentImageAtRotationHandle(rectangle: DocumentImageFrame, x: number, y: number, zoom: number): boolean {
  const handle = documentImageRotationHandle(rectangle, zoom), tolerance = 7 / Math.max(0.1, zoom / 100)
  return Math.hypot(x - handle.x, y - handle.y) <= tolerance
}
function rotationDelta(rectangle: DocumentImageFrame, start: { x: number; y: number }, x: number, y: number, snap: boolean): number {
  const centerX = rectangle.x + rectangle.width / 2, centerY = rectangle.y + rectangle.height / 2
  let angle = (Math.atan2(y - centerY, x - centerX) - Math.atan2(start.y - centerY, start.x - centerX)) * 180 / Math.PI
  angle = (angle + 540) % 360 - 180
  return snap ? Math.round(angle / 15) * 15 : angle
}
export function documentImageBatchGestureResult(gesture: DocumentImageBatchGesture, x: number, y: number, snap = false): DocumentImageBatchGestureResult {
  const old = gesture.rectangle
  let transform: DocumentImageBatchTransform
  if (gesture.rotate) transform = { angle: rotationDelta(old, gesture.start, x, y, snap) }
  else if (gesture.corner) {
    const origin = { x: old.x + (gesture.corner.endsWith('w') ? old.width : 0), y: old.y + (gesture.corner.startsWith('n') ? old.height : 0) }
    const startX = gesture.start.x - origin.x, startY = gesture.start.y - origin.y, lengthSquared = startX ** 2 + startY ** 2
    const minimum = Math.max(...gesture.members.map(member => 0.1 / Math.min(member.image.width, member.image.height)))
    const maximum = Math.min(...gesture.members.map(member => 100_000 / Math.max(member.image.width, member.image.height)))
    const scale = Math.max(minimum, Math.min(maximum, lengthSquared ? ((x - origin.x) * startX + (y - origin.y) * startY) / lengthSquared : 1))
    transform = { scale, origin }
  } else transform = { x: x - gesture.start.x, y: y - gesture.start.y }
  const updates = transformDocumentImageBatch(gesture.members, transform)
  const rectangle = updates.length ? documentImageBatchBounds(updates.map((update, index) => ({ ...gesture.members[index]!, rectangle: { ...gesture.members[index]!.rectangle, ...update.rectangle } }))) : old
  return { transform, updates, rectangle }
}
/** The pointer preview never changes the model; pointer-up commits one object command. */
export function documentImageGestureResult(gesture: DocumentImageGesture, x: number, y: number, unlockRatio = false, snapRotation = false): DocumentImageGestureResult {
  const dx = x - gesture.start.x, dy = y - gesture.start.y, old = gesture.rectangle
  if (gesture.rotate) {
    const angle = ((gesture.image.angle ?? 0) + rotationDelta(old, gesture.start, x, y, snapRotation) + 360) % 360
    const radians = angle * Math.PI / 180, paintWidth = old.paintWidth ?? old.width, paintHeight = old.paintHeight ?? old.height
    const width = Math.abs(Math.cos(radians)) * paintWidth + Math.abs(Math.sin(radians)) * paintHeight
    const height = Math.abs(Math.sin(radians)) * paintWidth + Math.abs(Math.cos(radians)) * paintHeight
    const rectangle: DocumentImageFrame = { ...old, x: old.x + (old.width - width) / 2, y: old.y + (old.height - height) / 2, width, height, paintWidth, paintHeight, angle }
    const patch: DocumentImagePatch = { angle }
    if ((gesture.image.wrappingStyle ?? 'inline') !== 'inline') {
      patch.positionH = { relativeFrom: 0, posOffset: rectangle.x - gesture.pageOrigin.x }
      patch.positionV = { relativeFrom: 0, posOffset: rectangle.y - gesture.pageOrigin.y }
    } else { rectangle.x = old.x; rectangle.y = old.y }
    return { patch, rectangle }
  }
  if (!gesture.corner) return { rectangle: { ...old, x: old.x + dx, y: old.y + dy }, patch: {
    wrappingStyle: gesture.image.wrappingStyle === 'inline' || !gesture.image.wrappingStyle ? 'inFrontOfText' : gesture.image.wrappingStyle,
    ...((gesture.image.wrappingStyle ?? 'inline') === 'inline' ? { width: old.paintWidth ?? old.width, height: old.paintHeight ?? old.height } : {}),
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
