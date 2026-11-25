import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { arrangeDocumentImageOrder, documentImageEntries, documentImageEntryMetadata, normalizeDocumentImage, patchDocumentImage, setDocumentImageEntry, type DocumentImageArrange, type DocumentImagePatch } from '../images/document-image-model'
import { createParagraph, clampPoint, collapsedSelection, findDocumentBlock, sameTarget, type DocumentImage, type DocumentSelection } from '../model/document-model'
import { documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from '../model/content-control-block-range'
import { reconcileDocumentContentControlBlockEdits } from '../model/content-control-block-edits'
import { getDocumentContentControls, validateDocumentContentControlRanges } from '../services/content-controls'
import { assertDocumentPermission, permissionTargetsForChangedBlocks, type DocumentPermissionContext } from '../permissions/document-permissions'
import { assertDocumentTextSelectionEditable } from '../services/document-text-formatting'
import { synchronizeDocumentPictureControlImages } from '../services/document-image-controls'

export type DocumentImageTextMutation = (model: DocumentModel, selection: DocumentSelection, text: string) => DocumentSelection

export type DocumentImageObjectCommand =
  | { id: 'doc.command.update-image'; imageId: string; patch: DocumentImagePatch }
  | { id: 'doc.command.replace-image'; imageId: string; image: DocumentImage }
  | { id: 'doc.command.remove-image'; imageId: string }
  | { id: 'doc.command.reset-image-size'; imageId: string }
  | { id: 'doc.command.arrange-image'; imageId: string; arrange: DocumentImageArrange }
export type DocumentImageCommand = DocumentImageObjectCommand
  | ({ id: 'doc.command.arrange-images'; imageIds: string[] } & ({ arrange: DocumentImageArrange; zOrder?: never } | { arrange?: never; zOrder: number }))
  | { id: 'doc.command.batch-images'; commands: DocumentImageObjectCommand[] }

export function isDocumentImageCommand(command: { id: string }): command is DocumentImageCommand {
  return ['doc.command.update-image', 'doc.command.replace-image', 'doc.command.remove-image', 'doc.command.reset-image-size', 'doc.command.arrange-image', 'doc.command.arrange-images', 'doc.command.batch-images'].includes(command.id)
}

function arrangeImages(candidate: DocumentModel, imageIds: string[], arrangement: DocumentImageArrange | number): void {
  const entries = documentImageEntries(candidate), first = entries.find(entry => entry.imageId === imageIds[0])
  if (!first) throw new RangeError('图片已删除或不属于当前内容区。')
  if (imageIds.some(id => !entries.some(entry => entry.imageId === id && entry.container === first.container))) throw new RangeError('图片排列需要同一容器中的对象。')
  for (const imageId of imageIds) assertDocumentImageEditable(candidate, imageId)
  const peers = entries.filter(entry => entry.container === first.container)
    .sort((left, right) => (left.image.layer ?? 0) - (right.image.layer ?? 0) || first.container.indexOf(left.block) - first.container.indexOf(right.block))
  const order = peers.map(peer => peer.imageId), arranged = arrangeDocumentImageOrder(order, imageIds, arrangement)
  if (arranged.every((id, index) => id === order[index])) return
  const byId = new Map(peers.map(peer => [peer.imageId, peer]))
  for (const [layer, id] of arranged.entries()) {
    const peer = byId.get(id)!
    if (peer.image.layer === layer) continue
    assertDocumentImageEditable(candidate, id)
    setDocumentImageEntry(peer, { ...peer.image, layer })
  }
}
export function assertDocumentImageEditable(model: DocumentModel, imageId: string, removing = false): void {
  const entry = documentImageEntries(model).find(entry => entry.imageId === imageId)
  if (!entry) throw new RangeError('图片已删除或不属于当前内容区。')
  if (entry.inline) assertDocumentTextSelectionEditable(model, {
    anchor: { ...entry.inline.target, offset: entry.inline.startOffset }, focus: { ...entry.inline.target, offset: entry.inline.endOffset } })
  for (const control of getDocumentContentControls(model)) {
    const owns = control.blockId === imageId || control.type === 'picture' && control.metadata?.pictureBlockId === imageId || control.blockRange
      && documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, control.blockRange)).has(entry.block.id)
    if (owns && (['contentLocked', 'sdtContentLocked'].includes(control.lock)
      || removing && ['controlLocked', 'sdtLocked'].includes(control.lock))) throw new Error('图片所属内容控件已锁定。')
  }
}

/** All object changes publish one transaction candidate through the editor history. */
export function executeDocumentImageCommand(model: DocumentModel, selection: DocumentSelection, command: DocumentImageCommand,
  context: DocumentPermissionContext, replaceText: DocumentImageTextMutation): DocumentSelection {
  if (command.id === 'doc.command.batch-images') {
    if (!Array.isArray(command.commands) || !command.commands.length || command.commands.length > 1000
      || command.commands.some(item => !item || !isDocumentImageCommand(item) || ['doc.command.batch-images', 'doc.command.arrange-images'].includes(Reflect.get(item, 'id')))) throw new TypeError('图片批量操作无效。')
    const candidate = cloneModel(model)
    let current = selection
    for (const item of command.commands) current = executeDocumentImageCommand(candidate, current, item, context, replaceText)
    assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedBlocks(model, candidate))
    Object.assign(model, candidate)
    return current
  }
  if (command.id === 'doc.command.arrange-images') {
    if (!Array.isArray(command.imageIds) || !command.imageIds.length || command.imageIds.length > 1000
      || command.zOrder !== undefined && (command.arrange !== undefined || typeof command.zOrder !== 'number')) throw new TypeError('图片排列参数无效。')
    const candidate = cloneModel(model)
    arrangeImages(candidate, command.imageIds, command.zOrder ?? command.arrange!)
    synchronizeDocumentPictureControlImages(model, candidate)
    validateDocumentContentControlRanges(candidate)
    assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedBlocks(model, candidate))
    Object.assign(model, candidate)
    return selection
  }
  if (typeof command.imageId !== 'string' || !command.imageId.trim()) throw new TypeError('图片身份无效。')
  assertDocumentImageEditable(model, command.imageId, command.id === 'doc.command.remove-image')
  const candidate = cloneModel(model), entry = documentImageEntries(candidate).find(entry => entry.imageId === command.imageId)!
  const old = entry.image
  let nextSelection = selection
  if (command.id === 'doc.command.remove-image') {
    if (entry.inline) {
      const atom = entry.inline
      replaceText(candidate, { anchor: { ...atom.target, offset: atom.startOffset }, focus: { ...atom.target, offset: atom.endOffset } }, '')
      const transform = (point: DocumentSelection['focus']) => sameTarget(point, atom.target)
        ? { ...point, offset: point.offset <= atom.startOffset ? point.offset : Math.max(atom.startOffset, point.offset - (atom.endOffset - atom.startOffset)) } : point
      nextSelection = { ...selection, anchor: transform(selection.anchor), focus: transform(selection.focus) }
    } else {
      entry.container.splice(entry.container.indexOf(entry.block), 1)
      if (!entry.container.length) entry.container.push(createParagraph())
    }
  } else if (command.id === 'doc.command.update-image') {
    setDocumentImageEntry(entry, patchDocumentImage(old, command.patch))
    if (command.patch.wrappingStyle !== undefined) {
      const source = documentImageEntryMetadata(entry).univerBodyDrawingSource
      if (source && typeof source === 'object' && !Array.isArray(source)) Reflect.deleteProperty(source, 'layoutType')
    }
  }
  else if (command.id === 'doc.command.replace-image') {
    const prepared = normalizeDocumentImage(command.image)
    setDocumentImageEntry(entry, normalizeDocumentImage({ ...old, ...prepared, width: old.width, height: old.height,
      angle: old.angle ?? 0, flipX: old.flipX ?? false, flipY: old.flipY ?? false, wrappingStyle: old.wrappingStyle ?? 'inline', positionH: old.positionH, positionV: old.positionV,
      distances: old.distances, layer: old.layer, keepAspectRatio: old.keepAspectRatio, crop: { top: 0, right: 0, bottom: 0, left: 0 } }))
  } else if (command.id === 'doc.command.reset-image-size') {
    if (!old.intrinsicWidth || !old.intrinsicHeight) throw new RangeError('图片原始尺寸尚未读取。')
    setDocumentImageEntry(entry, normalizeDocumentImage({ ...old, width: old.intrinsicWidth, height: old.intrinsicHeight,
      angle: 0, flipX: false, flipY: false, crop: { top: 0, right: 0, bottom: 0, left: 0 } }))
  } else {
    arrangeImages(candidate, [command.imageId], command.arrange)
  }
  synchronizeDocumentPictureControlImages(model, candidate)
  reconcileDocumentContentControlBlockEdits(model, candidate, { selection })
  validateDocumentContentControlRanges(candidate)
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedBlocks(model, candidate))
  Object.assign(model, candidate)
  const focus = clampPoint(model, nextSelection.focus)
  return findDocumentBlock(model, nextSelection.anchor.cellBlockId ?? nextSelection.anchor.blockId)
    && findDocumentBlock(model, nextSelection.focus.cellBlockId ?? nextSelection.focus.blockId) ? nextSelection : collapsedSelection(focus)
}
