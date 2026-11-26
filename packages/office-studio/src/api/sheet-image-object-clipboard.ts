import type { OfficeStudio } from './office-studio.ts'
import type { SheetClipboardCommandInput, SheetClipboardCopyResult, SheetClipboardLocation, SheetClipboardPasteResult } from '../features/sheet-workbench/sheet-clipboard-types.ts'
import { captureSheetImageObjects, planPasteSheetImageObjects, sheetImageObjectClipboardContent, type SheetImageObjectClipboard } from '../features/sheet-images/sheet-image-object-clipboard.ts'
import { sheetImageClipboardBounds } from '../features/sheet-images/sheet-image-clipboard.ts'
import { sheetClipboardObjectRange } from '../features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { canViewProtectedRanges } from '../features/sheet-protection/range-protection.ts'
import { SHEET_IMAGE_COMMANDS } from '../features/sheet-images/sheet-image-commands.ts'

export function copySheetImageObjects(api: OfficeStudio, input: SheetClipboardCommandInput, target: SheetClipboardLocation, cut: boolean): SheetClipboardCopyResult {
  const workbook = api.getWorkbookModel(target.unitId), sheet = workbook?.sheets.find(sheet => sheet.id === target.sheetId)
  if (!workbook || !sheet) return { accepted: false, plain: '', html: '' }
  api.assertCanViewWorksheet(target.unitId, target.sheetId)
  const revision = api.getWorkbookRevision(target.unitId), historyGeneration = api.editor.history.captureSavepoint(target.unitId).generation, userId = api.currentUserId
  const current = () => api.currentUserId === userId && api.getWorkbookModel(target.unitId) === workbook
    && api.getWorkbookRevision(target.unitId) === revision && api.editor.history.captureSavepoint(target.unitId).generation === historyGeneration
    && (input.expectedRevision === undefined || input.expectedRevision === revision)
    && (input.expectedHistoryGeneration === undefined || input.expectedHistoryGeneration === historyGeneration)
  if (!current()) return { accepted: false, plain: '', html: '' }
  const snapshot = captureSheetImageObjects(workbook, sheet, input.imageIds ?? [], userId, cut ? 'cut' : 'copy')
  const ranges = snapshot.entries.map(entry => sheetClipboardObjectRange(workbook, sheet.id, entry.bounds))
  if (!canViewProtectedRanges(sheet, ranges, api.currentUserId)) throw new Error('没有复制目标图片的权限。')
  if (cut) for (const range of ranges) api.assertCanEditRange(target.unitId, target.sheetId, range)
  const content = sheetImageObjectClipboardContent(snapshot), generation = api.sheetClipboard.clipboardGeneration
  if (!current()) return { accepted: false, plain: '', html: '' }
  const changed = cut && api.executeCommandSync(SHEET_IMAGE_COMMANDS.remove, { ...input, unitId: target.unitId, sheetId: target.sheetId,
    imageIds: snapshot.entries.map(entry => entry.image.id), expectedRevision: revision, expectedHistoryGeneration: historyGeneration })
  if (cut && !changed) return { accepted: false, plain: '', html: '' }
  const published = api.getWorkbookModel(target.unitId)
  if (published && api.currentUserId === userId && api.getWorkbookRevision(target.unitId) === revision + (changed ? 1 : 0) && api.sheetClipboard.clipboardGeneration === generation
    && api.editor.history.captureSavepoint(target.unitId).generation === historyGeneration) api.sheetClipboard.setImageClipboard(snapshot)
  return { accepted: true, changed, imageClipboard: snapshot, ...content }
}

export function pasteSheetImageObjects(api: OfficeStudio, input: SheetClipboardCommandInput, target: SheetClipboardLocation, clipboard: SheetImageObjectClipboard): SheetClipboardPasteResult {
  const workbook = api.getWorkbookModel(target.unitId), sheet = workbook?.sheets.find(sheet => sheet.id === target.sheetId)
  if (!workbook || !sheet) return { accepted: false, changed: false, ranges: [], issues: [], reason: 'stale-target' }
  const revision = api.getWorkbookRevision(target.unitId), historyGeneration = api.editor.history.captureSavepoint(target.unitId).generation, userId = api.currentUserId
  const current = () => api.currentUserId === userId && api.getWorkbookModel(target.unitId) === workbook
    && api.getWorkbookRevision(target.unitId) === revision && api.editor.history.captureSavepoint(target.unitId).generation === historyGeneration
    && (input.expectedRevision === undefined || input.expectedRevision === revision)
    && (input.expectedHistoryGeneration === undefined || input.expectedHistoryGeneration === historyGeneration)
  if (!current()) return { accepted: false, changed: false, ranges: [], issues: [], reason: 'stale-target' }
  if (clipboard.source.userId !== userId) throw new Error('图片剪贴板属于其他用户。')
  const source = api.getWorkbookModel(clipboard.source.unitId), sourceSheet = source?.sheets.find(sheet => sheet.id === clipboard.source.sheetId)
  if (source && sourceSheet) {
    api.assertCanViewWorksheet(clipboard.source.unitId, clipboard.source.sheetId)
    const ranges = clipboard.entries.map(entry => {
      const current = sourceSheet.images?.find(image => image.id === entry.image.id)
      return sheetClipboardObjectRange(source, sourceSheet.id, current ? sheetImageClipboardBounds(source, sourceSheet, current) : entry.bounds)
    })
    if (!canViewProtectedRanges(sourceSheet, ranges, api.currentUserId)) throw new Error('没有读取图片剪贴板来源的权限。')
  }
  const plan = planPasteSheetImageObjects(workbook, sheet, target.range, clipboard)
  if (!current()) return { accepted: false, changed: false, ranges: [], issues: [], reason: 'stale-target' }
  const changed = api.executeCommandSync(SHEET_IMAGE_COMMANDS.insertPrepared, { ...input, unitId: target.unitId, sheetId: target.sheetId,
    expectedRevision: revision, expectedHistoryGeneration: historyGeneration, images: plan.images })
  return { accepted: true, changed, ranges: plan.ranges, issues: [], imageIds: changed ? plan.images.map(image => image.id) : [] }
}
