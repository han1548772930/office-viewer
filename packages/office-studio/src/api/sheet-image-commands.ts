import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import { cloneModel } from '../domain/serialization.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { canViewProtectedRanges } from '../features/sheet-protection/range-protection.ts'
import { assertSheetRangeChangedValuesValid } from '../features/sheet-workbench/sheet-range-editor-validation.ts'
import { prepareSheetImageFiles, prepareSheetImageResources, prepareSheetImageUrl } from '../features/sheet-images/sheet-image-io.ts'
import { applySheetImagePlan, planSheetImageCommand, sheetImageCommandInput, SHEET_IMAGE_COMMANDS,
  type SheetImageCommandId, type SheetImageCommandInput } from '../features/sheet-images/sheet-image-commands.ts'
import { clearSheetImagePickers, requestSheetImagePicker } from './sheet-image-pickers.ts'
import { clearSheetImageSavers, requestSheetImageSaver } from './sheet-image-savers.ts'
import { clearSheetImageEditors, requestSheetImageEditor } from './sheet-image-editors.ts'
import { collectSheetCellImages, createSheetImageSaveRequest, sheetImageSaveConfig, sheetImageSaveReadRanges, validateSheetImageSaveConfig } from '../features/sheet-images/sheet-image-save.ts'
import { saveSheetImageFiles } from '../features/sheet-images/sheet-image-save-platform.ts'
import { formatProtectionRange } from '../features/sheet-protection/range-protection.ts'
import { sheetImagePlanAffectedRanges } from '../features/sheet-images/sheet-image-transaction.ts'
import { sheetClipboardObjectAxes, sheetClipboardObjectRange } from '../features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { sheetImageClipboardBounds } from '../features/sheet-images/sheet-image-clipboard.ts'
import { clearSheetImageCroppers, requestSheetImageCropper } from './sheet-image-croppers.ts'
import { SHEET_IMAGE_CROP_OPERATIONS } from '../features/sheet-images/sheet-image-crop-session.ts'
import { SHEET_IMAGE_ALIGN_OPERATIONS, sheetImageAlignTypeForOperation } from '../features/sheet-images/sheet-image-selection.ts'
import { clearSheetImageSelections, getSheetImageSelection } from './sheet-image-selection.ts'
import { runSheetDrawingCommand } from './sheet-drawing-commands.ts'

const asynchronousIds = [SHEET_IMAGE_COMMANDS.insertFloat, SHEET_IMAGE_COMMANDS.insertCell, SHEET_IMAGE_COMMANDS.addBackground] as const
function imageTarget(api: OfficeStudio, params: unknown, readOnly = false) {
  const input = sheetImageCommandInput(params)
  input.unitId ??= input.drawings?.[0]?.unitId; input.subUnitId ??= input.drawings?.[0]?.subUnitId
  const target = resolveSheetRangeCommandTarget(api, input)
  if (input.drawings?.some(drawing => drawing.unitId !== undefined && drawing.unitId !== target.unitId || drawing.subUnitId !== undefined && drawing.subUnitId !== target.sheetId)) throw new TypeError('图片操作需要同一工作簿及工作表。')
  const revision = api.getWorkbookRevision(target.unitId), historyGeneration = api.editor.history.captureSavepoint(target.unitId).generation, userId = api.currentUserId
  const primary = api.getSelectionState()?.primary, selection = api.getSelection()
  const anchor = { row: input.row ?? (selection?.unitId === target.unitId && selection.sheetId === target.sheetId ? primary?.startRow ?? selection.startRow : 0),
    column: input.column ?? (selection?.unitId === target.unitId && selection.sheetId === target.sheetId ? primary?.startColumn ?? selection.startColumn : 0) }
  const assertCurrent = () => {
    if (input.signal?.aborted) throw new DOMException('图片操作已取消。', 'AbortError')
    if (api.isDisposed || api.getWorkbookModel(target.unitId) !== target.workbook || api.getWorkbookRevision(target.unitId) !== revision
      || api.editor.history.captureSavepoint(target.unitId).generation !== historyGeneration || api.currentUserId !== userId
      || input.expectedRevision !== undefined && input.expectedRevision !== revision
      || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== historyGeneration) throw new Error('图片目标内容或历史已更改，请重新选择图片。')
    if (input.host && (input.host.unitId !== target.unitId || input.host.sheetId !== target.sheetId || api.activeWorkbookId !== target.unitId
      || api.editor.state.activeKind !== 'workbook' || target.workbook.activeSheetId !== target.sheetId)) throw new Error('图片宿主已切换，请重新选择图片。')
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    if (!readOnly) api.assertCanEditWorksheet(target.unitId, target.sheetId)
  }
  assertCurrent()
  return { input, target, anchor, revision, historyGeneration, assertCurrent }
}
function commitImage(api: OfficeStudio, state: ReturnType<typeof imageTarget>, id: SheetImageCommandId, input: SheetImageCommandInput): boolean {
  state.assertCurrent()
  const { target, anchor } = state
  const merge = target.sheet.mergedRanges.find(range => anchor.row >= range.startRow && anchor.row <= range.endRow && anchor.column >= range.startCol && anchor.column <= range.endCol)
  const cellAnchor = id === SHEET_IMAGE_COMMANDS.insertCell || (id === SHEET_IMAGE_COMMANDS.update || id === SHEET_IMAGE_COMMANDS.remove) && !target.sheet.images?.some(image => image.id === input.imageId)
    ? merge ? { row: merge.startRow, column: merge.startCol } : anchor : anchor
  const plan = planSheetImageCommand(target.sheet, id, { ...input, row: cellAnchor.row, column: cellAnchor.column }, cellAnchor,
    target.sheet.cells[`${cellAnchor.row}:${cellAnchor.column}`]?.image, sheetClipboardObjectAxes(target.workbook, target.sheet))
  const ranges = [...sheetImagePlanAffectedRanges(target.workbook, target.sheet, plan), ...(plan.cell && merge ? [merge] : [])]
  if (!canViewProtectedRanges(target.sheet, ranges, api.currentUserId)) throw new Error('没有查看目标图片范围的权限。')
  for (const range of ranges) api.assertCanEditRange(target.unitId, target.sheetId, range)
  state.assertCurrent()
  if (!plan.changed) return false
  const label = id === SHEET_IMAGE_COMMANDS.insertFloat || id === SHEET_IMAGE_COMMANDS.insertPrepared ? '插入浮动图片' : id === SHEET_IMAGE_COMMANDS.insertCell ? '插入单元格图片'
    : id === SHEET_IMAGE_COMMANDS.remove ? '删除图片' : id === SHEET_IMAGE_COMMANDS.arrange ? '排列图片' : sheetImageAlignTypeForOperation(id) !== undefined ? '对齐图片' : id === SHEET_IMAGE_COMMANDS.update || id === SHEET_IMAGE_COMMANDS.updateBatch ? '设置图片'
      : plan.backgroundImage ? '设置工作表背景图' : '删除工作表背景图'
  return api.commitWorkbook(target.unitId, label, model => {
    state.assertCurrent()
    const sheet = model.sheets.find(sheet => sheet.id === target.sheetId)!
    applySheetImagePlan(sheet, plan)
    if (plan.writesImages && id !== SHEET_IMAGE_COMMANDS.remove) for (const range of ranges) {
      if (sheet.rowCount !== undefined) sheet.rowCount = Math.max(sheet.rowCount, range.endRow + 1)
      if (sheet.columnCount !== undefined) sheet.columnCount = Math.max(sheet.columnCount, range.endCol + 1)
    }
    if (plan.cell) assertSheetRangeChangedValuesValid(target.sheet, model, target.sheetId,
      { userId: api.currentUserId, region: api.regionalSettings.region.value })
  }, { preserveStyleInheritance: true, recalculate: Boolean(plan.cell), expectedRevision: state.revision,
    expectedHistoryGeneration: state.historyGeneration, prepareView: model => { toExcelWorkbookView(model) } })
}
/** The picker, UI, Facade and file API share the same captured target and final content transaction. */
export function registerSheetImageCommands(api: OfficeStudio): Disposable {
  const registrations: Disposable[] = []
  const tasks = new Map<AbortController, string>()
  async function withTask<T>(unitId: string, signal: AbortSignal | undefined, run: (signal: AbortSignal) => Promise<T>): Promise<T> {
    const controller = new AbortController(), cancel = () => controller.abort()
    tasks.set(controller, unitId); signal?.addEventListener('abort', cancel, { once: true })
    if (signal?.aborted) controller.abort()
    try { return await run(controller.signal) }
    finally { signal?.removeEventListener('abort', cancel); tasks.delete(controller) }
  }
  for (const id of asynchronousIds) registrations.push(api.registerCommand<SheetImageCommandInput | undefined, Promise<boolean>>(id, async params => {
    const state = imageTarget(api, params), input = state.input
    return withTask(state.target.unitId, input.signal, async signal => {
      let resources = input.resources, images = input.images
      const options = { signal, current: () => { state.assertCurrent(); return true } }
      if (input.files !== undefined) {
        if (!input.files?.length) return false
        resources = await prepareSheetImageFiles(input.files, options)
      } else if (input.source !== undefined) resources = [await prepareSheetImageUrl(input.source, { signal, name: input.name })]
      else if (resources !== undefined) resources = await prepareSheetImageResources(resources, options)
      else if (images !== undefined) {
        const prepared = await prepareSheetImageResources(images.map(image => image.resource), options)
        images = images.map((image, index) => ({ ...image, resource: prepared[index]! }))
      } else {
      const requested = await requestSheetImagePicker(api, state.target.unitId, id, {
        ...input, signal, unitId: state.target.unitId, sheetId: state.target.sheetId, row: state.anchor.row, column: state.anchor.column,
        expectedRevision: state.revision, expectedHistoryGeneration: state.historyGeneration,
      })
      if (!requested) return false
      resources = requested.map(resource => cloneModel(resource))
      }
      if (signal.aborted) throw new DOMException('图片操作已取消。', 'AbortError')
      return commitImage(api, state, id, { ...input, resources, images })
    })
  }))
  for (const id of [SHEET_IMAGE_COMMANDS.insertPrepared, SHEET_IMAGE_COMMANDS.updateBatch, SHEET_IMAGE_COMMANDS.update, SHEET_IMAGE_COMMANDS.remove, SHEET_IMAGE_COMMANDS.setBackground, SHEET_IMAGE_COMMANDS.deleteBackground, SHEET_IMAGE_COMMANDS.arrange,
    ...Object.values(SHEET_IMAGE_ALIGN_OPERATIONS)] as const) {
    registrations.push(api.registerCommand<SheetImageCommandInput, boolean>(id, params => {
      if (id === SHEET_IMAGE_COMMANDS.arrange) return runSheetDrawingCommand(api, id, params)
      const state = imageTarget(api, params)
      if (sheetImageAlignTypeForOperation(id, state.input.alignType) !== undefined && state.input.imageIds === undefined && state.input.drawingIds === undefined && state.input.drawings === undefined)
        state.input.imageIds = getSheetImageSelection(api, state.target.unitId, state.target.sheetId)
      if ((id === SHEET_IMAGE_COMMANDS.insertPrepared || id === SHEET_IMAGE_COMMANDS.updateBatch) && state.input.images?.some(image => !/^data:/i.test(image.resource.source))) {
        throw new TypeError('同步批量图片命令需要已读取的图片资源，请先使用 buildAsync。')
      }
      return commitImage(api, state, id, state.input)
    }))
  }
  registrations.push(api.registerCommand<SheetImageCommandInput, Promise<boolean>>(SHEET_IMAGE_COMMANDS.replace, async params => {
    const state = imageTarget(api, params), input = state.input
    return withTask(state.target.unitId, input.signal, async signal => {
      const resources = input.files ? await prepareSheetImageFiles(input.files, { signal, current: () => { state.assertCurrent(); return true } })
        : input.source !== undefined ? [await prepareSheetImageUrl(input.source, { signal, name: input.name })]
          : await prepareSheetImageResources(input.resources ?? [], { signal, current: () => { state.assertCurrent(); return true } })
      if (!resources.length) return false
      if (resources.length !== 1) throw new RangeError('替换图片需要且只能使用一张图片。')
      if (signal.aborted) throw new DOMException('图片替换已取消。', 'AbortError')
      return commitImage(api, state, SHEET_IMAGE_COMMANDS.update, { ...input, patch: { ...input.patch, resource: resources[0]! } })
    })
  }))
  registrations.push(api.registerCommand<SheetImageCommandInput, boolean>(SHEET_IMAGE_COMMANDS.edit, params => {
    const state = imageTarget(api, params)
    return requestSheetImageEditor(api, state.target.unitId, { ...state.input, unitId: state.target.unitId, sheetId: state.target.sheetId,
      expectedRevision: state.revision, expectedHistoryGeneration: state.historyGeneration })
  }))
  for (const id of [SHEET_IMAGE_COMMANDS.saveCell, SHEET_IMAGE_COMMANDS.save] as const) registrations.push(api.registerCommand<SheetImageCommandInput | undefined, Promise<boolean>>(id, async params => {
    const state = imageTarget(api, params, true), { target, input } = state
    const floating = id === SHEET_IMAGE_COMMANDS.save && target.sheet.images?.find(image => image.id === input.imageId)
    const ranges = floating ? [sheetClipboardObjectRange(target.workbook, target.sheet.id, sheetImageClipboardBounds(target.workbook, target.sheet, floating))] : target.ranges
    const entries = floating ? [{ row: floating.anchor.row, column: floating.anchor.column, cellAddress: formatProtectionRange(ranges[0]!), image: cloneModel(floating) }] : collectSheetCellImages(target.sheet, ranges)
    if (!entries.length) return false
    if (!canViewProtectedRanges(target.sheet, floating ? ranges : entries.map(entry => ({ startRow: entry.row, endRow: entry.row, startCol: entry.column, endCol: entry.column })), api.currentUserId)) throw new Error('没有保存目标图片范围的权限。')
    const request = createSheetImageSaveRequest(target.sheet, ranges, entries, (row, column) => canViewProtectedRanges(target.sheet,
      [{ startRow: row, endRow: row, startCol: column, endCol: column }], api.currentUserId))
    return withTask(target.unitId, input.signal, async signal => {
      const save = async (value: import('../features/sheet-images/sheet-image-save.ts').SheetImageSaveConfig, uiSignal?: AbortSignal) => {
        const config = validateSheetImageSaveConfig(value)
        const assertCurrent = () => {
          if (signal.aborted || uiSignal?.aborted) throw new DOMException('图片保存已取消。', 'AbortError')
          state.assertCurrent()
          if (!canViewProtectedRanges(target.sheet, [...(floating ? ranges : []), ...sheetImageSaveReadRanges(request, config)], api.currentUserId)) throw new Error('没有保存图片或读取文件名列的权限。')
        }
        // The outer API owner and the dialog owner can independently cancel a save.
        // Link both signals so cancellation also stops an in-flight image fetch.
        return withTask(target.unitId, signal, outerSignal => withTask(target.unitId, uiSignal, uiTaskSignal => {
          const effective = new AbortController(), cancel = () => effective.abort()
          outerSignal.addEventListener('abort', cancel, { once: true }); uiTaskSignal.addEventListener('abort', cancel, { once: true })
          if (outerSignal.aborted || uiTaskSignal.aborted) effective.abort()
          return saveSheetImageFiles(request, config, { signal: effective.signal, single: entries.length === 1, assertCurrent }).finally(() => {
            outerSignal.removeEventListener('abort', cancel); uiTaskSignal.removeEventListener('abort', cancel)
          })
        }))
      }
      if (entries.length > 1 && input.showSaveDialog) return requestSheetImageSaver(api, target.unitId, request, { ...input, signal, unitId: target.unitId, sheetId: target.sheetId,
        expectedRevision: state.revision, expectedHistoryGeneration: state.historyGeneration }, save)
      return save(input.saveConfig ? validateSheetImageSaveConfig(input.saveConfig) : sheetImageSaveConfig())
    })
  }))
  for (const id of Object.values(SHEET_IMAGE_CROP_OPERATIONS)) registrations.push(api.registerCommand<SheetImageCommandInput | undefined, boolean>(id, params => {
    const state = imageTarget(api, params), result = requestSheetImageCropper(api, state.target.unitId, id,
      { ...state.input, unitId: state.target.unitId, sheetId: state.target.sheetId })
    if (!result.handled) return false
    state.assertCurrent()
    if (!result.image) return true
    if (result.expectedRevision !== undefined && result.expectedRevision !== state.revision) throw new Error('图片裁剪内容已过期。')
    return commitImage(api, state, SHEET_IMAGE_COMMANDS.updateBatch, { ...state.input, images: [result.image] })
  }))
  const disposed = api.events.addEvent('WorkbookDisposed', event => {
    clearSheetImagePickers(api, event.workbookId); clearSheetImageSavers(api, event.workbookId); clearSheetImageEditors(api, event.workbookId)
    clearSheetImageCroppers(api, event.workbookId)
    clearSheetImageSelections(api, event.workbookId)
    for (const [controller, unitId] of tasks) if (unitId === event.workbookId) controller.abort()
  })
  return { dispose: () => { tasks.forEach((_unitId, controller) => controller.abort()); tasks.clear(); registrations.forEach(registration => registration.dispose()); disposed.dispose(); clearSheetImagePickers(api); clearSheetImageSavers(api); clearSheetImageEditors(api); clearSheetImageCroppers(api); clearSheetImageSelections(api) } }
}
