import { computed, nextTick, onUnmounted, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { bindSheetImagePicker } from '../../api/sheet-image-pickers.ts'
import { bindSheetImageEditor } from '../../api/sheet-image-editors.ts'
import { bindSheetImageSaver, type SheetImageSaveAction } from '../../api/sheet-image-savers.ts'
import { bindSheetImageCropper } from '../../api/sheet-image-croppers.ts'
import { bindSheetImageSelection } from '../../api/sheet-image-selection.ts'
import { sheetImageAlignTypeForOperation } from '../../features/sheet-images/sheet-image-selection.ts'
import { SHEET_IMAGE_CROP_OPERATIONS, type SheetImageCropOperation, type SheetImageCropResult } from '../../features/sheet-images/sheet-image-crop-session.ts'
import type { SheetCellImage, SheetFloatingImage, SheetImageResource } from '../../domain/sheet-images.ts'
import { createBlankWorkbook } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { applyExcelWorkbookView, cellKey, parseCellKey, toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { assertSheetRangeChangedValuesValid } from '../../features/sheet-workbench/sheet-range-editor-validation.ts'
import { prepareSheetImageFiles, prepareSheetImageResources, prepareSheetImageUrl } from '../../features/sheet-images/sheet-image-io.ts'
import { collectSheetCellImages, createSheetImageSaveRequest, sheetImageSaveConfig, sheetImageSaveReadRanges, validateSheetImageSaveConfig,
  type SheetImageSaveConfig, type SheetImageSaveRequest } from '../../features/sheet-images/sheet-image-save.ts'
import { saveSheetImageFiles } from '../../features/sheet-images/sheet-image-save-platform.ts'
import type { CellRange } from '../../domain/model.ts'
import { formatProtectionRange } from '../../features/sheet-protection/range-protection.ts'
import { applySheetImagePlan, planSheetImageCommand, sheetImageCommandInput, SHEET_IMAGE_COMMANDS,
  type SheetImageCommandId, type SheetImageCommandInput } from '../../features/sheet-images/sheet-image-commands.ts'
import { captureSheetAxisCommandHost, type SheetAxisHostChangeOptions } from './sheet-axis-command-host.ts'
import { sheetImagePlanAffectedRanges } from '../../features/sheet-images/sheet-image-transaction.ts'
import { sheetClipboardObjectAxes, sheetClipboardObjectRange } from '../../features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { sheetImageClipboardBounds } from '../../features/sheet-images/sheet-image-clipboard.ts'
import type { SheetImageAxes } from '../../features/sheet-images/sheet-image-placement.ts'

interface SheetImageUiContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  selectedKey(): string
  ranges(): CellRange[]
  userId(): string
  active(): boolean
  finishEdit(): boolean
  prepareView?(view: ExcelWorkbookView, previous: ExcelWorkbookView): ExcelWorkbookView
  applyView(view: ExcelWorkbookView): void
  changed(options?: SheetAxisHostChangeOptions): void
  notify(message: string): void
  focus(): void
  crop?(id: SheetImageCropOperation, input: SheetImageCommandInput): SheetImageCropResult
}
interface CapturedImageHost {
  host: ReturnType<typeof captureSheetAxisCommandHost>
  input: SheetImageCommandInput
  signature: string
  model?: ReturnType<OfficeStudio['getWorkbookModel']>
  userId: string
  generation: number
}
export interface SheetImagePickerRequest { kind: 'floating' | 'cell' | 'background'; multiple: boolean; target: string }
interface PendingImagePicker extends CapturedImageHost {
  commandId: SheetImageCommandId
  request: SheetImagePickerRequest
  resolve(resources: readonly SheetImageResource[] | null): void
  controller?: AbortController
}
export interface SheetImageEditorRequest { image: SheetCellImage | SheetFloatingImage; row?: number; column?: number; axes?: SheetImageAxes }
interface PendingImageEditor extends CapturedImageHost { request: SheetImageEditorRequest; controller?: AbortController }
interface PendingImageSave extends CapturedImageHost { request: SheetImageSaveRequest; save: SheetImageSaveAction; resolve(saved: boolean): void; controller?: AbortController }

export function useSheetImages(context: SheetImageUiContext) {
  const pending = shallowRef<PendingImagePicker>(), editor = shallowRef<PendingImageEditor>(), saving = shallowRef<PendingImageSave>(), imageError = ref(''), busy = ref(false), progress = ref('')
  const selectedImageIds = ref<string[]>([]), selectedImageId = computed<string | undefined>({ get: () => selectedImageIds.value.at(-1), set: id => { selectedImageIds.value = id ? [id] : [] } })
  let generation = 0, closeCurrent: (() => boolean) | undefined
  const pickerRequest = computed(() => pending.value?.request ?? null), editorRequest = computed(() => editor.value?.request ?? null)
  const saveRequest = computed(() => saving.value?.request ?? null)
  function capture(params?: unknown, readOnly = false): CapturedImageHost {
    const input = sheetImageCommandInput(params), host = captureSheetAxisCommandHost(context), { view, studio, sheetId } = host
    if (!host.sameTarget() || input.unitId !== undefined && input.unitId !== view.id || (input.sheetId ?? input.subUnitId ?? sheetId) !== sheetId) throw new Error('请在目标工作表操作图片。')
    const sheet = view.sheets.find(item => item.id === sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, context.userId()) || !readOnly && !canEditWorksheetProtection(sheet, context.userId())) throw new Error('没有操作目标工作表图片的权限。')
    const anchor = parseCellKey(context.selectedKey()) ?? { row: 0, col: 0 }
    const revision = studio && view.id ? studio.getWorkbookRevision(view.id) : undefined
    const historyGeneration = studio && view.id ? studio.editor.history.captureSavepoint(view.id).generation : undefined
    if (input.expectedRevision !== undefined && input.expectedRevision !== revision || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== historyGeneration) throw new Error('图片目标内容或历史已更改。')
    return { host, signature: JSON.stringify(view), model: studio && view.id ? studio.getWorkbookModel(view.id) : undefined,
      userId: context.userId(), generation, input: { ...input, unitId: view.id, sheetId, row: input.row ?? anchor.row, column: input.column ?? anchor.col,
        expectedRevision: revision, expectedHistoryGeneration: historyGeneration, host: view.id ? { unitId: view.id, sheetId } : undefined } }
  }
  function assertCurrent(state: CapturedImageHost): void {
    const { host, input } = state, studio = host.studio, unitId = input.unitId
    if (state.input.signal?.aborted || state.generation !== generation || !host.sameTarget() || context.userId() !== state.userId
      || JSON.stringify(context.workbook()) !== state.signature) throw new Error('图片目标已更改，请重新打开图片工具。')
    if (studio && unitId && (studio.getWorkbookModel(unitId) !== state.model || studio.getWorkbookRevision(unitId) !== input.expectedRevision
      || studio.editor.history.captureSavepoint(unitId).generation !== input.expectedHistoryGeneration)) throw new Error('图片目标内容或历史已更改。')
  }
  function cancelPicker(): void {
    const state = pending.value
    state?.controller?.abort(); pending.value = undefined; state?.resolve(null); busy.value = false; progress.value = ''
    closeCurrent = state?.host.publishedCurrent()
  }
  function cancelEditor(): void { const state = editor.value; state?.controller?.abort(); editor.value = undefined; busy.value = false; closeCurrent = state?.host.publishedCurrent() }
  function cancelSave(): void { const state = saving.value; state?.controller?.abort(); saving.value = undefined; state?.resolve(false); busy.value = false; closeCurrent = state?.host.publishedCurrent() }
  function afterClose(): void {
    const current = closeCurrent
    closeCurrent = undefined
    void nextTick(() => { if (current?.() && context.active() && !pending.value && !editor.value && !saving.value) context.focus() })
  }
  async function requestPicker(id: SheetImageCommandId, input: SheetImageCommandInput): Promise<readonly SheetImageResource[] | null> {
    if (!context.active() || !context.finishEdit()) return null
    const state = capture(input)
    if (pending.value || editor.value || saving.value) return null
    const kind = id === SHEET_IMAGE_COMMANDS.insertFloat ? 'floating' : id === SHEET_IMAGE_COMMANDS.insertCell ? 'cell' : 'background'
    if (kind !== 'background') assertSelectionPermission(state)
    imageError.value = ''; progress.value = ''; busy.value = false
    return new Promise(resolve => {
      const cancel = () => { if (pending.value === request) cancelPicker() }
      const request: PendingImagePicker = { ...state, commandId: id, resolve: resources => { input.signal?.removeEventListener('abort', cancel); resolve(resources) },
        request: { kind, multiple: kind === 'floating', target: cellKey(state.input.row!, state.input.column!) } }
      pending.value = request; input.signal?.addEventListener('abort', cancel, { once: true })
      if (input.signal?.aborted) cancel()
    })
  }
  function assertSelectionPermission(state: CapturedImageHost): void {
    const sheet = state.host.view.sheets.find(item => item.id === state.input.sheetId)!
    const row = state.input.row!, column = state.input.column!, ranges = [{ startRow: row, endRow: row, startCol: column, endCol: column }]
    if (!canEditProtectedRanges(sheet, ranges, state.userId) || !canViewProtectedRanges(sheet, ranges, state.userId)) throw new Error('没有操作目标图片范围的权限。')
  }
  async function confirmPicker(files: readonly File[]): Promise<void> {
    const state = pending.value
    if (!state || busy.value || !files.length) return
    const controller = new AbortController(); state.controller = controller; busy.value = true; imageError.value = ''
    try {
      assertCurrent(state)
      if (!state.request.multiple && files.length > 1) throw new Error('此入口一次只能选择一张图片。')
      const resources = await prepareSheetImageFiles(files, { signal: controller.signal, current: () => { assertCurrent(state); return true },
        progress: (completed, total) => { if (pending.value === state) progress.value = `${completed} / ${total}` } })
      assertCurrent(state)
      if (pending.value !== state) return
      pending.value = undefined; closeCurrent = state.host.publishedCurrent(); state.resolve(resources)
    } catch (reason: unknown) {
      if (pending.value === state && !controller.signal.aborted) imageError.value = reason instanceof Error ? reason.message : '无法读取图片。'
    } finally { if (pending.value === state || !pending.value) { busy.value = false; progress.value = '' } }
  }
  async function runCaptured(state: CapturedImageHost, commandId: SheetImageCommandId, params: SheetImageCommandInput): Promise<boolean> {
    assertCurrent(state)
    const input = { ...state.input, ...params }, { host } = state, { studio, view } = host
    if (studio && view.id) {
      const changed = await studio.executeCommand<SheetImageCommandInput, boolean>(commandId, input)
      if (!changed) return false
      const model = studio.getWorkbookModel(view.id)
      if (!model || !host.sameTarget() || studio.getWorkbookRevision(view.id) !== input.expectedRevision! + 1 ||
        studio.editor.history.captureSavepoint(view.id).generation !== input.expectedHistoryGeneration) return true
      const next = toExcelWorkbookView(model), signature = JSON.stringify(next)
      context.applyView(next)
      const published = host.publishedCurrent(input.expectedRevision! + 1)
      const current = () => published() && studio.getWorkbookModel(view.id!) === model
        && studio.editor.history.captureSavepoint(view.id!).generation === input.expectedHistoryGeneration && JSON.stringify(context.workbook()) === signature
      closeCurrent = current
      if (current()) context.changed({ syncSelection: false, current })
      return true
    }
    const before = applyExcelWorkbookView(createBlankWorkbook(), view), candidate = cloneModel(before), sheet = candidate.sheets.find(item => item.id === host.sheetId)!
    const row = input.row!, column = input.column!, merge = sheet.mergedRanges.find(range => row >= range.startRow && row <= range.endRow && column >= range.startCol && column <= range.endCol)
    const isCell = commandId === SHEET_IMAGE_COMMANDS.insertCell || (commandId === SHEET_IMAGE_COMMANDS.update || commandId === SHEET_IMAGE_COMMANDS.remove) && !sheet.images?.some(image => image.id === input.imageId)
    const anchor = isCell && merge ? { row: merge.startRow, column: merge.startCol } : { row, column }
    const plan = planSheetImageCommand(sheet, commandId, { ...input, row: anchor.row, column: anchor.column }, anchor, sheet.cells[`${anchor.row}:${anchor.column}`]?.image, sheetClipboardObjectAxes(candidate, sheet))
    const ranges = [...sheetImagePlanAffectedRanges(candidate, sheet, plan), ...(plan.cell && merge ? [merge] : [])]
    if (!canViewProtectedRanges(sheet, ranges, state.userId) || !canEditProtectedRanges(sheet, ranges, state.userId)) throw new Error('没有操作目标图片范围的权限。')
    if (!plan.changed) return false
    applySheetImagePlan(sheet, plan)
    if (plan.cell) assertSheetRangeChangedValuesValid(before.sheets.find(item => item.id === sheet.id)!, candidate, sheet.id)
    const detached = toExcelWorkbookView(candidate), next = context.prepareView?.(detached, cloneModel(view)) ?? detached
    if (view.id === undefined) delete next.id
    else next.id = view.id
    assertCurrent(state)
    if (!host.sourceCurrent()) throw new Error('图片工作簿宿主已替换。')
    const signature = JSON.stringify(next)
    context.applyView(next)
    const published = host.publishedCurrent(), current = () => published() && JSON.stringify(context.workbook()) === signature
    closeCurrent = current
    if (current()) context.changed({ syncSelection: false, current })
    return true
  }
  async function runImageCommand(commandId: string, params?: unknown): Promise<boolean> {
    if (!context.active() || !context.finishEdit() || !context.active()) return false
    try {
      const id = commandId as SheetImageCommandId, readOnly = id === SHEET_IMAGE_COMMANDS.saveCell || id === SHEET_IMAGE_COMMANDS.save
      const state = capture(params, readOnly)
      let input = state.input
      if (sheetImageAlignTypeForOperation(id, input.alignType) !== undefined && input.imageIds === undefined && input.drawingIds === undefined && input.drawings === undefined)
        input = { ...input, imageIds: [...selectedImageIds.value] }
      if (id === SHEET_IMAGE_COMMANDS.edit) { if (!input.imageId) return false; openImageEditor(input.imageId, input.row, input.column, input); return Boolean(editor.value) }
      if (readOnly) return await runSave(state, id, input)
      const standalone = !state.host.studio || !state.input.unitId
      if ((Object.values(SHEET_IMAGE_CROP_OPERATIONS) as readonly string[]).includes(id)) {
        if (!standalone) return state.host.studio!.executeCommand(id, input)
        const result = context.crop?.(id as SheetImageCropOperation, input)
        if (!result?.handled) return false
        assertCurrent(state)
        return result.image ? await runCaptured(state, SHEET_IMAGE_COMMANDS.updateBatch, { ...input, images: [result.image] }) : true
      }
      if (standalone && input.files !== undefined) {
        if (!input.files?.length) return false
        input = { ...input, resources: await prepareSheetImageFiles(input.files, { signal: input.signal, current: () => { assertCurrent(state); return true } }) }
      } else if (standalone && input.source !== undefined) input = { ...input, resources: [await prepareSheetImageUrl(input.source, { signal: input.signal, name: input.name })] }
      else if (standalone && input.resources !== undefined) input = { ...input, resources: await prepareSheetImageResources(input.resources, { signal: input.signal, current: () => { assertCurrent(state); return true } }) }
      else if (standalone && input.images !== undefined) {
        const resources = await prepareSheetImageResources(input.images.map(image => image.resource), { signal: input.signal, current: () => { assertCurrent(state); return true } })
        input = { ...input, images: input.images.map((image, index) => ({ ...image, resource: resources[index]! })) }
      }
      if ([SHEET_IMAGE_COMMANDS.insertFloat, SHEET_IMAGE_COMMANDS.insertCell, SHEET_IMAGE_COMMANDS.addBackground].some(item => item === id)
        && input.resources === undefined && input.images === undefined && input.files === undefined && input.source === undefined && standalone) {
        const resources = await requestPicker(id, input)
        if (!resources) return false
        input = { ...input, resources }
      }
      const changed = await runCaptured(state, id, input)
      if (changed && id === SHEET_IMAGE_COMMANDS.remove) {
        const removed = new Set(input.imageIds ?? (input.imageId ? [input.imageId] : [])); selectedImageIds.value = selectedImageIds.value.filter(id => !removed.has(id))
      }
      return changed
    } catch (reason: unknown) { context.notify(reason instanceof Error ? reason.message : '图片操作失败。'); return false }
  }
  function openImageEditor(imageId: string, row?: number, column?: number, input?: SheetImageCommandInput): void {
    if (!context.active() || !context.finishEdit() || pending.value || editor.value || saving.value) return
    try {
      const state = capture({ ...input, row, column }), sheet = state.host.view.sheets.find(item => item.id === state.input.sheetId)!
      let image: SheetCellImage | SheetFloatingImage | undefined = sheet.images?.find(item => item.id === imageId)
      if (!image) {
        const entries = Object.entries(sheet.cells).filter(([, cell]) => cell.image?.id === imageId)
        const entry = row !== undefined && column !== undefined ? entries.find(([key]) => key === cellKey(row, column)) : entries.length === 1 ? entries[0] : undefined
        if (!entry && entries.length > 1) throw new Error('图片身份重复，请从目标单元格打开属性。')
        const point = entry && parseCellKey(entry[0])
        image = entry?.[1].image ?? undefined
        if (point) { state.input.row = point.row; state.input.column = point.col }
      }
      if (!image) throw new Error('找不到目标图片。')
      if ('anchor' in image) { state.input.row = image.anchor.row; state.input.column = image.anchor.column }
      assertSelectionPermission(state)
      const workbook = state.host.studio && state.input.unitId ? state.host.studio.getWorkbookModel(state.input.unitId) : applyExcelWorkbookView(createBlankWorkbook(), state.host.view)
      const sourceSheet = workbook?.sheets.find(sheet => sheet.id === state.input.sheetId)
      const axes = workbook && sourceSheet ? sheetClipboardObjectAxes(workbook, sourceSheet) : undefined
      if ('anchor' in image && workbook && sourceSheet && !canViewProtectedRanges(sourceSheet, [sheetClipboardObjectRange(workbook, sourceSheet.id, sheetImageClipboardBounds(workbook, sourceSheet, image))], state.userId)) throw new Error('没有读取目标图片范围的权限。')
      editor.value = { ...state, request: { image: cloneModel(image), axes, row: 'anchor' in image ? undefined : state.input.row, column: 'anchor' in image ? undefined : state.input.column } }
      imageError.value = ''; busy.value = false
    } catch (reason: unknown) { context.notify(reason instanceof Error ? reason.message : '无法打开图片属性。') }
  }
  async function confirmEditor(patch: SheetImageCommandInput['patch'], replaceFiles: readonly File[] = []): Promise<void> {
    const state = editor.value
    if (!state || busy.value) return
    const controller = new AbortController(); state.controller = controller; busy.value = true; imageError.value = ''
    try {
      assertCurrent(state)
      if (replaceFiles.length > 1) throw new Error('替换图片一次只能选择一个文件。')
      const replacement = replaceFiles.length ? (await prepareSheetImageFiles(replaceFiles, { signal: controller.signal,
        current: () => { assertCurrent(state); return true } }))[0] : undefined
      assertCurrent(state)
      await runCaptured(state, SHEET_IMAGE_COMMANDS.update, { ...state.input, imageId: state.request.image.id,
        patch: { ...patch, ...(replacement ? { resource: replacement } : {}) } })
      if (editor.value === state) { editor.value = undefined; closeCurrent ??= state.host.publishedCurrent() }
    } catch (reason: unknown) { if (editor.value === state && !controller.signal.aborted) imageError.value = reason instanceof Error ? reason.message : '无法修改图片。' }
    finally { if (editor.value === state || !editor.value) busy.value = false }
  }
  async function requestSaver(request: SheetImageSaveRequest, input: SheetImageCommandInput, save: SheetImageSaveAction): Promise<boolean> {
    if (!context.active() || !context.finishEdit() || pending.value || editor.value || saving.value) return false
    const state = capture(input, true); imageError.value = ''; progress.value = ''; busy.value = false
    return new Promise(resolve => {
      const cancel = () => { if (saving.value === task) cancelSave() }
      const task: PendingImageSave = { ...state, request: cloneModel(request), save,
        resolve: value => { input.signal?.removeEventListener('abort', cancel); resolve(value) } }
      saving.value = task; input.signal?.addEventListener('abort', cancel, { once: true }); if (input.signal?.aborted) cancel()
    })
  }
  async function confirmSave(config: SheetImageSaveConfig): Promise<void> {
    const state = saving.value
    if (!state || busy.value) return
    const controller = new AbortController(); state.controller = controller; busy.value = true; imageError.value = ''
    try {
      assertCurrent(state)
      const saved = await state.save(validateSheetImageSaveConfig(config), controller.signal)
      assertCurrent(state)
      if (saving.value === state && saved) { saving.value = undefined; closeCurrent = state.host.publishedCurrent(); state.resolve(true) }
    } catch (reason: unknown) { if (saving.value === state && !controller.signal.aborted) imageError.value = reason instanceof Error ? reason.message : '无法保存图片。' }
    finally { if (saving.value === state || !saving.value) busy.value = false }
  }
  async function runSave(state: CapturedImageHost, id: SheetImageCommandId, input: SheetImageCommandInput): Promise<boolean> {
    const { studio, view } = state.host
    if (studio && view.id) return studio.executeCommand(id, { ...input, ranges: input.ranges ?? context.ranges(), showSaveDialog: input.showSaveDialog ?? true })
    const model = applyExcelWorkbookView(createBlankWorkbook(), view), sheet = model.sheets.find(item => item.id === context.sheetId())!
    const floating = id === SHEET_IMAGE_COMMANDS.save && sheet.images?.find(image => image.id === input.imageId)
    const ranges = floating ? [sheetClipboardObjectRange(model, sheet.id, sheetImageClipboardBounds(model, sheet, floating))] : input.ranges ?? context.ranges()
    const entries = floating ? [{ row: floating.anchor.row, column: floating.anchor.column, cellAddress: formatProtectionRange(ranges[0]!), image: floating }] : collectSheetCellImages(sheet, ranges)
    if (!entries.length) return false
    if (!canViewProtectedRanges(sheet, floating ? ranges : entries.map(entry => ({ startRow: entry.row, endRow: entry.row, startCol: entry.column, endCol: entry.column })), state.userId)) throw new Error('没有保存目标图片范围的权限。')
    const request = createSheetImageSaveRequest(sheet, ranges, entries, (row, column) => canViewProtectedRanges(sheet,
      [{ startRow: row, endRow: row, startCol: column, endCol: column }], state.userId))
    const save = (config: SheetImageSaveConfig, signal?: AbortSignal) => saveSheetImageFiles(request, config, { single: entries.length === 1, signal,
      assertCurrent: () => { assertCurrent(state); if (!canViewProtectedRanges(sheet, [...(floating ? ranges : []), ...sheetImageSaveReadRanges(request, config)], state.userId)) throw new Error('没有保存图片或读取文件名列的权限。') } })
    return entries.length > 1 && input.showSaveDialog !== false ? requestSaver(request, input, save) : save(input.saveConfig ?? sheetImageSaveConfig(), input.signal)
  }
  function saveImage(imageId: string): void {
    void runImageCommand(SHEET_IMAGE_COMMANDS.save, { imageId })
  }
  watch([() => context.studio(), () => context.workbook().id], ([studio, unitId], _previous, onCleanup) => {
    if (!studio || !unitId || studio.isDisposed || !studio.getWorkbookModel(unitId)) return
    const owner = bindSheetImagePicker(studio, unitId, requestPicker)
    const saver = bindSheetImageSaver(studio, unitId, requestSaver)
    const imageEditor = bindSheetImageEditor(studio, unitId, input => { if (!input.imageId) return false; openImageEditor(input.imageId, input.row, input.column, input); return Boolean(editor.value) })
    const cropper = bindSheetImageCropper(studio, unitId, (id, input) => {
      if (!context.active() || pending.value || editor.value || saving.value || !context.finishEdit()) return { handled: false }
      const state = capture(input); assertCurrent(state)
      return context.crop?.(id, state.input) ?? { handled: false }
    })
    const selection = bindSheetImageSelection(studio, unitId, () => ({ sheetId: context.sheetId(), imageIds: context.active() ? selectedImageIds.value : [] }))
    onCleanup(() => { owner.dispose(); saver.dispose(); imageEditor.dispose(); cropper.dispose(); selection.dispose(); cancelPicker(); cancelEditor(); cancelSave() })
  }, { immediate: true, flush: 'sync' })
  watch([() => context.active(), () => context.workbook().id, () => context.sheetId(), () => context.userId()], () => {
    generation++; selectedImageId.value = undefined; cancelPicker(); cancelEditor(); cancelSave()
  }, { flush: 'sync' })
  onUnmounted(() => { generation++; cancelPicker(); cancelEditor(); cancelSave() })
  watch(() => context.workbook().sheets.find(sheet => sheet.id === context.sheetId())?.images?.map(image => image.id).join('\u0000'), () => {
    const ids = new Set(context.workbook().sheets.find(sheet => sheet.id === context.sheetId())?.images?.map(image => image.id) ?? [])
    selectedImageIds.value = selectedImageIds.value.filter(id => ids.has(id))
  })
  return { pending, editor, saving, pickerRequest, editorRequest, saveRequest, selectedImageId, selectedImageIds, imageError, busy, progress,
    runImageCommand, confirmPicker, cancelPicker, openImageEditor, confirmEditor, cancelEditor, confirmSave, cancelSave, afterClose, saveImage }
}
