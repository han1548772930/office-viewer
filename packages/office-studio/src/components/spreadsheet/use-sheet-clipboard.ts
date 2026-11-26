import { createPinia, getActivePinia, setActivePinia } from 'pinia'
import { computed, nextTick, onUnmounted, shallowRef, watch } from 'vue'
import { createBlankWorkbook, type CellRange, type SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { applyExcelWorkbookView, toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { captureSheetClipboard } from '../../features/sheet-workbench/sheet-clipboard-data.ts'
import { parseSheetClipboardHtml, parseSheetClipboardText, sheetClipboardContent, sheetClipboardIdFromHtml } from '../../features/sheet-workbench/sheet-clipboard-content.ts'
import { useSheetClipboardStore } from '../../features/sheet-workbench/sheet-clipboard-store.ts'
import { sheetClipboardCapabilities, sheetClipboardCommandCapability, type SheetClipboardCapabilitiesInput,
  type SheetClipboardCapabilityWorkbook, type SheetClipboardPermissionEvaluator } from '../../features/sheet-workbench/sheet-clipboard-capabilities.ts'
import { normalizeSheetClipboardContentAliases } from '../../features/sheet-workbench/sheet-clipboard-input.ts'
import { parseSheetAddress } from '../../features/sheet-workbench/sheet-workbench.ts'
import { applySheetClipboardPastePlan, assertSheetClipboardWorkbookChangesEditable, cloneSheetClipboardWorkbookResources,
  createSheetClipboardPastePlan, resolveSheetClipboardCommand, sheetClipboardPasteMode } from '../../features/sheet-workbench/sheet-clipboard.ts'
import type { SheetClipboardCommandInput, SheetClipboardCopyResult, SheetClipboardLocation, SheetClipboardPasteResult,
  SheetClipboardSnapshot, SheetClipboardWorkbookResources } from '../../features/sheet-workbench/sheet-clipboard-types.ts'
import type { SheetPasteMode } from '../../features/sheet-workbench/sheet-paste.ts'
import { captureSheetClipboardHost, type SheetClipboardHost, type SheetClipboardHostContext } from './sheet-clipboard-ui-host.ts'
import { sheetSystemClipboard, sheetSystemClipboardAvailable, type SheetSystemClipboard, type SheetSystemClipboardContent } from './sheet-system-clipboard.ts'
import { prepareSheetClipboardHtmlImages } from '../../features/sheet-images/sheet-image-html-preparation.ts'
import { captureSheetImageObjects, planPasteSheetImageObjects, sheetImageObjectClipboardContent, sheetImageObjectClipboardFromHtml, sheetImageObjectClipboardIdFromHtml, sheetImageObjectPasteRanges,
  type SheetImageObjectClipboard } from '../../features/sheet-images/sheet-image-object-clipboard.ts'
import { sheetClipboardObjectRange } from '../../features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { sheetImageClipboardBounds } from '../../features/sheet-images/sheet-image-clipboard.ts'
import { applySheetImagePlan, planSheetImageCommand, SHEET_IMAGE_COMMANDS } from '../../features/sheet-images/sheet-image-commands.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'

export interface SheetClipboardUiContext extends SheetClipboardHostContext {
  userId(): string
  finishEdit(): boolean
  prepareView(view: ExcelWorkbookView, previous: ExcelWorkbookView): ExcelWorkbookView
  applyView(view: ExcelWorkbookView): void
  selection(ranges: CellRange[]): void
  changed(options: { prepared: true; syncSelection?: boolean; current: () => boolean }): void
  focus(): void
  notify(value: { type: 'info' | 'warning' | 'error'; message: string }): void
  systemClipboard?: SheetSystemClipboard
  capabilitySelection?(): SheetClipboardCapabilitiesInput['selection']
  hidePasteOptions?: boolean | (() => boolean)
  exclusiveRange?(): boolean
  shapeTextEditing?(): boolean
  permission?: SheetClipboardPermissionEvaluator
  selectedImageIds?(): readonly string[]
  selectImageIds?(ids: readonly string[]): void
  pasteImageFiles?(files: readonly File[], input: SheetClipboardCommandInput & { signal?: AbortSignal }): Promise<boolean>
}

interface StandaloneRecentPaste {
  clipboard: SheetClipboardSnapshot
  target: SheetClipboardLocation
  ranges: CellRange[]
  mode: SheetPasteMode
  originalSheets: SpreadsheetSheet[]
  originalMetadata?: Record<string, unknown>
  metadataChanged: boolean
  originalResources?: SheetClipboardWorkbookResources
  view: ExcelWorkbookView
  signature: string
}
interface SheetClipboardUiInput {
  mode?: string
  value?: string
  type?: string
  text?: string
  html?: string
  clipboardId?: string
}
let nextStandaloneClipboard = 0
const stalePaste = (): SheetClipboardPasteResult => ({ accepted: false, changed: false, ranges: [], issues: [], reason: 'stale-target' })
const sameRange = (first: CellRange | undefined, second: CellRange | undefined): boolean => !!first && !!second
  && first.startRow === second.startRow && first.endRow === second.endRow && first.startCol === second.startCol && first.endCol === second.endCol
const nativeTextTarget = (event: Event): boolean => typeof Element !== 'undefined' && event.target instanceof Element
  && !!event.target.closest('input, textarea, [contenteditable]:not([contenteditable="false"])')
const resourceKeys = ['tables', 'pivotTables', 'charts', 'slicers'] as const

/** Browser events and Ribbon use the same domain commands. This controller owns no content undo stack. */
export function useSheetClipboard(context: SheetClipboardUiContext) {
  const standaloneUnitId = `sheet-clipboard-standalone-${++nextStandaloneClipboard}`
  let standaloneStore: ReturnType<typeof useSheetClipboardStore> | undefined
  let requestNumber = 0
  let pendingRead: number | undefined
  let imageRead: AbortController | undefined
  let imageWrite: AbortController | undefined
  let nativePasteMode: SheetPasteMode = 'all'
  const standaloneRecent = shallowRef<StandaloneRecentPaste>()
  const system = context.systemClipboard ?? sheetSystemClipboard
  function cancelImageRead(): void { imageRead?.abort(); imageRead = undefined; imageWrite?.abort(); imageWrite = undefined }
  async function writeClipboard(content: Parameters<SheetSystemClipboard['write']>[0], current: () => boolean): Promise<boolean> {
    imageWrite?.abort()
    const controller = new AbortController(); imageWrite = controller
    try { return await system.write(content, { signal: controller.signal, current }) }
    finally { if (imageWrite === controller) imageWrite = undefined }
  }
  async function prepareContent(host: SheetClipboardHost, content: SheetSystemClipboardContent | undefined, task: number): Promise<SheetSystemClipboardContent | undefined> {
    if (!content?.html || !/<img\b/i.test(content.html)) return content
    const token = sheetClipboardIdFromHtml(content.html)
    if (token && clipboardStore().getClipboard(token)) return content
    const imageToken = sheetImageObjectClipboardIdFromHtml(content.html)
    if (imageToken && (clipboardStore().getImageClipboard(imageToken) || sheetImageObjectClipboardFromHtml(content.html))) return content
    cancelImageRead()
    const controller = new AbortController()
    imageRead = controller; pendingRead = task
    try {
      const html = await prepareSheetClipboardHtmlImages(content.html, { signal: controller.signal, current: () => task === requestNumber && host.current() })
      return { ...content, html }
    } finally {
      if (imageRead === controller) imageRead = undefined
      if (pendingRead === task) pendingRead = undefined
    }
  }
  watch([() => context.active(), () => context.studio(), () => context.workbook().id, () => context.sheetId(), () => context.userId()], cancelImageRead, { flush: 'sync' })
  watch(() => context.studio(), (studio, _previous, onCleanup) => {
    const listener = studio?.events.addEvent('LifeCycleChanged', event => { if (event.stage === 'disposed') cancelImageRead() })
    onCleanup(() => { listener?.dispose(); cancelImageRead() })
  }, { immediate: true, flush: 'sync' })
  onUnmounted(cancelImageRead)

  function clipboardStore() {
    const studio = context.studio()
    if (studio && context.workbook().id) return studio.sheetClipboard
    if (!standaloneStore) {
      const previousPinia = getActivePinia()
      try { standaloneStore = useSheetClipboardStore(createPinia()) }
      finally { setActivePinia(previousPinia) }
    }
    return standaloneStore
  }
  const capabilityWorkbook = computed<SheetClipboardCapabilityWorkbook>(() => {
    const view = context.workbook(), studio = context.studio()
    const model = studio && view.id ? studio.getWorkbookModel(view.id) : undefined
    if (model) return model
    return { id: view.id ?? standaloneUnitId, tables: view.tables, sheets: view.sheets.map(sheet => ({
      id: sheet.id, visible: sheet.visible, rowCount: sheet.rowCount, columnCount: sheet.columnCount,
      worksheetProtection: sheet.worksheetProtection, rangeProtectionRules: sheet.rangeProtectionRules,
      filteredRows: sheet.filteredRows, autoFilter: sheet.autoFilter,
      mergedRanges: (sheet.merges ?? []).flatMap(address => {
        const parsed = parseSheetAddress(address, { rowCount: 1_048_576, columnCount: 16_384 })
        return parsed ? [{ startRow: parsed.range.startRow, endRow: parsed.range.endRow, startCol: parsed.range.startColumn, endCol: parsed.range.endColumn }] : []
      }),
    })) }
  })
  const imageCapabilities = computed(() => {
    const focusedIds = context.selectedImageIds?.() ?? [], imageClipboard = clipboardStore().imageClipboard
    if (!focusedIds.length && !imageClipboard) return {}
    const view = context.workbook(), studio = context.studio()
    const model = studio && view.id ? studio.getWorkbookModel(view.id) : applyExcelWorkbookView({ ...createBlankWorkbook(), id: view.id ?? standaloneUnitId }, view)
    const sheet = model?.sheets.find(sheet => sheet.id === context.sheetId()), selected = context.range()
    if (!model || !sheet) return {}
    return { imageClipboard, focusedImageRanges: (sheet.images ?? []).filter(image => focusedIds.includes(image.id))
      .map(image => sheetClipboardObjectRange(model, sheet.id, sheetImageClipboardBounds(model, sheet, image))),
      imagePasteRanges: imageClipboard && selected ? sheetImageObjectPasteRanges(model, sheet, selected, imageClipboard) : undefined }
  })
  function capabilityInput(clipboard: SheetClipboardSnapshot | undefined, systemClipboardAvailable: boolean,
    target?: SheetClipboardLocation, originalSheets?: readonly SpreadsheetSheet[], previousPasteMode?: SheetPasteMode): SheetClipboardCapabilitiesInput {
    const selected = context.range(), workbook = capabilityWorkbook.value
    const originals = new Map(originalSheets?.map(sheet => [sheet.id, sheet]))
    // Optional paste plans against the original geometry but must keep today's protection rules.
    const projected = originalSheets ? { ...workbook, sheets: workbook.sheets.map(sheet => {
      const original = originals.get(sheet.id)
      return original ? { ...sheet, mergedRanges: original.mergedRanges, autoFilter: original.autoFilter, filteredRows: original.autoFilter?.filteredRows } : sheet
    }) } : workbook
    return { workbook: projected, sheetId: target?.sheetId ?? context.sheetId(),
      selection: target ? { ranges: [target.range] } : context.capabilitySelection?.() ?? { ranges: selected ? [selected] : [] },
      userId: context.userId(), clipboard, systemClipboardAvailable, active: context.active(), editing: context.editing(),
      ...(clipboard ? {} : imageCapabilities.value),
      exclusiveRange: context.exclusiveRange?.(), shapeTextEditing: context.shapeTextEditing?.(), permission: context.permission, previousPasteMode }
  }
  const capabilities = computed(() => sheetClipboardCapabilities(capabilityInput(clipboardStore().clipboard, sheetSystemClipboardAvailable(system))))
  const recentPaste = () => context.studio() && context.workbook().id ? clipboardStore().recentPaste : standaloneRecent.value
  const pasteOptionsCapabilities = computed(() => {
    const recent = recentPaste()
    return recent ? sheetClipboardCapabilities(capabilityInput(recent.clipboard, false, recent.target, recent.originalSheets, recent.mode)) : undefined
  })
  const copyRange = computed(() => {
    const mark = clipboardStore().copyMark
    return context.active() && !context.editing() && mark?.unitId === (context.workbook().id ?? standaloneUnitId)
      && mark.sheetId === context.sheetId() ? cloneModel(mark.range) : undefined
  })
  const pasteOptions = computed(() => {
    if (typeof context.hidePasteOptions === 'function' ? context.hidePasteOptions() : context.hidePasteOptions) return undefined
    if (!context.active() || context.editing()) return undefined
    const studio = context.studio(), unitId = context.workbook().id ?? standaloneUnitId
    const recent = studio && context.workbook().id ? studio.sheetClipboard.recentPaste : standaloneRecent.value
    if (!recent || recent.target.unitId !== unitId || recent.target.sheetId !== context.sheetId()) return undefined
    if (studio && context.workbook().id) {
      const actual = studio.sheetClipboard.recentPaste
      if (!actual || studio.getWorkbookRevision(unitId) !== actual.contentRevision
        || !studio.editor.history.canAmend(unitId, actual.savepoint)) return undefined
    } else {
      const actual = standaloneRecent.value
      if (!actual || actual.view !== context.workbook() || actual.signature !== JSON.stringify(context.workbook())) return undefined
    }
    const range = recent.ranges.at(-1)
    return range ? { range: cloneModel(range), mode: recent.mode } : undefined
  })

  function begin(): SheetClipboardHost | undefined {
    const owner = context.studio(), unitId = context.workbook().id, sheetId = context.sheetId()
    if (!context.active() || context.shapeTextEditing?.() || !context.finishEdit() || context.studio() !== owner
      || context.workbook().id !== unitId || context.sheetId() !== sheetId) return undefined
    return captureSheetClipboardHost(context, standaloneUnitId)
  }
  function report(result: SheetClipboardPasteResult): void {
    const issue = result.issues.find(item => item.severity === 'stop') ?? result.issues[0]
    if (issue) context.notify({ type: issue.severity === 'stop' ? 'error' : issue.severity === 'warning' ? 'warning' : 'info', message: `${issue.title}：${issue.message}` })
    else if (!result.accepted && result.reason) {
      const messages = {
        'empty-clipboard': '剪贴板中没有可粘贴的文本或表格内容。',
        'stale-target': '粘贴目标已更改，本次粘贴已停止。',
        'stale-source': '剪切源内容已更改，请重新剪切。',
        'stale-paste': '上次粘贴之后内容已更改，不能切换其粘贴方式。',
        'cross-workbook-cut': '跨工作簿剪切移动尚未接入，源内容已保留。',
        validation: '目标区域的数据验证拒绝了粘贴。',
      }
      context.notify({ type: 'warning', message: messages[result.reason] })
    }
  }
  function focusWhenCurrent(host: SheetClipboardHost, current: () => boolean, task: number): void {
    void nextTick(() => { if (task === requestNumber && host.focusUnchanged() && current()) context.focus() })
  }
  function assertUiPermission(host: SheetClipboardHost, commandId: string, snapshot?: SheetClipboardSnapshot, mode?: SheetPasteMode,
    target?: SheetClipboardLocation, originalSheets?: readonly SpreadsheetSheet[], previousPasteMode?: SheetPasteMode): void {
    if (!context.permission) return
    const input = capabilityInput(snapshot, false, target ?? { unitId: host.unitId, sheetId: host.sheetId, range: host.range }, originalSheets, previousPasteMode)
    const entry = sheetClipboardCommandCapability(sheetClipboardCapabilities(input), commandId, mode)
    if (entry?.disabled) throw new Error(entry.reason ?? '当前用户没有执行此剪贴板操作的权限。')
  }
  function copy(host: SheetClipboardHost, commandId: string, formulaOnly: boolean, cut: boolean): SheetClipboardCopyResult {
    if (!host.current()) return { accepted: false, plain: '', html: '' }
    assertUiPermission(host, commandId)
    if (!host.current()) return { accepted: false, plain: '', html: '' }
    const imageIds = !formulaOnly ? context.selectedImageIds?.() : undefined
    if (imageIds?.length) {
      const ranges = (host.sheet.images ?? []).filter(image => imageIds.includes(image.id))
        .map(image => sheetClipboardObjectRange(host.model, host.sheetId, sheetImageClipboardBounds(host.model, host.sheet, image)))
      assertObjectPermissions(host, commandId, ranges, cut)
    }
    if (host.studio) {
      const result = host.studio.executeCommandSync<SheetClipboardCommandInput, SheetClipboardCopyResult>(commandId, { ...host.request,
        ...(imageIds?.length ? { imageIds } : {}) })
      if (result.changed) {
        publishStudio(host, { accepted: result.accepted, changed: true, ranges: [], issues: [] }, requestNumber)
        context.selectImageIds?.([])
      }
      return result
    }
    if (imageIds?.length) {
      const snapshot = captureSheetImageObjects(host.model, host.sheet, imageIds, context.userId(), cut ? 'cut' : 'copy')
      const ranges = snapshot.entries.map(entry => sheetClipboardObjectRange(host.model, host.sheetId, entry.bounds))
      assertObjectPermissions(host, commandId, ranges, cut)
      const content = sheetImageObjectClipboardContent(snapshot)
      if (cut) {
        const candidate = cloneModel(host.model), sheet = candidate.sheets.find(sheet => sheet.id === host.sheetId)!
        const plan = planSheetImageCommand(sheet, SHEET_IMAGE_COMMANDS.remove, { imageIds }, { row: host.range.startRow, column: host.range.startCol })
        applySheetImagePlan(sheet, plan)
        if (!publishStandaloneObjects(host, candidate, requestNumber)) return { accepted: false, plain: '', html: '' }
        context.selectImageIds?.([])
      }
      clipboardStore().setImageClipboard(snapshot); standaloneRecent.value = undefined
      return { accepted: true, changed: cut, imageClipboard: snapshot, ...content }
    }
    const snapshot = captureSheetClipboard(host.model, { unitId: host.unitId, sheetId: host.sheetId, range: host.range },
      { copyType: cut ? 'cut' : 'copy', formulaOnly, userId: context.userId() })
    const content = sheetClipboardContent(snapshot)
    if (!host.current()) return { accepted: false, plain: '', html: '' }
    clipboardStore().setClipboard(snapshot); standaloneRecent.value = undefined
    return { accepted: true, clipboard: snapshot, ...content }
  }
  function assertObjectPermissions(host: SheetClipboardHost, commandId: string, ranges: readonly CellRange[], edit: boolean): void {
    if (!canViewWorksheetProtection(host.sheet, context.userId()) || !canViewProtectedRanges(host.sheet, ranges, context.userId())) throw new Error('没有读取目标图片的权限。')
    if (edit && (!canEditWorksheetProtection(host.sheet, context.userId()) || !canEditProtectedRanges(host.sheet, ranges, context.userId()))) throw new Error('没有编辑目标图片的权限。')
    for (const range of ranges) assertUiPermission(host, commandId, undefined, undefined, { unitId: host.unitId, sheetId: host.sheetId, range })
  }
  function publishStandaloneObjects(host: SheetClipboardHost, model: SheetClipboardHost['model'], task: number): boolean {
    const next = context.prepareView(toExcelWorkbookView(model), cloneModel(host.view))
    if (!host.current() || task !== requestNumber) return false
    if (host.view.id === undefined) delete next.id
    else next.id = host.view.id
    const signature = JSON.stringify(next)
    context.applyView(next)
    const published = context.workbook(), current = () => task === requestNumber && host.sameTarget()
      && context.workbook() === published && JSON.stringify(published) === signature
    if (current()) context.changed({ prepared: true, syncSelection: false, current })
    return true
  }
  function copyCurrent(host: SheetClipboardHost, result: SheetClipboardCopyResult, task: number): boolean {
    if (task !== requestNumber || !host.sameTarget()) return false
    if (!result.changed) return host.current()
    if (host.studio) return host.studio.getWorkbookRevision(host.unitId) === host.revision! + 1
      && host.studio.editor.history.captureSavepoint(host.unitId).generation === host.historyGeneration
    return !!result.imageClipboard && clipboardStore().imageClipboard?.id === result.imageClipboard.id
  }
  function imageSource(content: SheetSystemClipboardContent | undefined, clipboardId?: string): SheetImageObjectClipboard | undefined {
    if (content === undefined) return clipboardStore().getImageClipboard(clipboardId)
    const token = content.html ? sheetImageObjectClipboardIdFromHtml(content.html) : undefined
    return token && (clipboardId === undefined || token === clipboardId)
      ? clipboardStore().getImageClipboard(token) ?? sheetImageObjectClipboardFromHtml(content.html!) : undefined
  }
  async function pasteFiles(host: SheetClipboardHost, files: readonly File[], task: number): Promise<SheetClipboardPasteResult> {
    if (!host.current() || task !== requestNumber) return stalePaste()
    if (!context.pasteImageFiles) throw new Error('当前宿主没有注册图片文件粘贴接口。')
    const controller = new AbortController()
    imageRead = controller; pendingRead = task
    try {
      const changed = await context.pasteImageFiles(files, { ...host.request, signal: controller.signal })
      return { accepted: changed, changed, ranges: [], issues: [] }
    } finally {
      if (imageRead === controller) imageRead = undefined
      if (pendingRead === task) pendingRead = undefined
    }
  }
  function source(content: SheetSystemClipboardContent | undefined, clipboardId?: string): SheetClipboardSnapshot | undefined {
    const store = clipboardStore()
    if (content === undefined) return store.getClipboard(clipboardId)
    const cached = content.html ? store.getClipboard(sheetClipboardIdFromHtml(content.html)) : undefined
    // Only an actual matching HTML token authorizes using the complete internal model.
    if (cached && content.html && sheetClipboardIdFromHtml(content.html) === cached.id
      && (clipboardId === undefined || clipboardId === cached.id)) return cached
    return content.html ? parseSheetClipboardHtml(content.html) ?? parseSheetClipboardText(content.text ?? '')
      : parseSheetClipboardText(content.text ?? '')
  }
  function publishStudio(host: SheetClipboardHost, result: SheetClipboardPasteResult, task: number): void {
    const studio = host.studio
    if (!studio || !result.accepted) return
    if (!result.changed && !host.current()) return
    const expectedRevision = host.revision! + (result.changed ? 1 : 0)
    const model = studio.getWorkbookModel(host.unitId)
    if (!model) return
    const current = () => host.sameTarget() && task === requestNumber && studio.getWorkbookModel(host.unitId) === model
      && studio.getWorkbookRevision(host.unitId) === expectedRevision
      && studio.editor.history.captureSavepoint(host.unitId).generation === host.historyGeneration
    if (!current()) return
    if (result.changed) {
      const next = toExcelWorkbookView(model)
      if (!current()) return
      const expectedView = JSON.stringify(next)
      context.applyView(next)
      if (!current() || JSON.stringify(context.workbook()) !== expectedView) return
      const published = context.workbook()
      context.changed({ prepared: true, syncSelection: false,
        current: () => current() && context.workbook() === published && JSON.stringify(context.workbook()) === expectedView })
      if (!current() || context.workbook() !== published || JSON.stringify(context.workbook()) !== expectedView) return
    }
    if (!current()) return
    const pastedRange = result.ranges.at(-1)
    if (sameRange(context.range(), pastedRange)) {
      const selection = JSON.stringify(context.range()), published = context.workbook(), signature = JSON.stringify(published)
      focusWhenCurrent(host, () => current() && context.workbook() === published && JSON.stringify(context.workbook()) === signature
        && JSON.stringify(context.range()) === selection, task)
    }
  }
  function pasteStandalone(host: SheetClipboardHost, snapshot: SheetClipboardSnapshot | undefined, mode: SheetPasteMode,
    task: number, recent?: StandaloneRecentPaste): SheetClipboardPasteResult {
    if (!snapshot) return { accepted: false, changed: false, ranges: [], issues: [], reason: 'empty-clipboard' }
    if (task !== requestNumber || !host.current()) return stalePaste()
    const model = cloneModel(host.model)
    if (recent) {
      const originals = new Map(recent.originalSheets.map(sheet => [sheet.id, cloneModel(sheet)]))
      model.sheets = model.sheets.map(sheet => originals.get(sheet.id) ?? sheet)
      if (recent.metadataChanged) {
        if (recent.originalMetadata === undefined) delete model.metadata
        else model.metadata = cloneModel(recent.originalMetadata)
      }
      if (recent.originalResources) for (const key of resourceKeys) {
        const value = recent.originalResources[key]
        if (value === undefined) delete model[key]
        else Object.assign(model, { [key]: cloneModel(value) })
      }
    }
    if (snapshot.formulaOnly && mode === 'all') mode = 'formulas'
    const target = recent?.target ?? { unitId: host.unitId, sheetId: host.sheetId, range: host.range }
    const plan = createSheetClipboardPastePlan(model, target, snapshot, mode, { userId: context.userId(), region: context.region() })
    if (!plan.result.accepted) return plan.result
    const originalSheets = new Map((recent?.originalSheets ?? []).map(sheet => [sheet.id, cloneModel(sheet)]))
    for (const sheet of model.sheets) if (!originalSheets.has(sheet.id) && plan.sheetChanges.some(changed => changed.id === sheet.id)) {
      originalSheets.set(sheet.id, cloneModel(sheet))
    }
    const metadata = recent ? recent.originalMetadata : model.metadata
    const originalMetadata = metadata === undefined ? undefined : cloneModel(metadata)
    const originalResources = recent?.originalResources ?? (plan.workbookResources ? cloneSheetClipboardWorkbookResources(model) : undefined)
    applySheetClipboardPastePlan(model, plan)
    if (recent) assertSheetClipboardWorkbookChangesEditable(host.model, model, context.userId())
    // Optional paste can restore the original content even when its new plan changes no original cell.
    const result = { ...plan.result, changed: JSON.stringify(model) !== JSON.stringify(host.model) }
    if (!result.changed) {
      if (recent && host.current() && task === requestNumber) standaloneRecent.value = { ...recent, ranges: cloneModel(result.ranges), mode }
      if (host.current() && task === requestNumber) clipboardStore().consumeCut(snapshot.id)
      if (sameRange(context.range(), result.ranges.at(-1))) focusWhenCurrent(host, () => host.current() && host.selectionUnchanged(), task)
      return result
    }
    const next = context.prepareView(toExcelWorkbookView(model), cloneModel(host.view))
    if (!host.current() || task !== requestNumber) return stalePaste()
    if (host.view.id === undefined) delete next.id
    else next.id = host.view.id
    const preparedSignature = JSON.stringify(next)
    context.applyView(next)
    const published = context.workbook()
    const current = () => host.sameTarget() && task === requestNumber && context.workbook() === published
      && JSON.stringify(context.workbook()) === preparedSignature
    if (!current()) return result
    if (host.selectionUnchanged()) context.selection(result.ranges)
    if (!current()) return result
    context.changed({ prepared: true, current })
    if (current()) {
      clipboardStore().consumeCut(snapshot.id)
      standaloneRecent.value = { clipboard: cloneModel(snapshot), target: cloneModel(target), ranges: cloneModel(result.ranges),
        mode, originalSheets: [...originalSheets.values()], originalMetadata, metadataChanged: !!recent?.metadataChanged || plan.metadataChanged,
        originalResources,
        view: context.workbook(), signature: JSON.stringify(context.workbook()) }
      if (sameRange(context.range(), result.ranges.at(-1))) {
        const selection = JSON.stringify(context.range())
        focusWhenCurrent(host, () => current() && JSON.stringify(context.range()) === selection, task)
      }
    }
    return result
  }
  function paste(host: SheetClipboardHost, commandId: string, mode: SheetPasteMode, task: number,
    content?: SheetSystemClipboardContent, clipboardId?: string, value?: string, explicitClipboardId?: string,
    type?: SheetClipboardCommandInput['type']): SheetClipboardPasteResult {
    if (!host.current() || task !== requestNumber) return stalePaste()
    const objectClipboard = imageSource(content, content === undefined ? clipboardId : explicitClipboardId)
    if (objectClipboard) {
      if (mode !== 'all' && mode !== 'exceptBorders') return { accepted: false, changed: false, ranges: [], issues: [], reason: 'empty-clipboard' }
      if (objectClipboard.source.userId !== context.userId()) throw new Error('图片剪贴板属于其他用户。')
      if (!host.studio && objectClipboard.source.unitId === host.unitId) {
        const sourceSheet = host.model.sheets.find(sheet => sheet.id === objectClipboard.source.sheetId)
        if (sourceSheet) {
          const ranges = objectClipboard.entries.map(entry => {
            const image = sourceSheet.images?.find(image => image.id === entry.image.id)
            return sheetClipboardObjectRange(host.model, sourceSheet.id, image ? sheetImageClipboardBounds(host.model, sourceSheet, image) : entry.bounds)
          })
          if (!canViewWorksheetProtection(sourceSheet, context.userId()) || !canViewProtectedRanges(sourceSheet, ranges, context.userId())) throw new Error('没有读取图片剪贴板来源的权限。')
          for (const point of ['view', 'copy'] as const) {
            const permission = context.permission?.({ point, phase: 'source', workbook: host.model, sheet: sourceSheet,
              unitId: host.unitId, sheetId: sourceSheet.id, ranges, userId: context.userId() })
            if (permission === false || permission && typeof permission === 'object' && !permission.allowed) throw new Error(typeof permission === 'object' && permission.reason || '没有读取图片剪贴板来源的权限。')
          }
        }
      }
      const plan = planPasteSheetImageObjects(host.model, host.sheet, host.range, objectClipboard)
      assertObjectPermissions(host, commandId, plan.ranges, true)
      if (!host.current() || task !== requestNumber) return stalePaste()
      let result: SheetClipboardPasteResult
      if (host.studio) {
        result = host.studio.executeCommandSync<SheetClipboardCommandInput, SheetClipboardPasteResult>('sheet.command.paste', {
          ...host.request, mode, clipboardId: objectClipboard.id, ...(content === undefined ? {} : { html: content.html, text: content.text }),
        })
        publishStudio(host, result, task)
      } else {
        if (objectClipboard.source.userId !== context.userId()) throw new Error('图片剪贴板属于其他用户。')
        const candidate = cloneModel(host.model), sheet = candidate.sheets.find(sheet => sheet.id === host.sheetId)!
        const imagePlan = planSheetImageCommand(sheet, SHEET_IMAGE_COMMANDS.insertPrepared, { images: plan.images }, { row: host.range.startRow, column: host.range.startCol })
        applySheetImagePlan(sheet, imagePlan)
        for (const range of plan.ranges) {
          if (sheet.rowCount !== undefined) sheet.rowCount = Math.max(sheet.rowCount, range.endRow + 1)
          if (sheet.columnCount !== undefined) sheet.columnCount = Math.max(sheet.columnCount, range.endCol + 1)
        }
        const changed = publishStandaloneObjects(host, candidate, task)
        result = { accepted: changed, changed, ranges: plan.ranges, issues: [], imageIds: changed ? plan.images.map(image => image.id) : [] }
      }
      if (result.changed && host.sameTarget() && task === requestNumber) context.selectImageIds?.(result.imageIds ?? [])
      return result
    }
    if (context.permission) {
      const snapshot = source(content, content === undefined ? clipboardId : explicitClipboardId)
      if (snapshot) assertUiPermission(host, commandId, snapshot, mode)
      if (!host.current() || task !== requestNumber) return stalePaste()
    }
    if (host.studio) {
      const requestClipboardId = content === undefined ? clipboardId : explicitClipboardId
      const request: SheetClipboardCommandInput = { ...host.request, mode,
        ...(content === undefined ? {} : { text: content.text ?? '', html: content.html ?? '' }),
        ...(requestClipboardId ? { clipboardId: requestClipboardId } : {}), ...(value ? { value } : {}), ...(type ? { type } : {}) }
      const result = host.studio.executeCommandSync<SheetClipboardCommandInput, SheetClipboardPasteResult>(commandId, request)
      publishStudio(host, result, task)
      return result
    }
    return pasteStandalone(host, source(content, content === undefined ? clipboardId : explicitClipboardId), mode, task)
  }
  function cancelCopyMark(): boolean {
    cancelImageRead()
    const task = ++requestNumber
    pendingRead = undefined; standaloneRecent.value = undefined
    const studio = context.studio(), view = context.workbook(), unitId = view.id, sheetId = context.sheetId()
    try {
      if (studio && unitId) return studio.executeCommandSync<SheetClipboardCommandInput, boolean>('sheet.operation.cancel-copy', { unitId, sheetId })
      clipboardStore().cancelCopyMark()
      return true
    } catch (error) {
      if (task === requestNumber && context.active() && context.studio() === studio
        && context.workbook() === view && context.sheetId() === sheetId) {
        context.notify({ type: 'error', message: error instanceof Error ? error.message : '取消复制失败。' })
      }
      return false
    }
  }
  function commandInput(params: unknown): SheetClipboardUiInput {
    if (params === undefined) return {}
    if (typeof params === 'string') return { value: params }
    if (!params || typeof params !== 'object' || Array.isArray(params)) throw new TypeError('剪贴板参数无效。')
    const source = params as Record<string, unknown>
    const input: SheetClipboardUiInput = normalizeSheetClipboardContentAliases(source)
    for (const key of ['mode', 'value', 'type', 'clipboardId'] as const) {
      const value = source[key]
      if (value === undefined) continue
      if (typeof value !== 'string' || key === 'clipboardId' && !value) throw new TypeError('剪贴板参数无效。')
      input[key] = value
    }
    return input
  }
  async function runClipboardCommand(commandId: string, params?: unknown): Promise<SheetClipboardCopyResult | SheetClipboardPasteResult | boolean | undefined> {
    const action = resolveSheetClipboardCommand(commandId)
    if (!action) return undefined
    if (action.action === 'cancel-copy') return cancelCopyMark()
    cancelImageRead()
    const task = ++requestNumber, host = begin()
    if (!host) return undefined
    try {
      if (action.action === 'copy' || action.action === 'cut') {
        const result = copy(host, action.action === 'cut' ? 'sheet.command.cut' : action.formulaOnly ? 'sheet.command.copy-formulas' : 'sheet.command.copy', !!action.formulaOnly, action.action === 'cut')
        if (result.accepted && copyCurrent(host, result, task)) {
          try {
            if (!await writeClipboard({ ...result, image: result.imageClipboard?.entries.length === 1 ? result.imageClipboard.entries[0]!.image.resource : undefined },
              () => copyCurrent(host, result, task))
              && host.sameTarget() && task === requestNumber) context.notify({ type: 'info', message: '系统剪贴板不可用，内容已保留在工作台内。' })
          } catch { if (host.sameTarget() && task === requestNumber) context.notify({ type: 'info', message: '系统剪贴板写入失败，内容已保留在工作台内。' }) }
          focusWhenCurrent(host, () => copyCurrent(host, result, task) && host.selectionUnchanged(), task)
        }
        return result
      }
      const input = commandInput(params)
      const mode = sheetClipboardPasteMode(input.mode, input.value, action.mode, input.type)
      const type = input.type as SheetClipboardCommandInput['type']
      if (action.action === 'optional-paste') {
        let result: SheetClipboardPasteResult
        const recent = recentPaste()
        if (recent && context.permission) {
          // The same prediction used by the menu checks both new writes and the previous mode's restoration.
          assertUiPermission(host, 'sheet.command.paste', recent.clipboard, mode, recent.target, recent.originalSheets, recent.mode)
          if (!host.current() || task !== requestNumber) return stalePaste()
        }
        if (host.studio) {
          result = host.studio.executeCommandSync<SheetClipboardCommandInput, SheetClipboardPasteResult>('sheet.command.optional-paste', {
            ...host.request, mode, ...(input.value ? { value: input.value } : {}), ...(type ? { type } : {}),
            ...(input.clipboardId ? { clipboardId: input.clipboardId } : {}),
          })
          publishStudio(host, result, task)
        } else {
          const standalone = standaloneRecent.value
          result = standalone && standalone.view === host.view && standalone.signature === JSON.stringify(host.view)
            && (input.clipboardId === undefined || input.clipboardId === standalone.clipboard.id)
            ? pasteStandalone(host, standalone.clipboard, mode, task, standalone)
            : { accepted: false, changed: false, ranges: [], issues: [], reason: 'stale-paste' }
        }
        if (host.sameTarget()) report(result)
        return result
      }
      const clipboardId = input.clipboardId ?? clipboardStore().imageClipboard?.id ?? clipboardStore().clipboard?.id
      let content: SheetSystemClipboardContent | undefined
      if (input.text !== undefined || input.html !== undefined) content = { text: input.text, html: input.html }
      else if (input.clipboardId === undefined && sheetSystemClipboardAvailable(system)) {
        pendingRead = task
        try { content = await system.read() }
        catch { content = undefined }
        finally { if (pendingRead === task) pendingRead = undefined }
      }
      if (!host.current() || task !== requestNumber) return stalePaste()
      if (content?.files?.length && !imageSource(content, input.clipboardId) && (mode === 'all' || mode === 'exceptBorders')) return pasteFiles(host, content.files, task)
      if (content?.html && /<img\b/i.test(content.html)) content = await prepareContent(host, content, task)
      if (!host.current() || task !== requestNumber) return stalePaste()
      const result = paste(host, commandId, mode, task, content, clipboardId, input.value, input.clipboardId, type)
      if (host.sameTarget()) report(result)
      return result
    } catch (error) {
      if (host.sameTarget() && task === requestNumber) context.notify({ type: 'error', message: error instanceof Error ? error.message : '剪贴板操作失败。' })
      return undefined
    }
  }
  function handleCopy(event: ClipboardEvent, cut = false): void {
    if (event.defaultPrevented || nativeTextTarget(event)) return
    event.preventDefault(); event.stopPropagation()
    cancelImageRead()
    const task = ++requestNumber, host = begin()
    if (!host) return
    try {
      const result = copy(host, cut ? 'sheet.command.cut' : 'sheet.command.copy', false, cut)
      if (result.accepted && copyCurrent(host, result, task)) {
        event.clipboardData?.setData('text/plain', result.plain)
        event.clipboardData?.setData('text/html', result.html)
        if (result.imageClipboard?.entries.length === 1) void writeClipboard({ plain: result.plain, html: result.html,
          image: result.imageClipboard.entries[0]!.image.resource }, () => copyCurrent(host, result, task)).catch(() => {
          if (copyCurrent(host, result, task)) context.notify({ type: 'info', message: '原生图片剪贴板写入失败，图片已保留在工作台和 HTML 剪贴板中。' })
        })
      }
    } catch (error) { if (host.sameTarget()) context.notify({ type: 'error', message: error instanceof Error ? error.message : '复制失败。' }) }
  }
  function handlePaste(event: ClipboardEvent): void {
    if (event.defaultPrevented || nativeTextTarget(event)) return
    event.preventDefault(); event.stopPropagation()
    const mode = nativePasteMode; nativePasteMode = 'all'
    cancelImageRead()
    const task = ++requestNumber, host = begin()
    if (!host) return
    const failed = (error: unknown) => {
      if (host.sameTarget() && task === requestNumber && !(error instanceof DOMException && error.name === 'AbortError')) {
        context.notify({ type: 'error', message: error instanceof Error ? error.message : '粘贴失败。' })
      }
    }
    try {
      const content: SheetSystemClipboardContent | undefined = event.clipboardData ? { text: event.clipboardData.getData('text/plain'), html: event.clipboardData.getData('text/html'),
        files: Array.from(event.clipboardData.files).filter(file => file.type.startsWith('image/')) } : undefined
      const apply = (prepared: SheetSystemClipboardContent | undefined) => {
        const result = paste(host, 'sheet.command.paste', mode, task, prepared, clipboardStore().imageClipboard?.id ?? clipboardStore().clipboard?.id)
        if (host.sameTarget() && task === requestNumber) report(result)
      }
      if (content?.files?.length && !imageSource(content) && (mode === 'all' || mode === 'exceptBorders')) void pasteFiles(host, content.files, task).catch(failed)
      else if (content?.html && /<img\b/i.test(content.html)) void prepareContent(host, content, task).then(apply).catch(failed)
      else apply(content)
    } catch (error) { failed(error) }
  }
  function handleClipboardKeydown(event: KeyboardEvent): boolean {
    if (!context.active() || event.defaultPrevented || event.isComposing || event.keyCode === 229 || nativeTextTarget(event)) return false
    if (event.key === 'Escape' && !context.editing() && (clipboardStore().copyMark || pasteOptions.value || standaloneRecent.value || pendingRead !== undefined)) {
      event.preventDefault(); cancelCopyMark(); return true
    }
    if ((event.ctrlKey || event.metaKey) && ['c', 'x', 'v'].includes(event.key.toLowerCase())) {
      nativePasteMode = event.key.toLowerCase() === 'v' && event.shiftKey ? 'values' : 'all'
      return true // Browser dispatches the actual copy/cut/paste event; keydown must keep its default action.
    }
    nativePasteMode = 'all'
    return false
  }
  return { capabilities, copyRange, pasteOptions, pasteOptionsCapabilities, runClipboardCommand, handleCopy, handlePaste, handleClipboardKeydown, cancelCopyMark }
}
