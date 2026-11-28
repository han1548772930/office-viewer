import { computed, nextTick, onScopeDispose, shallowRef, toRaw, watch, type Ref } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { Disposable } from '../../api/events.ts'
import type { CellRange, SheetConditionalFormatRule, SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { planSheetConditionalFormatsCommand, SET_SHEET_CONDITIONAL_FORMATS_COMMAND,
  type SetSheetConditionalFormatsInput, type SheetConditionalFormatSubmitResult } from '../../features/conditional-formatting/conditional-format-command.ts'
import type { ConditionalFormatPanelInitialType } from '../../features/conditional-formatting/quick-conditional-formatting.ts'

type ConditionalSheet = Pick<SpreadsheetSheet, 'id' | 'conditionalFormats' | 'worksheetProtection' | 'rangeProtectionRules'>
interface ConditionalWorkbenchState {
  studio?: OfficeStudio; model: ExcelWorkbookView; sheet?: ConditionalSheet; sheetId: string; userId: string
  active: boolean; blocked: boolean; inputSessionId?: string; ranges: readonly CellRange[]; defaultRange: string
}
interface ConditionalWorkbenchOptions {
  open: Ref<boolean>
  state(): ConditionalWorkbenchState
  refresh(studio: OfficeStudio, unitId: string, current: () => boolean): void
  applyStandalone(rules: readonly SheetConditionalFormatRule[], current: () => boolean, expectEcho: (snapshot: ExcelWorkbookView) => void): boolean
  focus(): void
  canRestoreFocus(): boolean
}
interface ConditionalPanelOwner {
  key: number; studio?: OfficeStudio; unitId?: string; sheetId: string; userId: string; inputSessionId?: string
  source: object; revision?: number; historyGeneration?: number; standaloneSnapshot?: string
  standaloneEcho?: object; standaloneEchoSnapshot?: string
  rules: SheetConditionalFormatRule[]; defaultRange: string; ranges: CellRange[]; initialType?: ConditionalFormatPanelInitialType
  setOpen(open: boolean): void; closed(): void; submit(rules: readonly SheetConditionalFormatRule[]): SheetConditionalFormatSubmitResult
  canEditRule(rule: SheetConditionalFormatRule): boolean; canApplyRules(rules: readonly SheetConditionalFormatRule[]): boolean
  preview(ranges: readonly CellRange[]): void
}

/** A drawer owns one input/model/history session. External edits invalidate it rather than merge an old full list. */
export function useSheetConditionalFormattingWorkbench(options: ConditionalWorkbenchOptions) {
  const context = shallowRef<ConditionalPanelOwner>(), highlighted = shallowRef<readonly CellRange[]>([])
  let sequence = 0, stopped = false, submitting = false
  function hostCurrent(owner: ConditionalPanelOwner): boolean {
    const state = options.state(), { studio, unitId } = owner
    return !stopped && owner.key === sequence && state.active && state.studio === studio && state.model.id === unitId
      && state.sheetId === owner.sheetId && state.sheet?.id === owner.sheetId && state.userId === owner.userId
      && state.inputSessionId === owner.inputSessionId && (!studio || !studio.isDisposed && unitId !== undefined
        && studio.activeWorkbookId === unitId && studio.editor.state.activeKind === 'workbook'
        && studio.getWorkbookModel(unitId)?.activeSheetId === owner.sheetId)
  }
  function current(owner: ConditionalPanelOwner): boolean {
    if (!hostCurrent(owner)) return false
    const { studio, unitId } = owner
    return studio && unitId !== undefined ? studio.getWorkbookModel(unitId) === owner.source
      && studio.getWorkbookRevision(unitId) === owner.revision && studio.currentUserId === owner.userId
      && studio.editor.history.captureSavepoint(unitId).generation === owner.historyGeneration
      : toRaw(options.state().model) === owner.source && JSON.stringify(options.state().model) === owner.standaloneSnapshot
  }
  function liveSheet(owner: ConditionalPanelOwner): ConditionalSheet | undefined {
    return owner.studio && owner.unitId ? owner.studio.getWorkbookModel(owner.unitId)?.sheets.find(sheet => sheet.id === owner.sheetId) : options.state().sheet
  }
  function rangesAllowed(owner: ConditionalPanelOwner, ranges: readonly CellRange[], edit: boolean): boolean {
    if (!current(owner)) return false
    const sheet = liveSheet(owner)
    return !!sheet && canViewWorksheetProtection(sheet, owner.userId) && canViewProtectedRanges(sheet, ranges, owner.userId)
      && (!edit || canEditWorksheetProtection(sheet, owner.userId) && canEditProtectedRanges(sheet, ranges, owner.userId))
  }
  function rulesAllowed(owner: ConditionalPanelOwner, rules: readonly SheetConditionalFormatRule[]): boolean {
    if (options.state().blocked || !current(owner)) return false
    const sheet = liveSheet(owner)
    if (!sheet) return false
    try { return rangesAllowed(owner, planSheetConditionalFormatsCommand(sheet, rules).ranges, true) } catch { return false }
  }
  function invalidate(): void {
    sequence++; options.open.value = false; context.value = undefined; highlighted.value = []
  }
  function inspect(): void {
    const owner = context.value
    if (owner && !submitting && !current(owner)) invalidate()
  }
  function receiveView(value: ExcelWorkbookView, apply: () => void): void {
    const owner = context.value
    if (!owner || owner.studio || owner.standaloneEcho !== toRaw(value) || !current(owner)) { apply(); return }
    // Only the exact snapshot emitted by this submission is a model echo. An equal external replacement still ends the owner.
    owner.standaloneEcho = undefined
    const wasSubmitting = submitting
    submitting = true
    try {
      apply()
      if (context.value === owner && hostCurrent(owner) && JSON.stringify(options.state().model) === owner.standaloneEchoSnapshot) {
        owner.source = toRaw(options.state().model)
      } else if (context.value === owner) invalidate()
    } finally { submitting = wasSubmitting; inspect() }
  }
  function setOpen(owner: ConditionalPanelOwner, value: boolean): void {
    if (context.value !== owner || value) return
    options.open.value = false; highlighted.value = []
  }
  function closed(owner: ConditionalPanelOwner): void {
    if (context.value !== owner || options.open.value) return
    context.value = undefined; highlighted.value = []
    void nextTick(() => {
      if (!options.open.value && !context.value && current(owner) && options.canRestoreFocus()) options.focus()
    })
  }
  function submit(owner: ConditionalPanelOwner, rules: readonly SheetConditionalFormatRule[]): SheetConditionalFormatSubmitResult {
    if (context.value !== owner || !options.open.value || options.state().blocked || !current(owner)) return { accepted: false, error: '条件格式目标已更改，请重新打开面板后编辑。' }
    submitting = true
    let changed = false
    try {
      const { studio, unitId } = owner
      if (studio && unitId) {
        changed = studio.executeCommandSync<SetSheetConditionalFormatsInput, boolean>(SET_SHEET_CONDITIONAL_FORMATS_COMMAND,
          { unitId, sheetId: owner.sheetId, rules, host: 'workbench', expectedRevision: owner.revision,
            expectedHistoryGeneration: owner.historyGeneration, expectedUserId: owner.userId })
        const source = studio.getWorkbookModel(unitId), revision = (owner.revision ?? 0) + (changed ? 1 : 0)
        const completed = { ...owner, source: source ?? owner.source, revision }
        if (!source || context.value !== owner || !current(completed)) {
          if (context.value === owner) invalidate()
          return { accepted: true, changed }
        }
        owner.source = source; owner.revision = revision
        if (changed) options.refresh(studio, unitId, () => context.value === owner && current(owner))
      } else {
        changed = options.applyStandalone(rules, () => context.value === owner && hostCurrent(owner), snapshot => {
          if (context.value === owner && hostCurrent(owner)) {
            owner.standaloneEcho = toRaw(snapshot); owner.standaloneEchoSnapshot = JSON.stringify(snapshot)
          }
        })
        if (context.value !== owner || !hostCurrent(owner)) return { accepted: true, changed }
        if (owner.standaloneEchoSnapshot && JSON.stringify(options.state().model) !== owner.standaloneEchoSnapshot) {
          invalidate(); return { accepted: true, changed }
        }
        owner.source = toRaw(options.state().model); owner.standaloneSnapshot = JSON.stringify(options.state().model)
        return { accepted: true, changed }
      }
      if (context.value === owner && !current(owner)) invalidate()
      return { accepted: true, changed }
    } catch (cause) {
      const message = cause instanceof Error ? cause.message : '无法应用条件格式，草稿已保留。'
      return { accepted: false, error: message.startsWith('Command was canceled:') ? '条件格式操作已取消，草稿已保留。' : message }
    } finally { submitting = false; inspect() }
  }
  function openPanel(initialType?: ConditionalFormatPanelInitialType): boolean {
    const state = options.state(), unitId = state.model.id, studio = state.studio, sheet = state.sheet
    if (stopped || !sheet || !state.active || state.blocked || options.open.value) return false
    const source = studio && unitId ? studio.getWorkbookModel(unitId) : toRaw(state.model)
    if (!source) return false
    const sourceSheet = studio && unitId ? studio.getWorkbookModel(unitId)?.sheets.find(item => item.id === sheet.id) : sheet
    if (!sourceSheet) return false
    const owner: ConditionalPanelOwner = { key: ++sequence, studio, unitId, sheetId: sheet.id, userId: state.userId,
      inputSessionId: state.inputSessionId, source, revision: studio && unitId ? studio.getWorkbookRevision(unitId) : undefined,
      historyGeneration: studio && unitId ? studio.editor.history.captureSavepoint(unitId).generation : undefined,
      standaloneSnapshot: studio ? undefined : JSON.stringify(state.model), rules: cloneModel(sourceSheet.conditionalFormats ?? []),
      defaultRange: state.defaultRange, ranges: cloneModel([...state.ranges]), initialType,
      setOpen: value => setOpen(owner, value), closed: () => closed(owner), submit: rules => submit(owner, rules),
      canEditRule: rule => !options.state().blocked && rangesAllowed(owner, rule.ranges, true),
      canApplyRules: rules => rulesAllowed(owner, rules),
      preview: ranges => { if (context.value === owner) highlighted.value = options.open.value && rangesAllowed(owner, ranges, false) ? cloneModel([...ranges]) : [] },
    }
    if (!hostCurrent(owner) || !rangesAllowed(owner, [], false)
      || initialType && (!owner.ranges.length || !rangesAllowed(owner, owner.ranges, true))) return false
    context.value = owner; highlighted.value = []; options.open.value = true
    return true
  }
  const disabled = computed(() => !context.value || !current(context.value) || options.state().blocked)
  const canCreate = computed(() => !!context.value && !disabled.value && !!context.value.ranges.length && rangesAllowed(context.value, context.value.ranges, true))
  const previewRanges = computed(() => options.open.value && context.value && !disabled.value
    && rangesAllowed(context.value, highlighted.value, false) ? highlighted.value : [])
  watch(() => options.state().studio, (studio, _previous, cleanup) => {
    const subscriptions: Disposable[] = studio ? [studio.events.addEvent('CommandExecuted', inspect), studio.events.addEvent('Undo', inspect),
      studio.events.addEvent('Redo', inspect), studio.events.addEvent('WorkbookCreated', inspect), studio.events.addEvent('WorkbookDisposed', inspect),
      studio.events.addEvent('SelectionsChanged', inspect), studio.events.addEvent('LifeCycleChanged', inspect)] : []
    cleanup(() => subscriptions.forEach(subscription => subscription.dispose()))
  }, { immediate: true, flush: 'sync' })
  // A host may leave and return during CommandExecuted or view refresh. That still ends the old input owner.
  watch(() => {
    const state = options.state(), studio = state.studio
    return [state.studio, state.model.id, state.sheetId, state.active, state.userId, state.inputSessionId,
      studio?.editor.state.activeKind, studio?.editor.workbook.model?.id, studio?.editor.workbook.model?.activeSheetId] as const
  }, (values, previous) => { if (previous && values.some((value, index) => value !== previous[index])) invalidate() }, { flush: 'sync' })
  watch(() => {
    const state = options.state(), studio = state.studio, unitId = state.model.id
    return [state.studio, unitId, state.sheetId, state.active, state.userId, state.inputSessionId,
      studio && unitId ? studio.getWorkbookModel(unitId) : toRaw(state.model),
      studio && unitId ? studio.getWorkbookRevision(unitId) : undefined,
      studio && unitId ? studio.editor.history.captureSavepoint(unitId).generation : undefined] as const
  }, inspect, { flush: 'sync' })
  onScopeDispose(() => { stopped = true; invalidate() })
  return { context, disabled, canCreate, previewRanges, openPanel, receiveView }
}
