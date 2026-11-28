import { computed, nextTick, onScopeDispose, shallowReactive, shallowRef, watch, type Ref } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { Disposable } from '../../api/events.ts'
import type { CellRange, SheetDataValidationRule, SheetDataValidationType, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { planSheetDataValidationCommand, SHEET_DATA_VALIDATION_COMMANDS, SHEET_DATA_VALIDATION_ASYNC_COMMANDS, sheetDataValidationCommandInput,
  type SheetDataValidationCommandInput, type SheetDataValidationSubmitResult } from '../../features/sheet-validation/validation-rule-command.ts'
import { sheetValidationListOptions, type SheetValidationListOption } from '../../features/sheet-validation/validation-list-input.ts'
import type { SheetCheckboxRuleCellProgress } from '../../features/sheet-validation/validation-checkbox-rule-cells-async.ts'

export type SheetDataValidationPanelAction = { action: 'set' | 'replace'; rule: SheetDataValidationRule }
  | { action: 'remove'; ruleId: string } | { action: 'clear'; ranges: CellRange[] }
  | { action: 'clearTypes'; ranges: CellRange[]; types: SheetDataValidationType[] }
interface ValidationWorkbenchState {
  studio?: OfficeStudio; model: ExcelWorkbookView; sheetId: string; userId: string; active: boolean; blocked: boolean
  inputSessionId?: string; inputOwner: string; ranges: readonly CellRange[]; defaultRange: string
}
interface ValidationWorkbenchOptions {
  open: Ref<boolean>; state(): ValidationWorkbenchState
  refresh(studio: OfficeStudio, unitId: string, current: () => boolean): void
  focus(): void; canRestoreFocus(): boolean
}
export interface SheetDataValidationPanelContext {
  key: number; sheetName: string; inputOwner: string; defaultRange: string; initialType?: SheetDataValidationType; initialRuleId?: string
  readonly rules: readonly SheetDataValidationRule[]; readonly disabled: boolean; readonly canAdd: boolean; readonly canRemoveAll: boolean
  readonly composing: boolean; readonly navigationRevision: number
  readonly pending: boolean; readonly progress?: SheetCheckboxRuleCellProgress
  isCurrent(): boolean; setOpen(open: boolean): void; closed(): void
  submit(action: SheetDataValidationPanelAction): SheetDataValidationSubmitResult
  submitAsync(action: SheetDataValidationPanelAction): Promise<SheetDataValidationSubmitResult>; cancel(): void
  canEditRule(rule: SheetDataValidationRule): boolean; canApplyRule(rule: SheetDataValidationRule): boolean
  listOptions(formula: string): readonly SheetValidationListOption[]
  showRule(ruleId: string): boolean; showRules(): void; setComposing(value: boolean): void
}
interface ValidationPanelOwner extends SheetDataValidationPanelContext {
  studio: OfficeStudio; unitId: string; sheetId: string; userId: string; inputSessionId?: string
  source: WorkbookModel; revision: number; historyGeneration: number; allRules: SheetDataValidationRule[]; rulesSignature: string
  composing: boolean; navigationRevision: number; pending: boolean; progress?: SheetCheckboxRuleCellProgress
}

/** The rule sidebar publishes only visible rules. One owner binds its full-list guard to the workbook, input and history session. */
export function useSheetDataValidationWorkbench(options: ValidationWorkbenchOptions) {
  const context = shallowRef<ValidationPanelOwner>()
  let sequence = 0, stopped = false, submitting = false
  let asyncController: AbortController | undefined
  function hostCurrent(owner: ValidationPanelOwner): boolean {
    const state = options.state(), studio = owner.studio
    return !stopped && owner.key === sequence && state.active && state.studio === studio && state.model.id === owner.unitId
      && state.sheetId === owner.sheetId && state.userId === owner.userId && state.inputSessionId === owner.inputSessionId
      && !studio.isDisposed && studio.activeWorkbookId === owner.unitId && studio.editor.state.activeKind === 'workbook'
      && studio.getWorkbookModel(owner.unitId)?.activeSheetId === owner.sheetId
  }
  function liveSheet(owner: ValidationPanelOwner): SpreadsheetSheet | undefined {
    return owner.studio.getWorkbookModel(owner.unitId)?.sheets.find(sheet => sheet.id === owner.sheetId)
  }
  function current(owner: ValidationPanelOwner): boolean {
    return hostCurrent(owner) && owner.studio.getWorkbookModel(owner.unitId) === owner.source
      && owner.studio.getWorkbookRevision(owner.unitId) === owner.revision && owner.studio.currentUserId === owner.userId
      && owner.studio.editor.history.captureSavepoint(owner.unitId).generation === owner.historyGeneration
      && JSON.stringify(liveSheet(owner)?.dataValidations ?? []) === owner.rulesSignature
  }
  function rangesAllowed(owner: ValidationPanelOwner, ranges: readonly CellRange[], edit: boolean): boolean {
    if (!current(owner)) return false
    const sheet = liveSheet(owner)
    return !!sheet && canViewWorksheetProtection(sheet, owner.userId) && canViewProtectedRanges(sheet, ranges, owner.userId)
      && (!edit || canEditWorksheetProtection(sheet, owner.userId) && canEditProtectedRanges(sheet, ranges, owner.userId))
  }
  function canApply(owner: ValidationPanelOwner, rule: SheetDataValidationRule): boolean {
    if (options.state().blocked || owner.composing || !current(owner)) return false
    const sheet = liveSheet(owner)
    if (!sheet) return false
    try {
      const input = sheetDataValidationCommandInput('replace', { unitId: owner.unitId, sheetId: owner.sheetId, rule })
      return rangesAllowed(owner, planSheetDataValidationCommand(sheet, input).ranges, true)
    } catch { return false }
  }
  function invalidate(): void { asyncController?.abort(); sequence++; options.open.value = false; context.value = undefined }
  function inspect(): void { if (!submitting && context.value && !current(context.value)) invalidate() }
  function closed(owner: ValidationPanelOwner): void {
    if (context.value !== owner || options.open.value) return
    context.value = undefined
    void nextTick(() => { if (!context.value && !options.open.value && current(owner) && options.canRestoreFocus()) options.focus() })
  }
  function submit(owner: ValidationPanelOwner, action: SheetDataValidationPanelAction): SheetDataValidationSubmitResult {
    if (submitting || context.value !== owner || !options.open.value || !current(owner) || options.state().blocked || owner.composing) return { accepted: false, error: '数据验证目标已更改或输入尚未结束，请完成输入后重试。' }
    submitting = true
    try {
      const { action: kind, ...payload } = action
      const input = sheetDataValidationCommandInput(kind, { ...payload, unitId: owner.unitId, sheetId: owner.sheetId,
        host: 'workbench', expectedRevision: owner.revision, expectedHistoryGeneration: owner.historyGeneration,
        expectedUserId: owner.userId, expectedRulesSignature: owner.rulesSignature })
      const changed = owner.studio.executeCommandSync<SheetDataValidationCommandInput, boolean>(SHEET_DATA_VALIDATION_COMMANDS[kind], input)
      const source = owner.studio.getWorkbookModel(owner.unitId), revision = owner.revision + (changed ? 1 : 0)
      if (!source || context.value !== owner || !hostCurrent(owner) || owner.studio.getWorkbookRevision(owner.unitId) !== revision
        || owner.studio.editor.history.captureSavepoint(owner.unitId).generation !== owner.historyGeneration) {
        if (context.value === owner) invalidate()
        return { accepted: true, changed }
      }
      owner.source = source; owner.revision = revision; owner.allRules = cloneModel(source.sheets.find(sheet => sheet.id === owner.sheetId)?.dataValidations ?? [])
      owner.rulesSignature = JSON.stringify(owner.allRules)
      if (changed) options.refresh(owner.studio, owner.unitId, () => context.value === owner && current(owner))
      if (context.value === owner && !current(owner)) invalidate()
      return { accepted: true, changed }
    } catch (cause) {
      const error = cause instanceof Error ? cause.message : '数据验证规则未能保存，草稿已保留。'
      return { accepted: false, error: error.startsWith('Command was canceled:') ? '数据验证操作已取消，草稿已保留。' : error }
    } finally { submitting = false; inspect() }
  }
  async function submitAsync(owner: ValidationPanelOwner, action: SheetDataValidationPanelAction): Promise<SheetDataValidationSubmitResult> {
    if (action.action !== 'set' && action.action !== 'replace') return submit(owner, action)
    if (submitting || context.value !== owner || !options.open.value || !current(owner) || options.state().blocked || owner.composing) {
      return { accepted: false, error: '数据验证目标已更改或输入尚未结束，请完成输入后重试。' }
    }
    const controller = new AbortController(); asyncController = controller; submitting = true; owner.pending = true; owner.progress = undefined
    try {
      const { action: kind, ...payload } = action
      const input = sheetDataValidationCommandInput(kind, { ...payload, unitId: owner.unitId, sheetId: owner.sheetId,
        host: 'workbench', expectedRevision: owner.revision, expectedHistoryGeneration: owner.historyGeneration,
        expectedUserId: owner.userId, expectedRulesSignature: owner.rulesSignature })
      const changed = await owner.studio.executeCommand(SHEET_DATA_VALIDATION_ASYNC_COMMANDS[kind], { ...input, signal: controller.signal,
        onProgress: (progress: SheetCheckboxRuleCellProgress) => {
          if (asyncController === controller && !controller.signal.aborted && context.value === owner && current(owner)) owner.progress = progress
        } })
      const source = owner.studio.getWorkbookModel(owner.unitId), revision = owner.revision + (changed ? 1 : 0)
      if (!source || context.value !== owner || !hostCurrent(owner) || owner.studio.getWorkbookRevision(owner.unitId) !== revision
        || owner.studio.editor.history.captureSavepoint(owner.unitId).generation !== owner.historyGeneration) {
        if (context.value === owner) invalidate()
        return { accepted: true, changed }
      }
      owner.source = source; owner.revision = revision; owner.allRules = cloneModel(source.sheets.find(sheet => sheet.id === owner.sheetId)?.dataValidations ?? [])
      owner.rulesSignature = JSON.stringify(owner.allRules)
      if (changed) options.refresh(owner.studio, owner.unitId, () => context.value === owner && current(owner))
      if (context.value === owner && !current(owner)) invalidate()
      return { accepted: true, changed }
    } catch (cause) {
      const error = cause instanceof Error ? cause.message : '数据验证规则未能保存，草稿已保留。'
      return { accepted: false, error: error.startsWith('Command was canceled:') || controller.signal.aborted ? '数据验证操作已取消，草稿已保留。' : error }
    } finally {
      if (asyncController === controller) asyncController = undefined
      owner.pending = false; owner.progress = undefined; submitting = false; inspect()
    }
  }
  function openPanel(initialType?: SheetDataValidationType, initialRuleId?: string): boolean {
    const state = options.state(), studio = state.studio, unitId = state.model.id
    if (stopped || !studio || !unitId || !state.active || state.blocked || options.open.value) return false
    const source = studio.getWorkbookModel(unitId), sheet = source?.sheets.find(candidate => candidate.id === state.sheetId)
    if (!source || !sheet) return false
    const owner: ValidationPanelOwner = shallowReactive({ key: ++sequence, studio, unitId, sheetId: sheet.id, sheetName: sheet.name,
      userId: state.userId, inputOwner: state.inputOwner, inputSessionId: state.inputSessionId, source,
      revision: studio.getWorkbookRevision(unitId), historyGeneration: studio.editor.history.captureSavepoint(unitId).generation,
      allRules: cloneModel(sheet.dataValidations ?? []), rulesSignature: JSON.stringify(sheet.dataValidations ?? []),
      defaultRange: state.defaultRange, initialType, initialRuleId, composing: false, navigationRevision: 0, pending: false,
      get rules() { return owner.allRules.filter(rule => rangesAllowed(owner, rule.ranges, false)) },
      get disabled() { return owner.pending || options.state().blocked || !current(owner) },
      get canAdd() { return !owner.disabled && !owner.composing && !!options.state().ranges.length && rangesAllowed(owner, options.state().ranges, true) },
      get canRemoveAll() { return !owner.disabled && !owner.composing && !!owner.rules.length && owner.rules.length === owner.allRules.length
        && rangesAllowed(owner, owner.allRules.flatMap(rule => rule.ranges), true) },
      isCurrent: () => context.value === owner && current(owner),
      canEditRule: rule => !owner.disabled && !owner.composing && owner.allRules.some(current => current.id === rule.id && JSON.stringify(current) === JSON.stringify(rule))
        && rangesAllowed(owner, rule.ranges, true),
      canApplyRule: rule => canApply(owner, rule),
      listOptions: formula => current(owner) ? sheetValidationListOptions({ id: 'validation-panel-preview', type: 'list', ranges: [], formula1: formula },
        owner.source, owner.sheetId, owner.userId, owner.studio.regionalSettings.region.value) : [],
      showRule: ruleId => { if (!current(owner) || owner.composing || !owner.rules.some(rule => rule.id === ruleId && owner.canEditRule(rule))) return false
        owner.initialType = undefined; owner.initialRuleId = ruleId; owner.navigationRevision++; return true },
      showRules: () => { if (context.value === owner && current(owner) && !owner.composing) { owner.initialType = undefined; owner.initialRuleId = undefined; owner.navigationRevision++ } },
      setComposing: value => { if (context.value === owner && current(owner)) owner.composing = value },
      submit: action => submit(owner, action),
      submitAsync: action => submitAsync(owner, action), cancel: () => { if (context.value === owner && owner.pending) asyncController?.abort() },
      setOpen: value => { if (context.value === owner && !value && !submitting && !owner.composing) options.open.value = false }, closed: () => closed(owner),
    })
    if (!hostCurrent(owner) || !rangesAllowed(owner, [], false)
      || initialType && (!state.ranges.length || !rangesAllowed(owner, state.ranges, true))
      || initialRuleId && !owner.rules.some(rule => rule.id === initialRuleId)) return false
    context.value = owner; options.open.value = true
    return true
  }
  watch(() => options.state().studio, (studio, _previous, cleanup) => {
    const subscriptions: Disposable[] = studio ? (['CommandExecuted', 'Undo', 'Redo', 'WorkbookCreated', 'WorkbookDisposed', 'SelectionsChanged', 'LifeCycleChanged'] as const)
      .map(event => studio.events.addEvent(event, inspect)) : []
    cleanup(() => subscriptions.forEach(subscription => subscription.dispose()))
  }, { immediate: true, flush: 'sync' })
  watch(() => {
    const state = options.state(), studio = state.studio
    return [state.studio, state.model.id, state.sheetId, state.active, state.userId, state.inputSessionId,
      studio?.editor.state.activeKind, studio?.editor.workbook.model?.id, studio?.editor.workbook.model?.activeSheetId] as const
  }, (values, previous) => { if (previous && values.some((value, index) => value !== previous[index])) invalidate() }, { flush: 'sync' })
  watch(() => {
    const state = options.state(), studio = state.studio, unitId = state.model.id
    const workbook = studio && unitId ? studio.getWorkbookModel(unitId) : undefined
    return [workbook, studio && unitId ? studio.getWorkbookRevision(unitId) : undefined,
      studio && unitId ? studio.editor.history.captureSavepoint(unitId).generation : undefined,
      JSON.stringify(workbook?.sheets.find(sheet => sheet.id === state.sheetId)?.dataValidations ?? [])] as const
  }, inspect, { flush: 'sync' })
  onScopeDispose(() => { stopped = true; invalidate() })
  return { context, openPanel, disabled: computed(() => !context.value || context.value.disabled) }
}
