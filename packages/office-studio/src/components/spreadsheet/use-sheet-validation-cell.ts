import { computed, nextTick, onScopeDispose, shallowReactive, shallowRef, toRaw, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { Disposable } from '../../api/events.ts'
import { createBlankWorkbook, type CellRange, type SheetDataValidationRule, type WorkbookModel } from '../../domain/model.ts'
import { applyExcelWorkbookView, cellKey, parseCellKey, sheetModelCellViewAt, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { createSheetCellEditSession, getSheetCellEditConflict, resolveSheetDate1904, type SheetCellEditSession } from '../../features/sheet-workbench/sheet-cell-edit-session.ts'
import { sheetCellInputConflictMessage, type SheetCellInputCommandResult, type SheetCellInputDefaults, type SheetCellInputPlan } from '../../features/sheet-workbench/sheet-cell-input-command.ts'
import { sheetCellNumberPattern, sheetCellNumberValue } from '../../features/sheet-workbench/sheet-number-presentation.ts'
import { dataValidationsAt } from '../../features/sheet-validation/validation-engine.ts'
import { COMMIT_SHEET_DATE_INPUT_COMMAND, planSheetDateInput, sheetDateInputCommandInput, sheetDateInputState,
  type SheetDateInputChange, type SheetDateInputState } from '../../features/sheet-validation/validation-date-input.ts'
import { COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, planSheetValidationListInput, sheetValidationListInput,
  sheetValidationListSelected, type SheetValidationListOption } from '../../features/sheet-validation/validation-list-input.ts'
import type { SheetValidationWidgetTargetInput } from '../../features/sheet-validation/validation-widget-input.ts'
import { attachSheetValidationDropdownController } from '../../api/sheet-validation-dropdown-operations.ts'
import { SHEET_VALIDATION_DROPDOWN_OPERATIONS, type SheetValidationDropdownRequest } from '../../features/sheet-validation/validation-dropdown.ts'
import { SheetValidationListReadCache } from '../../features/sheet-validation/validation-list-cache.ts'

export interface SheetValidationCellEntry {
  hostKey: string; key: string; kind: 'date' | 'list'; disabled: boolean; inputOwner: string; triggerVisible?: boolean; open(): boolean
}
export interface SheetValidationCellPopup {
  ownerKey: number; hostKey: string; key: string; kind: 'date' | 'list'; visible: boolean; inputOwner: string
  date?: SheetDateInputState; options: readonly SheetValidationListOption[]; selected: readonly string[]; multiple: boolean
  region: string; currentYear: number
  readonly disabled: boolean; readonly message: string
  readonly canEditRule: boolean
  readonly composing: boolean
  current(): boolean
  queryContext(): { unitId: string; baseRevision: number } | undefined
  submitDate(change: SheetDateInputChange): SheetCellInputCommandResult
  chooseList(values: readonly string[]): SheetCellInputCommandResult
  setOpen(open: boolean): void; afterOpenChange(open: boolean): void; editRule(): void
  setComposing(value: boolean): void
  detach(): void
}
interface ValidationCellState {
  studio?: OfficeStudio; model: ExcelWorkbookView; sheetId: string; key: string; userId: string
  active: boolean; blocked: boolean; inputSession?: SheetCellEditSession | null; region: string; currentYear: number; inputOwner: string
}
interface ValidationCellOptions {
  state(): ValidationCellState
  refresh(studio: OfficeStudio, unitId: string, current: () => boolean): void
  applyStandalone(plan: SheetCellInputPlan, workbook: WorkbookModel, current: () => boolean,
    expectEcho: (snapshot: ExcelWorkbookView) => void): SheetCellInputCommandResult
  endEditing(sessionId: string): void
  openRule(ruleId: string): void
  focus(session?: SheetCellEditSession): void
  canRestoreFocus(): boolean
  reveal?(key: string): boolean
}
interface ValidationCellOwner extends SheetValidationCellPopup {
  studio?: OfficeStudio; unitId?: string; sheetId: string; userId: string; ruleId: string; ruleSignature: string
  inputSessionId?: string; session: SheetCellEditSession; source: WorkbookModel; revision?: number; historyGeneration?: number
  viewSource: object; standaloneSnapshot?: string; standaloneEcho?: object; standaloneEchoSnapshot?: string; error: string
  composing: boolean; requested: boolean; selectionKeyAtOpen?: string; triggerVisible: boolean; hostRegion: string
}

/** One selected-cell popup owns a canonical edit session. Successful multiple choices rebase that session and keep it open. */
export function useSheetValidationCell(options: ValidationCellOptions) {
  const context = shallowRef<ValidationCellOwner>()
  const listCache = new SheetValidationListReadCache()
  const noListOptions: readonly SheetValidationListOption[] = Object.freeze([])
  const noSelection: readonly string[] = Object.freeze([])
  let sequence = 0, stopped = false, submitting = false, endingInput = false
  const sourceTokens = new WeakMap<object, number>()
  let sourceSequence = 0
  const sourceModel = computed(() => {
    const state = options.state(), unitId = state.model.id
    return state.studio && unitId ? state.studio.getWorkbookModel(unitId)
      : applyExcelWorkbookView(createBlankWorkbook(), state.model, state.region)
  })
  function info(requestedKey?: string, force = false) {
    const state = options.state(), workbook = sourceModel.value, sheet = workbook?.sheets.find(candidate => candidate.id === state.sheetId)
    const position = parseCellKey(requestedKey ?? state.key)
    if (!workbook || !sheet || !position) return undefined
    const merge = sheet.mergedRanges.find(range => position.row >= range.startRow && position.row <= range.endRow
      && position.col >= range.startCol && position.col <= range.endCol)
    const row = merge?.startRow ?? position.row, column = merge?.startCol ?? position.col
    const range: CellRange = merge ?? { startRow: row, endRow: row, startCol: column, endCol: column }
    if (!canViewWorksheetProtection(sheet, state.userId) || !canViewProtectedRanges(sheet, [range], state.userId)) return undefined
    const rule = dataValidationsAt(sheet, row, column).find(candidate => ['date', 'time', 'list', 'listMultiple'].includes(candidate.type)
      && (force || candidate.showDropDown !== false && candidate.renderMode !== 'text'))
    if (!rule) return undefined
    const key = cellKey(row, column), kind = rule.type === 'date' || rule.type === 'time' ? 'date' as const : 'list' as const
    const hostKey = JSON.stringify([state.model.id, sheet.id, key, rule.id])
    const source = toRaw(workbook)
    if (!sourceTokens.has(source)) sourceTokens.set(source, ++sourceSequence)
    const token = JSON.stringify([hostKey, sourceTokens.get(source), state.studio && state.model.id ? state.studio.getWorkbookRevision(state.model.id) : undefined,
      state.inputSession?.sessionId, state.userId, state.region, sequence, rule, sheet.cells[`${row}:${column}`], force])
    return { state, workbook, sheet, rule, range, row, column, key, kind, hostKey, token,
      triggerVisible: rule.showDropDown !== false && rule.renderMode !== 'text' }
  }
  function hostMatches(owner: ValidationCellOwner): boolean {
    const state = options.state(), studio = owner.studio, workbook = sourceModel.value
    return !stopped && state.active && state.studio === studio && state.model.id === owner.unitId
      && state.sheetId === owner.sheetId && state.userId === owner.userId && state.inputSession?.sessionId === owner.inputSessionId
      && state.region === owner.hostRegion
      && cellKeyForState(state) === (owner.requested ? owner.selectionKeyAtOpen : owner.key) && (!studio || !studio.isDisposed && owner.unitId !== undefined
        && studio.activeWorkbookId === owner.unitId && studio.editor.state.activeKind === 'workbook'
        && workbook?.activeSheetId === owner.sheetId)
  }
  function hostCurrent(owner: ValidationCellOwner): boolean { return owner.ownerKey === sequence && hostMatches(owner) }
  function cellKeyForState(state: ValidationCellState): string | undefined {
    const position = parseCellKey(state.key), sheet = sourceModel.value?.sheets.find(candidate => candidate.id === state.sheetId)
    if (!position || !sheet) return undefined
    const merge = sheet.mergedRanges.find(range => position.row >= range.startRow && position.row <= range.endRow
      && position.col >= range.startCol && position.col <= range.endCol)
    return cellKey(merge?.startRow ?? position.row, merge?.startCol ?? position.col)
  }
  function ownerCurrent(owner: ValidationCellOwner): boolean {
    if (!hostCurrent(owner)) return false
    const studio = owner.studio, workbook = sourceModel.value
    return !!workbook && (studio && owner.unitId ? studio.editor.history.captureSavepoint(owner.unitId).generation === owner.historyGeneration
      && workbook === owner.source && studio.getWorkbookRevision(owner.unitId) === owner.revision
      : toRaw(options.state().model) === owner.viewSource && JSON.stringify(options.state().model) === owner.standaloneSnapshot)
  }
  function ownerCanView(owner: ValidationCellOwner): boolean {
    const workbook = sourceModel.value, sheet = workbook?.sheets.find(candidate => candidate.id === owner.sheetId), position = parseCellKey(owner.key)
    if (!sheet || !position) return false
    const range = sheet.mergedRanges.find(candidate => position.row >= candidate.startRow && position.row <= candidate.endRow
      && position.col >= candidate.startCol && position.col <= candidate.endCol)
      ?? { startRow: position.row, endRow: position.row, startCol: position.col, endCol: position.col }
    return canViewWorksheetProtection(sheet, owner.userId) && canViewProtectedRanges(sheet, [range], owner.userId)
  }
  function popupCurrent(owner: ValidationCellOwner): boolean {
    const generation = sequence
    return ownerCurrent(owner) && ownerCanView(owner) && ownerCurrent(owner) && generation === sequence && context.value === owner
  }
  function ownerConflict(owner: ValidationCellOwner): string | undefined {
    const workbook = sourceModel.value, sheet = workbook?.sheets.find(candidate => candidate.id === owner.sheetId), position = parseCellKey(owner.key)
    if (!workbook || !sheet || !position) return '编辑目标已关闭，请重新打开选择控件。'
    const rule = dataValidationsAt(sheet, position.row, position.col).find(candidate => candidate.id === owner.ruleId)
    if (!rule || JSON.stringify(rule) !== owner.ruleSignature) return '数据验证规则已更改，请重新打开选择控件。'
    const range = sheet.mergedRanges.find(range => position.row >= range.startRow && position.row <= range.endRow
      && position.col >= range.startCol && position.col <= range.endCol)
      ?? { startRow: position.row, endRow: position.row, startCol: position.col, endCol: position.col }
    if (!canViewWorksheetProtection(sheet, owner.userId) || !canViewProtectedRanges(sheet, [range], owner.userId)
      || !canEditWorksheetProtection(sheet, owner.userId) || !canEditProtectedRanges(sheet, [range], owner.userId)) return '目标单元格受保护，当前用户没有编辑权限。'
    const conflict = getSheetCellEditConflict(owner.session, { documentId: workbook.id, sheetId: sheet.id, key: owner.key,
      date1904: resolveSheetDate1904(undefined, workbook), cell: sheetModelCellViewAt(workbook, sheet, position.row, position.col) })
    return conflict && sheetCellInputConflictMessage(conflict)
  }
  function invalidate(): void { sequence++; context.value = undefined }
  function inspect(): void {
    const owner = context.value
    if (!owner || submitting) return
    if (!hostCurrent(owner)) { invalidate(); return }
    const readable = ownerCanView(owner)
    if (context.value !== owner) return
    if (!readable) { invalidate(); return }
    if (owner.studio && owner.unitId) {
      const workbook = sourceModel.value, revision = owner.studio.getWorkbookRevision(owner.unitId)
      if (!workbook || owner.studio.editor.history.captureSavepoint(owner.unitId).generation !== owner.historyGeneration
        || workbook !== owner.source && revision === owner.revision) { invalidate(); return }
      // Unrelated content edits can advance the model; the captured cell and rule still guard this draft.
      owner.source = workbook; owner.revision = revision
    } else if (!ownerCurrent(owner)) { invalidate(); return }
  }
  function setOpen(owner: ValidationCellOwner, value: boolean): void {
    if (context.value !== owner || value || submitting || owner.composing) return
    owner.visible = false
  }
  function requestClose(owner: ValidationCellOwner, value: boolean): void {
    if (context.value !== owner || value || submitting || owner.composing || !hostCurrent(owner)) return
    if (!owner.studio) { setOpen(owner, value); return }
    try { owner.studio.executeCommandSync(SHEET_VALIDATION_DROPDOWN_OPERATIONS.hide, { unitId: owner.unitId, subUnitId: owner.sheetId }) }
    catch (cause) { owner.error = cause instanceof Error ? cause.message : '关闭选择控件失败。' }
  }
  function afterOpenChange(owner: ValidationCellOwner, value: boolean): void {
    if (value || context.value !== owner || owner.visible) return
    context.value = undefined
    void nextTick(() => {
      if (!context.value && ownerCurrent(owner) && options.canRestoreFocus()) options.focus(options.state().inputSession ?? undefined)
    })
  }
  function rebaseSession(owner: ValidationCellOwner): void {
    const workbook = sourceModel.value, sheet = workbook?.sheets.find(candidate => candidate.id === owner.sheetId), position = parseCellKey(owner.key)
    if (!workbook || !sheet || !position) return
    owner.session = createSheetCellEditSession({ documentId: workbook.id, sheetId: sheet.id, key: owner.key,
      date1904: resolveSheetDate1904(undefined, workbook), cell: sheetModelCellViewAt(workbook, sheet, position.row, position.col) },
    { region: owner.session.region, currentYear: owner.session.currentYear })
  }
  function submit(owner: ValidationCellOwner, command: string, values: { change: SheetDateInputChange } | { values: readonly string[] }): SheetCellInputCommandResult {
    inspect()
    const reject = (message: string): SheetCellInputCommandResult => ({ accepted: false, changed: false, issues: [], message })
    if (context.value !== owner || !owner.visible || owner.composing || options.state().blocked || !ownerCurrent(owner)) return reject('选择控件的目标已更改或正在输入组合文字，请重新操作。')
    const conflict = ownerConflict(owner)
    if (conflict) { owner.error = conflict; return reject(conflict) }
    const target: SheetValidationWidgetTargetInput = { unitId: owner.unitId ?? 'standalone-input', sheetId: owner.sheetId,
      row: parseCellKey(owner.key)!.row, column: parseCellKey(owner.key)!.col, ruleId: owner.ruleId, session: owner.session,
      host: 'workbench', expectedRevision: owner.revision, expectedHistoryGeneration: owner.historyGeneration,
      expectedUserId: owner.userId, expectedRuleSignature: owner.ruleSignature }
    submitting = true
    try {
      const studio = owner.studio, unitId = owner.unitId, baseRevision = owner.revision
      let result: SheetCellInputCommandResult
      if (studio && unitId) {
        result = studio.executeCommandSync(command, { ...target, ...values })
        const workbook = studio.getWorkbookModel(unitId), expectedRevision = (baseRevision ?? 0) + (result.changed ? 1 : 0)
        if (!workbook || context.value !== owner || !hostCurrent(owner)
          || studio.getWorkbookRevision(unitId) !== expectedRevision || studio.editor.history.captureSavepoint(unitId).generation !== owner.historyGeneration) {
          if (context.value === owner) invalidate()
          return result
        }
        owner.source = workbook; owner.revision = expectedRevision
      } else {
        const defaults: SheetCellInputDefaults = { region: owner.session.region, currentYear: owner.session.currentYear, userId: owner.userId }
        const model = sourceModel.value!
        const plan = 'change' in values ? planSheetDateInput(model, sheetDateInputCommandInput({ ...target, ...values }), defaults)
          : planSheetValidationListInput(model, sheetValidationListInput({ ...target, ...values }), defaults)
        result = options.applyStandalone(plan, model, () => context.value === owner && hostCurrent(owner), snapshot => {
          if (context.value !== owner || !hostCurrent(owner)) return
          owner.standaloneEcho = toRaw(snapshot); owner.standaloneEchoSnapshot = JSON.stringify(snapshot)
          owner.viewSource = toRaw(options.state().model); owner.standaloneSnapshot = JSON.stringify(options.state().model)
        })
        if (context.value !== owner || !hostCurrent(owner)
          || owner.standaloneEchoSnapshot && JSON.stringify(options.state().model) !== owner.standaloneEchoSnapshot) {
          if (context.value === owner) invalidate()
          return result
        }
        owner.viewSource = toRaw(options.state().model); owner.standaloneSnapshot = JSON.stringify(options.state().model)
        owner.source = sourceModel.value!
      }
      if (!result.accepted) { owner.error = result.message ?? '输入未能保存，草稿已保留。'; return result }
      owner.error = ''
      rebaseSession(owner)
      if (owner.inputSessionId) {
        endingInput = true
        try { options.endEditing(owner.inputSessionId); owner.inputSessionId = options.state().inputSession?.sessionId } finally { endingInput = false }
      }
      if (context.value !== owner || !ownerCurrent(owner)) return result
      if (studio && unitId && result.changed) options.refresh(studio, unitId, () => context.value === owner && ownerCurrent(owner))
      if (context.value === owner && ownerCurrent(owner) && (owner.kind === 'date' || !owner.multiple)) owner.visible = false
      return result
    } catch (cause) {
      owner.error = cause instanceof Error ? cause.message : '输入未能保存，草稿已保留。'
      if (owner.error.startsWith('Command was canceled:')) owner.error = '选择操作已取消，草稿已保留。'
      return reject(owner.error)
    } finally { submitting = false; inspect() }
  }
  function openEntry(expectedToken: string, requestedKey?: string, request?: SheetValidationDropdownRequest): boolean {
    inspect()
    const current = info(requestedKey, !!request)
    if (stopped || !current || current.token !== expectedToken || !current.state.active || current.state.blocked) return false
    if (!canEditWorksheetProtection(current.sheet, current.state.userId) || !canEditProtectedRanges(current.sheet, [current.range], current.state.userId)) return false
    if (request && (request.unitId !== current.state.model.id || request.sheetId !== current.sheet.id
      || request.ruleId !== current.rule.id || request.expectedRuleSignature !== JSON.stringify(current.rule)
      || request.expectedUserId !== current.state.userId || !current.state.studio
      || current.state.studio.getWorkbookRevision(request.unitId) !== request.expectedRevision
      || current.state.studio.editor.history.captureSavepoint(request.unitId).generation !== request.expectedHistoryGeneration)) return false
    if (request && current.state.inputSession && cellKeyForState(current.state) !== current.key) return false
    const previous = context.value, baseSequence = sequence
    if (previous?.visible) {
      if (!request) return false
      const owner = previous
      if (owner.hostKey === current.hostKey && ownerCurrent(owner)) return true
      if (owner.composing) return false
    }
    const { state, workbook, sheet, rule, key, kind, hostKey } = current, view = sheetModelCellViewAt(workbook, sheet, current.row, current.column)
    const session = state.inputSession ?? createSheetCellEditSession({ documentId: workbook.id, sheetId: sheet.id, key,
      date1904: resolveSheetDate1904(undefined, workbook), cell: view }, { region: state.region, currentYear: state.currentYear })
    const listOptions = computed(() => {
      const model = sourceModel.value, target = model?.sheets.find(candidate => candidate.id === sheet.id)
      const activeRule = target && dataValidationsAt(target, current.row, current.column).find(candidate => candidate.id === rule.id)
      return kind === 'list' && model && activeRule ? listCache.read(activeRule, model, sheet.id, state.userId, session.region) : []
    })
    const selected = computed(() => sheetValidationListSelected(listOptions.value,
      sourceModel.value?.sheets.find(candidate => candidate.id === sheet.id)?.cells[`${current.row}:${current.column}`]?.value, rule.type === 'listMultiple'))
    const owner: ValidationCellOwner = shallowReactive({ ownerKey: baseSequence + 1, hostKey, key, kind, visible: true, inputOwner: state.inputOwner,
      studio: state.studio, unitId: state.model.id, sheetId: sheet.id, userId: state.userId, ruleId: rule.id, ruleSignature: JSON.stringify(rule),
      inputSessionId: state.inputSession?.sessionId, session, source: workbook, viewSource: toRaw(state.model),
      revision: state.studio && state.model.id ? state.studio.getWorkbookRevision(state.model.id) : undefined,
      historyGeneration: state.studio && state.model.id ? state.studio.editor.history.captureSavepoint(state.model.id).generation : undefined,
      standaloneSnapshot: state.studio ? undefined : JSON.stringify(state.model), error: '', composing: false,
      requested: !!request, selectionKeyAtOpen: cellKeyForState(state), triggerVisible: current.triggerVisible,
      date: kind === 'date' ? sheetDateInputState(rule, sheetCellNumberValue(view), sheetCellNumberPattern(view), resolveSheetDate1904(undefined, workbook)) : undefined,
      get options() { if (!popupCurrent(owner)) return noListOptions
        const values = listOptions.value; return popupCurrent(owner) ? values : noListOptions },
      get selected() { if (!popupCurrent(owner)) return noSelection
        const values = selected.value; return popupCurrent(owner) ? values : noSelection }, multiple: rule.type === 'listMultiple',
      region: session.region, currentYear: session.currentYear, hostRegion: state.region,
      current: () => popupCurrent(owner),
      queryContext: () => { if (!popupCurrent(owner)) return undefined
        const unitId = owner.unitId ?? 'standalone-input', baseRevision = owner.revision ?? 0
        return popupCurrent(owner) && unitId === (owner.unitId ?? 'standalone-input') && baseRevision === (owner.revision ?? 0)
          ? { unitId, baseRevision } : undefined },
      get disabled() { return submitting || !ownerCurrent(owner) || options.state().blocked || !!ownerConflict(owner) },
      get message() { return ownerConflict(owner) ?? owner.error },
      get canEditRule() { const currentSheet = sourceModel.value?.sheets.find(candidate => candidate.id === owner.sheetId)
        return !!owner.studio && !!currentSheet && !owner.disabled && canViewProtectedRanges(currentSheet, rule.ranges, owner.userId)
          && canEditProtectedRanges(currentSheet, rule.ranges, owner.userId) },
      submitDate: change => submit(owner, COMMIT_SHEET_DATE_INPUT_COMMAND, { change }),
      chooseList: values => submit(owner, COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { values }),
      setOpen: value => requestClose(owner, value), afterOpenChange: value => afterOpenChange(owner, value),
      setComposing: value => { if (context.value === owner && hostCurrent(owner)) owner.composing = value },
      // The input bridge has disappeared (scroll clipping, Canvas replacement
      // or unmount). Retire its lease even during IME; an ordinary hide request
      // still preserves composition. No content or cell-input session commits.
      detach: () => { if (context.value === owner) { invalidate(); owner.visible = false; owner.composing = false } },
      editRule: () => { if (context.value === owner && ownerCurrent(owner) && !owner.composing && owner.canEditRule) { invalidate(); options.openRule(owner.ruleId) } },
    })
    // A failed reveal or a changed host keeps the previous popup and its draft. Publish the owner only after every preflight succeeds.
    if (!hostMatches(owner) || request && options.reveal && !options.reveal(owner.key)) return false
    if (sequence !== baseSequence || context.value !== previous || previous?.composing || !hostMatches(owner) || options.state().blocked
      || info(requestedKey, !!request)?.token !== current.token) return false
    sequence = owner.ownerKey
    if (previous) previous.visible = false
    context.value = owner
    return true
  }
  const entry = computed<SheetValidationCellEntry | undefined>(() => {
    const owner = context.value
    if (owner && hostCurrent(owner)) return { hostKey: owner.hostKey, key: owner.key, kind: owner.kind, inputOwner: owner.inputOwner,
      disabled: owner.disabled, triggerVisible: owner.triggerVisible, open: () => false }
    const current = info()
    if (!current || !current.state.active) return undefined
    return { hostKey: current.hostKey, key: current.key, kind: current.kind, inputOwner: current.state.inputOwner,
      triggerVisible: current.triggerVisible,
      disabled: current.state.blocked || !canEditWorksheetProtection(current.sheet, current.state.userId)
        || !canEditProtectedRanges(current.sheet, [current.range], current.state.userId), open: () => {
          const latest = info()
          if (!latest || latest.token !== current.token || !latest.state.active || latest.state.blocked) return false
          const studio = latest.state.studio, unitId = latest.state.model.id
          if (!studio || !unitId) return openEntry(latest.token)
          try {
            return studio.executeCommandSync(SHEET_VALIDATION_DROPDOWN_OPERATIONS.show, { unitId, subUnitId: latest.sheet.id,
              row: latest.row, column: latest.column, expectedRevision: studio.getWorkbookRevision(unitId),
              expectedHistoryGeneration: studio.editor.history.captureSavepoint(unitId).generation, expectedUserId: latest.state.userId })
          } catch (cause) {
            if (cause instanceof Error && cause.message.startsWith('Command was canceled:')) return false
            throw cause
          }
        } }
  })
  function receiveView(value: ExcelWorkbookView, apply: () => void): void {
    const owner = context.value
    if (!owner || owner.studio) { apply(); return }
    if (owner.standaloneEcho !== toRaw(value) || JSON.stringify(value) !== owner.standaloneEchoSnapshot || !hostCurrent(owner)) { invalidate(); apply(); return }
    owner.standaloneEcho = undefined
    const previous = submitting
    submitting = true
    try {
      apply()
      if (context.value === owner && hostCurrent(owner) && JSON.stringify(options.state().model) === owner.standaloneEchoSnapshot) {
        owner.viewSource = toRaw(options.state().model); owner.standaloneSnapshot = JSON.stringify(options.state().model)
      } else if (context.value === owner) invalidate()
    } finally { submitting = previous; inspect() }
  }
  watch(() => options.state().studio, (studio, _previous, cleanup) => {
    const subscriptions: Disposable[] = studio ? (['CommandExecuted', 'Undo', 'Redo', 'WorkbookCreated', 'WorkbookDisposed', 'SelectionsChanged', 'LifeCycleChanged'] as const)
      .map(event => studio.events.addEvent(event, inspect)) : []
    if (studio) subscriptions.push(attachSheetValidationDropdownController(studio, {
      supports: request => {
        const state = options.state()
        return !stopped && state.active && !state.blocked && state.studio === studio && state.model.id === request.unitId
          && state.sheetId === request.sheetId && state.userId === request.expectedUserId
      },
      show: request => {
        const key = cellKey(request.row, request.column), current = info(key, true)
        return !!current && openEntry(current.token, key, request)
      },
      hide: () => {
        inspect()
        const owner = context.value
        if (!owner) return true
        if (!hostCurrent(owner) || owner.composing) return false
        setOpen(owner, false)
        return !owner.visible
      },
      active: () => { const owner = context.value
        return owner?.visible && owner.unitId && hostCurrent(owner)
          ? { unitId: owner.unitId, sheetId: owner.sheetId, row: parseCellKey(owner.key)!.row, column: parseCellKey(owner.key)!.col } : undefined },
    }))
    cleanup(() => subscriptions.forEach(subscription => subscription.dispose()))
  }, { immediate: true, flush: 'sync' })
  watch(() => {
    const state = options.state(), studio = state.studio
    return [state.studio, state.model.id, state.sheetId, state.key, state.active, state.userId, state.inputSession?.sessionId, state.region,
      studio?.editor.state.activeKind, studio?.editor.workbook.model?.id, studio?.editor.workbook.model?.activeSheetId] as const
  }, (values, previous) => {
    if (!previous || values.every((value, index) => value === previous[index])) return
    if (endingInput && context.value && values.every((value, index) => index === 6 || value === previous[index])) { context.value.inputSessionId = options.state().inputSession?.sessionId; return }
    invalidate()
  }, { flush: 'sync' })
  watch(() => {
    const state = options.state(), workbook = sourceModel.value, sheet = workbook?.sheets.find(candidate => candidate.id === state.sheetId)
    const position = parseCellKey(cellKeyForState(state) ?? '')
    return [workbook, state.model, state.studio && state.model.id ? state.studio.getWorkbookRevision(state.model.id) : undefined,
      state.studio && state.model.id ? state.studio.editor.history.captureSavepoint(state.model.id).generation : undefined,
      position ? JSON.stringify([sheet?.cells[`${position.row}:${position.col}`], sheet && dataValidationsAt(sheet, position.row, position.col),
        sheet?.worksheetProtection, sheet?.rangeProtectionRules]) : undefined] as const
  }, (values, previous) => {
    // A close clears the visible owner before its deferred focus runs. Host
    // replacements during that interval still expire the original lease.
    if (!context.value && previous && values.some((value, index) => value !== previous[index])) sequence++
    inspect()
  }, { flush: 'sync' })
  watch(() => { const state = options.state(); return [state.studio, state.model.id, state.sheetId, state.userId, state.region, state.active] as const },
    () => listCache.clear(), { flush: 'sync' })
  onScopeDispose(() => { stopped = true; listCache.clear(); invalidate() })
  return { entry, popup: context, open: computed(() => context.value?.visible === true), receiveView }
}
