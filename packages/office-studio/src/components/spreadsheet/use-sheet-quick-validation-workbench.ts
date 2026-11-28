import { computed, onScopeDispose, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { CellRange } from '../../domain/model.ts'
import { planQuickValidation } from '../../features/sheet-validation/quick-validation.ts'
import { SHEET_DATA_VALIDATION_ASYNC_COMMANDS, SHEET_DATA_VALIDATION_COMMANDS } from '../../features/sheet-validation/validation-rule-command.ts'
import type { SheetCheckboxRuleCellProgress } from '../../features/sheet-validation/validation-checkbox-rule-cells-async.ts'
import type { SheetDataValidationPanelContext } from './use-sheet-data-validation-workbench.ts'

interface QuickValidationState { studio?: OfficeStudio; unitId?: string; sheetId: string; userId: string; active: boolean; blocked: boolean; ranges: readonly CellRange[] }
interface QuickValidationOptions {
  state(): QuickValidationState; translator(): (key: string, fallback: string) => string; owner(): SheetDataValidationPanelContext | undefined
  refresh(studio: OfficeStudio, unitId: string, current: () => boolean): void
  openRule(ruleId: string): void; notify(message: string): void
}
/** A quick rule owns its async content baseline and cancellation scope even when the rule panel is closed. */
export function useSheetQuickValidationWorkbench(options: QuickValidationOptions) {
  const pending = shallowRef<{ current(): boolean; cancel(): void; invalidate(): void }>(), progress = shallowRef<SheetCheckboxRuleCellProgress>()
  let stopped = false
  const cancel = () => pending.value?.cancel()
  async function apply(action: unknown): Promise<void> {
    const initial = options.state(), studio = initial.studio, unitId = initial.unitId
    if (stopped || pending.value || !studio || !unitId || !initial.active || initial.blocked) return
    const revision = studio.getWorkbookRevision(unitId), generation = studio.editor.history.captureSavepoint(unitId).generation
    let invalidated = false, submittingOwner: SheetDataValidationPanelContext | undefined
    const hostCurrent = () => {
      const state = options.state()
      return !stopped && !invalidated && state.active && !state.blocked && state.studio === studio && state.unitId === unitId && state.sheetId === initial.sheetId
        && state.userId === initial.userId && !studio.isDisposed && studio.currentUserId === initial.userId && studio.activeWorkbookId === unitId
        && studio.editor.state.activeKind === 'workbook' && studio.getWorkbookModel(unitId)?.activeSheetId === initial.sheetId
    }
    const controller = new AbortController(), task = { current: hostCurrent,
      cancel: () => { controller.abort(); submittingOwner?.cancel() },
      invalidate: () => { invalidated = true; controller.abort(); submittingOwner?.cancel() },
    }
    pending.value = task; progress.value = undefined
    try {
      const plan = planQuickValidation(action === 'dropdown' ? 'edit-dropdown' : action, [...initial.ranges], options.translator())
      if (!plan) return
      const owner = options.owner()
      submittingOwner = owner
      if (owner) {
        const result = await owner.submitAsync(plan.kind === 'set' ? { action: 'set', rule: plan.rule } : { action: 'clearTypes', ranges: [...plan.ranges], types: [...plan.types] })
        if (!result.accepted) throw new Error(result.error)
        if (owner.isCurrent() && plan.kind === 'set' && plan.edit) owner.showRule(plan.rule.id)
        return
      }
      if (!hostCurrent()) return
      const target = { unitId, sheetId: initial.sheetId, host: 'workbench' as const, expectedRevision: revision,
        expectedHistoryGeneration: generation, expectedUserId: initial.userId }
      const changed = plan.kind === 'set' ? await studio.executeCommand(SHEET_DATA_VALIDATION_ASYNC_COMMANDS.set,
        { ...target, rule: plan.rule, signal: controller.signal, onProgress: (value: SheetCheckboxRuleCellProgress) => {
          if (pending.value === task && !controller.signal.aborted && hostCurrent()) progress.value = value
        } })
        : studio.executeCommandSync(SHEET_DATA_VALIDATION_COMMANDS.clearTypes, { ...target, ranges: plan.ranges, types: plan.types })
      const completed = () => hostCurrent() && studio.getWorkbookRevision(unitId) === revision + 1
        && studio.editor.history.captureSavepoint(unitId).generation === generation
      if (!changed || !completed()) return
      options.refresh(studio, unitId, completed)
      if (completed() && plan.kind === 'set' && plan.edit) options.openRule(plan.rule.id)
    } catch (cause) {
      if (!stopped && hostCurrent()) options.notify(cause instanceof Error ? cause.message : '数据验证操作未能完成。')
    } finally { if (pending.value === task) { pending.value = undefined; progress.value = undefined } }
  }
  watch(() => {
    const state = options.state()
    return [state.studio, state.unitId, state.sheetId, state.userId, state.active, state.blocked] as const
  }, () => { if (pending.value && !pending.value.current()) pending.value.invalidate() }, { flush: 'sync' })
  onScopeDispose(() => { stopped = true; pending.value?.invalidate() })
  return { apply, cancel, busy: computed(() => !!pending.value), progress }
}
