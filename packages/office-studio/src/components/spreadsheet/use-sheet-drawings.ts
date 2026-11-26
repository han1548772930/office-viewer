import { nextTick, ref } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { sheetDrawingCommandInput, type SheetDrawingCommandId, type SheetDrawingCommandInput } from '../../api/sheet-drawing-commands.ts'
import { getSheetDrawingSelection } from '../../api/sheet-drawing-selection.ts'
import { createBlankWorkbook, type CellRange } from '../../domain/model.ts'
import type { SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { applyExcelWorkbookView, toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { planDrawingAlign, planDrawingArrange, planDrawingGroup, planDrawingGroupTransform, planDrawingMove, planDrawingSelectionTransform, planDrawingUngroup,
  SHEET_DRAWING_GROUP_COMMAND_IDS, type SheetDrawingPlanResult } from '../../features/drawing-groups/index.ts'
import { applySheetDrawingTransaction, createSheetDrawingScene, removeSheetDrawings, sheetDrawingRefAffectedRanges,
  sheetDrawingTransactionAffectedRanges } from '../../features/sheet-drawings/sheet-drawing-host.ts'
import { sheetDrawingMoveExtent } from '../../features/sheet-drawings/sheet-drawing-move.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { captureSheetAxisCommandHost, type SheetAxisHostChangeOptions } from './sheet-axis-command-host.ts'

interface SheetDrawingUiContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  selection(): readonly SheetDrawingRef[]
  select(refs: readonly SheetDrawingRef[]): void
  userId(): string
  active(): boolean
  finishEdit(): boolean
  blocked(): boolean
  prepareView?(next: ExcelWorkbookView, previous: ExcelWorkbookView): ExcelWorkbookView
  applyView(view: ExcelWorkbookView): void
  changed(options?: SheetAxisHostChangeOptions): void
  notify(message: string): void
  focus(): void
}
/** The view adapter only captures/publishes. Models and planners remain the common authority. */
export function useSheetDrawings(context: SheetDrawingUiContext) {
  const error = ref('')
  async function runDrawingCommand(id: string, params: SheetDrawingCommandInput = {}): Promise<boolean> {
    if (!Object.values(SHEET_DRAWING_GROUP_COMMAND_IDS).includes(id as SheetDrawingCommandId)) throw new TypeError('未知绘图命令。')
    if (!context.active() || context.blocked() || !context.finishEdit()) return false
    const command = id as SheetDrawingCommandId, host = captureSheetAxisCommandHost(context), userId = context.userId(), signature = JSON.stringify(host.view)
    let changed = false
    try {
      const input = sheetDrawingCommandInput(params), { studio, view, sheetId } = host,
        revision = studio && view.id ? studio.getWorkbookRevision(view.id) : undefined,
        historyGeneration = studio && view.id ? studio.editor.history.captureSavepoint(view.id).generation : undefined
      if (!host.sameTarget() || input.unitId !== undefined && input.unitId !== view.id || (input.sheetId ?? input.subUnitId ?? sheetId) !== sheetId
        || input.expectedRevision !== undefined && input.expectedRevision !== revision || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== historyGeneration) {
        throw new Error('绘图目标内容或历史已更改。')
      }
      const capturedSelection = context.selection().map(ref => ({ ...ref }))
      if (input.expectedSelection !== undefined && refSet(input.expectedSelection) !== refSet(capturedSelection)) throw new Error('绘图选择已更改。')
      input.selection ??= capturedSelection
      const current = () => {
        if (!host.sourceCurrent() || context.userId() !== userId || JSON.stringify(context.workbook()) !== signature
          || refSet(context.selection()) !== refSet(capturedSelection) || input.signal?.aborted) throw new Error('绘图宿主、选择或内容已更改。')
      }
      current(); error.value = ''
      if (studio && view.id) {
        changed = await studio.executeCommand(command, { ...input, unitId: view.id, sheetId, expectedRevision: revision,
          expectedHistoryGeneration: historyGeneration, host: { unitId: view.id, sheetId } })
        const model = studio.getWorkbookModel(view.id)
        if (changed && model && host.sameTarget() && studio.getWorkbookRevision(view.id) === revision! + 1) {
          context.applyView(toExcelWorkbookView(model))
          context.select(getSheetDrawingSelection(studio, view.id, sheetId))
        }
      } else {
        const model = applyExcelWorkbookView(createBlankWorkbook(view.name), cloneModel(view)), sheet = model.sheets.find(sheet => sheet.id === sheetId)
        if (!sheet || !canViewWorksheetProtection(sheet, userId) || !canEditWorksheetProtection(sheet, userId)) throw new Error('没有操作工作表绘图的权限。')
        const scene = createSheetDrawingScene(model, sheetId), selection = input.selection,
          before = cloneModel(model), ranges: CellRange[] = []
        let nextSelection: SheetDrawingRef[] = selection
        if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.remove) {
          const removed = removeSheetDrawings(model, sheetId, selection)
          if (!removed.removedRefs.length) return false
          ranges.push(...sheetDrawingRefAffectedRanges(before, sheetId, removed.affectedRefs), ...sheetDrawingRefAffectedRanges(model, sheetId, removed.affectedRefs))
          nextSelection = removed.selection
        } else {
          let plan: SheetDrawingPlanResult
          if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.group) plan = planDrawingGroup(scene, { selection, groupId: input.groupId ?? crypto.randomUUID() })
          else if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup) plan = planDrawingUngroup(scene, { selection })
          else if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.align) {
            if (input.alignType === undefined) throw new TypeError('需要指定绘图对齐方式。')
            plan = planDrawingAlign(scene, { selection, alignType: input.alignType })
          } else if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.move) {
            if (input.direction === undefined) throw new TypeError('需要指定绘图移动方向。')
            plan = planDrawingMove(scene, { selection, direction: input.direction, step: input.step, extent: sheetDrawingMoveExtent(model, sheetId) })
          } else if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.arrange) {
            const arrangeType = typeof input.arrangeType === 'number' ? (['forward', 'backward', 'front', 'back'] as const)[input.arrangeType] : input.arrangeType
            if ((arrangeType === undefined) === (input.zOrder === undefined)) throw new TypeError('需要且只能指定排列方式或层级。')
            plan = planDrawingArrange(scene, input.zOrder === undefined ? { selection, arrangeType: arrangeType! } : { selection, zOrder: input.zOrder })
          } else {
            if ((input.frame === undefined) === (input.delta === undefined)) throw new TypeError('需要且只能指定目标 frame 或世界 delta。')
            const transform = input.frame === undefined ? { delta: input.delta! } : { frame: input.frame }
            plan = input.groupId ? planDrawingGroupTransform(scene, { groupId: input.groupId, ...transform }) : planDrawingSelectionTransform(scene, { selection, ...transform })
          }
          if (plan.status === 'rejected') throw new Error(plan.message)
          if (plan.status === 'noop') return false
          ranges.push(...sheetDrawingTransactionAffectedRanges(model, sheetId, plan.transaction))
          nextSelection = applySheetDrawingTransaction(model, sheetId, plan.transaction).selection
        }
        if (!canViewProtectedRanges(sheet, ranges, userId) || !canEditProtectedRanges(sheet, ranges, userId)) throw new Error('没有操作绘图完整占用范围的权限。')
        const next = toExcelWorkbookView(model), prepared = context.prepareView?.(next, view) ?? next
        current()
        context.applyView(prepared); context.select(nextSelection)
        const published = host.publishedCurrent()
        if (published()) context.changed({ current: published, syncSelection: false })
        changed = true
      }
      const published = host.publishedCurrent()
      if (changed) await nextTick(() => { if (published() && context.active() && !context.blocked()) context.focus() })
      return changed
    } catch (reason) {
      const message = reason instanceof Error ? reason.message : String(reason)
      error.value = message
      if (host.sameTarget()) context.notify(message)
      return false
    }
  }
  return { runDrawingCommand, error }
}
const refSet = (refs: readonly SheetDrawingRef[]) => JSON.stringify(refs.map(ref => `${ref.kind}:${ref.id}`).sort())
