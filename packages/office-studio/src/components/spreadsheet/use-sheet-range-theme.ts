import { computed, nextTick, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { CellRange, SheetRangeThemeStyle, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { assertSheetRangesEditable, normalizeSheetCommandRange } from '../../features/sheet-workbench/sheet-range-command.ts'
import { applySheetRangeThemeWorkbookPlan, planWorkbookSheetRangeThemeCommand, sheetRangeThemeCommandInput, SHEET_RANGE_THEME_COMMANDS, type SheetRangeThemeCommandInput } from '../../features/sheet-workbench/sheet-range-theme-command.ts'
import { SHEET_BUILTIN_RANGE_THEME_NAMES } from '../../features/sheet-workbench/sheet-range-theme-presets.ts'
import { sheetRangeThemeByName, sheetUsedRangeThemeRule } from '../../features/sheet-workbench/sheet-range-theme-resolver.ts'
import { normalizeSheetRangeThemeStyle, type SheetRangeThemeStyleJson } from '../../features/sheet-workbench/sheet-range-theme-style.ts'
import { captureSheetAxisCommandHost, type SheetAxisHostChangeOptions } from './sheet-axis-command-host.ts'

interface SheetRangeThemeUiContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  ranges(): CellRange[]
  userId(): string
  active(): boolean
  finishEdit(): boolean
  prepareView?(view: ExcelWorkbookView, previous: ExcelWorkbookView): ExcelWorkbookView
  applyView(view: ExcelWorkbookView): void
  changed(options?: SheetAxisHostChangeOptions): void
  focus(): void
}
export interface SheetRangeThemeRequest {
  ranges: CellRange[]
  themeNames: string[]
  themeStyles: Record<string, SheetRangeThemeStyle>
  initialTheme: string
  canRemove: boolean
  mixed: boolean
}
interface PendingRangeTheme {
  request: SheetRangeThemeRequest
  input: SheetRangeThemeCommandInput
  signature?: string
  model?: WorkbookModel
  generation: number
  host: ReturnType<typeof captureSheetAxisCommandHost>
}

export function useSheetRangeTheme(context: SheetRangeThemeUiContext) {
  const pending = shallowRef<PendingRangeTheme | null>(null), rangeThemeError = ref('')
  let hostGeneration = 0
  let closeCurrent: (() => boolean) | undefined
  const request = computed(() => pending.value?.request ?? null)
  function cancelRangeTheme(): void {
    if (pending.value) {
      const current = pending.value.host.publishedCurrent(), generation = hostGeneration
      closeCurrent = () => generation === hostGeneration && current()
    }
    pending.value = null; rangeThemeError.value = ''
  }
  watch([context.studio, () => context.workbook().id ?? context.workbook().name, context.sheetId, context.active, context.userId], () => {
    hostGeneration++
    cancelRangeTheme()
    closeCurrent = undefined
  }, { flush: 'sync' })
  function afterRangeThemeClose(): void {
    const current = closeCurrent; closeCurrent = undefined
    void nextTick(() => { if (!pending.value && current?.()) context.focus() })
  }
  function runRangeThemeCommand(commandId: string, params?: unknown): boolean {
    if (commandId !== SHEET_RANGE_THEME_COMMANDS.set && commandId !== SHEET_RANGE_THEME_COMMANDS.remove || !context.active()) return false
    const input = sheetRangeThemeCommandInput(params)
    if (!context.finishEdit() || !context.active()) return false
    const host = captureSheetAxisCommandHost(context), { studio, view } = host, generation = hostGeneration
    const sameHost = () => generation === hostGeneration && host.sameTarget()
    if (!sameHost()) return false
    const sheetId = input.sheetId ?? input.subUnitId ?? host.sheetId
    const ranges = (input.ranges ?? input.selections ?? (input.range ? [input.range] : context.ranges())).map(normalizeSheetCommandRange)
    const target: SheetRangeThemeCommandInput = { ...input, unitId: input.unitId ?? view.id, sheetId, ranges }
    if (target.unitId !== view.id || sheetId !== host.sheetId) throw new Error('范围主题目标与当前工作台不一致。')
    if (studio && view.id) {
      const baseRevision = studio.getWorkbookRevision(view.id), historyGeneration = studio.editor.history.captureSavepoint(view.id).generation
      target.expectedRevision ??= baseRevision
      target.expectedHistoryGeneration ??= historyGeneration
      target.host = { unitId: view.id, sheetId }
      const changed = studio.executeCommandSync<SheetRangeThemeCommandInput, boolean>(commandId, target)
      if (!changed) return false
      if (!sameHost() || studio.getWorkbookRevision(view.id) !== baseRevision + 1
        || studio.editor.history.captureSavepoint(view.id).generation !== historyGeneration) return true
      const model = studio.getWorkbookModel(view.id)
      if (!model) return true
      const next = toExcelWorkbookView(model)
      if (!sameHost() || studio.getWorkbookModel(view.id) !== model || studio.getWorkbookRevision(view.id) !== baseRevision + 1) return true
      const signature = JSON.stringify(next)
      context.applyView(next)
      const publishedCurrent = host.publishedCurrent(baseRevision + 1)
      const current = () => sameHost() && publishedCurrent() && studio.getWorkbookModel(view.id!) === model
        && studio.editor.history.captureSavepoint(view.id!).generation === historyGeneration && JSON.stringify(context.workbook()) === signature
      if (current()) context.changed({ syncSelection: false, current })
      return true
    }
    if (input.unitId !== undefined && input.unitId !== view.id) throw new Error('找不到目标工作簿。')
    const signature = JSON.stringify(view), candidate = cloneModel(view), sheet = candidate.sheets.find(item => item.id === sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, context.userId())) throw new Error('找不到可访问的目标工作表。')
    const plan = planWorkbookSheetRangeThemeCommand(candidate, sheetId, ranges, commandId === SHEET_RANGE_THEME_COMMANDS.set ? 'set' : 'remove', input)
    for (const [affectedSheetId, affectedRanges] of Object.entries(plan.affected)) {
      const affectedSheet = candidate.sheets.find(item => item.id === affectedSheetId)
      if (!affectedSheet || !canViewWorksheetProtection(affectedSheet, context.userId())) throw new Error('找不到可访问的主题工作表。')
      assertSheetRangesEditable(affectedSheet, affectedRanges, context.userId())
    }
    if (!plan.changed) return false
    applySheetRangeThemeWorkbookPlan(candidate, plan)
    const next = context.prepareView?.(candidate, cloneModel(view)) ?? candidate
    if (!sameHost() || !host.sourceCurrent() || JSON.stringify(context.workbook()) !== signature) return false
    if (view.id === undefined) delete next.id
    else next.id = view.id
    const preparedSignature = JSON.stringify(next)
    context.applyView(next)
    const publishedCurrent = host.publishedCurrent()
    const current = () => sameHost() && publishedCurrent() && JSON.stringify(context.workbook()) === preparedSignature
    if (current()) context.changed({ syncSelection: false, current })
    return true
  }
  function openRangeThemeDialog(params?: unknown): void {
    if (!context.active() || !context.finishEdit() || !context.active()) return
    const input = sheetRangeThemeCommandInput(params), host = captureSheetAxisCommandHost(context), { view, studio } = host, generation = hostGeneration
    if (!host.sameTarget()) return
    if (input.unitId !== undefined && input.unitId !== view.id || (input.sheetId ?? input.subUnitId ?? host.sheetId) !== host.sheetId) throw new Error('请在目标工作表打开范围主题。')
    const sheet = view.sheets.find(item => item.id === host.sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, context.userId())) throw new Error('找不到可访问的目标工作表。')
    const ranges = cloneModel((input.ranges ?? input.selections ?? (input.range ? [input.range] : context.ranges())).map(normalizeSheetCommandRange))
    if (!ranges.length) return
    assertSheetRangesEditable(sheet, ranges, context.userId())
    const themeNames = [...new Set([...SHEET_BUILTIN_RANGE_THEME_NAMES, ...Object.keys(view.rangeThemes ?? sheet.rangeThemes ?? {})])]
    const themeStyles = Object.fromEntries(themeNames.flatMap(name => {
      const style = sheetRangeThemeByName(name, { rangeThemes: view.rangeThemes ?? sheet.rangeThemes })
      return style ? [[name, cloneModel(style)]] : []
    }))
    const current = ranges.map(range => sheetUsedRangeThemeRule(sheet, range)?.themeName)
    const mixed = current.some(name => name !== current[0])
    const preferred = !mixed && current[0] && themeNames.includes(current[0]) ? current[0] : 'default'
    const initialTheme = input.themeName ?? preferred
    if (!themeNames.includes(initialTheme)) throw new Error(`找不到范围主题：${initialTheme}`)
    const expectedRevision = studio && view.id ? studio.getWorkbookRevision(view.id) : undefined
    const expectedHistoryGeneration = studio && view.id ? studio.editor.history.captureSavepoint(view.id).generation : undefined
    const model = studio && view.id ? studio.getWorkbookModel(view.id) : undefined
    if (input.expectedRevision !== undefined && input.expectedRevision !== expectedRevision) throw new Error('工作簿已更改，请重新打开范围主题。')
    if (input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== expectedHistoryGeneration) throw new Error('工作簿历史已更改，请重新打开范围主题。')
    if (generation !== hostGeneration || !host.sameTarget()) return
    pending.value = { host, generation, model, signature: studio && view.id ? undefined : JSON.stringify(view),
      input: { ...input, unitId: view.id, sheetId: host.sheetId, ranges, expectedRevision, expectedHistoryGeneration },
      request: { ranges, themeNames, themeStyles, initialTheme, mixed, canRemove: current.some(name => name !== undefined) } }
    rangeThemeError.value = ''; closeCurrent = undefined
  }
  function confirmRangeTheme(action: 'apply' | 'remove', themeName?: string, draft?: SheetRangeThemeStyleJson): void {
    const state = pending.value
    if (!state || !context.active()) return
    try {
      if (!context.finishEdit()) return
      if (state.generation !== hostGeneration || !state.host.sameTarget()) throw new Error('目标工作簿或工作表已切换，请重新打开范围主题。')
      if (state.model && state.host.studio?.getWorkbookModel(state.input.unitId!) !== state.model) throw new Error('工作簿实例已更改，请重新打开范围主题。')
      if (state.signature !== undefined && !state.host.sourceCurrent()) throw new Error('范围主题工作簿宿主已更改，请重新打开范围主题。')
      if (state.signature !== undefined && state.signature !== JSON.stringify(context.workbook())) throw new Error('工作簿已更改，请重新打开范围主题。')
      const definition = action === 'apply' && draft !== undefined ? normalizeSheetRangeThemeStyle(draft) : undefined
      if (action === 'apply' && (!themeName || !definition && !state.request.themeNames.includes(themeName))) throw new Error('请选择可用的范围主题。')
      if (definition && (definition.name !== themeName || SHEET_BUILTIN_RANGE_THEME_NAMES.includes(definition.name))) throw new Error('请为编辑的主题选择可用的自定义名称。')
      runRangeThemeCommand(SHEET_RANGE_THEME_COMMANDS.set, { ...state.input, themeName: action === 'remove' ? null : themeName,
        rangeThemeStyle: undefined, rangeThemeStyleJson: definition })
      if (pending.value === state) cancelRangeTheme()
    } catch (error: unknown) { if (pending.value === state) rangeThemeError.value = error instanceof Error ? error.message : '无法设置范围主题。' }
  }
  return { pending, request, rangeThemeError, openRangeThemeDialog, runRangeThemeCommand, confirmRangeTheme, cancelRangeTheme, afterRangeThemeClose }
}
