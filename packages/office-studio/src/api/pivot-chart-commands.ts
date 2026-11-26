import type { ChartModel, PivotChartControls, PivotTableModel, WorkbookModel } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { equivalentModelData } from '../platform/commands/model-equivalence.ts'
import { createChartDefinition, updateChartDefinition, type ChartConfig } from '../features/charts/chart-engine.ts'
import { pruneChartImageResources } from '../features/charts/chart-image-resources.ts'
import { assertChartReadPermission, type ChartAuthoringPlan } from '../features/charts/services/chart-statistics-service.ts'
import { sheetDrawingRefAffectedRanges } from '../features/sheet-drawings/sheet-drawing-host.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../features/sheet-protection/range-protection.ts'
import { applyPivotChartEdit, createPivotChartReport, getPivotChartReport, PivotChartEditError,
  type PivotChartReportEdit } from '../features/pivot/services/pivot-chart-engine.ts'
import { createPivotTable } from '../features/pivot/services/pivot-engine.ts'
import { preparePivotChartPairedTable } from '../features/pivot/services/pivot-chart-create-table.ts'
import { normalizePivotChartControls, pivotChartControls } from '../features/pivot/services/pivot-chart-controls.ts'
import { PIVOT_CHART_COMMANDS, type PivotChartCommandContext, type PivotChartConfigureInput,
  type PivotChartInsertInfo, type PivotChartInfo, type PivotChartInsertOptions } from '../features/pivot/services/pivot-chart-api-types.ts'
import type { OfficeStudio } from './office-studio.ts'
import type { ApiEventMap } from './events.ts'

export interface PivotChartTarget extends PivotChartCommandContext {
  unitId: string
  sheetId: string
  pivotChartId: string
}

export interface PivotChartTargetInput extends PivotChartCommandContext {
  unitId: string
  sheetId?: string
  subUnitId?: string
  pivotChartId: string
}

const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
const identity = (value: unknown): value is string => typeof value === 'string' && !!value.trim()
const contentSignature = (workbook: WorkbookModel): string => JSON.stringify({ ...workbook, activeSheetId: undefined, modified: undefined })

/** Fixed Sheets commands name the host subUnitId; local callers also use sheetId. */
export function resolvePivotChartTarget<T extends PivotChartTargetInput>(raw: T): T & PivotChartTarget {
  const input = cloneModel(raw)
  if (!record(input) || input.sheetId !== undefined && !identity(input.sheetId)
    || input.subUnitId !== undefined && !identity(input.subUnitId)
    || input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) {
    throw new TypeError('透视图命令的宿主身份无效或不一致。')
  }
  return { ...input, sheetId: input.sheetId ?? input.subUnitId ?? '' }
}

function requireChart(workbook: WorkbookModel, input: PivotChartTarget): ChartModel {
  const chart = workbook.charts?.find(candidate => candidate.id === input.pivotChartId && candidate.sheetId === input.sheetId)
  if (!chart) throw new PivotChartEditError('invalid-resource', '透视图或其工作表已被删除。')
  getPivotChartReport(workbook, chart)
  return chart
}

function assertEditPermission(workbook: WorkbookModel, chart: ChartModel, userId: string, reportEdit: boolean): void {
  assertChartReadPermission(workbook, chart, userId)
  const sheet = workbook.sheets.find(candidate => candidate.id === chart.sheetId)
  const ranges = sheetDrawingRefAffectedRanges(workbook, chart.sheetId, [{ kind: 'chart', id: chart.id }])
  if (!sheet || !canViewWorksheetProtection(sheet, userId) || !canViewProtectedRanges(sheet, ranges, userId)
    || !canEditWorksheetProtection(sheet, userId) || !canEditProtectedRanges(sheet, ranges, userId)) {
    throw new Error('没有编辑透视图完整范围的权限。')
  }
  if (reportEdit && chart.pivotTableId) {
    const report = getPivotChartReport(workbook, chart), host = workbook.sheets.find(candidate => candidate.id === report.anchor.subUnitId)
    const output = report.outputRange ?? { startRow: report.anchor.row, endRow: report.anchor.row, startCol: report.anchor.col, endCol: report.anchor.col }
    if (!host || !canViewWorksheetProtection(host, userId) || !canViewProtectedRanges(host, [output], userId)
      || !canEditWorksheetProtection(host, userId) || !canEditProtectedRanges(host, [output], userId)) {
      throw new Error('没有编辑关联透视表输出范围的权限。')
    }
  }
}

/** Menus use the same complete drawing/report permission check as content commands. */
export function canEditPivotChart(workbook: WorkbookModel, chart: ChartModel, userId: string, reportEdit = true): boolean {
  try { assertEditPermission(workbook, chart, userId, reportEdit); return true }
  catch { return false }
}

function targetGuard(studio: OfficeStudio, input: PivotChartTarget): { workbook: WorkbookModel; assertCurrent(): void } {
  if (!record(input) || !identity(input.unitId) || !identity(input.sheetId) || !identity(input.pivotChartId)) throw new TypeError('透视图命令身份无效。')
  if (input.host !== undefined && input.host !== 'workbench') throw new TypeError('透视图交互宿主无效。')
  if (input.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0)
    || input.expectedHistoryGeneration !== undefined && (!Number.isSafeInteger(input.expectedHistoryGeneration) || input.expectedHistoryGeneration <= 0)
    || input.expectedUserId !== undefined && !identity(input.expectedUserId)) throw new TypeError('透视图内容、历史或用户上下文无效。')
  const workbook = studio.getWorkbookModel(input.unitId)
  if (!workbook || !workbook.sheets.some(sheet => sheet.id === input.sheetId)) throw new PivotChartEditError('invalid-resource', '透视图工作簿或宿主不可用。')
  const revision = studio.getWorkbookRevision(input.unitId), userId = studio.currentUserId,
    history = studio.editor.history.captureSavepoint(input.unitId), signature = contentSignature(workbook)
  const assertCurrent = (): void => {
    const current = studio.getWorkbookModel(input.unitId), currentHistory = studio.editor.history.captureSavepoint(input.unitId)
    if (studio.isDisposed || current !== workbook || studio.getWorkbookRevision(input.unitId) !== revision
      || contentSignature(workbook) !== signature || studio.currentUserId !== userId
      || currentHistory.revision !== history.revision || currentHistory.generation !== history.generation
      || input.expectedRevision !== undefined && input.expectedRevision !== revision
      || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== history.generation
      || input.expectedUserId !== undefined && input.expectedUserId !== userId) throw new Error('透视图目标内容、历史或用户已更改。')
    if (input.host === 'workbench' && (studio.editor.state.activeKind !== 'workbook'
      || studio.activeWorkbookId !== input.unitId || current.activeSheetId !== input.sheetId)) throw new Error('透视图工作台宿主已切换。')
  }
  assertCurrent()
  return { workbook, assertCurrent }
}

function beforeChange(studio: OfficeStudio, input: PivotChartTarget, commandId: string, chart?: ChartModel, reportEdit = false,
  linkedReport?: PivotTableModel, assertCurrent?: () => void): void {
  const event = { unitId: input.unitId, subUnitId: input.sheetId, pivotChartId: input.pivotChartId, commandId, cancel: false }
  const report = reportEdit && chart?.pivotTableId ? getPivotChartReport(studio.getWorkbookModel(input.unitId)!, chart) : linkedReport
  const pivotEvent = report ? { unitId: input.unitId, subUnitId: report.anchor.subUnitId, pivotTableId: report.id, commandId, cancel: false } : undefined
  // Both targets are captured before a listener can mutate or replace a report.
  studio.events.emit('BeforePivotChartChange', event)
  if (event.cancel) throw new Error('透视图修改已取消。')
  assertCurrent?.()
  if (pivotEvent) {
    studio.events.emit('BeforePivotTableChange', pivotEvent)
    if (pivotEvent.cancel) throw new Error('关联透视表修改已取消。')
  }
}

function publishRelated(studio: OfficeStudio, before: WorkbookModel, after: WorkbookModel): void {
  const added: ApiEventMap['PivotChartAdded'][] = [], changed: ApiEventMap['PivotChartChanged'][] = [],
    removed: ApiEventMap['PivotChartRemoved'][] = [], pivots: ApiEventMap['PivotTableRendered'][] = [], controls: ApiEventMap['SlicerChanged'][] = []
  for (const chart of after.charts ?? []) {
    if ((!chart.ownedPivot && !chart.pivotTableId) || equivalentModelData(before.charts?.find(previous => previous.id === chart.id), chart)) continue
    const previous = before.charts?.find(candidate => candidate.id === chart.id), report = chart.ownedPivot ?? after.pivotTables?.find(candidate => candidate.id === chart.pivotTableId)
    // Retained unavailable imports have diagnostics; they are not an invented report result.
    if (!report) continue
    if (!previous) added.push({ unitId: after.id!, subUnitId: chart.sheetId, pivotChartId: chart.id, sourceKind: chart.ownedPivot ? 'range' : 'pivotTable' })
    changed.push({ unitId: after.id!, subUnitId: chart.sheetId, pivotChartId: chart.id, revision: chart.revision, reportRevision: report.revision })
  }
  for (const chart of before.charts ?? []) if ((chart.ownedPivot || chart.pivotTableId) && !after.charts?.some(candidate => candidate.id === chart.id)) {
    removed.push({ unitId: before.id!, subUnitId: chart.sheetId, pivotChartId: chart.id })
  }
  for (const report of after.pivotTables ?? []) {
    if (equivalentModelData(before.pivotTables?.find(previous => previous.id === report.id), report)) continue
    pivots.push({ unitId: after.id!, subUnitId: report.anchor.subUnitId, pivotTableId: report.id, revision: report.revision,
      ...(report.outputRange ? { outputRange: cloneModel(report.outputRange) } : {}) })
  }
  for (const view of after.slicers ?? []) if (!equivalentModelData(before.slicers?.find(previous => previous.id === view.id), view)) {
    controls.push({ unitId: after.id!, hostSheetId: view.hostSheetId, slicerId: view.id, revision: view.revision })
  }
  // Capture all payloads before any observer can start a different transaction.
  for (const payload of added) studio.publishObserverEvent('PivotChartAdded', () => ({ ...payload }))
  for (const payload of changed) studio.publishObserverEvent('PivotChartChanged', () => ({ ...payload }))
  for (const payload of removed) studio.publishObserverEvent('PivotChartRemoved', () => ({ ...payload }))
  for (const payload of pivots) {
    studio.publishObserverEvent('PivotTableChanged', () => ({ unitId: payload.unitId, subUnitId: payload.subUnitId,
      pivotTableId: payload.pivotTableId, revision: payload.revision }))
    studio.publishObserverEvent('PivotTableRendered', () => cloneModel(payload))
  }
  for (const payload of controls) studio.publishObserverEvent('SlicerChanged', () => ({ ...payload }))
}

function modernSource(info: PivotChartInsertInfo): info is PivotChartInfo {
  return record(info.source) && 'kind' in info.source
}

function createDefinition(workbook: WorkbookModel, input: PivotChartTarget, info: PivotChartInsertInfo, pairedTable?: PivotTableModel): ChartModel {
  if (!record(info) || !record(info.source)) throw new TypeError('透视图创建信息无效。')
  let config: ChartConfig, fieldsConfig: PivotChartInfo['initialFieldsConfig']
  if (modernSource(info)) {
    const { source, initialFieldsConfig, ...visual } = info
    fieldsConfig = initialFieldsConfig
    if (Object.hasOwn(info, 'pivotTableId') || Object.hasOwn(info, 'ownedPivot')) throw new TypeError('透视图不能混合新旧来源身份。')
    if (source.kind === 'range') {
      if (!record(source.sourceRangeInfo) || Object.keys(source).some(key => !['kind', 'sourceRangeInfo'].includes(key))) throw new TypeError('独立透视图来源无效。')
      if (pairedTable) {
        const report = createPivotTable(workbook, source.sourceRangeInfo, pairedTable.anchor, pairedTable.id)
        config = { ...visual, pivotTableId: report.id, source: { sheetId: report.anchor.subUnitId,
          range: cloneModel(report.outputRange ?? { startRow: report.anchor.row, endRow: report.anchor.row, startCol: report.anchor.col, endCol: report.anchor.col }), orientation: 'columns' } }
      } else {
        const report = createPivotChartReport(workbook, source.sourceRangeInfo, input.sheetId, `${input.pivotChartId}:report`)
        config = { ...visual, ownedPivot: report, source: { sheetId: report.source.subUnitId, range: cloneModel(report.source.range), orientation: 'columns' } }
      }
    } else if (source.kind === 'pivotTable') {
      const reference = source.pivotTableRef
      if (!record(reference) || !identity(reference.pivotTableId) || reference.unitId !== workbook.id
        || Object.keys(reference).some(key => !['unitId', 'subUnitId', 'pivotTableId'].includes(key))
        || Object.keys(source).some(key => !['kind', 'pivotTableRef'].includes(key))) throw new TypeError('关联透视图来源无效。')
      const report = workbook.pivotTables?.find(candidate => candidate.id === reference.pivotTableId && candidate.anchor.subUnitId === reference.subUnitId)
      if (!report) throw new PivotChartEditError('invalid-resource', '关联透视表或其宿主不存在。')
      const range = report.outputRange ?? { startRow: report.anchor.row, endRow: report.anchor.row, startCol: report.anchor.col, endCol: report.anchor.col }
      config = { ...visual, pivotTableId: report.id, source: { sheetId: report.anchor.subUnitId, range: cloneModel(range), orientation: 'columns' } }
    } else throw new TypeError('透视图来源类型无效。')
  } else {
    if (!identity(info.pivotTableId) || info.ownedPivot) throw new TypeError('关联透视图须引用一个透视表。')
    const report = workbook.pivotTables?.find(candidate => candidate.id === info.pivotTableId)
    if (!report) throw new PivotChartEditError('invalid-resource', '关联透视表不存在。')
    const range = report.outputRange ?? { startRow: report.anchor.row, endRow: report.anchor.row, startCol: report.anchor.col, endCol: report.anchor.col }
    config = { ...info, source: { sheetId: report.anchor.subUnitId, range: cloneModel(range), orientation: 'columns' } }
  }
  const chart = createChartDefinition(workbook, input.sheetId, config, input.pivotChartId)
  workbook.charts ??= []
  workbook.charts.push(chart)
  getPivotChartReport(workbook, chart)
  if (fieldsConfig !== undefined) applyPivotChartEdit(workbook, chart, { type: 'setFieldsConfig', config: fieldsConfig })
  return chart
}

async function commitExisting(studio: OfficeStudio, input: PivotChartTarget, commandId: string,
  update: (workbook: WorkbookModel, chart: ChartModel, authoring: ChartAuthoringPlan) => void, reportEdit = true): Promise<boolean> {
  const { workbook, assertCurrent } = targetGuard(studio, input), chart = requireChart(workbook, input)
  assertEditPermission(workbook, chart, studio.currentUserId, reportEdit)
  beforeChange(studio, input, commandId, chart, reportEdit, undefined, assertCurrent)
  assertCurrent()
  const authoring: ChartAuthoringPlan = { previous: cloneModel(chart) }
  const changed = await studio.commitChartChange(input.unitId, '编辑透视图', input.pivotChartId, candidate => {
    assertCurrent()
    const target = requireChart(candidate, input)
    update(candidate, target, authoring)
    assertEditPermission(candidate, target, studio.currentUserId, reportEdit)
  }, authoring, assertCurrent)
  if (changed) publishRelated(studio, workbook, studio.getWorkbookModel(input.unitId)!)
  return true
}

/** Fixed public action aliases retain their command ID while sharing the exact report transaction. */
export function commitPivotChartReportEdit(studio: OfficeStudio, input: PivotChartTarget,
  edit: PivotChartReportEdit, commandId: string = PIVOT_CHART_COMMANDS.edit): Promise<boolean> {
  const request = cloneModel({ ...input, edit })
  return commitExisting(studio, request, commandId, (candidate, chart) => { applyPivotChartEdit(candidate, chart, request.edit) })
}

export function registerPivotChartCommands(studio: OfficeStudio): void {
  studio.registerCommand<PivotChartTarget & { chart: PivotChartInsertInfo; options?: PivotChartInsertOptions }, Promise<boolean>>(PIVOT_CHART_COMMANDS.insert, async input => {
    input = cloneModel(input)
    const { workbook, assertCurrent } = targetGuard(studio, input)
    if (workbook.charts?.some(chart => chart.id === input.pivotChartId)) throw new PivotChartEditError('invalid-resource', '透视图 ID 已存在。')
    studio.assertCanEditWorksheet(input.unitId, input.sheetId)
    if (!record(input.chart) || !record(input.chart.source)) throw new TypeError('透视图创建信息无效。')
    const pairedTable = preparePivotChartPairedTable(workbook, input.pivotChartId, input.chart, input.options, studio.currentUserId)
    let linkedReport: PivotTableModel | undefined = pairedTable
    if (modernSource(input.chart) && input.chart.source.kind === 'pivotTable' && input.chart.initialFieldsConfig !== undefined) {
      const reference = input.chart.source.pivotTableRef
      if (!record(reference) || reference.unitId !== workbook.id || !identity(reference.pivotTableId) || !identity(reference.subUnitId)) throw new TypeError('关联透视图来源无效。')
      linkedReport = workbook.pivotTables?.find(report => report.id === reference.pivotTableId && report.anchor.subUnitId === reference.subUnitId)
      if (!linkedReport) throw new PivotChartEditError('invalid-resource', '关联透视表或其宿主不存在。')
    }
    beforeChange(studio, input, PIVOT_CHART_COMMANDS.insert, undefined, false, linkedReport, assertCurrent)
    assertCurrent()
    const changed = await studio.commitChartChange(input.unitId, '插入透视图', input.pivotChartId, candidate => {
      assertCurrent()
      const chart = createDefinition(candidate, input, input.chart, pairedTable)
      assertEditPermission(candidate, chart, studio.currentUserId, Boolean(pairedTable) || modernSource(input.chart) && input.chart.initialFieldsConfig !== undefined)
    }, { initializeDefaults: true }, assertCurrent)
    if (changed) publishRelated(studio, workbook, studio.getWorkbookModel(input.unitId)!)
    return changed
  }, { asynchronous: true })
  studio.registerCommand<PivotChartTarget & { edit: PivotChartReportEdit }, Promise<boolean>>(PIVOT_CHART_COMMANDS.edit, input => {
    return commitPivotChartReportEdit(studio, input, input.edit)
  }, { asynchronous: true })
  studio.registerCommand<PivotChartTarget, Promise<boolean>>(PIVOT_CHART_COMMANDS.refresh, raw => {
    const input = cloneModel(raw)
    return commitExisting(studio, input, PIVOT_CHART_COMMANDS.refresh, (candidate, chart) => { applyPivotChartEdit(candidate, chart, { type: 'refresh' }) })
  }, { asynchronous: true })
  studio.registerCommand<PivotChartTargetInput & { controls: Partial<PivotChartControls> }, Promise<boolean>>(PIVOT_CHART_COMMANDS.setControls, raw => {
    const input = resolvePivotChartTarget(raw)
    normalizePivotChartControls(input.controls)
    return commitExisting(studio, input, PIVOT_CHART_COMMANDS.setControls, (_candidate, chart) => {
      const controls = normalizePivotChartControls({ ...pivotChartControls(chart), ...input.controls })
      if (equivalentModelData(pivotChartControls(chart), controls)) return
      chart.controls = controls; chart.revision += 1
    }, false)
  }, { asynchronous: true })
  studio.registerCommand<PivotChartTarget & PivotChartConfigureInput, Promise<boolean>>(PIVOT_CHART_COMMANDS.configure, raw => {
    const input = cloneModel(raw)
    return commitExisting(studio, input, PIVOT_CHART_COMMANDS.configure, (candidate, chart, authoring) => {
      if (!record(input.patch)) throw new TypeError('透视图作者设置无效。')
      const patch = cloneModel(input.patch), originalRevision = chart.revision
      for (const key of ['source', 'pivotTableId', 'ownedPivot'] as const) {
        if (Object.hasOwn(patch, key) && !equivalentModelData(patch[key], chart[key])) throw new TypeError('作者设置不能替换透视图报告来源或归属。')
        delete patch[key]
      }
      if (input.fieldsConfig !== undefined) applyPivotChartEdit(candidate, chart, { type: 'setFieldsConfig', config: input.fieldsConfig })
      if (Object.keys(patch).length) {
        const next = updateChartDefinition(candidate, chart, patch)
        if (!equivalentModelData(next, chart, ['revision'])) Object.assign(chart, next)
      }
      authoring.patch = patch
      if (chart.revision !== originalRevision) chart.revision = originalRevision + 1
    }, input.fieldsConfig !== undefined)
  }, { asynchronous: true })
  studio.registerCommand<PivotChartTarget, boolean>(PIVOT_CHART_COMMANDS.remove, raw => {
    const input = cloneModel(raw), { workbook, assertCurrent } = targetGuard(studio, input), chart = requireChart(workbook, input)
    assertEditPermission(workbook, chart, studio.currentUserId, false)
    beforeChange(studio, input, PIVOT_CHART_COMMANDS.remove, chart, false, undefined, assertCurrent)
    assertCurrent()
    const changed = studio.commitWorkbook(input.unitId, '删除透视图', candidate => {
      assertCurrent()
      candidate.charts = candidate.charts?.filter(target => target.id !== input.pivotChartId)
      pruneChartImageResources(candidate)
    }, { expectedRevision: studio.getWorkbookRevision(input.unitId), expectedHistoryGeneration: studio.editor.history.captureSavepoint(input.unitId).generation })
    if (changed) publishRelated(studio, workbook, studio.getWorkbookModel(input.unitId)!)
    return changed
  })
}
